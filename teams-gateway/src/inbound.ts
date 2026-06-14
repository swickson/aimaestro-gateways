/**
 * Teams gateway — inbound pipeline (personal scope, v1).
 *
 * Teams Activity -> dedupe -> user-resolve -> tenant-scoped trust -> content
 * scan/wrap (shared scanner) -> typed EnrichedContext envelope -> AMP
 * `/api/v1/route`, originating from the bot's own AMP identity.
 *
 * SDK DECOUPLING: this module never imports `@microsoft/teams.apps`. `server.ts`
 * extracts the fields it needs from the SDK `ctx` (including mention-stripping the
 * text and grabbing `ctx.ref`) into a plain `InboundActivity`, then hands it here.
 * That keeps the security-critical pipeline unit-testable with mock activities and
 * no live Azure endpoint (PLAN: "testable WITHOUT a live Azure endpoint").
 *
 * ACK-FAST (red-team §0.1 C1): the `@microsoft/teams.apps` ExpressAdapter sends
 * HTTP 200 only AFTER the route handler's promise resolves (verified in the SDK
 * source). So `server.ts` fires `handleInbound` WITHOUT awaiting and returns
 * immediately — the heavy `/api/v1/route` round-trip runs async, off the response
 * path, so Teams never hits its ~10-15s redelivery window. A post-200 route
 * failure is therefore invisible to Teams — it MUST be observable in
 * `[TEAMS]`/`[AMP]` logs (NON-NEGOTIABLE), never silently dropped.
 *
 * DEDUPE (C1): Bot Framework redelivers an Activity (same `activity.id`) when it
 * doesn't see a fast 200. The seen-set, checked+marked before the first `await`,
 * collapses retries to a single route.
 */

import { Cache } from '@aimaestro/common/cache.js';
import type { ConversationReference } from '@microsoft/teams.api';
import type { AMPAttachmentV1, AMPRouteRequest, AMPRouteResponse, AttachmentPolicy, EnrichedContext, OperatorAadRef } from './types.js';
import { resolveTrust, sanitizeTeamsMessage } from './content-security.js';
import { ingestAttachments, type DownloadAttachment, type RawInboundAttachment } from './attachments-inbound.js';
import type { ThreadStore } from './thread-store.js';
import type { UserResolver } from './user-resolver.js';

const CONVERSATION_TIMEOUT_MS = 30 * 60 * 1000; // 30 min — personal-scope "new conversation" window.
const MAX_TOPIC_HINTS = 3;

// ---------------------------------------------------------------------------
// Topic hints (ported from discord-gateway/src/inbound.ts; transitional dup —
// candidate for packages/common later).
// ---------------------------------------------------------------------------

const STOP_WORDS = new Set([
  'a', 'an', 'the', 'is', 'are', 'was', 'were', 'be', 'been', 'being',
  'have', 'has', 'had', 'do', 'does', 'did', 'will', 'would', 'could',
  'should', 'may', 'might', 'can', 'shall', 'to', 'of', 'in', 'for',
  'on', 'with', 'at', 'by', 'from', 'as', 'into', 'through', 'during',
  'before', 'after', 'above', 'below', 'between', 'out', 'off', 'over',
  'under', 'again', 'further', 'then', 'once', 'and', 'but', 'or', 'nor',
  'not', 'so', 'yet', 'both', 'either', 'neither', 'each', 'every', 'all',
  'any', 'few', 'more', 'most', 'other', 'some', 'such', 'no', 'only',
  'same', 'than', 'too', 'very', 'just', 'because', 'if', 'when', 'where',
  'how', 'what', 'which', 'who', 'whom', 'this', 'that', 'these', 'those',
  'i', 'me', 'my', 'we', 'our', 'you', 'your', 'he', 'him', 'his', 'she',
  'her', 'it', 'its', 'they', 'them', 'their', 'about', 'up',
]);

/** Up to 3 frequency-ranked keywords, stop-words + short tokens filtered. */
export function extractTopicHints(text: string): string[] {
  const words = text
    .toLowerCase()
    .replace(/[^a-z0-9\s-]/g, ' ')
    .split(/\s+/)
    // length>2, not a stop word, and at least one alphanumeric char — so a
    // hyphen-only token like '---' (a markdown rule) is never a "keyword".
    .filter((w) => w.length > 2 && !STOP_WORDS.has(w) && /[a-z0-9]/.test(w));

  const freq = new Map<string, number>();
  for (const w of words) {
    freq.set(w, (freq.get(w) || 0) + 1);
  }

  return [...freq.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, MAX_TOPIC_HINTS)
    .map(([word]) => word);
}

// ---------------------------------------------------------------------------
// SDK-decoupled inbound DTO + per-bot dependencies
// ---------------------------------------------------------------------------

/**
 * The SDK-independent shape `server.ts` extracts from a Teams message Activity.
 * `text` is ALREADY mention-stripped (bot mention only) by the extractor — the
 * scanner must see the unaltered injection payload, so only the bot's `<at>` tag
 * is removed upstream, nothing else.
 */
export interface InboundActivity {
  activityId: string;
  conversationId: string;
  conversationType: string;
  /** Sender AAD object id (canonical Teams user id). May be absent for non-user senders. */
  aadObjectId?: string;
  /** Bot Framework account id of the sender (fallback identity). */
  fromId: string;
  fromName: string;
  /** Mention-stripped message text. */
  text: string;
  tenantId?: string;
  serviceUrl?: string;
  /** Full conversation reference (`ctx.ref`) — drives outbound `continueConversation`. */
  reference: ConversationReference;
  /**
   * SDK-decoupled attachment descriptors extracted from the activity (w3). The
   * `downloadAttachment` closure (below) resolves their bytes. Absent/empty for a
   * text-only message.
   */
  attachments?: RawInboundAttachment[];
  /**
   * Byte downloader for `attachments`, bound by `server.ts` to the bot connector /
   * pre-auth Teams URL. Kept as a closure (not an SDK type) so this module stays
   * SDK-free — the one function that crosses, precedented by `reference: ctx.ref`.
   */
  downloadAttachment?: DownloadAttachment;
}

/** One bot's identity + the per-bot AMP credentials it routes as. */
export interface InboundBotContext {
  slug: string;
  defaultAgent: string;
  agentName: string;
  /** This bot's AMP address (sender identity). */
  ampAddress: string;
  /** This bot's AMP api key (the Bearer token = the sending identity on the route). */
  ampApiKey: string;
}

export interface InboundDeps {
  bot: InboundBotContext;
  maestroUrl: string;
  operatorAadObjectIds: OperatorAadRef[];
  userResolver: UserResolver;
  threadStore: ThreadStore;
  /** Shared activity.id seen-set (dedupe). */
  dedupe: Cache<true>;
  /** Gateway-side attachment caps + deny policy (w3). */
  attachmentPolicy: AttachmentPolicy;
  timeoutMs: number;
  debug: boolean;
  /** Injectable clock for deterministic tests. Defaults to `Date.now`. */
  now?: () => number;
}

/**
 * Process one inbound Teams message. Returns a small status for tests/observability;
 * `server.ts` fires this WITHOUT awaiting (ack-fast) and only logs rejections.
 */
export async function handleInbound(
  activity: InboundActivity,
  deps: InboundDeps,
): Promise<'routed' | 'duplicate' | 'dropped' | 'failed'> {
  const { bot } = deps;
  const now = deps.now ?? Date.now;
  const log = (msg: string) => console.log(`[TEAMS] (${bot.slug}) ${msg}`);

  // 1. Personal-scope gate (v1). Channel / groupChat enumerated in a later phase.
  if (activity.conversationType !== 'personal') {
    log(`dropping ${activity.conversationType} activity ${activity.activityId} — personal scope only (v1).`);
    return 'dropped';
  }

  // 2. Dedupe by activity.id (checked+marked before the first await — collapses
  //    Bot Framework retries to one route).
  if (deps.dedupe.get(activity.activityId)) {
    log(`duplicate activity ${activity.activityId} — already routed, skipping.`);
    return 'duplicate';
  }
  deps.dedupe.set(activity.activityId, true);

  // AAD object id is the canonical Teams user id; fall back to the BF account id
  // if absent (defensive — a sender without an aadObjectId can never match the
  // tenant-scoped operator whitelist, so it stays external = fail-closed).
  const platformUserId = activity.aadObjectId ?? activity.fromId;
  const displayName = activity.fromName || platformUserId;

  // 3. Resolve sender against the Maestro user directory (auto-create on miss).
  //    Pass the activity tenant so an auto-created teams mapping is tenant-bound
  //    (enables tenant-scoped directory-operator trust on later resolves).
  const resolvedUser = await deps.userResolver.resolve(platformUserId, displayName, activity.tenantId, bot.slug);

  // 4. Tenant-scoped trust (user directory preferred; legacy env fallback requires
  //    a (tenantId, aadObjectId) match; unknown/missing tenant fails closed).
  const trust = resolveTrust(activity.tenantId, platformUserId, deps.operatorAadObjectIds, resolvedUser);

  // 5. Scan + wrap via the SHARED scanner (operator bypasses; external is wrapped).
  const { sanitized, flags } = sanitizeTeamsMessage({
    text: activity.text,
    senderAadObjectId: platformUserId,
    senderDisplayName: displayName,
    senderTenantId: activity.tenantId,
    trustLevel: trust.level,
  });
  if (deps.debug) {
    log(`trust=${trust.level} flags=${flags.length} (${trust.reason})`);
  }

  // 6. Thread heuristics (personal scope): new if no recent thread or stale > 30 min.
  const recent = deps.threadStore.findRecentByConversation(bot.slug, activity.conversationId);
  const isNewConversation = !recent || now() - recent.createdAt > CONVERSATION_TIMEOUT_MS;

  // 7. Build the typed EnrichedContext envelope (locked contract; no `userId`,
  //    `inReplyTo` omitted when absent, topicHints capped at 3).
  const context: EnrichedContext = {
    sender: {
      platformUserId,
      platform: 'teams',
      displayName: resolvedUser?.displayName ?? displayName,
      handle: displayName,
      ...(resolvedUser && { trustLevel: resolvedUser.trustLevel, role: resolvedUser.role }),
    },
    thread: {
      threadId: activity.conversationId,
      isNewConversation,
      ...(recent && { inReplyTo: recent.ampMessageId }),
    },
    topicHints: extractTopicHints(activity.text),
  };

  // 7.5 Ingest attachments (w3) BEFORE routing — the AMPAttachmentV1[] is cited in
  //     the payload, no bytes on /route. Runs off the ack-fast path (handleInbound
  //     is fire-and-forget). FAIL-OPEN: a failed attachment drops itself, never the
  //     message — an attachment-only message still routes its (empty) text.
  let attachments: AMPAttachmentV1[] | undefined;
  if (activity.attachments?.length) {
    if (!activity.downloadAttachment || !bot.ampApiKey) {
      console.error(
        `[AMP] (${bot.slug}) ${activity.attachments.length} attachment(s) on activity ${activity.activityId} but ` +
          `${!bot.ampApiKey ? 'no AMP api key' : 'no downloader'} — routing text only.`,
      );
    } else {
      const ingest = await ingestAttachments(activity.attachments, {
        maestroUrl: deps.maestroUrl,
        ampApiKey: bot.ampApiKey,
        botSlug: bot.slug,
        policy: deps.attachmentPolicy,
        downloadAttachment: activity.downloadAttachment,
        timeoutMs: deps.timeoutMs,
      });
      if (ingest.attachments.length > 0) attachments = ingest.attachments;
      if (ingest.failed > 0 || ingest.skipped > 0) {
        log(`attachments: ${ingest.attachments.length} routed, ${ingest.failed} failed, ${ingest.skipped} skipped.`);
      }
    }
  }

  // 8. Target = this bot's default agent (the bot IS the agent selector; @AIM
  //    cross-targeting is deferred for v1).
  const routeRequest: AMPRouteRequest = {
    to: bot.defaultAgent,
    subject: `Teams message from ${displayName}`,
    priority: 'normal',
    payload: { type: 'request', message: sanitized, context, ...(attachments && { attachments }) },
  };

  // 9. Route async. A post-200 failure is invisible to Teams -> log loudly.
  if (!bot.ampApiKey) {
    console.error(`[AMP] (${bot.slug}) no AMP api key (dry-run / unregistered) — cannot route activity ${activity.activityId}.`);
    return 'failed';
  }

  let result: AMPRouteResponse;
  try {
    const response = await fetch(`${deps.maestroUrl}/api/v1/route`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${bot.ampApiKey}`,
      },
      body: JSON.stringify(routeRequest),
      signal: AbortSignal.timeout(deps.timeoutMs),
    });
    if (!response.ok) {
      const body = await response.text().catch(() => '');
      console.error(`[AMP] (${bot.slug}) route failed (${response.status}) for activity ${activity.activityId}: ${body}`);
      return 'failed';
    }
    result = (await response.json()) as AMPRouteResponse;
  } catch (error) {
    console.error(`[AMP] (${bot.slug}) route request failed for activity ${activity.activityId}:`, (error as Error).message);
    return 'failed';
  }

  // 10. Persist the (botSlug, AMP message id) -> ConversationReference mapping so
  //     the agent's reply can be posted back under this bot (consumed in Phase 3).
  deps.threadStore.record({
    botSlug: bot.slug,
    conversationId: activity.conversationId,
    ampMessageId: result.id,
    // TARGET user for a future proactive DM — the same canonical id used as the
    // directory key above (aadObjectId, BF account-id fallback), so a Maestro DM's
    // `platformUserId` resolves this entry via the by-user index.
    aadObjectId: platformUserId,
    context: {
      reference: activity.reference,
      rootActivityId: activity.activityId,
      tenantId: activity.tenantId ?? '',
    },
    createdAt: now(),
  });

  log(`routed activity ${activity.activityId} -> ${bot.defaultAgent} (amp ${result.id}, trust=${trust.level}, new=${isNewConversation}).`);
  return 'routed';
}
