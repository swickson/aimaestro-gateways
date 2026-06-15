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

// Maestro `/api/v1/route` 400s on an empty `payload.message` ("must have type and
// message fields"). An attachment-only Teams message (a photo with no caption) has
// empty text whether its attachments upload cleanly (Mode 2 success) or all fail —
// either way an empty message means TOTAL loss. We substitute one of these CONSTANT
// gateway placeholders so the agent still sees that something arrived. They are
// injected AFTER the content scanner, so they MUST be constant: no filename, sender,
// or other untrusted text interpolation (that would be an injection vector).
const ATTACHMENT_ONLY_PLACEHOLDER = '[Teams: attachment received — see attachments]';
const ATTACHMENT_FAILED_PLACEHOLDER = '[Teams: an attachment was received but could not be retrieved]';

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
// Conversation scope + stable thread-root derivation (#12)
// ---------------------------------------------------------------------------

export type ConversationScope = 'personal' | 'channel' | 'groupChat';

export interface ThreadIdentity {
  scope: ConversationScope;
  /**
   * Stable per-thread id — the SAME value for the thread ROOT message and every
   * reply in it. Used both as the top-level route `thread_id` (so Maestro memory
   * coheres per-thread) AND as the outbound reply target (so a reply lands in the
   * originating thread, not as a new top-level post).
   */
  stableThreadId: string;
  /**
   * Channel thread-root message id. Omitted when THIS message IS the root (per the
   * locked contract: "omit for channel-root"), and for personal/groupChat (no
   * sub-threading there).
   */
  threadRootId?: string;
}

const CHANNEL_THREAD_MARKER = ';messageid=';

/**
 * Classify a Teams `conversationType` and derive the stable thread-root identity.
 *
 * Teams channel conversation ids look like `19:<ch>@thread.tacv2` for a thread-ROOT
 * message and `19:<ch>@thread.tacv2;messageid=<rootId>` for replies within that
 * thread. We parse the `;messageid=` suffix and rebuild a canonical
 * `<base>;messageid=<rootId>` so the root and all its replies collapse to ONE
 * stable id (for the root we synthesize the suffix from its own activity id). A
 * group chat has no sub-threading, so the whole chat is one stable thread.
 *
 * D2 (greenlit): the suffix parse is defensive (fallback = this activity's id); the
 * exact live channel-ref shape is a deploy-time watch item, so this is unit-tested
 * against mock refs and the derived id is logged for the live pass to eyeball.
 *
 * NOTE: only called for already-accepted scopes (the caller drops unknown
 * `conversationType`s first), so the `personal` return is reached only for a true
 * 1:1 — its `stableThreadId` is unused (personal omits the top-level `thread_id`).
 */
export function deriveThreadIdentity(
  conversationType: string,
  conversationId: string,
  activityId: string,
): ThreadIdentity {
  if (conversationType === 'channel') {
    const idx = conversationId.indexOf(CHANNEL_THREAD_MARKER);
    const suffix = idx >= 0 ? conversationId.slice(idx + CHANNEL_THREAD_MARKER.length) : '';
    const base = idx >= 0 ? conversationId.slice(0, idx) : conversationId;
    if (suffix !== '') {
      // Reply within an existing thread — root id is the parsed suffix.
      return { scope: 'channel', stableThreadId: `${base}${CHANNEL_THREAD_MARKER}${suffix}`, threadRootId: suffix };
    }
    // Thread-ROOT message — synthesize the suffix from this activity's id; omit
    // threadRootId so the root message carries no `room.threadRootId`.
    return { scope: 'channel', stableThreadId: `${base}${CHANNEL_THREAD_MARKER}${activityId}` };
  }
  if (conversationType === 'groupChat') {
    return { scope: 'groupChat', stableThreadId: conversationId };
  }
  return { scope: 'personal', stableThreadId: conversationId };
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
  /**
   * True when THIS bot was @mentioned in the activity (#12). Computed SDK-side in
   * `server.ts` from the mention entities against the bot's recipient id. Drives
   * the channel/groupChat mention gate; ignored for personal (DMs are implicitly
   * addressed). Absent/false = not mentioned.
   */
  mentionsBot?: boolean;
  /** Teams team (group) id for a channel message (#12); absent for personal/groupChat. */
  teamId?: string;
  /** Teams channel id for a channel message (#12); absent for personal/groupChat. */
  channelId?: string;
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

  // 1. Scope gate (#12). Accept personal (1:1 DM), channel, and groupChat; drop any
  //    other conversationType. (Replaces the v1 personal-only drop.)
  const conversationType = activity.conversationType;
  const isPersonal = conversationType === 'personal';
  if (conversationType !== 'personal' && conversationType !== 'channel' && conversationType !== 'groupChat') {
    log(`dropping unsupported conversationType '${conversationType}' activity ${activity.activityId}.`);
    return 'dropped';
  }

  // 1b. Mention gate (#12). The NEW security invariant (replacing "drop all
  //     non-personal"): a channel/groupChat message proceeds ONLY when THIS bot is
  //     @mentioned — gated here, BEFORE dedupe / user-resolution / scan. A 1:1 DM is
  //     implicitly addressed, so no mention is required there. From here on every
  //     non-personal message has passed the mention gate and still flows through the
  //     same tenant-scoped trust + scanner path below.
  if (!isPersonal && !activity.mentionsBot) {
    log(`dropping ${conversationType} activity ${activity.activityId} — this bot was not @mentioned.`);
    return 'dropped';
  }

  // 2. Dedupe by activity.id (checked+marked before the first await — collapses
  //    Bot Framework retries to one route).
  if (deps.dedupe.get(activity.activityId)) {
    log(`duplicate activity ${activity.activityId} — already routed, skipping.`);
    return 'duplicate';
  }
  deps.dedupe.set(activity.activityId, true);

  // The PROVEN AAD object id is the ONLY identity allowed to drive trust elevation
  // (#12 security fix). When it is ABSENT (a Bot-Framework-only sender), trust is
  // forced external in resolveTrust below — BEFORE any directory/legacy check — so
  // the BF `fromId` fallback can never match an operator mapping or whitelist. The
  // fallback still drives conversation/threading/display identity, never trust.
  const aadObjectId = activity.aadObjectId;
  const platformUserId = aadObjectId ?? activity.fromId;
  const displayName = activity.fromName || platformUserId;

  // 3. Resolve sender against the Maestro user directory (auto-create on miss).
  //    Pass the activity tenant so an auto-created teams mapping is tenant-bound
  //    (enables tenant-scoped directory-operator trust on later resolves).
  const resolvedUser = await deps.userResolver.resolve(platformUserId, displayName, activity.tenantId, bot.slug);

  // 4. Tenant-scoped trust (user directory preferred; legacy env fallback requires
  //    a (tenantId, aadObjectId) match; unknown/missing tenant fails closed).
  //    Pass the PROVEN aad id (not the fromId fallback): an absent aadObjectId
  //    resolves to external regardless of any directory record the fallback hit.
  const trust = resolveTrust(activity.tenantId, aadObjectId, deps.operatorAadObjectIds, resolvedUser);

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

  // 6. Thread heuristics: new if no recent thread or stale > 30 min.
  const recent = deps.threadStore.findRecentByConversation(bot.slug, activity.conversationId);
  const isNewConversation = !recent || now() - recent.createdAt > CONVERSATION_TIMEOUT_MS;

  // 6b. Stable thread-root identity (#12). For channel/groupChat this is the memory
  //     key (top-level thread_id) AND the outbound reply target; personal omits it.
  const identity = deriveThreadIdentity(conversationType, activity.conversationId, activity.activityId);

  // 7. Build the typed EnrichedContext envelope (locked contract; no `userId`,
  //    `inReplyTo` omitted when absent, topicHints capped at 3). Channel/groupChat
  //    additively carry advisory `room` + per-sender `trust`; the personal envelope
  //    is byte-identical to v1 (D1 greenlit).
  const context: EnrichedContext = {
    sender: {
      platformUserId,
      platform: 'teams',
      displayName: resolvedUser?.displayName ?? displayName,
      handle: displayName,
      ...(resolvedUser && { trustLevel: resolvedUser.trustLevel, role: resolvedUser.role }),
      // Advisory gate decision (operator|external), per-participant. Non-personal only.
      ...(!isPersonal && { trust: trust.level }),
    },
    thread: {
      // Non-personal threads on the stable root id (cohere replies); personal keeps
      // the 1:1 conversation id unchanged.
      threadId: isPersonal ? activity.conversationId : identity.stableThreadId,
      isNewConversation,
      ...(recent && { inReplyTo: recent.ampMessageId }),
    },
    topicHints: extractTopicHints(activity.text),
    // Advisory room descriptor — channel/groupChat only; personal omits it entirely.
    ...(!isPersonal && {
      room: {
        scope: identity.scope,
        ...(activity.teamId && { teamId: activity.teamId }),
        ...(activity.channelId && { channelId: activity.channelId }),
        ...(identity.threadRootId && { threadRootId: identity.threadRootId }),
      },
    }),
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

  // 7.6 Empty-message guard: never POST an empty `payload.message` (Maestro 400 =
  //     total loss). Fires for an attachment-bearing message whose scanned text is
  //     empty — whether the attachments routed (no-caption photo) or all failed.
  //     Placeholder is a CONSTANT (count-/name-free) injected after the scanner.
  let message = sanitized;
  if (message.trim() === '' && activity.attachments?.length) {
    message = attachments?.length ? ATTACHMENT_ONLY_PLACEHOLDER : ATTACHMENT_FAILED_PLACEHOLDER;
    log(`empty text with ${activity.attachments.length} attachment(s) — substituting placeholder (${attachments?.length ?? 0} routed).`);
  }

  // 8. Target = this bot's default agent (the bot IS the agent selector; @AIM
  //    cross-targeting is deferred for v1).
  const routeRequest: AMPRouteRequest = {
    to: bot.defaultAgent,
    subject: `Teams message from ${displayName}`,
    priority: 'normal',
    // Top-level thread_id (#12): the stable channel/groupChat thread-root, so Maestro
    // memory coheres per-thread. Personal omits it (unchanged — threads via in_reply_to).
    ...(!isPersonal && { thread_id: identity.stableThreadId }),
    payload: { type: 'request', message, context, ...(attachments && { attachments }) },
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
      // #12: where outbound posts the reply so it threads into the originating
      // conversation. Channel/groupChat -> the stable thread-root; personal omits it
      // (outbound falls back to reference.conversation.id = unchanged 1:1 behavior).
      ...(!isPersonal && { replyConversationId: identity.stableThreadId }),
    },
    createdAt: now(),
  });

  // Log the derived stableThreadId for non-personal (D2: deploy-time live pass eyeballs it).
  const threadLog = isPersonal ? '' : `, scope=${identity.scope}, thread=${identity.stableThreadId}`;
  log(`routed activity ${activity.activityId} -> ${bot.defaultAgent} (amp ${result.id}, trust=${trust.level}, new=${isNewConversation}${threadLog}).`);
  return 'routed';
}
