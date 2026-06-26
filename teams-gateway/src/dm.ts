/**
 * Teams gateway — proactive DM delivery (Phase 5).
 *
 * Maestro's `notifyUser` POSTs `/api/gateway/dm` to push an agent message to a
 * user who has NO open reply thread (`in_reply_to` is absent). v1 is
 * CAPTURE-ON-FIRST-CONTACT: we can only reach a user the gateway has already seen
 * inbound (their `ConversationReference` is in the thread-store, indexed by user).
 * Cold-start `createConversation` is feature-flagged OFF by default. When enabled,
 * it only works for directory-known users where Maestro supplies the tenantId +
 * botSlug it learned from prior inbound contact with some Teams bot. A truly
 * never-seen user is still unresolvable — no Graph lookup or operator-supplied
 * identity path lives here.
 *
 * SEND-BOT RESOLUTION (contract, locked with Watson): `body.botSlug ?? by-user
 * single-bot reuse ?? 409`. The caller may pin a bot; otherwise a single live bot
 * can be reused. Multi-bot users must pin `botSlug`; the gateway will not guess via
 * recency because that can send under the wrong bot identity. A pinned bot that is
 * not in the registry is a 400; a pinned bot the user never messaged is a 409.
 *
 * AUTH BOUNDARY (§0.2): this router is mounted at `/api/gateway` and gated as a
 * GROUP by `createAuthMiddleware(ADMIN_TOKEN)` — a targeted gate, NOT a blanket
 * `/api` gate, so it never touches the bot `/api/<slug>/messages` JWT paths.
 *
 * SDK-DECOUPLED (mirrors inbound/outbound): the delivery core `deliverDm` takes a
 * `sendChunk` closure (server.ts binds it to the originating bot's `App.send`) and
 * never imports `@microsoft/teams.*`, so it is unit-testable with a capturing mock
 * sender + an in-memory thread-store — no live Azure.
 */

import express, { type Router } from 'express';
import { randomUUID } from 'node:crypto';
import { createAuthMiddleware } from '@aimaestro/common/auth.js';
import { chunkText, TEAMS_MAX_LENGTH } from './format.js';
import type { ThreadStore } from './thread-store.js';
import type { ThreadContext } from './types.js';

export interface CreateColdStartConversationInput {
  botSlug: string;
  tenantId: string;
  aadObjectId: string;
}

export interface CreateColdStartConversationResult {
  conversationId: string;
  reference: ThreadContext['reference'];
  rootActivityId: string;
}

export interface DmDeps {
  threadStore: ThreadStore;
  /** Registry allowlist — a caller-pinned botSlug must be one of these (§0.2). */
  knownBots: Set<string>;
  /** Post one chunk under `botSlug` into `conversationId` (bound to that bot's App.send). */
  sendChunk(botSlug: string, conversationId: string, text: string, markdown: boolean): Promise<void>;
  /**
   * Ensure a personal 1:1 conversation exists and return its id + reference WITHOUT
   * posting any inline activity (an existing 1:1 would silently drop it — #25).
   * All chunks are delivered separately via `sendChunk`. Bound in server.ts to the
   * per-bot SDK App; mocked in unit tests.
   */
  createColdStartConversation?: (input: CreateColdStartConversationInput) => Promise<CreateColdStartConversationResult>;
  /** Feature flag: default false preserves the no_prior_contact 409 contract. */
  coldStartEnabled: boolean;
  /** Render markdown (default) vs plain text. */
  markdownDefault: boolean;
}

export interface DmResult {
  status: number;
  json: Record<string, unknown>;
}

function badRequest(detail: string): DmResult {
  return { status: 400, json: { error: 'bad_request', detail } };
}

function coldStartError(err: unknown): DmResult {
  // Null-guard the rejection (#18): a null/undefined throw must fall through to the
  // default 502 mapping, never throw again inside the mapper and surface as a 500.
  const e = (err ?? {}) as {
    status?: unknown;
    statusCode?: unknown;
    code?: unknown;
    response?: { status?: unknown };
  };
  const status =
    typeof e.status === 'number'
      ? e.status
      : typeof e.statusCode === 'number'
        ? e.statusCode
        : typeof e.response?.status === 'number'
          ? e.response.status
          : undefined;
  const code = typeof e.code === 'string' ? e.code : '';

  if (status === 401 || status === 403) {
    return { status: 409, json: { error: 'undeliverable', reason: 'bot_not_installed_or_forbidden' } };
  }
  if (status === 404) {
    return { status: 409, json: { error: 'undeliverable', reason: 'user_or_tenant_not_found' } };
  }
  if (status === 400 || code === 'wrong_tenant') {
    return { status: 409, json: { error: 'undeliverable', reason: 'wrong_tenant_or_unreachable' } };
  }
  return { status: 502, json: { error: 'undeliverable', reason: 'cold_start_failed' } };
}

/**
 * Pure delivery core. Validates the body, resolves the send-bot + conversation,
 * chunks the message (subject prepended as a bold leading line when present), and
 * sends each chunk via `sendChunk`. Returns the HTTP status + JSON the route emits.
 * Throws only if `sendChunk` rejects (the route maps that to a 500).
 */
export async function deliverDm(deps: DmDeps, body: unknown): Promise<DmResult> {
  if (typeof body !== 'object' || body === null) {
    return badRequest('request body must be a JSON object');
  }
  const b = body as Record<string, unknown>;

  const platformUserId = typeof b.platformUserId === 'string' ? b.platformUserId.trim() : '';
  const message = typeof b.message === 'string' ? b.message : '';
  const subject = typeof b.subject === 'string' ? b.subject : undefined;
  const tenantId = typeof b.tenantId === 'string' ? b.tenantId.trim() : '';

  // botSlug is optional, but if present it must be a non-empty string.
  let botSlug: string | undefined;
  if (b.botSlug !== undefined) {
    if (typeof b.botSlug !== 'string' || b.botSlug.trim() === '') {
      return badRequest('botSlug, when provided, must be a non-empty string');
    }
    botSlug = b.botSlug.trim();
  }

  if (!platformUserId) return badRequest('platformUserId is required');
  if (message.trim() === '') return badRequest('message is required and cannot be empty');

  // Resolve the conversation: pinned bot -> (user, bot) lookup; else single-bot reuse.
  let entry;
  if (botSlug) {
    if (!deps.knownBots.has(botSlug)) {
      return badRequest(`botSlug '${botSlug}' is not a registered bot`);
    }
    entry = deps.threadStore.findByUserAndBot(platformUserId, botSlug);
  } else {
    const candidateBots = deps.threadStore.distinctBotsForUser(platformUserId);
    if (candidateBots.length > 1) {
      console.log(
        `[TEAMS] /api/gateway/dm ambiguous — platformUserId=${platformUserId} has multiple bot mappings; ` +
          'caller must pin botSlug.',
      );
      return {
        status: 409,
        json: {
          error: 'undeliverable',
          reason: 'ambiguous_bot',
          note: 'user has conversations with multiple bots; pin botSlug',
          candidates: candidateBots,
        },
      };
    }
    entry = deps.threadStore.findLatestByUser(platformUserId);
  }

  if (!entry) {
    if (!deps.coldStartEnabled) {
      console.log(
        `[TEAMS] /api/gateway/dm undeliverable — no prior contact for platformUserId=${platformUserId}` +
          `${botSlug ? ` bot=${botSlug}` : ''} (cold-start disabled).`,
      );
      return {
        status: 409,
        json: { error: 'undeliverable', reason: 'no_prior_contact', note: 'cold-start conversation creation is disabled' },
      };
    }
    if (!botSlug) {
      console.log(`[TEAMS] /api/gateway/dm undeliverable — no prior contact and no botSlug for platformUserId=${platformUserId}.`);
      return { status: 409, json: { error: 'undeliverable', reason: 'no_send_bot' } };
    }
    if (!tenantId) {
      return badRequest('tenantId is required for cold-start DM delivery');
    }
    if (!deps.createColdStartConversation) {
      console.error('[TEAMS] /api/gateway/dm cold-start enabled but no createConversation dependency is configured.');
      return { status: 502, json: { error: 'undeliverable', reason: 'cold_start_unavailable' } };
    }

    const prefix = subject ? (deps.markdownDefault ? `**${subject}**\n\n` : `${subject}\n\n`) : '';
    const chunks = chunkText(prefix + message, TEAMS_MAX_LENGTH);
    try {
      // Ensure the 1:1 exists (idempotent, NO inline activity — so an EXISTING
      // conversation can no longer silently drop chunk[0]; #25), then deliver EVERY
      // chunk via the proven warm send leg (uniform for new + existing 1:1s).
      const created = await deps.createColdStartConversation({
        botSlug,
        tenantId,
        aadObjectId: platformUserId,
      });
      for (const chunk of chunks) {
        await deps.sendChunk(botSlug, created.conversationId, chunk, deps.markdownDefault);
      }
      // Persist the reference only AFTER every chunk is delivered (#17): a
      // mid-sequence failure throws above, leaving NO record, so a caller retry
      // re-runs this idempotent cold path rather than a warm-path re-send of chunk[0].
      deps.threadStore.record({
        botSlug,
        conversationId: created.conversationId,
        ampMessageId: `dm:${botSlug}:${platformUserId}:${randomUUID()}`,
        aadObjectId: platformUserId,
        context: {
          reference: created.reference,
          rootActivityId: created.rootActivityId,
          tenantId,
        },
        createdAt: Date.now(),
      });
      console.log(`[TEAMS] /api/gateway/dm cold-start delivered to ${platformUserId} via ${botSlug} (${chunks.length} chunk(s)).`);
      return { status: 200, json: { delivered: true, botSlug, chunks: chunks.length, coldStart: true } };
    } catch (err) {
      const mapped = coldStartError(err);
      // Null-safe (#18): a null/undefined rejection must not throw here either.
      const detail = err instanceof Error ? err.message : String(err);
      console.error(
        `[TEAMS] /api/gateway/dm cold-start failed for platformUserId=${platformUserId} bot=${botSlug} ` +
          `tenant=${tenantId}: ${detail}`,
      );
      return mapped;
    }
  }

  const sendBot = entry.botSlug;
  // The resolved bot was a live bot at record time, but it may not be in THIS
  // process's registry (config changed across a restart) — guard before sending.
  if (!deps.knownBots.has(sendBot)) {
    console.error(`[TEAMS] /api/gateway/dm — resolved bot '${sendBot}' is not currently registered; cannot deliver.`);
    return { status: 409, json: { error: 'undeliverable', reason: 'bot_unavailable' } };
  }

  const prefix = subject ? (deps.markdownDefault ? `**${subject}**\n\n` : `${subject}\n\n`) : '';
  const chunks = chunkText(prefix + message, TEAMS_MAX_LENGTH);
  for (const chunk of chunks) {
    await deps.sendChunk(sendBot, entry.conversationId, chunk, deps.markdownDefault);
  }

  console.log(`[TEAMS] /api/gateway/dm delivered to ${platformUserId} via ${sendBot} (${chunks.length} chunk(s)).`);
  return { status: 200, json: { delivered: true, botSlug: sendBot, chunks: chunks.length } };
}

/**
 * Build the `/api/gateway` sub-router: JSON body + ADMIN_TOKEN group auth + the
 * `POST /dm` handler. Fail-closed by construction (createAuthMiddleware throws if
 * ADMIN_TOKEN is blank). A `sendChunk` rejection (e.g. App.send network failure)
 * is logged and surfaced as 500 — never a silent drop.
 */
export function createDmRouter(deps: DmDeps & { adminToken: string }): Router {
  const router = express.Router();
  router.use(express.json());
  router.use(createAuthMiddleware(deps.adminToken));

  router.post('/dm', (req, res) => {
    void deliverDm(deps, req.body)
      .then((result) => {
        res.status(result.status).json(result.json);
      })
      .catch((err) => {
        console.error('[TEAMS] /api/gateway/dm delivery error:', (err as Error).message);
        res.status(500).json({ error: 'delivery_failed' });
      });
  });

  return router;
}
