/**
 * Teams gateway — proactive DM delivery (Phase 5).
 *
 * Maestro's `notifyUser` POSTs `/api/gateway/dm` to push an agent message to a
 * user who has NO open reply thread (`in_reply_to` is absent). v1 is
 * CAPTURE-ON-FIRST-CONTACT: we can only reach a user the gateway has already seen
 * inbound (their `ConversationReference` is in the thread-store, indexed by user).
 * Cold-start `createConversation` (never-contacted user) is v2 — OUT OF SCOPE; an
 * unseen user is a clean 409, not an error.
 *
 * SEND-BOT RESOLUTION (contract, locked with Watson): `body.botSlug ?? by-user
 * last-seen bot ?? 409`. The caller may pin a bot; otherwise the most-recently-
 * inbound bot wins (the `findLatestByUser` recency pointer). A pinned bot that is
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
import { createAuthMiddleware } from '@aimaestro/common/auth.js';
import { chunkText, TEAMS_MAX_LENGTH } from './format.js';
import type { ThreadStore } from './thread-store.js';

export interface DmDeps {
  threadStore: ThreadStore;
  /** Registry allowlist — a caller-pinned botSlug must be one of these (§0.2). */
  knownBots: Set<string>;
  /** Post one chunk under `botSlug` into `conversationId` (bound to that bot's App.send). */
  sendChunk(botSlug: string, conversationId: string, text: string, markdown: boolean): Promise<void>;
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

  // Resolve the conversation: pinned bot -> (user, bot) lookup; else most-recent bot.
  let entry;
  if (botSlug) {
    if (!deps.knownBots.has(botSlug)) {
      return badRequest(`botSlug '${botSlug}' is not a registered bot`);
    }
    entry = deps.threadStore.findByUserAndBot(platformUserId, botSlug);
  } else {
    entry = deps.threadStore.findLatestByUser(platformUserId);
  }

  if (!entry) {
    console.log(
      `[TEAMS] /api/gateway/dm undeliverable — no prior contact for platformUserId=${platformUserId}` +
        `${botSlug ? ` bot=${botSlug}` : ''} (cold-start conversation = v2, out of scope).`,
    );
    return {
      status: 409,
      json: { error: 'undeliverable', reason: 'no_prior_contact', note: 'cold-start conversation creation is v2 (out of scope)' },
    };
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
