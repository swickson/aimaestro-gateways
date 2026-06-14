/**
 * Teams gateway — outbound response poller (Phase 3).
 *
 * Polls each bot's AMP filesystem inbox for agent replies and posts them back to
 * the originating Teams conversation under that SAME bot's identity, via the
 * stored full `ConversationReference` (proactive `App.send`).
 *
 * IDENTITY IS INBOX-AUTHORITATIVE (red-team finding C — load-bearing): the bot a
 * reply posts under is decided ONLY by which bot's inbox the file arrived in —
 * never by an agent-echoed `payload.context`. Unlike the slack/discord pollers
 * (single identity, so they can trust `payload.context.<platform>` first), Teams
 * runs N bots through one gateway; trusting agent-supplied routing would let a
 * reply post under the wrong bot. So there is NO `payload.context` precedence
 * here: the only lookup is `threadStore.findByAmpMessageId(bot.slug, in_reply_to)`,
 * scoped to the inbox's own bot.
 *
 * DELIVERY DECISIONS:
 *   - No `in_reply_to`, or no `(bot.slug, in_reply_to)` mapping -> SKIP + LEAVE the
 *     file for a later retry (the mapping may be mid-restore). Never delete an
 *     agent's reply (accepted v1 risk: undeliverable files accumulate -> Phase-6
 *     dead-letter follow-up).
 *   - Empty / whitespace-only reply -> nothing to post; delete the file (it is not
 *     "undeliverable agent data", it is no data) so it does not re-poll forever.
 *   - Successful send -> delete the file (mirrors siblings). A crash AFTER send but
 *     BEFORE delete can double-post (no outbound dedupe) — sibling parity, v1-accepted.
 *
 * SDK-DECOUPLED: this module never imports `@microsoft/teams.*`. `server.ts`
 * supplies each bot's `send` (the proactive `App.send`) and its configured
 * `serviceUrl`, so the poller is unit-testable with a mocked sender + mock inbox
 * files (PLAN constraint: "continueConversation with a mocked adapter").
 */

import * as fs from 'node:fs';
import * as path from 'node:path';
import type { AMPAttachmentV1, AMPMessage } from './types.js';
import type { ThreadStore } from './thread-store.js';
import { formatReply } from './format.js';

/** Per-attachment HTTP pull timeout (signed-url GET). */
const ATTACHMENT_PULL_TIMEOUT_MS = 20_000;

/** Bytes pulled from a signed download url, ready for the Teams send (w3). */
export interface OutboundAttachment {
  filename: string;
  contentType: string;
  bytes: Uint8Array;
}

/** One bot's outbound surface: its inbox + a proactive sender bound to its identity. */
export interface OutboundBot {
  slug: string;
  /** Absolute path to this bot's AMP inbox dir (sender-nested `*.json` files). */
  inboxDir: string;
  /**
   * The bot adapter's configured `serviceUrl`. Used ONLY to log a Fork-O1
   * observability warning when it diverges from the stored reference's serviceUrl
   * (regional/GCC routing) — `App.send` rebuilds the ref from this value, so a
   * silent delivery miss there would otherwise be a black hole.
   */
  configuredServiceUrl?: string;
  /**
   * Post a chunk proactively under THIS bot, into `conversationId`. `attachments`
   * (w3) ride a final attachment-carrying activity; empty/omitted = text-only.
   */
  send(conversationId: string, text: string, markdown: boolean, attachments?: OutboundAttachment[]): Promise<void>;
}

export interface OutboundDeps {
  bots: OutboundBot[];
  threadStore: ThreadStore;
  pollIntervalMs: number;
  /** Render replies as markdown (default) vs plain text. */
  markdownDefault: boolean;
  debug: boolean;
}

/**
 * Pull bytes for each cited attachment from its HMAC-signed download url (the url
 * IS the auth — NO Bearer). Returns only the successfully-pulled attachments; a
 * failed pull is logged and dropped (the caller decides leave-for-retry vs deliver
 * text-only). Never throws.
 */
async function pullOutboundAttachments(
  slug: string,
  declared: AMPAttachmentV1[],
  fileName: string,
): Promise<OutboundAttachment[]> {
  const out: OutboundAttachment[] = [];
  for (const att of declared) {
    try {
      if (!att.url) throw new Error('descriptor has no signed url');
      const res = await fetch(att.url, { signal: AbortSignal.timeout(ATTACHMENT_PULL_TIMEOUT_MS) });
      if (!res.ok) throw new Error(`download ${res.status}`);
      out.push({
        filename: att.filename,
        contentType: att.content_type,
        bytes: new Uint8Array(await res.arrayBuffer()),
      });
    } catch (e) {
      console.error(`[OUTBOUND] (${slug}) failed to pull attachment '${att.filename}' (${att.id}) for ${fileName}: ${(e as Error).message}`);
    }
  }
  return out;
}

export function startOutboundPoller(deps: OutboundDeps): () => void {
  let isPolling = false;
  let pollTimeoutId: NodeJS.Timeout | null = null;
  // Log-on-transition state for undeliverable replies. An unmapped file sits in
  // the inbox and re-polls every tick (the mapping may be mid-restore); without
  // throttling it would log "no thread mapping" on EVERY tick → console spam as
  // files pile up. Keyed by full file path (basenames collide across sender
  // dirs). Warn once on first undeliverable; cleared when the file is delivered
  // /removed, and pruned (below) for files that vanish from disk so the set can
  // never grow unbounded.
  const warnedUndeliverable = new Set<string>();
  // Every file path touched in the current scan — drives the prune above.
  const seenThisScan = new Set<string>();

  function debug(message: string, ...args: unknown[]): void {
    if (deps.debug) console.log(`[DEBUG] [OUTBOUND] ${message}`, ...args);
  }

  async function processMessageFile(bot: OutboundBot, filePath: string): Promise<boolean> {
    seenThisScan.add(filePath);
    let msg: AMPMessage;
    try {
      msg = JSON.parse(fs.readFileSync(filePath, 'utf-8')) as AMPMessage;
    } catch (err) {
      console.error(`[OUTBOUND] (${bot.slug}) failed to read/parse ${path.basename(filePath)}:`, (err as Error).message);
      return false;
    }

    try {
      const inReplyTo = msg.envelope?.in_reply_to;
      if (!inReplyTo) {
        console.log(`[OUTBOUND] (${bot.slug}) ${path.basename(filePath)} has no in_reply_to — cannot resolve a conversation, leaving.`);
        return false;
      }

      // INBOX-AUTHORITATIVE: scope strictly to this inbox's bot. No payload.context.
      const entry = deps.threadStore.findByAmpMessageId(bot.slug, inReplyTo);
      if (!entry) {
        // Log-on-transition: warn only the first time this file is undeliverable,
        // not on every poll tick. The file is left for retry regardless.
        if (!warnedUndeliverable.has(filePath)) {
          console.log(`[OUTBOUND] (${bot.slug}) no thread mapping for in_reply_to=${inReplyTo} (evicted/expired/unknown) — undeliverable, leaving for retry.`);
          warnedUndeliverable.add(filePath);
        }
        return false;
      }

      const reference = entry.context.reference;
      const conversationId = reference?.conversation?.id;
      if (!conversationId) {
        console.error(`[OUTBOUND] (${bot.slug}) stored reference for in_reply_to=${inReplyTo} has no conversation id — cannot deliver, leaving.`);
        return false;
      }

      // Fork-O1 observability: App.send uses the bot's configured serviceUrl, not
      // the stored one. If they diverge, a regional/GCC reply could silently miss.
      if (
        bot.configuredServiceUrl &&
        reference.serviceUrl &&
        reference.serviceUrl !== bot.configuredServiceUrl
      ) {
        console.warn(`[TEAMS] (${bot.slug}) serviceUrl mismatch — stored=${reference.serviceUrl} configured=${bot.configuredServiceUrl}; sending via configured (Fork-O1 v1).`);
      }

      const displayName = msg.envelope?.from?.split('@')[0] || 'Agent';
      const rawMessage = msg.payload?.message;
      const responseText =
        typeof rawMessage === 'string' ? rawMessage : rawMessage ? JSON.stringify(rawMessage) : '';

      const { chunks, markdown } = formatReply({
        displayName,
        message: responseText,
        markdown: deps.markdownDefault,
      });

      // w3 attachments: pull bytes for any cited descriptors (signed url = auth).
      const declared = Array.isArray(msg.payload?.attachments) ? msg.payload.attachments : [];
      const attachments = declared.length > 0 ? await pullOutboundAttachments(bot.slug, declared, path.basename(filePath)) : [];

      // The outbound.ts:139 fix: "nothing to post" is now text AND attachments empty.
      // An attachment-only reply (empty text + attachments) must DELIVER, not delete.
      if (chunks.length === 0 && declared.length === 0) {
        console.log(`[OUTBOUND] (${bot.slug}) ${path.basename(filePath)} is an empty reply — nothing to post, deleting.`);
        fs.unlinkSync(filePath);
        warnedUndeliverable.delete(filePath);
        return true;
      }

      // Attachment-ONLY reply whose every pull failed: never drop agent data — leave
      // for retry (warn-once, mirrors the no-mapping path) rather than delete.
      if (chunks.length === 0 && attachments.length === 0) {
        if (!warnedUndeliverable.has(filePath)) {
          console.error(`[OUTBOUND] (${bot.slug}) ${path.basename(filePath)} is attachment-only but no attachment could be pulled — leaving for retry.`);
          warnedUndeliverable.add(filePath);
        }
        return false;
      }

      for (const chunk of chunks) {
        await bot.send(conversationId, chunk, markdown);
      }
      // Attachments ride a final attachment-carrying activity (separate bubble; the
      // outbound App.send attachment shape is the live-Azure watch item).
      if (attachments.length > 0) {
        await bot.send(conversationId, '', markdown, attachments);
      }

      // Text delivered but some/all attachments failed to pull: never block text —
      // log loud (the lost attachment is accepted per the item-8 has-text policy).
      if (declared.length > attachments.length) {
        console.error(`[OUTBOUND] (${bot.slug}) delivered text + ${attachments.length}/${declared.length} attachment(s) for ${path.basename(filePath)}; ${declared.length - attachments.length} could not be pulled.`);
      }

      console.log(`[-> Teams] (${bot.slug}) reply from ${displayName} -> conversation ${conversationId} (${chunks.length} chunk(s), ${attachments.length} attachment(s)).`);

      fs.unlinkSync(filePath);
      warnedUndeliverable.delete(filePath);
      debug(`(${bot.slug}) deleted processed reply ${filePath}`);
      return true;
    } catch (err) {
      console.error(`[OUTBOUND] (${bot.slug}) failed to deliver ${path.basename(filePath)}:`, (err as Error).message);
      return false;
    }
  }

  async function scanBotInbox(bot: OutboundBot): Promise<void> {
    if (!fs.existsSync(bot.inboxDir)) {
      debug(`(${bot.slug}) inbox does not exist yet: ${bot.inboxDir}`);
      return;
    }

    let entries: fs.Dirent[];
    try {
      entries = fs.readdirSync(bot.inboxDir, { withFileTypes: true });
    } catch (err) {
      debug(`(${bot.slug}) inbox read error:`, (err as Error).message);
      return;
    }

    for (const dirent of entries) {
      if (!dirent.isDirectory()) continue; // inbox is sender-nested: inbox/{sender}/{id}.json
      const senderDir = path.join(bot.inboxDir, dirent.name);

      let files: string[];
      try {
        files = fs.readdirSync(senderDir).filter((f) => f.endsWith('.json'));
      } catch {
        continue;
      }

      for (const file of files) {
        await processMessageFile(bot, path.join(senderDir, file));
      }

      // Tidy empty sender dirs (best-effort).
      try {
        if (fs.readdirSync(senderDir).length === 0) {
          fs.rmdirSync(senderDir);
          debug(`(${bot.slug}) cleaned empty sender dir: ${dirent.name}`);
        }
      } catch {
        /* ignore cleanup races */
      }
    }
  }

  async function scanAll(): Promise<void> {
    if (isPolling) return; // reentrancy guard — a slow tick must not overlap the next.
    isPolling = true;
    seenThisScan.clear();
    try {
      for (const bot of deps.bots) {
        await scanBotInbox(bot);
      }
      // Prune warn-state for files that no longer exist on disk (delivered, or
      // removed out-of-band) so a stuck-file warned-set can't grow unbounded.
      for (const filePath of warnedUndeliverable) {
        if (!seenThisScan.has(filePath)) warnedUndeliverable.delete(filePath);
      }
    } finally {
      isPolling = false;
    }
  }

  const poll = async (): Promise<void> => {
    await scanAll();
    pollTimeoutId = setTimeout(() => void poll(), deps.pollIntervalMs);
  };

  void poll();
  console.log(`[OUTBOUND] Filesystem polling started at ${deps.pollIntervalMs}ms across ${deps.bots.length} bot inbox(es).`);
  for (const bot of deps.bots) {
    console.log(`[OUTBOUND]   ${bot.slug}: ${bot.inboxDir}`);
  }

  return () => {
    if (pollTimeoutId) {
      clearTimeout(pollTimeoutId);
      pollTimeoutId = null;
    }
    console.log('[OUTBOUND] Poller stopped.');
  };
}
