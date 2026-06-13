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
import type { AMPMessage } from './types.js';
import type { ThreadStore } from './thread-store.js';
import { formatReply } from './format.js';

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
  /** Post a chunk proactively under THIS bot, into `conversationId`. */
  send(conversationId: string, text: string, markdown: boolean): Promise<void>;
}

export interface OutboundDeps {
  bots: OutboundBot[];
  threadStore: ThreadStore;
  pollIntervalMs: number;
  /** Render replies as markdown (default) vs plain text. */
  markdownDefault: boolean;
  debug: boolean;
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

      if (chunks.length === 0) {
        console.log(`[OUTBOUND] (${bot.slug}) ${path.basename(filePath)} is an empty reply — nothing to post, deleting.`);
        fs.unlinkSync(filePath);
        warnedUndeliverable.delete(filePath);
        return true;
      }

      for (const chunk of chunks) {
        await bot.send(conversationId, chunk, markdown);
      }

      console.log(`[-> Teams] (${bot.slug}) reply from ${displayName} -> conversation ${conversationId} (${chunks.length} chunk(s)).`);

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
