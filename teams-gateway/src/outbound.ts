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
import type { AMPAttachmentV1, AMPMessage, AttachmentPolicy } from './types.js';
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
  /** Configured Maestro base URL (this gateway's own AMP endpoint). */
  maestroUrl: string;
  /**
   * Trusted-origin allowlist for outbound attachment download urls: the union of
   * this gateway's `maestroUrl` origin and every ENABLED mesh-host origin from
   * `~/.aimaestro/hosts.json` (loaded once at startup). A signed download url is
   * accepted ONLY when its `origin` is in this set — signed urls carry the ORIGIN
   * host's Tailscale origin (`getSelfHost().url`), not necessarily `maestroUrl`.
   * Restricting to known mesh hosts keeps SSRF closed (never an arbitrary origin).
   */
  allowedOrigins: ReadonlySet<string>;
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
  /**
   * Gateway attachment policy (w3 hardening). The OUTBOUND consume path reads
   * `payload.attachments` from the AGENT-controlled inbox JSON, so every cited
   * descriptor is re-validated against this policy (caps, scan_status, url
   * origin+path) BEFORE any byte is pulled — mirroring the inbound enforcement.
   */
  policy: AttachmentPolicy;
  debug: boolean;
}

/**
 * Outcome of pulling one message's cited attachments. The two failure tallies drive
 * the caller's file-lifecycle decision (DROP vs leave-for-retry):
 *   - `validationDrops` — descriptor failed POLICY validation (malformed/hostile:
 *     bad kind, non-routable scan_status, missing id/filename, over-cap or lying
 *     size, off-origin/off-path url, deny-listed type, count over cap). These never
 *     self-heal, so a reply whose attachments are ALL validation-dropped is deleted,
 *     not retried.
 *   - `pullFailures` — descriptor was VALID but the network pull failed (non-2xx,
 *     timeout, redirect rejected, connection error). Possibly transient → the caller
 *     leaves the file for a later retry (never drop legitimate agent data).
 */
interface PullResult {
  pulled: OutboundAttachment[];
  validationDrops: number;
  pullFailures: number;
}

/** A size-cap violation discovered DURING the pull (lying/oversize body) — a DROP, not a retry. */
class AttachmentOverCapError extends Error {}

/**
 * Validate ONE agent-supplied descriptor against the gateway policy. Returns a
 * human-readable reason string when invalid, or `null` when it passes. The input
 * is typed `unknown`: `payload.attachments` is agent-controlled JSON, so an element
 * can be `null`, a string, an array, or any non-object — the static `AMPAttachmentV1`
 * type is NOT a runtime guarantee. This function is TOTAL: it FIRST proves `att` is a
 * non-null, non-array object before reading any field, so it never throws a TypeError
 * on a hostile shape — it always returns a drop reason instead. Every field needed for
 * delivery is then runtime-checked. The url's origin must be in the trusted mesh-host
 * `allowedOrigins` set AND its path EXACTLY `/api/v1/attachments/<id>/download` with
 * `<id>` === the descriptor's own id — closing the "trusted-origin but arbitrary
 * internal path" (e.g. `/api/v1/agents`) SSRF angle on top of the origin allowlist.
 *
 * KIND/DIGEST tolerate the agent `amp-send --attach` WIRE form (Watson-locked): the
 * CLI may omit `kind` (only an explicit `legacy` is rejected — mirrors Maestro
 * `/route`), and a digest may carry an optional `sha256:` prefix. Shape inference
 * requires id + url(+origin+path+id-match) + digest + scan_status + size + filename +
 * content_type — all present/valid — so a non-amp-v1 descriptor still can't slip through.
 */
const SHA256_HEX = /^[0-9a-f]{64}$/i;

/** Validate the digest field, tolerating an optional case-insensitive `sha256:` prefix. */
function isValidDigest(value: unknown): boolean {
  if (typeof value !== 'string') return false;
  const hex = value.replace(/^sha256:/i, '');
  return SHA256_HEX.test(hex);
}

export function validateOutboundDescriptor(
  att: unknown,
  allowedOrigins: ReadonlySet<string>,
  policy: AttachmentPolicy,
): string | null {
  // FIRST guard — prove `att` is a non-null, non-array object before any field
  // dereference. Without this, `att.kind` on a `null`/string/array element throws a
  // TypeError that escapes the pull loop, breaks the never-throws contract, and leaves
  // a hostile descriptor spinning on retry instead of policy-dropping.
  if (att === null || typeof att !== 'object' || Array.isArray(att)) {
    const got = att === null ? 'null' : Array.isArray(att) ? 'array' : typeof att;
    return `descriptor is not an object (got ${got})`;
  }
  const d = att as Record<string, unknown>;
  // KIND: reject ONLY an explicit `legacy`; absent/amp-v1/other infer amp-v1 by shape
  // (the CLI wire form omits kind; Maestro /route itself only hard-rejects 'legacy').
  if (d.kind === 'legacy') return "kind 'legacy' is not routable";
  if (d.scan_status !== 'clean' && d.scan_status !== 'basic_clean') {
    return `scan_status '${String(d.scan_status)}' is not routable`;
  }
  if (typeof d.id !== 'string' || d.id.trim() === '') return 'missing/empty id';
  if (typeof d.filename !== 'string' || d.filename.trim() === '') return 'missing/empty filename';
  if (typeof d.content_type !== 'string' || d.content_type.trim() === '') {
    return 'missing/empty content_type';
  }
  if (!isValidDigest(d.digest)) return `invalid/missing digest '${String(d.digest)}'`;
  if (typeof d.size !== 'number' || !Number.isFinite(d.size) || d.size <= 0) {
    return `invalid size ${String(d.size)}`;
  }
  if (d.size > policy.maxBytes) return `declared size ${d.size}B exceeds cap ${policy.maxBytes}B`;
  const ct = d.content_type.toLowerCase();
  if (policy.denyContentTypes.some((deny) => ct.includes(deny))) {
    return `content-type '${d.content_type}' is deny-listed`;
  }
  if (typeof d.url !== 'string' || d.url.trim() === '') return 'missing/empty url';
  let url: URL;
  try {
    url = new URL(d.url);
  } catch {
    return 'unparseable url';
  }
  if (!allowedOrigins.has(url.origin)) return `url origin '${url.origin}' is not a trusted mesh host`;
  const expectedPath = `/api/v1/attachments/${d.id}/download`;
  if (url.pathname !== expectedPath) return `url path '${url.pathname}' != '${expectedPath}'`;
  return null;
}

/**
 * Read a response body into bytes WITHOUT unbounded materialization: reject up front
 * if a declared `Content-Length` exceeds the cap, then cap the ACTUAL streamed bytes
 * (defeats an oversize body that omits Content-Length). Throws `AttachmentOverCapError`
 * on a size violation. Replaces a blind `res.arrayBuffer()`.
 */
async function readBoundedBody(res: Response, maxBytes: number): Promise<Uint8Array> {
  const lenHeader = res.headers.get('content-length');
  if (lenHeader !== null) {
    const declaredLen = Number(lenHeader);
    if (Number.isFinite(declaredLen) && declaredLen > maxBytes) {
      throw new AttachmentOverCapError(`Content-Length ${declaredLen}B exceeds cap ${maxBytes}B`);
    }
  }
  const body = res.body;
  if (!body) {
    // No stream available — fall back to arrayBuffer but still enforce the cap.
    const buf = new Uint8Array(await res.arrayBuffer());
    if (buf.byteLength > maxBytes) {
      throw new AttachmentOverCapError(`body ${buf.byteLength}B exceeds cap ${maxBytes}B`);
    }
    return buf;
  }
  const reader = body.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    if (!value) continue;
    total += value.byteLength;
    if (total > maxBytes) {
      await reader.cancel().catch(() => undefined); // stop pulling bytes immediately
      throw new AttachmentOverCapError(`streamed body exceeds cap ${maxBytes}B`);
    }
    chunks.push(value);
  }
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return out;
}

/**
 * Pull bytes for each cited attachment from its HMAC-signed download url (the url
 * IS the auth — NO Bearer). Each descriptor is POLICY-VALIDATED before any byte is
 * fetched; the fetch follows NO redirects (`redirect:'error'`, so a 3xx off the
 * pinned origin can't bypass the origin pin) and the body is BOUNDED at the size
 * cap. Returns the successfully-pulled attachments plus failure tallies; never throws.
 */
async function pullOutboundAttachments(
  bot: OutboundBot,
  declared: unknown[],
  fileName: string,
  policy: AttachmentPolicy,
): Promise<PullResult> {
  const result: PullResult = { pulled: [], validationDrops: 0, pullFailures: 0 };

  // Count cap: never pull more than maxCount; the extras are a policy DROP.
  let toProcess = declared;
  if (declared.length > policy.maxCount) {
    const dropped = declared.length - policy.maxCount;
    console.error(`[OUTBOUND] (${bot.slug}) ${fileName}: ${declared.length} cited attachments exceeds cap ${policy.maxCount} — dropping ${dropped}.`);
    result.validationDrops += dropped;
    toProcess = declared.slice(0, policy.maxCount);
  }

  for (const att of toProcess) {
    const invalid = validateOutboundDescriptor(att, bot.allowedOrigins, policy);
    if (invalid) {
      // Malformed/hostile descriptor — DROP it (won't self-heal). Loud, never silent.
      // `att` may be a non-object here, so extract any filename defensively (no throw).
      const label = String((att as { filename?: unknown } | null | undefined)?.filename);
      result.validationDrops += 1;
      console.error(`[OUTBOUND] (${bot.slug}) ${fileName}: rejecting attachment descriptor '${label}' (${invalid}) — dropping.`);
      continue;
    }
    // Validation proved the runtime shape: a non-null object with non-empty string
    // id/filename/content_type/url and a finite positive size. Safe to narrow from unknown.
    const valid = att as AMPAttachmentV1;
    try {
      const res = await fetch(valid.url, {
        signal: AbortSignal.timeout(ATTACHMENT_PULL_TIMEOUT_MS),
        redirect: 'error', // a 3xx off the pinned origin must NOT be followed (Watson F2)
      });
      if (!res.ok) throw new Error(`download ${res.status}`);
      const bytes = await readBoundedBody(res, policy.maxBytes);
      result.pulled.push({ filename: valid.filename, contentType: valid.content_type, bytes });
    } catch (e) {
      if (e instanceof AttachmentOverCapError) {
        // Body violated the size cap — a lying/hostile descriptor, DROP (retry won't help).
        result.validationDrops += 1;
        console.error(`[OUTBOUND] (${bot.slug}) ${fileName}: attachment '${valid.filename}' (${valid.id}) over size cap — dropping: ${e.message}`);
      } else {
        // Valid descriptor, transient/network pull failure — leave for retry.
        result.pullFailures += 1;
        console.error(`[OUTBOUND] (${bot.slug}) failed to pull attachment '${valid.filename}' (${valid.id}) for ${fileName}: ${(e as Error).message}`);
      }
    }
  }
  return result;
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

      // w3 attachments: VALIDATE each agent-cited descriptor against policy, then pull
      // bytes (signed url = auth; bounded read; no redirects). Validation failures are
      // dropped (hostile/malformed); transient pull failures are tallied for retry.
      const declared = Array.isArray(msg.payload?.attachments) ? msg.payload.attachments : [];
      const pull =
        declared.length > 0
          ? await pullOutboundAttachments(bot, declared, path.basename(filePath), deps.policy)
          : { pulled: [], validationDrops: 0, pullFailures: 0 };
      const attachments = pull.pulled;

      // The outbound.ts:139 fix: "nothing to post" is now text AND attachments empty.
      // An attachment-only reply (empty text + attachments) must DELIVER, not delete.
      if (chunks.length === 0 && declared.length === 0) {
        console.log(`[OUTBOUND] (${bot.slug}) ${path.basename(filePath)} is an empty reply — nothing to post, deleting.`);
        fs.unlinkSync(filePath);
        warnedUndeliverable.delete(filePath);
        return true;
      }

      // Attachment-ONLY reply with nothing delivered. Distinguish WHY (Columbo P1):
      //   - a transient pull failure may self-heal → leave for retry (warn-once),
      //     never drop legitimate agent data;
      //   - but if EVERY descriptor failed policy validation (malformed/hostile) with
      //     no transient failure, retry is pointless → DROP it (delete), don't spin.
      if (chunks.length === 0 && attachments.length === 0) {
        if (pull.pullFailures > 0) {
          if (!warnedUndeliverable.has(filePath)) {
            console.error(`[OUTBOUND] (${bot.slug}) ${path.basename(filePath)} is attachment-only but no attachment could be pulled — leaving for retry.`);
            warnedUndeliverable.add(filePath);
          }
          return false;
        }
        console.error(`[OUTBOUND] (${bot.slug}) ${path.basename(filePath)} is attachment-only but all ${declared.length} cited descriptor(s) were rejected by policy — dropping (not retrying).`);
        fs.unlinkSync(filePath);
        warnedUndeliverable.delete(filePath);
        return true;
      }

      for (const chunk of chunks) {
        await bot.send(conversationId, chunk, markdown);
      }
      // Attachments ride a final attachment-carrying activity (separate bubble; the
      // outbound App.send attachment shape is the live-Azure watch item).
      if (attachments.length > 0) {
        await bot.send(conversationId, '', markdown, attachments);
      }

      // Text delivered but some/all attachments dropped (policy-rejected or pull-failed):
      // never block text — log loud (the lost attachment is accepted per item-8 has-text).
      if (declared.length > attachments.length) {
        console.error(`[OUTBOUND] (${bot.slug}) delivered text + ${attachments.length}/${declared.length} attachment(s) for ${path.basename(filePath)}; ${declared.length - attachments.length} rejected-by-policy or could-not-be-pulled.`);
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
