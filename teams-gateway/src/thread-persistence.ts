/**
 * Teams gateway — thread-store persistence (Phase 3).
 *
 * The AMP inbox is durable: an agent reply can land on disk while the gateway is
 * down. It is only deliverable if the `(botSlug, ampMessageId) -> ConversationReference`
 * mapping survived the restart too — otherwise the reply has no conversation to
 * post into. This module bridges that gap by persisting the in-memory
 * `ThreadStore` to a JSON snapshot file and restoring it on boot.
 *
 * Durability shape:
 *   - `restoreThreadStore` on boot — missing/corrupt/wrong-shape file -> empty
 *     store, never a crash (a bad snapshot must not wedge startup).
 *   - `saveThreadStore` is ATOMIC (write temp + rename) so a crash mid-write can
 *     never leave a half-written file that fails the next boot's parse.
 *   - `startSnapshotTimer` runs a periodic save (crash-safety floor); the graceful
 *     SIGTERM/SIGINT path calls `saveThreadStore` once more for a clean final state.
 *
 * The snapshot itself is age-pruned by the store (`snapshot()` drops aged-out
 * entries), so the file never grows unbounded with stale references.
 */

import * as fs from 'node:fs';
import * as path from 'node:path';
import type { ThreadEntry, ThreadStore, ThreadStoreSnapshot } from './thread-store.js';

/**
 * Load a snapshot from disk into the store. Returns the number of entries
 * restored. Any failure (absent file, unreadable, malformed JSON, wrong shape)
 * is logged and treated as "start empty" — never throws.
 */
export function restoreThreadStore(store: ThreadStore, filePath: string): number {
  let raw: string;
  try {
    raw = fs.readFileSync(filePath, 'utf-8');
  } catch (err) {
    const e = err as NodeJS.ErrnoException;
    if (e.code === 'ENOENT') {
      console.log(`[TEAMS] thread-store: no snapshot at ${filePath} — starting empty.`);
    } else {
      console.error(`[TEAMS] thread-store: could not read snapshot ${filePath} — starting empty:`, e.message);
    }
    return 0;
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch (err) {
    console.error(`[TEAMS] thread-store: corrupt snapshot ${filePath} — starting empty:`, (err as Error).message);
    return 0;
  }

  if (!isSnapshot(parsed)) {
    console.error(`[TEAMS] thread-store: unrecognized snapshot shape in ${filePath} — starting empty.`);
    return 0;
  }

  store.restore(parsed);
  const n = store.size();
  console.log(`[TEAMS] thread-store: restored ${n} entry(ies) from ${filePath}.`);
  return n;
}

/**
 * Persist the store to disk atomically (temp file + rename — rename is atomic on
 * the same filesystem, so a reader/next-boot never sees a partial file). Creates
 * the parent directory on first save. Never throws — a failed save is logged and
 * the gateway keeps running on the in-memory store.
 */
export function saveThreadStore(store: ThreadStore, filePath: string): boolean {
  const snapshot = store.snapshot();
  try {
    fs.mkdirSync(path.dirname(filePath), { recursive: true });
    // Unique-enough temp suffix without Math.random (banned in some runtimes):
    // pid + hrtime keeps concurrent saves from racing on the same temp name.
    const tmp = `${filePath}.tmp-${process.pid}-${process.hrtime.bigint().toString(36)}`;
    fs.writeFileSync(tmp, JSON.stringify(snapshot), 'utf-8');
    fs.renameSync(tmp, filePath);
    return true;
  } catch (err) {
    console.error(`[TEAMS] thread-store: failed to save snapshot ${filePath}:`, (err as Error).message);
    return false;
  }
}

/**
 * Start a periodic snapshot timer. Returns a stop function (the timer is
 * `unref`'d so it never keeps the process alive on its own). The periodic save
 * is the crash-safety floor — a SIGKILL between ticks loses at most one interval
 * of mappings; the graceful shutdown path saves once more for a clean state.
 */
export function startSnapshotTimer(
  store: ThreadStore,
  filePath: string,
  intervalMs: number,
): () => void {
  const timer = setInterval(() => {
    saveThreadStore(store, filePath);
  }, intervalMs);
  timer.unref();
  return () => clearInterval(timer);
}

function isSnapshot(value: unknown): value is ThreadStoreSnapshot {
  if (typeof value !== 'object' || value === null) return false;
  const v = value as { version?: unknown; entries?: unknown };
  if (v.version !== 1) return false;
  if (!Array.isArray(v.entries)) return false;
  // Fail-empty contract: a snapshot whose envelope is well-formed but whose
  // entries are malformed is still wrong-shape. Validate EVERY entry — one bad
  // entry rejects the whole file rather than poisoning the indexes with
  // undefined composite keys / invalid references that "restore" silently then
  // blow up at outbound time.
  return v.entries.every(isThreadEntry);
}

function isFiniteNumber(value: unknown): boolean {
  return typeof value === 'number' && Number.isFinite(value);
}

/**
 * Structural guard for a persisted thread entry. Mirrors `ThreadEntry` +
 * `ThreadContext`: a stored mapping is only deliverable if every field outbound
 * relies on is present and the right type. Anything short of that is treated as
 * a corrupt entry (caller fails the whole snapshot empty).
 */
function isThreadEntry(value: unknown): value is ThreadEntry {
  if (typeof value !== 'object' || value === null) return false;
  const e = value as Record<string, unknown>;

  if (typeof e.botSlug !== 'string') return false;
  if (typeof e.conversationId !== 'string') return false;
  if (typeof e.ampMessageId !== 'string') return false;
  if (!isFiniteNumber(e.createdAt)) return false;
  // recordedAt is optional, but if present it must be a usable clock value.
  if (e.recordedAt !== undefined && !isFiniteNumber(e.recordedAt)) return false;
  // aadObjectId (Phase 5) is GRACEFUL: a pre-Phase-5 snapshot predates the field,
  // so absent is accepted (schema-version skew, not corruption) — such an entry
  // stays reply-deliverable, just not DM-indexed (the store skips by-user indexing
  // when it's missing). But if PRESENT it must be a string (no junk poisons the
  // by-user keys). New writes always set it (the write path requires it).
  if (e.aadObjectId !== undefined && typeof e.aadObjectId !== 'string') return false;

  // context: full ConversationReference + rootActivityId + tenantId.
  if (typeof e.context !== 'object' || e.context === null) return false;
  const ctx = e.context as Record<string, unknown>;
  if (typeof ctx.rootActivityId !== 'string') return false;
  if (typeof ctx.tenantId !== 'string') return false;

  // reference.conversation.id is the field outbound resolves to post the reply.
  if (typeof ctx.reference !== 'object' || ctx.reference === null) return false;
  const ref = ctx.reference as Record<string, unknown>;
  if (typeof ref.conversation !== 'object' || ref.conversation === null) return false;
  const conv = ref.conversation as Record<string, unknown>;
  if (typeof conv.id !== 'string') return false;

  return true;
}
