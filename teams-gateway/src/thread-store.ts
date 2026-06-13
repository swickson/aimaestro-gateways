/**
 * Teams gateway — thread store (inbound writes; outbound consumption is Phase 3).
 *
 * Maps a routed message back to the Teams conversation it came from, so a later
 * agent reply can be posted under the ORIGINATING bot's identity via
 * `continueConversation` against the stored `ConversationReference`.
 *
 * NAMESPACING (red-team §0.3): every entry is namespaced by `botSlug` so Demo's
 * reply never posts under Echo, and two bots sharing one channel (same
 * `conversationId`) never collide. The stored value is the FULL
 * `ConversationReference` (+ `rootActivityId` + `tenantId`) — the thin
 * `{conversationId, activityId, serviceUrl}` triple cannot drive a proactive
 * reply and orphans channel threads.
 *
 * Two access patterns are supported:
 *   - `findRecentByConversation(botSlug, conversationId)` — INBOUND, runs BEFORE
 *     the AMP route to compute `thread.inReplyTo` + `isNewConversation`.
 *   - `findByAmpMessageId(botSlug, ampMessageId)` — OUTBOUND (Phase 3), resolves
 *     `(botSlug, in_reply_to)` to the conversation to reply into.
 *
 * The write happens in the ASYNC route path, AFTER `/api/v1/route` returns the
 * AMP message id (the key outbound matches on).
 *
 * PERSISTENCE: `snapshot()` / `restore()` are provided for Phase 3 (the AMP inbox
 * is durable, so a reply can land while the gateway is down — deliverable only if
 * the `(botSlug, ampMessageId) -> ConversationReference` mapping survived the
 * restart). Boot-load + shutdown-save wiring lands in Phase 3; this phase only
 * builds + populates the in-memory store.
 */

import type { ThreadContext } from './types.js';

const DEFAULT_MAX_ENTRIES = 5000;
/**
 * Default recency horizon: a stored conversation reference older than this is
 * pruned. An agent reply that lands a full day after its inbound is vanishingly
 * rare, and a day-old `serviceUrl` may well be stale anyway — so 24h bounds the
 * store by age in addition to the count cap (red-team / Maestro core "bounded store").
 */
const DEFAULT_MAX_AGE_MS = 24 * 60 * 60 * 1000;
/** NUL — safe composite-key separator (never appears in a slug/GUID/AMP id). */
const SEP = '\u0000';

export interface ThreadEntry {
  botSlug: string;
  conversationId: string;
  /** The AMP message id returned by `/api/v1/route` — outbound's `in_reply_to`. */
  ampMessageId: string;
  context: ThreadContext;
  /**
   * Caller-supplied creation time (inbound's clock) — drives the
   * `isNewConversation` heuristic. NOT used for age-eviction: a caller may stamp
   * this on a different clock than the store's.
   */
  createdAt: number;
  /**
   * Store-stamped record time (the store's OWN clock), set at `record()` and
   * preserved across snapshot/restore. Age-based eviction measures from here, so
   * expiry is immune to a caller `createdAt` on a skewed clock and stays correct
   * across a restart. Optional only so external callers needn't supply it.
   */
  recordedAt?: number;
}

/** Serializable form for Phase-3 persistence. */
export interface ThreadStoreSnapshot {
  version: 1;
  entries: ThreadEntry[];
}

export interface ThreadStore {
  record(entry: ThreadEntry): void;
  findByAmpMessageId(botSlug: string, ampMessageId: string): ThreadEntry | null;
  findRecentByConversation(botSlug: string, conversationId: string): ThreadEntry | null;
  size(): number;
  snapshot(): ThreadStoreSnapshot;
  restore(snapshot: ThreadStoreSnapshot): void;
}

export interface ThreadStoreOptions {
  /** Insertion-order eviction bound (oldest dropped first). */
  maxEntries?: number;
  /**
   * Recency horizon in ms. Entries older than this are lazily expired on lookup
   * and dropped from `snapshot()`/`restore()`. Pass `Infinity` to disable
   * age-based pruning (count cap still applies). Defaults to 24h.
   */
  maxAgeMs?: number;
  /** Injected clock for deterministic tests. Defaults to `Date.now`. */
  now?: () => number;
}

function ampKey(botSlug: string, ampMessageId: string): string {
  return `${botSlug}${SEP}${ampMessageId}`;
}
function convKey(botSlug: string, conversationId: string): string {
  return `${botSlug}${SEP}${conversationId}`;
}

export function createThreadStore(options: ThreadStoreOptions = {}): ThreadStore {
  const maxEntries = options.maxEntries ?? DEFAULT_MAX_ENTRIES;
  const maxAgeMs = options.maxAgeMs ?? DEFAULT_MAX_AGE_MS;
  const now = options.now ?? Date.now;
  // Primary index: (botSlug, ampMessageId) -> entry. Insertion-ordered (Map) for
  // O(1) oldest-eviction.
  const byAmpId = new Map<string, ThreadEntry>();
  // Recency index: (botSlug, conversationId) -> the latest entry for it.
  const recentByConv = new Map<string, ThreadEntry>();

  /**
   * True once an entry has aged past the recency horizon. Measured from
   * `recordedAt` (the store's own clock), falling back to `createdAt` only for an
   * entry that somehow lacks a record stamp (defensive).
   */
  function isExpired(entry: ThreadEntry): boolean {
    return now() - (entry.recordedAt ?? entry.createdAt) > maxAgeMs;
  }

  /** Drop an entry from both indexes (used by lazy expiry). */
  function drop(entry: ThreadEntry): void {
    byAmpId.delete(ampKey(entry.botSlug, entry.ampMessageId));
    const ck = convKey(entry.botSlug, entry.conversationId);
    if (recentByConv.get(ck) === entry) {
      recentByConv.delete(ck);
    }
  }

  function evictIfNeeded(): void {
    while (byAmpId.size > maxEntries) {
      const oldestKey = byAmpId.keys().next().value;
      if (oldestKey === undefined) break;
      const evicted = byAmpId.get(oldestKey);
      byAmpId.delete(oldestKey);
      // Only clear the recency pointer if it still points at the evicted entry.
      if (evicted) {
        const ck = convKey(evicted.botSlug, evicted.conversationId);
        if (recentByConv.get(ck) === evicted) {
          recentByConv.delete(ck);
        }
      }
    }
  }

  function record(entry: ThreadEntry): void {
    // Stamp the store's own clock for age-eviction, but PRESERVE an existing
    // stamp so restore() carries the original record time across a restart.
    const stored: ThreadEntry = { ...entry, recordedAt: entry.recordedAt ?? now() };
    byAmpId.set(ampKey(stored.botSlug, stored.ampMessageId), stored);
    const ck = convKey(stored.botSlug, stored.conversationId);
    const current = recentByConv.get(ck);
    // Keep the newest entry as the conversation's recency pointer.
    if (!current || stored.createdAt >= current.createdAt) {
      recentByConv.set(ck, stored);
    }
    evictIfNeeded();
  }

  function findByAmpMessageId(botSlug: string, ampMessageId: string): ThreadEntry | null {
    const entry = byAmpId.get(ampKey(botSlug, ampMessageId));
    if (!entry) return null;
    if (isExpired(entry)) {
      drop(entry);
      return null;
    }
    return entry;
  }

  function findRecentByConversation(botSlug: string, conversationId: string): ThreadEntry | null {
    const entry = recentByConv.get(convKey(botSlug, conversationId));
    if (!entry) return null;
    if (isExpired(entry)) {
      drop(entry);
      return null;
    }
    return entry;
  }

  function snapshot(): ThreadStoreSnapshot {
    // Persist only live entries — never carry aged-out references across a restart.
    return { version: 1, entries: [...byAmpId.values()].filter((e) => !isExpired(e)) };
  }

  function restore(snap: ThreadStoreSnapshot): void {
    byAmpId.clear();
    recentByConv.clear();
    // Restore in insertion order so eviction ordering + recency stay consistent;
    // skip anything that aged out while the gateway was down.
    for (const entry of snap.entries) {
      if (!isExpired(entry)) {
        record(entry);
      }
    }
  }

  return {
    record,
    findByAmpMessageId,
    findRecentByConversation,
    size: () => byAmpId.size,
    snapshot,
    restore,
  };
}
