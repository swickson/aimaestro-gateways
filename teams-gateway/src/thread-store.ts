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
/**
 * Default DM-target horizon (#26): the `byUser` / `byUserAndBot` indexes drive
 * proactive-DM delivery, and a Teams 1:1's `conversationId` is stable for the life
 * of the conversation — so pruning that mapping at the 24h reply-recency horizon
 * caused a spurious cold-start (which then hit the #25 chunk[0] loss). The DM-target
 * mapping is therefore DURABLE by default (unbounded age); the `maxEntries` count cap
 * is its backstop. A long-lived `serviceUrl`/reference may go stale, but #25 delivers
 * via `sendChunk` against the re-ensured conversation, so that's acceptable.
 */
const DEFAULT_USER_MAX_AGE_MS = Infinity;
/** NUL — safe composite-key separator (never appears in a slug/GUID/AMP id). */
const SEP = '\u0000';

export interface ThreadEntry {
  botSlug: string;
  conversationId: string;
  /** The AMP message id returned by `/api/v1/route` — outbound's `in_reply_to`. */
  ampMessageId: string;
  /**
   * The TARGET user this conversation reaches — the inbound sender's AAD object id
   * (the value inbound uses as the directory key: `aadObjectId ?? fromId`). Drives
   * the Phase-5 by-user index so a PROACTIVE DM (no `in_reply_to`) can resolve a
   * conversation to deliver into. REQUIRED on the write path; tolerated absent only
   * on a pre-Phase-5 restored snapshot entry (graceful schema-version skew — such
   * an entry is reply-deliverable but NOT indexed for DM until the user re-contacts).
   */
  aadObjectId?: string;
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
  /**
   * Most-recent entry for a target user across ALL bots (Phase 5). Used only after
   * the unpinned proactive-DM path has proven the user maps to at most one live bot.
   * Null if the user has no live mapping.
   */
  findLatestByUser(aadObjectId: string): ThreadEntry | null;
  /** Most-recent entry for a specific (user, bot) pair (Phase 5 — caller botSlug override). */
  findByUserAndBot(aadObjectId: string, botSlug: string): ThreadEntry | null;
  /** Distinct live bot slugs for a target user, sorted for deterministic caller errors. */
  distinctBotsForUser(aadObjectId: string): string[];
  size(): number;
  snapshot(): ThreadStoreSnapshot;
  restore(snapshot: ThreadStoreSnapshot): void;
}

export interface ThreadStoreOptions {
  /** Insertion-order eviction bound (oldest dropped first). */
  maxEntries?: number;
  /**
   * REPLY-recency horizon in ms (governs `byAmpId` / `recentByConv` — the
   * `in_reply_to` + `isNewConversation` lookups). Entries older than this are lazily
   * expired on the reply-path lookups. Pass `Infinity` to disable age-based pruning
   * (count cap still applies). Defaults to 24h.
   */
  maxAgeMs?: number;
  /**
   * DM-TARGET horizon in ms (governs `byUser` / `byUserAndBot` — proactive-DM
   * resolution; #26). Decoupled from `maxAgeMs` so the durable DM mapping outlives
   * the reply-recency window. Defaults to `Infinity` (durable; `maxEntries` is the
   * backstop). Reply-recency pruning is unaffected by this value.
   */
  userMaxAgeMs?: number;
  /** Injected clock for deterministic tests. Defaults to `Date.now`. */
  now?: () => number;
}

function ampKey(botSlug: string, ampMessageId: string): string {
  return `${botSlug}${SEP}${ampMessageId}`;
}
function convKey(botSlug: string, conversationId: string): string {
  return `${botSlug}${SEP}${conversationId}`;
}
function userBotKey(aadObjectId: string, botSlug: string): string {
  return `${aadObjectId}${SEP}${botSlug}`;
}

export function createThreadStore(options: ThreadStoreOptions = {}): ThreadStore {
  const maxEntries = options.maxEntries ?? DEFAULT_MAX_ENTRIES;
  const maxAgeMs = options.maxAgeMs ?? DEFAULT_MAX_AGE_MS;
  const userMaxAgeMs = options.userMaxAgeMs ?? DEFAULT_USER_MAX_AGE_MS;
  const now = options.now ?? Date.now;
  // Primary index: (botSlug, ampMessageId) -> entry. Insertion-ordered (Map) for
  // O(1) oldest-eviction.
  const byAmpId = new Map<string, ThreadEntry>();
  // Recency index: (botSlug, conversationId) -> the latest entry for it.
  const recentByConv = new Map<string, ThreadEntry>();
  // By-user recency (Phase 5): aadObjectId -> the user's latest entry across ALL
  // bots (drives findLatestByUser = last-seen-bot tiebreak).
  const byUser = new Map<string, ThreadEntry>();
  // By-(user, bot) recency (Phase 5): aadObjectId\0botSlug -> latest for that pair
  // (drives findByUserAndBot when the DM caller pins a botSlug).
  const byUserAndBot = new Map<string, ThreadEntry>();

  /** Age in ms from the store's own clock (`recordedAt`), defensively falling back
   * to `createdAt` for an entry that somehow lacks a record stamp. */
  function ageMs(entry: ThreadEntry): number {
    return now() - (entry.recordedAt ?? entry.createdAt);
  }
  /** Past the REPLY-recency horizon (`byAmpId` / `recentByConv`). */
  function isExpired(entry: ThreadEntry): boolean {
    return ageMs(entry) > maxAgeMs;
  }
  /** Past the DM-TARGET horizon (`byUser` / `byUserAndBot`); #26. Decoupled from
   * `isExpired` so the durable DM mapping outlives the reply-recency window. */
  function isUserExpired(entry: ThreadEntry): boolean {
    return ageMs(entry) > userMaxAgeMs;
  }

  // Lazy expiry is SPLIT per index family (#26): a reply-path lookup aging out an
  // entry must NOT also tear down the still-live DM-target mapping (the single
  // all-index drop() was the root cause of the spurious cold-start). Each family
  // prunes only its own indexes; the count-cap eviction is the only all-index drop.

  /** Drop an entry from the reply-recency indexes only (`byAmpId` + `recentByConv`). */
  function dropReplyIndexes(entry: ThreadEntry): void {
    byAmpId.delete(ampKey(entry.botSlug, entry.ampMessageId));
    const ck = convKey(entry.botSlug, entry.conversationId);
    if (recentByConv.get(ck) === entry) {
      recentByConv.delete(ck);
    }
  }

  /** Drop an entry from the DM-target indexes only (`byUser` + `byUserAndBot`).
   * Clear a pointer only when it still points AT this entry — the pointer always
   * holds the user's NEWEST entry and expiry is age-monotonic (newest expires last),
   * so a cleared pointer can never strand a newer live entry behind it. */
  function dropUserIndexes(entry: ThreadEntry): void {
    if (!entry.aadObjectId) return;
    if (byUser.get(entry.aadObjectId) === entry) {
      byUser.delete(entry.aadObjectId);
    }
    const ubk = userBotKey(entry.aadObjectId, entry.botSlug);
    if (byUserAndBot.get(ubk) === entry) {
      byUserAndBot.delete(ubk);
    }
  }

  function evictIfNeeded(): void {
    while (byAmpId.size > maxEntries) {
      const oldestKey = byAmpId.keys().next().value;
      if (oldestKey === undefined) break;
      const evicted = byAmpId.get(oldestKey);
      byAmpId.delete(oldestKey);
      // Only clear the recency pointers if they still point at the evicted entry.
      // Insertion-order eviction drops the OLDEST entry first, so a user's newest
      // entry (the by-user pointer) is evicted last — clearing without promotion is
      // safe (same invariant as lazy expiry in `drop`).
      if (evicted) {
        const ck = convKey(evicted.botSlug, evicted.conversationId);
        if (recentByConv.get(ck) === evicted) {
          recentByConv.delete(ck);
        }
        if (evicted.aadObjectId) {
          if (byUser.get(evicted.aadObjectId) === evicted) {
            byUser.delete(evicted.aadObjectId);
          }
          const ubk = userBotKey(evicted.aadObjectId, evicted.botSlug);
          if (byUserAndBot.get(ubk) === evicted) {
            byUserAndBot.delete(ubk);
          }
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
    // By-user indexes (Phase 5): only an entry that carries the target user id is
    // DM-routable. A pre-Phase-5 restored entry without one is intentionally skipped
    // (reply-deliverable, just not DM-indexed) so the DM consumer never resolves to
    // an entry it can't attribute to a user.
    if (stored.aadObjectId) {
      const ucur = byUser.get(stored.aadObjectId);
      if (!ucur || stored.createdAt >= ucur.createdAt) {
        byUser.set(stored.aadObjectId, stored);
      }
      const ubk = userBotKey(stored.aadObjectId, stored.botSlug);
      const ubcur = byUserAndBot.get(ubk);
      if (!ubcur || stored.createdAt >= ubcur.createdAt) {
        byUserAndBot.set(ubk, stored);
      }
    }
    evictIfNeeded();
  }

  function findByAmpMessageId(botSlug: string, ampMessageId: string): ThreadEntry | null {
    const entry = byAmpId.get(ampKey(botSlug, ampMessageId));
    if (!entry) return null;
    if (isExpired(entry)) {
      dropReplyIndexes(entry);
      return null;
    }
    return entry;
  }

  function findRecentByConversation(botSlug: string, conversationId: string): ThreadEntry | null {
    const entry = recentByConv.get(convKey(botSlug, conversationId));
    if (!entry) return null;
    if (isExpired(entry)) {
      dropReplyIndexes(entry);
      return null;
    }
    return entry;
  }

  function findLatestByUser(aadObjectId: string): ThreadEntry | null {
    const entry = byUser.get(aadObjectId);
    if (!entry) return null;
    if (isUserExpired(entry)) {
      dropUserIndexes(entry);
      return null;
    }
    return entry;
  }

  function findByUserAndBot(aadObjectId: string, botSlug: string): ThreadEntry | null {
    const entry = byUserAndBot.get(userBotKey(aadObjectId, botSlug));
    if (!entry) return null;
    if (isUserExpired(entry)) {
      dropUserIndexes(entry);
      return null;
    }
    return entry;
  }

  function distinctBotsForUser(aadObjectId: string): string[] {
    const bots = new Set<string>();
    const prefix = `${aadObjectId}${SEP}`;
    for (const [key, entry] of byUserAndBot.entries()) {
      if (!key.startsWith(prefix)) continue;
      if (isUserExpired(entry)) {
        dropUserIndexes(entry);
        continue;
      }
      bots.add(entry.botSlug);
    }
    return [...bots].sort();
  }

  function snapshot(): ThreadStoreSnapshot {
    // Persist every entry still live under EITHER horizon: reply-recency entries so a
    // reply can land post-restart, AND DM-target entries (#26) so a proactive DM still
    // resolves — including one whose reply-path lookup already lazily removed it from
    // byAmpId. Union the three index families and dedup by identity (byAmpId first to
    // preserve insertion/eviction order; DM-only-live entries append after).
    const seen = new Set<ThreadEntry>();
    const entries: ThreadEntry[] = [];
    const consider = (e: ThreadEntry): void => {
      if (seen.has(e)) return;
      seen.add(e);
      if (!isExpired(e) || !isUserExpired(e)) entries.push(e);
    };
    for (const e of byAmpId.values()) consider(e);
    for (const e of byUser.values()) consider(e);
    for (const e of byUserAndBot.values()) consider(e);
    return { version: 1, entries };
  }

  function restore(snap: ThreadStoreSnapshot): void {
    byAmpId.clear();
    recentByConv.clear();
    byUser.clear();
    byUserAndBot.clear();
    // Restore in CHRONOLOGICAL order so byAmpId's insertion order = eviction order. snapshot()
    // appends DM-only-live entries (already reply-expired, so older) AFTER the byAmpId block,
    // so the raw snap.entries order would record those oldest entries LAST — making byAmpId
    // treat them as newest and evictIfNeeded drop newer active entries first (#26 eviction-order
    // bug). Sort by the store's own clock (recordedAt), falling back to createdAt, before the
    // record loop. Skip anything aged out of BOTH horizons while the gateway was down; record()
    // repopulates all indexes — the reply indexes then self-prune on the next reply-path lookup.
    const sortedEntries = [...snap.entries].sort((a, b) => {
      const timeA = a.recordedAt ?? a.createdAt;
      const timeB = b.recordedAt ?? b.createdAt;
      return timeA - timeB;
    });
    for (const entry of sortedEntries) {
      if (!isExpired(entry) || !isUserExpired(entry)) {
        record(entry);
      }
    }
  }

  return {
    record,
    findByAmpMessageId,
    findRecentByConversation,
    findLatestByUser,
    findByUserAndBot,
    distinctBotsForUser,
    size: () => byAmpId.size,
    snapshot,
    restore,
  };
}
