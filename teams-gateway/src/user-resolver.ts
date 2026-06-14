/**
 * Teams gateway — user resolution (Maestro user directory).
 *
 * Ported from `discord-gateway/src/user-resolver.ts`, with `platform` fixed to
 * `teams` and the hand-rolled Map swapped for `@aimaestro/common`'s TTL `Cache`.
 *
 * User identity in Teams is the AAD object id (a tenant-scoped GUID). Resolution
 * is platform-wide (NOT per-bot) — every Teams bot in the gateway resolves the
 * same sender to the same directory record — so ONE shared resolver (one cache,
 * one auth key) serves all bots. The Maestro `/api/users/*` API wraps its
 * response in `{ user: {...} }`; we unwrap with `body.user ?? body` (CLAUDE.md
 * User Directory note).
 */

import { Cache } from '@aimaestro/common/cache.js';
import type { ResolvedUser, TeamsPlatformContext } from './types.js';

const PLATFORM = 'teams';
const DEFAULT_CACHE_TTL_MS = 300_000; // 5 min (CACHE_USER_TTL_MS parity)
const DEFAULT_TIMEOUT_MS = 10_000;

export interface UserResolverOptions {
  maestroUrl: string;
  /** AMP api key used as the Bearer token for the Maestro user-directory API. */
  apiKey: string;
  cacheTtlMs?: number;
  timeoutMs?: number;
  debug?: boolean;
}

export interface UserResolver {
  /**
   * Resolve (or auto-create) a Teams sender by AAD object id. Null on hard failure.
   *
   * `tenantId` (the activity's AAD tenant) is written into the platform mapping's
   * `context` ONLY on auto-create, so tenant-scoped directory-operator trust has a
   * binding to match against later (see `TeamsPlatformContext`). Lookup is by
   * `(platform, aadObjectId)` alone — the object id is already tenant-unique — so
   * the tenant is not part of the resolve query or cache key.
   */
  resolve(aadObjectId: string, handle?: string, tenantId?: string, botSlug?: string): Promise<ResolvedUser | null>;
  clearCache(): void;
}

export function createUserResolver(options: UserResolverOptions): UserResolver {
  const cache = new Cache<ResolvedUser>(options.cacheTtlMs ?? DEFAULT_CACHE_TTL_MS);
  const timeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  const baseUrl = options.maestroUrl;
  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
    Authorization: `Bearer ${options.apiKey}`,
  };

  // Warn-once latch for the shape-(b) /last-seen route 404ing before Watson's
  // Maestro route lands (SEQUENCING): a missing route must NEVER crash/block
  // inbound, but it should surface ONCE so a permanent 404 in prod is visible.
  let warnedLastSeen404 = false;

  function debug(message: string, ...args: unknown[]): void {
    if (options.debug) {
      console.log(`[DEBUG] [UserResolver] ${message}`, ...args);
    }
  }

  function cacheKey(aadObjectId: string): string {
    return `${PLATFORM}:${aadObjectId}`;
  }

  /**
   * Resolve (or auto-create) a sender, firing shape (b) exactly once for every
   * successful resolution. Shape (b) is fired in `resolve()` — the single
   * unconditional fire-point below — NOT in this lookup, so that EVERY success
   * path (cache-hit, resolve-success, auto-create) reports last-seen and no new
   * path can silently skip it (Whistler review: first-contact previously drifted).
   */
  async function lookup(
    aadObjectId: string,
    handle?: string,
    tenantId?: string,
    botSlug?: string,
  ): Promise<ResolvedUser | null> {
    const key = cacheKey(aadObjectId);

    const cached = cache.get(key);
    if (cached) {
      debug(`Cache hit for ${key}`);
      return cached;
    }

    // 1. Resolve endpoint.
    try {
      const resolveUrl = `${baseUrl}/api/users/resolve?platform=${PLATFORM}&platformUserId=${encodeURIComponent(aadObjectId)}`;
      const resolveRes = await fetch(resolveUrl, {
        headers,
        signal: AbortSignal.timeout(timeoutMs),
      });

      if (resolveRes.ok) {
        const body = (await resolveRes.json()) as { user?: ResolvedUser } & ResolvedUser;
        const user: ResolvedUser = body.user ?? body;
        cache.set(key, user);
        debug(`Resolved ${key} -> ${user.displayName} (${user.role})`);
        return user;
      }

      if (resolveRes.status !== 404) {
        console.error(`[UserResolver] Resolve failed (${resolveRes.status}) for ${key}`);
        return null;
      }
    } catch (error) {
      console.error(`[UserResolver] Resolve request failed for ${key}:`, (error as Error).message);
      return null;
    }

    // 2. Not found — auto-create as an external user. When the activity carried a
    //    tenant id, bind it into the platform mapping's `context` NOW: Maestro
    //    stores context verbatim at first create and never backfills it on a later
    //    resolve, so tenant-scoped directory-operator trust depends on this write
    //    happening on the FIRST create.
    try {
      // Shape (a): bind tenantId (tenant-scoped trust) AND botSlug (DM tiebreak)
      // into the mapping context at first create — Maestro stores it verbatim and
      // never backfills, so both must be written on the FIRST contact.
      const teamsContext: TeamsPlatformContext | undefined =
        tenantId || botSlug ? { ...(tenantId && { tenantId }), ...(botSlug && { botSlug }) } : undefined;
      const createRes = await fetch(`${baseUrl}/api/users/auto-create`, {
        method: 'POST',
        headers,
        body: JSON.stringify({
          platform: PLATFORM,
          platformUserId: aadObjectId,
          handle: handle || aadObjectId,
          ...(teamsContext && { context: teamsContext }),
        }),
        signal: AbortSignal.timeout(timeoutMs),
      });

      if (createRes.ok) {
        const createBody = (await createRes.json()) as { user?: ResolvedUser } & ResolvedUser;
        const user: ResolvedUser = createBody.user ?? createBody;
        cache.set(key, user);
        console.log(`[UserResolver] Auto-created external user: ${user.displayName} (${PLATFORM}:${aadObjectId})`);
        return user;
      }

      console.error(`[UserResolver] Auto-create failed (${createRes.status}) for ${key}`);
      return null;
    } catch (error) {
      console.error(`[UserResolver] Auto-create request failed for ${key}:`, (error as Error).message);
      return null;
    }
  }

  async function resolve(
    aadObjectId: string,
    handle?: string,
    tenantId?: string,
    botSlug?: string,
  ): Promise<ResolvedUser | null> {
    const user = await lookup(aadObjectId, handle, tenantId, botSlug);
    if (user) {
      // Shape (b) — the SINGLE unconditional every-inbound fire-point. Covers
      // ALL success paths (cache-hit, resolve-success, auto-create/first-contact)
      // with exactly one (b) per inbound, so Maestro's stored `context.botSlug`
      // always tracks the most-recently-inbound bot for the DM tiebreak. Routing
      // every success through here is what prevents a path from silently skipping
      // (b) again. Fire-and-forget + 404-graceful; never blocks resolution.
      updateLastSeen(user.id, aadObjectId, botSlug);
    }
    return user;
  }

  /**
   * Shape (b) — fire-and-forget every-inbound last-seen report. REPLACES the old
   * `PATCH /api/users/:id {lastSeenPerPlatform}`: emits Watson's exact
   * `PATCH /api/users/<userId>/last-seen` with `{platform, platformUserId, context:{botSlug}}`
   * so Maestro refreshes the most-recently-inbound bot for the DM tiebreak. Never
   * blocks resolution; a 404 (route not yet deployed) is swallowed with a single
   * warning. Skipped when no botSlug is available (nothing to report).
   */
  function updateLastSeen(userId: string, aadObjectId: string, botSlug?: string): void {
    if (!botSlug) return;
    fetch(`${baseUrl}/api/users/${userId}/last-seen`, {
      method: 'PATCH',
      headers,
      body: JSON.stringify({
        platform: PLATFORM,
        platformUserId: aadObjectId,
        context: { botSlug },
      }),
      signal: AbortSignal.timeout(5000),
    })
      .then((res) => {
        if (res.ok) return;
        if (res.status === 404) {
          if (!warnedLastSeen404) {
            warnedLastSeen404 = true;
            console.warn(
              `[UserResolver] /api/users/:id/last-seen returned 404 — Maestro route not yet deployed; suppressing further warnings (inbound unaffected).`,
            );
          }
          return;
        }
        if (options.debug) {
          console.log(`[DEBUG] [UserResolver] last-seen update non-ok (${res.status}) for ${PLATFORM}:${aadObjectId}`);
        }
      })
      .catch((err) => {
        if (options.debug) {
          console.log(`[DEBUG] [UserResolver] last-seen update failed: ${(err as Error).message}`);
        }
      });
  }

  function clearCache(): void {
    cache.clear();
  }

  return { resolve, clearCache };
}
