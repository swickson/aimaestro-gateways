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
  resolve(aadObjectId: string, handle?: string, tenantId?: string): Promise<ResolvedUser | null>;
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

  function debug(message: string, ...args: unknown[]): void {
    if (options.debug) {
      console.log(`[DEBUG] [UserResolver] ${message}`, ...args);
    }
  }

  function cacheKey(aadObjectId: string): string {
    return `${PLATFORM}:${aadObjectId}`;
  }

  async function resolve(aadObjectId: string, handle?: string, tenantId?: string): Promise<ResolvedUser | null> {
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
        updateLastSeen(user.id);
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
      const teamsContext: TeamsPlatformContext | undefined = tenantId ? { tenantId } : undefined;
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

  /** Fire-and-forget lastSeen update — non-critical, never blocks resolution. */
  function updateLastSeen(userId: string): void {
    fetch(`${baseUrl}/api/users/${userId}`, {
      method: 'PATCH',
      headers,
      body: JSON.stringify({ lastSeenPerPlatform: { [PLATFORM]: new Date().toISOString() } }),
      signal: AbortSignal.timeout(5000),
    }).catch((err) => {
      if (options.debug) {
        console.log(`[DEBUG] [UserResolver] lastSeen update failed: ${(err as Error).message}`);
      }
    });
  }

  function clearCache(): void {
    cache.clear();
  }

  return { resolve, clearCache };
}
