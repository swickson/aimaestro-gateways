/**
 * Discord Gateway - User Resolution (User Directory)
 *
 * Resolves Discord users against the Maestro user directory.
 * Caches resolved users locally with TTL to avoid hammering the API.
 * Auto-creates unknown senders as external users on first contact.
 */

import type { GatewayConfig, ResolvedUser } from './types.js';

interface CacheEntry {
  user: ResolvedUser;
  cachedAt: number;
}

export interface UserResolver {
  resolve(platform: string, platformUserId: string, handle?: string): Promise<ResolvedUser | null>;
  clearCache(): void;
}

export function createUserResolver(config: GatewayConfig): UserResolver {
  const cache = new Map<string, CacheEntry>();
  const cacheTtlMs = config.cache.userTtlMs;

  function debug(message: string, ...args: unknown[]): void {
    if (config.debug) {
      console.log(`[DEBUG] [UserResolver] ${message}`, ...args);
    }
  }

  function cacheKey(platform: string, platformUserId: string): string {
    return `${platform}:${platformUserId}`;
  }

  function getCached(key: string): ResolvedUser | null {
    const entry = cache.get(key);
    if (!entry) return null;
    if (Date.now() - entry.cachedAt > cacheTtlMs) {
      cache.delete(key);
      return null;
    }
    return entry.user;
  }

  function setCache(key: string, user: ResolvedUser): void {
    cache.set(key, { user, cachedAt: Date.now() });
  }

  /**
   * Resolve a platform user via the Maestro user directory.
   * Returns the user record, or null if resolution fails entirely.
   *
   * Flow:
   * 1. Check local cache
   * 2. Call GET /api/users/resolve?platform=...&platformUserId=...
   * 3. If 404: auto-create via POST /api/users/auto-create
   * 4. Cache and return
   */
  async function resolve(
    platform: string,
    platformUserId: string,
    handle?: string
  ): Promise<ResolvedUser | null> {
    const key = cacheKey(platform, platformUserId);

    // 1. Check cache
    const cached = getCached(key);
    if (cached) {
      debug(`Cache hit for ${key}`);
      return cached;
    }

    const baseUrl = config.amp.maestroUrl;
    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${config.amp.apiKey}`,
    };

    // 2. Try resolve endpoint
    try {
      const resolveUrl = `${baseUrl}/api/users/resolve?platform=${encodeURIComponent(platform)}&platformUserId=${encodeURIComponent(platformUserId)}`;
      const resolveRes = await fetch(resolveUrl, {
        headers,
        signal: AbortSignal.timeout(config.polling.timeoutMs),
      });

      if (resolveRes.ok) {
        const body = await resolveRes.json() as { user?: ResolvedUser } & ResolvedUser;
        const user: ResolvedUser = body.user ?? body;
        setCache(key, user);
        debug(`Resolved ${key} -> ${user.displayName} (${user.role})`);

        // Update lastSeen in background (fire-and-forget)
        updateLastSeen(baseUrl, headers, user.id, platform);

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

    // 3. Not found — auto-create as external user
    try {
      const createRes = await fetch(`${baseUrl}/api/users/auto-create`, {
        method: 'POST',
        headers,
        body: JSON.stringify({
          platform,
          platformUserId,
          handle: handle || platformUserId,
        }),
        signal: AbortSignal.timeout(config.polling.timeoutMs),
      });

      if (createRes.ok) {
        const createBody = await createRes.json() as { user?: ResolvedUser } & ResolvedUser;
        const user: ResolvedUser = createBody.user ?? createBody;
        setCache(key, user);
        console.log(`[UserResolver] Auto-created external user: ${user.displayName} (${platform}:${platformUserId})`);
        return user;
      }

      console.error(`[UserResolver] Auto-create failed (${createRes.status}) for ${key}`);
      return null;
    } catch (error) {
      console.error(`[UserResolver] Auto-create request failed for ${key}:`, (error as Error).message);
      return null;
    }
  }

  /**
   * Fire-and-forget PATCH to update lastSeen timestamp.
   */
  function updateLastSeen(
    baseUrl: string,
    headers: Record<string, string>,
    userId: string,
    platform: string
  ): void {
    fetch(`${baseUrl}/api/users/${userId}`, {
      method: 'PATCH',
      headers,
      body: JSON.stringify({
        lastSeenPerPlatform: { [platform]: new Date().toISOString() },
      }),
      signal: AbortSignal.timeout(5000),
    }).catch((err) => {
      // Non-critical — don't log unless debug
      if (config.debug) {
        console.log(`[DEBUG] [UserResolver] lastSeen update failed: ${err.message}`);
      }
    });
  }

  function clearCache(): void {
    cache.clear();
  }

  return { resolve, clearCache };
}
