/**
 * Teams gateway — bot registry parse + validation.
 *
 * The gateway reads a multi-bot registry (PLAN §6) rather than a single App ID.
 * Each entry is a self-contained Azure AD app + AMP routing target. Validation is
 * FAIL-CLOSED at startup (red-team §0.4 "Registry validation fails startup on"):
 * a misconfigured registry throws before any adapter is constructed or any port
 * is opened — never a partial/ambiguous identity surface.
 */

import type { BotConfig } from '@aimaestro/common/amp-bootstrap.js';
import type { BotRegistryEntry } from './types.js';

/** Slug: lowercase alnum + dashes, leading alnum. The `/api/<slug>/messages` segment. */
const SLUG_RE = /^[a-z0-9][a-z0-9-]*$/;
/** AMP agent name constraint (Maestro `/api/v1/register`), mirrored from PLAN §6. */
const AGENT_NAME_RE = /^[a-z0-9][a-z0-9-]{0,61}[a-z0-9]$/;

/**
 * Slugs that would collide with a non-bot HTTP path. `admin` is the management
 * sub-router mount (`/api/admin/*`); `health` guards against `/api/health/...`
 * shadowing intent. Reserving them keeps the bot message paths
 * (`/api/<slug>/messages`) provably disjoint from the management surface, so the
 * auth-by-structure split (mgmt gated, bot paths Bot-Service-JWT-authed) holds.
 */
const RESERVED_SLUGS = new Set(['admin', 'health', 'api']);

interface RawBot {
  slug?: unknown;
  appId?: unknown;
  appPassword?: unknown;
  appTenantId?: unknown;
  defaultAgent?: unknown;
  agentName?: unknown;
}

function reqStr(value: unknown, field: string, slug: string): string {
  if (typeof value !== 'string' || value.trim() === '') {
    throw new Error(`[REGISTRY] bot '${slug}': missing or empty required field '${field}'.`);
  }
  return value.trim();
}

/**
 * Parse + validate the `TEAMS_BOTS` JSON array. Throws at startup on: absent/blank
 * value, invalid JSON, empty array, any missing credential, invalid slug, reserved
 * slug, invalid derived agentName, or a duplicate slug / appId / agentName /
 * defaultAgent (each must map to one distinct identity + adapter).
 */
export function loadBotRegistry(raw: string | undefined): BotRegistryEntry[] {
  if (!raw || raw.trim() === '') {
    throw new Error('[REGISTRY] TEAMS_BOTS is required (a JSON array of bot identities) — fail-closed.');
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch (err) {
    throw new Error(`[REGISTRY] TEAMS_BOTS is not valid JSON: ${(err as Error).message}`);
  }
  if (!Array.isArray(parsed) || parsed.length === 0) {
    throw new Error('[REGISTRY] TEAMS_BOTS must be a non-empty JSON array.');
  }

  const seen = {
    slug: new Set<string>(),
    appId: new Set<string>(),
    agentName: new Set<string>(),
    defaultAgent: new Set<string>(),
  };
  const entries: BotRegistryEntry[] = [];

  for (const item of parsed as RawBot[]) {
    const slug = reqStr(item.slug, 'slug', typeof item.slug === 'string' ? item.slug : '?');
    if (!SLUG_RE.test(slug)) {
      throw new Error(`[REGISTRY] bot '${slug}': invalid slug (must match ${SLUG_RE}).`);
    }
    if (RESERVED_SLUGS.has(slug)) {
      throw new Error(`[REGISTRY] bot '${slug}': slug is reserved (${[...RESERVED_SLUGS].join(', ')}).`);
    }

    const appId = reqStr(item.appId, 'appId', slug);
    const appPassword = reqStr(item.appPassword, 'appPassword', slug);
    const appTenantId = reqStr(item.appTenantId, 'appTenantId', slug);
    const defaultAgent = reqStr(item.defaultAgent, 'defaultAgent', slug);
    const agentName = item.agentName !== undefined
      ? reqStr(item.agentName, 'agentName', slug)
      : `teams-${slug}-bot`;
    if (!AGENT_NAME_RE.test(agentName)) {
      throw new Error(`[REGISTRY] bot '${slug}': agentName '${agentName}' must match ${AGENT_NAME_RE}.`);
    }

    // Fail-closed on collisions — each field keys a distinct identity / adapter.
    const dupChecks: Array<[Set<string>, string, string]> = [
      [seen.slug, slug, 'slug'],
      [seen.appId, appId, 'appId'],
      [seen.agentName, agentName, 'agentName'],
      [seen.defaultAgent, defaultAgent, 'defaultAgent'],
    ];
    for (const [set, value, label] of dupChecks) {
      if (set.has(value)) {
        throw new Error(`[REGISTRY] duplicate ${label}: '${value}' (each bot must be unique).`);
      }
      set.add(value);
    }

    entries.push({ slug, appId, appPassword, appTenantId, defaultAgent, agentName });
  }

  return entries;
}

/**
 * Map the validated registry to `@aimaestro/common`'s multi-bot bootstrap input.
 * The `gateway`/`botSlug` metadata tags each AMP identity so it can be filtered
 * out of the human agent roster (PLAN §6: register hardcodes program='Claude
 * Code'; these are gateway bridge identities, not interactive agents).
 */
export function toBotConfigs(bots: BotRegistryEntry[]): BotConfig[] {
  return bots.map((bot) => ({
    slug: bot.slug,
    agentName: bot.agentName,
    metadata: { gateway: 'teams', botSlug: bot.slug },
  }));
}
