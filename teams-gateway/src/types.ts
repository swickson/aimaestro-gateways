/**
 * Teams gateway — gateway-local types.
 *
 * Boundary rule (packages/common-PLAN §2 / red-team §0.4): shared AMP +
 * user-directory types live in `@aimaestro/common` and are NEVER redefined here.
 * Only platform-specific shapes are gateway-local:
 *   - `GatewayConfig`   — Teams carries per-bot AAD credentials, not one token.
 *   - `BotRegistryEntry`— one Azure AD app registration + its AMP routing target.
 *   - `ThreadContext`   — the FULL Bot Framework `ConversationReference`, which
 *                         would leak the Teams SDK type into the shared package
 *                         if it lived in common.
 *
 * The shared AMP types are re-exported below so Phase-2/3 code has a single
 * import surface (`./types.js`) without ever redeclaring the contract.
 */

import type { ConversationReference } from '@microsoft/teams.api';

export type {
  AMPEnvelope,
  AMPPayload,
  AMPMessage,
  AMPRouteRequest,
  AMPRouteResponse,
  ResolvedUser,
  EnrichedContext,
  EnrichedSender,
  EnrichedThread,
} from '@aimaestro/common/types.js';

/**
 * One Teams bot identity: an Azure AD app registration paired with the AMP
 * agent it speaks as. `slug` is the routing key (the `/api/<slug>/messages`
 * path); `agentName` is the per-bot AMP identity registered at bootstrap.
 */
export interface BotRegistryEntry {
  /** Routing key — the `/api/<slug>/messages` path segment. */
  slug: string;
  /** Azure AD app (client) ID — the JWT audience this bot's adapter validates. */
  appId: string;
  /** Azure AD client secret used to authenticate outbound sends. */
  appPassword: string;
  /** Single-tenant Azure AD tenant the app is registered in. */
  appTenantId: string;
  /** AMP address this bot routes inbound messages to by default. */
  defaultAgent: string;
  /** Unique AMP agent name (`^[a-z0-9][a-z0-9-]{0,61}[a-z0-9]$`). */
  agentName: string;
}

/**
 * Tenant-scoped operator reference (red-team §0.2). Operator trust MUST be keyed
 * on `(tenantId, aadObjectId)`, never a bare aadObjectId — otherwise an external
 * tenant whose object-id collides with an operator's would bypass the scanner.
 */
export interface OperatorAadRef {
  tenantId: string;
  aadObjectId: string;
}

/**
 * Shape the gateway WRITES into a Teams user's platform-mapping `context` at
 * auto-create — and READS back in tenant-scoped trust resolution. The shared
 * `UserPlatformMapping.context` is an opaque `Record<string, unknown>` on BOTH
 * sides (Maestro stores + returns it verbatim, no typed guarantee), so this
 * gateway-local type is the only thing that makes the producer reliably emit
 * `tenantId`. Without it the directory-operator path is a silent dead path: the
 * read side (`resolveTrust`) would never find a tenant to match and would always
 * fail closed to external. The consumer still reads defensively (the record is
 * Maestro-owned and may be malformed); this type binds the producer.
 */
export interface TeamsPlatformContext {
  /** Azure AD tenant the sender's Teams mapping is bound to (tenant-scoped trust). */
  tenantId: string;
}

export interface GatewayConfig {
  port: number;
  host: string[];
  adminToken: string;
  amp: {
    maestroUrl: string;
    /** AMP org/tenant to register bot identities under (e.g. 'example'). When
     *  unset, bootstrap discovers the provider tenant (falls back to 'default',
     *  which a tenant-scoped Maestro will reject). */
    tenant?: string;
  };
  bots: BotRegistryEntry[];
  operatorAadObjectIds: OperatorAadRef[];
  /** When true, bootstrap resolves + logs the bot plan but performs NO network register. */
  dryRunBootstrap: boolean;
  /** Outbound inbox-poll cadence (Phase 3). */
  polling: {
    intervalMs: number;
  };
  /** Render outbound replies as markdown (default) or plain text (TEAMS_MARKDOWN=0). */
  markdownDefault: boolean;
  /** User-directory resolver cache TTL (CACHE_USER_TTL_MS; default 5 min). */
  cacheUserTtlMs: number;
  /** Where the thread-store JSON snapshot is persisted across restarts (Phase 3). */
  threadStorePath: string;
  /** Periodic thread-store snapshot cadence — crash-safety floor between graceful saves. */
  snapshotIntervalMs: number;
  debug: boolean;
}

/**
 * Thread-store value (red-team finding B / §0.3). Stores the FULL
 * `ConversationReference` — `continueConversation` needs every field
 * (serviceUrl, channelId, conversation, bot, …) — plus the root activity id and
 * the tenant for namespacing. The thin `{conversationId, activityId, serviceUrl}`
 * triple cannot drive a proactive reply and orphans channel threads.
 *
 * Keyed (in the Phase-3 store) by `(botSlug, conversationId, messageId)` so a
 * reply never posts under the wrong bot and two bots can share a channel.
 */
export interface ThreadContext {
  reference: ConversationReference;
  rootActivityId: string;
  tenantId: string;
}
