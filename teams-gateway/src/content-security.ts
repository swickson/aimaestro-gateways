/**
 * Teams gateway — trust resolution (the only gateway-local security logic).
 *
 * The 34-pattern injection SCANNER + `<external-content>` wrapping live in
 * `@aimaestro/common/content-security` and are consumed byte-for-byte — there is
 * NO Teams-local copy of the scanner (PLAN §0.2: "content-security scan/wrap from
 * packages/common, NO new copy of the scanner"). This module only resolves a
 * sender's `TrustLevel` against the Teams-specific trust sources, then delegates
 * scanning/wrapping to common's `sanitizeMessage`.
 *
 * TENANT-SCOPED operator trust (red-team §0.2): BOTH trust sources are keyed on
 * the FULL `(tenantId, aadObjectId)` pair, never a bare aadObjectId. The legacy
 * env fallback matches against `OPERATOR_AAD_OBJECT_IDS`; the user-directory path
 * (preferred, Maestro-owned) honors an operator/full record ONLY when the sender's
 * Teams platform mapping is EXPLICITLY bound to the activity's tenant
 * (`platforms[].context.tenantId === activity.tenantId`). A directory record is
 * platform-wide and tenant-agnostic on its own, so without that binding an
 * operator's object-id arriving from the WRONG (or a missing) tenant cannot win
 * via the directory path; only an explicit tenant-scoped legacy whitelist match
 * can still grant operator before the final fail-closed external result.
 */

import {
  resolveTrust as resolveTrustGeneric,
  sanitizeMessage,
  type TrustLevel,
  type TrustResult,
  type InjectionFlag,
} from '@aimaestro/common/content-security.js';
import type { OperatorAadRef, ResolvedUser } from './types.js';

export type { TrustLevel, TrustResult, InjectionFlag };

/**
 * Resolve the trust level for a Teams sender.
 *
 * Order (fail-closed):
 *   0. PROVEN-AAD invariant (#12 security fix): trust elevation requires a real
 *      AAD object id. When `senderAadObjectId` is absent (a Bot-Framework-only
 *      sender with no `aadObjectId`), resolve to external IMMEDIATELY — before any
 *      directory or legacy check — so a fallback identity (e.g. the BF `fromId`)
 *      can never match an operator directory mapping or whitelist entry.
 *   1. User-directory record (preferred, authoritative): an `operator` role or
 *      `full` trustLevel is honored ONLY when the sender's Teams platform mapping
 *      is bound to the activity's tenant (`platforms[].context.tenantId ===
 *      senderTenantId`, both present). A missing tenant id, a missing/unbound
 *      teams mapping, or a tenant mismatch does not grant operator.
 *   2. Legacy env fallback: the sender's `(tenantId, aadObjectId)` must match a
 *      configured `OPERATOR_AAD_OBJECT_IDS` entry EXACTLY. A missing tenant id on
 *      the activity, or a tenant mismatch, can never resolve to operator.
 *   3. Otherwise external.
 */
export function resolveTrust(
  senderTenantId: string | undefined,
  senderAadObjectId: string | undefined,
  operatorAadObjectIds: OperatorAadRef[],
  resolvedUser?: ResolvedUser | null,
): TrustResult {
  // (0) No provable AAD identity => external, fail-closed, before any trust source.
  if (!senderAadObjectId) {
    return {
      level: 'external',
      reason: 'sender has no aadObjectId — not a provable AAD identity; fail closed to external',
    };
  }
  return resolveTrustGeneric({
    resolvedUser,
    senderDescription: senderTenantId
      ? `AAD (${senderTenantId}, ${senderAadObjectId})`
      : `AAD ${senderAadObjectId}`,
    isDirOperatorValid: (user) => {
      const mappingTenantId = teamsMappingTenantId(user, senderAadObjectId);
      if (senderTenantId && mappingTenantId && mappingTenantId === senderTenantId) {
        return {
          isValid: true,
          reasonExtension: `teams mapping bound to activity tenant ${senderTenantId}`,
        };
      }
      return {
        isValid: false,
        reasonExtension: `the teams mapping is not bound to activity tenant ${senderTenantId ?? '(none)'} — fail closed`,
      };
    },
    isLegacyOperator: () => {
      if (senderTenantId) {
        const isOperator = operatorAadObjectIds.some(
          (ref) => ref.tenantId === senderTenantId && ref.aadObjectId === senderAadObjectId,
        );
        if (isOperator) {
          return {
            isOperator: true,
            reason: `AAD (${senderTenantId}, ${senderAadObjectId}) is in tenant-scoped operator whitelist (legacy)`,
          };
        }
      }
      return {
        isOperator: false,
        reason: senderTenantId
          ? `AAD (${senderTenantId}, ${senderAadObjectId}) is not a recognized operator`
          : `AAD ${senderAadObjectId} arrived without a tenant id — fail closed`,
      };
    },
  });
}

/**
 * Tenant id bound to a user's Teams platform mapping, if any. The shared
 * `UserPlatformMapping.context` is an opaque `Record<string, unknown>`
 * (Maestro-owned), so the value is read defensively: only a string `tenantId` on
 * the mapping whose `(type, platformUserId)` matches this Teams sender counts;
 * anything else (no mapping, no `tenantId`, non-string) yields `undefined`, which
 * `resolveTrust` treats as "not provably bound" → fail closed.
 */
function teamsMappingTenantId(
  resolvedUser: ResolvedUser,
  senderAadObjectId: string,
): string | undefined {
  const mapping = resolvedUser.platforms.find(
    (p) => p.type === 'teams' && p.platformUserId === senderAadObjectId,
  );
  const tenantId = mapping?.context?.tenantId;
  return typeof tenantId === 'string' ? tenantId : undefined;
}

/**
 * Scan + wrap a Teams message via the shared scanner. Operator messages bypass
 * scanning entirely (common contract); external messages are scanned against the
 * 34 patterns and wrapped in `<external-content source="teams" …>` when flagged.
 *
 * `senderAadObjectId` becomes the `teams-user-id` attribute on the wrap; the
 * tenant id rides along as an extra escaped attribute for downstream audit.
 */
export function sanitizeTeamsMessage(options: {
  text: string;
  senderAadObjectId: string;
  senderDisplayName: string;
  senderTenantId: string | undefined;
  trustLevel: TrustLevel;
}): { sanitized: string; flags: InjectionFlag[] } {
  return sanitizeMessage({
    text: options.text,
    source: 'teams',
    senderPlatformId: options.senderAadObjectId,
    senderDisplayName: options.senderDisplayName,
    trustLevel: options.trustLevel,
    additionalAttrs: options.senderTenantId ? { 'teams-tenant-id': options.senderTenantId } : undefined,
  });
}
