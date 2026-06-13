/**
 * Shared AMP + User-directory type definitions.
 *
 * Extracted from the per-gateway `types.ts` copies. Only the platform-agnostic
 * AMP protocol and user-directory types live here. Platform-specific types stay
 * gateway-local by design (per packages/common-PLAN §2 boundary rules):
 *   - `GatewayConfig`  — differs per platform (Discord token vs Teams AAD creds)
 *   - `ThreadContext`  — Discord channel/message IDs vs Teams ConversationReference
 *   - `WatchWebhookEntry` etc. — Discord-only
 */

// ---------------------------------------------------------------------------
// Enriched envelope context — the gateway -> Maestro CONTRACT (locked shape)
// ---------------------------------------------------------------------------

/**
 * Producer-side enriched context attached to a routed message's payload. This is
 * the LOCKED gateway -> Maestro contract — Maestro core signed it off verbatim against
 * the Maestro consumer (amp-service.ts) on 2026-06-09. Do NOT deviate from the
 * field shapes below; Maestro's memory-retrieval middleware reads them.
 *
 * `platformUserId` + `platform` are the load-bearing REQUIRED pair (this killed
 * the earlier `userId`-vs-`platformUserId` drift — there is deliberately NO
 * `userId` field here). `topicHints` is capped at 3 PRODUCER-SIDE; Maestro does
 * NOT enforce the cap.
 */
export interface EnrichedSender {
  platformUserId: string;
  platform: string;
  displayName?: string;
  handle?: string;
  trustLevel?: string;
  role?: string;
}

export interface EnrichedThread {
  threadId: string;
  /** Prior AMP message id in this thread, when known. Omitted (never `null`) otherwise. */
  inReplyTo?: string;
  isNewConversation: boolean;
}

export interface EnrichedContext {
  sender: EnrichedSender;
  thread: EnrichedThread;
  /** Up to 3 topic keywords; capped producer-side. */
  topicHints: string[];
}

// ---------------------------------------------------------------------------
// AMP Protocol Types
// ---------------------------------------------------------------------------

export interface AMPEnvelope {
  version: string;
  id: string;
  from: string;
  to: string;
  subject: string;
  priority: 'urgent' | 'high' | 'normal' | 'low';
  timestamp: string;
  signature: string | null;
  in_reply_to?: string | null;
  thread_id?: string;
  expires_at?: string | null;
}

export interface AMPPayload {
  type: string;
  message: string;
  context?: EnrichedContext | null;
}

export interface AMPMessage {
  envelope: AMPEnvelope;
  payload: AMPPayload;
  metadata?: {
    status?: string;
    queued_at?: string;
    delivery_attempts?: number;
  };
  local?: {
    received_at?: string;
    delivery_method?: string;
    status?: string;
  };
}

export interface AMPRouteRequest {
  to: string;
  subject: string;
  priority?: 'urgent' | 'high' | 'normal' | 'low';
  in_reply_to?: string | null;
  payload: {
    type: string;
    message: string;
    context?: EnrichedContext;
  };
}

export interface AMPRouteResponse {
  id: string;
  status: 'delivered' | 'queued' | 'failed';
  method?: string;
  delivered_at?: string;
  error?: string;
  message?: string;
}

// ---------------------------------------------------------------------------
// User Directory (resolved from Maestro /api/users/resolve)
// ---------------------------------------------------------------------------

export interface UserPlatformMapping {
  type: string;
  platformUserId: string;
  handle: string;
  context: Record<string, unknown>;
}

export interface ResolvedUser {
  id: string;
  displayName: string;
  aliases: string[];
  platforms: UserPlatformMapping[];
  role: 'operator' | 'external';
  trustLevel: 'full' | 'none';
  preferredPlatform?: string;
  createdAt: string;
  updatedAt: string;
}

// ---------------------------------------------------------------------------
// Agent Resolution (simplified for AMP)
// ---------------------------------------------------------------------------

export interface LookupResult {
  address: string;
  displayName: string;
}
