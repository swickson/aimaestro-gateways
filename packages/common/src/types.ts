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

/**
 * Attachment descriptor — the LOCKED gateway <-> Maestro wire type (mirror of
 * ai-maestro `lib/types/amp.ts:190`; Maestro's `routeMessage` accepts `kind:'amp-v1'`
 * VERBATIM and hard-rejects legacy shapes). The gateway never invents these: it
 * runs Maestro's `upload -> PUT -> confirm -> status` flow and assembles this from
 * the `GET /api/v1/attachments/:id/status` response (the only call that yields the
 * signed `url` + server-sanitized `filename`). `url` is the HMAC-signed download
 * link — present only once `scan_status` is `clean`/`basic_clean` — and is itself
 * the auth for `GET .../download` (no Bearer).
 */
export interface AMPAttachmentV1 {
  kind: 'amp-v1';
  /** Maestro attachment id. */
  id: string;
  /** Server-sanitized filename (from /status — confirm does not return it). */
  filename: string;
  /** Server-sniffed content type (authoritative). */
  content_type: string;
  /** Byte size (server-authoritative). */
  size: number;
  /** sha256 hex, server-computed from the streamed bytes. */
  digest: string;
  /** HMAC-signed download URL; present only when scan_status is clean/basic_clean. */
  url: string;
  /**
   * Scan lifecycle: pending -> basic_clean (confirm) / rejected (confirm fail) /
   * clean. `suspicious` is a terminal non-routable verdict Maestro may emit (Watson
   * F1) — NOT in {clean,basic_clean}, so the gateway treats it as undeliverable.
   */
  scan_status: 'pending' | 'basic_clean' | 'clean' | 'rejected' | 'suspicious';
  uploaded_at: string;
  expires_at: string;
}

export interface AMPPayload {
  type: string;
  message: string;
  context?: EnrichedContext | null;
  /** Attachments cited on a routed/replied message (optional; gateways that don't
   *  support attachments neither set nor read this — discord is untouched). */
  attachments?: AMPAttachmentV1[];
  /** Optional rendering hint (e.g. 'status_summary') to select a rich layout platform-side. */
  render?: string;
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
    /** Attachments cited on the routed message (optional; bytes are NOT sent on
     *  /route — only the post-confirm AMPAttachmentV1 descriptors). */
    attachments?: AMPAttachmentV1[];
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
