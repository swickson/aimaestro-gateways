/**
 * WhatsApp Gateway - Type Definitions (AMP Protocol)
 */

// ---------------------------------------------------------------------------
// AMP Protocol Types
// ---------------------------------------------------------------------------

export interface AMPEnvelope {
  version: string;
  id: string;
  from: string;
  to: string;
  subject: string;
  priority: 'critical' | 'high' | 'normal' | 'low';
  timestamp: string;
  signature?: string;
  in_reply_to?: string;
  thread_id?: string;
}

export interface AMPPayload {
  type: string;
  message: string;
  context?: Record<string, any>;
  attachments?: any[];
}

export interface AMPMessage {
  envelope: AMPEnvelope;
  payload: AMPPayload;
  /**
   * Card B receiver-added enrichment (memory recall). Gateways DEGRADE: ignore it and
   * relay only payload.message — never render recall to a platform user (memory leak).
   * Typed `unknown` to force a deliberate cast; canonical type = @aimaestro/common Enrichment.
   */
  enrichment?: unknown;
}

export interface AMPRouteRequest {
  to: string;
  subject: string;
  priority?: string;
  in_reply_to?: string;
  payload: {
    type: string;
    message: string;
    context?: Record<string, any>;
    attachments?: any[];
  };
  signature?: string;
}

export interface AMPRouteResponse {
  id: string;
  status: string;
  method?: string;
  delivered_at?: string;
  error?: string;
}

// ---------------------------------------------------------------------------
// Gateway Config
// ---------------------------------------------------------------------------

export interface RouteTarget {
  agent: string;
}

export interface GatewayConfig {
  port: number;
  debug: boolean;
  amp: {
    apiKey: string;
    agentAddress: string;
    maestroUrl: string;
    defaultAgent: string;
    tenant: string;
    inboxDir: string;
  };
  whatsapp: {
    stateDir: string;
    allowFrom: string[];
    dmPolicy: 'allowlist' | 'open' | 'disabled';
    sendReadReceipts: boolean;
    textChunkLimit: number;
  };
  routing: {
    phones: Record<string, RouteTarget>;
    default: RouteTarget;
  };
  outbound: {
    pollIntervalMs: number;
  };
  operatorPhones: string[];
  adminToken: string;
}

// ---------------------------------------------------------------------------
// WhatsApp Message Types
// ---------------------------------------------------------------------------

export interface WhatsAppInboundMessage {
  from: string;
  fromName: string;
  chatJid: string;
  messageId: string;
  isGroup: boolean;
  groupJid: string | null;
  groupName: string | null;
  textBody: string;
  quotedMessage: QuotedMessage | null;
  hasMedia: boolean;
  mediaType: string | null;
  timestamp: string;
}

export interface QuotedMessage {
  id: string;
  sender: string;
  body: string;
}

export interface WhatsAppSendPayload {
  to: string;
  message: string;
  quotedMessageId?: string;
  accountId?: string;
}
