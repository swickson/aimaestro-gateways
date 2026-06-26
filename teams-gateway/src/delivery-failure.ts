import type { AMPMessage } from './types.js';

export const DELIVERY_FAILURE_KIND = 'teams-delivery-failure-v1';

export type DeliveryFailureReason =
  | 'no_conversation'
  | 'mapping_expired'
  | 'teams_api_error'
  | 'bot_unreachable_or_forbidden'
  | 'attachment_rejected'
  | 'attachment_unavailable';

export interface DeliveryFailure {
  kind: typeof DELIVERY_FAILURE_KIND;
  originalMessageId: string;
  botSlug: string;
  reason: DeliveryFailureReason;
  detail: string;
  retryable: boolean;
  attempts: number;
  attemptedAt: string;
}

export interface DeliveryFailureRouteRequest {
  to: string;
  subject: string;
  priority: 'normal';
  in_reply_to: string;
  payload: {
    type: 'notification';
    message: string;
    context: {
      deliveryFailure: DeliveryFailure;
    };
  };
}

export interface SendDeliveryFailureNackInput {
  maestroUrl: string;
  apiKey: string;
  toAgent: string;
  failure: DeliveryFailure;
  timeoutMs: number;
}

export function isDeliveryFailureMessage(msg: AMPMessage): boolean {
  const context = msg.payload?.context as { deliveryFailure?: { kind?: unknown } } | null | undefined;
  return context?.deliveryFailure?.kind === DELIVERY_FAILURE_KIND;
}

export function deliveryFailureMessage(failure: DeliveryFailure): string {
  return `Your Teams message could not be delivered: ${failure.detail}.`;
}

export function buildDeliveryFailureRouteRequest(
  toAgent: string,
  failure: DeliveryFailure,
): DeliveryFailureRouteRequest {
  return {
    to: toAgent,
    subject: `Delivery failed: Teams (${failure.botSlug})`,
    priority: 'normal',
    in_reply_to: failure.originalMessageId,
    payload: {
      type: 'notification',
      message: deliveryFailureMessage(failure),
      context: { deliveryFailure: failure },
    },
  };
}

export async function sendDeliveryFailureNack(input: SendDeliveryFailureNackInput): Promise<void> {
  const resp = await fetch(`${input.maestroUrl}/api/v1/route`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${input.apiKey}`,
    },
    body: JSON.stringify(buildDeliveryFailureRouteRequest(input.toAgent, input.failure)),
    signal: AbortSignal.timeout(input.timeoutMs),
  });

  if (!resp.ok) {
    const body = await resp.text().catch(() => '');
    throw new Error(`AMP delivery-failure route failed (${resp.status})${body ? `: ${body}` : ''}`);
  }
}
