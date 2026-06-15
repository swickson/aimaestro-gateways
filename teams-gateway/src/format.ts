/**
 * Teams gateway — outbound formatting (Phase 3).
 *
 * Turns an agent's AMP reply text into one or more Teams-ready message chunks.
 * SDK-decoupled by design (no `@microsoft/teams.*` import) so it stays unit-
 * testable with plain strings; `outbound.ts` wraps each chunk into a message
 * Activity.
 *
 * Decisions (Phase-0 A/B, Shane-ratified):
 *   - MARKDOWN DEFAULT. Teams renders a markdown subset in the message `text`
 *     field, so the v1 deliverable is markdown text + a PLAIN-TEXT fallback
 *     (the `markdown` flag) — reliable on every tenant, no schema to reject.
 *   - ADAPTIVE CARDS are OPT-IN and NOT on the default path. `buildCardScaffold`
 *     exists so the wiring is ready, but nothing calls it; the PR is explicitly
 *     NOT gated on markdown->card rendering (invalid card JSON => Teams silently
 *     rejects the whole message — the risk Phase 0 chose to avoid for v1).
 *   - ~28KB CHUNKING. Teams caps a message near 28KB; long replies are split on
 *     a paragraph/word boundary (ported from discord-gateway/src/outbound.ts,
 *     limit raised from 2000 to 28000).
 */

/** Teams message-size ceiling (~28KB). Chunk boundary for long agent replies. */
export const TEAMS_MAX_LENGTH = 28_000;

export interface FormatOptions {
  /** Sender label prefixed to the reply (the agent's short name). */
  displayName: string;
  /** The agent's reply text. */
  message: string;
  /** Render as markdown (default) or plain text (fallback). */
  markdown: boolean;
  /** Override the chunk ceiling (tests). Defaults to {@link TEAMS_MAX_LENGTH}. */
  maxLength?: number;
}

export interface FormattedReply {
  /** One or more chunks, each within the size ceiling. Empty when nothing to send. */
  chunks: string[];
  /** Whether the chunks are intended to render as markdown. */
  markdown: boolean;
}

/**
 * Split text into chunks no longer than `maxLength`, preferring to break on a
 * newline, then a space, then a hard cut. Ported from the discord poller.
 */
export function chunkText(text: string, maxLength: number = TEAMS_MAX_LENGTH): string[] {
  if (text.length <= maxLength) return [text];

  const chunks: string[] = [];
  let remaining = text;

  while (remaining.length > 0) {
    if (remaining.length <= maxLength) {
      chunks.push(remaining);
      break;
    }

    let splitAt = remaining.lastIndexOf('\n', maxLength);
    if (splitAt < maxLength * 0.5) {
      splitAt = remaining.lastIndexOf(' ', maxLength);
    }
    if (splitAt < maxLength * 0.5) {
      splitAt = maxLength;
    }

    chunks.push(remaining.substring(0, splitAt));
    remaining = remaining.substring(splitAt).trimStart();
  }

  return chunks;
}

/**
 * Build the chunked reply for a Teams send. Prepends a sender label (bold in
 * markdown mode, bare in plain mode) and splits to the size ceiling. A blank /
 * whitespace-only message yields zero chunks — the caller skips the send rather
 * than posting an empty bubble.
 */
export function formatReply(options: FormatOptions): FormattedReply {
  const { displayName, message, markdown } = options;
  const maxLength = options.maxLength ?? TEAMS_MAX_LENGTH;

  const body = (message ?? '').trim();
  if (body === '') {
    return { chunks: [], markdown };
  }

  const label = (displayName || 'Agent').trim() || 'Agent';
  const prefix = markdown ? `**[${label}]** ` : `[${label}] `;

  return { chunks: chunkText(prefix + body, maxLength), markdown };
}

/**
 * OPT-IN Adaptive Card scaffold (NOT on the default path). Wraps text in a
 * minimal valid AdaptiveCard so the opt-in wiring is ready, but no code calls
 * this in v1 — markdown text is the shipped default (see file header). Returns a
 * loosely-typed object: the SDK's card builder is the real construction path
 * when cards are enabled in a later phase.
 */
export function buildCardScaffold(text: string): Record<string, unknown> {
  return {
    type: 'AdaptiveCard',
    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
    version: '1.5',
    body: [{ type: 'TextBlock', text, wrap: true }],
  };
}

export interface StatusSummaryFact {
  title: string;
  value: string;
}

export interface StatusSummary {
  title: string;
  status: 'success' | 'warning' | 'info' | 'error';
  description?: string;
  facts?: StatusSummaryFact[];
}

/**
 * Formats a StatusSummary into a clean markdown fallback text (SDK-decoupled).
 */
export function formatStatusSummaryFallback(data: StatusSummary): string {
  const title = (data.title ?? 'Status Summary').trim();
  const status = (data.status ?? 'unknown').toUpperCase();

  let markdown = `**[${title}]**\n\n`;
  markdown += `Status: **${status}**\n\n`;

  if (data.description && data.description.trim() !== '') {
    markdown += `${data.description.trim()}\n\n`;
  }

  if (data.facts && data.facts.length > 0) {
    for (const fact of data.facts) {
      if (fact && typeof fact === 'object' && fact.title && fact.value) {
        markdown += `- **${fact.title.trim()}**: ${fact.value.trim()}\n`;
      }
    }
  }

  return markdown.trim();
}
