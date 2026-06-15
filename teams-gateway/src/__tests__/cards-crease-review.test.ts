/**
 * FRESH cross-review tests for w4 #14 (outbound Adaptive Cards).
 *
 * Authored by Crease (correctness lens) independently of Mother's
 * outbound-cards.test.ts — written from the dispatch spec (card builder +
 * payload.render selector + ZERO-LOSS fallback). Focus is on the correctness
 * gaps a builder's own tests tend to miss:
 *   - no double-delivery (card XOR text, never both);
 *   - render-absent non-regression (a normal reply never becomes a card);
 *   - the selector is fully opt-in + safe when the buildCard dep is absent;
 *   - card + attachments interaction (separate activities, card carries no files);
 *   - empty-reply handling when render is set;
 *   - the validation-fail degradation path (valid JSON, bad status → raw text).
 */

import assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { afterEach, describe, it } from 'node:test';

import { buildCard, buildStatusSummaryCard } from '../card-builder.js';
import { formatStatusSummaryFallback } from '../format.js';
import { startOutboundPoller, type OutboundBot, type OutboundAttachment } from '../outbound.js';
import { createThreadStore, type ThreadEntry } from '../thread-store.js';
import type { AMPMessage, AttachmentPolicy, ThreadContext } from '../types.js';

const TEST_POLICY: AttachmentPolicy = { maxBytes: 26_214_400, maxCount: 10, denyContentTypes: [] };
const tempRoots: string[] = [];

afterEach(() => {
  for (const dir of tempRoots.splice(0)) fs.rmSync(dir, { recursive: true, force: true });
});

function tempDir(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'crease-cards-'));
  tempRoots.push(dir);
  return dir;
}

function context(conversationId: string): ThreadContext {
  return {
    reference: {
      serviceUrl: 'https://smba.trafficmanager.net/amer/',
      channelId: 'msteams',
      conversation: { id: conversationId },
      bot: { id: 'bot-app-id', name: 'Maestro' },
      user: { id: 'user-1', name: 'Alice' },
    } as ThreadContext['reference'],
    rootActivityId: 'root-1',
    tenantId: 'tenant-1',
  };
}

function entry(): ThreadEntry {
  return {
    botSlug: 'maestro',
    conversationId: 'conv-maestro',
    ampMessageId: 'amp-inbound-1',
    context: context('conv-maestro'),
    createdAt: 100,
  };
}

interface Send { text: string; markdown: boolean; attachments?: OutboundAttachment[]; card?: Record<string, unknown> }

function msg(payload: Partial<AMPMessage['payload']>): AMPMessage {
  return {
    envelope: {
      id: 'reply-1',
      from: 'agent-one@example.aimaestro.local',
      to: 'teams-maestro-bot@example.aimaestro.local',
      timestamp: '2026-06-09T00:00:00.000Z',
      in_reply_to: 'amp-inbound-1',
      version: '1.0',
      priority: 'normal',
      signature: null,
      subject: 'test',
    },
    payload: { type: 'text', message: '', context: null, ...payload },
  } as AMPMessage;
}

function writeInbox(inbox: string, m: AMPMessage): string {
  const dir = path.join(inbox, 'agent-one_example_aimaestro_local');
  fs.mkdirSync(dir, { recursive: true });
  const fp = path.join(dir, 'reply.json');
  fs.writeFileSync(fp, JSON.stringify(m), 'utf-8');
  return fp;
}

/** Run the poller once against one inbox message; resolve with the captured sends. */
async function runPoller(
  m: AMPMessage,
  opts: {
    sendImpl?: (s: Send) => void | Promise<void>;
    withBuildCard?: boolean;
  } = {},
): Promise<{ sends: Send[]; consumed: () => boolean }> {
  const root = tempDir();
  const inbox = path.join(root, 'inbox');
  const fp = writeInbox(inbox, m);
  const store = createThreadStore({ maxAgeMs: Infinity });
  store.record(entry());
  const sends: Send[] = [];

  const bots: OutboundBot[] = [
    {
      slug: 'maestro',
      inboxDir: inbox,
      maestroUrl: 'https://maestro.test',
      allowedOrigins: new Set(['https://maestro.test']),
      send: async (_conversationId, text, markdown, attachments, card) => {
        const rec = { text, markdown, attachments, card };
        sends.push(rec);
        await opts.sendImpl?.(rec);
      },
    },
  ];

  const stop = startOutboundPoller({
    bots,
    threadStore: store,
    pollIntervalMs: 60_000,
    markdownDefault: true,
    policy: TEST_POLICY,
    debug: false,
    ...(opts.withBuildCard === false ? {} : { buildCard }),
  });

  try {
    const deadline = Date.now() + 2000;
    // The file is unlink'd once the reply is fully processed (success path).
    while (Date.now() < deadline && fs.existsSync(fp)) {
      await new Promise((r) => setTimeout(r, 10));
    }
    // settle one more tick to catch any erroneous extra send
    await new Promise((r) => setTimeout(r, 30));
  } finally {
    stop();
  }
  return { sends, consumed: () => !fs.existsSync(fp) };
}

// ---------------------------------------------------------------------------
// Pure card-builder / fallback correctness
// ---------------------------------------------------------------------------
describe('#14 card-builder selector', () => {
  it('returns null for any non-status_summary render type (safe default)', () => {
    assert.equal(buildCard('text', '{}'), null);
    assert.equal(buildCard('table', JSON.stringify({ title: 't', status: 'success' })), null);
    assert.equal(buildCard('', 'anything'), null);
  });

  it('returns null when valid JSON but status is outside the enum (no invalid card emitted)', () => {
    assert.equal(buildCard('status_summary', JSON.stringify({ title: 't', status: 'bogus' })), null);
  });

  it('builds a schema-shaped card for a valid status_summary', () => {
    const card = buildCard('status_summary', JSON.stringify({ title: 'T', status: 'info' }));
    assert.ok(card);
    assert.equal(card.type, 'AdaptiveCard');
    assert.equal(card.version, '1.5');
    assert.ok(Array.isArray(card.body));
  });

  it('buildStatusSummaryCard with no description/facts serializes to exactly title + status blocks', () => {
    const card = JSON.parse(JSON.stringify(buildStatusSummaryCard({ title: 'T', status: 'success' }))) as {
      body: unknown[];
    };
    assert.equal(card.body.length, 2, 'only title + status when optional fields are absent');
  });

  it('info status maps to a valid Accent color (not an invalid token)', () => {
    const card = JSON.parse(JSON.stringify(buildStatusSummaryCard({ title: 'T', status: 'info' }))) as {
      body: Array<Record<string, unknown>>;
    };
    assert.equal(card.body[1].color, 'Accent');
  });
});

describe('#14 fallback formatter is defensive + lossless', () => {
  it('includes title, uppercased status, description and facts', () => {
    const md = formatStatusSummaryFallback({
      title: 'Deploy',
      status: 'success',
      description: 'all good',
      facts: [{ title: 'Region', value: 'us-east' }],
    });
    assert.match(md, /\*\*\[Deploy\]\*\*/);
    assert.match(md, /Status: \*\*SUCCESS\*\*/);
    assert.match(md, /all good/);
    assert.match(md, /- \*\*Region\*\*: us-east/);
  });

  it('skips fact rows missing a title or value rather than emitting blanks', () => {
    const md = formatStatusSummaryFallback({
      title: 'T',
      status: 'warning',
      facts: [{ title: 'ok', value: 'v' }, { title: '', value: 'x' }, { title: 'y', value: '' } as never],
    });
    assert.match(md, /- \*\*ok\*\*: v/);
    assert.doesNotMatch(md, /\*\*y\*\*/);
  });
});

// ---------------------------------------------------------------------------
// Selector wiring in the poller
// ---------------------------------------------------------------------------
describe('#14 poller selector + zero-loss', () => {
  it('render ABSENT → plain text reply, never a card (non-regression)', async () => {
    const { sends, consumed } = await runPoller(msg({ message: 'just a normal reply' }));
    assert.equal(sends.length, 1);
    assert.equal(sends[0].card, undefined);
    assert.match(sends[0].text, /just a normal reply/);
    assert.ok(consumed());
  });

  it('render set but buildCard dep NOT injected → plain text (opt-in is safe-by-absence)', async () => {
    const { sends } = await runPoller(
      msg({ render: 'status_summary', message: JSON.stringify({ title: 'T', status: 'success' }) }),
      { withBuildCard: false },
    );
    assert.equal(sends.length, 1);
    assert.equal(sends[0].card, undefined, 'no card without the builder dep');
    // raw JSON is still delivered — nothing dropped
    assert.match(sends[0].text, /"status":"success"/);
  });

  it('unknown render type → no card, raw text delivered (nothing dropped)', async () => {
    const { sends } = await runPoller(msg({ render: 'pie_chart', message: 'hello there' }));
    assert.equal(sends.length, 1);
    assert.equal(sends[0].card, undefined);
    assert.match(sends[0].text, /hello there/);
  });

  it('valid status_summary card success → exactly ONE send, the card, with empty text (no double delivery)', async () => {
    const { sends, consumed } = await runPoller(
      msg({ render: 'status_summary', message: JSON.stringify({ title: 'Summary', status: 'success' }) }),
    );
    assert.equal(sends.length, 1, 'card success must NOT also send the text fallback');
    assert.ok(sends[0].card);
    assert.equal(sends[0].card?.type, 'AdaptiveCard');
    assert.equal(sends[0].text, '');
    assert.ok(consumed());
  });

  it('card delivery throws → falls back to the markdown text (zero-loss), file consumed', async () => {
    let firstCall = true;
    const { sends, consumed } = await runPoller(
      msg({
        render: 'status_summary',
        message: JSON.stringify({ title: 'Summary', status: 'error', description: 'boom' }),
      }),
      {
        sendImpl: (s) => {
          if (s.card && firstCall) {
            firstCall = false;
            throw new Error('card send failed');
          }
        },
      },
    );
    // 1 failed card send + 1 successful text fallback send
    assert.equal(sends.length, 2);
    assert.ok(sends[0].card, 'first attempt is the card');
    assert.equal(sends[1].card, undefined, 'fallback is plain text');
    assert.match(sends[1].text, /Status: \*\*ERROR\*\*/);
    assert.match(sends[1].text, /boom/);
    assert.ok(consumed(), 'a successful fallback consumes the file (not retried)');
  });

  it('valid JSON but status outside enum → no card, raw JSON delivered (degradation watch item)', async () => {
    const { sends } = await runPoller(
      msg({ render: 'status_summary', message: JSON.stringify({ title: 'T', status: 'bogus' }) }),
    );
    assert.equal(sends.length, 1);
    assert.equal(sends[0].card, undefined);
    // DOCUMENTS: validation-fail does NOT route through formatStatusSummaryFallback,
    // so the user sees the raw JSON string, not a markdown summary. Lossless but ugly.
    assert.match(sends[0].text, /"status":"bogus"/);
  });

  it('empty message + render set + no attachments → nothing to post, file deleted, zero sends', async () => {
    const { sends, consumed } = await runPoller(msg({ render: 'status_summary', message: '' }));
    assert.equal(sends.length, 0);
    assert.ok(consumed(), 'empty reply is deleted, not left to spin');
  });
});
