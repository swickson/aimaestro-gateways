/**
 * w3 — inbound route-payload citation (inbound.ts handleInbound).
 *
 * Drives the full inbound pipeline with a combined mock `fetch` (attachment legs
 * upload/PUT/confirm/status + the /api/v1/route call). Asserts: a successful
 * ingest is cited in payload.attachments; an attachment-only (empty text) message
 * still routes; and FAIL-OPEN — a failed ingest still routes the text with NO
 * attachments (a failed attachment leg never drops the message).
 */

import assert from 'node:assert/strict';
import { afterEach, beforeEach, describe, it } from 'node:test';

import { Cache } from '@aimaestro/common/cache.js';
import { handleInbound, type InboundActivity, type InboundDeps } from '../inbound.js';
import { createThreadStore } from '../thread-store.js';
import type { AMPRouteRequest, AttachmentPolicy, OperatorAadRef } from '../types.js';
import type { RawInboundAttachment } from '../attachments-inbound.js';

const POLICY: AttachmentPolicy = { maxBytes: 26_214_400, maxCount: 10, denyContentTypes: [] };

const origFetch = globalThis.fetch;
const origLog = console.log;
const origError = console.error;
beforeEach(() => {
  console.log = () => undefined;
  console.error = () => undefined;
});
afterEach(() => {
  globalThis.fetch = origFetch;
  console.log = origLog;
  console.error = origError;
});

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
}

/** Combined attachment-flow + route mock; captures the routed AMPRouteRequest(s). */
function installFetch(opts: { routed: AMPRouteRequest[]; uploadStatus?: number } = { routed: [] }): void {
  globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = String(input);
    if (url.includes('/api/v1/route')) {
      opts.routed.push(JSON.parse(String(init?.body)) as AMPRouteRequest);
      return json({ id: 'amp-out-1', status: 'delivered' });
    }
    if (url.includes('/attachments/upload')) {
      return json({ attachment_id: 'att-1', upload_url: 'https://maestro.test/signed/att-1' }, opts.uploadStatus ?? 200);
    }
    if (url.includes('/signed/')) return new Response(null, { status: 200 });
    if (url.includes('/confirm')) return json({ ok: true });
    if (url.includes('/status')) {
      return json({
        attachment_id: 'att-1', filename: 'clean.pdf', content_type: 'application/pdf', size: 4,
        digest: 'd', scan_status: 'basic_clean', uploaded_at: '2026-06-14T00:00:00.000Z',
        expires_at: '2026-06-15T00:00:00.000Z', url: 'https://maestro.test/api/v1/attachments/att-1/download?sig=x',
      });
    }
    throw new Error(`unexpected fetch ${url}`);
  }) as typeof fetch;
}

function deps(overrides: Partial<InboundDeps> = {}): InboundDeps {
  return {
    bot: { slug: 'maestro', defaultAgent: 'ops@example.aimaestro.local', agentName: 'teams-maestro-bot', ampAddress: 'teams-maestro-bot@example.aimaestro.local', ampApiKey: 'bot-key' },
    maestroUrl: 'https://maestro.test',
    operatorAadObjectIds: [],
    userResolver: { resolve: async () => null, clearCache: () => undefined },
    threadStore: createThreadStore(),
    dedupe: new Cache<true>(60_000),
    attachmentPolicy: POLICY,
    timeoutMs: 1000,
    debug: false,
    now: () => 1_000_000,
    ...overrides,
  };
}

function activity(text: string, attachments?: RawInboundAttachment[], downloader?: InboundActivity['downloadAttachment']): InboundActivity {
  return {
    activityId: 'act-' + Math.round(1e6 * Number('0.' + text.length)), // varies by input
    conversationId: 'conversation-1',
    conversationType: 'personal',
    aadObjectId: 'aad-1',
    fromId: 'bf-1',
    fromName: 'Alice',
    text,
    tenantId: 'tenant-1',
    serviceUrl: 'https://smba.trafficmanager.net/amer/',
    reference: { serviceUrl: 'https://smba.trafficmanager.net/amer/', channelId: 'msteams', conversation: { id: 'conversation-1' }, bot: { id: 'bot-app-id', name: 'Maestro' }, user: { id: 'bf-1', name: 'Alice' } } as InboundActivity['reference'],
    attachments,
    downloadAttachment: downloader,
  };
}

const file: RawInboundAttachment = { name: 'doc.pdf', contentType: 'application/pdf', downloadUrl: 'https://teams.test/dl' };
const goodDownloader = async () => new Uint8Array([1, 2, 3, 4]);

// The activity() helper's sender is (tenant-1, aad-1). As an OPERATOR they bypass the
// scanner, so empty text stays empty — the exact live-repro condition (the operator DM'd a
// photo) where the empty-message guard must fire. (An EXTERNAL sender's empty text is
// wrapped in a non-empty <external-content> envelope, so it never hits the 400 and the
// placeholder is intentionally not needed.)
const OPERATOR: OperatorAadRef[] = [{ tenantId: 'tenant-1', aadObjectId: 'aad-1' }];

describe('w3 inbound route-payload citation', () => {
  it('cites a successfully-ingested attachment in payload.attachments', async () => {
    const routed: AMPRouteRequest[] = [];
    installFetch({ routed });
    const result = await handleInbound(activity('here is a file', [file], goodDownloader), deps());

    assert.equal(result, 'routed');
    assert.equal(routed.length, 1);
    assert.equal(routed[0].payload.attachments?.length, 1);
    const a = routed[0].payload.attachments![0];
    assert.equal(a.kind, 'amp-v1');
    assert.equal(a.id, 'att-1');
    assert.equal(a.url, 'https://maestro.test/api/v1/attachments/att-1/download?sig=x');
  });

  it('Part B — attachment-only (empty text) that SUCCEEDS: cites the attachment AND substitutes a non-empty placeholder', async () => {
    const routed: AMPRouteRequest[] = [];
    installFetch({ routed });
    const result = await handleInbound(activity('', [file], goodDownloader), deps({ operatorAadObjectIds: OPERATOR }));

    assert.equal(result, 'routed');
    assert.equal(routed.length, 1);
    assert.equal(routed[0].payload.attachments?.length, 1);
    // Mode-2 success: empty text would 400 even though the attachment uploaded.
    assert.notEqual(routed[0].payload.message.trim(), '');
    assert.match(routed[0].payload.message, /attachment/i);
    assert.doesNotMatch(routed[0].payload.message, /could not be retrieved/i);
  });

  it('Part B — attachment-only (empty text) whose attachments ALL FAIL: routes a non-empty "could not retrieve" placeholder, no attachments field', async () => {
    const routed: AMPRouteRequest[] = [];
    installFetch({ routed });
    const failingDownloader = async () => { throw new Error('download failed (401)'); };
    const result = await handleInbound(activity('', [file], failingDownloader), deps({ operatorAadObjectIds: OPERATOR }));

    assert.equal(result, 'routed');
    assert.equal(routed.length, 1);
    assert.notEqual(routed[0].payload.message.trim(), ''); // never the empty-message 400
    assert.match(routed[0].payload.message, /could not be retrieved/i);
    assert.equal(routed[0].payload.attachments, undefined);
  });

  it('Part B — placeholder only fires on EMPTY text: a caption survives unchanged even when the attachment fails', async () => {
    const routed: AMPRouteRequest[] = [];
    installFetch({ routed });
    const failingDownloader = async () => { throw new Error('download failed (401)'); };
    const result = await handleInbound(activity('look at this photo', [file], failingDownloader), deps());

    assert.equal(result, 'routed');
    assert.match(routed[0].payload.message, /look at this photo/);
    assert.doesNotMatch(routed[0].payload.message, /could not be retrieved/i);
  });

  it('Part B invariant — no attachment-bearing message ever routes an empty payload.message', async () => {
    const sink: AMPRouteRequest[] = [];
    const cases: Array<() => Promise<void>> = [
      async () => { installFetch({ routed: sink }); await handleInbound(activity('', [file], goodDownloader), deps({ operatorAadObjectIds: OPERATOR })); },
      async () => { installFetch({ routed: sink }); await handleInbound(activity('   ', [file], async () => { throw new Error('boom'); }), deps({ operatorAadObjectIds: OPERATOR })); },
    ];
    for (const run of cases) await run();
    assert.equal(sink.length, 2);
    for (const r of sink) assert.notEqual(r.payload.message.trim(), '');
  });

  it('FAIL-OPEN: a failed ingest still routes the text with NO attachments field', async () => {
    const routed: AMPRouteRequest[] = [];
    installFetch({ routed });
    const failingDownloader = async () => { throw new Error('teams 403'); };
    const result = await handleInbound(activity('text must survive', [file], failingDownloader), deps());

    assert.equal(result, 'routed');
    assert.equal(routed.length, 1);
    assert.match(routed[0].payload.message, /text must survive/);
    assert.equal(routed[0].payload.attachments, undefined); // dropped, not cited
  });

  it('attachments present but no downloader -> routes text-only (no throw)', async () => {
    const routed: AMPRouteRequest[] = [];
    installFetch({ routed });
    const result = await handleInbound(activity('no downloader here', [file], undefined), deps());

    assert.equal(result, 'routed');
    assert.equal(routed[0].payload.attachments, undefined);
  });

  it('a text-only message (no attachments) routes unchanged, no attachments field', async () => {
    const routed: AMPRouteRequest[] = [];
    installFetch({ routed });
    const result = await handleInbound(activity('just text'), deps());

    assert.equal(result, 'routed');
    assert.equal(routed[0].payload.attachments, undefined);
  });
});
