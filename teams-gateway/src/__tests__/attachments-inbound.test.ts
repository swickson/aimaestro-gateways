/**
 * w3 — inbound attachment ingestion (attachments-inbound.ts).
 *
 * Exercises the upload -> PUT -> confirm -> status flow with a mock `fetch`
 * sequencer + a mock downloader. NO live Azure/Maestro. Asserts: the call sequence
 * + auth (bot Bearer on upload/confirm/status, NO Bearer on PUT), the PUT body
 * length == declared size, the sha256 digest, policy (deny/size/count), and the
 * FAIL-OPEN contract (a failed leg drops just that attachment, never throws).
 */

import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { afterEach, beforeEach, describe, it } from 'node:test';

import { ingestAttachments, type IngestDeps, type RawInboundAttachment } from '../attachments-inbound.js';
import type { AttachmentPolicy } from '../types.js';

const POLICY: AttachmentPolicy = { maxBytes: 1000, maxCount: 3, denyContentTypes: ['application/x-msdownload'] };

interface RecordedCall {
  url: string;
  method: string;
  auth?: string;
  bodyText?: string;
  bodyBytes?: number;
}

/** A scripted fetch: maps URL substrings to canned Responses + records each call. */
function installFetch(handlers: Array<{ match: string; respond: () => Response }>): RecordedCall[] {
  const calls: RecordedCall[] = [];
  globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = String(input);
    const method = init?.method ?? 'GET';
    const auth = (init?.headers as Record<string, string> | undefined)?.['Authorization'];
    const call: RecordedCall = { url, method, auth };
    if (init?.body instanceof Blob) call.bodyBytes = init.body.size;
    else if (typeof init?.body === 'string') call.bodyText = init.body;
    calls.push(call);
    const handler = handlers.find((h) => url.includes(h.match));
    if (!handler) throw new Error(`no mock handler for ${method} ${url}`);
    return handler.respond();
  }) as typeof fetch;
  return calls;
}

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
}

function statusBody(id: string, overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    attachment_id: id,
    filename: 'safe-name.pdf',
    content_type: 'application/pdf',
    size: 4,
    digest: 'server-digest',
    scan_status: 'basic_clean',
    uploaded_at: '2026-06-14T00:00:00.000Z',
    expires_at: '2026-06-15T00:00:00.000Z',
    url: 'https://maestro.test/api/v1/attachments/' + id + '/download?sig=abc',
    ...overrides,
  };
}

function happyHandlers(id = 'att-1'): Array<{ match: string; respond: () => Response }> {
  return [
    { match: '/attachments/upload', respond: () => json({ attachment_id: id, upload_url: 'https://maestro.test/signed/put/' + id }) },
    { match: '/signed/put/', respond: () => new Response(null, { status: 200 }) },
    { match: '/confirm', respond: () => json({ ok: true }) },
    { match: '/status', respond: () => json(statusBody(id)) },
  ];
}

function deps(overrides: Partial<IngestDeps> = {}): IngestDeps {
  return {
    maestroUrl: 'https://maestro.test',
    ampApiKey: 'bot-key',
    botSlug: 'maestro',
    policy: POLICY,
    timeoutMs: 1000,
    downloadAttachment: async () => new Uint8Array([1, 2, 3, 4]),
    ...overrides,
  };
}

function file(overrides: Partial<RawInboundAttachment> = {}): RawInboundAttachment {
  return { name: 'doc.pdf', contentType: 'application/pdf', downloadUrl: 'https://teams.test/dl/doc', ...overrides };
}

const origFetch = globalThis.fetch;
const origError = console.error;
const origLog = console.log;
beforeEach(() => {
  console.log = () => undefined;
  console.error = () => undefined;
});
afterEach(() => {
  globalThis.fetch = origFetch;
  console.error = origError;
  console.log = origLog;
});

describe('w3 inbound attachment ingestion', () => {
  it('happy path: upload -> PUT -> confirm -> status, returns the status-assembled AMPAttachmentV1', async () => {
    const calls = installFetch(happyHandlers('att-1'));
    const res = await ingestAttachments([file()], deps());

    assert.equal(res.failed, 0);
    assert.equal(res.skipped, 0);
    assert.equal(res.attachments.length, 1);
    const a = res.attachments[0];
    // Descriptor is assembled from /status (the only call yielding url + sanitized filename).
    assert.deepEqual(a, {
      kind: 'amp-v1',
      id: 'att-1',
      filename: 'safe-name.pdf',
      content_type: 'application/pdf',
      size: 4,
      digest: 'server-digest',
      url: 'https://maestro.test/api/v1/attachments/att-1/download?sig=abc',
      scan_status: 'basic_clean',
      uploaded_at: '2026-06-14T00:00:00.000Z',
      expires_at: '2026-06-15T00:00:00.000Z',
    });

    // Exact call sequence.
    assert.deepEqual(
      calls.map((c) => c.method + ' ' + new URL(c.url).pathname.replace('/api/v1/attachments', '')),
      ['POST /upload', 'PUT /signed/put/att-1', 'POST /att-1/confirm', 'GET /att-1/status'],
    );
  });

  it('auth: bot Bearer on upload/confirm/status, NONE on the signed PUT', async () => {
    const calls = installFetch(happyHandlers('att-1'));
    await ingestAttachments([file()], deps());

    const byPath = (frag: string) => calls.find((c) => c.url.includes(frag))!;
    assert.equal(byPath('/upload').auth, 'Bearer bot-key');
    assert.equal(byPath('/confirm').auth, 'Bearer bot-key');
    assert.equal(byPath('/status').auth, 'Bearer bot-key');
    assert.equal(byPath('/signed/put/').auth, undefined); // signed url IS the auth
  });

  it('declares the real byte length on PUT and the sha256 digest on /upload', async () => {
    const bytes = new Uint8Array([10, 20, 30, 40, 50]); // 5 bytes
    const calls = installFetch(happyHandlers('att-1'));
    await ingestAttachments([file()], deps({ downloadAttachment: async () => bytes }));

    const upload = JSON.parse(calls.find((c) => c.url.includes('/upload'))!.bodyText!);
    assert.equal(upload.size, 5);
    assert.equal(upload.digest, createHash('sha256').update(bytes).digest('hex'));
    // PUT body length must equal the declared size (Maestro 400s otherwise).
    assert.equal(calls.find((c) => c.url.includes('/signed/put/'))!.bodyBytes, 5);
  });

  it('FAIL-OPEN: a download failure drops the attachment (failed=1), never throws', async () => {
    installFetch(happyHandlers());
    const res = await ingestAttachments([file()], deps({ downloadAttachment: async () => { throw new Error('teams 403'); } }));
    assert.equal(res.attachments.length, 0);
    assert.equal(res.failed, 1);
  });

  it('FAIL-OPEN: a confirm 422 (rejected) drops the attachment, never throws', async () => {
    installFetch([
      { match: '/attachments/upload', respond: () => json({ attachment_id: 'x', upload_url: 'https://maestro.test/signed/put/x' }) },
      { match: '/signed/put/', respond: () => new Response(null, { status: 200 }) },
      { match: '/confirm', respond: () => json({ error: 'rejected' }, 422) },
      { match: '/status', respond: () => json(statusBody('x')) },
    ]);
    const res = await ingestAttachments([file()], deps());
    assert.equal(res.attachments.length, 0);
    assert.equal(res.failed, 1);
  });

  it('FAIL-OPEN: status without a signed url drops the attachment', async () => {
    installFetch([
      { match: '/attachments/upload', respond: () => json({ attachment_id: 'x', upload_url: 'https://maestro.test/signed/put/x' }) },
      { match: '/signed/put/', respond: () => new Response(null, { status: 200 }) },
      { match: '/confirm', respond: () => json({ ok: true }) },
      { match: '/status', respond: () => json(statusBody('x', { url: '', scan_status: 'pending' })) },
    ]);
    const res = await ingestAttachments([file()], deps());
    assert.equal(res.attachments.length, 0);
    assert.equal(res.failed, 1);
  });

  it('one failed attachment does not block a healthy sibling (partial success)', async () => {
    // First file downloads fine; second throws on download. Healthy one still ingests.
    let n = 0;
    installFetch(happyHandlers('att-1'));
    const res = await ingestAttachments(
      [file({ name: 'good.pdf' }), file({ name: 'bad.pdf' })],
      deps({
        downloadAttachment: async () => {
          n += 1;
          if (n === 2) throw new Error('download blew up');
          return new Uint8Array([1, 2, 3, 4]);
        },
      }),
    );
    assert.equal(res.attachments.length, 1);
    assert.equal(res.failed, 1);
  });

  it('policy: deny-listed content type is skipped (no network)', async () => {
    const calls = installFetch(happyHandlers());
    const res = await ingestAttachments([file({ contentType: 'application/x-msdownload', name: 'evil.exe' })], deps());
    assert.equal(res.skipped, 1);
    assert.equal(res.attachments.length, 0);
    assert.equal(calls.length, 0); // never touched the network
  });

  it('policy: over-cap size is skipped (not a failure)', async () => {
    installFetch(happyHandlers());
    const big = new Uint8Array(POLICY.maxBytes + 1);
    const res = await ingestAttachments([file()], deps({ downloadAttachment: async () => big }));
    assert.equal(res.skipped, 1);
    assert.equal(res.failed, 0);
    assert.equal(res.attachments.length, 0);
  });

  it('policy: count cap drops the overflow (skipped) and ingests the rest', async () => {
    installFetch(happyHandlers('att-1'));
    const five = Array.from({ length: 5 }, (_, i) => file({ name: `f${i}.pdf` }));
    const res = await ingestAttachments(five, deps()); // maxCount = 3
    assert.equal(res.attachments.length, 3);
    assert.equal(res.skipped, 2);
  });

  it('empty input is a no-op', async () => {
    const calls = installFetch(happyHandlers());
    const res = await ingestAttachments([], deps());
    assert.deepEqual(res, { attachments: [], failed: 0, skipped: 0 });
    assert.equal(calls.length, 0);
  });
});
