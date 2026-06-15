/**
 * w5 cross-review (Crease, correctness lens) — inbound CITED-descriptor digest
 * normalization in attachments-inbound.ts (`normalizeDescriptorDigest`).
 *
 * Written FRESH from the dispatch spec, blind to whistler's own tests. Exercises
 * the REAL ingestAttachments pipeline (mock fetch + mock downloader, the module's
 * SDK-decoupled design) and asserts the cited AMPAttachmentV1.digest the gateway
 * emits on /route. Confirms:
 *   (1) bare /status hex            -> sha256:<hex>      (normalized)
 *   (2) already sha256:-prefixed    -> unchanged         (idempotent, no sha256:sha256:)
 *   (3) the prefixed value carries the REAL sha256(bytes) hex unchanged
 *       (integrity not weakened — amp-download bytes-match still holds)
 */

import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { afterEach, beforeEach, describe, it } from 'node:test';

import { ingestAttachments, type IngestDeps, type RawInboundAttachment } from '../attachments-inbound.js';
import type { AttachmentPolicy } from '../types.js';

const POLICY: AttachmentPolicy = { maxBytes: 1000, maxCount: 3, denyContentTypes: [] };
const BYTES = new Uint8Array([9, 8, 7, 6, 5]);
const REAL_HEX = createHash('sha256').update(BYTES).digest('hex');

/** Scripted fetch: maps URL substrings to canned Responses. */
function installFetch(statusDigest: string): void {
  const handlers: Array<{ match: string; respond: () => Response }> = [
    { match: '/attachments/upload', respond: () => json({ attachment_id: 'a1', upload_url: 'https://m.test/signed/put/a1' }) },
    { match: '/signed/put/', respond: () => new Response(null, { status: 200 }) },
    { match: '/confirm', respond: () => json({ ok: true }) },
    {
      match: '/status',
      respond: () =>
        json({
          attachment_id: 'a1',
          filename: 'safe.pdf',
          content_type: 'application/pdf',
          size: BYTES.byteLength,
          digest: statusDigest,
          scan_status: 'basic_clean',
          uploaded_at: '2026-06-14T00:00:00.000Z',
          expires_at: '2026-06-15T00:00:00.000Z',
          url: 'https://m.test/api/v1/attachments/a1/download?sig=z',
        }),
    },
  ];
  globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = String(input);
    const handler = handlers.find((h) => url.includes(h.match));
    if (!handler) throw new Error(`no mock handler for ${init?.method ?? 'GET'} ${url}`);
    return handler.respond();
  }) as typeof fetch;
}

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
}

function deps(): IngestDeps {
  return {
    maestroUrl: 'https://m.test',
    ampApiKey: 'k',
    botSlug: 'maestro',
    policy: POLICY,
    timeoutMs: 1000,
    downloadAttachment: async () => BYTES,
  };
}

const FILE: RawInboundAttachment = { name: 'doc.pdf', contentType: 'application/pdf', downloadUrl: 'https://t.test/dl' };

const origFetch = globalThis.fetch;
const origLog = console.log;
const origErr = console.error;
beforeEach(() => {
  console.log = () => undefined;
  console.error = () => undefined;
});
afterEach(() => {
  globalThis.fetch = origFetch;
  console.log = origLog;
  console.error = origErr;
});

describe('w5 cited-descriptor digest normalization (Crease correctness cross-review)', () => {
  it('(1) bare /status hex -> the cited descriptor digest is sha256:<hex>', async () => {
    installFetch(REAL_HEX);
    const res = await ingestAttachments([FILE], deps());

    assert.equal(res.attachments.length, 1);
    assert.equal(res.attachments[0].digest, `sha256:${REAL_HEX}`);
  });

  it('(2) already sha256:-prefixed /status digest is left unchanged (no sha256:sha256:)', async () => {
    installFetch(`sha256:${REAL_HEX}`);
    const res = await ingestAttachments([FILE], deps());

    assert.equal(res.attachments.length, 1);
    assert.equal(res.attachments[0].digest, `sha256:${REAL_HEX}`);
    assert.ok(!res.attachments[0].digest.includes('sha256:sha256:'), 'must not double-prefix');
    // idempotent: prefixing the already-prefixed wire form yields the same string.
    assert.equal(res.attachments[0].digest.match(/sha256:/g)?.length, 1);
  });

  it('(3) the prefixed value carries the REAL sha256(bytes) hex unchanged — bytes-match holds, no integrity weakening', async () => {
    installFetch(REAL_HEX);
    const res = await ingestAttachments([FILE], deps());

    const cited = res.attachments[0].digest;
    assert.ok(cited.startsWith('sha256:'), 'cited descriptor is in sha256: wire form');
    const hexPart = cited.slice('sha256:'.length);
    // The hex an amp-download peer would compare against == the actual sha256 of the bytes.
    assert.equal(hexPart, createHash('sha256').update(BYTES).digest('hex'));
  });
});
