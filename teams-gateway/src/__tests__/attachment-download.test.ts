/**
 * w3 follow-up — inbound attachment byte download + connector-token auth
 * (attachment-download.ts). Unit-tests the auth/retry policy with a mock `fetch`
 * and an injected token-getter — no live Azure.
 *
 * Asserts: a connector `contentUrl` is fetched WITH a Bearer token up front; a
 * pre-auth `downloadUrl` goes bare; a `data:` URI decodes with NO network; an
 * unexpected 401 on a bare path retries ONCE with the token; and FAIL-OPEN — a
 * null token or a still-401 authed attempt throws (the caller drops the attachment).
 * Also pins `redirect: 'error'` on the GET (connector parity with outbound).
 */

import assert from 'node:assert/strict';
import { afterEach, beforeEach, describe, it } from 'node:test';

import { downloadAttachmentBytes, type AttachmentDownloadDeps } from '../attachment-download.js';
import type { RawInboundAttachment } from '../attachments-inbound.js';

interface Call {
  url: string;
  auth: string | null;
  redirect: RequestInit['redirect'];
}

const origFetch = globalThis.fetch;
const origError = console.error;
let calls: Call[];

beforeEach(() => {
  calls = [];
  console.error = () => undefined;
});
afterEach(() => {
  globalThis.fetch = origFetch;
  console.error = origError;
});

/** Mock fetch that records each call and returns the queued status (200 by default). */
function installFetch(statuses: number[] = [200], bytes = new Uint8Array([1, 2, 3, 4])): void {
  let i = 0;
  globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    const headers = new Headers(init?.headers);
    calls.push({ url: String(input), auth: headers.get('authorization'), redirect: init?.redirect });
    const status = statuses[Math.min(i, statuses.length - 1)];
    i += 1;
    if (status === 200) return new Response(bytes, { status: 200 });
    return new Response(null, { status });
  }) as typeof fetch;
}

function deps(getConnectorToken: AttachmentDownloadDeps['getConnectorToken']): AttachmentDownloadDeps {
  return { getConnectorToken, timeoutMs: 1000, slug: 'leoai' };
}

const tokenOk = async () => 'CONNECTOR.JWT.TOKEN';
const tokenNull = async () => null;

const connectorAtt: RawInboundAttachment = {
  name: 'image.png',
  contentType: 'image/png',
  contentUrl: 'https://smba.trafficmanager.net/amer/v3/attachments/abc/views/original',
};
const fileAtt: RawInboundAttachment = {
  name: 'doc.pdf',
  contentType: 'application/pdf',
  downloadUrl: 'https://teams.test/files/doc.pdf?token=preauth',
};

describe('w3 attachment-download — connector-token auth', () => {
  it('connector contentUrl is fetched WITH a Bearer token up front', async () => {
    installFetch([200]);
    const bytes = await downloadAttachmentBytes(connectorAtt, deps(tokenOk));

    assert.equal(calls.length, 1);
    assert.equal(calls[0].auth, 'Bearer CONNECTOR.JWT.TOKEN');
    assert.equal(calls[0].redirect, 'error');
    assert.equal(bytes.byteLength, 4);
  });

  it('pre-auth downloadUrl goes bare (no Authorization header)', async () => {
    installFetch([200]);
    await downloadAttachmentBytes(fileAtt, deps(tokenOk));

    assert.equal(calls.length, 1);
    assert.equal(calls[0].auth, null);
    assert.equal(calls[0].url, fileAtt.downloadUrl);
  });

  it('data: URI decodes locally — no network call', async () => {
    installFetch([200]);
    const dataAtt: RawInboundAttachment = {
      name: 'inline.txt',
      contentType: 'text/plain',
      contentUrl: 'data:text/plain;base64,' + Buffer.from('hello').toString('base64'),
    };
    const bytes = await downloadAttachmentBytes(dataAtt, deps(tokenNull));

    assert.equal(calls.length, 0);
    assert.equal(Buffer.from(bytes).toString('utf8'), 'hello');
  });

  it('safety net: an unexpected 401 on a bare pre-auth path retries ONCE with the token', async () => {
    installFetch([401, 200]);
    const bytes = await downloadAttachmentBytes(fileAtt, deps(tokenOk));

    assert.equal(calls.length, 2);
    assert.equal(calls[0].auth, null); // first bare
    assert.equal(calls[1].auth, 'Bearer CONNECTOR.JWT.TOKEN'); // retried with token
    assert.equal(bytes.byteLength, 4);
  });

  it('FAIL-OPEN: a null connector token on the connector path 401s and throws (no retry token to add)', async () => {
    installFetch([401]);
    await assert.rejects(
      downloadAttachmentBytes(connectorAtt, deps(tokenNull)),
      /download failed \(401\)/,
    );
    assert.equal(calls.length, 1); // bare attempt; no token available to retry
    assert.equal(calls[0].auth, null);
  });

  it('FAIL-OPEN: a 401 even WITH the token throws (caller drops the attachment)', async () => {
    installFetch([401, 401]);
    await assert.rejects(
      downloadAttachmentBytes(connectorAtt, deps(tokenOk)),
      /download failed \(401\)/,
    );
    // connector path authed up front (call 1 has token); no second auth attempt
    // because the up-front attempt already tried auth.
    assert.equal(calls.length, 1);
    assert.equal(calls[0].auth, 'Bearer CONNECTOR.JWT.TOKEN');
  });

  it('throws when an attachment has no usable url', async () => {
    installFetch([200]);
    await assert.rejects(
      downloadAttachmentBytes({ name: 'x', contentType: 'image/png' }, deps(tokenOk)),
      /no downloadUrl\/contentUrl/,
    );
    assert.equal(calls.length, 0);
  });
});
