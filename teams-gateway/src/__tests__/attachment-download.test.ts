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
  return { getConnectorToken, timeoutMs: 1000, bodyStallMs: 200, slug: 'echo' };
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

/**
 * w3 body-transfer resilience — the connector intermittently STALLS the response
 * BODY after headers. The download streams the body under a per-chunk stall
 * timeout and retries ONCE on a stall/transient, then fails open (caller drops
 * the attachment, text still routes). Bodies are mocked as ReadableStreams whose
 * chunk cadence we control; `bodyStallMs` is small so the suite runs fast.
 */
describe('w3 attachment-download — body-transfer resilience', () => {
  const enc = new TextEncoder();

  /**
   * Build a streaming body. Emits `chunks` with `gapMs` between each; if `stall`
   * is set, after the chunks the stream HANGS (no further bytes) so a reader idles
   * and the per-chunk stall timer fires. The hang is settled cleanly on `cancel()`
   * (and pending timers cleared) so the test leaves no dangling promise.
   */
  function streamBody(
    chunks: Uint8Array[],
    opts: { gapMs?: number; stall?: boolean } = {},
  ): ReadableStream<Uint8Array> {
    let i = 0;
    let settlePending: (() => void) | null = null;
    let timer: ReturnType<typeof setTimeout> | null = null;
    const clear = () => {
      if (timer) {
        clearTimeout(timer);
        timer = null;
      }
    };
    return new ReadableStream<Uint8Array>({
      pull(controller) {
        if (i < chunks.length) {
          const chunk = chunks[i];
          i += 1;
          return new Promise<void>((resolve) => {
            settlePending = resolve;
            timer = setTimeout(() => {
              timer = null;
              settlePending = null;
              controller.enqueue(chunk);
              resolve();
            }, opts.gapMs ?? 0);
          });
        }
        if (opts.stall) {
          // Hang until the reader is cancelled (the stall timer fires upstream).
          return new Promise<void>((resolve) => {
            settlePending = resolve;
          });
        }
        controller.close();
        return undefined;
      },
      cancel() {
        clear();
        if (settlePending) {
          const resolve = settlePending;
          settlePending = null;
          resolve();
        }
      },
    });
  }

  /** Mock fetch that returns a 200 with the queued streaming body, one per call. */
  function installBodyFetch(bodies: ReadableStream<Uint8Array>[]): void {
    let i = 0;
    globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
      const headers = new Headers(init?.headers);
      calls.push({ url: String(input), auth: headers.get('authorization'), redirect: init?.redirect });
      const body = bodies[Math.min(i, bodies.length - 1)];
      i += 1;
      return new Response(body, { status: 200 });
    }) as typeof fetch;
  }

  it('body stall on the first attempt → retries once → succeeds', async () => {
    // Attempt 1 delivers one chunk then hangs; attempt 2 streams cleanly.
    installBodyFetch([
      streamBody([enc.encode('part-')], { stall: true }),
      streamBody([enc.encode('hello'), enc.encode('-world')]),
    ]);

    const bytes = await downloadAttachmentBytes(fileAtt, deps(tokenOk));

    assert.equal(calls.length, 2); // one retry
    assert.equal(Buffer.from(bytes).toString('utf8'), 'hello-world');
  });

  it('truly-hung body → fails open within a BOUNDED time (not ~unbounded)', async () => {
    installBodyFetch([
      streamBody([], { stall: true }),
      streamBody([], { stall: true }),
    ]);

    const start = Date.now();
    await assert.rejects(downloadAttachmentBytes(connectorAtt, deps(tokenOk)));
    const elapsed = Date.now() - start;

    assert.equal(calls.length, 2); // exactly two attempts (one transient retry), no unbounded loop
    // Bound ≈ 2 × bodyStallMs(200) ≈ 400ms; assert well under a multiple — proves
    // it does NOT hang for ~unbounded time. (Real prod bound ≈ 2 × 8s.)
    assert.ok(elapsed < 2000, `fail-open took ${elapsed}ms — expected bounded (~2×stall)`);
  });

  it('slow-but-progressing body (gap < stall) → succeeds, no retry', async () => {
    // Chunks 60ms apart, under the 200ms stall window — progress, not a stall.
    installBodyFetch([
      streamBody([enc.encode('a'), enc.encode('b'), enc.encode('c')], { gapMs: 60 }),
    ]);

    const bytes = await downloadAttachmentBytes(fileAtt, deps(tokenOk));

    assert.equal(calls.length, 1); // no retry — slow is not stalled
    assert.equal(Buffer.from(bytes).toString('utf8'), 'abc');
  });

  it('a header-phase transport error retries once then succeeds', async () => {
    let i = 0;
    globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
      const headers = new Headers(init?.headers);
      calls.push({ url: String(input), auth: headers.get('authorization'), redirect: init?.redirect });
      i += 1;
      if (i === 1) throw new TypeError('network reset'); // transient on first attempt
      return new Response(streamBody([enc.encode('ok-bytes')]), { status: 200 });
    }) as typeof fetch;

    const bytes = await downloadAttachmentBytes(fileAtt, deps(tokenOk));

    assert.equal(calls.length, 2); // retried the transient
    assert.equal(Buffer.from(bytes).toString('utf8'), 'ok-bytes');
  });
});

/**
 * w3 follow-up — resource-leak teardown. After headers arrive, TWO paths abandon a
 * `Response` WITHOUT consuming its body: (a) the 401/403 token-retry `continue` and
 * (b) the non-ok fail-open `throw`. An un-consumed body leaks the socket to
 * undici/GC. The download must `cancel()` the body before EITHER path. Bodies here
 * track their own `cancel()` count so we can assert the abandoned body was released.
 */
describe('w3 attachment-download — abandoned-Response body teardown', () => {
  const enc = new TextEncoder();

  /** A streaming body that streams `bytes` and counts how many times it is cancelled. */
  function countingBody(
    bytes: Uint8Array | null,
  ): { stream: ReadableStream<Uint8Array>; cancelCount: () => number } {
    let count = 0;
    const stream = new ReadableStream<Uint8Array>({
      pull(controller) {
        if (bytes) controller.enqueue(bytes);
        controller.close();
      },
      cancel() {
        count += 1;
      },
    });
    return { stream, cancelCount: () => count };
  }

  it('401 token-retry: the FIRST (401) response body is cancelled before the retry', async () => {
    // Bare pre-auth path unexpectedly 401s WITH a body, then succeeds with the token.
    const first = countingBody(enc.encode('401-error-page')); // abandoned on retry
    const second = countingBody(enc.encode('ok')); // consumed on the 2nd attempt
    let i = 0;
    globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
      const headers = new Headers(init?.headers);
      calls.push({ url: String(input), auth: headers.get('authorization'), redirect: init?.redirect });
      i += 1;
      return i === 1
        ? new Response(first.stream, { status: 401 })
        : new Response(second.stream, { status: 200 });
    }) as typeof fetch;

    const bytes = await downloadAttachmentBytes(fileAtt, deps(tokenOk));

    assert.equal(calls.length, 2); // 401 then token retry
    assert.ok(first.cancelCount() >= 1, 'the abandoned 401 body must be cancelled, not leaked');
    assert.equal(Buffer.from(bytes).toString('utf8'), 'ok');
  });

  it('non-ok fail-open: the final response body is cancelled before the throw', async () => {
    // Connector path authed up front → single 401 (no further auth retry) → fail open.
    const only = countingBody(enc.encode('403-or-401-page'));
    globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
      const headers = new Headers(init?.headers);
      calls.push({ url: String(input), auth: headers.get('authorization'), redirect: init?.redirect });
      return new Response(only.stream, { status: 401 });
    }) as typeof fetch;

    await assert.rejects(
      downloadAttachmentBytes(connectorAtt, deps(tokenOk)),
      /download failed \(401\)/,
    );

    assert.equal(calls.length, 1); // authed up front; no second auth attempt
    assert.ok(only.cancelCount() >= 1, 'the abandoned non-ok body must be cancelled, not leaked');
  });
});
