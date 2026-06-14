/**
 * w3 — outbound attachment delivery (outbound.ts).
 *
 * Drives the real poller against a temp inbox with a mock signed-url `fetch` + a
 * capturing `send`. Covers the outbound.ts:139 fix (attachment-only reply DELIVERS
 * instead of being deleted as "empty"), the no-Bearer signed-url pull, and the
 * item-8 failure policy (attachment-only pull-fail LEAVES for retry; text + failed
 * attachment still delivers the text).
 */

import assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { afterEach, describe, it } from 'node:test';

import { startOutboundPoller, type OutboundAttachment, type OutboundBot } from '../outbound.js';
import { createThreadStore, type ThreadEntry } from '../thread-store.js';
import type { AMPAttachmentV1, AMPMessage, AttachmentPolicy, ThreadContext } from '../types.js';

/** Default outbound test policy: 25MB cap, 10 attachments, no deny-list. */
const DEFAULT_TEST_POLICY: AttachmentPolicy = {
  maxBytes: 26_214_400,
  maxCount: 10,
  denyContentTypes: [],
};

const tempRoots: string[] = [];
const origFetch = globalThis.fetch;
const origLog = console.log;
const origError = console.error;

afterEach(() => {
  for (const dir of tempRoots.splice(0)) fs.rmSync(dir, { recursive: true, force: true });
  globalThis.fetch = origFetch;
  console.log = origLog;
  console.error = origError;
});

function tempDir(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-attach-out-'));
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
    conversationId: 'conversation-maestro',
    ampMessageId: 'amp-inbound-1',
    aadObjectId: 'aad-1',
    context: context('conversation-maestro'),
    createdAt: 100,
  };
}

function attachment(overrides: Partial<AMPAttachmentV1> = {}): AMPAttachmentV1 {
  return {
    kind: 'amp-v1',
    id: 'att-1',
    filename: 'report.pdf',
    content_type: 'application/pdf',
    size: 4,
    digest: 'd',
    url: 'https://maestro.test/api/v1/attachments/att-1/download?sig=abc',
    scan_status: 'basic_clean',
    uploaded_at: '2026-06-14T00:00:00.000Z',
    expires_at: '2026-06-15T00:00:00.000Z',
    ...overrides,
  };
}

function writeReply(inboxDir: string, message: string, attachments?: AMPAttachmentV1[]): string {
  const senderDir = path.join(inboxDir, 'agent-one_example_aimaestro_local');
  fs.mkdirSync(senderDir, { recursive: true });
  const filePath = path.join(senderDir, 'reply.json');
  const msg = {
    envelope: { id: 'reply-1', from: 'agent-one@example.aimaestro.local', to: 'teams-maestro-bot@example.aimaestro.local', timestamp: '2026-06-14T00:00:00.000Z', in_reply_to: 'amp-inbound-1' },
    payload: { type: 'response', message, ...(attachments ? { attachments } : {}) },
  } as AMPMessage;
  fs.writeFileSync(filePath, JSON.stringify(msg), 'utf-8');
  return filePath;
}

/**
 * Write a reply whose `payload.attachments` is an ARBITRARY array — including hostile
 * non-object elements (`null`, strings, nested arrays) that the typed `writeReply`
 * cannot express. `payload.attachments` is agent-controlled JSON, so the gateway must
 * survive any element shape without throwing.
 */
function writeRawReply(inboxDir: string, message: string, attachments: unknown[]): string {
  const senderDir = path.join(inboxDir, 'agent-one_example_aimaestro_local');
  fs.mkdirSync(senderDir, { recursive: true });
  const filePath = path.join(senderDir, 'reply.json');
  const msg = {
    envelope: { id: 'reply-1', from: 'agent-one@example.aimaestro.local', to: 'teams-maestro-bot@example.aimaestro.local', timestamp: '2026-06-14T00:00:00.000Z', in_reply_to: 'amp-inbound-1' },
    payload: { type: 'response', message, attachments },
  };
  fs.writeFileSync(filePath, JSON.stringify(msg), 'utf-8');
  return filePath;
}

interface SentRecord {
  text: string;
  attachments?: OutboundAttachment[];
}

/** Capturing bot + a record of signed-url fetches (to assert no Bearer). */
function setup(): { bot: OutboundBot; sends: SentRecord[]; fetched: Array<{ url: string; auth?: string }>; inbox: string } {
  const inbox = path.join(tempDir(), 'maestro-inbox');
  const sends: SentRecord[] = [];
  const bot: OutboundBot = {
    slug: 'maestro',
    inboxDir: inbox,
    maestroUrl: 'https://maestro.test',
    send: async (_conversationId, text, _markdown, attachments) => {
      sends.push({ text, attachments });
    },
  };
  const fetched: Array<{ url: string; auth?: string }> = [];
  globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = String(input);
    fetched.push({ url, auth: (init?.headers as Record<string, string> | undefined)?.['Authorization'] });
    if (url.includes('FAIL')) return new Response('nope', { status: 404 });
    return new Response(new Uint8Array([1, 2, 3, 4]), { status: 200 });
  }) as typeof fetch;
  return { bot, sends, fetched, inbox };
}

async function runPoller(bot: OutboundBot, policy: AttachmentPolicy = DEFAULT_TEST_POLICY): Promise<void> {
  const store = createThreadStore({ maxAgeMs: Infinity });
  store.record(entry());
  const stop = startOutboundPoller({ bots: [bot], threadStore: store, pollIntervalMs: 60_000, markdownDefault: true, policy, debug: false });
  try {
    await waitFor(() => true, 'first tick');
    await new Promise((r) => setTimeout(r, 40));
  } finally {
    stop();
  }
}

async function waitFor(predicate: () => boolean, label: string): Promise<void> {
  const deadline = Date.now() + 2000;
  while (Date.now() < deadline) {
    if (predicate()) return;
    await new Promise((r) => setTimeout(r, 10));
  }
  throw new Error(`timed out waiting for ${label}`);
}

describe('w3 outbound attachment delivery', () => {
  it('attachment-only reply (empty text) DELIVERS the attachment and deletes the file (the :139 fix)', async () => {
    console.log = () => undefined;
    const { bot, sends, inbox } = setup();
    const filePath = writeReply(inbox, '   ', [attachment()]);
    await runPoller(bot);

    // One send, carrying the attachment, with no text bubble.
    const withAttach = sends.filter((s) => s.attachments && s.attachments.length > 0);
    assert.equal(withAttach.length, 1);
    assert.equal(withAttach[0].attachments![0].filename, 'report.pdf');
    assert.equal(withAttach[0].attachments![0].bytes.byteLength, 4);
    // File consumed (NOT left as "empty").
    assert.equal(fs.existsSync(filePath), false);
  });

  it('text + attachment: delivers the text chunk AND the attachment, then deletes', async () => {
    console.log = () => undefined;
    const { bot, sends, inbox } = setup();
    const filePath = writeReply(inbox, 'here is the file', [attachment()]);
    await runPoller(bot);

    assert.ok(sends.some((s) => /here is the file/.test(s.text)), 'text chunk delivered');
    assert.ok(sends.some((s) => (s.attachments?.length ?? 0) > 0), 'attachment delivered');
    assert.equal(fs.existsSync(filePath), false);
  });

  it('pulls the signed url with NO Bearer (the signed url is the auth)', async () => {
    console.log = () => undefined;
    const { bot, fetched, inbox } = setup();
    writeReply(inbox, 'x', [attachment()]);
    await runPoller(bot);

    const pull = fetched.find((f) => f.url.includes('/download'));
    assert.ok(pull, 'signed-url GET happened');
    assert.equal(pull!.auth, undefined);
  });

  it('does not fetch attachment URLs outside the configured Maestro origin (validation DROP)', async () => {
    console.log = () => undefined;
    console.error = () => undefined;
    const { bot, sends, fetched, inbox } = setup();
    const filePath = writeReply(inbox, '', [attachment({ url: 'https://evil.test/internal' })]);
    await runPoller(bot);

    assert.equal(fetched.length, 0); // never even fetched — off-origin descriptor rejected pre-pull
    assert.equal(sends.length, 0);
    // Hardening change: an off-origin descriptor is a VALIDATION failure (hostile/malformed),
    // not a transient pull failure — attachment-only-all-invalid is DROPPED, not left to spin.
    assert.equal(fs.existsSync(filePath), false);
  });

  it('attachment-only whose VALID descriptor fails the network pull is left for retry', async () => {
    console.log = () => undefined;
    console.error = () => undefined;
    const { bot, sends, inbox } = setup();
    // Valid same-origin /download path (passes validation) whose pull 404s (mock keys on FAIL).
    const filePath = writeReply(inbox, '', [attachment({ url: 'https://maestro.test/api/v1/attachments/att-1/download?x=FAIL' })]);
    await runPoller(bot);

    assert.equal(sends.length, 0);
    assert.equal(fs.existsSync(filePath), true); // transient pull failure — left for retry, never drop
  });

  it('text + a failed attachment still DELIVERS the text and deletes (item-8 has-text policy)', async () => {
    console.log = () => undefined;
    console.error = () => undefined;
    const { bot, sends, inbox } = setup();
    const filePath = writeReply(inbox, 'text survives', [attachment({ url: 'https://maestro.test/FAIL/download' })]);
    await runPoller(bot);

    assert.ok(sends.some((s) => /text survives/.test(s.text)), 'text delivered despite attachment pull failure');
    assert.ok(!sends.some((s) => (s.attachments?.length ?? 0) > 0), 'no attachment delivered');
    assert.equal(fs.existsSync(filePath), false); // text delivered -> file consumed
  });

  it('plain empty reply (no text, no attachments) is still deleted as nothing-to-post', async () => {
    console.log = () => undefined;
    const { bot, sends, inbox } = setup();
    const filePath = writeReply(inbox, '   '); // no attachments at all
    await runPoller(bot);

    assert.equal(sends.length, 0);
    assert.equal(fs.existsSync(filePath), false);
  });
});

/**
 * w3 FIX-LOOP — outbound consume-path hardening (Columbo P1 + Watson F2).
 * The agent-controlled `payload.attachments` is now policy-validated BEFORE any byte
 * is pulled; the pull is bounded + follows no redirects. These cover each hostile /
 * malformed descriptor shape and the bounded-read / redirect rejections.
 */
describe('w3 outbound attachment hardening (descriptor trust)', () => {
  /** Bot + capturing send + custom fetch; lets each test inject the pull response. */
  function harden(fetchImpl: typeof globalThis.fetch): {
    bot: OutboundBot;
    sends: SentRecord[];
    fetched: Array<{ url: string; redirect?: RequestRedirect }>;
    inbox: string;
  } {
    const inbox = path.join(tempDir(), 'maestro-inbox');
    const sends: SentRecord[] = [];
    const bot: OutboundBot = {
      slug: 'maestro',
      inboxDir: inbox,
      maestroUrl: 'https://maestro.test',
      send: async (_conversationId, text, _markdown, attachments) => {
        sends.push({ text, attachments });
      },
    };
    const fetched: Array<{ url: string; redirect?: RequestRedirect }> = [];
    globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
      fetched.push({ url: String(input), redirect: init?.redirect });
      return fetchImpl(input, init);
    }) as typeof fetch;
    return { bot, sends, fetched, inbox };
  }

  /** A response whose body streams `n` bytes and carries NO content-length header. */
  function streamingResponse(n: number): Response {
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new Uint8Array(n));
        controller.close();
      },
    });
    return new Response(stream, { status: 200 });
  }

  const okBytes = (async () => new Response(new Uint8Array([1, 2, 3, 4]), { status: 200 })) as typeof fetch;

  it('drops a descriptor whose DECLARED size exceeds the cap — never fetched', async () => {
    console.log = console.error = () => undefined;
    const { bot, sends, fetched, inbox } = harden(okBytes);
    const filePath = writeReply(inbox, '', [attachment({ size: 11 })]);
    await runPoller(bot, { maxBytes: 10, maxCount: 10, denyContentTypes: [] });

    assert.equal(fetched.length, 0); // rejected pre-pull
    assert.equal(sends.length, 0);
    assert.equal(fs.existsSync(filePath), false); // all-invalid → dropped
  });

  it('rejects an oversize STREAMED body with no Content-Length (bounded read)', async () => {
    console.log = console.error = () => undefined;
    const { bot, sends, fetched, inbox } = harden((async () => streamingResponse(50)) as typeof fetch);
    const filePath = writeReply(inbox, '', [attachment({ size: 4 })]); // descriptor lies small
    await runPoller(bot, { maxBytes: 10, maxCount: 10, denyContentTypes: [] });

    assert.equal(fetched.length, 1); // it DID fetch (descriptor looked valid) ...
    assert.ok(!sends.some((s) => (s.attachments?.length ?? 0) > 0)); // ... but body capped → no delivery
    assert.equal(fs.existsSync(filePath), false); // over-cap body is a DROP, not a retry
  });

  it('rejects an oversize Content-Length up front', async () => {
    console.log = console.error = () => undefined;
    const big = (async () =>
      new Response(new ReadableStream<Uint8Array>({ start: (c) => { c.enqueue(new Uint8Array(4)); c.close(); } }), {
        status: 200,
        headers: { 'content-length': '999999' },
      })) as typeof fetch;
    const { bot, sends, inbox } = harden(big);
    const filePath = writeReply(inbox, '', [attachment({ size: 4 })]);
    await runPoller(bot, { maxBytes: 10, maxCount: 10, denyContentTypes: [] });

    assert.ok(!sends.some((s) => (s.attachments?.length ?? 0) > 0));
    assert.equal(fs.existsSync(filePath), false);
  });

  it('drops a descriptor with a non-clean scan_status (e.g. suspicious)', async () => {
    console.log = console.error = () => undefined;
    const { bot, sends, fetched, inbox } = harden(okBytes);
    const filePath = writeReply(inbox, '', [attachment({ scan_status: 'suspicious' })]);
    await runPoller(bot);

    assert.equal(fetched.length, 0);
    assert.equal(sends.length, 0);
    assert.equal(fs.existsSync(filePath), false);
  });

  it('drops a same-origin url whose path is NOT /attachments/<id>/download', async () => {
    console.log = console.error = () => undefined;
    const { bot, sends, fetched, inbox } = harden(okBytes);
    // Same Maestro origin, but pointed at an internal route — must NOT be fetched.
    const filePath = writeReply(inbox, '', [attachment({ url: 'https://maestro.test/api/v1/agents' })]);
    await runPoller(bot);

    assert.equal(fetched.length, 0);
    assert.equal(sends.length, 0);
    assert.equal(fs.existsSync(filePath), false);
  });

  it('drops a same-origin /download url whose <id> does not match the descriptor id', async () => {
    console.log = console.error = () => undefined;
    const { bot, sends, fetched, inbox } = harden(okBytes);
    // id=att-1 but the url cites a DIFFERENT id — path/id must agree.
    const filePath = writeReply(inbox, '', [attachment({ url: 'https://maestro.test/api/v1/attachments/other/download' })]);
    await runPoller(bot);

    assert.equal(fetched.length, 0);
    assert.equal(sends.length, 0);
    assert.equal(fs.existsSync(filePath), false);
  });

  it('passes redirect:"error" to the pull and rejects a redirecting response', async () => {
    console.log = console.error = () => undefined;
    // Simulate undici rejecting a 3xx under redirect:'error'.
    const redirecting = (async () => {
      throw new TypeError('unexpected redirect');
    }) as typeof fetch;
    const { bot, sends, fetched, inbox } = harden(redirecting);
    const filePath = writeReply(inbox, '', [attachment()]);
    await runPoller(bot);

    assert.equal(fetched.length, 1);
    assert.equal(fetched[0].redirect, 'error'); // the fix: no redirects followed
    assert.ok(!sends.some((s) => (s.attachments?.length ?? 0) > 0)); // nothing delivered
    assert.equal(fs.existsSync(filePath), true); // transient-class pull failure → left for retry
  });

  it('enforces the count cap — pulls at most maxCount, drops the rest', async () => {
    console.log = console.error = () => undefined;
    const { bot, sends, inbox } = harden(okBytes);
    const filePath = writeReply(inbox, '', [
      attachment({ id: 'att-1', url: 'https://maestro.test/api/v1/attachments/att-1/download' }),
      attachment({ id: 'att-2', url: 'https://maestro.test/api/v1/attachments/att-2/download' }),
      attachment({ id: 'att-3', url: 'https://maestro.test/api/v1/attachments/att-3/download' }),
    ]);
    await runPoller(bot, { maxBytes: 26_214_400, maxCount: 2, denyContentTypes: [] });

    const withAttach = sends.filter((s) => (s.attachments?.length ?? 0) > 0);
    assert.equal(withAttach.length, 1);
    assert.equal(withAttach[0].attachments!.length, 2); // capped at 2, third dropped
    assert.equal(fs.existsSync(filePath), false);
  });

  it('happy path: a fully valid descriptor still delivers', async () => {
    console.log = console.error = () => undefined;
    const { bot, sends, fetched, inbox } = harden(okBytes);
    const filePath = writeReply(inbox, '', [attachment()]);
    await runPoller(bot);

    assert.equal(fetched.length, 1);
    assert.equal(fetched[0].redirect, 'error');
    const withAttach = sends.filter((s) => (s.attachments?.length ?? 0) > 0);
    assert.equal(withAttach.length, 1);
    assert.equal(withAttach[0].attachments![0].bytes.byteLength, 4);
    assert.equal(fs.existsSync(filePath), false);
  });

  // FIX-LOOP r2 (Whistler NEEDS-CHANGES): a hostile non-object element (null / string /
  // array) must be a total VALIDATION drop — validateOutboundDescriptor must NOT throw a
  // TypeError dereferencing `.kind` on it. A throw would escape the pull loop into the
  // outer catch, send NO text, and leave the file to spin on retry. These assert the
  // never-throws contract: text still delivers, hostile descriptors drop-not-retry, and
  // no fetch is ever attempted for a non-object descriptor.
  it('text + a null attachment element DELIVERS the text, deletes, and never fetches (total guard)', async () => {
    console.log = console.error = () => undefined;
    const { bot, sends, fetched, inbox } = harden(okBytes);
    const filePath = writeRawReply(inbox, 'text survives', [null]);
    await runPoller(bot);

    assert.ok(sends.some((s) => /text survives/.test(s.text)), 'text delivered despite null descriptor');
    assert.ok(!sends.some((s) => (s.attachments?.length ?? 0) > 0), 'no attachment delivered');
    assert.equal(fetched.length, 0, 'a non-object descriptor is never fetched');
    assert.equal(fs.existsSync(filePath), false, 'text delivered → file consumed');
  });

  it('attachment-only reply with a null element is DROPPED (deleted), never fetched', async () => {
    console.log = console.error = () => undefined;
    const { bot, sends, fetched, inbox } = harden(okBytes);
    const filePath = writeRawReply(inbox, '', [null]);
    await runPoller(bot);

    assert.equal(fetched.length, 0, 'never fetched a null descriptor');
    assert.equal(sends.length, 0, 'nothing delivered');
    assert.equal(fs.existsSync(filePath), false, 'all-invalid → DROP, not retry');
  });

  it('attachment-only reply with a non-object STRING element is DROPPED (deleted), never fetched', async () => {
    console.log = console.error = () => undefined;
    const { bot, sends, fetched, inbox } = harden(okBytes);
    const filePath = writeRawReply(inbox, '', ['x']);
    await runPoller(bot);

    assert.equal(fetched.length, 0, 'never fetched a string descriptor');
    assert.equal(sends.length, 0, 'nothing delivered');
    assert.equal(fs.existsSync(filePath), false, 'all-invalid → DROP, not retry');
  });

  it('attachment-only reply with a nested-ARRAY element is DROPPED (deleted), never fetched', async () => {
    console.log = console.error = () => undefined;
    const { bot, sends, fetched, inbox } = harden(okBytes);
    const filePath = writeRawReply(inbox, '', [[1, 2]]);
    await runPoller(bot);

    assert.equal(fetched.length, 0, 'never fetched an array descriptor');
    assert.equal(sends.length, 0, 'nothing delivered');
    assert.equal(fs.existsSync(filePath), false, 'all-invalid → DROP, not retry');
  });
});
