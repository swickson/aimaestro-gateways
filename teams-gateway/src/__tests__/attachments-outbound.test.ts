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
import type { AMPAttachmentV1, AMPMessage, ThreadContext } from '../types.js';

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

async function runPoller(bot: OutboundBot): Promise<void> {
  const store = createThreadStore({ maxAgeMs: Infinity });
  store.record(entry());
  const stop = startOutboundPoller({ bots: [bot], threadStore: store, pollIntervalMs: 60_000, markdownDefault: true, debug: false });
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

  it('does not fetch attachment URLs outside the configured Maestro origin', async () => {
    console.log = () => undefined;
    console.error = () => undefined;
    const { bot, sends, fetched, inbox } = setup();
    const filePath = writeReply(inbox, '', [attachment({ url: 'https://evil.test/internal' })]);
    await runPoller(bot);

    assert.equal(fetched.length, 0);
    assert.equal(sends.length, 0);
    assert.equal(fs.existsSync(filePath), true);
  });

  it('attachment-only whose pull FAILS is left for retry (not deleted, no send)', async () => {
    console.log = () => undefined;
    console.error = () => undefined;
    const { bot, sends, inbox } = setup();
    const filePath = writeReply(inbox, '', [attachment({ url: 'https://maestro.test/FAIL/download' })]);
    await runPoller(bot);

    assert.equal(sends.length, 0);
    assert.equal(fs.existsSync(filePath), true); // left for retry — never drop agent data
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
