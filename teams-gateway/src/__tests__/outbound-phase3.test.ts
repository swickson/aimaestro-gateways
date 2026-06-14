import assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { afterEach, describe, it } from 'node:test';

import { chunkText, formatReply, TEAMS_MAX_LENGTH } from '../format.js';
import { startOutboundPoller, type OutboundBot } from '../outbound.js';
import { restoreThreadStore, saveThreadStore } from '../thread-persistence.js';
import { createThreadStore, type ThreadEntry } from '../thread-store.js';
import type { AMPMessage, AttachmentPolicy, ThreadContext } from '../types.js';

/** Permissive policy — Phase-3 tests predate attachments; cap/validate is covered in attachments-outbound.test.ts. */
const TEST_POLICY: AttachmentPolicy = { maxBytes: 26_214_400, maxCount: 10, denyContentTypes: [] };

const tempRoots: string[] = [];

afterEach(() => {
  for (const dir of tempRoots.splice(0)) {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

function tempDir(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-outbound-test-'));
  tempRoots.push(dir);
  return dir;
}

function context(conversationId: string, serviceUrl = 'https://smba.trafficmanager.net/amer/'): ThreadContext {
  return {
    reference: {
      serviceUrl,
      channelId: 'msteams',
      conversation: { id: conversationId },
      bot: { id: 'bot-app-id', name: 'Maestro' },
      user: { id: 'user-1', name: 'Alice' },
    } as ThreadContext['reference'],
    rootActivityId: 'root-activity-1',
    tenantId: 'tenant-1',
  };
}

function entry(overrides: Partial<ThreadEntry> = {}): ThreadEntry {
  return {
    botSlug: 'maestro',
    conversationId: 'conversation-maestro',
    ampMessageId: 'amp-inbound-1',
    context: context('conversation-maestro'),
    createdAt: 100,
    ...overrides,
  };
}

function message(overrides: Partial<AMPMessage> = {}): AMPMessage {
  return {
    envelope: {
      id: 'reply-1',
      from: 'agent-one@example.aimaestro.local',
      to: 'teams-maestro-bot@example.aimaestro.local',
      timestamp: '2026-06-09T00:00:00.000Z',
      in_reply_to: 'amp-inbound-1',
    },
    payload: {
      message: 'hello from the agent',
      context: { botSlug: 'echo', conversationId: 'conversation-echo' },
    },
    ...overrides,
  } as AMPMessage;
}

function writeInboxMessage(inboxDir: string, msg: AMPMessage, file = 'reply.json'): string {
  const senderDir = path.join(inboxDir, 'agent-one_example_aimaestro_local');
  fs.mkdirSync(senderDir, { recursive: true });
  const filePath = path.join(senderDir, file);
  fs.writeFileSync(filePath, JSON.stringify(msg), 'utf-8');
  return filePath;
}

async function waitFor(predicate: () => boolean, label: string): Promise<void> {
  const deadline = Date.now() + 2000;
  while (Date.now() < deadline) {
    if (predicate()) return;
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
  throw new Error(`timed out waiting for ${label}`);
}

async function settlePollTick(): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, 25));
}

describe('Teams outbound reply poller', () => {
  it('routes by inbox botSlug, never by agent-echoed payload context, and deletes only after successful send', async () => {
    const root = tempDir();
    const maestroInbox = path.join(root, 'maestro-inbox');
    const echoInbox = path.join(root, 'echo-inbox');
    const filePath = writeInboxMessage(maestroInbox, message());
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry());
    store.record(entry({
      botSlug: 'echo',
      conversationId: 'conversation-echo',
      context: context('conversation-echo'),
    }));

    const sends: Array<{ slug: string; conversationId: string; text: string; markdown: boolean }> = [];
    const bots: OutboundBot[] = [
      {
        slug: 'maestro',
        inboxDir: maestroInbox,
        maestroUrl: 'https://maestro.test',
        send: async (conversationId, text, markdown) => {
          sends.push({ slug: 'maestro', conversationId, text, markdown });
        },
      },
      {
        slug: 'echo',
        inboxDir: echoInbox,
        maestroUrl: 'https://maestro.test',
        send: async (conversationId, text, markdown) => {
          sends.push({ slug: 'echo', conversationId, text, markdown });
        },
      },
    ];

    const stop = startOutboundPoller({ bots, threadStore: store, pollIntervalMs: 60_000, markdownDefault: true, policy: TEST_POLICY, debug: false });
    try {
      await waitFor(() => sends.length === 1 && !fs.existsSync(filePath), 'maestro send and delete');
      await settlePollTick();
    } finally {
      stop();
    }

    assert.deepEqual(sends.map(({ slug, conversationId, markdown }) => ({ slug, conversationId, markdown })), [
      { slug: 'maestro', conversationId: 'conversation-maestro', markdown: true },
    ]);
    assert.match(sends[0]?.text ?? '', /\*\*\[agent-one\]\*\* hello from the agent/);
    assert.equal(fs.existsSync(filePath), false);
  });

  it('leaves replies on disk when in_reply_to is missing, unmapped, or send fails', async () => {
    const root = tempDir();
    const inbox = path.join(root, 'inbox');
    const noReplyFile = writeInboxMessage(inbox, message({
      envelope: { ...message().envelope, in_reply_to: undefined },
    } as Partial<AMPMessage>), 'missing-reply.json');
    const unmappedFile = writeInboxMessage(inbox, message(), 'unmapped.json');
    const failingFile = writeInboxMessage(inbox, message({
      envelope: { ...message().envelope, in_reply_to: 'amp-inbound-2' },
    } as Partial<AMPMessage>), 'send-fails.json');
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry({ ampMessageId: 'amp-inbound-2' }));

    let sendCalls = 0;
    const stop = startOutboundPoller({
      bots: [{
        slug: 'maestro',
        inboxDir: inbox,
        maestroUrl: 'https://maestro.test',
        send: async () => {
          sendCalls += 1;
          throw new Error('transport down');
        },
      }],
      threadStore: store,
      pollIntervalMs: 60_000,
      markdownDefault: true,
      policy: TEST_POLICY,
      debug: false,
    });
    try {
      await waitFor(() => sendCalls === 1, 'failing send attempt');
      await settlePollTick();
    } finally {
      stop();
    }

    assert.equal(fs.existsSync(noReplyFile), true);
    assert.equal(fs.existsSync(unmappedFile), true);
    assert.equal(fs.existsSync(failingFile), true);
  });

  it('warns once for a stuck unmapped reply across repeated poll ticks', async () => {
    const root = tempDir();
    const inbox = path.join(root, 'inbox');
    const unmappedFile = writeInboxMessage(inbox, message(), 'unmapped.json');
    const store = createThreadStore({ maxAgeMs: Infinity });
    const logs: string[] = [];
    const originalLog = console.log;
    console.log = (...args: unknown[]) => {
      logs.push(args.map(String).join(' '));
    };

    const stop = startOutboundPoller({
      bots: [{
        slug: 'maestro',
        inboxDir: inbox,
        maestroUrl: 'https://maestro.test',
        send: async () => {
          throw new Error('should not send without a thread mapping');
        },
      }],
      threadStore: store,
      pollIntervalMs: 10,
      markdownDefault: true,
      policy: TEST_POLICY,
      debug: false,
    });
    try {
      await waitFor(() => logs.some((line) => line.includes('no thread mapping')), 'first unmapped warning');
      await new Promise((resolve) => setTimeout(resolve, 50));
    } finally {
      stop();
      console.log = originalLog;
    }

    assert.equal(fs.existsSync(unmappedFile), true);
    assert.equal(logs.filter((line) => line.includes('no thread mapping')).length, 1);
  });

  it('deletes whitespace-only replies without sending and sends one call per formatted chunk', async () => {
    const root = tempDir();
    const inbox = path.join(root, 'inbox');
    const blankFile = writeInboxMessage(inbox, message({
      envelope: { ...message().envelope, in_reply_to: 'amp-blank' },
      payload: { message: '   \n\t   ' },
    } as Partial<AMPMessage>), 'blank.json');
    writeInboxMessage(inbox, message({
      envelope: { ...message().envelope, in_reply_to: 'amp-long' },
      payload: { message: 'b'.repeat(TEAMS_MAX_LENGTH * 2 + 10) },
    } as Partial<AMPMessage>), 'long.json');
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry({ ampMessageId: 'amp-blank' }));
    store.record(entry({ ampMessageId: 'amp-long' }));

    const chunks: string[] = [];
    const stop = startOutboundPoller({
      bots: [{
        slug: 'maestro',
        inboxDir: inbox,
        maestroUrl: 'https://maestro.test',
        send: async (_conversationId, text) => {
          chunks.push(text);
        },
      }],
      threadStore: store,
      pollIntervalMs: 60_000,
      markdownDefault: false,
      policy: TEST_POLICY,
      debug: false,
    });
    try {
      await waitFor(() => chunks.length > 1, 'chunked send');
      await settlePollTick();
    } finally {
      stop();
    }

    assert.equal(fs.existsSync(blankFile), false);
    assert.ok(chunks.length > 1);
    assert.ok(chunks.every((chunk) => chunk.length <= TEAMS_MAX_LENGTH));
  });

  it('does not start overlapping sends when a poll tick is still in flight', async () => {
    const root = tempDir();
    const inbox = path.join(root, 'inbox');
    writeInboxMessage(inbox, message());
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry());

    let releaseSend!: () => void;
    const sendStarted = new Promise<void>((resolve) => {
      releaseSend = resolve;
    });
    let sendCalls = 0;
    const stop = startOutboundPoller({
      bots: [{
        slug: 'maestro',
        inboxDir: inbox,
        maestroUrl: 'https://maestro.test',
        send: async () => {
          sendCalls += 1;
          await sendStarted;
        },
      }],
      threadStore: store,
      pollIntervalMs: 1,
      markdownDefault: true,
      policy: TEST_POLICY,
      debug: false,
    });

    await waitFor(() => sendCalls === 1, 'first send start');
    await new Promise((resolve) => setTimeout(resolve, 40));
    assert.equal(sendCalls, 1);

    releaseSend();
    try {
      await waitFor(() => sendCalls >= 1, 'send release');
      await settlePollTick();
    } finally {
      stop();
    }
  });
});

describe('Teams thread-store persistence', () => {
  it('round-trips snapshots through atomic save and restore', () => {
    const root = tempDir();
    const snapshotPath = path.join(root, 'nested', 'threads.json');
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry({ recordedAt: 1234 }));

    assert.equal(saveThreadStore(store, snapshotPath), true);
    assert.equal(fs.existsSync(snapshotPath), true);
    assert.deepEqual(fs.readdirSync(path.dirname(snapshotPath)).filter((name) => name.includes('.tmp-')), []);

    const restored = createThreadStore({ maxAgeMs: Infinity });
    assert.equal(restoreThreadStore(restored, snapshotPath), 1);
    assert.deepEqual(restored.snapshot(), store.snapshot());
  });

  it('restores missing, corrupt, and wrong-shape snapshots as empty without throwing', () => {
    const root = tempDir();
    const store = createThreadStore();

    assert.doesNotThrow(() => assert.equal(restoreThreadStore(store, path.join(root, 'missing.json')), 0));

    const corrupt = path.join(root, 'corrupt.json');
    fs.writeFileSync(corrupt, '{', 'utf-8');
    assert.doesNotThrow(() => assert.equal(restoreThreadStore(store, corrupt), 0));

    const wrongShape = path.join(root, 'wrong-shape.json');
    fs.writeFileSync(wrongShape, JSON.stringify({ version: 1, entries: {} }), 'utf-8');
    assert.doesNotThrow(() => assert.equal(restoreThreadStore(store, wrongShape), 0));
    assert.equal(store.size(), 0);
  });

  it('treats malformed snapshot entries as wrong-shape and restores empty', () => {
    const root = tempDir();
    const validEntry = entry();
    const cases: Array<{ name: string; badEntry: unknown }> = [
      { name: 'missing composite keys', badEntry: { botSlug: 'maestro' } },
      { name: 'non-finite createdAt', badEntry: { ...validEntry, createdAt: Number.NaN } },
      { name: 'non-finite recordedAt', badEntry: { ...validEntry, recordedAt: Infinity } },
      { name: 'missing context object', badEntry: { ...validEntry, context: null } },
      { name: 'wrong rootActivityId type', badEntry: { ...validEntry, context: { ...validEntry.context, rootActivityId: 42 } } },
      { name: 'wrong tenantId type', badEntry: { ...validEntry, context: { ...validEntry.context, tenantId: false } } },
      {
        name: 'missing nested conversation id',
        badEntry: {
          ...validEntry,
          context: {
            ...validEntry.context,
            reference: {
              ...validEntry.context.reference,
              conversation: {},
            },
          },
        },
      },
    ];

    for (const { name, badEntry } of cases) {
      const malformedEntries = path.join(root, `${name.replaceAll(' ', '-')}.json`);
      fs.writeFileSync(malformedEntries, JSON.stringify({
        version: 1,
        entries: [validEntry, badEntry],
      }), 'utf-8');

      const store = createThreadStore();
      const restored = restoreThreadStore(store, malformedEntries);
      assert.equal(restored, 0, name);
      assert.equal(store.size(), 0, name);
    }
  });

  it('accepts a pre-Phase-5 entry missing aadObjectId (graceful schema skew)', () => {
    const root = tempDir();
    const file = path.join(root, 'legacy.json');
    const legacy = entry();
    delete (legacy as { aadObjectId?: string }).aadObjectId;
    fs.writeFileSync(file, JSON.stringify({ version: 1, entries: [legacy] }), 'utf-8');

    const store = createThreadStore({ maxAgeMs: Infinity });
    // Not dropped — the snapshot restores, but the entry is NOT DM-indexed.
    assert.equal(restoreThreadStore(store, file), 1);
    assert.equal(store.findByAmpMessageId('maestro', 'amp-inbound-1')?.ampMessageId, 'amp-inbound-1');
  });

  it('rejects an entry whose aadObjectId is present but not a string', () => {
    const root = tempDir();
    const file = path.join(root, 'bad-aad.json');
    const bad = { ...entry(), aadObjectId: 42 };
    fs.writeFileSync(file, JSON.stringify({ version: 1, entries: [bad] }), 'utf-8');

    const store = createThreadStore({ maxAgeMs: Infinity });
    assert.equal(restoreThreadStore(store, file), 0);
    assert.equal(store.size(), 0);
  });
});

describe('Teams thread-store expiry and recency', () => {
  it('expires by store-owned recordedAt, not caller-created createdAt, and preserves recordedAt through restore', () => {
    let now = 1_000;
    const store = createThreadStore({ maxAgeMs: 100, now: () => now });
    store.record(entry({ ampMessageId: 'future-created', createdAt: 1_000_000 }));

    now = 1_050;
    assert.equal(store.findByAmpMessageId('maestro', 'future-created')?.recordedAt, 1_000);

    const restored = createThreadStore({ maxAgeMs: 100, now: () => now });
    restored.restore(store.snapshot());
    assert.equal(restored.findByAmpMessageId('maestro', 'future-created')?.recordedAt, 1_000);

    now = 1_101;
    assert.equal(restored.findByAmpMessageId('maestro', 'future-created'), null);
    assert.deepEqual(restored.snapshot().entries, []);
  });

  it('keeps the most recent conversation pointer while count eviction drops oldest entries', () => {
    const store = createThreadStore({ maxEntries: 2, maxAgeMs: Infinity, now: () => 5000 });
    store.record(entry({ ampMessageId: 'old', conversationId: 'conversation-1', context: context('conversation-1'), createdAt: 100 }));
    store.record(entry({ ampMessageId: 'new', conversationId: 'conversation-1', context: context('conversation-1'), createdAt: 200 }));
    store.record(entry({ ampMessageId: 'other', conversationId: 'conversation-2', context: context('conversation-2'), createdAt: 300 }));

    assert.equal(store.findByAmpMessageId('maestro', 'old'), null);
    assert.equal(store.findRecentByConversation('maestro', 'conversation-1')?.ampMessageId, 'new');
    assert.equal(store.findRecentByConversation('maestro', 'conversation-2')?.ampMessageId, 'other');
    assert.equal(store.size(), 2);
  });
});

describe('Teams reply formatting', () => {
  it('formats short markdown and plain-text replies with the expected sender prefix', () => {
    assert.deepEqual(formatReply({ displayName: 'agent-one', message: 'hello', markdown: true, maxLength: 100 }), {
      chunks: ['**[agent-one]** hello'],
      markdown: true,
    });
    assert.deepEqual(formatReply({ displayName: 'agent-one', message: 'hello', markdown: false, maxLength: 100 }), {
      chunks: ['[agent-one] hello'],
      markdown: false,
    });
  });

  it('skips empty replies and falls back to Agent when the display label is blank', () => {
    assert.deepEqual(formatReply({ displayName: 'agent-one', message: '   ', markdown: true }), {
      chunks: [],
      markdown: true,
    });
    assert.deepEqual(formatReply({ displayName: '   ', message: 'hello', markdown: false, maxLength: 100 }), {
      chunks: ['[Agent] hello'],
      markdown: false,
    });
  });

  it('splits long text on newline before space before hard cut', () => {
    assert.equal(chunkText(`aaaaaa\nbb ${'c'.repeat(20)}`, 10)[0], 'aaaaaa');
    assert.equal(chunkText(`aaaa bbbb ${'c'.repeat(20)}`, 10)[0], 'aaaa bbbb');
    assert.deepEqual(chunkText('x'.repeat(25), 10), ['x'.repeat(10), 'x'.repeat(10), 'x'.repeat(5)]);
  });
});
