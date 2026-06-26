import assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { afterEach, describe, it } from 'node:test';

import { sendDeliveryFailureNack, type DeliveryFailure } from '../delivery-failure.js';
import { startOutboundPoller, type OutboundBot } from '../outbound.js';
import { createThreadStore, type ThreadEntry } from '../thread-store.js';
import type { AMPAttachmentV1, AMPMessage, AttachmentPolicy, ThreadContext } from '../types.js';

const TEST_POLICY: AttachmentPolicy = { maxBytes: 26_214_400, maxCount: 10, denyContentTypes: [] };
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
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-nack-out-'));
  tempRoots.push(dir);
  return dir;
}

function context(conversationId: string): ThreadContext {
  return {
    reference: {
      serviceUrl: 'https://smba.trafficmanager.net/amer/',
      channelId: 'msteams',
      conversation: { id: conversationId },
      bot: { id: 'bot-app-id', name: 'Bridge' },
      user: { id: 'user-1', name: 'User One' },
    } as ThreadContext['reference'],
    rootActivityId: 'root-activity-1',
    tenantId: 'tenant-1',
  };
}

function entry(overrides: Partial<ThreadEntry> = {}): ThreadEntry {
  return {
    botSlug: 'bridge',
    conversationId: 'conversation-bridge',
    ampMessageId: 'amp-inbound-1',
    aadObjectId: 'aad-1',
    context: context('conversation-bridge'),
    createdAt: 100,
    ...overrides,
  };
}

function message(overrides: Partial<AMPMessage> = {}): AMPMessage {
  return {
    envelope: {
      id: 'reply-1',
      from: 'sender-agent@example.aimaestro.local',
      to: 'teams-bridge-bot@example.aimaestro.local',
      subject: 'generic outbound reply',
      priority: 'normal',
      timestamp: '2026-06-26T00:00:00.000Z',
      signature: null,
      in_reply_to: 'amp-inbound-1',
      version: 'amp/0.1',
    },
    payload: { type: 'response', message: 'hello from a generic agent' },
    ...overrides,
  } as AMPMessage;
}

function attachment(overrides: Partial<AMPAttachmentV1> = {}): AMPAttachmentV1 {
  return {
    kind: 'amp-v1',
    id: 'att-1',
    filename: 'generic.txt',
    content_type: 'text/plain',
    size: 4,
    digest: 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855',
    url: 'https://maestro.test/api/v1/attachments/att-1/download',
    scan_status: 'basic_clean',
    uploaded_at: '2026-06-26T00:00:00.000Z',
    expires_at: '2026-06-27T00:00:00.000Z',
    ...overrides,
  };
}

function writeInboxMessage(inboxDir: string, msg: AMPMessage, file = 'reply.json'): string {
  const senderDir = path.join(inboxDir, 'sender-agent_example_aimaestro_local');
  fs.mkdirSync(senderDir, { recursive: true });
  const filePath = path.join(senderDir, file);
  fs.writeFileSync(filePath, JSON.stringify(msg), 'utf-8');
  return filePath;
}

function deadLetterFor(filePath: string, inboxDir: string): string {
  return path.join(inboxDir, 'dead-letter', path.relative(inboxDir, filePath));
}

async function waitFor(predicate: () => boolean, label: string): Promise<void> {
  const deadline = Date.now() + 2000;
  while (Date.now() < deadline) {
    if (predicate()) return;
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
  throw new Error(`timed out waiting for ${label}`);
}

interface PollerHarness {
  inbox: string;
  bot: OutboundBot;
  nacks: Array<{ toAgent: string; failure: DeliveryFailure; fromBotSlug: string }>;
  sendCalls: number;
}

function harness(send?: OutboundBot['send']): PollerHarness {
  const inbox = path.join(tempDir(), 'bridge-inbox');
  const nacks: PollerHarness['nacks'] = [];
  let sendCalls = 0;
  return {
    inbox,
    nacks,
    get sendCalls() {
      return sendCalls;
    },
    bot: {
      slug: 'bridge',
      inboxDir: inbox,
      maestroUrl: 'https://maestro.test',
      getAllowedOrigins: () => new Set(['https://maestro.test']),
      send: async (...args) => {
        sendCalls += 1;
        if (send) await send(...args);
      },
    },
  };
}

async function runHarness(
  h: PollerHarness,
  store = createThreadStore({ maxAgeMs: Infinity }),
  maxDeliveryAttempts = 1,
): Promise<() => void> {
  return startOutboundPoller({
    bots: [h.bot],
    threadStore: store,
    pollIntervalMs: 10,
    markdownDefault: true,
    policy: TEST_POLICY,
    debug: false,
    maxDeliveryAttempts,
    nack: async (toAgent, failure, fromBotSlug) => {
      h.nacks.push({ toAgent, failure, fromBotSlug });
    },
  });
}

describe('Teams outbound delivery failure NACK + dead-letter', () => {
  it('no in_reply_to emits one no_conversation NACK and dead-letters with sender nesting', async () => {
    console.log = () => undefined;
    const h = harness();
    const filePath = writeInboxMessage(h.inbox, message({
      envelope: { ...message().envelope, in_reply_to: undefined },
    } as Partial<AMPMessage>), 'no-in-reply-to.json');
    const stop = await runHarness(h);
    try {
      await waitFor(() => h.nacks.length === 1 && !fs.existsSync(filePath), 'terminal no_conversation');
      await new Promise((resolve) => setTimeout(resolve, 50));
    } finally {
      stop();
    }

    const deadLetter = deadLetterFor(filePath, h.inbox);
    assert.equal(fs.existsSync(deadLetter), true);
    assert.equal(h.nacks.length, 1);
    assert.equal(h.nacks[0].toAgent, 'sender-agent@example.aimaestro.local');
    assert.equal(h.nacks[0].fromBotSlug, 'bridge');
    assert.equal(h.nacks[0].failure.kind, 'teams-delivery-failure-v1');
    assert.equal(h.nacks[0].failure.reason, 'no_conversation');
    assert.equal(h.nacks[0].failure.retryable, false);
    assert.equal(h.nacks[0].failure.attempts, 1);
  });

  it('never NACKs a NACK; it silently dead-letters the failed notification', async () => {
    console.log = () => undefined;
    const h = harness();
    const filePath = writeInboxMessage(h.inbox, message({
      payload: {
        type: 'notification',
        message: 'delivery failed',
        context: {
          deliveryFailure: {
            kind: 'teams-delivery-failure-v1',
            originalMessageId: 'reply-original',
            botSlug: 'bridge',
            reason: 'no_conversation',
            detail: 'original failed',
            retryable: false,
            attempts: 1,
            attemptedAt: '2026-06-26T00:00:00.000Z',
          },
        },
      } as unknown as AMPMessage['payload'],
    }), 'nack.json');
    const stop = await runHarness(h);
    try {
      await waitFor(() => !fs.existsSync(filePath), 'nack dead-letter');
    } finally {
      stop();
    }

    assert.equal(h.nacks.length, 0);
    assert.equal(fs.existsSync(deadLetterFor(filePath, h.inbox)), true);
  });

  it('unmapped replies retry only up to the configured attempt bound, then mapping_expired NACK + dead-letter', async () => {
    console.log = () => undefined;
    const h = harness();
    const filePath = writeInboxMessage(h.inbox, message(), 'unmapped.json');
    const stop = await runHarness(h, createThreadStore({ maxAgeMs: Infinity }), 3);
    try {
      await waitFor(() => h.nacks.length === 1 && !fs.existsSync(filePath), 'bounded mapping retry');
    } finally {
      stop();
    }

    assert.equal(h.nacks[0].failure.reason, 'mapping_expired');
    assert.equal(h.nacks[0].failure.attempts, 3);
    assert.equal(fs.existsSync(deadLetterFor(filePath, h.inbox)), true);
  });

  it('attachment-only pull failures retry only up to the configured bound, then attachment_unavailable', async () => {
    console.log = console.error = () => undefined;
    globalThis.fetch = (async () => new Response('missing', { status: 404 })) as typeof fetch;
    const h = harness();
    const filePath = writeInboxMessage(h.inbox, message({
      payload: { type: 'response', message: '', attachments: [attachment()] },
    } as Partial<AMPMessage>), 'attachment-unavailable.json');
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry());
    const stop = await runHarness(h, store, 2);
    try {
      await waitFor(() => h.nacks.length === 1 && !fs.existsSync(filePath), 'bounded attachment retry');
    } finally {
      stop();
    }

    assert.equal(h.nacks[0].failure.reason, 'attachment_unavailable');
    assert.equal(h.nacks[0].failure.retryable, true);
    assert.equal(h.nacks[0].failure.attempts, 2);
  });

  it('attachment-only policy rejection emits attachment_rejected before dead-letter', async () => {
    console.log = console.error = () => undefined;
    const h = harness();
    const filePath = writeInboxMessage(h.inbox, message({
      payload: { type: 'response', message: '', attachments: [attachment({ url: 'https://blocked.test/internal' })] },
    } as Partial<AMPMessage>), 'attachment-rejected.json');
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry());
    const stop = await runHarness(h, store);
    try {
      await waitFor(() => h.nacks.length === 1 && !fs.existsSync(filePath), 'attachment rejected');
    } finally {
      stop();
    }

    assert.equal(h.nacks[0].failure.reason, 'attachment_rejected');
    assert.equal(h.nacks[0].failure.retryable, false);
    assert.equal(fs.existsSync(deadLetterFor(filePath, h.inbox)), true);
  });

  it('Teams 403 send failures are terminal bot_unreachable_or_forbidden', async () => {
    console.log = console.error = () => undefined;
    const err = new Error('forbidden') as Error & { status: number };
    err.status = 403;
    const h = harness(async () => {
      throw err;
    });
    const filePath = writeInboxMessage(h.inbox, message(), 'forbidden.json');
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry());
    const stop = await runHarness(h, store, 5);
    try {
      await waitFor(() => h.nacks.length === 1 && !fs.existsSync(filePath), 'forbidden terminal');
    } finally {
      stop();
    }

    assert.equal(h.sendCalls, 1);
    assert.equal(h.nacks[0].failure.reason, 'bot_unreachable_or_forbidden');
    assert.equal(h.nacks[0].failure.attempts, 1);
  });

  it('Teams transient send failures exhaust attempts before teams_api_error NACK', async () => {
    console.log = console.error = () => undefined;
    const err = new Error('upstream unavailable') as Error & { status: number };
    err.status = 503;
    const h = harness(async () => {
      throw err;
    });
    const filePath = writeInboxMessage(h.inbox, message(), 'transient-send.json');
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry());
    const stop = await runHarness(h, store, 3);
    try {
      await waitFor(() => h.nacks.length === 1 && !fs.existsSync(filePath), 'transient send exhausted');
    } finally {
      stop();
    }

    assert.equal(h.sendCalls, 3);
    assert.equal(h.nacks[0].failure.reason, 'teams_api_error');
    assert.equal(h.nacks[0].failure.retryable, true);
    assert.equal(h.nacks[0].failure.attempts, 3);
  });

  it('terminal failures still dead-letter when nack is absent or throws', async () => {
    console.log = console.error = () => undefined;
    const inbox = path.join(tempDir(), 'bridge-inbox');
    const bot: OutboundBot = {
      slug: 'bridge',
      inboxDir: inbox,
      maestroUrl: 'https://maestro.test',
      getAllowedOrigins: () => new Set(['https://maestro.test']),
      send: async () => undefined,
    };
    const filePath = writeInboxMessage(inbox, message({
      envelope: { ...message().envelope, in_reply_to: undefined },
    } as Partial<AMPMessage>), 'no-nack-dep.json');
    const stop = startOutboundPoller({
      bots: [bot],
      threadStore: createThreadStore({ maxAgeMs: Infinity }),
      pollIntervalMs: 10,
      markdownDefault: true,
      policy: TEST_POLICY,
      debug: false,
    });
    try {
      await waitFor(() => !fs.existsSync(filePath), 'dead-letter without nack dependency');
    } finally {
      stop();
    }
    assert.equal(fs.existsSync(deadLetterFor(filePath, inbox)), true);

    const h = harness();
    const throwingFile = writeInboxMessage(h.inbox, message({
      envelope: { ...message().envelope, in_reply_to: undefined, id: 'reply-throwing-nack' },
    } as Partial<AMPMessage>), 'throwing-nack.json');
    const stopThrowing = startOutboundPoller({
      bots: [h.bot],
      threadStore: createThreadStore({ maxAgeMs: Infinity }),
      pollIntervalMs: 10,
      markdownDefault: true,
      policy: TEST_POLICY,
      debug: false,
      nack: async () => {
        throw new Error('route down');
      },
    });
    try {
      await waitFor(() => !fs.existsSync(throwingFile), 'dead-letter after nack throw');
    } finally {
      stopThrowing();
    }
    assert.equal(fs.existsSync(deadLetterFor(throwingFile, h.inbox)), true);
  });

  it('route emitter posts the deliveryFailure v1 payload with bot apiKey auth', async () => {
    const failure: DeliveryFailure = {
      kind: 'teams-delivery-failure-v1',
      originalMessageId: 'reply-1',
      botSlug: 'bridge',
      reason: 'no_conversation',
      detail: 'message has no in_reply_to',
      retryable: false,
      attempts: 1,
      attemptedAt: '2026-06-26T00:00:00.000Z',
    };
    const calls: Array<{ url: string; auth?: string; body: unknown }> = [];
    globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
      calls.push({
        url: String(input),
        auth: (init?.headers as Record<string, string> | undefined)?.Authorization,
        body: JSON.parse(String(init?.body)),
      });
      return new Response(JSON.stringify({ id: 'nack-1', status: 'delivered' }), { status: 200 });
    }) as typeof fetch;

    await sendDeliveryFailureNack({
      maestroUrl: 'https://maestro.test',
      apiKey: 'bot-api-key',
      toAgent: 'sender-agent@example.aimaestro.local',
      failure,
      timeoutMs: 1000,
    });

    assert.equal(calls.length, 1);
    assert.equal(calls[0].url, 'https://maestro.test/api/v1/route');
    assert.equal(calls[0].auth, 'Bearer bot-api-key');
    assert.deepEqual(calls[0].body, {
      to: 'sender-agent@example.aimaestro.local',
      subject: 'Delivery failed: Teams (bridge)',
      priority: 'normal',
      in_reply_to: 'reply-1',
      payload: {
        type: 'notification',
        message: 'Your Teams message could not be delivered: message has no in_reply_to.',
        context: { deliveryFailure: failure },
      },
    });
  });
});
