import assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { afterEach, describe, it } from 'node:test';

import { buildCard } from '../card-builder.js';
import { startOutboundPoller, type OutboundBot } from '../outbound.js';
import { createThreadStore, type ThreadEntry } from '../thread-store.js';
import type { AMPMessage, AttachmentPolicy, ThreadContext } from '../types.js';

const TEST_POLICY: AttachmentPolicy = { maxBytes: 26_214_400, maxCount: 10, denyContentTypes: [] };
const tempRoots: string[] = [];

afterEach(() => {
  for (const dir of tempRoots.splice(0)) fs.rmSync(dir, { recursive: true, force: true });
});

function tempDir(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'w5-cards-whistler-'));
  tempRoots.push(dir);
  return dir;
}

function context(): ThreadContext {
  return {
    reference: {
      serviceUrl: 'https://smba.trafficmanager.net/amer/',
      channelId: 'msteams',
      conversation: { id: 'conversation-maestro' },
      bot: { id: 'bot-app-id', name: 'Maestro' },
      user: { id: 'user-1', name: 'Alice' },
    } as ThreadContext['reference'],
    rootActivityId: 'root-activity-1',
    tenantId: 'tenant-1',
  };
}

function entry(): ThreadEntry {
  return {
    botSlug: 'maestro',
    conversationId: 'conversation-maestro',
    ampMessageId: 'amp-inbound-1',
    context: context(),
    createdAt: 100,
  };
}

function message(payload: Partial<AMPMessage['payload']>): AMPMessage {
  return {
    envelope: {
      id: 'reply-1',
      from: 'agent-one@example.aimaestro.local',
      to: 'teams-maestro-bot@example.aimaestro.local',
      timestamp: '2026-06-15T00:00:00.000Z',
      in_reply_to: 'amp-inbound-1',
      version: '1.0',
      priority: 'normal',
      signature: null,
      subject: 'test',
    },
    payload: { type: 'text', message: 'plain reply', context: null, ...payload },
  } as AMPMessage;
}

function writeInboxMessage(inboxDir: string, msg: AMPMessage): string {
  const senderDir = path.join(inboxDir, 'agent-one_example_aimaestro_local');
  fs.mkdirSync(senderDir, { recursive: true });
  const filePath = path.join(senderDir, 'reply.json');
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

async function runPoller(msg: AMPMessage): Promise<{
  sends: Array<{ text: string; card?: Record<string, unknown> }>;
  selectors: string[];
}> {
  const root = tempDir();
  const inbox = path.join(root, 'inbox');
  const filePath = writeInboxMessage(inbox, msg);
  const store = createThreadStore({ maxAgeMs: Infinity });
  store.record(entry());

  const sends: Array<{ text: string; card?: Record<string, unknown> }> = [];
  const selectors: string[] = [];
  const bots: OutboundBot[] = [
    {
      slug: 'maestro',
      inboxDir: inbox,
      maestroUrl: 'https://maestro.test',
      getAllowedOrigins: () => new Set(['https://maestro.test']),
      send: async (_conversationId, text, _markdown, _attachments, card) => {
        sends.push({ text, card });
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
    buildCard: (render, text) => {
      selectors.push(render);
      return buildCard(render, text);
    },
  });

  try {
    await waitFor(() => !fs.existsSync(filePath), 'message processed');
    await new Promise((resolve) => setTimeout(resolve, 30));
  } finally {
    stop();
  }

  return { sends, selectors };
}

describe('w5 cards render selector reachability', () => {
  const cardJson = JSON.stringify({ title: 'Deploy', status: 'success', description: 'green' });

  it('keeps top-level payload.render authoritative over context.render', async () => {
    const { sends, selectors } = await runPoller(message({
      render: 'text',
      message: cardJson,
      context: { render: 'status_summary' } as unknown as AMPMessage['payload']['context'],
    }));

    assert.deepEqual(selectors, ['text']);
    assert.equal(sends.length, 1);
    assert.equal(sends[0].card, undefined);
    assert.match(sends[0].text, /"status":"success"/);
  });

  it('uses string payload.context.render as a fallback when payload.render is absent', async () => {
    const { sends, selectors } = await runPoller(message({
      message: cardJson,
      context: { render: 'status_summary' } as unknown as AMPMessage['payload']['context'],
    }));

    assert.deepEqual(selectors, ['status_summary']);
    assert.equal(sends.length, 1);
    assert.equal(sends[0].card?.type, 'AdaptiveCard');
    assert.equal(sends[0].text, '');
  });

  it('ignores non-string payload.context.render and falls through to text', async () => {
    const { sends, selectors } = await runPoller(message({
      message: cardJson,
      context: { render: ['status_summary'] } as unknown as AMPMessage['payload']['context'],
    }));

    assert.deepEqual(selectors, []);
    assert.equal(sends.length, 1);
    assert.equal(sends[0].card, undefined);
    assert.match(sends[0].text, /"status":"success"/);
  });

  it('keeps the no-render text fallback unchanged', async () => {
    const { sends, selectors } = await runPoller(message({ message: 'ordinary reply' }));

    assert.deepEqual(selectors, []);
    assert.equal(sends.length, 1);
    assert.equal(sends[0].card, undefined);
    assert.match(sends[0].text, /ordinary reply/);
  });

  it('keeps the existing top-level card path working', async () => {
    const { sends, selectors } = await runPoller(message({
      render: 'status_summary',
      message: cardJson,
    }));

    assert.deepEqual(selectors, ['status_summary']);
    assert.equal(sends.length, 1);
    assert.equal(sends[0].card?.type, 'AdaptiveCard');
    assert.equal(sends[0].text, '');
  });
});
