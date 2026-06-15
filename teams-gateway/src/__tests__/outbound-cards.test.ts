import assert from 'node:assert/strict';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import { afterEach, describe, it } from 'node:test';

import { buildCard } from '../card-builder.js';
import { formatStatusSummaryFallback } from '../format.js';
import { startOutboundPoller, type OutboundBot } from '../outbound.js';
import { createThreadStore, type ThreadEntry } from '../thread-store.js';
import type { AMPMessage, AttachmentPolicy, ThreadContext } from '../types.js';

const TEST_POLICY: AttachmentPolicy = { maxBytes: 26_214_400, maxCount: 10, denyContentTypes: [] };

const tempRoots: string[] = [];

afterEach(() => {
  for (const dir of tempRoots.splice(0)) {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

function tempDir(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-cards-test-'));
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
    rootActivityId: 'root-activity-1',
    tenantId: 'tenant-1',
  };
}

function entry(): ThreadEntry {
  return {
    botSlug: 'maestro',
    conversationId: 'conversation-maestro',
    ampMessageId: 'amp-inbound-1',
    context: context('conversation-maestro'),
    createdAt: 100,
  };
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

describe('Teams Adaptive Cards Outbound', () => {
  describe('Card Builder & Fallback formatting', () => {
    it('buildCard returns null for normal message type', () => {
      assert.equal(buildCard('text', 'hello'), null);
    });

    it('buildCard returns null for invalid JSON status_summary', () => {
      assert.equal(buildCard('status_summary', 'not-a-json'), null);
    });

    it('buildCard returns null for missing title or status', () => {
      assert.equal(buildCard('status_summary', JSON.stringify({ title: 'test' })), null);
      assert.equal(buildCard('status_summary', JSON.stringify({ status: 'success' })), null);
      assert.equal(buildCard('status_summary', JSON.stringify({ title: 'test', status: 'invalid-status' })), null);
    });

    it('buildCard returns valid AdaptiveCard for valid status_summary', () => {
      const payload = {
        title: 'System Diagnostics',
        status: 'warning',
        description: 'Completed with warnings',
        facts: [{ title: 'CPU', value: '98%' }]
      };
      const card = buildCard('status_summary', JSON.stringify(payload));
      assert.ok(card);
      assert.equal(card.type, 'AdaptiveCard');
      assert.equal(card.version, '1.5');
      assert.ok(Array.isArray(card.body));

      const body = card.body as Array<Record<string, unknown>>;
      assert.equal(body[0].text, 'System Diagnostics');
      assert.equal(body[1].text, 'Status: WARNING');
      assert.equal(body[1].color, 'Warning');
      assert.equal(body[2].text, 'Completed with warnings');
      assert.equal(body[3].type, 'FactSet');

      const factSet = body[3] as { facts: Array<{ title: string; value: string }> };
      assert.equal(factSet.facts[0].title, 'CPU');
      assert.equal(factSet.facts[0].value, '98%');
    });

    it('formatStatusSummaryFallback formats expected markdown text', () => {
      const payload = {
        title: 'System Diagnostics',
        status: 'warning' as const,
        description: 'Completed with warnings',
        facts: [{ title: 'CPU', value: '98%' }]
      };
      const markdown = formatStatusSummaryFallback(payload);
      assert.match(markdown, /\*\*\[System Diagnostics\]\*\*/);
      assert.match(markdown, /Status: \*\*WARNING\*\*/);
      assert.match(markdown, /Completed with warnings/);
      assert.match(markdown, /- \*\*CPU\*\*: 98%/);
    });
  });

  describe('Outbound poller integration', () => {
    it('sends Adaptive Card when type is status_summary and card delivery succeeds', async () => {
      const root = tempDir();
      const inbox = path.join(root, 'inbox');
      const payloadJSON = JSON.stringify({
        title: 'Summary',
        status: 'success',
        description: 'Everything green'
      });
      const msg: AMPMessage = {
        envelope: {
          id: 'reply-card',
          from: 'agent-one@example.aimaestro.local',
          to: 'teams-maestro-bot@example.aimaestro.local',
          timestamp: '2026-06-09T00:00:00.000Z',
          in_reply_to: 'amp-inbound-1',
          version: '1.0',
          priority: 'normal',
          signature: null,
          subject: 'test',
        },
        payload: {
          type: 'text',
          render: 'status_summary',
          message: payloadJSON,
          context: null,
        }
      };
      const filePath = writeInboxMessage(inbox, msg);
      const store = createThreadStore({ maxAgeMs: Infinity });
      store.record(entry());

      const sends: Array<{
        conversationId: string;
        text: string;
        markdown: boolean;
        card?: Record<string, unknown>;
      }> = [];

      const bots: OutboundBot[] = [
        {
          slug: 'maestro',
          inboxDir: inbox,
          maestroUrl: 'https://maestro.test',
          allowedOrigins: new Set(['https://maestro.test']),
          send: async (conversationId, text, markdown, attachments, card) => {
            sends.push({ conversationId, text, markdown, card });
          },
        }
      ];

      const stop = startOutboundPoller({
        bots,
        threadStore: store,
        pollIntervalMs: 60_000,
        markdownDefault: true,
        policy: TEST_POLICY,
        debug: false,
        buildCard,
      });

      try {
        await waitFor(() => sends.length === 1 && !fs.existsSync(filePath), 'card sent');
        await settlePollTick();
      } finally {
        stop();
      }

      assert.equal(sends.length, 1);
      assert.ok(sends[0].card);
      assert.equal(sends[0].card.type, 'AdaptiveCard');
      assert.equal(sends[0].text, '');
    });

    it('falls back to markdown text when card delivery throws an error (send-with-fallback)', async () => {
      const root = tempDir();
      const inbox = path.join(root, 'inbox');
      const payloadJSON = JSON.stringify({
        title: 'Summary',
        status: 'error',
        description: 'Failed system check',
        facts: [{ title: 'Error Code', value: '500' }]
      });
      const msg: AMPMessage = {
        envelope: {
          id: 'reply-card-fallback',
          from: 'agent-one@example.aimaestro.local',
          to: 'teams-maestro-bot@example.aimaestro.local',
          timestamp: '2026-06-09T00:00:00.000Z',
          in_reply_to: 'amp-inbound-1',
          version: '1.0',
          priority: 'normal',
          signature: null,
          subject: 'test',
        },
        payload: {
          type: 'text',
          render: 'status_summary',
          message: payloadJSON,
          context: null,
        }
      };
      const filePath = writeInboxMessage(inbox, msg);
      const store = createThreadStore({ maxAgeMs: Infinity });
      store.record(entry());

      const sends: Array<{
        conversationId: string;
        text: string;
        markdown: boolean;
        card?: Record<string, unknown>;
      }> = [];

      let firstCall = true;
      const bots: OutboundBot[] = [
        {
          slug: 'maestro',
          inboxDir: inbox,
          maestroUrl: 'https://maestro.test',
          allowedOrigins: new Set(['https://maestro.test']),
          send: async (conversationId, text, markdown, attachments, card) => {
            if (card && firstCall) {
              firstCall = false;
              throw new Error('Teams rejected card JSON');
            }
            sends.push({ conversationId, text, markdown, card });
          },
        }
      ];

      const stop = startOutboundPoller({
        bots,
        threadStore: store,
        pollIntervalMs: 60_000,
        markdownDefault: true,
        policy: TEST_POLICY,
        debug: false,
        buildCard,
      });

      try {
        await waitFor(() => sends.length === 1 && !fs.existsSync(filePath), 'text fallback sent');
        await settlePollTick();
      } finally {
        stop();
      }

      assert.equal(sends.length, 1);
      assert.equal(sends[0].card, undefined);
      assert.match(sends[0].text, /\*\*\[agent-one\]\*\* \*\*\[Summary\]\*\*/);
      assert.match(sends[0].text, /Status: \*\*ERROR\*\*/);
      assert.match(sends[0].text, /Failed system check/);
      assert.match(sends[0].text, /- \*\*Error Code\*\*: 500/);
    });

    it('falls back to raw text if payload message JSON is malformed', async () => {
      const root = tempDir();
      const inbox = path.join(root, 'inbox');
      const msg: AMPMessage = {
        envelope: {
          id: 'reply-malformed',
          from: 'agent-one@example.aimaestro.local',
          to: 'teams-maestro-bot@example.aimaestro.local',
          timestamp: '2026-06-09T00:00:00.000Z',
          in_reply_to: 'amp-inbound-1',
          version: '1.0',
          priority: 'normal',
          signature: null,
          subject: 'test',
        },
        payload: {
          type: 'text',
          render: 'status_summary',
          message: 'not-valid-json',
          context: null,
        }
      };
      const filePath = writeInboxMessage(inbox, msg);
      const store = createThreadStore({ maxAgeMs: Infinity });
      store.record(entry());

      const sends: Array<{
        conversationId: string;
        text: string;
        markdown: boolean;
        card?: Record<string, unknown>;
      }> = [];

      const bots: OutboundBot[] = [
        {
          slug: 'maestro',
          inboxDir: inbox,
          maestroUrl: 'https://maestro.test',
          allowedOrigins: new Set(['https://maestro.test']),
          send: async (conversationId, text, markdown, attachments, card) => {
            sends.push({ conversationId, text, markdown, card });
          },
        }
      ];

      const stop = startOutboundPoller({
        bots,
        threadStore: store,
        pollIntervalMs: 60_000,
        markdownDefault: true,
        policy: TEST_POLICY,
        debug: false,
        buildCard,
      });

      try {
        await waitFor(() => sends.length === 1 && !fs.existsSync(filePath), 'raw text sent');
        await settlePollTick();
      } finally {
        stop();
      }

      assert.equal(sends.length, 1);
      assert.equal(sends[0].card, undefined);
      assert.match(sends[0].text, /\*\*\[agent-one\]\*\* not-valid-json/);
    });
  });
});
