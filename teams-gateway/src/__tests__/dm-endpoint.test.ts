import assert from 'node:assert/strict';
import type { AddressInfo } from 'node:net';
import type { Server } from 'node:http';
import express from 'express';
import { afterEach, describe, it } from 'node:test';

import { deliverDm, createDmRouter, type DmDeps } from '../dm.js';
import { createThreadStore, type ThreadEntry, type ThreadStore } from '../thread-store.js';
import type { ThreadContext } from '../types.js';

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

function entry(overrides: Partial<ThreadEntry> = {}): ThreadEntry {
  const conversationId = overrides.conversationId ?? 'conv-maestro';
  return {
    botSlug: 'maestro',
    conversationId,
    ampMessageId: 'amp-1',
    aadObjectId: 'aad-user-1',
    context: context(conversationId),
    createdAt: 100,
    ...overrides,
  };
}

interface Sent {
  botSlug: string;
  conversationId: string;
  text: string;
  markdown: boolean;
}

function deps(store: ThreadStore, sent: Sent[], markdownDefault = true): DmDeps {
  return {
    threadStore: store,
    knownBots: new Set(['maestro', 'echo']),
    markdownDefault,
    sendChunk: async (botSlug, conversationId, text, markdown) => {
      sent.push({ botSlug, conversationId, text, markdown });
    },
  };
}

describe('deliverDm (proactive DM core)', () => {
  it('400s on a missing platformUserId', async () => {
    const sent: Sent[] = [];
    const r = await deliverDm(deps(createThreadStore(), sent), { message: 'hi' });
    assert.equal(r.status, 400);
    assert.equal(sent.length, 0);
  });

  it('400s on a missing/blank message', async () => {
    const sent: Sent[] = [];
    const r = await deliverDm(deps(createThreadStore(), sent), { platformUserId: 'aad-user-1', message: '   ' });
    assert.equal(r.status, 400);
    assert.equal(sent.length, 0);
  });

  it('400s on a botSlug not in the registry', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry());
    const sent: Sent[] = [];
    const r = await deliverDm(deps(store, sent), { platformUserId: 'aad-user-1', botSlug: 'ghost', message: 'hi' });
    assert.equal(r.status, 400);
    assert.equal(r.json.error, 'bad_request');
    assert.equal(sent.length, 0);
  });

  it('409s undeliverable when the user has no prior contact (cold start = v2)', async () => {
    const sent: Sent[] = [];
    const r = await deliverDm(deps(createThreadStore(), sent), { platformUserId: 'aad-user-1', message: 'hi' });
    assert.equal(r.status, 409);
    assert.equal(r.json.reason, 'no_prior_contact');
    assert.equal(sent.length, 0);
  });

  it('delivers under a caller-pinned botSlug (happy path)', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry({ botSlug: 'echo', conversationId: 'conv-echo', ampMessageId: 'amp-e' }));
    const sent: Sent[] = [];
    const r = await deliverDm(deps(store, sent), { platformUserId: 'aad-user-1', botSlug: 'echo', message: 'ping' });
    assert.equal(r.status, 200);
    assert.deepEqual(r.json, { delivered: true, botSlug: 'echo', chunks: 1 });
    assert.equal(sent.length, 1);
    assert.equal(sent[0]?.botSlug, 'echo');
    assert.equal(sent[0]?.conversationId, 'conv-echo');
    assert.equal(sent[0]?.text, 'ping');
  });

  it('falls back to the most-recent bot when no botSlug is supplied', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry({ botSlug: 'maestro', conversationId: 'conv-m', ampMessageId: 'amp-m', createdAt: 100 }));
    store.record(entry({ botSlug: 'echo', conversationId: 'conv-e', ampMessageId: 'amp-e', createdAt: 200 }));
    const sent: Sent[] = [];
    const r = await deliverDm(deps(store, sent), { platformUserId: 'aad-user-1', message: 'ping' });
    assert.equal(r.status, 200);
    assert.equal(r.json.botSlug, 'echo', 'last-seen bot wins');
    assert.equal(sent[0]?.conversationId, 'conv-e');
  });

  it('prepends a bold subject line when subject is present', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry());
    const sent: Sent[] = [];
    const r = await deliverDm(deps(store, sent), { platformUserId: 'aad-user-1', message: 'body', subject: 'Heads up' });
    assert.equal(r.status, 200);
    assert.equal(sent[0]?.text, '**Heads up**\n\nbody');
  });

  it('chunks a >28KB message into multiple sends', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry());
    const sent: Sent[] = [];
    const big = 'a'.repeat(30_000);
    const r = await deliverDm(deps(store, sent), { platformUserId: 'aad-user-1', message: big });
    assert.equal(r.status, 200);
    assert.equal(r.json.chunks, 2);
    assert.equal(sent.length, 2);
  });
});

describe('createDmRouter (auth boundary)', () => {
  const servers: Server[] = [];
  afterEach(() => {
    for (const s of servers.splice(0)) s.close();
  });

  async function mount(): Promise<string> {
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry());
    const app = express();
    app.use('/api/gateway', createDmRouter({ ...deps(store, []), adminToken: 'secret-token' }));
    const server = await new Promise<Server>((resolve) => {
      const s = app.listen(0, '127.0.0.1', () => resolve(s));
    });
    servers.push(server);
    const { port } = server.address() as AddressInfo;
    return `http://127.0.0.1:${port}`;
  }

  it('401s without a valid ADMIN_TOKEN bearer', async () => {
    const base = await mount();
    const noAuth = await fetch(`${base}/api/gateway/dm`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ platformUserId: 'aad-user-1', message: 'hi' }),
    });
    assert.equal(noAuth.status, 401);

    const badAuth = await fetch(`${base}/api/gateway/dm`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: 'Bearer wrong' },
      body: JSON.stringify({ platformUserId: 'aad-user-1', message: 'hi' }),
    });
    assert.equal(badAuth.status, 401);
  });

  it('delivers with the correct ADMIN_TOKEN bearer', async () => {
    const base = await mount();
    const res = await fetch(`${base}/api/gateway/dm`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: 'Bearer secret-token' },
      body: JSON.stringify({ platformUserId: 'aad-user-1', message: 'hi' }),
    });
    assert.equal(res.status, 200);
    const body = (await res.json()) as { delivered: boolean; botSlug: string };
    assert.equal(body.delivered, true);
    assert.equal(body.botSlug, 'maestro');
  });
});
