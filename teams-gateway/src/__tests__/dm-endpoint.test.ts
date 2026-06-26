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

interface Created {
  botSlug: string;
  tenantId: string;
  aadObjectId: string;
  // FLIGHT1 #25: text/markdown dropped — createColdStartConversation no longer posts
  // an inline activity (CreateColdStartConversationInput shed these fields).
}

function deps(
  store: ThreadStore,
  sent: Sent[],
  markdownDefault = true,
  overrides: Partial<DmDeps> = {},
): DmDeps {
  return {
    threadStore: store,
    knownBots: new Set(['maestro', 'echo']),
    coldStartEnabled: false,
    markdownDefault,
    sendChunk: async (botSlug, conversationId, text, markdown) => {
      sent.push({ botSlug, conversationId, text, markdown });
    },
    ...overrides,
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

  it('reuses the only live bot when botSlug is omitted for a single-bot user', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry({ botSlug: 'maestro', conversationId: 'conv-maestro', ampMessageId: 'amp-maestro' }));
    const sent: Sent[] = [];

    const r = await deliverDm(deps(store, sent), { platformUserId: 'aad-user-1', message: 'single-bot reuse' });

    assert.equal(r.status, 200);
    assert.equal(r.json.botSlug, 'maestro');
    assert.equal(sent.length, 1);
    assert.equal(sent[0]?.botSlug, 'maestro');
    assert.equal(sent[0]?.conversationId, 'conv-maestro');
  });

  it('409s ambiguous_bot when a multi-bot user omits botSlug instead of recency-guessing', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry({ botSlug: 'leoai', conversationId: 'conv-leoai', ampMessageId: 'amp-leoai', createdAt: 100 }));
    store.record(entry({ botSlug: 'zach', conversationId: 'conv-zach', ampMessageId: 'amp-zach', createdAt: 200 }));
    const sent: Sent[] = [];

    const r = await deliverDm(deps(store, sent), { platformUserId: 'aad-user-1', message: 'incident regression' });

    assert.equal(r.status, 409);
    assert.equal(r.json.error, 'undeliverable');
    assert.equal(r.json.reason, 'ambiguous_bot');
    assert.deepEqual(r.json.candidates, ['leoai', 'zach']);
    assert.equal(sent.length, 0, 'must not deliver via the most-recent zach mapping');
  });

  it('honors a pinned valid bot even when another bot is more recent', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry({ botSlug: 'leoai', conversationId: 'conv-leoai', ampMessageId: 'amp-leoai', createdAt: 100 }));
    store.record(entry({ botSlug: 'zach', conversationId: 'conv-zach', ampMessageId: 'amp-zach', createdAt: 200 }));
    const sent: Sent[] = [];

    const r = await deliverDm(
      deps(store, sent, true, { knownBots: new Set(['leoai', 'zach']) }),
      { platformUserId: 'aad-user-1', botSlug: 'leoai', message: 'pinned' },
    );

    assert.equal(r.status, 200);
    assert.equal(r.json.botSlug, 'leoai');
    assert.equal(sent.length, 1);
    assert.equal(sent[0]?.botSlug, 'leoai');
    assert.equal(sent[0]?.conversationId, 'conv-leoai');
  });

  it('409s undeliverable when the user has no prior contact and cold-start is disabled', async () => {
    const sent: Sent[] = [];
    const r = await deliverDm(deps(createThreadStore(), sent), { platformUserId: 'aad-user-1', message: 'hi' });
    assert.equal(r.status, 409);
    assert.equal(r.json.reason, 'no_prior_contact');
    assert.equal(sent.length, 0);
  });

  it('400s when cold-start needs a tenantId but the caller omitted it', async () => {
    const sent: Sent[] = [];
    const r = await deliverDm(
      deps(createThreadStore(), sent, true, {
        coldStartEnabled: true,
        createColdStartConversation: async () => {
          throw new Error('should not create without tenant');
        },
      }),
      { platformUserId: 'aad-user-1', botSlug: 'maestro', message: 'hi' },
    );
    assert.equal(r.status, 400);
    assert.equal(r.json.error, 'bad_request');
    assert.equal(sent.length, 0);
  });

  it('409s when cold-start is enabled but no send bot can be resolved', async () => {
    const sent: Sent[] = [];
    const r = await deliverDm(
      deps(createThreadStore(), sent, true, { coldStartEnabled: true }),
      { platformUserId: 'aad-user-1', tenantId: 'tenant-1', message: 'hi' },
    );
    assert.equal(r.status, 409);
    assert.equal(r.json.reason, 'no_send_bot');
    assert.equal(sent.length, 0);
  });

  it('cold-starts a pinned bot, records the ConversationReference, then reuses it', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    const sent: Sent[] = [];
    const created: Created[] = [];
    const d = deps(store, sent, true, {
      coldStartEnabled: true,
      createColdStartConversation: async (input) => {
        created.push(input);
        return {
          conversationId: 'cold-conv-1',
          rootActivityId: 'cold-activity-1',
          reference: context('cold-conv-1').reference,
        };
      },
    });

    const first = await deliverDm(d, {
      platformUserId: 'aad-user-1',
      botSlug: 'maestro',
      tenantId: 'tenant-1',
      message: 'hello cold',
    });
    assert.equal(first.status, 200);
    assert.deepEqual(first.json, { delivered: true, botSlug: 'maestro', chunks: 1, coldStart: true });
    assert.equal(created.length, 1);
    assert.equal(created[0]?.aadObjectId, 'aad-user-1');
    assert.equal(created[0]?.tenantId, 'tenant-1');
    assert.equal(sent.length, 1);
    assert.equal(sent[0]?.conversationId, 'cold-conv-1');
    assert.equal(sent[0]?.text, 'hello cold');

    const recorded = store.findByUserAndBot('aad-user-1', 'maestro');
    assert.equal(recorded?.conversationId, 'cold-conv-1');
    assert.equal(recorded?.context.tenantId, 'tenant-1');

    const second = await deliverDm(d, { platformUserId: 'aad-user-1', botSlug: 'maestro', message: 'reuse' });
    assert.equal(second.status, 200);
    assert.equal(created.length, 1, 'second DM reuses the stored conversation');
    assert.equal(sent.length, 2);
    assert.equal(sent[1]?.conversationId, 'cold-conv-1');
    assert.equal(sent[1]?.text, 'reuse');
  });

  it('cold-start sends every chunk via App.send after createConversation ensures the 1:1', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    const sent: Sent[] = [];
    const created: Created[] = [];
    const d = deps(store, sent, true, {
      coldStartEnabled: true,
      createColdStartConversation: async (input) => {
        created.push(input);
        return {
          conversationId: 'cold-conv-2',
          rootActivityId: 'cold-activity-2',
          reference: context('cold-conv-2').reference,
        };
      },
    });

    const big = 'a'.repeat(30_000);
    const r = await deliverDm(d, {
      platformUserId: 'aad-user-1',
      botSlug: 'maestro',
      tenantId: 'tenant-1',
      message: big,
    });
    assert.equal(r.status, 200);
    assert.equal(r.json.chunks, 2);
    assert.equal(created.length, 1);
    assert.equal(sent.length, 2);
    assert.equal(sent.map((s) => s.text).join(''), big);
    assert.ok(sent.every((s) => s.conversationId === 'cold-conv-2'));
  });

  it('#25 existing 1:1 cold-start still delivers chunk[0] via App.send', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    const sent: Sent[] = [];
    const created: Created[] = [];
    const d = deps(store, sent, true, {
      coldStartEnabled: true,
      createColdStartConversation: async (input) => {
        created.push(input);
        return {
          conversationId: 'existing-personal-conv',
          rootActivityId: 'existing-root',
          reference: context('existing-personal-conv').reference,
        };
      },
    });

    const r = await deliverDm(d, {
      platformUserId: 'aad-user-1',
      botSlug: 'maestro',
      tenantId: 'tenant-1',
      message: 'hello existing',
    });

    assert.equal(r.status, 200);
    assert.equal(created.length, 1);
    assert.equal(sent.length, 1);
    assert.equal(sent[0]?.conversationId, 'existing-personal-conv');
    assert.equal(sent[0]?.text, 'hello existing');
  });

  it('maps Bot Connector createConversation failures to clear undeliverable reasons', async () => {
    const sent: Sent[] = [];
    const r = await deliverDm(
      deps(createThreadStore(), sent, true, {
        coldStartEnabled: true,
        createColdStartConversation: async () => {
          const err = new Error('forbidden') as Error & { response: { status: number } };
          err.response = { status: 403 };
          throw err;
        },
      }),
      { platformUserId: 'aad-user-1', botSlug: 'maestro', tenantId: 'tenant-1', message: 'hi' },
    );
    assert.equal(r.status, 409);
    assert.equal(r.json.reason, 'bot_not_installed_or_forbidden');
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

  it('does not fall back to the most-recent bot when no botSlug is supplied for a multi-bot user', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry({ botSlug: 'maestro', conversationId: 'conv-m', ampMessageId: 'amp-m', createdAt: 100 }));
    store.record(entry({ botSlug: 'echo', conversationId: 'conv-e', ampMessageId: 'amp-e', createdAt: 200 }));
    const sent: Sent[] = [];
    const r = await deliverDm(deps(store, sent), { platformUserId: 'aad-user-1', message: 'ping' });
    assert.equal(r.status, 409);
    assert.equal(r.json.reason, 'ambiguous_bot');
    assert.deepEqual(r.json.candidates, ['echo', 'maestro']);
    assert.equal(sent.length, 0);
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
