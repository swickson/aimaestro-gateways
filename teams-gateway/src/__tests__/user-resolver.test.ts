import assert from 'node:assert/strict';
import { afterEach, describe, it } from 'node:test';

import { createUserResolver } from '../user-resolver.js';

interface Call {
  url: string;
  method: string;
  body: unknown;
}

interface MockResponse {
  ok: boolean;
  status: number;
  json?: unknown;
}

const realFetch = globalThis.fetch;

function installFetch(handler: (call: Call) => MockResponse): Call[] {
  const calls: Call[] = [];
  globalThis.fetch = (async (input: unknown, init?: { method?: string; body?: string }) => {
    const url = String(input);
    const method = init?.method ?? 'GET';
    const body = init?.body ? JSON.parse(init.body) : undefined;
    const call = { url, method, body };
    calls.push(call);
    const r = handler(call);
    return {
      ok: r.ok,
      status: r.status,
      json: async () => r.json ?? {},
      text: async () => JSON.stringify(r.json ?? {}),
    };
  }) as unknown as typeof fetch;
  return calls;
}

async function tick(): Promise<void> {
  // Let fire-and-forget last-seen .then handlers settle.
  await new Promise((resolve) => setTimeout(resolve, 10));
}

afterEach(() => {
  globalThis.fetch = realFetch;
});

describe('user-resolver Phase 5 shapes', () => {
  it('auto-create (shape a) emits context.tenantId AND context.botSlug', async () => {
    const calls = installFetch((call) => {
      if (call.url.includes('/api/users/resolve')) return { ok: false, status: 404 };
      if (call.url.includes('/api/users/auto-create')) {
        return { ok: true, status: 200, json: { user: { id: 'u-1', displayName: 'Alice', role: 'external' } } };
      }
      return { ok: false, status: 500 };
    });

    const resolver = createUserResolver({ maestroUrl: 'http://maestro', apiKey: 'k' });
    const user = await resolver.resolve('aad-1', 'Alice', 'tenant-9', 'echo');
    assert.equal(user?.id, 'u-1');

    const create = calls.find((c) => c.url.includes('/api/users/auto-create'));
    assert.ok(create, 'auto-create called');
    assert.deepEqual(create?.body, {
      platform: 'teams',
      platformUserId: 'aad-1',
      handle: 'Alice',
      context: { tenantId: 'tenant-9', botSlug: 'echo' },
    });
  });

  it('first-contact also PATCHes /last-seen after auto-create (shape b every inbound)', async () => {
    const calls = installFetch((call) => {
      if (call.url.includes('/api/users/resolve')) return { ok: false, status: 404 };
      if (call.url.includes('/api/users/auto-create')) {
        return { ok: true, status: 200, json: { user: { id: 'u-1', displayName: 'Alice', role: 'external' } } };
      }
      if (call.url.endsWith('/api/users/u-1/last-seen')) return { ok: true, status: 200 };
      return { ok: false, status: 500 };
    });

    const resolver = createUserResolver({ maestroUrl: 'http://maestro', apiKey: 'k' });
    await resolver.resolve('aad-1', 'Alice', 'tenant-9', 'echo');
    await tick();

    const lastSeen = calls.find((c) => c.url.endsWith('/api/users/u-1/last-seen'));
    assert.ok(lastSeen, 'last-seen PATCH fired on first-contact auto-create');
    assert.equal(lastSeen?.method, 'PATCH');
    assert.deepEqual(lastSeen?.body, {
      platform: 'teams',
      platformUserId: 'aad-1',
      context: { botSlug: 'echo' },
    });
  });

  it('every-inbound (shape b) PATCHes /last-seen with the exact body', async () => {
    const calls = installFetch((call) => {
      if (call.url.includes('/api/users/resolve')) {
        return { ok: true, status: 200, json: { user: { id: 'u-7', displayName: 'Bob', role: 'external' } } };
      }
      if (call.url.endsWith('/api/users/u-7/last-seen')) return { ok: true, status: 200 };
      return { ok: false, status: 500 };
    });

    const resolver = createUserResolver({ maestroUrl: 'http://maestro', apiKey: 'k' });
    await resolver.resolve('aad-7', 'Bob', 'tenant-1', 'maestro');

    const lastSeen = calls.find((c) => c.url.endsWith('/api/users/u-7/last-seen'));
    assert.ok(lastSeen, 'last-seen PATCH fired');
    assert.equal(lastSeen?.method, 'PATCH');
    assert.deepEqual(lastSeen?.body, {
      platform: 'teams',
      platformUserId: 'aad-7',
      context: { botSlug: 'maestro' },
    });
  });

  it('fires last-seen on a cache hit too (every inbound)', async () => {
    const calls = installFetch((call) => {
      if (call.url.includes('/api/users/resolve')) {
        return { ok: true, status: 200, json: { user: { id: 'u-7', displayName: 'Bob', role: 'external' } } };
      }
      return { ok: true, status: 200 };
    });

    const resolver = createUserResolver({ maestroUrl: 'http://maestro', apiKey: 'k' });
    await resolver.resolve('aad-7', 'Bob', 'tenant-1', 'maestro'); // miss -> resolve + last-seen
    await resolver.resolve('aad-7', 'Bob', 'tenant-1', 'echo'); // cache hit -> last-seen only

    const lastSeenCalls = calls.filter((c) => c.url.endsWith('/api/users/u-7/last-seen'));
    assert.equal(lastSeenCalls.length, 2, 'last-seen fired on both inbounds');
    assert.equal((lastSeenCalls[1]?.body as { context: { botSlug: string } }).context.botSlug, 'echo');
    // The resolve endpoint was only hit once (second was a cache hit).
    assert.equal(calls.filter((c) => c.url.includes('/api/users/resolve')).length, 1);
  });

  it('swallows a 404 from /last-seen without throwing (route not yet deployed)', async () => {
    installFetch((call) => {
      if (call.url.includes('/api/users/resolve')) {
        return { ok: true, status: 200, json: { user: { id: 'u-9', displayName: 'Carol', role: 'external' } } };
      }
      if (call.url.endsWith('/api/users/u-9/last-seen')) return { ok: false, status: 404 };
      return { ok: false, status: 500 };
    });

    const resolver = createUserResolver({ maestroUrl: 'http://maestro', apiKey: 'k' });
    const user = await assert.doesNotReject(() => resolver.resolve('aad-9', 'Carol', 'tenant-1', 'maestro'));
    void user;
    await tick(); // let the 404-handling .then run

    // Resolution still succeeded despite the 404 last-seen.
    const again = await resolver.resolve('aad-9', 'Carol', 'tenant-1', 'maestro');
    assert.equal(again?.id, 'u-9');
  });

  it('warns once on repeated /last-seen 404s', async () => {
    installFetch((call) => {
      if (call.url.includes('/api/users/resolve')) {
        return { ok: true, status: 200, json: { user: { id: 'u-9', displayName: 'Carol', role: 'external' } } };
      }
      return { ok: false, status: 404 };
    });

    const warnings: string[] = [];
    const realWarn = console.warn;
    console.warn = (msg?: unknown) => {
      warnings.push(String(msg));
    };
    try {
      const resolver = createUserResolver({ maestroUrl: 'http://maestro', apiKey: 'k' });
      await resolver.resolve('aad-9', 'Carol', 'tenant-1', 'maestro'); // resolve + last-seen 404
      await tick();
      await resolver.resolve('aad-9', 'Carol', 'tenant-1', 'echo'); // cache hit -> last-seen 404 again
      await tick();
    } finally {
      console.warn = realWarn;
    }

    const lastSeenWarnings = warnings.filter((w) => w.includes('/api/users/:id/last-seen'));
    assert.equal(lastSeenWarnings.length, 1, 'warned exactly once across two 404s');
  });
});
