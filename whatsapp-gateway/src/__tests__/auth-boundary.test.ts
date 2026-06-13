import assert from 'node:assert/strict';
import { once } from 'node:events';
import type { Server } from 'node:http';
import { describe, it } from 'node:test';

import { loadConfig } from '../config.js';
import { createHttpApp } from '../server.js';
import type { GatewayConfig } from '../types.js';

async function withEnv<T>(env: Record<string, string | undefined>, fn: () => Promise<T>): Promise<T> {
  const previous = new Map<string, string | undefined>();
  for (const key of Object.keys(env)) {
    previous.set(key, process.env[key]);
    const value = env[key];
    if (value === undefined) {
      delete process.env[key];
    } else {
      process.env[key] = value;
    }
  }
  try {
    return await fn();
  } finally {
    for (const [key, value] of previous) {
      if (value === undefined) {
        delete process.env[key];
      } else {
        process.env[key] = value;
      }
    }
  }
}

async function listen(app: ReturnType<typeof createHttpApp>): Promise<{ server: Server; baseUrl: string }> {
  const server = app.listen(0, '127.0.0.1');
  await once(server, 'listening');
  const address = server.address();
  assert.ok(typeof address === 'object' && address !== null);
  return { server, baseUrl: `http://127.0.0.1:${address.port}` };
}

async function close(server: Server): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    server.close((err) => err ? reject(err) : resolve());
  });
}

function config(adminToken = 'admin-secret'): GatewayConfig {
  return {
    port: 0,
    debug: false,
    amp: {
      apiKey: 'amp-key',
      agentAddress: 'whatsapp-bot@example.aimaestro.local',
      maestroUrl: 'http://127.0.0.1:1',
      defaultAgent: 'ops@example.aimaestro.local',
      tenant: 'example',
      inboxDir: '/tmp/whatsapp-inbox',
    },
    whatsapp: {
      stateDir: '/tmp/whatsapp-state',
      allowFrom: [],
      dmPolicy: 'allowlist',
      sendReadReceipts: true,
      textChunkLimit: 4000,
    },
    routing: {
      phones: {},
      default: { agent: 'ops@example.aimaestro.local' },
    },
    outbound: {
      pollIntervalMs: 5000,
    },
    operatorPhones: [],
    adminToken,
  };
}

describe('WhatsApp gateway auth boundary', () => {
  it('fails config load when ADMIN_TOKEN is missing or blank', async () => {
    for (const token of ['', '   ', undefined]) {
      await withEnv({ ADMIN_TOKEN: token }, async () => {
        await assert.rejects(() => loadConfig(), /ADMIN_TOKEN/);
      });
    }
  });

  it('keeps /health and /status public and rejects unauthenticated or malformed /api requests', async () => {
    const { server, baseUrl } = await listen(createHttpApp(config()));
    try {
      const health = await fetch(`${baseUrl}/health`);
      assert.equal(health.status, 200);

      const status = await fetch(`${baseUrl}/status`);
      assert.equal(status.status, 200);

      const invalidHeaders: Array<HeadersInit | undefined> = [
        undefined,
        { authorization: 'Bearer wrong' },
        { authorization: 'Basic admin-secret' },
      ];

      for (const headers of invalidHeaders) {
        const res = await fetch(`${baseUrl}/api/activity`, { headers });
        assert.equal(res.status, 401);
        assert.deepEqual(await res.json(), { error: 'Unauthorized' });
      }
    } finally {
      await close(server);
    }
  });

  it('allows /api requests with the valid ADMIN_TOKEN bearer', async () => {
    const { server, baseUrl } = await listen(createHttpApp(config()));
    try {
      const res = await fetch(`${baseUrl}/api/activity`, {
        headers: { authorization: 'Bearer admin-secret' },
      });
      assert.equal(res.status, 200);
      assert.equal((await res.json() as { count?: number }).count, 0);
    } finally {
      await close(server);
    }
  });
});
