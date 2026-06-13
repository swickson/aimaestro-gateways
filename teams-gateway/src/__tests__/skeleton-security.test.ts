import assert from 'node:assert/strict';
import { spawn, type ChildProcess } from 'node:child_process';
import { once } from 'node:events';
import { createServer, type AddressInfo } from 'node:net';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, it } from 'node:test';

import { loadBotRegistry } from '../bot-registry.js';
import { loadConfig } from '../config.js';

const TEAMS_DIR = resolve(dirname(fileURLToPath(import.meta.url)), '../..');

const BOT = {
  slug: 'maestro',
  appId: '00000000-0000-0000-0000-000000000000',
  appPassword: 'secret',
  appTenantId: '11111111-1111-1111-1111-111111111111',
  defaultAgent: 'ops-agent@example.aimaestro.local',
};

function botsJson(overrides: Partial<typeof BOT> = {}): string {
  return JSON.stringify([{ ...BOT, ...overrides }]);
}

function withEnv<T>(env: Record<string, string | undefined>, fn: () => T): T {
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
    return fn();
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

async function freePort(): Promise<number> {
  const server = createServer();
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  const address = server.address();
  assert.ok(isAddressInfo(address));
  const port = address.port;
  await new Promise<void>((resolveClose, reject) => {
    server.close((err) => err ? reject(err) : resolveClose());
  });
  return port;
}

function isAddressInfo(address: string | AddressInfo | null): address is AddressInfo {
  return typeof address === 'object' && address !== null;
}

async function waitForHealth(baseUrl: string, proc: ChildProcess): Promise<unknown> {
  const deadline = Date.now() + 15_000;
  let lastErr: unknown;
  while (Date.now() < deadline) {
    if (proc.exitCode !== null) {
      throw new Error(`server exited early with ${proc.exitCode}`);
    }
    try {
      const res = await fetch(`${baseUrl}/health`);
      if (res.ok) return await res.json();
    } catch (err) {
      lastErr = err;
    }
    await new Promise((resolveWait) => setTimeout(resolveWait, 100));
  }
  throw new Error(`server did not become healthy: ${String(lastErr)}`);
}

async function stop(proc: ChildProcess): Promise<void> {
  if (proc.exitCode !== null) return;
  if (proc.pid !== undefined) {
    try {
      process.kill(-proc.pid, 'SIGTERM');
    } catch {
      proc.kill('SIGTERM');
    }
  } else {
    proc.kill('SIGTERM');
  }
  const timeout = setTimeout(() => {
    if (proc.pid !== undefined) {
      try {
        process.kill(-proc.pid, 'SIGKILL');
      } catch {
        proc.kill('SIGKILL');
      }
    } else {
      proc.kill('SIGKILL');
    }
  }, 5_000);
  try {
    await once(proc, 'exit');
  } finally {
    clearTimeout(timeout);
  }
}

function safeJson(body: string): unknown {
  try {
    return JSON.parse(body);
  } catch {
    return body;
  }
}

describe('Teams bot registry validation', () => {
  it('accepts a valid bot and derives a unique AMP agent name', () => {
    const bots = loadBotRegistry(botsJson());
    assert.equal(bots.length, 1);
    assert.equal(bots[0]?.slug, 'maestro');
    assert.equal(bots[0]?.agentName, 'teams-maestro-bot');
  });

  it('fails closed for absent, invalid, empty, reserved, malformed, and incomplete registries', () => {
    const invalidRegistries = [
      undefined,
      '',
      'not json',
      '[]',
      botsJson({ slug: 'admin' }),
      botsJson({ slug: 'api' }),
      botsJson({ slug: 'UpperCase' }),
      JSON.stringify([{ ...BOT, appPassword: '' }]),
      JSON.stringify([{ ...BOT, appTenantId: '   ' }]),
    ];

    for (const raw of invalidRegistries) {
      assert.throws(() => loadBotRegistry(raw), /REGISTRY/);
    }
  });

  it('fails closed on duplicate identity fields across bots', () => {
    const duplicateCases = [
      [{ ...BOT }, { ...BOT, appId: '22222222-2222-2222-2222-222222222222', defaultAgent: 'other@example.aimaestro.local' }],
      [{ ...BOT }, { ...BOT, slug: 'demo', defaultAgent: 'other@example.aimaestro.local' }],
      [{ ...BOT }, { ...BOT, slug: 'demo', appId: '22222222-2222-2222-2222-222222222222' }],
      [{ ...BOT, agentName: 'shared-agent' }, { ...BOT, slug: 'demo', appId: '22222222-2222-2222-2222-222222222222', defaultAgent: 'other@example.aimaestro.local', agentName: 'shared-agent' }],
    ];

    for (const value of duplicateCases) {
      assert.throws(() => loadBotRegistry(JSON.stringify(value)), /duplicate/);
    }
  });
});

describe('Teams gateway config fail-closed behavior', () => {
  it('throws at startup when ADMIN_TOKEN is empty or blank', () => {
    for (const token of ['', '   ', undefined]) {
      withEnv({
        ADMIN_TOKEN: token,
        TEAMS_BOTS: botsJson(),
        PORT: '3024',
      }, () => {
        assert.throws(() => loadConfig(), /ADMIN_TOKEN/);
      });
    }
  });

  it('accepts only tenant-scoped operator refs and does not preserve bare Aad IDs', () => {
    withEnv({
      ADMIN_TOKEN: 'admin-secret',
      TEAMS_BOTS: botsJson(),
      OPERATOR_AAD_OBJECT_IDS: 'tenant-a:operator-a,bare-object-id,tenant-b:operator-b',
      PORT: '3024',
      HOST: '127.0.0.1',
      TEAMS_DRY_RUN: '1',
    }, () => {
      const config = loadConfig();
      assert.deepEqual(config.operatorAadObjectIds, [
        { tenantId: 'tenant-a', aadObjectId: 'operator-a' },
        { tenantId: 'tenant-b', aadObjectId: 'operator-b' },
      ]);
    });
  });
});

describe('Teams gateway skeleton HTTP boundary', () => {
  it('boots in dry-run without network registration and gates only management routes with ADMIN_TOKEN', async () => {
    const port = await freePort();
    const env = {
      ...process.env,
      PORT: String(port),
      HOST: '127.0.0.1',
      ADMIN_TOKEN: 'admin-secret',
      AIMAESTRO_URL: 'http://127.0.0.1:1',
      TEAMS_BOTS: botsJson(),
      TEAMS_DRY_RUN: '1',
      DEBUG: '',
    };
    const proc = spawn('npm', ['start'], {
      cwd: TEAMS_DIR,
      env,
      detached: true,
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    let output = '';
    proc.stdout.on('data', (chunk) => {
      output += chunk.toString();
    });
    proc.stderr.on('data', (chunk) => {
      output += chunk.toString();
    });

    try {
      const baseUrl = `http://127.0.0.1:${port}`;
      const health = await waitForHealth(baseUrl, proc) as {
        status?: string;
        service?: string;
        bootstrap?: string;
        adapter?: { ready?: boolean; sharedExpressAdapter?: boolean; botCount?: number };
        bots?: Array<{ slug?: string; messagingPath?: string; authEnabled?: boolean; ampAddress?: string | null }>;
      };
      assert.equal(health.status, 'ok');
      assert.equal(health.service, 'teams-gateway');
      assert.equal(health.bootstrap, 'dry-run');
      assert.deepEqual(health.adapter, { ready: true, sharedExpressAdapter: true, botCount: 1 });
      assert.equal(health.bots?.[0]?.slug, 'maestro');
      assert.equal(health.bots?.[0]?.messagingPath, '/api/maestro/messages');
      assert.equal(health.bots?.[0]?.authEnabled, true);
      assert.equal(health.bots?.[0]?.ampAddress, null);

      const unauthAdmin = await fetch(`${baseUrl}/api/admin/stats`);
      assert.equal(unauthAdmin.status, 401);
      assert.deepEqual(await unauthAdmin.json(), { error: 'Unauthorized' });

      const authAdmin = await fetch(`${baseUrl}/api/admin/stats`, {
        headers: { authorization: 'Bearer admin-secret' },
      });
      assert.equal(authAdmin.status, 200);
      const stats = await authAdmin.json() as { service?: string; bootstrap?: string; bots?: unknown[] };
      assert.equal(stats.service, 'teams-gateway');
      assert.equal(stats.bootstrap, 'dry-run');
      assert.equal(stats.bots?.length, 1);

      const botPath = await fetch(`${baseUrl}/api/maestro/messages`, { method: 'POST' });
      const botBody = await botPath.text();
      assert.notDeepEqual(safeJson(botBody), { error: 'Unauthorized' });

      assert.match(output, /DRY-RUN/);
      assert.doesNotMatch(output, /\[FATAL\]/);
    } finally {
      await stop(proc);
    }
  });
});
