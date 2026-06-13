import assert from 'node:assert/strict';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, describe, it } from 'node:test';

import { bootstrapAMP, bootstrapGateway, sanitizeAddressForPath, updateIndex } from '../amp-bootstrap.js';

const originalHome = process.env.HOME;
const originalFetch = globalThis.fetch;
const originalEnv = {
  AMP_API_KEY: process.env.AMP_API_KEY,
  AMP_AGENT_ADDRESS: process.env.AMP_AGENT_ADDRESS,
  AMP_TENANT: process.env.AMP_TENANT,
  AMP_INBOX_DIR: process.env.AMP_INBOX_DIR,
};

function restoreEnv(): void {
  process.env.HOME = originalHome;
  for (const [key, value] of Object.entries(originalEnv)) {
    if (value === undefined) delete process.env[key];
    else process.env[key] = value;
  }
  globalThis.fetch = originalFetch;
}

function useTempHome(): string {
  const home = mkdtempSync(join(tmpdir(), 'common-bootstrap-'));
  process.env.HOME = home;
  delete process.env.AMP_API_KEY;
  delete process.env.AMP_AGENT_ADDRESS;
  delete process.env.AMP_TENANT;
  delete process.env.AMP_INBOX_DIR;
  return home;
}

function jsonResponse(data: unknown): Response {
  return new Response(JSON.stringify(data), {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
  });
}

function installFetchMock(registerIds: string[]) {
  const calls: Array<{ url: string; body?: string }> = [];
  let registerIndex = 0;
  globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
    const url = input.toString();
    calls.push({ url, body: typeof init?.body === 'string' ? init.body : undefined });
    if (url.endsWith('/.well-known/agent-messaging.json')) {
      return jsonResponse({ tenant: 'tenant-one', domain: 'example.aimaestro.local' });
    }
    if (url.endsWith('/api/v1/register')) {
      const id = registerIds[registerIndex++] ?? `uuid-${registerIndex}`;
      const body = init?.body ? JSON.parse(init.body.toString()) : {};
      return jsonResponse({
        agent_id: id,
        address: `${body.name}@example.aimaestro.local`,
        api_key: `api-key-${id}`,
        tenant: body.tenant,
      });
    }
    throw new Error(`unexpected fetch ${url}`);
  }) as typeof fetch;
  return calls;
}

function writeRegisteredBot(home: string, gatewayName: string, slug: string, fingerprint: string): void {
  const agentId = `${slug}-uuid`;
  const agentDir = join(home, '.agent-messaging', 'agents', agentId);
  mkdirSync(join(agentDir, 'keys'), { recursive: true });
  writeFileSync(join(agentDir, 'keys', 'private.pem'), 'private');
  writeFileSync(join(agentDir, 'config.json'), JSON.stringify({ agent: { id: agentId, name: slug } }));

  const registrationDir = join(home, '.agent-messaging', 'gateways', gatewayName);
  mkdirSync(registrationDir, { recursive: true });
  const registrationFile = join(registrationDir, 'registration.json');
  const registration = existsSync(registrationFile)
    ? JSON.parse(readFileSync(registrationFile, 'utf-8'))
    : { bots: {} };
  registration.bots[slug] = {
    agentName: `${slug}-bot`,
    agentId,
    address: `${slug}-bot@example.aimaestro.local`,
    apiKey: `${slug}-api-key`,
    inboxDir: join(agentDir, 'messages', 'inbox'),
    publicKeyFingerprint: fingerprint,
  };
  writeFileSync(registrationFile, JSON.stringify(registration, null, 2));
}

afterEach(() => {
  const home = process.env.HOME;
  restoreEnv();
  if (home?.startsWith(tmpdir())) rmSync(home, { recursive: true, force: true });
});

describe('amp-bootstrap helpers', () => {
  it('sanitizes sender addresses and writes additive lowercase index mappings atomically', () => {
    const home = useTempHome();
    assert.equal(
      sanitizeAddressForPath('Demo.Bot+evil/path@example.aimaestro.local'),
      'Demo_Botevilpath_example_aimaestro_local',
    );
    updateIndex('ExistingBot', 'uuid-1');
    updateIndex('DemoBot', 'uuid-2');
    const agentsDir = join(home, '.agent-messaging', 'agents');
    assert.deepEqual(JSON.parse(readFileSync(join(agentsDir, '.index.json'), 'utf-8')), {
      existingbot: 'uuid-1',
      demobot: 'uuid-2',
    });
    assert.equal(existsSync(join(agentsDir, '.index.json.tmp')), false);
  });
});

describe('bootstrapGateway', () => {
  it('registers missing bots under agentId-keyed dirs with config, registry, inbox, and lowercase index', async () => {
    const home = useTempHome();
    const calls = installFetchMock(['uuid-demo']);
    const registration = await bootstrapGateway({
      gatewayName: 'teams',
      maestroUrl: 'http://maestro.test',
      bots: [{ slug: 'demo', agentName: 'Teams-Demo-Bot' }],
    });
    const agentDir = join(home, '.agent-messaging', 'agents', 'uuid-demo');
    assert.equal(registration.bots.demo.agentId, 'uuid-demo');
    assert.equal(registration.bots.demo.inboxDir, join(agentDir, 'messages', 'inbox'));
    assert.equal(existsSync(join(agentDir, 'config.json')), true);
    assert.equal(existsSync(join(agentDir, 'keys', 'public.pem')), true);
    assert.equal(existsSync(join(agentDir, 'messages', 'inbox')), true);
    assert.equal(existsSync(join(agentDir, 'registrations', 'example.aimaestro.local.json')), true);
    assert.equal(existsSync(join(home, '.agent-messaging', 'agents', 'Teams-Demo-Bot')), false);
    const index = JSON.parse(readFileSync(join(home, '.agent-messaging', 'agents', '.index.json'), 'utf-8'));
    assert.deepEqual(index, { 'teams-demo-bot': 'uuid-demo' });
    assert.equal(calls.filter((call) => call.url.endsWith('/api/v1/register')).length, 1);
  });

  it('skips already-registered bots only when registration, keys, config, and index mapping exist', async () => {
    const home = useTempHome();
    writeRegisteredBot(home, 'teams', 'demo', 'sha256-existing');
    updateIndex('demo-bot', 'demo-uuid');
    const calls = installFetchMock(['should-not-register']);
    const registration = await bootstrapGateway({
      gatewayName: 'teams',
      maestroUrl: 'http://maestro.test',
      bots: [{ slug: 'demo', agentName: 'demo-bot' }],
    });
    assert.equal(registration.bots.demo.agentId, 'demo-uuid');
    assert.equal(calls.length, 0);
  });

  it('repairs a missing canonical index mapping without re-registering an intact bot identity', async () => {
    const home = useTempHome();
    writeRegisteredBot(home, 'teams', 'demo', 'sha256-existing');
    const calls = installFetchMock(['uuid-reregistered']);
    const registration = await bootstrapGateway({
      gatewayName: 'teams',
      maestroUrl: 'http://maestro.test',
      bots: [{ slug: 'demo', agentName: 'demo-bot' }],
    });
    assert.equal(registration.bots.demo.agentId, 'demo-uuid');
    assert.equal(calls.filter((call) => call.url.endsWith('/api/v1/register')).length, 0);
    const index = JSON.parse(readFileSync(join(home, '.agent-messaging', 'agents', '.index.json'), 'utf-8'));
    assert.equal(index['demo-bot'], 'demo-uuid');
  });

  it('throws fail-closed when active bots share a public key fingerprint', async () => {
    const home = useTempHome();
    writeRegisteredBot(home, 'teams', 'demo', 'sha256-duplicate');
    writeRegisteredBot(home, 'teams', 'echo', 'sha256-duplicate');
    await assert.rejects(
      bootstrapGateway({
        gatewayName: 'teams',
        maestroUrl: 'http://maestro.test',
        bots: [
          { slug: 'demo', agentName: 'demo-bot' },
          { slug: 'echo', agentName: 'echo-bot' },
        ],
      }),
      /share public-key fingerprint/,
    );
  });
});

describe('bootstrapAMP legacy compatibility', () => {
  it('honors AMP_INBOX_DIR override while using UUID-keyed registration layout', async () => {
    const home = useTempHome();
    const override = join(home, 'custom-inbox');
    process.env.AMP_INBOX_DIR = override;
    installFetchMock(['uuid-single']);
    const result = await bootstrapAMP({
      agentName: 'single-bot',
      maestroUrl: 'http://maestro.test',
      envFile: join(home, '.env'),
    });
    assert.equal(result.agentId, 'uuid-single');
    assert.equal(result.inboxDir, override);
    assert.equal(existsSync(join(home, '.agent-messaging', 'agents', 'uuid-single', 'config.json')), true);
    assert.match(readFileSync(join(home, '.env'), 'utf-8'), /AMP_INBOX_DIR=/);
  });

  it('leaves existing legacy name-keyed dirs untouched and does not re-register', async () => {
    const home = useTempHome();
    const legacyDir = join(home, '.agent-messaging', 'agents', 'legacy-bot');
    mkdirSync(join(legacyDir, 'keys'), { recursive: true });
    mkdirSync(join(legacyDir, 'messages', 'inbox'), { recursive: true });
    writeFileSync(join(legacyDir, 'config.json'), JSON.stringify({ agent_id: 'legacy-id', keep: 'sentinel' }));
    writeFileSync(join(legacyDir, 'keys', 'private.pem'), 'legacy-private-key');
    process.env.AMP_API_KEY = 'legacy-api-key';
    process.env.AMP_AGENT_ADDRESS = 'legacy-bot@example.aimaestro.local';
    const calls = installFetchMock(['should-not-register']);
    const result = await bootstrapAMP({
      agentName: 'legacy-bot',
      maestroUrl: 'http://maestro.test',
      envFile: join(home, '.env'),
    });
    assert.equal(result.apiKey, 'legacy-api-key');
    assert.equal(result.agentId, 'legacy-id');
    assert.equal(result.inboxDir, join(legacyDir, 'messages', 'inbox'));
    assert.equal(readdirSync(join(home, '.agent-messaging', 'agents')).includes('legacy-id'), false);
    assert.equal(calls.length, 0);
  });
});
