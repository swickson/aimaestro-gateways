/**
 * AMP Auto-Bootstrap (redesigned, shared)
 *
 * Registers gateway bot identities as AMP bridge agents: generates Ed25519 keys,
 * calls `/api/v1/register`, and persists everything under `~/.agent-messaging/`.
 *
 * REDESIGN vs the per-gateway copies (conforms to Maestro core's signed-off contract):
 *  - **agentId-keyed dirs (Bug #3 fix):** agent files are written under
 *    `agents/{agentId}/` (UUID), NOT `agents/{agentName}/`. Maestro's delivery
 *    resolver prefers the UUID dir ONLY IF it contains `config.json`, so we
 *    always write `config.json` into the UUID dir on registration.
 *  - **Canonical `.index.json`:** `{ name.toLowerCase(): agentId }`, additive
 *    merge, atomic tmp+rename — mirrors Maestro `lib/amp-inbox-writer.ts`.
 *  - **Sender-nested inbox awareness:** Maestro delivers to
 *    `agents/{agentId}/messages/inbox/{sanitizedSender}/{messageId}.json`;
 *    `sanitizeAddressForPath` is the shared helper for that nesting.
 *  - **Multi-bot registry + idempotency:** `bootstrapGateway` registers N bots,
 *    skips already-registered ones, retries only missing ones, and fails closed
 *    on duplicate public-key fingerprints.
 *  - **Legacy single-bot compatibility:** `bootstrapAMP` honors an explicit
 *    `AMP_INBOX_DIR` override and leaves pre-existing legacy `agents/{name}/`
 *    records untouched (no re-register, no overwrite).
 */

import * as crypto from 'crypto';
import * as fs from 'fs';
import * as path from 'path';

// ---------------------------------------------------------------------------
// Path helpers (resolved per-call so tests can point HOME at a tmp dir)
// ---------------------------------------------------------------------------

function homeDir(): string {
  return process.env.HOME || '/root';
}
function ampHome(): string {
  return path.join(homeDir(), '.agent-messaging');
}
function agentsDir(): string {
  return path.join(ampHome(), 'agents');
}
function indexFile(): string {
  return path.join(agentsDir(), '.index.json');
}
function gatewaysDir(): string {
  return path.join(ampHome(), 'gateways');
}

/**
 * Sanitize an AMP address into a filesystem-safe path segment, VERBATIM with
 * Maestro `lib/amp-inbox-writer.ts:sanitizeAddressForPath`. ORDER MATTERS:
 * replace `@`/`.` with `_` FIRST, then strip any remaining non-allowed chars.
 */
export function sanitizeAddressForPath(address: string): string {
  return address.replace(/[@.]/g, '_').replace(/[^a-zA-Z0-9_-]/g, '');
}

/** SHA-256 fingerprint of a PEM public key, prefixed `sha256-`. */
export function computeFingerprint(publicKeyPem: string): string {
  return 'sha256-' + crypto.createHash('sha256').update(publicKeyPem).digest('hex');
}

// ---------------------------------------------------------------------------
// Canonical .index.json writer (mirrors Maestro lib/amp-inbox-writer.ts)
// ---------------------------------------------------------------------------

function readIndex(): Record<string, string> {
  try {
    return JSON.parse(fs.readFileSync(indexFile(), 'utf-8'));
  } catch {
    return {};
  }
}

function writeIndex(index: Record<string, string>): void {
  const dir = agentsDir();
  fs.mkdirSync(dir, { recursive: true });
  const target = indexFile();
  const tmp = target + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(index, null, 2));
  fs.renameSync(tmp, target);
}

/**
 * Additive merge: set `index[name.toLowerCase()] = agentId` and atomically
 * rewrite `.index.json`. Never clobbers other entries.
 */
export function updateIndex(agentName: string, agentId: string): void {
  const index = readIndex();
  index[agentName.toLowerCase()] = agentId;
  writeIndex(index);
}

// ---------------------------------------------------------------------------
// Provider discovery + registration (unchanged behavior from the per-gateway copies)
// ---------------------------------------------------------------------------

interface ProviderInfo { tenant?: string; default_tenant?: string; domain?: string }

async function discoverProvider(maestroUrl: string): Promise<{ tenant: string; domain: string }> {
  const endpoints = [
    { url: `${maestroUrl}/.well-known/agent-messaging.json`, extract: (d: ProviderInfo) => ({ tenant: d.tenant || d.default_tenant, domain: d.domain }) },
    { url: `${maestroUrl}/api/v1/info`, extract: (d: ProviderInfo) => ({ tenant: d.tenant || d.default_tenant, domain: d.domain }) },
    { url: `${maestroUrl}/api/v1/health`, extract: (d: ProviderInfo) => ({ tenant: d.tenant || 'default', domain: d.domain || 'default.aimaestro.local' }) },
  ];

  for (const ep of endpoints) {
    try {
      const resp = await fetch(ep.url, { signal: AbortSignal.timeout(5000) });
      if (resp.ok) {
        const data = await resp.json();
        const result = ep.extract(data);
        if (result.tenant && result.domain) return { tenant: result.tenant, domain: result.domain };
      }
    } catch { /* Try next */ }
  }

  return { tenant: 'default', domain: 'default.aimaestro.local' };
}

function generateKeyPair(): { publicKeyPem: string; privateKeyPem: string } {
  const { publicKey, privateKey } = crypto.generateKeyPairSync('ed25519', {
    publicKeyEncoding: { type: 'spki', format: 'pem' },
    privateKeyEncoding: { type: 'pkcs8', format: 'pem' },
  });
  return { publicKeyPem: publicKey, privateKeyPem: privateKey };
}

async function registerAgent(
  maestroUrl: string, tenant: string, name: string,
  publicKeyPem: string, alias: string, metadata: Record<string, string>
): Promise<{ agent_id: string; address: string; api_key: string; tenant: string }> {
  const resp = await fetch(`${maestroUrl}/api/v1/register`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ tenant, name, public_key: publicKeyPem, key_algorithm: 'Ed25519', alias, metadata }),
    signal: AbortSignal.timeout(10000),
  });

  if (!resp.ok) {
    const error = await resp.text();
    throw new Error(`AMP registration failed (${resp.status}): ${error}`);
  }

  return resp.json();
}

/**
 * Write the agent's on-disk record under `agents/{agentId}/` (UUID-keyed).
 * `config.json` is written INTO the UUID dir (Maestro's resolver gates on its
 * existence). Returns the agent's inbox directory.
 */
function saveAgentFiles(
  agentName: string, agentId: string, address: string, tenant: string,
  domain: string, apiKey: string, publicKeyPem: string, privateKeyPem: string
): string {
  const agentDir = path.join(agentsDir(), agentId);
  const keysDir = path.join(agentDir, 'keys');
  const inboxDir = path.join(agentDir, 'messages', 'inbox');
  const sentDir = path.join(agentDir, 'messages', 'sent');
  const regDir = path.join(agentDir, 'registrations');

  for (const dir of [keysDir, inboxDir, sentDir, regDir]) {
    fs.mkdirSync(dir, { recursive: true });
  }

  fs.writeFileSync(path.join(keysDir, 'public.pem'), publicKeyPem, { mode: 0o644 });
  fs.writeFileSync(path.join(keysDir, 'private.pem'), privateKeyPem, { mode: 0o600 });
  // config.json shape mirrors what Maestro itself writes (resolver gates on the
  // file existing; extra keys are tolerated and kept for tool-compat).
  fs.writeFileSync(
    path.join(agentDir, 'config.json'),
    JSON.stringify(
      { version: 'amp/0.1', agent: { name: agentName, id: agentId }, address, tenant, domain, created_at: new Date().toISOString() },
      null, 2,
    ),
  );
  fs.writeFileSync(
    path.join(regDir, `${domain}.json`),
    JSON.stringify({ provider: domain, address, agent_id: agentId, api_key: apiKey, registered_at: new Date().toISOString() }, null, 2),
    { mode: 0o600 },
  );

  return inboxDir;
}

// ---------------------------------------------------------------------------
// Multi-bot gateway bootstrap (per-bot AMP identity, idempotent)
// ---------------------------------------------------------------------------

export interface BotConfig {
  /** Stable per-gateway slug, e.g. 'demo' (the routing key). */
  slug: string;
  /** Unique AMP agent name, `^[a-z0-9][a-z0-9-]{0,61}[a-z0-9]$`. */
  agentName: string;
  alias?: string;
  tenant?: string;
  metadata?: Record<string, string>;
}

export interface BotRegistration {
  agentName: string;
  agentId: string;
  address: string;
  apiKey: string;
  inboxDir: string;
  publicKeyFingerprint: string;
}

export interface GatewayRegistration {
  bots: Record<string, BotRegistration>;
}

export interface GatewayBootstrapOptions {
  /** Gateway identifier; registration state lives at gateways/{gatewayName}/registration.json. */
  gatewayName: string;
  maestroUrl: string;
  bots: BotConfig[];
  /** Optional tenant override applied to every bot lacking its own. */
  tenant?: string;
}

function registrationFile(gatewayName: string): string {
  return path.join(gatewaysDir(), gatewayName, 'registration.json');
}

function readRegistration(gatewayName: string): GatewayRegistration {
  try {
    const parsed = JSON.parse(fs.readFileSync(registrationFile(gatewayName), 'utf-8'));
    if (parsed && typeof parsed === 'object' && parsed.bots) return parsed as GatewayRegistration;
  } catch { /* no prior registration */ }
  return { bots: {} };
}

function writeRegistration(gatewayName: string, reg: GatewayRegistration): void {
  const file = registrationFile(gatewayName);
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = file + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(reg, null, 2), { mode: 0o600 });
  fs.renameSync(tmp, file);
}

/** A bot is considered already-registered only if its record AND its key files exist. */
function isBotRegistered(entry: BotRegistration | undefined): boolean {
  if (!entry || !entry.agentId || !entry.apiKey) return false;
  const keyFile = path.join(agentsDir(), entry.agentId, 'keys', 'private.pem');
  const cfgFile = path.join(agentsDir(), entry.agentId, 'config.json');
  return fs.existsSync(keyFile) && fs.existsSync(cfgFile);
}

/**
 * Bootstrap N bot identities for a gateway. Idempotent across restarts: bots
 * already present in `registration.json` (with key files on disk) are skipped;
 * only missing bots are registered. Fails closed if any two active bots share a
 * public-key fingerprint. The registration store is updated ADDITIVELY.
 */
export async function bootstrapGateway(options: GatewayBootstrapOptions): Promise<GatewayRegistration> {
  const { gatewayName, maestroUrl, bots } = options;
  console.log(`[BOOTSTRAP] Gateway ${gatewayName}: ${bots.length} bot identity(ies)`);

  const reg = readRegistration(gatewayName);
  let provider: { tenant: string; domain: string } | null = null;

  for (const bot of bots) {
    const existing = reg.bots[bot.slug];
    if (isBotRegistered(existing)) {
      // Self-heal a missing/stale canonical index mapping WITHOUT re-registering.
      // Identity is intact here (registration record + keys + config.json all
      // present), so the only thing that can be wrong is the name->uuid entry in
      // `.index.json`. A full re-register would mint a NEW identity via
      // /api/v1/register and orphan this bot — so repair just the index entry.
      if (readIndex()[existing.agentName.toLowerCase()] !== existing.agentId) {
        updateIndex(existing.agentName, existing.agentId);
        console.log(`[BOOTSTRAP]   ${bot.slug}: canonical index mapping missing/stale — repaired (no re-register)`);
      }
      console.log(`[BOOTSTRAP]   ${bot.slug}: already registered (${existing.agentId}) — skip`);
      continue;
    }

    if (!provider) {
      provider = await discoverProvider(maestroUrl);
      console.log(`[BOOTSTRAP]   discovered tenant=${provider.tenant} domain=${provider.domain}`);
    }
    const tenant = bot.tenant || options.tenant || provider.tenant;

    const { publicKeyPem, privateKeyPem } = generateKeyPair();
    const fingerprint = computeFingerprint(publicKeyPem);

    const result = await registerAgent(
      maestroUrl, tenant, bot.agentName, publicKeyPem,
      bot.alias || bot.agentName, bot.metadata || {},
    );
    const inboxDir = saveAgentFiles(
      bot.agentName, result.agent_id, result.address, tenant,
      provider.domain, result.api_key, publicKeyPem, privateKeyPem,
    );
    updateIndex(bot.agentName, result.agent_id);

    reg.bots[bot.slug] = {
      agentName: bot.agentName,
      agentId: result.agent_id,
      address: result.address,
      apiKey: result.api_key,
      inboxDir,
      publicKeyFingerprint: fingerprint,
    };
    console.log(`[BOOTSTRAP]   ${bot.slug}: registered ${result.address} (${result.agent_id})`);
  }

  // Fail-closed: no two active bots may share a public-key fingerprint.
  assertNoFingerprintCollision(reg);

  writeRegistration(gatewayName, reg);
  console.log(`[BOOTSTRAP] Gateway ${gatewayName}: bootstrap complete`);
  return reg;
}

function assertNoFingerprintCollision(reg: GatewayRegistration): void {
  const seen = new Map<string, string>(); // fingerprint -> slug
  for (const [slug, entry] of Object.entries(reg.bots)) {
    const fp = entry.publicKeyFingerprint;
    if (!fp) continue;
    const prior = seen.get(fp);
    if (prior) {
      throw new Error(
        `[BOOTSTRAP] fail-closed: bots '${prior}' and '${slug}' share public-key fingerprint ${fp}`,
      );
    }
    seen.set(fp, slug);
  }
}

// ---------------------------------------------------------------------------
// Legacy single-bot bootstrap (backward compatibility)
// ---------------------------------------------------------------------------

export interface BootstrapOptions {
  agentName: string;
  maestroUrl: string;
  tenant?: string;
  alias?: string;
  envFile: string;
  metadata?: Record<string, string>;
}

export interface BootstrapResult {
  apiKey: string;
  address: string;
  agentId: string;
  tenant: string;
  inboxDir: string;
}

function persistToEnv(envFile: string, vars: Record<string, string>): void {
  let content = '';
  try { content = fs.readFileSync(envFile, 'utf-8'); } catch { /* File doesn't exist */ }

  for (const [key, value] of Object.entries(vars)) {
    const regex = new RegExp(`^${key}=.*$`, 'm');
    if (regex.test(content)) {
      content = content.replace(regex, `${key}=${value}`);
    } else {
      content += `\n${key}=${value}`;
    }
  }

  fs.writeFileSync(envFile, content.trimStart());
  console.log(`[BOOTSTRAP] Config persisted to ${envFile}`);
}

/**
 * Single-bot bootstrap with legacy compatibility.
 *
 * - If `AMP_INBOX_DIR` is set in the environment, that directory is honored as
 *   the inbox (verbatim override).
 * - If `AMP_API_KEY` is already present AND a legacy `agents/{agentName}/`
 *   record exists on disk, the record is left UNTOUCHED (no re-register, no
 *   overwrite); only `AMP_INBOX_DIR` is (re)written to `.env`.
 * - Otherwise a fresh agent is registered under the UUID-keyed layout.
 */
export async function bootstrapAMP(options: BootstrapOptions): Promise<BootstrapResult> {
  const inboxOverride = process.env.AMP_INBOX_DIR && process.env.AMP_INBOX_DIR.trim() !== ''
    ? process.env.AMP_INBOX_DIR
    : null;

  // Legacy detection: already-registered single-bot using the old name-keyed layout.
  if (process.env.AMP_API_KEY && process.env.AMP_API_KEY.trim() !== '') {
    const legacyDir = path.join(agentsDir(), options.agentName);
    const legacyConfig = path.join(legacyDir, 'config.json');
    const legacyKey = path.join(legacyDir, 'keys', 'private.pem');
    if (fs.existsSync(legacyConfig) && fs.existsSync(legacyKey)) {
      const inboxDir = inboxOverride || path.join(legacyDir, 'messages', 'inbox');
      let agentId = '';
      try {
        const cfg = JSON.parse(fs.readFileSync(legacyConfig, 'utf-8'));
        agentId = cfg.agent?.id || cfg.agent_id || '';
      } catch { /* tolerate */ }
      const tenant = options.tenant || process.env.AMP_TENANT || 'default';
      persistToEnv(options.envFile, { AMP_INBOX_DIR: inboxDir });
      console.log(`[BOOTSTRAP] Legacy single-bot record detected for ${options.agentName} — left untouched`);
      return {
        apiKey: process.env.AMP_API_KEY,
        address: process.env.AMP_AGENT_ADDRESS || '',
        agentId,
        tenant,
        inboxDir,
      };
    }
  }

  console.log('[BOOTSTRAP] No existing registration — starting auto-registration...');
  console.log(`[BOOTSTRAP] Agent: ${options.agentName}`);

  const provider = await discoverProvider(options.maestroUrl);
  const tenant = options.tenant || provider.tenant;
  console.log(`[BOOTSTRAP] Discovered tenant: ${tenant}, domain: ${provider.domain}`);

  const { publicKeyPem, privateKeyPem } = generateKeyPair();
  console.log('[BOOTSTRAP] Generated Ed25519 key pair');

  const result = await registerAgent(
    options.maestroUrl, tenant, options.agentName, publicKeyPem,
    options.alias || options.agentName, options.metadata || {},
  );
  console.log(`[BOOTSTRAP] Registered: ${result.address} (ID: ${result.agent_id})`);

  const defaultInbox = saveAgentFiles(
    options.agentName, result.agent_id, result.address, tenant,
    provider.domain, result.api_key, publicKeyPem, privateKeyPem,
  );
  const inboxDir = inboxOverride || defaultInbox;
  updateIndex(options.agentName, result.agent_id);
  console.log(`[BOOTSTRAP] Saved agent files to ~/.agent-messaging/agents/${result.agent_id}/`);

  persistToEnv(options.envFile, {
    AMP_API_KEY: result.api_key,
    AMP_AGENT_ADDRESS: result.address,
    AMP_TENANT: tenant,
    AMP_INBOX_DIR: inboxDir,
  });
  console.log('[BOOTSTRAP] Auto-registration complete!');

  return { apiKey: result.api_key, address: result.address, agentId: result.agent_id, tenant, inboxDir };
}
