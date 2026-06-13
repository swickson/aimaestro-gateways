/**
 * Teams gateway — configuration loading.
 *
 * Loads env, validates the multi-bot registry (delegated to bot-registry.ts), and
 * FAILS CLOSED at startup: empty `ADMIN_TOKEN` throws here (and again in common's
 * `createAuthMiddleware`), an invalid `TEAMS_BOTS` throws, an invalid `PORT`
 * throws. No bypass-when-empty branch (CLAUDE.md bugs #1/#6).
 */

import * as os from 'node:os';
import * as path from 'node:path';
import * as dotenv from 'dotenv';
import { loadBotRegistry } from './bot-registry.js';
import type { GatewayConfig, OperatorAadRef } from './types.js';

dotenv.config();

const TRUTHY = new Set(['1', 'true', 'yes', 'on']);
const DEFAULT_POLL_INTERVAL_MS = 3000;
const DEFAULT_CACHE_USER_TTL_MS = 300_000; // 5 min (shared CACHE_USER_TTL_MS default).
const DEFAULT_SNAPSHOT_INTERVAL_MS = 60_000; // 60s crash-safety floor between graceful saves.

/**
 * Parse a positive-integer env var with a default. Invalid (non-numeric, <=0)
 * falls back to the default with a warning rather than throwing — these are
 * tuning knobs, not security-critical (unlike PORT/ADMIN_TOKEN, which fail closed).
 */
function parsePositiveInt(raw: string | undefined, fallback: number, name: string): number {
  if (raw === undefined || raw.trim() === '') return fallback;
  const n = Number(raw);
  if (!Number.isInteger(n) || n <= 0) {
    console.warn(`[CONFIG] Invalid ${name}: "${raw}" (expected positive integer) — using ${fallback}.`);
    return fallback;
  }
  return n;
}

function isTruthy(raw: string | undefined): boolean {
  return raw !== undefined && TRUTHY.has(raw.trim().toLowerCase());
}

/** Comma-separated bind hosts (discord parity). Defaults to loopback. */
function parseHosts(raw: string | undefined): string[] {
  if (!raw || raw.trim() === '') return ['127.0.0.1'];
  return raw.split(',').map((s) => s.trim()).filter(Boolean);
}

/**
 * Parse `OPERATOR_AAD_OBJECT_IDS` as `tenantId:aadObjectId[,…]` — tenant-scoped
 * per red-team §0.2. Malformed entries (not exactly two non-empty parts) are
 * skipped with a warning rather than silently treated as a bare object-id.
 */
function parseOperatorAadRefs(raw: string | undefined): OperatorAadRef[] {
  if (!raw || raw.trim() === '') return [];
  return raw
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean)
    .map((entry): OperatorAadRef | null => {
      const parts = entry.split(':');
      if (parts.length !== 2 || parts.some((p) => p.trim() === '')) {
        console.warn(`[CONFIG] Skipping malformed OPERATOR_AAD_OBJECT_IDS entry: "${entry}" (expected tenantId:aadObjectId).`);
        return null;
      }
      const [tenantId, aadObjectId] = parts.map((p) => p.trim());
      return { tenantId, aadObjectId };
    })
    .filter((e): e is OperatorAadRef => e !== null);
}

export function loadConfig(): GatewayConfig {
  const adminToken = process.env.ADMIN_TOKEN ?? '';
  if (adminToken.trim() === '') {
    throw new Error('[CONFIG] ADMIN_TOKEN is required and cannot be empty (fail-closed).');
  }

  const port = Number(process.env.PORT ?? '3024');
  if (!Number.isInteger(port) || port <= 0 || port > 65535) {
    throw new Error(`[CONFIG] Invalid PORT: "${process.env.PORT}" (expected 1–65535).`);
  }

  const maestroUrl =
    process.env.AIMAESTRO_URL || process.env.AMP_MAESTRO_URL || 'http://127.0.0.1:23000';

  const bots = loadBotRegistry(process.env.TEAMS_BOTS);

  const threadStorePath =
    process.env.THREAD_STORE_PATH && process.env.THREAD_STORE_PATH.trim() !== ''
      ? process.env.THREAD_STORE_PATH.trim()
      : path.join(os.homedir(), '.agent-messaging', 'teams-gateway', 'thread-store.json');

  return {
    port,
    host: parseHosts(process.env.HOST),
    adminToken,
    amp: {
      maestroUrl,
      tenant: process.env.AMP_TENANT && process.env.AMP_TENANT.trim() !== '' ? process.env.AMP_TENANT.trim() : undefined,
    },
    bots,
    operatorAadObjectIds: parseOperatorAadRefs(process.env.OPERATOR_AAD_OBJECT_IDS),
    dryRunBootstrap: isTruthy(process.env.TEAMS_DRY_RUN),
    polling: {
      intervalMs: parsePositiveInt(process.env.POLL_INTERVAL_MS, DEFAULT_POLL_INTERVAL_MS, 'POLL_INTERVAL_MS'),
    },
    // Markdown is the default render; TEAMS_MARKDOWN=0/false flips to plain text.
    markdownDefault: process.env.TEAMS_MARKDOWN === undefined ? true : isTruthy(process.env.TEAMS_MARKDOWN),
    cacheUserTtlMs: parsePositiveInt(process.env.CACHE_USER_TTL_MS, DEFAULT_CACHE_USER_TTL_MS, 'CACHE_USER_TTL_MS'),
    threadStorePath,
    snapshotIntervalMs: parsePositiveInt(
      process.env.SNAPSHOT_INTERVAL_MS,
      DEFAULT_SNAPSHOT_INTERVAL_MS,
      'SNAPSHOT_INTERVAL_MS',
    ),
    debug: isTruthy(process.env.DEBUG),
  };
}
