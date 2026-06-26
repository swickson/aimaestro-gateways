/**
 * Teams gateway — trusted mesh-host origin allowlist (w3 outbound recalibration).
 *
 * The OUTBOUND consume path (outbound.ts) pulls cited attachment bytes from an
 * HMAC-signed `/api/v1/attachments/<id>/download` url. That url's origin is the
 * ORIGIN host's Tailscale address (`getSelfHost().url` on the Maestro that minted
 * it), NOT this gateway's own `maestroUrl` (which is typically `127.0.0.1`). Pinning
 * the consumer to `maestroUrl` therefore rejected every legit cross-/same-host
 * signed url. This loader derives the SSRF allowlist from the operator's trusted
 * mesh topology so a descriptor url is accepted iff its origin is a known mesh host.
 *
 * Source: `~/.aimaestro/hosts.json` (host-verified shape) —
 *   { "hosts": [ { id, name?, url, type?, enabled, aliases: string[] }, ... ] }
 * For each ENABLED host we take its canonical `url` PLUS every alias that PARSES as
 * a URL, normalized to its `origin` (scheme://host:port). Bare-hostname / bare-IP
 * aliases are deliberately SKIPPED — they have no scheme, so they are not origins,
 * and synthesizing `http://` for them would over-broaden the allowlist (SSRF). A
 * missing / unreadable / malformed file yields an EMPTY set (never throws): the
 * caller unions in its own `maestroUrl` origin and warns loudly that mesh-host
 * outbound attachments are disabled until the file is fixed + the gateway restarted.
 */

import { readFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';

interface RawHost {
  url?: unknown;
  enabled?: unknown;
  aliases?: unknown;
}

export type LoadMeshOriginsResult = { ok: boolean; origins?: Set<string> };

export interface LoadMeshOriginsOptions {
  /** Override the source path (defaults to ~/.aimaestro/hosts.json). */
  path?: string;
  /** Inject already-parsed JSON (tests) — bypasses the filesystem read entirely. */
  json?: unknown;
}

/** Default mesh-topology file location. */
export function defaultHostsPath(): string {
  return join(homedir(), '.aimaestro', 'hosts.json');
}

/** Normalize one candidate to its origin, or null if it is not a parseable URL. */
function toOrigin(candidate: unknown): string | null {
  if (typeof candidate !== 'string' || candidate.trim() === '') return null;
  try {
    return new URL(candidate).origin;
  } catch {
    return null; // bare hostname / IP / garbage — not an origin, skip.
  }
}

/** Pull the host array out of array / `{hosts:[...]}` / id-keyed-map tolerantly. */
function extractHosts(parsed: unknown): RawHost[] | null {
  if (Array.isArray(parsed)) return parsed as RawHost[];
  if (parsed && typeof parsed === 'object') {
    const obj = parsed as Record<string, unknown>;
    if (Array.isArray(obj.hosts)) return obj.hosts as RawHost[];
    // id-keyed map: { mesh-a: {...}, mesh-b: {...} } — cheap defensive tolerance.
    const values = Object.values(obj);
    if (values.every((v) => v && typeof v === 'object' && !Array.isArray(v))) {
      return values as RawHost[];
    }
  }
  return null;
}

/**
 * Build the trusted-origin allowlist from the mesh topology. Returns the set of
 * `scheme://host:port` origins for every ENABLED host. Never throws — a missing or
 * malformed source logs a `[MESH]` warning and returns an empty set.
 */
export function loadMeshOrigins(options: LoadMeshOriginsOptions = {}): LoadMeshOriginsResult {
  const origins = new Set<string>();

  let parsed: unknown;
  if (options.json !== undefined) {
    parsed = options.json;
  } else {
    const path = options.path ?? defaultHostsPath();
    let raw: string;
    try {
      raw = readFileSync(path, 'utf-8');
    } catch (err) {
      console.warn(`[MESH] could not read hosts file at ${path} — mesh-host origin allowlist is EMPTY: ${(err as Error).message}`);
      return { ok: false };
    }
    try {
      parsed = JSON.parse(raw);
    } catch (err) {
      console.warn(`[MESH] hosts file at ${path} is not valid JSON — mesh-host origin allowlist is EMPTY: ${(err as Error).message}`);
      return { ok: false };
    }
  }

  const hosts = extractHosts(parsed);
  if (hosts === null) {
    return { ok: false };
  }

  for (const host of hosts) {
    if (!host || typeof host !== 'object' || host.enabled !== true) continue;
    const candidates: unknown[] = [host.url, ...(Array.isArray(host.aliases) ? host.aliases : [])];
    for (const candidate of candidates) {
      const origin = toOrigin(candidate);
      if (origin) origins.add(origin);
    }
  }
  return { ok: true, origins };
}

/**
 * Mutable holder for the outbound-attachment origin allowlist. Ensures the gate
 * is live-resolved without restarting, while enforcing the security fail-safe:
 * a malformed reload NEVER replaces the live set with an empty/degraded one.
 */
export class OriginHolder {
  private currentOrigins: ReadonlySet<string>;

  constructor(private maestroUrl: string, initialMesh: Set<string>) {
    this.currentOrigins = this.buildOrigins(initialMesh);
  }

  private buildOrigins(mesh: Set<string>): ReadonlySet<string> {
    return new Set([new URL(this.maestroUrl).origin, ...mesh]);
  }

  /** Gets the live set of allowed origins (maestroUrl + enabled mesh hosts). */
  get current(): ReadonlySet<string> {
    return this.currentOrigins;
  }

  /**
   * Reload the mesh origins. If the file is missing/malformed (empty set),
   * the reload is ABORTED and the previous valid allowlist is retained.
   */
  reload(options?: LoadMeshOriginsOptions): void {
    const result = loadMeshOrigins(options);
    if (!result.ok || !result.origins) {
      console.warn(
        '[MESH] ⚠️  Reload triggered but source is missing or malformed. ' +
          'RETAINING previous allowlist to prevent silent outbound attachment drops.'
      );
      return;
    }
    this.currentOrigins = this.buildOrigins(result.origins);
    if (result.origins.size === 0) {
      console.log('[MESH] reloaded trusted outbound attachment origins: <none> (all mesh hosts disabled)');
    } else {
      console.log(`[MESH] reloaded trusted outbound attachment origins: ${[...result.origins].join(', ')}`);
    }
  }
}
