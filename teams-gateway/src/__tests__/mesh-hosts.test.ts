/**
 * w3 outbound recalibration — mesh-host origin allowlist loader (mesh-hosts.ts).
 *
 * Verifies loadMeshOrigins against the host-verified ~/.aimaestro/hosts.json shape:
 * enabled hosts contribute their `url` origin plus every URL-PARSEABLE alias origin;
 * disabled hosts are excluded; bare-hostname / bare-IP aliases are skipped (not
 * synthesized into origins — that would over-broaden the SSRF allowlist); and a
 * missing / malformed source returns an EMPTY set without throwing.
 */

import assert from 'node:assert/strict';
import { afterEach, beforeEach, describe, it } from 'node:test';

import { loadMeshOrigins, OriginHolder } from '../mesh-hosts.js';

const origWarn = console.warn;
const origLog = console.log;
beforeEach(() => {
  console.warn = () => undefined;
  console.log = () => undefined;
});
afterEach(() => {
  console.warn = origWarn;
  console.log = origLog;
});

// The real host-verified shape (sanitized).
const REAL_SHAPE = {
  hosts: [
    {
      id: 'holmes', name: 'holmes', url: 'http://203.0.113.11:23000', type: null, enabled: true,
      aliases: ['holmes', '10.10.40.59', '203.0.113.11', 'http://10.10.40.59:23000', 'http://203.0.113.11:23000'],
    },
    { id: 'redacted-1', url: 'http://203.0.113.13:23000', type: 'remote', enabled: true, aliases: [] },
    { id: 'host-a-mbp', name: 'redacted-3', url: 'http://203.0.113.12:23000', type: 'remote', enabled: true, aliases: ['host-a-mbp', 'host-b.internal'] },
  ],
};

describe('loadMeshOrigins', () => {
  it('collects url + URL-parseable alias origins for every enabled host; skips bare hostnames/IPs', () => {
    const origins = loadMeshOrigins({ json: REAL_SHAPE });

    // canonical urls
    assert.ok(origins.has('http://203.0.113.11:23000'));
    assert.ok(origins.has('http://203.0.113.13:23000'));
    assert.ok(origins.has('http://203.0.113.12:23000'));
    // the http:// alias contributes the LAN-IP origin variant
    assert.ok(origins.has('http://10.10.40.59:23000'));
    // bare hostname / bare IP aliases are NOT turned into origins
    assert.ok(!origins.has('holmes'));
    assert.ok(!origins.has('10.10.40.59'));
    assert.ok(!origins.has('203.0.113.11'));
    assert.ok(!origins.has('host-b.internal'));
    // exactly the 4 real origins, no synthesized extras
    assert.equal(origins.size, 4);
  });

  it('excludes disabled hosts entirely', () => {
    const origins = loadMeshOrigins({
      json: { hosts: [
        { id: 'on', url: 'http://100.0.0.1:23000', enabled: true, aliases: [] },
        { id: 'off', url: 'http://100.0.0.2:23000', enabled: false, aliases: ['http://100.0.0.2:23000'] },
        { id: 'missing-flag', url: 'http://100.0.0.3:23000', aliases: [] },
      ] },
    });
    assert.deepEqual([...origins], ['http://100.0.0.1:23000']);
  });

  it('tolerates a bare host array (no { hosts } wrapper)', () => {
    const origins = loadMeshOrigins({ json: [{ id: 'a', url: 'http://100.5.5.5:23000', enabled: true, aliases: [] }] });
    assert.ok(origins.has('http://100.5.5.5:23000'));
    assert.equal(origins.size, 1);
  });

  it('returns an EMPTY set (no throw) for a missing file', () => {
    const origins = loadMeshOrigins({ path: '/nonexistent/definitely/not/here/hosts.json' });
    assert.equal(origins.size, 0);
  });

  it('returns an EMPTY set (no throw) for malformed JSON / non-host shapes', () => {
    assert.equal(loadMeshOrigins({ json: 'not an object' }).size, 0);
    assert.equal(loadMeshOrigins({ json: null }).size, 0);
    assert.equal(loadMeshOrigins({ json: { hosts: 'nope' } }).size, 0);
  });

  it('skips a host whose url itself is not a parseable origin but keeps its valid aliases', () => {
    const origins = loadMeshOrigins({
      json: { hosts: [{ id: 'x', url: 'not-a-url', enabled: true, aliases: ['http://100.9.9.9:23000'] }] },
    });
    assert.deepEqual([...origins], ['http://100.9.9.9:23000']);
  });
});

describe('OriginHolder', () => {
  const MAESTRO_URL = 'http://127.0.0.1:23000';

  it('initializes with maestroUrl plus initial mesh origins', () => {
    const initialMesh = new Set(['http://10.10.10.1:23000']);
    const holder = new OriginHolder(MAESTRO_URL, initialMesh);
    assert.equal(holder.current.size, 2);
    assert.ok(holder.current.has(MAESTRO_URL));
    assert.ok(holder.current.has('http://10.10.10.1:23000'));
  });

  it('reloads and updates origins on success, keeping maestroUrl', () => {
    const holder = new OriginHolder(MAESTRO_URL, new Set(['http://old:23000']));
    assert.ok(holder.current.has('http://old:23000'));
    
    // reload with a valid shape
    holder.reload({ json: { hosts: [{ id: 'new', url: 'http://new:23000', enabled: true, aliases: [] }] } });
    
    assert.equal(holder.current.size, 2);
    assert.ok(holder.current.has(MAESTRO_URL));
    assert.ok(holder.current.has('http://new:23000'));
    assert.ok(!holder.current.has('http://old:23000')); // old origin removed
  });

  it('FAIL-SAFE: retains previous valid set and warns on empty/malformed source', () => {
    const holder = new OriginHolder(MAESTRO_URL, new Set(['http://good:23000']));
    const originalSet = holder.current;
    assert.equal(originalSet.size, 2);
    
    let warned = false;
    const localOrigWarn = console.warn;
    console.warn = () => { warned = true; };
    try {
      holder.reload({ json: 'malformed' });
    } finally {
      console.warn = localOrigWarn;
    }
    
    // The reference must be identically preserved (never swapped)
    assert.strictEqual(holder.current, originalSet);
    assert.ok(warned, 'Expected fail-safe warning to be logged');
    assert.ok(holder.current.has('http://good:23000'));
  });
});
