/**
 * w3 outbound recalibration — validateOutboundDescriptor (outbound.ts) matrix.
 *
 * Direct unit coverage of the consume-boundary validator after the Watson-locked
 * recalibration: KIND (reject only `legacy`, infer amp-v1 by shape), ORIGIN (accept
 * any trusted mesh-host origin in the allowlist, REJECT arbitrary external = SSRF
 * closed), DIGEST (tolerate an optional `sha256:` prefix), while the path-pin +
 * id-match + size/scan_status/deny-list gates stay intact.
 */

import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { validateOutboundDescriptor } from '../outbound.js';
import type { AttachmentPolicy } from '../types.js';

const POLICY: AttachmentPolicy = { maxBytes: 26_214_400, maxCount: 10, denyContentTypes: ['application/x-msdownload'] };

// Trusted mesh allowlist: the gateway's own maestroUrl origin + a remote mesh host.
const ALLOWED = new Set(['http://127.0.0.1:23000', 'http://100.83.160.34:23000']);

const HEX = 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855';

/** A fully valid WIRE-form descriptor (CLI omits `kind`), on the local mesh origin. */
function descriptor(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    id: 'att-9',
    filename: 'photo.png',
    content_type: 'image/png',
    size: 1234,
    digest: HEX,
    scan_status: 'basic_clean',
    url: 'http://127.0.0.1:23000/api/v1/attachments/att-9/download?sig=abc',
    ...overrides,
  };
}

const ok = (att: unknown) => validateOutboundDescriptor(att, ALLOWED, POLICY);

describe('validateOutboundDescriptor — recalibrated consume boundary', () => {
  it('ACCEPTS a valid wire descriptor with NO kind field (CLI omits kind)', () => {
    assert.equal(ok(descriptor()), null);
  });

  it('ACCEPTS an explicit kind: amp-v1 (inbound-assembled form)', () => {
    assert.equal(ok(descriptor({ kind: 'amp-v1' })), null);
  });

  it("REJECTS kind: 'legacy' (the only hard-rejected kind)", () => {
    assert.match(String(ok(descriptor({ kind: 'legacy' }))), /legacy/);
  });

  it('ACCEPTS a url on a remote mesh-host origin in the allowlist', () => {
    assert.equal(ok(descriptor({
      url: 'http://100.83.160.34:23000/api/v1/attachments/att-9/download?sig=x',
    })), null);
  });

  it('REJECTS an arbitrary external origin (SSRF closed)', () => {
    const reason = ok(descriptor({ url: 'https://evil.test/api/v1/attachments/att-9/download' }));
    assert.match(String(reason), /not a trusted mesh host/);
  });

  it('REJECTS a same-origin url whose path is not /attachments/<id>/download (internal-route SSRF)', () => {
    assert.match(String(ok(descriptor({ url: 'http://127.0.0.1:23000/api/v1/agents' }))), /path/);
  });

  it('REJECTS a /download url whose <id> does not match the descriptor id', () => {
    assert.match(String(ok(descriptor({ url: 'http://127.0.0.1:23000/api/v1/attachments/OTHER/download' }))), /path/);
  });

  it('ACCEPTS a sha256:-prefixed digest', () => {
    assert.equal(ok(descriptor({ digest: `sha256:${HEX}` })), null);
  });

  it('ACCEPTS a bare-hex digest', () => {
    assert.equal(ok(descriptor({ digest: HEX })), null);
  });

  it('ACCEPTS an uppercase SHA256: prefix (case-insensitive)', () => {
    assert.equal(ok(descriptor({ digest: `SHA256:${HEX.toUpperCase()}` })), null);
  });

  it('REJECTS a missing digest', () => {
    assert.match(String(ok(descriptor({ digest: undefined }))), /digest/);
  });

  it('REJECTS a non-hex / wrong-length digest', () => {
    assert.match(String(ok(descriptor({ digest: 'd' }))), /digest/);
    assert.match(String(ok(descriptor({ digest: 'sha256:nothex' }))), /digest/);
  });

  it('still enforces scan_status, size cap, deny-list, and the never-throw hostile-shape guard', () => {
    assert.match(String(ok(descriptor({ scan_status: 'pending' }))), /scan_status/);
    assert.match(String(ok(descriptor({ size: POLICY.maxBytes + 1 }))), /exceeds cap/);
    assert.match(String(ok(descriptor({ content_type: 'application/x-msdownload' }))), /deny-listed/);
    assert.match(String(ok(descriptor({ filename: '' }))), /filename/);
    // hostile non-object shapes never throw — always a drop reason string
    assert.match(String(ok(null)), /not an object/);
    assert.match(String(ok('a string')), /not an object/);
    assert.match(String(ok([descriptor()])), /not an object/);
  });
});
