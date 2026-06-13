import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath } from 'node:url';

import { createAuthMiddleware } from '../auth.js';

function response() {
  return {
    statusCode: 200,
    body: undefined as unknown,
    status(code: number) {
      this.statusCode = code;
      return this;
    },
    json(payload: unknown) {
      this.body = payload;
      return this;
    },
  };
}

describe('createAuthMiddleware', () => {
  it('throws at init for empty or blank tokens', () => {
    assert.throws(() => createAuthMiddleware(''), /ADMIN_TOKEN is required/);
    assert.throws(() => createAuthMiddleware('   '), /ADMIN_TOKEN is required/);
    assert.throws(() => createAuthMiddleware(undefined as unknown as string), /ADMIN_TOKEN is required/);
  });

  it('passes valid bearer tokens and rejects invalid, missing, and length-mismatched tokens', () => {
    const middleware = createAuthMiddleware('secret-token');
    let nextCalled = false;
    const okRes = response();
    middleware({ headers: { authorization: 'Bearer secret-token' } } as never, okRes as never, () => {
      nextCalled = true;
    });
    assert.equal(nextCalled, true);
    assert.equal(okRes.statusCode, 200);

    for (const authorization of ['Bearer wrong-token', 'Bearer secret-token-extra', undefined]) {
      const res = response();
      nextCalled = false;
      assert.doesNotThrow(() => middleware({ headers: { authorization } } as never, res as never, () => {
        nextCalled = true;
      }));
      assert.equal(nextCalled, false);
      assert.equal(res.statusCode, 401);
      assert.deepEqual(res.body, { error: 'Unauthorized' });
    }
  });

  it('rejects multibyte invalid tokens with matching character length without throwing', () => {
    const middleware = createAuthMiddleware('secret-token');
    const authorization = 'Bearer secret-tokeé';
    assert.equal(authorization.length, 'Bearer secret-token'.length);
    assert.notEqual(Buffer.from(authorization).length, Buffer.from('Bearer secret-token').length);

    const res = response();
    let nextCalled = false;
    assert.doesNotThrow(() => middleware({ headers: { authorization } } as never, res as never, () => {
      nextCalled = true;
    }));
    assert.equal(nextCalled, false);
    assert.equal(res.statusCode, 401);
    assert.deepEqual(res.body, { error: 'Unauthorized' });
  });

  it('uses timingSafeEqual for the secret comparison', () => {
    const dir = dirname(fileURLToPath(import.meta.url));
    const source = readFileSync(join(dir, '..', 'auth.ts'), 'utf-8');
    assert.match(source, /timingSafeEqual/);
    assert.doesNotMatch(source, /auth\s*={2,3}\s*expected/);
    assert.doesNotMatch(source, /expected\s*={2,3}\s*auth/);
  });
});
