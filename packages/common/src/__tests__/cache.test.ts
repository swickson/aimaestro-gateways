import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { Cache } from '../cache.js';

const wait = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

describe('Cache', () => {
  it('evicts by TTL, size limit, cleanup, delete, clear, and periodic cleanup', async () => {
    const ttl = new Cache<string>(20, 2);
    ttl.set('key', 'value');
    assert.equal(ttl.get('key'), 'value');
    await wait(30);
    assert.equal(ttl.get('key'), null);

    const bounded = new Cache<number>(1000, 2);
    bounded.set('a', 1);
    bounded.set('b', 2);
    bounded.set('c', 3);
    assert.equal(bounded.get('a'), null);
    assert.equal(bounded.get('b'), 2);
    assert.equal(bounded.get('c'), 3);

    const cleanup = new Cache<string>(20);
    cleanup.set('old', 'old');
    await wait(30);
    cleanup.set('fresh', 'fresh');
    cleanup.cleanup();
    assert.equal(cleanup.get('old'), null);
    assert.equal(cleanup.get('fresh'), 'fresh');

    cleanup.delete('fresh');
    assert.equal(cleanup.get('fresh'), null);
    cleanup.set('clear', 'clear');
    cleanup.clear();
    assert.equal(cleanup.size(), 0);

    cleanup.set('periodic', 'value');
    cleanup.startCleanup(5);
    await wait(30);
    cleanup.stopCleanup();
    assert.equal(cleanup.get('periodic'), null);
  });
});
