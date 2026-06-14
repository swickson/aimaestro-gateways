import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { createThreadStore, type ThreadEntry } from '../thread-store.js';
import type { ThreadContext } from '../types.js';

function context(conversationId: string): ThreadContext {
  return {
    reference: {
      serviceUrl: 'https://smba.trafficmanager.net/amer/',
      channelId: 'msteams',
      conversation: { id: conversationId },
      bot: { id: 'bot-app-id', name: 'Maestro' },
      user: { id: 'user-1', name: 'Alice' },
    } as ThreadContext['reference'],
    rootActivityId: 'root-activity-1',
    tenantId: 'tenant-1',
  };
}

function entry(overrides: Partial<ThreadEntry> = {}): ThreadEntry {
  const conversationId = overrides.conversationId ?? 'conversation-maestro';
  return {
    botSlug: 'maestro',
    conversationId,
    ampMessageId: 'amp-1',
    aadObjectId: 'aad-user-1',
    context: context(conversationId),
    createdAt: 100,
    ...overrides,
  };
}

describe('thread-store by-user index (Phase 5)', () => {
  it('indexes a recorded entry by user and by (user, bot)', () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry());

    const byUser = store.findLatestByUser('aad-user-1');
    assert.equal(byUser?.ampMessageId, 'amp-1');
    assert.equal(byUser?.botSlug, 'maestro');

    const byUserBot = store.findByUserAndBot('aad-user-1', 'maestro');
    assert.equal(byUserBot?.ampMessageId, 'amp-1');

    assert.equal(store.findLatestByUser('nobody'), null);
    assert.equal(store.findByUserAndBot('aad-user-1', 'echo'), null);
  });

  it('findLatestByUser returns the most-recently-inbound bot (last-seen tiebreak)', () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    // Same user talks to two bots; echo is more recent (higher createdAt).
    store.record(entry({ botSlug: 'maestro', conversationId: 'c-maestro', ampMessageId: 'amp-m', createdAt: 100 }));
    store.record(entry({ botSlug: 'echo', conversationId: 'c-echo', ampMessageId: 'amp-e', createdAt: 200 }));

    assert.equal(store.findLatestByUser('aad-user-1')?.botSlug, 'echo');
    // Both per-(user,bot) entries remain independently resolvable.
    assert.equal(store.findByUserAndBot('aad-user-1', 'maestro')?.ampMessageId, 'amp-m');
    assert.equal(store.findByUserAndBot('aad-user-1', 'echo')?.ampMessageId, 'amp-e');

    // A newer maestro message flips the latest pointer back to maestro.
    store.record(entry({ botSlug: 'maestro', conversationId: 'c-maestro', ampMessageId: 'amp-m2', createdAt: 300 }));
    assert.equal(store.findLatestByUser('aad-user-1')?.botSlug, 'maestro');
    assert.equal(store.findLatestByUser('aad-user-1')?.ampMessageId, 'amp-m2');
  });

  it('does not index entries that lack an aadObjectId (graceful pre-Phase-5 entry)', () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry({ aadObjectId: undefined, ampMessageId: 'amp-legacy' }));

    // Still reply-deliverable via the amp-id index, but never DM-indexed.
    assert.equal(store.findByAmpMessageId('maestro', 'amp-legacy')?.ampMessageId, 'amp-legacy');
    assert.equal(store.findLatestByUser('aad-user-1'), null);
  });

  it('prunes the by-user pointer on count eviction', () => {
    const store = createThreadStore({ maxEntries: 1, maxAgeMs: Infinity });
    store.record(entry({ ampMessageId: 'amp-old', conversationId: 'c-old', createdAt: 100 }));
    // Second entry (different user) evicts the first (insertion-order, maxEntries=1).
    store.record(entry({ aadObjectId: 'aad-user-2', ampMessageId: 'amp-new', conversationId: 'c-new', createdAt: 200 }));

    assert.equal(store.findLatestByUser('aad-user-1'), null, 'evicted user pointer cleared');
    assert.equal(store.findByUserAndBot('aad-user-1', 'maestro'), null);
    assert.equal(store.findLatestByUser('aad-user-2')?.ampMessageId, 'amp-new');
  });

  it('lazily expires the by-user pointer past the age horizon', () => {
    let clock = 1000;
    const store = createThreadStore({ maxAgeMs: 500, now: () => clock });
    store.record(entry({ createdAt: 1000 }));

    assert.equal(store.findLatestByUser('aad-user-1')?.ampMessageId, 'amp-1');
    clock = 2000; // 1000ms later > 500ms horizon
    assert.equal(store.findLatestByUser('aad-user-1'), null, 'expired user pointer dropped on lookup');
    assert.equal(store.findByUserAndBot('aad-user-1', 'maestro'), null);
  });

  it('round-trips aadObjectId through snapshot/restore and rebuilds the by-user index', () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    store.record(entry({ botSlug: 'maestro', conversationId: 'c-m', ampMessageId: 'amp-m', createdAt: 100 }));
    store.record(entry({ botSlug: 'echo', conversationId: 'c-e', ampMessageId: 'amp-e', createdAt: 200 }));

    const snap = store.snapshot();
    assert.ok(snap.entries.every((e) => e.aadObjectId === 'aad-user-1'), 'aadObjectId persisted');

    const restored = createThreadStore({ maxAgeMs: Infinity });
    restored.restore(snap);
    assert.equal(restored.findLatestByUser('aad-user-1')?.botSlug, 'echo', 'by-user rebuilt with recency');
    assert.equal(restored.findByUserAndBot('aad-user-1', 'maestro')?.ampMessageId, 'amp-m');
    assert.equal(restored.findByUserAndBot('aad-user-1', 'echo')?.ampMessageId, 'amp-e');
  });
});
