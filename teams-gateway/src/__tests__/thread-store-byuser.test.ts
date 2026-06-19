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
    store.record(entry({ ampMessageId: 'amp-old', conversationId: 'c-old', context: context('c-old'), createdAt: 100 }));
    // Second entry (different user) evicts the first (insertion-order, maxEntries=1).
    store.record(entry({ aadObjectId: 'aad-user-2', ampMessageId: 'amp-new', conversationId: 'c-new', context: context('c-new'), createdAt: 200 }));

    assert.equal(store.findByAmpMessageId('maestro', 'amp-old'), null);
    assert.equal(store.findRecentByConversation('maestro', 'c-old'), null);
    assert.equal(store.findLatestByUser('aad-user-1'), null, 'evicted user pointer cleared');
    assert.equal(store.findByUserAndBot('aad-user-1', 'maestro'), null);
    assert.equal(store.findLatestByUser('aad-user-2')?.ampMessageId, 'amp-new');
    assert.equal(store.findByAmpMessageId('maestro', 'amp-new')?.aadObjectId, 'aad-user-2');
  });

  it('#26 keeps by-user mappings past the 24h reply-recency horizon by default', () => {
    let clock = 1000;
    const store = createThreadStore({ maxAgeMs: 500, now: () => clock });
    store.record(entry({ createdAt: 1000 }));

    assert.equal(store.findLatestByUser('aad-user-1')?.ampMessageId, 'amp-1');
    clock = 2000; // 1000ms later > 500ms horizon
    assert.equal(store.findByAmpMessageId('maestro', 'amp-1'), null, 'reply-path amp id expires at the reply horizon');
    assert.equal(store.findRecentByConversation('maestro', 'conversation-maestro'), null, 'reply-path conversation recency expires');
    assert.equal(store.findLatestByUser('aad-user-1')?.ampMessageId, 'amp-1', 'DM-target mapping remains durable');
    assert.equal(store.findByUserAndBot('aad-user-1', 'maestro')?.ampMessageId, 'amp-1');
  });

  it('lazily expires the by-user pointer only past an explicit user horizon', () => {
    let clock = 1000;
    const store = createThreadStore({ maxAgeMs: Infinity, userMaxAgeMs: 500, now: () => clock });
    store.record(entry({ createdAt: 1000 }));

    assert.equal(store.findLatestByUser('aad-user-1')?.ampMessageId, 'amp-1');
    clock = 2000;
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

  it('#26 snapshot/restore preserves DM-target mappings after reply entries age out', () => {
    let clock = 0;
    const store = createThreadStore({ maxAgeMs: 500, now: () => clock });
    store.record(entry({ ampMessageId: 'amp-old', conversationId: 'c-old', context: context('c-old'), createdAt: 0 }));

    clock = 1000;
    const snap = store.snapshot();
    assert.equal(snap.entries.length, 1, 'snapshot includes entries live under the user horizon');

    const restored = createThreadStore({ maxAgeMs: 500, now: () => clock });
    restored.restore(snap);
    assert.equal(restored.findByAmpMessageId('maestro', 'amp-old'), null);
    assert.equal(restored.findRecentByConversation('maestro', 'c-old'), null);
    assert.equal(restored.findByUserAndBot('aad-user-1', 'maestro')?.conversationId, 'c-old');
    assert.equal(restored.findLatestByUser('aad-user-1')?.conversationId, 'c-old');
  });

  it('#26 split-drop: reply-path expiry does not drop the by-user mapping', () => {
    let clock = 0;
    const store = createThreadStore({ maxAgeMs: 500, now: () => clock });
    store.record(entry({ ampMessageId: 'amp-old', conversationId: 'c-old', context: context('c-old'), createdAt: 0 }));

    clock = 1000;
    assert.equal(store.findRecentByConversation('maestro', 'c-old'), null);
    assert.equal(store.findByAmpMessageId('maestro', 'amp-old'), null);
    assert.equal(store.findByUserAndBot('aad-user-1', 'maestro')?.conversationId, 'c-old');
    assert.equal(store.findLatestByUser('aad-user-1')?.conversationId, 'c-old');
  });

  it('#26 restore preserves chronological eviction order when DM-only entries snapshot after active replies', () => {
    let clock = 0;
    const source = createThreadStore({ maxAgeMs: 500, now: () => clock });

    source.record(entry({
      aadObjectId: 'aad-old-1',
      ampMessageId: 'amp-old-1',
      conversationId: 'c-old-1',
      context: context('c-old-1'),
      createdAt: 0,
    }));
    clock = 100;
    source.record(entry({
      aadObjectId: 'aad-old-2',
      ampMessageId: 'amp-old-2',
      conversationId: 'c-old-2',
      context: context('c-old-2'),
      createdAt: 100,
    }));
    clock = 1000;
    source.record(entry({
      aadObjectId: 'aad-new-1',
      ampMessageId: 'amp-new-1',
      conversationId: 'c-new-1',
      context: context('c-new-1'),
      createdAt: 1000,
    }));
    clock = 1100;
    source.record(entry({
      aadObjectId: 'aad-new-2',
      ampMessageId: 'amp-new-2',
      conversationId: 'c-new-2',
      context: context('c-new-2'),
      createdAt: 1100,
    }));

    clock = 1200;
    assert.equal(source.findByAmpMessageId('maestro', 'amp-old-1'), null);
    assert.equal(source.findByAmpMessageId('maestro', 'amp-old-2'), null);
    assert.equal(source.findByUserAndBot('aad-old-1', 'maestro')?.ampMessageId, 'amp-old-1');
    assert.equal(source.findByUserAndBot('aad-old-2', 'maestro')?.ampMessageId, 'amp-old-2');

    const snap = source.snapshot();
    assert.deepEqual(snap.entries.map((e) => e.ampMessageId), [
      'amp-new-1',
      'amp-new-2',
      'amp-old-1',
      'amp-old-2',
    ], 'snapshot order reproduces the pre-fix hazard: active replies before older DM-only entries');

    const restored = createThreadStore({ maxEntries: 4, maxAgeMs: 500, now: () => clock });
    restored.restore(snap);
    restored.record(entry({
      aadObjectId: 'aad-extra',
      ampMessageId: 'amp-extra',
      conversationId: 'c-extra',
      context: context('c-extra'),
      createdAt: 1200,
    }));

    assert.equal(restored.findByUserAndBot('aad-old-1', 'maestro'), null, 'oldest-by-recordedAt entry evicted first');
    assert.equal(restored.findByUserAndBot('aad-old-2', 'maestro')?.ampMessageId, 'amp-old-2', 'second-oldest DM mapping survives');
    assert.equal(restored.findByAmpMessageId('maestro', 'amp-new-1')?.ampMessageId, 'amp-new-1', 'newer reply-active entry was not evicted first');
    assert.equal(restored.findByAmpMessageId('maestro', 'amp-new-2')?.ampMessageId, 'amp-new-2');
    assert.equal(restored.findByAmpMessageId('maestro', 'amp-extra')?.ampMessageId, 'amp-extra');
    assert.equal(restored.size(), 4);
  });
});
