/**
 * FRESH cross-review tests for w4 #13 (proactive DM cold-start).
 *
 * Authored by Crease (senior-eng correctness lens) BLIND to Whistler's own
 * dm-endpoint.test.ts — written from the dispatch spec, not the implementation:
 *   (a) flag OFF (default) preserves the pinned no_prior_contact 409 contract
 *       AND never invokes createConversation;
 *   (b) createConversation sends the FIRST chunk itself — no duplicate App.send;
 *   (c) the cold-start ConversationReference is persisted so the NEXT DM reuses
 *       it and does NOT create again (both pinned-bot and last-seen paths);
 *   (d) Bot Connector failure mapping is exhaustive + never silent, and a failed
 *       cold-start records NOTHING (a retry re-attempts rather than reusing a
 *       phantom conversation);
 *   (e) the scope bound is enforced — cold-start requires botSlug + tenantId;
 *       truly-cold (no identity) cannot create.
 *
 * Self-contained fixtures (no import of the author's helpers).
 */

import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { deliverDm, type DmDeps, type CreateColdStartConversationInput, type CreateColdStartConversationResult } from '../dm.js';
import { createThreadStore, type ThreadStore } from '../thread-store.js';
import type { ThreadContext } from '../types.js';

function reference(conversationId: string): ThreadContext['reference'] {
  return {
    serviceUrl: 'https://smba.trafficmanager.net/amer/',
    channelId: 'msteams',
    conversation: { id: conversationId, conversationType: 'personal' },
    bot: { id: 'bot-app-id', name: 'maestro' },
    user: { id: 'aad-77', name: 'aad-77' },
  } as ThreadContext['reference'];
}

interface Sent { botSlug: string; conversationId: string; text: string; markdown: boolean }

/** A DmDeps with cold-start ON and a capturing create fn. */
function coldDeps(
  store: ThreadStore,
  sent: Sent[],
  created: CreateColdStartConversationInput[],
  opts: {
    coldStartEnabled?: boolean;
    markdownDefault?: boolean;
    createResult?: (input: CreateColdStartConversationInput) => CreateColdStartConversationResult;
    createThrows?: () => never;
    omitCreateFn?: boolean;
  } = {},
): DmDeps {
  const base: DmDeps = {
    threadStore: store,
    knownBots: new Set(['maestro', 'echo']),
    coldStartEnabled: opts.coldStartEnabled ?? true,
    markdownDefault: opts.markdownDefault ?? true,
    sendChunk: async (botSlug, conversationId, text, markdown) => {
      sent.push({ botSlug, conversationId, text, markdown });
    },
  };
  if (!opts.omitCreateFn) {
    base.createColdStartConversation = async (input) => {
      created.push(input);
      if (opts.createThrows) opts.createThrows();
      const r = opts.createResult?.(input);
      return (
        r ?? {
          conversationId: `cold-${input.aadObjectId}`,
          rootActivityId: `act-${input.aadObjectId}`,
          reference: reference(`cold-${input.aadObjectId}`),
        }
      );
    };
  }
  return base;
}

// ---------------------------------------------------------------------------
// (a) Flag OFF — preserves the 409 contract, never creates.
// ---------------------------------------------------------------------------
describe('#13 (a) cold-start OFF preserves the no_prior_contact 409', () => {
  it('returns the exact contract triple and never invokes createConversation', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    const sent: Sent[] = [];
    const created: CreateColdStartConversationInput[] = [];
    // create fn IS configured — proving the flag (not the missing dep) gates it.
    const deps = coldDeps(store, sent, created, { coldStartEnabled: false });

    const r = await deliverDm(deps, { platformUserId: 'aad-77', botSlug: 'maestro', tenantId: 't-1', message: 'hi' });

    assert.equal(r.status, 409);
    assert.equal(r.json.error, 'undeliverable');
    assert.equal(r.json.reason, 'no_prior_contact');
    assert.equal(created.length, 0, 'flag OFF must not create a conversation');
    assert.equal(sent.length, 0);
    assert.equal(store.size(), 0, 'flag OFF records nothing');
  });
});

// ---------------------------------------------------------------------------
// (b) First chunk via createConversation; no duplicate App.send.
// ---------------------------------------------------------------------------
describe('#13 (b) createConversation owns the first chunk', () => {
  it('single chunk: create gets it, sendChunk is never called', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    const sent: Sent[] = [];
    const created: CreateColdStartConversationInput[] = [];
    const r = await deliverDm(coldDeps(store, sent, created), {
      platformUserId: 'aad-77',
      botSlug: 'maestro',
      tenantId: 't-1',
      message: 'only one',
    });
    assert.equal(r.status, 200);
    assert.equal(r.json.coldStart, true);
    assert.equal(r.json.chunks, 1);
    assert.equal(created.length, 1);
    assert.equal(created[0]?.text, 'only one');
    assert.equal(sent.length, 0, 'no App.send for a single-chunk cold-start');
  });

  it('multi chunk: create gets ONLY chunk[0]; the rest go via sendChunk; chunk[0] is never re-sent', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    const sent: Sent[] = [];
    const created: CreateColdStartConversationInput[] = [];
    const big = 'x'.repeat(30_000); // > 28KB → 2 chunks
    const r = await deliverDm(coldDeps(store, sent, created), {
      platformUserId: 'aad-77',
      botSlug: 'maestro',
      tenantId: 't-1',
      message: big,
    });
    assert.equal(r.status, 200);
    assert.equal(r.json.chunks, 2);
    assert.equal(created.length, 1);
    assert.equal(sent.length, 1, 'exactly chunks-1 App.send calls');
    const firstChunk = created[0]?.text ?? '';
    // The single sendChunk payload must be the SECOND chunk, never a repeat of the first.
    assert.notEqual(sent[0]?.text, firstChunk, 'chunk[0] must not be re-sent via App.send');
    // Total bytes preserved across the two legs (no loss, no dup).
    assert.equal((firstChunk + sent[0]?.text).length, big.length);
  });

  it('subject is prepended to the first (creation) chunk, bold in markdown mode', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    const sent: Sent[] = [];
    const created: CreateColdStartConversationInput[] = [];
    await deliverDm(coldDeps(store, sent, created), {
      platformUserId: 'aad-77',
      botSlug: 'maestro',
      tenantId: 't-1',
      message: 'body',
      subject: 'Heads up',
    });
    assert.equal(created[0]?.text, '**Heads up**\n\nbody');
  });
});

// ---------------------------------------------------------------------------
// (c) Persistence — the NEXT DM reuses, does not create again.
// ---------------------------------------------------------------------------
describe('#13 (c) cold-start persists the reference for reuse', () => {
  it('pinned-bot second DM reuses the recorded conversation (no second create)', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    const sent: Sent[] = [];
    const created: CreateColdStartConversationInput[] = [];
    const deps = coldDeps(store, sent, created);

    await deliverDm(deps, { platformUserId: 'aad-77', botSlug: 'maestro', tenantId: 't-1', message: 'first' });
    assert.equal(created.length, 1);

    const recorded = store.findByUserAndBot('aad-77', 'maestro');
    assert.equal(recorded?.conversationId, 'cold-aad-77');
    assert.equal(recorded?.aadObjectId, 'aad-77');
    assert.equal(recorded?.context.tenantId, 't-1');

    const r2 = await deliverDm(deps, { platformUserId: 'aad-77', botSlug: 'maestro', message: 'second' });
    assert.equal(r2.status, 200);
    assert.equal(created.length, 1, 'second DM must NOT create again');
    assert.equal(sent.length, 1);
    assert.equal(sent[0]?.conversationId, 'cold-aad-77');
    assert.equal(sent[0]?.text, 'second');
    assert.equal(r2.json.coldStart, undefined, 'warm reuse is not flagged coldStart');
  });

  it('no-botSlug second DM resolves via findLatestByUser (last-seen bot) and reuses', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    const sent: Sent[] = [];
    const created: CreateColdStartConversationInput[] = [];
    const deps = coldDeps(store, sent, created);

    await deliverDm(deps, { platformUserId: 'aad-77', botSlug: 'maestro', tenantId: 't-1', message: 'first' });
    const r2 = await deliverDm(deps, { platformUserId: 'aad-77', message: 'no-bot-second' });
    assert.equal(r2.status, 200);
    assert.equal(r2.json.botSlug, 'maestro');
    assert.equal(created.length, 1);
    assert.equal(sent[0]?.conversationId, 'cold-aad-77');
  });
});

// ---------------------------------------------------------------------------
// (d) Failure mapping — exhaustive, never silent, records nothing on failure.
// ---------------------------------------------------------------------------
describe('#13 (d) Bot Connector failure mapping', () => {
  const cases: Array<{ name: string; err: unknown; status: number; reason: string }> = [
    { name: '401 → 409 forbidden', err: Object.assign(new Error('x'), { status: 401 }), status: 409, reason: 'bot_not_installed_or_forbidden' },
    { name: '403 → 409 forbidden', err: Object.assign(new Error('x'), { status: 403 }), status: 409, reason: 'bot_not_installed_or_forbidden' },
    { name: '404 → 409 not found', err: Object.assign(new Error('x'), { statusCode: 404 }), status: 409, reason: 'user_or_tenant_not_found' },
    { name: '400 → 409 wrong tenant', err: Object.assign(new Error('x'), { response: { status: 400 } }), status: 409, reason: 'wrong_tenant_or_unreachable' },
    { name: 'code wrong_tenant → 409', err: Object.assign(new Error('x'), { code: 'wrong_tenant' }), status: 409, reason: 'wrong_tenant_or_unreachable' },
    { name: '500 → 502 cold_start_failed', err: Object.assign(new Error('x'), { status: 500 }), status: 502, reason: 'cold_start_failed' },
    { name: 'opaque error → 502', err: new Error('boom'), status: 502, reason: 'cold_start_failed' },
  ];

  for (const c of cases) {
    it(`${c.name}; records nothing, sends nothing`, async () => {
      const store = createThreadStore({ maxAgeMs: Infinity });
      const sent: Sent[] = [];
      const created: CreateColdStartConversationInput[] = [];
      const deps = coldDeps(store, sent, created, { createThrows: () => { throw c.err; } });
      const r = await deliverDm(deps, { platformUserId: 'aad-77', botSlug: 'maestro', tenantId: 't-1', message: 'hi' });
      assert.equal(r.status, c.status);
      assert.equal(r.json.reason, c.reason);
      assert.equal(created.length, 1, 'create was attempted');
      assert.equal(sent.length, 0);
      assert.equal(store.size(), 0, 'a FAILED cold-start must persist nothing (retry re-attempts, no phantom reuse)');
    });
  }

  it('cold-start enabled but no create dependency wired → 502 cold_start_unavailable (not silent)', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    const sent: Sent[] = [];
    const created: CreateColdStartConversationInput[] = [];
    const deps = coldDeps(store, sent, created, { omitCreateFn: true });
    const r = await deliverDm(deps, { platformUserId: 'aad-77', botSlug: 'maestro', tenantId: 't-1', message: 'hi' });
    assert.equal(r.status, 502);
    assert.equal(r.json.reason, 'cold_start_unavailable');
    assert.equal(store.size(), 0);
  });
});

// ---------------------------------------------------------------------------
// (e) Scope bound — cold-start needs botSlug + tenantId; truly-cold cannot create.
// ---------------------------------------------------------------------------
describe('#13 (e) scope bound enforcement', () => {
  it('no botSlug → 409 no_send_bot (cannot pick a bot to create under)', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    const sent: Sent[] = [];
    const created: CreateColdStartConversationInput[] = [];
    const r = await deliverDm(coldDeps(store, sent, created), { platformUserId: 'aad-77', tenantId: 't-1', message: 'hi' });
    assert.equal(r.status, 409);
    assert.equal(r.json.reason, 'no_send_bot');
    assert.equal(created.length, 0);
    assert.equal(store.size(), 0);
  });

  it('botSlug present but no tenantId → 400 tenant_required (cannot create cross-tenant)', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    const sent: Sent[] = [];
    const created: CreateColdStartConversationInput[] = [];
    const r = await deliverDm(coldDeps(store, sent, created), { platformUserId: 'aad-77', botSlug: 'maestro', message: 'hi' });
    assert.equal(r.status, 400);
    assert.equal(r.json.error, 'bad_request');
    assert.equal(created.length, 0);
    assert.equal(store.size(), 0);
  });

  it('unregistered pinned botSlug is rejected 400 BEFORE any cold-start path', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    const sent: Sent[] = [];
    const created: CreateColdStartConversationInput[] = [];
    const r = await deliverDm(coldDeps(store, sent, created), { platformUserId: 'aad-77', botSlug: 'ghost', tenantId: 't-1', message: 'hi' });
    assert.equal(r.status, 400);
    assert.equal(created.length, 0);
  });
});

// ---------------------------------------------------------------------------
// Edge case documented for Holmes: multi-chunk cold-start, first chunk created +
// RECORDED, then a later App.send fails. The conversation is already persisted,
// so the mapped error is returned BUT the next retry will reuse (warm path) and
// re-send the already-delivered first chunk → potential duplicate. Captured here
// as observed behavior, not asserted-correct.
// ---------------------------------------------------------------------------
describe('#13 edge: partial-failure after a successful create', () => {
  it('records the conversation even though a later chunk send fails (retry would duplicate chunk[0])', async () => {
    const store = createThreadStore({ maxAgeMs: Infinity });
    const created: CreateColdStartConversationInput[] = [];
    let calls = 0;
    const deps: DmDeps = {
      threadStore: store,
      knownBots: new Set(['maestro']),
      coldStartEnabled: true,
      markdownDefault: true,
      sendChunk: async () => { calls += 1; throw Object.assign(new Error('net'), { status: 503 }); },
      createColdStartConversation: async (input) => {
        created.push(input);
        return { conversationId: 'cold-x', rootActivityId: 'act-x', reference: reference('cold-x') };
      },
    };
    const big = 'y'.repeat(30_000); // 2 chunks → 1 sendChunk that throws
    const r = await deliverDm(deps, { platformUserId: 'aad-77', botSlug: 'maestro', tenantId: 't-1', message: big });
    // create succeeded; the failed second-chunk send is mapped (503 → 502 catch-all).
    assert.equal(calls, 1);
    assert.equal(r.status, 502);
    // OBSERVED: the conversation IS recorded despite the partial failure.
    assert.equal(store.size(), 1, 'documents the partial-delivery/duplicate-on-retry watch item');
    assert.equal(store.findByUserAndBot('aad-77', 'maestro')?.conversationId, 'cold-x');
  });
});
