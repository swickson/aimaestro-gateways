/**
 * Issue #11 — cross-review regression suite (authored from the SPEC, not the impl).
 *
 * Contract under test: an inbound Teams message from an EXTERNAL sender that is
 * attachment-only (empty original text) and whose attachments ALL fail to ingest
 * must carry a gateway-authored "dropped" signal so the recipient is not silently
 * starved. The scanner ALWAYS wraps external content in <external-content …> — even
 * when the text is empty — so the legacy `sanitized.trim() === ''` guard never fires
 * for external senders. The fix keys the notice off the RAW `activity.text`.
 *
 * Four boundary cases the spec names (dispatch-issue-11.md §"Required behavior" +
 * §"Constraints"):
 *   (a) external + empty text + ALL attachments failed  -> count-only notice, OUTSIDE
 *       the data-only fence, no filename/type/size/sender metadata. N = inbound count.
 *   (b) operator + empty text + ALL failed              -> existing constant placeholder
 *       path still fires (count-only notice must NOT hijack the operator path).
 *   (c) external + empty text + >=1 attachment SUCCEEDS -> keeps the <external-content>
 *       wrapper + payload.attachments; notice must NOT fire.
 *   (d) external + non-empty text + ALL failed          -> scanned/wrapped text routes;
 *       NO gateway notice (the recipient already has the sender's words).
 *
 * Self-contained harness (mirrors inbound-security.test.ts): mock `fetch` for the
 * Maestro /route call (+ the attachment upload->PUT->confirm->status legs in case c),
 * a mock `downloadAttachment` closure for the ingest path. No live Azure/Maestro.
 */

import assert from 'node:assert/strict';
import { afterEach, beforeEach, describe, it } from 'node:test';

import { Cache } from '@aimaestro/common/cache.js';

import type { RawInboundAttachment } from '../attachments-inbound.js';
import { handleInbound, type InboundActivity, type InboundDeps } from '../inbound.js';
import { createThreadStore, type ThreadStore } from '../thread-store.js';
import type { AMPRouteRequest, ResolvedUser } from '../types.js';

const originalFetch = globalThis.fetch;
const originalConsoleError = console.error;
const originalConsoleLog = console.log;

beforeEach(() => {
  // Silence the gateway's [CONTEXT] logs; tests assert on routed payloads, not stdout.
  console.log = () => undefined;
});
afterEach(() => {
  globalThis.fetch = originalFetch;
  console.error = originalConsoleError;
  console.log = originalConsoleLog;
});

function user(overrides: Partial<ResolvedUser> = {}): ResolvedUser {
  return {
    id: 'user-1',
    displayName: 'Mallory External',
    aliases: [],
    platforms: [],
    role: 'external',
    trustLevel: 'none',
    createdAt: '2026-06-09T00:00:00.000Z',
    updatedAt: '2026-06-09T00:00:00.000Z',
    ...overrides,
  };
}

/** A raw inbound attachment descriptor (pre-ingest), e.g. a no-caption photo. */
function rawAttachment(overrides: Partial<RawInboundAttachment> = {}): RawInboundAttachment {
  return { name: 'secret-plans.pdf', contentType: 'application/pdf', downloadUrl: 'https://teams.test/dl/x', ...overrides };
}

function activity(overrides: Partial<InboundActivity> = {}): InboundActivity {
  return {
    activityId: 'activity-1',
    conversationId: 'conversation-1',
    conversationType: 'personal',
    aadObjectId: 'aad-sender',
    fromId: 'bf-user-1',
    fromName: 'Mallory',
    text: 'status report',
    tenantId: 'tenant-1',
    serviceUrl: 'https://smba.trafficmanager.net/amer/',
    reference: {
      serviceUrl: 'https://smba.trafficmanager.net/amer/',
      channelId: 'msteams',
      conversation: { id: 'conversation-1' },
      bot: { id: 'bot-app-id', name: 'Maestro' },
      user: { id: 'bf-user-1', name: 'Mallory' },
    } as InboundActivity['reference'],
    ...overrides,
  };
}

function routeResponse(id = 'amp-1'): Response {
  return new Response(JSON.stringify({ id, status: 'delivered' }), {
    status: 200,
    headers: { 'content-type': 'application/json' },
  });
}

function makeDeps(overrides: Partial<InboundDeps> = {}): InboundDeps & { threadStore: ThreadStore } {
  return {
    bot: {
      slug: 'maestro',
      defaultAgent: 'ops-agent@example.aimaestro.local',
      agentName: 'teams-maestro-bot',
      ampAddress: 'teams-maestro-bot@example.aimaestro.local',
      ampApiKey: 'amp-secret',
    },
    maestroUrl: 'http://maestro.test',
    operatorAadObjectIds: [],
    userResolver: {
      resolve: async () => user(),
      clearCache: () => undefined,
    },
    threadStore: createThreadStore(),
    dedupe: new Cache<true>(60_000),
    attachmentPolicy: { maxBytes: 26_214_400, maxCount: 10, denyContentTypes: [] },
    timeoutMs: 1000,
    debug: false,
    now: () => 1_000_000,
    ...overrides,
  };
}

/** Force the default sender to resolve OPERATOR via the tenant-scoped legacy whitelist. */
function operatorDeps(overrides: Partial<InboundDeps> = {}): InboundDeps & { threadStore: ThreadStore } {
  return makeDeps({
    operatorAadObjectIds: [{ tenantId: 'tenant-1', aadObjectId: 'aad-sender' }],
    ...overrides,
  });
}

/** Captures the /route POST body. All-failed cases never touch the attachment legs
 *  (downloadAttachment throws before any fetch), so /route is the only call. */
function installRouteFetch(calls: AMPRouteRequest[]): void {
  globalThis.fetch = (async (_input: RequestInfo | URL, init?: RequestInit) => {
    calls.push(JSON.parse(String(init?.body)) as AMPRouteRequest);
    return routeResponse();
  }) as typeof fetch;
}

/** A downloader that always throws — models every attachment failing to ingest. */
const failingDownloader = async (): Promise<Uint8Array> => {
  throw new Error('teams 403 — download refused');
};

const NOTICE_RE = /could not be retrieved/;

describe('Issue #11 — external attachment-only all-failed dropped signal', () => {
  // ── (a) external + empty text + ALL failed -> count-only notice outside the fence ──
  it('(a) emits a count-only notice OUTSIDE the data-only fence when an external attachment-only message has all attachments fail', async () => {
    const routed: AMPRouteRequest[] = [];
    installRouteFetch(routed);
    const deps = makeDeps();

    const status = await handleInbound(
      activity({ text: '', attachments: [rawAttachment(), rawAttachment({ name: 'b.pdf' })], downloadAttachment: failingDownloader }),
      deps,
    );

    assert.equal(status, 'routed');
    const msg = routed[0]?.payload.message ?? '';
    // N = inbound attachment count (2), NOT routed (0).
    assert.equal(msg, '[Teams: 2 attachment(s) received but could not be retrieved]');
    // Gateway truth, NOT sender data: the notice must not be inside the wrapper.
    assert.doesNotMatch(msg, /^<external-content/);
    assert.doesNotMatch(msg, /external-content/);
    // No untrusted metadata leaked into the gateway-authored notice.
    assert.doesNotMatch(msg, /secret-plans|b\.pdf|application\/pdf|Mallory|aad-sender/);
    // All attachments failed -> none cited on the route.
    assert.equal(Object.hasOwn(routed[0]?.payload ?? {}, 'attachments'), false);
  });

  it('(a) whitespace-only original text is treated as empty (trim contract)', async () => {
    const routed: AMPRouteRequest[] = [];
    installRouteFetch(routed);
    const deps = makeDeps();

    await handleInbound(
      activity({ text: '   \n\t ', attachments: [rawAttachment()], downloadAttachment: failingDownloader }),
      deps,
    );

    assert.equal(routed[0]?.payload.message, '[Teams: 1 attachment(s) received but could not be retrieved]');
  });

  // ── (b) operator + empty text + ALL failed -> existing constant placeholder, no regression ──
  it('(b) operator attachment-only all-failed still routes the existing CONSTANT placeholder (count-only notice does NOT hijack the operator path)', async () => {
    const routed: AMPRouteRequest[] = [];
    installRouteFetch(routed);
    const deps = operatorDeps();

    const status = await handleInbound(
      activity({ text: '', attachments: [rawAttachment(), rawAttachment({ name: 'b.pdf' })], downloadAttachment: failingDownloader }),
      deps,
    );

    assert.equal(status, 'routed');
    // Operator path bypasses the scanner (sanitized === '' for empty text) -> legacy
    // ATTACHMENT_FAILED_PLACEHOLDER fires. Count-free constant, NOT the #11 count notice.
    assert.equal(routed[0]?.payload.message, '[Teams: an attachment was received but could not be retrieved]');
    assert.doesNotMatch(routed[0]?.payload.message ?? '', /\d+ attachment\(s\)/);
  });

  // ── (c) external + empty text + >=1 SUCCESS -> wrapper + attachments preserved, no notice ──
  it('(c) external attachment-only message with >=1 attachment INGESTING keeps the <external-content> wrapper + payload.attachments and does NOT fire the notice', async () => {
    const routed: AMPRouteRequest[] = [];
    // Combined fetch: drive the upload->PUT->confirm->status ingest legs to SUCCESS,
    // then capture the /route POST.
    globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = String(input);
      if (url.includes('/api/v1/route')) {
        routed.push(JSON.parse(String(init?.body)) as AMPRouteRequest);
        return routeResponse();
      }
      if (url.includes('/attachments/upload')) {
        return new Response(JSON.stringify({ attachment_id: 'att-c', upload_url: 'http://maestro.test/signed/put/att-c' }), {
          status: 200, headers: { 'content-type': 'application/json' },
        });
      }
      if (url.includes('/signed/put/')) return new Response(null, { status: 200 });
      if (url.includes('/confirm')) return new Response(JSON.stringify({ ok: true }), { status: 200, headers: { 'content-type': 'application/json' } });
      if (url.includes('/status')) {
        return new Response(JSON.stringify({
          attachment_id: 'att-c', filename: 'safe-name.pdf', content_type: 'application/pdf', size: 4,
          digest: 'a'.repeat(64), scan_status: 'basic_clean',
          uploaded_at: '2026-06-14T00:00:00.000Z', expires_at: '2026-06-15T00:00:00.000Z',
          url: 'http://maestro.test/api/v1/attachments/att-c/download?sig=abc',
        }), { status: 200, headers: { 'content-type': 'application/json' } });
      }
      throw new Error(`no mock handler for ${init?.method ?? 'GET'} ${url}`);
    }) as typeof fetch;
    const deps = makeDeps();

    const status = await handleInbound(
      activity({ text: '', attachments: [rawAttachment()], downloadAttachment: async () => new Uint8Array([1, 2, 3, 4]) }),
      deps,
    );

    assert.equal(status, 'routed');
    const msg = routed[0]?.payload.message ?? '';
    // The no-caption-photo-that-uploaded path is unchanged: wrapper preserved...
    assert.match(msg, /^<external-content /);
    // ...and the #11 notice must NOT fire.
    assert.doesNotMatch(msg, NOTICE_RE);
    // The successfully-ingested attachment is cited.
    assert.equal(routed[0]?.payload.attachments?.length, 1);
  });

  // ── (d) external + non-empty text + ALL failed -> scanned/wrapped text, no notice ──
  it('(d) external NON-empty text with all attachments failed routes the wrapped/scanned text and does NOT add a gateway notice', async () => {
    const routed: AMPRouteRequest[] = [];
    installRouteFetch(routed);
    const deps = makeDeps();

    const status = await handleInbound(
      activity({ text: 'please review the attached file', attachments: [rawAttachment()], downloadAttachment: failingDownloader }),
      deps,
    );

    assert.equal(status, 'routed');
    const msg = routed[0]?.payload.message ?? '';
    // Sender's words survive inside the data-only fence...
    assert.match(msg, /^<external-content /);
    assert.match(msg, /please review the attached file/);
    // ...and the gateway does NOT bolt on a drop notice (the recipient has the text).
    assert.doesNotMatch(msg, NOTICE_RE);
  });
});
