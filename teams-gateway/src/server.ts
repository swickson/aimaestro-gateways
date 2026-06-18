/**
 * Teams gateway — server bootstrap (Phase-1.3 SKELETON: infra + first identity).
 *
 * What this phase proves (the wiring), and what it defers (the logic):
 *   - N `@microsoft/teams.apps` `App`s, ONE per bot, each constructed from ONLY
 *     that bot's Azure AD credentials and sharing a single `ExpressAdapter`
 *     (spike-validated). Each path's JWT audience is therefore fixed BY
 *     CONSTRUCTION — the BotIdentity binding invariant (§0.2) — not a post-hoc
 *     compare. `skipAuth` stays OFF; the public `/api/<slug>/messages` paths are
 *     Bot-Service-JWT-authed by the adapter.
 *   - Inbound handler is a logged no-op STUB. Phase 2 replaces it with the real
 *     ack-fast / dedupe-by-activity.id / async `/api/v1/route` + Discord-shaped
 *     enriched envelope. Outbound poller is NOT started (Phase 3).
 *
 * AUTH-BY-STRUCTURE (path layout — do NOT reintroduce a blanket `/api` gate):
 *   - `/health`                 — public, unauthenticated.
 *   - `/api/<slug>/messages`    — public, Bot-Service-JWT-authed by the adapter.
 *                                 MUST NOT carry ADMIN_TOKEN (§0.2 don't-conflate).
 *   - `/api/admin/*`            — management, gated as a GROUP by createAuthMiddleware
 *                                 on the sub-router. A new mgmt route added under
 *                                 this router is auth-protected by structure, not by
 *                                 remembering a per-route guard.
 *   The `admin` slug is reserved (bot-registry.ts) so no bot path can shadow the
 *   management mount.
 */

import express from 'express';
import type { Server } from 'node:http';
import { App, ExpressAdapter } from '@microsoft/teams.apps';
import type { Account, ConversationReference } from '@microsoft/teams.api';
import { createAuthMiddleware } from '@aimaestro/common/auth.js';
import { bootstrapGateway, type BotRegistration } from '@aimaestro/common/amp-bootstrap.js';
import { Cache } from '@aimaestro/common/cache.js';
import { loadConfig } from './config.js';
import { toBotConfigs } from './bot-registry.js';
import { handleInbound, type InboundActivity, type InboundDeps } from './inbound.js';
import type { RawInboundAttachment } from './attachments-inbound.js';
import { buildAttachmentDownloader, type ConnectorTokenGetter } from './attachment-download.js';
import { botWasMentioned, rejectMismatchedRecipient } from './recipient-binding.js';
import { createThreadStore, type ThreadStore } from './thread-store.js';
import { createDmRouter } from './dm.js';
import { createUserResolver, type UserResolver } from './user-resolver.js';
import { startOutboundPoller, type OutboundBot } from './outbound.js';
import { loadMeshOrigins } from './mesh-hosts.js';
import { restoreThreadStore, saveThreadStore, startSnapshotTimer } from './thread-persistence.js';
import { buildCard } from './card-builder.js';
import type { GatewayConfig } from './types.js';

const GATEWAY_NAME = 'teams-gateway';
const SHUTDOWN_TIMEOUT_MS = 10_000;
/** AMP `/api/v1/route` request timeout. */
const ROUTE_TIMEOUT_MS = 10_000;
/** Dedupe seen-set TTL — Bot Framework retries land within seconds; 10 min is generous. */
const DEDUPE_TTL_MS = 10 * 60 * 1000;

/** Shared, bot-agnostic inbound services (one set for the whole gateway). */
interface InboundServices {
  userResolver: UserResolver;
  threadStore: ThreadStore;
  /** activity.id seen-set, shared across bots (ids are globally unique). */
  dedupe: Cache<true>;
}

/**
 * Mention-strip the bot's own `<at>` tag (and only the bot's) from a message
 * activity, preferring the SDK's `stripMentionsText` and falling back to the raw
 * text. Typed structurally so this module stays decoupled from the SDK's exact
 * (churn-prone) activity type. Stripping only the bot mention preserves any
 * injection payload for the scanner (red-team §0.1).
 */
function extractStrippedText(activity: {
  text?: string;
  recipient?: { id?: string };
  stripMentionsText?: (opts?: { accountId?: string }) => unknown;
}): string {
  if (typeof activity.stripMentionsText === 'function') {
    try {
      const opts = activity.recipient?.id ? { accountId: activity.recipient.id } : undefined;
      const stripped = activity.stripMentionsText(opts);
      if (typeof stripped === 'string') return stripped;
    } catch {
      /* fall back to raw text */
    }
  }
  return activity.text ?? '';
}

/** Teams file-send wrapper content type — carries the real download URL + name. */
const TEAMS_FILE_DOWNLOAD_INFO = 'application/vnd.microsoft.teams.file.download.info';
/** Per-attachment HTTP leg timeout (download header arrival + Maestro upload/confirm/status). */
const ATTACHMENT_TIMEOUT_MS = 20_000;
/**
 * Per-chunk body IDLE (stall) timeout for the inbound attachment download. The
 * Bot Framework connector intermittently STALLS the response body after headers;
 * this aborts a truly-stalled stream (no bytes for this long) without killing a
 * slow-but-progressing transfer. A stall retries once before failing open.
 */
const ATTACHMENT_BODY_STALL_MS = 8_000;

/**
 * Minimal extension -> MIME map. The gateway-declared content type is ADVISORY
 * (Maestro magic-byte sniffs authoritatively at confirm); this exists only so the
 * gateway deny-list can catch known-dangerous executables by name. Not exhaustive.
 */
const EXT_MIME: Record<string, string> = {
  exe: 'application/x-msdownload', dll: 'application/x-msdownload', com: 'application/x-msdownload',
  bat: 'application/x-msdos-program', cmd: 'application/x-msdos-program', msi: 'application/x-msdownload',
  sh: 'application/x-sh', bin: 'application/x-executable',
  pdf: 'application/pdf', txt: 'text/plain', csv: 'text/csv', json: 'application/json',
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp',
  doc: 'application/msword', docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xls: 'application/vnd.ms-excel', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  zip: 'application/zip',
};

function mimeFromName(name: string | undefined): string {
  const ext = name?.split('.').pop()?.toLowerCase() ?? '';
  return EXT_MIME[ext] ?? 'application/octet-stream';
}

/** Structural (SDK-churn-decoupled) view of a Teams activity attachment. */
interface SdkAttachment {
  contentType?: string;
  contentUrl?: string;
  name?: string;
  content?: unknown;
}

/**
 * Extract the SDK-decoupled inbound attachments from a Teams activity. Teams file
 * sends arrive as a `file.download.info` wrapper carrying `content.downloadUrl`
 * (a pre-authenticated URL) + the real file name; inline content (images, etc.)
 * carries a `contentUrl` (http or `data:` URI). Anything without a fetchable URL
 * is skipped here (nothing to download).
 */
function extractInboundAttachments(activity: { attachments?: SdkAttachment[] }): RawInboundAttachment[] {
  const list = activity.attachments;
  if (!Array.isArray(list) || list.length === 0) return [];
  const out: RawInboundAttachment[] = [];
  for (const att of list) {
    const contentType = att.contentType ?? '';
    if (contentType === TEAMS_FILE_DOWNLOAD_INFO) {
      const content = (att.content ?? {}) as { downloadUrl?: string };
      if (content.downloadUrl) {
        out.push({ name: att.name ?? 'file', contentType: mimeFromName(att.name), downloadUrl: content.downloadUrl });
      }
      continue;
    }
    if (att.contentUrl) {
      out.push({ name: att.name ?? 'attachment', contentType, contentUrl: att.contentUrl });
    }
  }
  return out;
}

/** Per-bot runtime state surfaced on `/health`. */
interface BotRuntime {
  slug: string;
  agentName: string;
  defaultAgent: string;
  messagingPath: string;
  /** Resolved AMP address (null under dry-run — no network register). */
  ampAddress: string | null;
  /** Always true: public bot paths keep Bot-Service-JWT auth (skipAuth OFF). */
  authEnabled: boolean;
}

function messagingPath(slug: string): `/${string}` {
  return `/api/${slug}/messages`;
}

/**
 * Construct one `App` per bot from its own credentials, all sharing `adapter`.
 * `app.initialize()` registers that bot's `messagingEndpoint` POST route on the
 * shared Express app WITHOUT starting a listener (we listen once ourselves).
 *
 * Each bot's `message` handler is ACK-FAST: it extracts the activity synchronously
 * and fires `handleInbound` WITHOUT awaiting, then returns — the SDK sends HTTP
 * 200 as soon as the handler promise resolves, so the heavy `/api/v1/route`
 * round-trip never blocks the response (red-team §0.1 C1). An unhandled rejection
 * in the async pipeline is logged loudly — never silently dropped.
 */
async function buildBotApps(
  adapter: ExpressAdapter,
  config: GatewayConfig,
  registrations: Map<string, BotRegistration>,
  services: InboundServices,
): Promise<Map<string, App>> {
  // slug -> initialized App, so the outbound poller can post a reply under the
  // ORIGINATING bot's identity (App.send is bound to that bot's credentials).
  const apps = new Map<string, App>();
  for (const bot of config.bots) {
    const app = new App({
      clientId: bot.appId,
      clientSecret: bot.appPassword,
      tenantId: bot.appTenantId,
      httpServerAdapter: adapter,
      messagingEndpoint: messagingPath(bot.slug),
      // skipAuth intentionally omitted (defaults OFF) — never ship a public bot
      // path with auth disabled (spike finding #5).
    });

    const registration = registrations.get(bot.slug);
    const deps: InboundDeps = {
      bot: {
        slug: bot.slug,
        defaultAgent: bot.defaultAgent,
        agentName: bot.agentName,
        ampAddress: registration?.address ?? '',
        ampApiKey: registration?.apiKey ?? '',
      },
      maestroUrl: config.amp.maestroUrl,
      operatorAadObjectIds: config.operatorAadObjectIds,
      userResolver: services.userResolver,
      threadStore: services.threadStore,
      dedupe: services.dedupe,
      attachmentPolicy: config.attachments,
      timeoutMs: ROUTE_TIMEOUT_MS,
      debug: config.debug,
    };
    // Connector-token source for authenticating inbound inline-image `contentUrl`
    // downloads (Bot Framework connector). SDK coupling stays HERE: the per-bot App
    // mints its own connector token. Fail-open → null so the downloader still tries
    // (and drops the attachment on 401) rather than crashing the inbound pipeline.
    const getConnectorToken: ConnectorTokenGetter = async () => {
      try {
        const token = await app.tokenManager.getBotToken();
        return token?.toString() ?? null;
      } catch (err) {
        console.error(`[TEAMS] (${bot.slug}) connector token fetch failed: ${(err as Error).message}`);
        return null;
      }
    };
    const downloadAttachment = buildAttachmentDownloader(
      bot.slug,
      getConnectorToken,
      ATTACHMENT_TIMEOUT_MS,
      ATTACHMENT_BODY_STALL_MS,
    );

    app.on('message', async (ctx) => {
      // ACK-FAST: build the SDK-decoupled DTO synchronously, fire the pipeline
      // without awaiting, then return so the adapter can 200 immediately.
      const a = ctx.activity;

      // RECIPIENT-IDENTITY BINDING (§0.2): cheap defense-in-depth reject BEFORE any
      // resolve/scan/thread-store/route. The adapter already audience-validated the
      // Bot-Service JWT against this bot's appId; if the activity nonetheless names a
      // different recipient, drop it here — never resolve a user or route under the
      // wrong bot identity. Absent recipient id falls through to the adapter check.
      if (rejectMismatchedRecipient({
        recipientId: a.recipient?.id,
        appId: bot.appId,
        slug: bot.slug,
        activityId: a.id,
      })) {
        return;
      }

      const activity: InboundActivity = {
        activityId: a.id ?? '',
        conversationId: a.conversation?.id ?? '',
        conversationType: String(a.conversation?.conversationType ?? 'unknown'),
        aadObjectId: a.from?.aadObjectId,
        fromId: a.from?.id ?? '',
        fromName: a.from?.name ?? '',
        text: extractStrippedText(a),
        // #12: addressing + channel/team context for the scope+mention gate and the
        // advisory room descriptor. Personal scope ignores all three.
        mentionsBot: botWasMentioned(a),
        teamId: a.channelData?.team?.id,
        channelId: a.channelData?.channel?.id,
        tenantId: a.channelData?.tenant?.id,
        serviceUrl: a.serviceUrl,
        reference: ctx.ref,
        attachments: extractInboundAttachments(a),
        downloadAttachment,
      };
      void handleInbound(activity, deps).catch((err) => {
        console.error(`[TEAMS] (${bot.slug}) unhandled inbound error for activity ${activity.activityId}:`, err);
      });
    });

    await app.initialize();
    apps.set(bot.slug, app);
    console.log(`[TEAMS] bot '${bot.slug}' -> ${bot.agentName} mounted at ${messagingPath(bot.slug)} (auth ON)`);
  }
  return apps;
}

/**
 * Build the outbound poller's per-bot surface from the live registrations. Only
 * bots that registered for real (non-empty `inboxDir` + `apiKey`) get an inbox to
 * poll — under dry-run nothing is registered, so this returns empty and the poller
 * is never started. `send` is the proactive `App.send` bound to the bot's own
 * identity (Fork-O1 Option A: rebuilds the ref from the bot's configured
 * serviceUrl; the poller logs a mismatch against the stored serviceUrl).
 */
function buildOutboundBots(
  config: GatewayConfig,
  registrations: Map<string, BotRegistration>,
  apps: Map<string, App>,
  meshOrigins: ReadonlySet<string>,
): OutboundBot[] {
  // Trusted-origin allowlist for outbound attachment pulls: this gateway's own
  // maestroUrl origin (baseline — keeps same-host config working + degrades safely
  // when hosts.json is absent) UNION the enabled mesh-host origins. Computed once;
  // a hosts.json change needs a gateway restart to take effect.
  const allowedOrigins: ReadonlySet<string> = new Set([
    new URL(config.amp.maestroUrl).origin,
    ...meshOrigins,
  ]);
  const bots: OutboundBot[] = [];
  for (const bot of config.bots) {
    const registration = registrations.get(bot.slug);
    const app = apps.get(bot.slug);
    if (!registration?.inboxDir || !registration.apiKey || !app) continue;
    bots.push({
      slug: bot.slug,
      inboxDir: registration.inboxDir,
      maestroUrl: config.amp.maestroUrl,
      allowedOrigins,
      configuredServiceUrl: app.api.serviceUrl,
      send: async (conversationId, text, markdown, attachments, card) => {
        await app.send(conversationId, {
          type: 'message',
          ...(card
            ? {
                attachments: [
                  {
                    contentType: 'application/vnd.microsoft.card.adaptive',
                    content: card,
                  },
                ],
              }
            : {
                text,
                // markdown is the Teams default; only set the field to fall back to plain.
                ...(markdown ? {} : { textFormat: 'plain' }),
                // w3: inline each attachment as a base64 data URI. Small images render
                // inline; larger files surface as a download. The exact Teams attachment
                // shape (hosted vs file-consent vs Graph) is a LIVE-AZURE WATCH ITEM —
                // verified at deploy, same class as the Fork-O1 App.send unknown.
                ...(attachments?.length
                  ? {
                      attachments: attachments.map((a) => ({
                        contentType: a.contentType || 'application/octet-stream',
                        contentUrl: `data:${a.contentType || 'application/octet-stream'};base64,${Buffer.from(a.bytes).toString('base64')}`,
                        name: a.filename,
                      })),
                    }
                  : {}),
              }),
        });
      },
    });
  }
  return bots;
}

/**
 * Management sub-router, gated as a GROUP by createAuthMiddleware. Mounted at
 * `/api/admin`. Inbound/outbound + DM logic is deferred — these are wiring stubs
 * that prove the auth boundary is enforced by structure.
 */
function buildManagementRouter(
  config: GatewayConfig,
  getBots: () => BotRuntime[],
): express.Router {
  const router = express.Router();
  router.use(express.json());
  // Whole-group gate: fail-closed by construction (throws if ADMIN_TOKEN blank).
  router.use(createAuthMiddleware(config.adminToken));

  router.get('/stats', (_req, res) => {
    res.json({
      service: GATEWAY_NAME,
      bots: getBots(),
      bootstrap: config.dryRunBootstrap ? 'dry-run' : 'registered',
    });
  });

  // Proactive DM delivery is mounted separately at `/api/gateway/dm` (its own
  // ADMIN_TOKEN-gated group — the path Maestro's notifyUser POSTs), so it is NOT
  // a route on this management router. See `mountGatewayDm` in main().

  return router;
}

/**
 * Bootstrap per-bot AMP identities. DRY-RUN resolves + logs the plan with ZERO
 * network side effect (used for the skeleton boot smoke). Real registration is a
 * deploy-time, coordinated live-directory change (Phase 5/6).
 */
async function runBootstrap(config: GatewayConfig): Promise<Map<string, BotRegistration>> {
  const registrations = new Map<string, BotRegistration>();
  if (config.dryRunBootstrap) {
    console.log(`[BOOTSTRAP] DRY-RUN — would register ${config.bots.length} bot identity(ies) (no network):`);
    for (const bot of config.bots) {
      console.log(`[BOOTSTRAP]   ${bot.slug} -> ${bot.agentName} (default agent ${bot.defaultAgent})`);
    }
    return registrations;
  }

  const registration = await bootstrapGateway({
    gatewayName: GATEWAY_NAME,
    maestroUrl: config.amp.maestroUrl,
    tenant: config.amp.tenant,
    bots: toBotConfigs(config.bots),
  });
  for (const [slug, reg] of Object.entries(registration.bots)) {
    registrations.set(slug, reg);
  }
  return registrations;
}

async function main(): Promise<void> {
  const config = loadConfig();
  console.log('========================================');
  console.log(`${GATEWAY_NAME} starting (AMP Protocol)`);
  console.log(`  bots:      ${config.bots.map((b) => b.slug).join(', ')}`);
  console.log(`  maestro:   ${config.amp.maestroUrl}`);
  console.log(`  bootstrap: ${config.dryRunBootstrap ? 'DRY-RUN' : 'live register'}`);
  console.log('========================================');

  const registrations = await runBootstrap(config);

  const botRuntimes: BotRuntime[] = config.bots.map((bot) => ({
    slug: bot.slug,
    agentName: bot.agentName,
    defaultAgent: bot.defaultAgent,
    messagingPath: messagingPath(bot.slug),
    ampAddress: registrations.get(bot.slug)?.address ?? null,
    authEnabled: true,
  }));

  // Shared inbound services (bot-agnostic; one set for the whole gateway). User
  // resolution is platform-wide, so it authenticates with any one registered
  // bot's AMP key. Pick the first NON-EMPTY key: a partially-registered fleet can
  // leave an earlier bot keyless, and an empty Bearer token would 401 every
  // resolve (empty under dry-run — exercised only by real activities).
  const sharedApiKey = [...registrations.values()].map((r) => r.apiKey).find((k) => k && k.trim() !== '') ?? '';
  const threadStore = createThreadStore();
  // Boot-restore the thread-store BEFORE the poller starts — a reply that landed
  // while we were down must be able to resolve its conversation reference.
  restoreThreadStore(threadStore, config.threadStorePath);

  const inboundServices: InboundServices = {
    userResolver: createUserResolver({
      maestroUrl: config.amp.maestroUrl,
      apiKey: sharedApiKey,
      cacheTtlMs: config.cacheUserTtlMs,
      debug: config.debug,
    }),
    threadStore,
    dedupe: new Cache<true>(DEDUPE_TTL_MS),
  };

  const httpApp = express();
  const adapter = new ExpressAdapter(httpApp);

  // Public health check (unauthenticated) — registered before the bot routes;
  // paths are disjoint so order is not load-bearing, but health stays outside
  // the management auth group by design.
  httpApp.get('/health', (_req, res) => {
    res.json({
      status: 'ok',
      service: GATEWAY_NAME,
      protocol: 'AMP',
      bootstrap: config.dryRunBootstrap ? 'dry-run' : 'registered',
      adapter: { ready: true, sharedExpressAdapter: true, botCount: botRuntimes.length },
      bots: botRuntimes,
      amp: { maestro: config.amp.maestroUrl },
      timestamp: new Date().toISOString(),
    });
  });

  // Management group — gated as a whole by createAuthMiddleware on the sub-router.
  httpApp.use('/api/admin', buildManagementRouter(config, () => botRuntimes));

  // Per-bot Teams adapters — register the public `/api/<slug>/messages` routes.
  const apps = await buildBotApps(adapter, config, registrations, inboundServices);

  // Proactive DM endpoint (Phase 5) — mounted AFTER buildBotApps so the slug->App
  // map exists. Its own ADMIN_TOKEN-gated group at `/api/gateway` (the path
  // Maestro's notifyUser POSTs); `sendChunk` binds delivery to the originating
  // bot's own `App.send` (same proactive send shape the outbound poller uses).
  httpApp.use(
    '/api/gateway',
    createDmRouter({
      threadStore,
      knownBots: new Set(apps.keys()),
      markdownDefault: config.markdownDefault,
      coldStartEnabled: config.dmColdStartEnabled,
      adminToken: config.adminToken,
      sendChunk: async (botSlug, conversationId, text, markdown) => {
        const app = apps.get(botSlug);
        if (!app) throw new Error(`no App for bot '${botSlug}'`);
        await app.send(conversationId, {
          type: 'message',
          text,
          ...(markdown ? {} : { textFormat: 'plain' }),
        });
      },
      createColdStartConversation: async ({ botSlug, tenantId, aadObjectId }) => {
        const app = apps.get(botSlug);
        const bot = config.bots.find((b) => b.slug === botSlug);
        if (!app || !bot) throw new Error(`no App for bot '${botSlug}'`);

        const user: Account = { id: aadObjectId, aadObjectId, name: aadObjectId, role: 'user' };
        const botAccount: Account = { id: bot.appId, name: bot.slug, role: 'bot' };
        // Ensure-only: NO inline `activity`. On an existing 1:1 the inline activity is
        // silently dropped by Bot Framework (#25) — deliverDm posts every chunk via
        // App.send instead. `created.activityId` is then empty; guards below handle it.
        const created = await app.api.conversations.create({
          isGroup: false,
          tenantId,
          bot: botAccount,
          members: [user],
        });
        const reference: ConversationReference = {
          serviceUrl: created.serviceUrl,
          channelId: 'msteams',
          conversation: { id: created.id, conversationType: 'personal', tenantId, isGroup: false },
          bot: botAccount,
          user,
          ...(created.activityId && { activityId: created.activityId }),
        };
        return {
          conversationId: created.id,
          reference,
          rootActivityId: created.activityId || `dm:${botSlug}:${aadObjectId}`,
        };
      },
    }),
  );

  // Outbound delivery (Phase 3): poll each registered bot's AMP inbox for agent
  // replies and post them back under that bot. Under dry-run nothing registered,
  // so there are no inboxes to poll and the poller is not started.
  // Load the trusted mesh-host origin allowlist ONCE at startup (restart to refresh).
  // Outbound attachment download urls are accepted only when their origin is in this
  // set ∪ the maestroUrl origin. An EMPTY mesh set means signed urls carrying a remote
  // host's Tailscale origin will be rejected — i.e. outbound attachments are effectively
  // OFF — so warn loud and unmistakably in the startup log.
  const meshOrigins = loadMeshOrigins();
  if (meshOrigins.size === 0) {
    console.warn(
      '[MESH] ⚠️  NO mesh-host origins loaded — OUTBOUND ATTACHMENTS to remote/mesh hosts are DISABLED ' +
        '(only this gateway\'s own maestroUrl origin is trusted). Fix ~/.aimaestro/hosts.json and restart to enable.',
    );
  } else {
    console.log(`[MESH] trusted outbound attachment origins: ${[...meshOrigins].join(', ')}`);
  }
  const outboundBots = buildOutboundBots(config, registrations, apps, meshOrigins);
  let stopOutbound: (() => void) | null = null;
  let stopSnapshot: (() => void) | null = null;
  if (outboundBots.length > 0) {
    stopOutbound = startOutboundPoller({
      bots: outboundBots,
      threadStore,
      pollIntervalMs: config.polling.intervalMs,
      markdownDefault: config.markdownDefault,
      policy: config.attachments,
      debug: config.debug,
      buildCard,
    });
    // Periodic crash-safety snapshot; the graceful shutdown path saves once more.
    stopSnapshot = startSnapshotTimer(threadStore, config.threadStorePath, config.snapshotIntervalMs);
    console.log(`[TEAMS] thread-store persistence at ${config.threadStorePath} (snapshot every ${config.snapshotIntervalMs}ms).`);
  } else {
    console.log('[OUTBOUND] No registered bot inboxes (dry-run / unregistered) — outbound poller not started.');
  }

  const servers: Server[] = config.host.map((host) =>
    httpApp.listen(config.port, host, () => {
      console.log(`[HTTP] teams-gateway listening on http://${host}:${config.port}`);
    }),
  );

  console.log('');
  console.log('Endpoints:');
  console.log('  GET  /health             - Health check (public)');
  console.log('  GET  /api/admin/stats    - Gateway state (ADMIN_TOKEN)');
  console.log('  POST /api/gateway/dm     - Proactive DM (ADMIN_TOKEN)');
  for (const bot of config.bots) {
    console.log(`  POST ${messagingPath(bot.slug)}  - ${bot.slug} inbound (Bot Service JWT)`);
  }
  console.log('========================================');
  console.log(
    outboundBots.length > 0
      ? 'Gateway ready. Inbound + outbound pipelines LIVE (personal scope).'
      : 'Gateway ready. Inbound pipeline LIVE (personal scope); outbound idle (no registered inboxes).',
  );

  let isShuttingDown = false;
  async function shutdown(signal: string): Promise<void> {
    if (isShuttingDown) return;
    isShuttingDown = true;
    console.log(`\n[SHUTDOWN] Received ${signal}, shutting down...`);

    const forceExit = setTimeout(() => {
      console.error('[SHUTDOWN] Timed out — forcing exit.');
      process.exit(1);
    }, SHUTDOWN_TIMEOUT_MS);
    forceExit.unref();

    for (const server of servers) {
      server.close(() => {});
    }
    console.log('[SHUTDOWN] HTTP listeners closed');

    // Stop the poller + periodic snapshot, then take one final clean snapshot so a
    // reply that arrives next boot can still resolve its conversation reference.
    if (stopOutbound) stopOutbound();
    if (stopSnapshot) stopSnapshot();
    if (outboundBots.length > 0) {
      if (saveThreadStore(threadStore, config.threadStorePath)) {
        console.log(`[SHUTDOWN] thread-store snapshot saved (${threadStore.size()} entries).`);
      }
    }
    void apps; // per-bot Apps hold no open sockets of their own (shared adapter).

    console.log('[SHUTDOWN] Complete');
    process.exit(0);
  }

  process.on('SIGTERM', () => void shutdown('SIGTERM'));
  process.on('SIGINT', () => void shutdown('SIGINT'));
}

main().catch((err) => {
  console.error('[FATAL] teams-gateway failed to start:', err);
  process.exit(1);
});
