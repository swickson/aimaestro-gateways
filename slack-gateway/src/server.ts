/**
 * AI Maestro - Slack Gateway (AMP Protocol)
 *
 * Connects to Slack via Bolt (Socket Mode) and bridges messages with
 * AI Maestro agents using the AMP protocol. Runs as a long-lived
 * service managed by pm2.
 *
 * Features:
 * - Bidirectional Slack <-> AMP messaging
 * - AMP auto-registration on first boot
 * - Filesystem-based inbox polling (no HTTP overhead)
 * - Content security (34 injection pattern detection)
 * - Activity logging (ring buffer, 500 events)
 * - Health endpoint and management APIs
 * - Thread context persistence across restarts
 * - Graceful shutdown
 */

import * as path from 'path';
import { pathToFileURL } from 'url';
import express from 'express';
import { App } from '@slack/bolt';
import { createAuthMiddleware } from '@aimaestro/common/auth.js';
import { loadConfig } from './config.js';
import { loadSecurityConfig, type SecurityConfig } from './content-security.js';
import { createAgentResolver } from './agent-resolver.js';
import { ThreadStore } from './thread-store.js';
import { registerInboundHandlers } from './inbound.js';
import { startOutboundPoller } from './outbound.js';
import { createConfigRouter } from './api/config-api.js';
import { createActivityRouter } from './api/activity-api.js';
import { createStatsRouter } from './api/stats-api.js';
import type { GatewayConfig } from './types.js';

interface HttpAppDeps {
  config: GatewayConfig;
  securityConfig: SecurityConfig;
  updateSecurityConfig: (config: SecurityConfig) => void;
  threadCount: () => number;
  slackConnected?: boolean;
}

export function createHttpApp({
  config,
  securityConfig,
  updateSecurityConfig,
  threadCount,
  slackConnected = true,
}: HttpAppDeps) {
  let currentSecurityConfig = securityConfig;
  const httpApp = express();
  httpApp.use(express.json());

  // Health check (public, no auth required)
  httpApp.get('/health', (_req, res) => {
    res.json({
      status: 'ok',
      service: 'slack-gateway',
      protocol: 'AMP',
      slack: { connected: slackConnected },
      amp: {
        agent: config.amp.agentAddress,
        maestro: config.amp.maestroUrl,
        tenant: config.amp.tenant,
      },
      threads: threadCount(),
      timestamp: new Date().toISOString(),
    });
  });

  // Auth middleware for management APIs. Fails closed if ADMIN_TOKEN is blank.
  httpApp.use('/api', createAuthMiddleware(config.adminToken));

  // Management APIs
  httpApp.use(
    '/api/config',
    createConfigRouter(
      () => config,
      () => currentSecurityConfig,
      (newConfig) => {
        currentSecurityConfig = newConfig;
        updateSecurityConfig(newConfig);
      },
      config.adminToken
    )
  );

  httpApp.use('/api/activity', createActivityRouter());

  httpApp.use('/api/stats', createStatsRouter(() => config));

  return httpApp;
}

async function main(): Promise<void> {
  // Load config (async — may trigger AMP auto-registration)
  let config: GatewayConfig;
  let securityConfig: SecurityConfig;
  try {
    config = await loadConfig();
    securityConfig = loadSecurityConfig();
  } catch (err) {
    console.error('[FATAL] Failed to load config:', err);
    process.exit(1);
  }

  console.log('========================================');
  console.log('AI Maestro - Slack Gateway (AMP)');
  console.log('========================================');
  console.log(`Port: ${config.port}`);
  console.log(`Agent: ${config.amp.agentAddress}`);
  console.log(`Default agent: ${config.amp.defaultAgent}`);
  console.log(`Tenant: ${config.amp.tenant}`);
  console.log(`Maestro: ${config.amp.maestroUrl}`);
  console.log(`Inbox: ${config.amp.inboxDir}`);
  console.log(`Poll interval: ${config.polling.intervalMs}ms`);
  console.log(`Security: ${securityConfig.operatorSlackIds.length} operator Slack ID(s) whitelisted`);
  console.log(`Debug: ${config.debug}`);

  // Create Slack Bolt app (Socket Mode)
  const slackApp = new App({
    token: config.slack.botToken,
    appToken: config.slack.appToken,
    signingSecret: config.slack.signingSecret,
    socketMode: true,
  });

  // Create agent resolver
  const resolver = createAgentResolver(config, slackApp);

  // Create thread store with persistence
  const threadStore = new ThreadStore();
  const threadStorePath = path.resolve(
    path.dirname(new URL(import.meta.url).pathname),
    '..',
    'thread-store.json'
  );
  threadStore.load(threadStorePath);
  threadStore.startCleanup(60000);

  // Register Slack event handlers
  registerInboundHandlers(slackApp, config, resolver, securityConfig, threadStore);

  // Start the Slack app (Socket Mode connection)
  await slackApp.start();
  console.log('Connected to Slack (Socket Mode)');

  // Start polling AMP inbox for agent responses
  const stopPoller = startOutboundPoller(config, slackApp, threadStore);

  // Express server for health checks and management APIs
  const httpApp = createHttpApp({
    config,
    securityConfig,
    updateSecurityConfig: (newConfig) => {
      securityConfig = newConfig;
    },
    threadCount: () => threadStore.size(),
  });

  const server = httpApp.listen(config.port, '127.0.0.1', () => {
    console.log(`[HTTP] Management API on http://127.0.0.1:${config.port}`);
  });

  console.log('');
  console.log('Endpoints:');
  console.log('  GET  /health        - Health check');
  console.log('  GET  /api/config    - Gateway config');
  console.log('  GET  /api/stats     - Gateway metrics');
  console.log('  GET  /api/activity  - Activity log');
  console.log('========================================');
  console.log('');
  console.log('Gateway ready! (AMP Protocol)');
  console.log('  - DM the bot or @mention in channels');
  console.log('  - Use @AIM:agent-name to route to specific agents');
  console.log('  - Messages routed via AMP protocol');
  console.log('  - Responses delivered via filesystem inbox');

  // Graceful shutdown
  let isShuttingDown = false;

  async function shutdown(signal: string): Promise<void> {
    if (isShuttingDown) return;
    isShuttingDown = true;

    console.log(`\n[SHUTDOWN] Received ${signal}, shutting down...`);

    stopPoller();
    threadStore.stopCleanup();
    threadStore.save(threadStorePath);
    console.log('[SHUTDOWN] Thread store saved');

    resolver.clearCaches();

    server.close(() => {
      console.log('[SHUTDOWN] HTTP server closed');
    });

    try {
      await slackApp.stop();
      console.log('[SHUTDOWN] Slack connection closed');
    } catch (error) {
      if (config.debug) {
        console.log('[SHUTDOWN] Error closing Slack:', error);
      }
    }

    console.log('[SHUTDOWN] Complete');
    process.exit(0);
  }

  process.on('SIGTERM', () => shutdown('SIGTERM'));
  process.on('SIGINT', () => shutdown('SIGINT'));
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((err) => {
    console.error('Fatal error:', err);
    process.exit(1);
  });
}
