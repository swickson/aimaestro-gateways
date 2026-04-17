/**
 * Discord Gateway - Inbound Message Handlers (AMP Protocol)
 *
 * Registers Discord event handlers (messageCreate) and routes messages
 * to agents via AMP POST /api/v1/route.
 */

import type { Client, Message, TextChannel } from 'discord.js';
import type { GatewayConfig, AMPRouteRequest, ResolvedUser, WatchWebhookEntry } from './types.js';
import type { AgentResolver } from './agent-resolver.js';
import type { UserResolver } from './user-resolver.js';
import type { ThreadStore } from './thread-store.js';
import { sanitizeDiscordMessage, type SecurityConfig } from './content-security.js';
import { logEvent } from './api/activity-log.js';

// ---------------------------------------------------------------------------
// Topic Hints Extraction
// ---------------------------------------------------------------------------

const STOP_WORDS = new Set([
  'a', 'an', 'the', 'is', 'are', 'was', 'were', 'be', 'been', 'being',
  'have', 'has', 'had', 'do', 'does', 'did', 'will', 'would', 'could',
  'should', 'may', 'might', 'can', 'shall', 'to', 'of', 'in', 'for',
  'on', 'with', 'at', 'by', 'from', 'as', 'into', 'through', 'during',
  'before', 'after', 'above', 'below', 'between', 'out', 'off', 'over',
  'under', 'again', 'further', 'then', 'once', 'and', 'but', 'or', 'nor',
  'not', 'so', 'yet', 'both', 'either', 'neither', 'each', 'every', 'all',
  'any', 'few', 'more', 'most', 'other', 'some', 'such', 'no', 'only',
  'same', 'than', 'too', 'very', 'just', 'because', 'if', 'when', 'where',
  'how', 'what', 'which', 'who', 'whom', 'this', 'that', 'these', 'those',
  'i', 'me', 'my', 'we', 'our', 'you', 'your', 'he', 'him', 'his', 'she',
  'her', 'it', 'its', 'they', 'them', 'their', 'about', 'up',
]);

/**
 * Extract lightweight topic hints from message text.
 * Returns up to 3 keywords, filtered for stop words and short tokens.
 */
function extractTopicHints(text: string): string[] {
  const words = text
    .toLowerCase()
    .replace(/[^a-z0-9\s-]/g, ' ')
    .split(/\s+/)
    .filter(w => w.length > 2 && !STOP_WORDS.has(w));

  // Count frequency, take top 3
  const freq = new Map<string, number>();
  for (const w of words) {
    freq.set(w, (freq.get(w) || 0) + 1);
  }

  return [...freq.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3)
    .map(([word]) => word);
}

/**
 * Parse @AIM:agent-name routing from message text.
 * Allows full AMP addresses like @AIM:agent@tenant.domain
 */
function parseAgentRouting(
  text: string,
  defaultAgent: string
): { agent: string; message: string } {
  const match = text.match(/@AIM:([a-zA-Z0-9_@.\-]+)/i);

  if (match) {
    const message = text
      .replace(match[0], '')
      .replace(/\s+/g, ' ')
      .replace(/^[,.\s]+/, '')
      .trim();

    return {
      agent: match[1],
      message: message || '(no message)',
    };
  }

  return {
    agent: defaultAgent,
    message: text,
  };
}

/**
 * Remove bot mention from message text.
 */
function stripBotMention(text: string, botId: string): string {
  return text.replace(new RegExp(`<@!?${botId}>`, 'g'), '').trim();
}

/**
 * Send a message to an agent via AMP route API.
 */
async function sendToAgent(
  config: GatewayConfig,
  targetAddress: string,
  text: string,
  channelId: string,
  messageId: string,
  displayName: string,
  discordUserId: string,
  securityConfig: SecurityConfig,
  threadStore: ThreadStore,
  resolvedUser: ResolvedUser | null,
  isDM: boolean,
  monitorMode: boolean = false
): Promise<void> {
  const { sanitized, trust, flags } = sanitizeDiscordMessage(
    text,
    discordUserId,
    displayName,
    securityConfig,
    resolvedUser
  );

  if (flags.length > 0) {
    console.log(
      `[SECURITY] ${flags.length} injection pattern(s) flagged from ${displayName} (trust: ${trust.level})`
    );
    logEvent('security', `Injection patterns flagged from ${displayName}`, {
      from: displayName,
      to: targetAddress,
      subject: text.substring(0, 80),
      securityFlags: flags.map((f) => `${f.category}: ${f.match}`),
    });
  }

  // Determine if this is a new conversation
  const recentThread = threadStore.findByChannel(channelId);
  const conversationTimeoutMs = 30 * 60 * 1000; // 30 minutes
  const isNewConversation = isDM
    ? !recentThread || (Date.now() - recentThread.createdAt > conversationTimeoutMs)
    : true; // Guild @mentions are always treated as new conversations

  const topicHints = extractTopicHints(text);

  const ampRequest: AMPRouteRequest = {
    to: targetAddress,
    subject: `Discord message from ${displayName}`,
    priority: 'normal',
    payload: {
      type: 'request',
      message: sanitized,
      context: {
        // Legacy fields (kept for backward compatibility)
        channel: {
          type: 'discord',
          sender: displayName,
          sender_id: discordUserId,
          thread_id: channelId,
          bridge_agent: config.amp.agentAddress,
          received_at: new Date().toISOString(),
        },
        discord: { channelId, messageId, user: displayName },
        security: {
          trust: trust.level,
          source: 'discord',
          scanned: true,
          injection_flags: flags.map((f) => f.category),
          wrapped: trust.level !== 'operator',
          scanned_at: new Date().toISOString(),
        },
        // New enriched context (Phase 2)
        sender: {
          platformUserId: discordUserId,
          platform: 'discord',
          handle: displayName,
          ...(resolvedUser && {
            userId: resolvedUser.id,
            displayName: resolvedUser.displayName,
            trustLevel: resolvedUser.trustLevel,
            role: resolvedUser.role,
          }),
        },
        thread: {
          threadId: channelId,
          inReplyTo: recentThread?.ampMessageId || null,
          isNewConversation,
        },
        ...(topicHints.length > 0 && { topicHints }),
        ...(monitorMode && { mode: 'monitor' as const }),
      },
    },
  };

  const response = await fetch(`${config.amp.maestroUrl}/api/v1/route`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${config.amp.apiKey}`,
    },
    body: JSON.stringify(ampRequest),
    signal: AbortSignal.timeout(config.polling.timeoutMs),
  });

  if (!response.ok) {
    const errorBody = await response.text().catch(() => '');
    if (response.status === 404) {
      throw new Error(`agent_not_found: ${targetAddress}`);
    }
    if (response.status === 429) {
      throw new Error(`rate_limited: ${targetAddress}`);
    }
    throw new Error(`AMP route failed (${response.status}): ${errorBody}`);
  }

  const result = await response.json();

  // Skip thread-store mapping in monitor mode so any accidental reply from the
  // target agent does not land back in the watched channel.
  if (result.id && !monitorMode) {
    threadStore.set(result.id, {
      channelId,
      messageId,
      user: discordUserId,
      userName: displayName,
      ampMessageId: result.id,
      createdAt: Date.now(),
    });
  }

  const agentName = targetAddress.split('@')[0];
  console.log(
    `[-> ${targetAddress}] Message from ${displayName} (trust: ${trust.level}): ${text.substring(0, 50)}...`
  );

  logEvent('inbound', `Discord message routed: ${displayName} -> ${agentName}`, {
    from: displayName,
    to: targetAddress,
    subject: text.substring(0, 80),
    ampMessageId: result.id,
    deliveryStatus: result.status,
  });
}

/**
 * Route a Discord message to the appropriate agent via AMP.
 */
async function routeMessage(
  config: GatewayConfig,
  resolver: AgentResolver,
  securityConfig: SecurityConfig,
  threadStore: ThreadStore,
  userResolver: UserResolver,
  text: string,
  channelId: string,
  messageId: string,
  displayName: string,
  discordUserId: string,
  isDM: boolean,
  reply: (text: string) => Promise<void>
): Promise<void> {
  const { agent, message } = parseAgentRouting(text, config.amp.defaultAgent);
  const { address } = resolver.lookupAgent(agent);

  // Resolve sender against user directory
  const resolvedUser = await userResolver.resolve('discord', discordUserId, displayName);

  try {
    await sendToAgent(
      config,
      address,
      message,
      channelId,
      messageId,
      displayName,
      discordUserId,
      securityConfig,
      threadStore,
      resolvedUser,
      isDM
    );
  } catch (error) {
    const errMsg = (error as Error).message;

    if (errMsg.startsWith('agent_not_found:')) {
      await reply(
        `Agent \`${agent}\` not found.\n\nUse \`@AIM:agent-name message\` to route to a specific agent.`
      );
      logEvent('error', `Agent not found: ${agent}`, { from: displayName, to: agent });
      return;
    }

    if (errMsg.startsWith('rate_limited:')) {
      await reply(`Agent \`${agent}\` is rate limited. Please try again in a moment.`);
      return;
    }

    throw error;
  }
}

/**
 * Lookup a watch-webhook config entry for a message, if any.
 * Matches on both channelId and webhookId — either alone is not enough.
 */
function matchWatchWebhook(
  message: Message,
  entries: WatchWebhookEntry[]
): WatchWebhookEntry | null {
  if (!message.webhookId || entries.length === 0) return null;
  return (
    entries.find(
      e => e.channelId === message.channelId && e.webhookId === message.webhookId
    ) || null
  );
}

/**
 * Register all inbound Discord event handlers.
 */
export function registerInboundHandlers(
  client: Client,
  config: GatewayConfig,
  resolver: AgentResolver,
  securityConfig: SecurityConfig,
  threadStore: ThreadStore,
  userResolver: UserResolver
): void {
  client.on('messageCreate', async (message: Message) => {
    // Watch-webhook fast path: route messages from whitelisted webhooks to a
    // fixed agent in monitor mode (no reactions, no reply, no threadStore).
    const watch = matchWatchWebhook(message, config.watchWebhooks);
    if (watch) {
      const text = message.content?.trim();
      if (!text) return;

      const webhookName =
        message.author?.username || `webhook:${message.webhookId}`;

      console.log(
        `[Discord <- watch] #${(message.channel as TextChannel).name || message.channelId} from ${webhookName} -> ${watch.targetAgent}: ${text.substring(0, 50)}...`
      );

      try {
        const { address } = resolver.lookupAgent(watch.targetAgent);
        await sendToAgent(
          config,
          address,
          text,
          message.channelId,
          message.id,
          webhookName,
          message.webhookId!,
          securityConfig,
          threadStore,
          null,
          false,
          true
        );
      } catch (error) {
        console.error('[WATCH] Failed to route webhook message:', error);
        logEvent('error', `Watch-webhook route failed: ${webhookName}`, {
          from: webhookName,
          to: watch.targetAgent,
          error: (error as Error).message,
        });
      }
      return;
    }

    if (message.author.bot) return;

    const isDM = !message.guild;
    const isMentioned = message.mentions.has(client.user!);

    if (!isDM && !isMentioned) return;

    const displayName = message.author.displayName || message.author.username;
    let text = message.content;

    if (isMentioned && client.user) {
      text = stripBotMention(text, client.user.id);
    }

    if (!text.trim()) return;

    const source = isDM ? 'DM' : `#${(message.channel as TextChannel).name || message.channelId}`;
    console.log(`[Discord <-] ${source} from ${displayName}: ${text.substring(0, 50)}...`);

    try {
      await message.react('\u{1F440}').catch(() => {});

      if ('sendTyping' in message.channel) {
        await message.channel.sendTyping().catch(() => {});
      }

      const reply = async (replyText: string) => {
        await message.reply(replyText);
      };

      await routeMessage(
        config,
        resolver,
        securityConfig,
        threadStore,
        userResolver,
        text,
        message.channelId,
        message.id,
        displayName,
        message.author.id,
        isDM,
        reply
      );
    } catch (error) {
      console.error('Error routing message:', error);
      await message.reply('Failed to route message. Please try again.').catch(() => {});
    }
  });
}
