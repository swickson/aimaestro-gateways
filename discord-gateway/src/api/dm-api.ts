/**
 * Discord Gateway - Outbound DM API
 *
 * POST /api/gateway/dm — Send a direct message to a Discord user.
 * Called by Maestro's /api/users/:id/notify endpoint for outbound routing.
 */

import { Router, Request, Response } from 'express';
import type { Client } from 'discord.js';
import { logEvent } from './activity-log.js';

/** Discord's max message length */
const DISCORD_MAX_LENGTH = 2000;

interface DMRequest {
  platformUserId: string;
  message: string;
  subject?: string;
}

interface DMSuccessResponse {
  success: true;
  messageId: string;
}

interface DMErrorResponse {
  success: false;
  error: string;
  reason: string;
}

export function createDMRouter(
  getClient: () => Client
): Router {
  const router = Router();

  router.post('/', async (req: Request, res: Response) => {
    const { platformUserId, message, subject } = req.body as DMRequest;

    // Validate required fields
    if (!platformUserId || typeof platformUserId !== 'string') {
      res.status(400).json({
        success: false,
        error: 'invalid_request',
        reason: 'platformUserId is required and must be a string',
      } satisfies DMErrorResponse);
      return;
    }

    if (!message || typeof message !== 'string') {
      res.status(400).json({
        success: false,
        error: 'invalid_request',
        reason: 'message is required and must be a string',
      } satisfies DMErrorResponse);
      return;
    }

    const client = getClient();

    if (!client.isReady()) {
      res.status(503).json({
        success: false,
        error: 'gateway_unavailable',
        reason: 'Discord client is not connected',
      } satisfies DMErrorResponse);
      return;
    }

    try {
      // Fetch the user — this validates the user ID exists
      const user = await client.users.fetch(platformUserId).catch(() => null);

      if (!user) {
        res.status(404).json({
          success: false,
          error: 'user_not_found',
          reason: `Discord user ${platformUserId} not found`,
        } satisfies DMErrorResponse);
        return;
      }

      // Check mutual guild membership (required for DMs)
      const mutualGuild = client.guilds.cache.find(
        guild => guild.members.cache.has(platformUserId)
      );

      if (!mutualGuild) {
        // Try fetching members in case the cache is incomplete
        let hasMutualGuild = false;
        for (const guild of client.guilds.cache.values()) {
          try {
            await guild.members.fetch(platformUserId);
            hasMutualGuild = true;
            break;
          } catch {
            // Not a member of this guild
          }
        }

        if (!hasMutualGuild) {
          res.status(422).json({
            success: false,
            error: 'no_mutual_guild',
            reason: `Cannot DM user ${user.tag} — no mutual guild membership with the bot`,
          } satisfies DMErrorResponse);
          return;
        }
      }

      // Open DM channel and send
      const dmChannel = await user.createDM();
      const truncatedMessage = message.length > DISCORD_MAX_LENGTH
        ? message.substring(0, DISCORD_MAX_LENGTH - 3) + '...'
        : message;

      const sent = await dmChannel.send(truncatedMessage);

      console.log(`[-> Discord DM] Message sent to ${user.tag} (${platformUserId})`);
      logEvent('outbound', `DM sent to ${user.tag}`, {
        to: user.tag,
        subject: subject || '',
        deliveryStatus: 'delivered',
      });

      res.json({
        success: true,
        messageId: sent.id,
      } satisfies DMSuccessResponse);
    } catch (error) {
      const errMsg = (error as Error).message;
      console.error(`[DM API] Failed to send DM to ${platformUserId}:`, errMsg);

      logEvent('error', `DM delivery failed to ${platformUserId}`, {
        error: errMsg,
      });

      res.status(500).json({
        success: false,
        error: 'delivery_failed',
        reason: errMsg,
      } satisfies DMErrorResponse);
    }
  });

  return router;
}
