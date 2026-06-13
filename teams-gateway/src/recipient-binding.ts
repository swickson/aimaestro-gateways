/**
 * Recipient-identity binding check (PLAN §0.2 defense-in-depth). Returns whether
 * the inbound activity's `recipient.id` belongs to THIS bot.
 *
 * The adapter's Bot-Service-JWT audience validation (per-bot `clientId`) is the
 * PRIMARY control; this guard only ADDS a cheap pre-resolve rejection — it never
 * weakens the adapter's gate. Teams sends the bot's channel-account id as
 * `28:<appId>`, so both that prefixed form and the bare `appId` are accepted. An
 * ABSENT recipient id passes through (we defer to the adapter rather than reject a
 * shape we can't authoritatively judge).
 */
export function recipientMatchesBot(recipientId: string | undefined, appId: string): boolean {
  if (!recipientId) return true;
  return recipientId === appId || recipientId === `28:${appId}`;
}

/**
 * Reject a message activity whose recipient account belongs to a different bot.
 * Returns true when the caller must stop before user resolution / routing.
 */
export function rejectMismatchedRecipient(options: {
  recipientId: string | undefined;
  appId: string;
  slug: string;
  activityId?: string;
  warn?: (...args: unknown[]) => void;
}): boolean {
  if (recipientMatchesBot(options.recipientId, options.appId)) return false;
  const warn = options.warn ?? console.warn;
  warn(
    `[TEAMS] (${options.slug}) rejecting activity ${options.activityId ?? '(no id)'} — recipient.id '${options.recipientId}' does not match bot appId '${options.appId}'.`,
  );
  return true;
}
