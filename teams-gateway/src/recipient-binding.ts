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
 * Canonicalize a Teams account/user id for mention-equality comparison (#12 fix).
 * Teams encodes the bot's channel account as `28:<appId>` inside mention entities,
 * while the activity `recipient.id` may arrive bare (or vice versa), and the GUID
 * case can skew between the two — a strict `===` then misses a real @mention and
 * the message is SILENTLY DROPPED. Strip a leading `28:` channel-account prefix and
 * case-fold so both sides reduce to one canonical id. Empty/undefined => `''`, which
 * never matches a real id (keeps the gate fail-closed).
 */
export function normalizeTeamsId(id: string | undefined): string {
  if (!id) return '';
  return id.replace(/^28:/, '').toLowerCase();
}

/**
 * True when THIS bot is @mentioned in the activity (#12). Teams encodes mentions as
 * `entities[]` of `{ type:'mention', mentioned:{ id } }`; the bot is addressed when a
 * mention's normalized id equals the normalized activity recipient (this bot's) id.
 * Typed structurally so this stays decoupled from the SDK's churn-prone activity type.
 * Conservative/fail-closed: no recipient id or no mention entities => not mentioned
 * (the gate then drops a non-personal message).
 */
export function botWasMentioned(activity: {
  entities?: unknown;
  recipient?: { id?: string };
}): boolean {
  const recipientId = normalizeTeamsId(activity.recipient?.id);
  if (!recipientId || !Array.isArray(activity.entities)) return false;
  return activity.entities.some((e) => {
    const ent = e as { type?: string; mentioned?: { id?: string } };
    return ent?.type === 'mention' && normalizeTeamsId(ent.mentioned?.id) === recipientId;
  });
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
