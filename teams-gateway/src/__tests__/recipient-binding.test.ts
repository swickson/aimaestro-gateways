import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { botWasMentioned, normalizeTeamsId, recipientMatchesBot, rejectMismatchedRecipient } from '../recipient-binding.js';

describe('Teams recipient identity binding', () => {
  const appId = '00000000-0000-0000-0000-000000000000';

  it('accepts bare app id, Teams 28: app id, and absent recipient ids', () => {
    assert.equal(recipientMatchesBot(appId, appId), true);
    assert.equal(recipientMatchesBot(`28:${appId}`, appId), true);
    assert.equal(recipientMatchesBot(undefined, appId), true);
  });

  it('rejects a genuinely different bot recipient id', () => {
    assert.equal(recipientMatchesBot('11111111-1111-1111-1111-111111111111', appId), false);
    assert.equal(recipientMatchesBot('28:11111111-1111-1111-1111-111111111111', appId), false);
  });

  it('logs recipient mismatch at warn level and stops before the inbound pipeline', () => {
    const warnings: string[] = [];
    let proceeded = false;

    function runMessageGuard(recipientId: string | undefined): void {
      if (rejectMismatchedRecipient({
        recipientId,
        appId,
        slug: 'maestro',
        activityId: 'activity-1',
        warn: (...args: unknown[]) => {
          warnings.push(args.map(String).join(' '));
        },
      })) {
        return;
      }
      proceeded = true;
    }

    runMessageGuard('28:11111111-1111-1111-1111-111111111111');

    assert.equal(proceeded, false);
    assert.equal(warnings.length, 1);
    assert.match(warnings[0] ?? '', /^\[TEAMS\] \(maestro\) rejecting activity activity-1/);
    assert.match(warnings[0] ?? '', /recipient\.id '28:11111111-1111-1111-1111-111111111111' does not match bot appId/);
  });
});

describe('Teams @mention gate (#12 — botWasMentioned)', () => {
  const appId = '00000000-0000-0000-0000-000000000000';
  const mention = (mentionedId: string) => ({ type: 'mention', mentioned: { id: mentionedId } });

  it('normalizeTeamsId strips a leading 28: prefix and case-folds; empty => empty', () => {
    assert.equal(normalizeTeamsId(`28:${appId.toUpperCase()}`), appId);
    assert.equal(normalizeTeamsId(appId.toUpperCase()), appId);
    assert.equal(normalizeTeamsId(undefined), '');
    assert.equal(normalizeTeamsId(''), '');
  });

  it('matches a 28:-prefixed mention id against a bare recipient id (and vice versa)', () => {
    // Mention entity carries the 28:<appId> MRI form; recipient.id arrives bare.
    assert.equal(botWasMentioned({ recipient: { id: appId }, entities: [mention(`28:${appId}`)] }), true);
    // Inverse skew: recipient.id is 28:-prefixed, mention is bare.
    assert.equal(botWasMentioned({ recipient: { id: `28:${appId}` }, entities: [mention(appId)] }), true);
  });

  it('matches across case-mismatched ids', () => {
    assert.equal(
      botWasMentioned({ recipient: { id: appId.toLowerCase() }, entities: [mention(`28:${appId.toUpperCase()}`)] }),
      true,
    );
  });

  it('still drops a non-matching mention id (a different bot was mentioned)', () => {
    const otherBot = '11111111-1111-1111-1111-111111111111';
    assert.equal(botWasMentioned({ recipient: { id: appId }, entities: [mention(`28:${otherBot}`)] }), false);
  });

  it('is fail-closed: no recipient id, no entities, or no mention entity => not mentioned', () => {
    assert.equal(botWasMentioned({ recipient: { id: undefined }, entities: [mention(`28:${appId}`)] }), false);
    assert.equal(botWasMentioned({ recipient: { id: appId } }), false);
    assert.equal(botWasMentioned({ recipient: { id: appId }, entities: [] }), false);
    assert.equal(botWasMentioned({ recipient: { id: appId }, entities: [{ type: 'clientInfo' }] }), false);
  });
});
