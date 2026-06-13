import assert from 'node:assert/strict';
import { describe, it } from 'node:test';

import { recipientMatchesBot, rejectMismatchedRecipient } from '../recipient-binding.js';

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
