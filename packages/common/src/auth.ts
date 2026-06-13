/**
 * Admin Auth Middleware (shared)
 *
 * Express middleware enforcing ADMIN_TOKEN bearer auth for management / DM API
 * routes. Extracted to kill the copy-paste bypass-when-empty vulnerability
 * (CLAUDE.md bugs #1 / #6): the factory throws at construction if the token is
 * empty/blank, so a misconfigured gateway fails CLOSED at startup rather than
 * silently running with an open `/api/*` surface.
 */

import type { Request, Response, NextFunction } from 'express';
import { timingSafeEqual } from 'crypto';

/**
 * Build an Express middleware that requires `Authorization: Bearer <adminToken>`.
 *
 * Fail-closed by construction: throws immediately if `adminToken` is empty or
 * whitespace-only — there is no bypass-when-unset branch. The request-time check
 * is a BYTE-length-guarded `timingSafeEqual`: `Buffer.from` runs FIRST, then the
 * guard compares byte lengths (`timingSafeEqual` throws on unequal-length
 * buffers). Guarding on `String.length` instead would let a multibyte header
 * with a matching character count but a differing byte count slip past the guard
 * and crash `timingSafeEqual` with a TypeError (-> 500).
 */
export function createAuthMiddleware(adminToken: string) {
  if (!adminToken || adminToken.trim() === '') {
    throw new Error('createAuthMiddleware: ADMIN_TOKEN is required and cannot be empty.');
  }

  const expected = `Bearer ${adminToken}`;
  const expectedBuf = Buffer.from(expected);

  return (req: Request, res: Response, next: NextFunction) => {
    const auth = req.headers.authorization || '';
    const authBuf = Buffer.from(auth);

    if (authBuf.length === expectedBuf.length && timingSafeEqual(authBuf, expectedBuf)) {
      return next();
    }
    res.status(401).json({ error: 'Unauthorized' });
  };
}
