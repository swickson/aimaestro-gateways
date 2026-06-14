/**
 * Teams gateway — inbound attachment BYTE download (w3 follow-up fix).
 *
 * The connector-token authentication + 401-retry policy for pulling an inbound
 * Teams attachment's bytes, factored OUT of `server.ts` so it is unit-testable
 * with an injected token-getter + mock `fetch` (no live Azure). `server.ts` stays
 * the only SDK-coupled file: it builds the `getConnectorToken` closure from the
 * bot's `App` (`app.tokenManager.getBotToken()`) and forwards to here.
 *
 * Two inbound URL classes (resolved upstream by `extractInboundAttachments`):
 *   - `downloadUrl` — Teams `file.download.info`; PRE-AUTHENTICATED → bare GET.
 *   - `contentUrl`  — inline image/content served by the Bot Framework CONNECTOR
 *     (`smba.trafficmanager.net/.../v3/attachments`); a bare GET 401s. Needs the
 *     bot's CONNECTOR access token (audience `api.botframework.com`) as Bearer.
 *   - `data:` URIs decode locally — no network.
 *
 * Strategy (detect-connector-then-auth + safety net): a `contentUrl` (connector)
 * gets the Bearer token UP FRONT; a `downloadUrl` (pre-auth) goes bare. EITHER, if
 * the first attempt 401/403s and a token has not yet been tried, retries ONCE with
 * the token. FAIL-OPEN: if the token mint returns null or the authed attempt still
 * fails, this throws and `ingestAttachments` drops THAT attachment (the text still
 * routes), logged loudly. `redirect: 'error'` on the connector GET — a connector
 * `/v3/attachments` GET must not redirect (parity with the outbound hardening).
 */

import type { DownloadAttachment, RawInboundAttachment } from './attachments-inbound.js';

/** Mints this bot's connector (Bot Framework) access token, or null if unavailable. */
export type ConnectorTokenGetter = () => Promise<string | null>;

export interface AttachmentDownloadDeps {
  /** Bound in `server.ts` to `app.tokenManager.getBotToken()` (fail-open → null). */
  getConnectorToken: ConnectorTokenGetter;
  /** Per-leg timeout (ms). */
  timeoutMs: number;
  /** For `[TEAMS] (slug)` log prefixing. */
  slug: string;
}

/**
 * Decode a `data:` URI into bytes (inline attachment fast-path — no network).
 * Returns null if `value` is not a base64 data URI.
 */
function decodeDataUri(value: string): Uint8Array | null {
  const match = /^data:[^;,]*;base64,(.*)$/s.exec(value);
  if (!match) return null;
  return new Uint8Array(Buffer.from(match[1], 'base64'));
}

/**
 * Download one inbound attachment's bytes, authenticating connector `contentUrl`s
 * with the bot's connector token. Throws on any failure (caller fail-opens).
 */
export async function downloadAttachmentBytes(
  attachment: RawInboundAttachment,
  deps: AttachmentDownloadDeps,
): Promise<Uint8Array> {
  // A pre-authenticated Teams file URL takes precedence; a connector contentUrl
  // (no downloadUrl) is the auth-required path.
  const preAuthUrl = attachment.downloadUrl;
  const url = preAuthUrl ?? attachment.contentUrl;
  if (!url) throw new Error('attachment has no downloadUrl/contentUrl');

  const inline = decodeDataUri(url);
  if (inline) return inline;

  const isConnector = !preAuthUrl;

  let token: string | null = null;
  let triedAuth = false;

  const doFetch = (bearer: string | null): Promise<Response> =>
    fetch(url, {
      redirect: 'error',
      signal: AbortSignal.timeout(deps.timeoutMs),
      ...(bearer ? { headers: { Authorization: `Bearer ${bearer}` } } : {}),
    });

  // First attempt: connector content gets the token up front; pre-auth goes bare.
  if (isConnector) {
    token = await deps.getConnectorToken();
    triedAuth = token !== null;
  }
  let res = await doFetch(isConnector ? token : null);

  // Safety net: a denied first attempt we have not yet authed → retry once with the
  // token (covers a pre-auth URL that unexpectedly requires auth).
  if ((res.status === 401 || res.status === 403) && !triedAuth) {
    token = await deps.getConnectorToken();
    if (token) {
      triedAuth = true;
      res = await doFetch(token);
    }
  }

  if (!res.ok) {
    if (res.status === 401 || res.status === 403) {
      console.error(
        `[TEAMS] (${deps.slug}) attachment download ${res.status} (connector-token path) for '${attachment.name}' ` +
          `(auth ${triedAuth ? 'attempted' : 'unavailable'}).`,
      );
    }
    throw new Error(`download failed (${res.status})`);
  }
  return new Uint8Array(await res.arrayBuffer());
}

/**
 * Build the per-bot `DownloadAttachment` closure used by the inbound pipeline. The
 * SDK-coupled token source is injected as `getConnectorToken` from `server.ts`.
 */
export function buildAttachmentDownloader(
  slug: string,
  getConnectorToken: ConnectorTokenGetter,
  timeoutMs: number,
): DownloadAttachment {
  return (attachment: RawInboundAttachment): Promise<Uint8Array> =>
    downloadAttachmentBytes(attachment, { getConnectorToken, timeoutMs, slug });
}
