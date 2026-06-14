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
 *
 * BODY-TRANSFER RESILIENCE (w3 follow-up): the live root cause of an intermittent
 * inbound-image DROP was the Bot Framework connector intermittently STALLING the
 * response BODY — headers return fast, then the byte stream hangs. A single
 * whole-request timeout cannot tell a hung body from a slow-but-progressing one,
 * and it burns the full budget before failing open. So the body is read as a
 * STREAM with a per-chunk IDLE (stall) timeout: the timer resets on every chunk,
 * aborting only on a true stall (no bytes for `bodyStallMs`), never on a slow
 * transfer. A stall (or any transport error, header- or body-phase) RETRIES ONCE
 * — an intermittent stall usually clears on a fresh connection, converting a drop
 * into a delivered image — then fails open. The fast path adds ZERO latency.
 */

import type { DownloadAttachment, RawInboundAttachment } from './attachments-inbound.js';

/** Mints this bot's connector (Bot Framework) access token, or null if unavailable. */
export type ConnectorTokenGetter = () => Promise<string | null>;

export interface AttachmentDownloadDeps {
  /** Bound in `server.ts` to `app.tokenManager.getBotToken()` (fail-open → null). */
  getConnectorToken: ConnectorTokenGetter;
  /** Header/response-arrival timeout (ms) — caps how long we wait for headers. */
  timeoutMs: number;
  /** Per-chunk body IDLE (stall) timeout (ms) — abort if no bytes arrive for this long. */
  bodyStallMs: number;
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

/** Host of the chosen URL only — never the path/query/signature (for the telemetry line). */
function safeHost(url: string): string {
  try {
    return new URL(url).host;
  } catch {
    return '?';
  }
}

/**
 * Classify a thrown `fetch`/body rejection for the telemetry line. An aborted
 * header timer or a stalled body surfaces as a `TimeoutError` — its own `TIMEOUT`
 * outcome, distinct from any other transport `error(...)`. No URL/sig/token here.
 */
function classifyError(err: unknown): string {
  const name = err instanceof Error ? err.name : '';
  const msg = err instanceof Error ? err.message : String(err);
  if (name === 'TimeoutError' || /timeout|abort|stall/i.test(msg)) return 'TIMEOUT';
  return `error(${msg})`;
}

/**
 * Cancel an un-consumed `Response` body so an abandoned post-headers attempt does
 * NOT leak the socket to undici/GC. Called on EVERY path that drops a `Response`
 * after its headers arrive but WITHOUT handing the body to
 * `readBodyWithStallTimeout` — i.e. the 401/403 token-retry `continue` and the
 * non-ok fail-open `throw`. Best-effort: a `cancel()` rejection (already-errored /
 * locked body) is swallowed — there is nothing left to release in that case. The
 * happy path (body consumed) and the stall path (already abort()+cancel()'d) never
 * route here.
 */
async function discardBody(res: Response): Promise<void> {
  try {
    await res.body?.cancel();
  } catch {
    // body already errored or locked — socket teardown is out of our hands
  }
}

/**
 * Fetch a URL's HEADERS under a header-arrival timeout, returning the `Response`
 * (body not yet consumed) and the live `AbortController`. The header timer is
 * cleared the moment headers arrive — the controller then governs the BODY read
 * (the stall-timeout calls `controller.abort()` to tear the socket down).
 */
async function fetchHeaders(
  url: string,
  bearer: string | null,
  headerTimeoutMs: number,
): Promise<{ res: Response; controller: AbortController }> {
  const controller = new AbortController();
  // NOT unref'd: this timer is the load-bearing header-stall detector — it must
  // fire even in an otherwise-quiescent loop. It is always cleared below.
  const headerTimer = setTimeout(
    () => controller.abort(new DOMException('attachment header timeout', 'TimeoutError')),
    headerTimeoutMs,
  );
  try {
    const res = await fetch(url, {
      redirect: 'error',
      signal: controller.signal,
      ...(bearer ? { headers: { Authorization: `Bearer ${bearer}` } } : {}),
    });
    return { res, controller };
  } finally {
    clearTimeout(headerTimer);
  }
}

/**
 * Read a response body to completion, aborting if it STALLS — no bytes for
 * `stallMs`. A per-chunk idle timer (reset on every chunk) distinguishes a true
 * hang from a slow-but-progressing transfer. On stall (or any read error) we
 * `controller.abort()` to release the socket, then throw (caller retries once,
 * then fails open). Falls back to `arrayBuffer()` if the body is not streamable.
 */
async function readBodyWithStallTimeout(
  res: Response,
  controller: AbortController,
  stallMs: number,
): Promise<Uint8Array> {
  if (!res.body) return new Uint8Array(await res.arrayBuffer());

  const reader = res.body.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  try {
    for (;;) {
      let timer: ReturnType<typeof setTimeout> | undefined;
      let resolveIdle: (() => void) | undefined;
      const idle = new Promise<never>((resolve, reject) => {
        resolveIdle = resolve as () => void;
        // NOT unref'd: the body-stall detector must fire even in a quiescent loop;
        // it is always cleared on the next line's `finally`.
        timer = setTimeout(
          () => reject(new DOMException('attachment body stalled', 'TimeoutError')),
          stallMs,
        );
      });
      let chunk: ReadableStreamReadResult<Uint8Array>;
      try {
        chunk = await Promise.race([reader.read(), idle]);
      } finally {
        if (timer) clearTimeout(timer);
        // Settle the idle promise so a won race never leaves it lingering pending.
        resolveIdle?.();
      }
      if (chunk.done) break;
      if (chunk.value) {
        chunks.push(chunk.value);
        total += chunk.value.byteLength;
      }
    }
  } catch (err) {
    controller.abort();
    throw err;
  } finally {
    void reader.cancel().catch(() => undefined);
  }

  const out = new Uint8Array(total);
  let offset = 0;
  for (const c of chunks) {
    out.set(c, offset);
    offset += c.byteLength;
  }
  return out;
}

/**
 * Download one inbound attachment's bytes, authenticating connector `contentUrl`s
 * with the bot's connector token and tolerating an intermittent connector body
 * stall (stream-read + per-chunk stall-timeout + one bounded retry). Throws on any
 * unrecoverable failure (caller fail-opens). Emits ONE concise telemetry line.
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
  if (inline) {
    console.log(
      `[TEAMS] (${deps.slug}) attachment '${attachment.name}' path=data: host=- ` +
        `outcome=ok bytes=${inline.length} (no network)`,
    );
    return inline;
  }

  const isConnector = !preAuthUrl;
  const host = safeHost(url);

  // Telemetry accumulators — one always-on info line per attachment, emitted in
  // `finally` so it lands on the success AND the fail-open path. No url/sig/token.
  let tokenPresent = false;
  let mintMs = -1;
  let headerMs = -1;
  let bodyMs = -1;
  let outcome = 'n/a';
  let retries = 0;
  let bytesOut = -1;

  let token: string | null = null;
  let triedAuth = false;

  try {
    // Connector content gets the token up front; a pre-auth URL goes bare.
    if (isConnector) {
      const mintStart = Date.now();
      token = await deps.getConnectorToken();
      mintMs = Date.now() - mintStart;
      tokenPresent = token !== null;
      triedAuth = token !== null;
    }

    // Bounded attempt loop: the existing 401/403 auth net (mint + retry ONCE with
    // a token) plus ONE shared transient retry for a header/body timeout or stall.
    let transientRetried = false;
    for (;;) {
      let res: Response;
      let controller: AbortController;
      try {
        const headerStart = Date.now();
        ({ res, controller } = await fetchHeaders(
          url,
          isConnector || triedAuth ? token : null,
          deps.timeoutMs,
        ));
        headerMs = Date.now() - headerStart;
      } catch (err) {
        // Header-phase timeout / transport error → one bounded retry, else fail open.
        if (!transientRetried) {
          transientRetried = true;
          retries += 1;
          continue;
        }
        outcome = classifyError(err);
        throw err;
      }

      // Safety net: a denied attempt we have not yet authed → mint + retry once with
      // the token (covers a pre-auth URL that unexpectedly requires auth).
      if ((res.status === 401 || res.status === 403) && !triedAuth) {
        const mintStart = Date.now();
        token = await deps.getConnectorToken();
        mintMs = Date.now() - mintStart;
        tokenPresent = token !== null;
        if (token) {
          // Abandon this 401 Response → release its un-consumed body before retrying.
          await discardBody(res);
          triedAuth = true;
          retries += 1;
          continue;
        }
        // No token to add → fall through to the fail-open below.
      }

      if (!res.ok) {
        if (res.status === 401 || res.status === 403) {
          console.error(
            `[TEAMS] (${deps.slug}) attachment download ${res.status} (connector-token path) for '${attachment.name}' ` +
              `(auth ${triedAuth ? 'attempted' : 'unavailable'}).`,
          );
        }
        outcome = `status ${res.status}`;
        // Fail open → release this non-ok Response's un-consumed body before throwing.
        await discardBody(res);
        throw new Error(`download failed (${res.status})`);
      }

      // Headers OK → stream the body under the per-chunk stall timeout.
      try {
        const bodyStart = Date.now();
        const out = await readBodyWithStallTimeout(res, controller, deps.bodyStallMs);
        bodyMs = Date.now() - bodyStart;
        outcome = 'ok';
        bytesOut = out.byteLength;
        return out;
      } catch (err) {
        // Body stall / mid-stream error → one bounded retry (a fresh connection
        // usually clears an intermittent stall), else fail open.
        if (!transientRetried) {
          transientRetried = true;
          retries += 1;
          continue;
        }
        outcome = classifyError(err);
        throw err;
      }
    }
  } finally {
    console.log(
      `[TEAMS] (${deps.slug}) attachment '${attachment.name}' ` +
        `path=${isConnector ? 'connector' : 'pre-auth'} host=${host} ` +
        `tokenPresent=${tokenPresent ? 'y' : 'n'} mintMs=${mintMs} headerMs=${headerMs} bodyMs=${bodyMs} ` +
        `outcome=${outcome} retries=${retries}` +
        (bytesOut >= 0 ? ` bytes=${bytesOut}` : ''),
    );
  }
}

/**
 * Build the per-bot `DownloadAttachment` closure used by the inbound pipeline. The
 * SDK-coupled token source is injected as `getConnectorToken` from `server.ts`.
 */
export function buildAttachmentDownloader(
  slug: string,
  getConnectorToken: ConnectorTokenGetter,
  timeoutMs: number,
  bodyStallMs: number,
): DownloadAttachment {
  return (attachment: RawInboundAttachment): Promise<Uint8Array> =>
    downloadAttachmentBytes(attachment, { getConnectorToken, timeoutMs, bodyStallMs, slug });
}
