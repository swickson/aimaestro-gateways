/**
 * Teams gateway — inbound attachment ingestion (w3).
 *
 * Runs Maestro's LOCKED attachment upload flow once per Teams attachment and
 * returns the post-confirm `AMPAttachmentV1` descriptors for `inbound.ts` to cite
 * in `payload.attachments` on `/api/v1/route`. This runs OFF the ack-fast path
 * (`handleInbound` is fire-and-forget), so the multi-call round-trip never blocks
 * the Teams 200.
 *
 * SDK-DECOUPLED: byte download is an injected `downloadAttachment` closure
 * (server.ts binds it to the bot connector / pre-auth Teams URL), so this whole
 * module is unit-testable with a mock downloader + mock `fetch` — no live Azure,
 * no live Maestro (PLAN: "testable WITHOUT a live Azure endpoint").
 *
 * Per-attachment sequence (contract resolved from ai-maestro source, locked by
 * redacted-2 — see ORCHESTRATOR_PLAN w3 contract block):
 *   1. download bytes (closure)                      -> exact byteLength (declared size)
 *   2. POST /api/v1/attachments/upload   (bot Bearer) -> { attachment_id, upload_url }
 *   3. PUT  <upload_url>                  (NO Bearer)  -> body length MUST == declared size
 *   4. POST /api/v1/attachments/:id/confirm (bot Bearer) -> pending -> basic_clean (422 on fail)
 *   5. GET  /api/v1/attachments/:id/status  (bot Bearer) -> { url, sanitized filename, ... }
 *   6. assemble AMPAttachmentV1 from /status (the ONLY call that yields url + filename)
 *
 * FAIL-OPEN (NON-NEGOTIABLE): any leg that throws drops THAT attachment with a loud
 * `[TEAMS]`/`[AMP]` log and the others continue; the caller still routes the text.
 * A failed attachment leg must NEVER drop the message.
 */

import { createHash } from 'node:crypto';
import type { AMPAttachmentV1 } from './types.js';
import type { AttachmentPolicy } from './types.js';

/**
 * SDK-decoupled inbound attachment, as extracted by `server.ts` from the Teams
 * activity. `contentType` is best-effort (real MIME for inline images; for Teams
 * file sends, derived from the file name/type) — it is ADVISORY: Maestro sniffs
 * magic bytes at confirm and is authoritative. `downloadUrl` (Teams
 * file.download.info, pre-authenticated) or `contentUrl` (inline, connector-token)
 * is resolved by the injected downloader, not here.
 */
export interface RawInboundAttachment {
  name: string;
  contentType: string;
  downloadUrl?: string;
  contentUrl?: string;
}

/** Fetch the raw bytes for one attachment (bound to the bot connector in server.ts). */
export type DownloadAttachment = (attachment: RawInboundAttachment) => Promise<Uint8Array>;

export interface IngestDeps {
  maestroUrl: string;
  /** This bot's AMP api key — the Bearer for upload/confirm/status (same identity as /route). */
  ampApiKey: string;
  /** For `[TEAMS] (slug)` log prefixing. */
  botSlug: string;
  policy: AttachmentPolicy;
  downloadAttachment: DownloadAttachment;
  /** Per-call timeout for each Maestro/Teams HTTP leg. */
  timeoutMs: number;
}

export interface IngestResult {
  /** Successfully uploaded + confirmed + status-resolved descriptors, in input order. */
  attachments: AMPAttachmentV1[];
  /** Attachments that failed a leg (download/upload/PUT/confirm/status) — fail-open dropped. */
  failed: number;
  /** Attachments skipped by gateway policy (deny content-type, size cap, count cap). */
  skipped: number;
}

interface UploadResponse {
  attachment_id: string;
  upload_url: string;
}

interface StatusResponse {
  attachment_id: string;
  filename: string;
  content_type: string;
  size: number;
  digest: string;
  scan_status: AMPAttachmentV1['scan_status'];
  uploaded_at: string;
  expires_at: string;
  /** Signed download url; `null` until scan_status is routable (clean/basic_clean). */
  url: string | null;
}

/** Resolve a possibly-relative upload URL against the Maestro base. */
function resolveUploadUrl(uploadUrl: string, maestroUrl: string): string {
  if (/^https?:\/\//i.test(uploadUrl)) return uploadUrl;
  return new URL(uploadUrl, maestroUrl).toString();
}

/** Advisory content type for the upload body — only used if it looks like a MIME. */
function declaredContentType(raw: string): string {
  return raw && raw.includes('/') ? raw : 'application/octet-stream';
}

function isRoutableScanStatus(status: AMPAttachmentV1['scan_status']): boolean {
  return status === 'clean' || status === 'basic_clean';
}

function normalizeDescriptorDigest(digest: string): string {
  // Case-insensitive prefix check to match outbound isValidDigest (`/^sha256:/i`).
  // Guards against a `SHA256:`-prefixed input being re-prefixed to `sha256:SHA256:…`.
  // No effect today (Maestro /status returns bare lowercase hex per contract). (#23)
  return /^sha256:/i.test(digest) ? digest : `sha256:${digest}`;
}

/**
 * Ingest all attachments on one inbound Teams message. Enforces the gateway count
 * cap (extras dropped + logged), per-attachment deny-list + size cap (fail-fast),
 * then runs the Maestro flow. Always resolves — never throws — so the caller's
 * route is never blocked by an attachment failure.
 */
export async function ingestAttachments(
  rawAttachments: RawInboundAttachment[],
  deps: IngestDeps,
): Promise<IngestResult> {
  const log = (msg: string) => console.log(`[TEAMS] (${deps.botSlug}) ${msg}`);
  const err = (msg: string) => console.error(`[AMP] (${deps.botSlug}) ${msg}`);

  const result: IngestResult = { attachments: [], failed: 0, skipped: 0 };
  if (rawAttachments.length === 0) return result;

  // Count cap: process at most maxCount; drop + log the rest (never silent).
  let toProcess = rawAttachments;
  if (rawAttachments.length > deps.policy.maxCount) {
    const dropped = rawAttachments.length - deps.policy.maxCount;
    log(`attachment count ${rawAttachments.length} exceeds cap ${deps.policy.maxCount} — dropping ${dropped} (kept first ${deps.policy.maxCount}).`);
    result.skipped += dropped;
    toProcess = rawAttachments.slice(0, deps.policy.maxCount);
  }

  for (const att of toProcess) {
    // Deny-list (defense-in-depth; Maestro magic-byte sniff at confirm is authoritative).
    const ct = (att.contentType || '').toLowerCase();
    if (deps.policy.denyContentTypes.some((deny) => ct.includes(deny))) {
      log(`attachment '${att.name}' content-type '${att.contentType}' is deny-listed — skipping.`);
      result.skipped += 1;
      continue;
    }

    // 1. Download bytes — the exact byteLength is the size we declare at /upload, so
    //    the PUT body length matches (Maestro 400s on a mismatch). A download failure
    //    is fail-OPEN: drop this attachment, keep going.
    let bytes: Uint8Array;
    try {
      bytes = await deps.downloadAttachment(att);
    } catch (e) {
      result.failed += 1;
      err(`attachment '${att.name}' download failed — dropping (text still routes): ${(e as Error).message}`);
      continue;
    }

    // Authoritative size cap (the real byte count — the load-bearing gate before
    // Maestro's own 25MB). Over-cap / empty are clean policy SKIPS, not failures.
    if (bytes.byteLength > deps.policy.maxBytes) {
      log(`attachment '${att.name}' size ${bytes.byteLength}B exceeds cap ${deps.policy.maxBytes}B — skipping.`);
      result.skipped += 1;
      continue;
    }
    if (bytes.byteLength === 0) {
      log(`attachment '${att.name}' is empty (0 bytes) — skipping.`);
      result.skipped += 1;
      continue;
    }

    try {
      const descriptor = await uploadAttachment(att, bytes, deps);
      result.attachments.push(descriptor);
      log(`attachment '${att.name}' ingested -> ${descriptor.id} (${descriptor.size}B, ${descriptor.scan_status}).`);
    } catch (e) {
      // FAIL-OPEN: drop this attachment, keep going, never throw up to the route.
      result.failed += 1;
      err(`attachment '${att.name}' failed to ingest — dropping (text still routes): ${(e as Error).message}`);
    }
  }

  return result;
}

/**
 * Run Maestro's upload -> PUT -> confirm -> status flow for one already-downloaded
 * attachment and assemble the LOCKED descriptor. Throws on any leg failure (the
 * caller fail-opens).
 */
async function uploadAttachment(att: RawInboundAttachment, bytes: Uint8Array, deps: IngestDeps): Promise<AMPAttachmentV1> {
  const size = bytes.byteLength;
  const digest = createHash('sha256').update(bytes).digest('hex');
  const contentType = declaredContentType(att.contentType);

  // 2. POST /upload — reserve an id + signed upload url. Declared size MUST be the
  //    real byte length (Q1 caveat). Advisory digest sent for cross-check.
  const upload = await postJson<UploadResponse>(
    `${deps.maestroUrl}/api/v1/attachments/upload`,
    deps.ampApiKey,
    { filename: att.name, content_type: contentType, size, digest },
    deps.timeoutMs,
  );
  if (!upload.attachment_id || !upload.upload_url) {
    throw new Error('upload response missing attachment_id/upload_url');
  }

  // 3. PUT raw bytes to the signed upload url (NO Bearer — the url is the auth).
  const putUrl = resolveUploadUrl(upload.upload_url, deps.maestroUrl);
  const putRes = await fetch(putUrl, {
    method: 'PUT',
    headers: { 'Content-Type': contentType },
    // Wrap in a Blob: a BodyInit whose byte length is exactly `size` (Maestro 400s
    // on a length mismatch). The `as BlobPart` cast sidesteps the TS5.7
    // Uint8Array<ArrayBufferLike> vs <ArrayBuffer> lib friction; the bytes (and
    // their length) are unchanged at runtime.
    body: new Blob([bytes as BlobPart]),
    signal: AbortSignal.timeout(deps.timeoutMs),
  });
  if (!putRes.ok) {
    throw new Error(`PUT bytes failed (${putRes.status}): ${await safeBody(putRes)}`);
  }

  // 4. POST /confirm — TRANSITIONS pending -> basic_clean; re-verifies size + digest
  //    + magic-byte<->content_type. 422 = rejected (MUST-fail).
  const confirmRes = await fetch(`${deps.maestroUrl}/api/v1/attachments/${upload.attachment_id}/confirm`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${deps.ampApiKey}` },
    signal: AbortSignal.timeout(deps.timeoutMs),
  });
  if (!confirmRes.ok) {
    throw new Error(`confirm failed (${confirmRes.status}): ${await safeBody(confirmRes)}`);
  }

  // 5. GET /status — the ONLY call that yields the signed `url` + server-sanitized
  //    `filename`. url is present only once scan_status is clean/basic_clean.
  const status = await getJson<StatusResponse>(
    `${deps.maestroUrl}/api/v1/attachments/${upload.attachment_id}/status`,
    deps.ampApiKey,
    deps.timeoutMs,
  );
  if (!status.url) {
    throw new Error(`status returned no signed url (scan_status=${status.scan_status})`);
  }
  if (!isRoutableScanStatus(status.scan_status)) {
    throw new Error(`status scan_status=${status.scan_status} is not routable`);
  }

  // 6. Assemble the LOCKED wire descriptor, server-authoritative fields from /status.
  return {
    kind: 'amp-v1',
    id: status.attachment_id,
    filename: status.filename,
    content_type: status.content_type,
    size: status.size,
    digest: normalizeDescriptorDigest(status.digest),
    url: status.url,
    scan_status: status.scan_status,
    uploaded_at: status.uploaded_at,
    expires_at: status.expires_at,
  };
}

async function postJson<T>(url: string, apiKey: string, body: unknown, timeoutMs: number): Promise<T> {
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${apiKey}` },
    body: JSON.stringify(body),
    signal: AbortSignal.timeout(timeoutMs),
  });
  if (!res.ok) {
    throw new Error(`POST ${new URL(url).pathname} failed (${res.status}): ${await safeBody(res)}`);
  }
  return (await res.json()) as T;
}

async function getJson<T>(url: string, apiKey: string, timeoutMs: number): Promise<T> {
  const res = await fetch(url, {
    method: 'GET',
    headers: { Authorization: `Bearer ${apiKey}` },
    signal: AbortSignal.timeout(timeoutMs),
  });
  if (!res.ok) {
    throw new Error(`GET ${new URL(url).pathname} failed (${res.status}): ${await safeBody(res)}`);
  }
  return (await res.json()) as T;
}

async function safeBody(res: Response): Promise<string> {
  return res.text().catch(() => '');
}
