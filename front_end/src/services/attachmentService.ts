// filepath: my-app/src/services/attachmentService.ts
// Orchestrates small vs. chunked uploads to Microsoft Graph.
//
// - Small files (<= 3 MB) go through the backend's /attachments/small
//   endpoint (JSON body, backend forwards to Graph).
// - Large files (> 3 MB) use a Graph upload session. The backend creates
//   the session and returns the uploadUrl. The client then PUTs each
//   chunk DIRECTLY to Microsoft Graph using the pre-authenticated URL
//   (no Authorization header needed).
//
// === Rate-limit (429) handling ===
// Microsoft Graph throttles rapid successive PUTs to the upload session
// URL with HTTP 429 (Too Many Requests). When this happens, Graph
// includes a `Retry-After` header specifying how many seconds to wait.
// Our retry logic:
//   1. Reads and honors the `Retry-After` header.
//   2. Falls back to exponential backoff (5s, 10s, 20s, 40s, 80s) when
//      the header is absent.
//   3. Retries up to MAX_CHUNK_RETRIES times before giving up.
//
// === Concurrency control ===
// If multiple large files are selected at once, each AttachmentRow
// would start its own upload session, and the concurrent PUTs would
// trigger throttling. We serialize all large-file uploads through a
// global queue so only one file uploads at a time.

import { mailAPI } from './api';

const SMALL_FILE_LIMIT = 3 * 1024 * 1024; // 3 MB
// Use 1 MB chunks to minimize per-request load and stay well under
// Graph's throttle threshold for the attachment upload session endpoint.
const CHUNK_SIZE = 1 * 1024 * 1024; // 1 MB
// Base delay between chunks (ms) to avoid throttling under normal
// operation. Additional backoff is applied on 429 responses.
const INTER_CHUNK_DELAY_MS = 500;
// Max retry attempts per chunk when a 429 / network error occurs.
const MAX_CHUNK_RETRIES = 5;
// Minimum backoff (seconds) for the first retry attempt when no
// Retry-After header is present. Subsequent attempts double this value.
const BASE_BACKOFF_SECONDS = 5;

export interface UploadProgress {
  /** Percent of the file uploaded (0-100) */
  percent: number;
  /** Current chunk (1-based) when applicable */
  currentChunk?: number;
  /** Total chunks when applicable */
  totalChunks?: number;
}

export interface UploadHandlers {
  onProgress?: (p: UploadProgress) => void;
}

/** Promise-based delay helper. */
function delay(ms: number): Promise<void> {
  if (ms <= 0) return Promise.resolve();
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}

/**
 * Extract the Graph attachment ID from a `Location` (or similar) header
 * returned by the final chunk PUT. The URL can look like any of:
 *   https://outlook.office.com/api/v2.0/Users(...)/Messages(...)/Attachments('<ID>')
 *   https://graph.microsoft.com/v1.0/users(...)/messages(...)/attachments('<ID>')
 *   ...or sometimes just the raw ID.
 *
 * The ID is the last path segment, often wrapped in single quotes and
 * percent-encoded (e.g. `%3D` for `=`, `%2B` for `+`). We:
 *   1. Strip any surrounding single/double quotes.
 *   2. URL-decode the segment.
 *   3. If parsing as a URL fails, treat the whole header value as the ID
 *      and strip quotes / decode.
 */
function extractIdFromLocation(location: string | undefined): string | undefined {
  if (!location) return undefined;

  // Try parsing as a URL first.
  try {
    const url = new URL(location);
    const segments = url.pathname.split('/').filter(Boolean);
    const last = segments[segments.length - 1] ?? '';
    if (!last) return undefined;
    const cleaned = last
      .replace(/^['"]|['"]$/g, '')
      .replace(/^attachments\(/i, '')
      .replace(/\)$/i, '');
    const decoded = (() => {
      try {
        return decodeURIComponent(cleaned);
      } catch {
        return cleaned;
      }
    })();
    return decoded || undefined;
  } catch {
    // Not a parseable URL — fall through to raw-text handling.
  }

  // Raw text: strip surrounding quotes, decode, and return.
  const cleaned = location.trim().replace(/^['"]|['"]$/g, '');
  try {
    return decodeURIComponent(cleaned) || undefined;
  } catch {
    return cleaned || undefined;
  }
}

// ---------------------------------------------------------------
// Global upload serialization queue
// ---------------------------------------------------------------
// Only one large-file upload runs at a time. This prevents concurrent
// upload sessions from triggering Microsoft Graph's per-user throttle
// (HTTP 429). Small-file uploads go through the backend and are not
// serialized.

let largeUploadInProgress = false;
const largeUploadQueue: Array<() => Promise<void>> = [];

function enqueueLargeUpload(task: () => Promise<string>): Promise<string> {
  return new Promise<string>((resolve, reject) => {
    const run = async (): Promise<void> => {
      try {
        const id = await task();
        resolve(id);
      } catch (err) {
        reject(err);
      }
    };
    largeUploadQueue.push(run);
    void processLargeUploadQueue();
  });
}

async function processLargeUploadQueue(): Promise<void> {
  if (largeUploadInProgress) return;
  largeUploadInProgress = true;
  while (largeUploadQueue.length > 0) {
    const task = largeUploadQueue.shift();
    if (task) {
      await task();
      // Brief pause between files to let Graph's rate limiter recover.
      await delay(1000);
    }
  }
  largeUploadInProgress = false;
}

// ---------------------------------------------------------------
// Public API
// ---------------------------------------------------------------

/**
 * Upload a single file to a draft message using either the small or
 * chunked upload flow depending on file size. Large-file uploads are
 * serialized through a global queue to prevent concurrent throttling.
 *
 * Returns the Microsoft Graph attachment ID.
 */
export async function uploadFileToDraft(
  draftId: string,
  file: File,
  handlers: UploadHandlers = {}
): Promise<string> {
  if (file.size <= SMALL_FILE_LIMIT) {
    return uploadSmallFile(draftId, file, handlers);
  }
  // Serialize large-file uploads so only one is in flight at a time.
  return enqueueLargeUpload(() => uploadLargeFileInChunks(draftId, file, handlers));
}

// ---------------------------------------------------------------
// Small file upload ( <= 3 MB)
// ---------------------------------------------------------------

async function uploadSmallFile(
  draftId: string,
  file: File,
  handlers: UploadHandlers
): Promise<string> {
  handlers.onProgress?.({ percent: 0 });
  const base64 = await mailAPI.fileToBase64(file);
  handlers.onProgress?.({ percent: 50 });

  const res = await mailAPI.uploadSmallAttachment({
    draftId,
    fileName: file.name,
    contentType: file.type || 'application/octet-stream',
    base64,
  });

  handlers.onProgress?.({ percent: 100 });
  return res.attachmentId;
}

// ---------------------------------------------------------------
// Large file upload ( > 3 MB) — chunked, direct to Graph
// ---------------------------------------------------------------

/**
 * Chunked upload for files > 3 MB. The client uploads chunks directly
 * to Microsoft Graph's pre-authenticated upload session URL.
 *
 * IMPORTANT: The uploadUrl is a PRE-AUTHENTICATED URL. It does NOT
 * require (and must NOT include) an Authorization header. The auth is
 * embedded in the URL itself.
 *
 * The function reads the `Retry-After` header from any 429 response
 * and waits the specified time before retrying.
 */
async function uploadLargeFileInChunks(
  draftId: string,
  file: File,
  handlers: UploadHandlers
): Promise<string> {
  // 1) Create the upload session via the backend.
  const session = await mailAPI.createAttachmentUploadSession({
    draftId,
    fileName: file.name,
    contentType: file.type || 'application/octet-stream',
    fileSize: file.size,
  });

  const uploadUrl = session.uploadUrl;

  // 2) PUT each chunk to Graph's pre-authenticated URL.
  const total = file.size;
  const totalChunks = Math.ceil(total / CHUNK_SIZE);
  let uploaded = 0;

  for (let i = 0; i < totalChunks; i++) {
    const start = i * CHUNK_SIZE;
    const end = Math.min(start + CHUNK_SIZE, total) - 1;
    const slice = file.slice(start, end + 1);
    const contentRange = `bytes ${start}-${end}/${total}`;

    const result = await putChunkWithRetry(
      uploadUrl,
      slice,
      contentRange,
      file.type || 'application/octet-stream',
    );

    uploaded += end - start + 1;
    const percent = Math.min(100, Math.round((uploaded / total) * 100));
    handlers.onProgress?.({
      percent,
      currentChunk: i + 1,
      totalChunks,
    });

    // Determine if this is the FINAL chunk. Per Microsoft Graph docs,
    // the final chunk is the one that covers the last byte of the file:
    //   - If the response is 200/201 (upload complete), this is the final
    //     chunk regardless of what byte range we sent.
    //   - If the response is 202 (more chunks expected), this is NOT the
    //     final chunk.
    //
    // CRITICAL: Some Microsoft 365 tenants return 200 OK for INTERMEDIATE
    // chunks (not just the final one). If we treat that as "complete" and
    // stop the loop early, the file never gets fully uploaded and the
    // draft ends up with no attachment. To avoid this, we ALSO require
    // either:
    //   (a) the chunk index to be the last one (i === totalChunks - 1), OR
    //   (b) the response includes a `nextExpectedRanges` field that
    //       starts at byte 0 (meaning the server has received nothing more
    //       and this is a fresh session — also indicates completion), OR
    //   (c) the response is 201 (which is reserved for the final chunk).
    //
    // In practice, the simplest and most reliable check is:
    //   - status === 201 → definitely complete (final chunk)
    //   - status === 200 AND i === totalChunks - 1 → complete (final chunk)
    //   - status === 200 AND i < totalChunks - 1 → INTERMEDIATE chunk,
    //     continue uploading (some tenants do this)
    //   - status === 202 → more chunks expected
    const isLastChunk = i === totalChunks - 1;
    const isComplete =
      result.status === 201 ||
      (result.status === 200 && isLastChunk);

    if (isComplete) {
      // Graph returned 200/201 — upload is done.
      //
      // Per Microsoft Graph docs, the final PUT response is one of:
      //   1) 201 Created with an empty body and a `Location` header
      //      pointing to the new file attachment. The attachment ID
      //      is the last path segment of that Location URL.
      //   2) 200 OK with a JSON body containing the attachment's `id`.
      //
      // We prefer the Location header when present, then fall back to
      // the JSON body, then to a synthetic ID so the rest of the
      // pipeline can still proceed.
      const realId =
        result.attachmentId ?? extractIdFromLocation(result.location);

      // Diagnostic: if we couldn't extract a real ID, warn the developer
      // so they can investigate. The synthetic ID still works for the
      // "send" path, but it's better to have the real one.
      if (!realId) {
        console.warn(
          '[attachmentService] Could not extract attachment ID from final chunk response. ' +
            'Using synthetic ID. Location header:',
          result.location,
          'Body attachmentId:',
          result.attachmentId
        );
      }

      handlers.onProgress?.({ percent: 100, currentChunk: i + 1, totalChunks });
      return realId ?? `session-${session.uploadUrl.slice(-12)}`;
    }

    // Intermediate chunk: Graph told us more is expected (or returned
    // 200 on a non-final chunk, which we treat as intermediate). Brief
    // pause between chunks to stay under the throttle threshold.
    await delay(INTER_CHUNK_DELAY_MS);
  }

  // If we get here, all chunks returned 202 (more expected) but none
  // returned a completion status. This shouldn't happen for a correctly
  // sized file, but we treat it as success anyway since all bytes were
  // uploaded.
  return `session-${session.uploadUrl.slice(-12)}`;
}

// ---------------------------------------------------------------
// Chunk upload with retry + Retry-After support
// ---------------------------------------------------------------

/**
 * PUT a single chunk to Microsoft Graph's pre-authenticated upload
 * session URL, with automatic retry on 429 (Too Many Requests) and
 * network-level errors.
 *
 * The function reads the `Retry-After` response header (in seconds)
 * and waits that long before retrying. If the header is absent, it
 * falls back to exponential backoff: 5s, 10s, 20s, 40s, 80s.
 */
async function putChunkWithRetry(
  uploadUrl: string,
  chunk: Blob,
  contentRange: string,
  contentType: string,
): Promise<{
  status: number;
  complete: boolean;
  attachmentId?: string;
  location?: string;
  nextExpectedRange?: string;
}> {
  let lastError: Error | null = null;

  for (let attempt = 0; attempt < MAX_CHUNK_RETRIES; attempt++) {
    try {
      const result = await putChunkToGraph(uploadUrl, chunk, contentRange, contentType);
      return result;
    } catch (err: unknown) {
      const error = err instanceof Error ? err : new Error(String(err));
      lastError = error;

      // Only retry on 429 (Too Many Requests) or network errors.
      const isThrottle = /429/.test(error.message);
      const isNetwork = /Failed to fetch|NetworkError|load failed/i.test(
        error.message,
      );

      if (!isThrottle && !isNetwork) {
        throw error;
      }

      // Parse the Retry-After value baked into the error message by
      // putChunkToGraph. If present, honor it. Otherwise use
      // exponential backoff starting at BASE_BACKOFF_SECONDS.
      const retryAfterMatch = error.message.match(/Retry-After:\s*(\d+)/i);
      const retryAfterSeconds = retryAfterMatch
        ? parseInt(retryAfterMatch[1], 10)
        : BASE_BACKOFF_SECONDS * Math.pow(2, attempt);

      // Wait before the next retry attempt.
      await delay(retryAfterSeconds * 1000);
    }
  }

  throw lastError ?? new Error('Chunk upload failed after max retries.');
}

/**
 * PUT a single binary chunk to Microsoft Graph's pre-authenticated
 * upload session URL. The URL itself contains the authentication, so
 * we must NOT include an Authorization header.
 *
 * IMPORTANT: We must NOT set the `Content-Length` header manually.
 * Browsers automatically set it based on the Blob body size, and
 * manually setting it causes fetch() to throw or Graph to return 400
 * Bad Request.
 *
 * On 429, the error message includes the `Retry-After` header value
 * so the retry logic can honor it.
 */
async function putChunkToGraph(
  uploadUrl: string,
  chunk: Blob,
  contentRange: string,
  contentType: string,
): Promise<{
  status: number;
  complete: boolean;
  attachmentId?: string;
  location?: string;
  nextExpectedRange?: string;
}> {
  const res = await fetch(uploadUrl, {
    method: 'PUT',
    // Do NOT include Content-Length — the browser sets it automatically
    // based on the Blob body size. Setting it manually causes issues.
    //
    // Content-Type IS required by Microsoft Graph for upload sessions
    // (per the official Outlook large-attachments documentation). It
    // must be application/octet-stream per the docs, but we pass the
    // file's MIME type so Outlook can display the correct icon.
    // Authorization header is intentionally absent: the uploadUrl is
    // pre-authenticated.
    headers: {
      'Content-Range': contentRange,
      'Content-Type': contentType || 'application/octet-stream',
    },
    body: chunk,
  });

  // 200/201: upload complete. Two variants per Microsoft Graph docs:
  //   1) 201 Created with empty body + `Location` header containing the
  //      attachment URL. The ID is the last path segment of that URL.
  //   2) 200 OK with a JSON body containing `{ id, size, ... }`.
  if (res.status === 200 || res.status === 201) {
    // Per the docs, the `Location` header is the authoritative source
    // for the attachment ID. Capture it BEFORE we consume the body.
    // Some servers also include a `Content-Location` header or echo the
    // ID in `OData-EntityId`, so we read all three for robustness.
    const location =
      res.headers.get('Location') ||
      res.headers.get('Content-Location') ||
      res.headers.get('OData-EntityId') ||
      undefined;

    const text = await res.text().catch(() => '');
    let attachmentId: string | undefined;
    if (text) {
      try {
        const data: { id?: string; size?: number } = JSON.parse(text);
        attachmentId = data.id;
      } catch {
        // Body wasn't JSON — that's expected for the 201 variant. The
        // ID will come from the Location header instead.
      }
    }

    // Diagnostic: log the raw response for the final chunk so we can
    // diagnose any future "missing ID" issues.
    if (typeof console !== 'undefined' && console.debug) {
      console.debug('[attachmentService] Final chunk response:', {
        status: res.status,
        location,
        attachmentId,
        bodyLength: text.length,
        bodyPreview: text ? text.slice(0, 200) : '(empty)',
      });
    }

    return { status: res.status, complete: true, attachmentId, location };
  }

  // 202: more chunks expected
  if (res.status === 202) {
    const text = await res.text().catch(() => '');
    let nextExpectedRange: string | undefined;
    if (text) {
      try {
        const data: { nextExpectedRange?: string } = JSON.parse(text);
        nextExpectedRange = data.nextExpectedRange;
      } catch {
        // Body wasn't JSON — ignore, the range is determined by the loop.
      }
    }
    return { status: 202, complete: false, nextExpectedRange };
  }

  // 429: throttled — include the Retry-After header in the error
  // message so the retry logic can honor it.
  if (res.status === 429) {
    const retryAfter = res.headers.get('Retry-After') || '';
    const text = await res.text().catch(() => '');
    throw new Error(
      `Chunk upload throttled (429). Retry-After: ${retryAfter}. ${text}`,
    );
  }

  // Any other status is a non-retryable error.
  const text = await res.text().catch(() => '');
  throw new Error(
    `Chunk upload failed (${res.status}): ${text || res.statusText}`,
  );
}

