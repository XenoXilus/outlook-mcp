/**
 * Allowlisted billing-PDF fetch (FR-2, NFR-2).
 *
 * HTTPS-only, allowlist-only (including every redirect hop), Content-Type +
 * %PDF magic-byte validated, size- and time-bounded. Fetched bytes are never
 * executed or eval'd — they are returned as a Buffer for saving only.
 */

import { isPdfBuffer } from '../../utils/receiptFiles.js';

const DEFAULT_ALLOWLIST = ['pay.stripe.com', 'invoice.stripe.com', 'files.stripe.com', 'm.stripe.network'];
const REDIRECT_STATUSES = [301, 302, 303, 307, 308];

export function getBillingAllowlist() {
  const raw = (process.env.BILLING_DOMAIN_ALLOWLIST || '').trim();
  if (!raw) return [...DEFAULT_ALLOWLIST];
  return raw.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
}

export function assertAllowlisted(urlString, allowlist) {
  let parsed;
  try {
    parsed = new URL(urlString);
  } catch {
    throw new Error(`Invalid URL: ${urlString}`);
  }
  if (parsed.protocol !== 'https:') {
    throw new Error(`Only https:// URLs may be fetched (got ${parsed.protocol}//)`);
  }
  const host = parsed.hostname.toLowerCase();
  if (!allowlist.includes(host)) {
    throw new Error(`Host '${host}' is not in BILLING_DOMAIN_ALLOWLIST (${allowlist.join(', ')})`);
  }
  if (parsed.port !== '' && parsed.port !== '443') {
    throw new Error(`Non-default port '${parsed.port}' is not allowed on allowlisted hosts (only 443/default HTTPS port)`);
  }
  return parsed;
}

export async function fetchPdf(urlString, options = {}) {
  const {
    maxBytes = 25 * 1024 * 1024,
    timeoutMs = 30000,
    maxRedirects = 5,
    allowlist = getBillingAllowlist(),
    fetchImpl = fetch
  } = options;

  let current = urlString;
  for (let hop = 0; hop <= maxRedirects; hop++) {
    assertAllowlisted(current, allowlist);

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    let response;
    try {
      response = await fetchImpl(current, { redirect: 'manual', signal: controller.signal });
    } catch (error) {
      clearTimeout(timer);
      throw new Error(`Fetch failed for ${current}: ${error.name === 'AbortError' ? `timed out after ${timeoutMs}ms` : error.message}`);
    }

    // Non-body paths: clear the timer immediately before continuing/throwing.
    if (REDIRECT_STATUSES.includes(response.status)) {
      clearTimeout(timer);
      const location = response.headers.get('location');
      if (!location) throw new Error(`Redirect (${response.status}) without a Location header`);
      current = new URL(location, current).toString();
      continue;
    }

    if (!response.ok) {
      clearTimeout(timer);
      throw new Error(`Fetch failed with HTTP ${response.status} for ${current}`);
    }

    const contentType = (response.headers.get('content-type') || '').split(';')[0].trim().toLowerCase();
    if (contentType !== 'application/pdf') {
      clearTimeout(timer);
      throw new Error(`Expected Content-Type application/pdf, got '${contentType || 'none'}'`);
    }

    const declared = parseInt(response.headers.get('content-length') || '0', 10);
    if (declared > maxBytes) {
      clearTimeout(timer);
      throw new Error(`PDF exceeds maximum size: ${declared} bytes > ${maxBytes} bytes`);
    }

    // Body-read phase: keep the timer live so a slow body is still bounded in
    // time.  The finally clears it once the body is fully consumed (or on any
    // error path — including the size-exceeded throw below).
    let buffer;
    try {
      if (response.body?.getReader) {
        // Incremental size enforcement: reject the moment we exceed maxBytes
        // rather than buffering the whole body first.
        const reader = response.body.getReader();
        const chunks = [];
        let totalBytes = 0;
        while (true) {
          let readResult;
          try {
            readResult = await reader.read();
          } catch (err) {
            if (err.name === 'AbortError') {
              throw new Error(`Fetch failed for ${current}: timed out after ${timeoutMs}ms`);
            }
            throw err;
          }
          const { done, value } = readResult;
          if (done) break;
          totalBytes += value.length;
          if (totalBytes > maxBytes) {
            controller.abort();
            try { reader.cancel(); } catch (_) { /* ignore */ }
            throw new Error(`PDF exceeds maximum size: ${totalBytes} bytes > ${maxBytes} bytes`);
          }
          chunks.push(value);
        }
        buffer = Buffer.concat(chunks.map(c => Buffer.from(c)));
      } else {
        // Fallback for environments / mocks that expose only arrayBuffer().
        // The timer is still live, so a slow body will be aborted.
        let ab;
        try {
          ab = await response.arrayBuffer();
        } catch (err) {
          if (err.name === 'AbortError') {
            throw new Error(`Fetch failed for ${current}: timed out after ${timeoutMs}ms`);
          }
          throw err;
        }
        buffer = Buffer.from(ab);
        if (buffer.length > maxBytes) {
          throw new Error(`PDF exceeds maximum size: ${buffer.length} bytes > ${maxBytes} bytes`);
        }
      }
    } finally {
      clearTimeout(timer);
    }

    if (!isPdfBuffer(buffer)) {
      throw new Error('Fetched content does not start with %PDF magic bytes');
    }

    return { buffer, finalUrl: current, contentType };
  }
  throw new Error(`Too many redirects (more than ${maxRedirects})`);
}
