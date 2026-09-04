/**
 * Fetch a billing PDF from an allowlisted URL in an email body (FR-2).
 *
 * Fallback for when a receipt message has no PDF attachment. Unauthenticated
 * HTTPS GET of pre-signed Stripe links only — see pdfFetch.js for the
 * egress-security rules (NFR-2).
 */

import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';
import { getMailboxBase } from '../../graph/graphHelpers.js';
import { saveReceiptFile } from '../../utils/receiptFiles.js';
import { fetchPdf, getBillingAllowlist } from './pdfFetch.js';
import { extractBillingPdfUrl } from './receiptParser.js';

function throwIfMcpError(result) {
  if (result && result.isError !== undefined && result.content) throw result;
  return result;
}

export async function fetchBillingPdfCore(graphApiClient, args, deps = {}) {
  const { messageId, url, destDir, fileName, onExisting = 'skip', mailbox } = args;
  const { fetchImpl } = deps;

  let sourceUrl = url;
  if (!sourceUrl) {
    const message = throwIfMcpError(await graphApiClient.makeRequest(
      `${getMailboxBase(mailbox)}/messages/${messageId}`,
      { select: 'id,subject,body' }
    ));
    sourceUrl = extractBillingPdfUrl(message.body?.content || '', getBillingAllowlist());
    if (!sourceUrl) {
      throw new Error(`No allowlisted billing link found in message ${messageId} (allowlist: ${getBillingAllowlist().join(', ')})`);
    }
  }

  const fetchOptions = fetchImpl ? { fetchImpl } : {};
  const { buffer, finalUrl } = await fetchPdf(sourceUrl, fetchOptions);
  const saved = await saveReceiptFile(buffer, destDir, fileName, { onExisting });

  return { ...saved, sourceUrl, finalUrl };
}

export async function fetchBillingPdfTool(authManager, args, deps = {}) {
  if (!args?.fileName) return createValidationError('fileName', 'Parameter is required');
  if (!args?.url && !args?.messageId) {
    return createValidationError('url', 'Either url or messageId is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    const result = await fetchBillingPdfCore(graphApiClient, args, deps);
    return createSafeResponse(result);
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to fetch billing PDF');
  }
}
