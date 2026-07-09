/**
 * Structured receipt extraction tool (FR-3).
 *
 * Returns a compact typed summary of a receipt email (amount, currency,
 * numbers, product label, billing link, attachment ids) — a few KB at most,
 * never the raw 60-100 KB HTML body.
 */

import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';
import { getMailboxBase } from '../../graph/graphHelpers.js';
import { parseReceiptMessage } from './receiptParser.js';
import { getBillingAllowlist } from './pdfFetch.js';
import { loadReceiptRules } from './receiptRules.js';

function throwIfMcpError(result) {
  if (result && result.isError !== undefined && result.content) throw result;
  return result;
}

export async function extractReceiptCore(graphApiClient, messageId, rules = loadReceiptRules()) {
  const base = getMailboxBase();
  const message = throwIfMcpError(await graphApiClient.makeRequest(
    `${base}/messages/${messageId}`,
    { select: 'id,subject,from,receivedDateTime,body,hasAttachments' }
  ));

  let attachments = [];
  if (message.hasAttachments) {
    const list = throwIfMcpError(await graphApiClient.makeRequest(
      `${base}/messages/${messageId}/attachments`,
      { select: 'id,name,contentType,size,isInline' }
    ));
    attachments = list.value || [];
  }

  return parseReceiptMessage(message, attachments, getBillingAllowlist(), rules);
}

export async function extractReceiptTool(authManager, args) {
  if (!args?.messageId) return createValidationError('messageId', 'Parameter is required');

  try {
    const rules = loadReceiptRules();
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    const result = await extractReceiptCore(graphApiClient, args.messageId, rules);
    return createSafeResponse(result);
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to extract receipt');
  }
}
