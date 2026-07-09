/**
 * Render a receipt email as a PDF (FR-7 fallback).
 *
 * Used when a receipt has no PDF attachment and no allowlisted billing link
 * (e.g. app-store order receipts). The result is an audit-trail document, marked as
 * rendered — not the formal vendor invoice.
 */

import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';
import { getMailboxBase } from '../../graph/graphHelpers.js';
import { saveReceiptFile } from '../../utils/receiptFiles.js';
import { buildRenderDocument, renderHtmlToPdfBuffer } from './htmlToPdf.js';

function throwIfMcpError(result) {
  if (result && result.isError !== undefined && result.content) throw result;
  return result;
}

export async function renderEmailPdfCore(graphApiClient, args, deps = {}) {
  const { messageId, destDir, fileName, onExisting = 'skip' } = args;
  const { renderImpl } = deps;

  const message = throwIfMcpError(await graphApiClient.makeRequest(
    `${getMailboxBase()}/messages/${messageId}`,
    { select: 'id,subject,from,receivedDateTime,body' }
  ));

  const doc = buildRenderDocument({
    subject: message.subject,
    fromAddress: message.from?.emailAddress?.address,
    receivedDateTime: message.receivedDateTime,
    bodyHtml: message.body?.content || ''
  });

  const buffer = await (renderImpl || renderHtmlToPdfBuffer)(doc);
  const saved = await saveReceiptFile(buffer, destDir, fileName, { onExisting });

  return { ...saved, method: 'rendered', messageId };
}

export async function renderEmailPdfTool(authManager, args, deps = {}) {
  if (!args?.messageId) return createValidationError('messageId', 'Parameter is required');
  if (!args?.fileName) return createValidationError('fileName', 'Parameter is required');

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    const result = await renderEmailPdfCore(graphApiClient, args, deps);
    return createSafeResponse(result);
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to render email as PDF');
  }
}
