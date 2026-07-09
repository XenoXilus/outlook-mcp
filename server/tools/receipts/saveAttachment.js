/**
 * Save a message attachment's original bytes to disk (FR-1).
 *
 * No text extraction, no transformation, no timestamp suffix — the exact
 * attachment bytes land at destDir/fileName under the receipts/work dir.
 */

import { Buffer } from 'buffer';
import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';
import { getMailboxBase } from '../../graph/graphHelpers.js';
import { saveReceiptFile, renderFilenameTemplate, getFilenameTemplate } from '../../utils/receiptFiles.js';
import { selectPdfAttachment, detectVendor } from './receiptParser.js';
import { loadReceiptRules } from './receiptRules.js';

function throwIfMcpError(result) {
  if (result && result.isError !== undefined && result.content) throw result;
  return result;
}

export async function saveAttachmentCore(graphApiClient, args, rules = loadReceiptRules()) {
  const {
    messageId,
    attachmentId,
    attachmentName,
    contentType: contentTypeSelector,
    prefer = 'invoice',
    destDir,
    fileName,
    filenameTemplate,
    vendor,
    onExisting = 'skip'
  } = args;

  const base = getMailboxBase();
  let targetId = attachmentId;

  if (!targetId) {
    const list = throwIfMcpError(await graphApiClient.makeRequest(
      `${base}/messages/${messageId}/attachments`,
      { select: 'id,name,contentType,size,isInline' }
    ));
    const attachments = list.value || [];
    let selected = null;
    if (attachmentName) {
      selected = attachments.find(a => a.name === attachmentName);
    } else if (contentTypeSelector) {
      selected = attachments.find(a =>
        !a.isInline &&
        (a.contentType || '').toLowerCase() === contentTypeSelector.toLowerCase()
      );
    } else {
      selected = selectPdfAttachment(attachments, prefer);
    }
    if (!selected) {
      const names = attachments.map(a => a.name).join(', ') || 'none';
      if (contentTypeSelector) {
        throw new Error(`No non-inline attachment with contentType '${contentTypeSelector}' found on message ${messageId} (attachments: ${names})`);
      }
      throw new Error(`No suitable PDF attachment found on message ${messageId} (attachments: ${names})`);
    }
    targetId = selected.id;
  }

  // Fetch message details only when the template path is triggered
  const useTemplatePath = !fileName && (filenameTemplate !== undefined || vendor !== undefined);
  let resolvedVendor, resolvedDate;
  if (useTemplatePath) {
    const msg = throwIfMcpError(await graphApiClient.makeRequest(
      `${base}/messages/${messageId}`,
      { select: 'receivedDateTime,from,subject' }
    ));
    resolvedDate = msg.receivedDateTime;
    const fromAddress = msg.from?.emailAddress?.address || '';
    const subject = msg.subject || '';
    resolvedVendor = vendor || detectVendor(fromAddress, subject, rules.vendorSenders);
    if (!resolvedVendor) {
      throw new Error(
        'Cannot render filename template: vendor could not be auto-detected. Provide a `vendor` argument.'
      );
    }
  }

  const full = throwIfMcpError(await graphApiClient.makeRequest(
    `${base}/messages/${messageId}/attachments/${targetId}`,
    { select: 'id,name,contentType,size,isInline,contentBytes' }
  ));
  if (!full.contentBytes) {
    throw new Error(`Attachment '${full.name || targetId}' has no contentBytes (not a file attachment)`);
  }

  const buffer = Buffer.from(full.contentBytes, 'base64');

  let finalFileName;
  if (fileName) {
    finalFileName = fileName;
  } else if (useTemplatePath) {
    const template = filenameTemplate || getFilenameTemplate();
    finalFileName = renderFilenameTemplate(template, { vendor: resolvedVendor, date: resolvedDate });
  } else {
    finalFileName = full.name;
  }

  const saved = await saveReceiptFile(buffer, destDir, finalFileName, { onExisting });

  return {
    ...saved,
    attachmentId: full.id,
    attachmentName: full.name,
    contentType: full.contentType,
    messageId
  };
}

export async function saveAttachmentTool(authManager, args) {
  if (!args?.messageId) return createValidationError('messageId', 'Parameter is required');

  try {
    const rules = loadReceiptRules();
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    const result = await saveAttachmentCore(graphApiClient, args, rules);
    return createSafeResponse(result);
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to save attachment');
  }
}
