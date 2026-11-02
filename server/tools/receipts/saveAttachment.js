/**
 * Save a message attachment's original bytes to disk (FR-1).
 *
 * No text extraction, no transformation, no timestamp suffix — the exact
 * attachment bytes land at destDir/fileName under the receipts/work dir.
 */

import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';
import { getMailboxBase } from '../../graph/graphHelpers.js';
import { saveReceiptFile, renderFilenameTemplate, getFilenameTemplate } from '../../utils/receiptFiles.js';
import { fetchAttachmentContent } from '../../graph/attachmentContent.js';
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
    onExisting = 'skip',
    mailbox
  } = args;

  const base = getMailboxBase(mailbox);
  let targetId = attachmentId;
  let knownName = attachmentName;

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
    knownName = selected.name;
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

  // Throws an Error carrying `fetchDetails` (attempted routes, ids, last Graph error)
  // when the bytes cannot be fetched.
  const { buffer, attachment: full } = await fetchAttachmentContent(
    graphApiClient, base, messageId, targetId, { attachmentName: knownName }
  );

  let finalFileName;
  if (fileName) {
    finalFileName = fileName;
  } else if (useTemplatePath) {
    const template = filenameTemplate || getFilenameTemplate();
    finalFileName = renderFilenameTemplate(template, { vendor: resolvedVendor, date: resolvedDate });
  } else {
    finalFileName = full.name || knownName;
  }

  const saved = await saveReceiptFile(buffer, destDir, finalFileName, { onExisting });

  return {
    ...saved,
    attachmentId: full.id || targetId,
    attachmentName: full.name || knownName,
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
