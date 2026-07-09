/**
 * Batch "collect receipts for period" orchestration (FR-6).
 *
 * For each vendor rule: discover receipt messages by sender/subject/date
 * across the whole mailbox (FR-5), extract structured details (FR-3), then
 * save the PDF via attachment (FR-1), allowlisted link (FR-2) or rendered
 * fallback (FR-7). Returns a manifest plus a missing[] list. Idempotent under
 * re-runs (NFR-3); one structured log line per message (NFR-4).
 */

import path from 'path';
import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';
import { getMailboxBase } from '../../graph/graphHelpers.js';
import { getFilenameTemplate, renderFilenameTemplate } from '../../utils/receiptFiles.js';
import { loadReceiptRules } from './receiptRules.js';
import { extractReceiptCore } from './extractReceipt.js';
import { saveAttachmentCore } from './saveAttachment.js';
import { fetchBillingPdfCore } from './fetchBillingPdf.js';
import { renderEmailPdfCore } from './renderEmailPdf.js';

function throwIfMcpError(result) {
  if (result && result.isError !== undefined && result.content) throw result;
  return result;
}

function mcpErrorText(error) {
  return error?.content?.[0]?.text || error?.message || String(error);
}

function logCollectEvent(event) {
  console.error(`receipt-collect ${JSON.stringify(event)}`);
}

// Deterministic in-run dedup: 2nd same-named file becomes "name (2).pdf" etc.
function dedupeFilename(fileName, usedNames) {
  const count = usedNames.get(fileName) || 0;
  usedNames.set(fileName, count + 1);
  if (count === 0) return fileName;
  const ext = path.extname(fileName);
  const base = ext ? fileName.slice(0, -ext.length) : fileName;
  return `${base} (${count + 1})${ext}`;
}

const ISO_DATE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/;

export async function collectReceiptsTool(authManager, args, deps = {}) {
  const { periodStart, periodEnd, vendors, destDir, onExisting = 'skip' } = args || {};

  if (!periodStart || !ISO_DATE.test(periodStart)) {
    return createValidationError('periodStart', 'Required ISO date-time, e.g. 2026-06-01T00:00:00Z');
  }
  if (!periodEnd || !ISO_DATE.test(periodEnd)) {
    return createValidationError('periodEnd', 'Required ISO date-time, e.g. 2026-06-30T23:59:59Z');
  }
  if (!Array.isArray(vendors) || vendors.length === 0) {
    return createValidationError('vendors', 'At least one vendor rule is required');
  }
  for (const rule of vendors) {
    if (!rule.vendor) return createValidationError('vendors', 'Each rule needs a vendor name');
    if (!rule.from && !rule.subjectContains) {
      return createValidationError('vendors', `Rule '${rule.vendor}' needs a from and/or subjectContains matcher`);
    }
  }

  try {
    const rules = loadReceiptRules();
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    const base = getMailboxBase();

    const manifest = [];
    const missing = [];
    const usedNames = new Map();

    for (const rule of vendors) {
      let messages = [];
      try {
        const filters = [
          `receivedDateTime ge ${periodStart}`,
          `receivedDateTime le ${periodEnd}`
        ];
        if (rule.from) filters.push(`from/emailAddress/address eq '${rule.from.replace(/'/g, "''")}'`);
        if (rule.subjectContains) filters.push(`contains(subject,'${rule.subjectContains.replace(/'/g, "''")}')`);

        const searchResult = throwIfMcpError(await graphApiClient.makeRequest(`${base}/messages`, {
          filter: filters.join(' and '),
          select: 'id,subject,from,receivedDateTime,hasAttachments',
          orderby: 'receivedDateTime desc',
          top: 50
        }));
        messages = searchResult.value || [];
      } catch (error) {
        const text = mcpErrorText(error);
        logCollectEvent({ vendor: rule.vendor, stage: 'search', outcome: 'error', error: text });
        manifest.push({ vendor: rule.vendor, status: 'ambiguous', error: `Search failed: ${text}` });
        continue;
      }

      if (messages.length === 0) {
        missing.push(rule.vendor);
        manifest.push({ vendor: rule.vendor, status: 'missing' });
        logCollectEvent({ vendor: rule.vendor, stage: 'search', outcome: 'missing' });
        continue;
      }

      for (const message of messages) {
        try {
          const receipt = await extractReceiptCore(graphApiClient, message.id, rules);

          const template = rule.filenameTemplate || getFilenameTemplate();
          const baseName = renderFilenameTemplate(template, {
            vendor: rule.vendor,
            date: message.receivedDateTime
          });
          const fileName = dedupeFilename(baseName, usedNames);

          let saved;
          let method;
          if (receipt.hasPdfAttachment) {
            saved = await saveAttachmentCore(graphApiClient, {
              messageId: message.id,
              prefer: rule.prefer || 'invoice',
              destDir,
              fileName,
              onExisting
            }, rules);
            method = 'attachment';
          } else if (receipt.billingPdfUrl) {
            saved = await fetchBillingPdfCore(graphApiClient, {
              url: receipt.billingPdfUrl,
              destDir,
              fileName,
              onExisting
            }, deps);
            method = 'link';
          } else {
            saved = await renderEmailPdfCore(graphApiClient, {
              messageId: message.id,
              destDir,
              fileName,
              onExisting
            }, deps);
            method = 'rendered';
          }

          const entry = {
            vendor: rule.vendor,
            status: 'saved',
            savedPath: saved.savedPath,
            action: saved.action,
            amount: receipt.amount,
            currency: receipt.currency,
            date: message.receivedDateTime,
            sourceMessageId: message.id,
            method,
            sha256: saved.sha256
          };
          if (rule.expectedCurrency && receipt.currency && rule.expectedCurrency !== receipt.currency) {
            entry.currencyMismatch = `expected ${rule.expectedCurrency}, found ${receipt.currency}`;
          }
          manifest.push(entry);
          logCollectEvent({
            vendor: rule.vendor, stage: 'save', outcome: saved.action,
            method, bytes: saved.size, sha256: saved.sha256, messageId: message.id
          });
        } catch (error) {
          const text = mcpErrorText(error);
          manifest.push({
            vendor: rule.vendor,
            status: 'ambiguous',
            date: message.receivedDateTime,
            sourceMessageId: message.id,
            error: text
          });
          logCollectEvent({ vendor: rule.vendor, stage: 'save', outcome: 'error', messageId: message.id, error: text });
        }
      }
    }

    return createSafeResponse({
      period: { start: periodStart, end: periodEnd },
      manifest,
      missing
    });
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to collect receipts');
  }
}
