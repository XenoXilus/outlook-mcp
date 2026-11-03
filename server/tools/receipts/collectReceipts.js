/**
 * Batch "collect receipts for period" orchestration.
 *
 * Discovery: each vendor rule searches the whole mailbox by sender/subject and
 * date. Messages in Sent Items and Drafts are dropped (a user's own reply or
 * forward quoting a receipt is never a vendor receipt) and a message found by
 * several rule searches is processed once.
 *
 * Claiming: every rule is evaluated against every discovered message (sender,
 * normalised subject, optional productContains/productExcludes). A message
 * claimed by one rule belongs to it; a positive productContains match beats a
 * rule without product matchers; anything still contested, or whose product
 * cannot be determined for a rule that needs it, becomes an `ambiguous` entry.
 *
 * Saving: copies of one receipt (same receipt/invoice/order number) collapse to
 * one representative and the rest are `duplicate`. The representative's PDF
 * comes from its attachment, else a per-receipt allowlisted link, else a
 * rendering of the email. A formal source that fails is `fetch_failed` (also
 * listed in failed[]) and gets a rendered copy that is never presented as the
 * invoice. Amounts prefer the saved PDF's text over the email body.
 *
 * Idempotent under re-runs (onExisting); one structured log line per outcome.
 */

import fs from 'fs';
import path from 'path';
import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';
import { getMailboxBase, toStartDateTime, endDateFilterClause } from '../../graph/graphHelpers.js';
import { getFilenameTemplate, renderFilenameTemplate } from '../../utils/receiptFiles.js';
import { loadReceiptRules } from './receiptRules.js';
import { extractReceiptCore } from './extractReceipt.js';
import { saveAttachmentCore } from './saveAttachment.js';
import { fetchBillingPdfCore } from './fetchBillingPdf.js';
import { renderEmailPdfCore } from './renderEmailPdf.js';
import { parseReceiptPdf } from './receiptPdfText.js';
import { normaliseSubject, selectPdfAttachment } from './receiptParser.js';

const STATUSES = ['saved', 'skipped', 'not_found', 'fetch_failed', 'ambiguous', 'duplicate'];
const MESSAGE_SELECT = 'id,subject,from,receivedDateTime,hasAttachments,parentFolderId';
// Well-known folders whose messages are the user's own mail, never a vendor's.
const EXCLUDED_FOLDERS = ['sentitems', 'drafts'];
const NUMBER_KINDS = ['receiptNumber', 'invoiceNumber', 'orderNumber'];
// Money fields dropped from a duplicate so totals are never counted twice.
const MONEY_FIELDS = ['amount', 'currency', 'taxAmount', 'amounts', 'amountSource', 'amountMismatch', 'currencyMismatch'];
// Save actions meaning this run wrote a new file (so removing it loses nothing).
const CREATED_ACTIONS = new Set(['saved', 'versioned']);
const PRODUCT_UNKNOWN_REASON =
  'product could not be determined (no product label or line items), so productContains/productExcludes cannot decide which rule it belongs to';

function throwIfMcpError(result) {
  if (result && result.isError !== undefined && result.content) throw result;
  return result;
}

function mcpErrorText(error) {
  return error?.content?.[0]?.text || error?.message || String(error);
}

/** Error text plus the Graph error and attachment fetch details, whichever the error carries. */
function describeFailure(error) {
  const details = error?.fetchDetails || null;
  const graphError = details?.graphError || error?._errorDetails?.graphError || error?.graphError || null;
  return { text: mcpErrorText(error), graphError: graphError ? { ...graphError } : null, details };
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

function withSuffix(fileName, suffix) {
  const ext = path.extname(fileName);
  const base = ext ? fileName.slice(0, -ext.length) : fileName;
  return `${base}${suffix}${ext}`;
}

const ISO_DATE = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2})?/;

function escapeODataString(value) {
  return String(value).replace(/'/g, "''");
}

/**
 * Lower-case an address and strip any plus tag from the local part, so
 * billing+statements+acct_1@example.com and billing@example.com compare equal.
 */
function normaliseEmailAddress(address) {
  const lower = String(address || '').toLowerCase().trim();
  const at = lower.lastIndexOf('@');
  if (at === -1) return lower;
  return `${lower.slice(0, at).split('+')[0]}@${lower.slice(at + 1)}`;
}

function senderOf(message) {
  return message.from?.emailAddress?.address || null;
}

/** Validation message for a productContains/productExcludes value, or null when valid or absent. */
function productTermsError(value) {
  if (value === undefined || value === null) return null;
  const terms = Array.isArray(value) ? value : [value];
  if (terms.length === 0 || terms.some(t => typeof t !== 'string' || !t.trim())) {
    return 'must be a non-empty string or an array of non-empty strings';
  }
  return null;
}

function toTerms(value) {
  if (value === undefined || value === null) return [];
  return (Array.isArray(value) ? value : [value]).map(t => t.trim());
}

/** Which of the rule's sender/subject signals this message satisfied. */
function headerSignals(rule, message) {
  const matched = [];
  const fromAddress = (senderOf(message) || '').toLowerCase();
  if (rule.from) {
    if (fromAddress === rule.from.toLowerCase()) {
      matched.push('from');
    } else if (normaliseEmailAddress(fromAddress) === normaliseEmailAddress(rule.from)) {
      matched.push('from-normalised');
    }
  }
  // Tested against the subject with Re:/Fw: prefixes stripped.
  if (rule.subjectContains &&
      normaliseSubject(message.subject).toLowerCase().includes(rule.subjectContains.toLowerCase())) {
    matched.push('subject');
  }
  return matched;
}

/**
 * Product matchers, case-insensitive substrings against the product label and
 * every line-item description. Verdict: 'none' (rule has no matchers),
 * 'unknown' (receipt has no product text), 'reject' or 'accept'.
 */
function evaluateProduct(rule, receipt) {
  const contains = toTerms(rule.productContains);
  const excludes = toTerms(rule.productExcludes);
  if (contains.length === 0 && excludes.length === 0) return { verdict: 'none', signals: [] };

  const texts = [receipt.productLabel, ...(receipt.lineItems || []).map(item => item?.description)]
    .filter(t => typeof t === 'string' && t.trim())
    .map(t => t.toLowerCase());
  if (texts.length === 0) return { verdict: 'unknown', signals: [] };

  const hit = (term) => texts.some(t => t.includes(term.toLowerCase()));
  if (excludes.some(hit)) return { verdict: 'reject', signals: [] };
  const signals = [];
  if (contains.length > 0) {
    const matched = contains.filter(hit);
    if (matched.length === 0) return { verdict: 'reject', signals: [] };
    signals.push(...matched.map(term => `product:contains:${term}`));
  }
  if (excludes.length > 0) signals.push('product:excludes');
  return { verdict: 'accept', containsMatched: contains.length > 0, signals };
}

/**
 * Pick the owning rule among the rules that claim a message: { winner } or
 * { contenders, reason } when it must be reported as ambiguous.
 */
function resolveClaims(claims, vendorOf) {
  if (claims.some(c => c.product.verdict === 'unknown')) {
    return { contenders: claims, reason: PRODUCT_UNKNOWN_REASON };
  }
  if (claims.length === 1) return { winner: claims[0] };
  // A rule whose productContains matched is more specific than one without product matchers.
  const positive = claims.filter(c => c.product.containsMatched);
  if (positive.length === 1) return { winner: positive[0] };
  const contenders = positive.length > 1 ? positive : claims;
  return {
    contenders,
    reason: `claimed by ${contenders.length} vendor rules (${contenders.map(c => vendorOf(c.index)).join(', ')}); ` +
      'narrow from/subjectContains/productContains so exactly one rule matches'
  };
}

/** The number that identifies a receipt across copies: receipt, else invoice, else order number. */
function dedupeKey(receipt) {
  for (const kind of NUMBER_KINDS) {
    const value = typeof receipt[kind] === 'string' ? receipt[kind].trim() : '';
    if (value) return { kind, value };
  }
  return null;
}

/** sender, forwarded and (for a forward with a quoted header) the original sender. */
function senderFields(rule, message, receipt) {
  const sender = senderOf(message);
  const subject = String(message.subject || '');
  const replyOrForward = receipt
    ? Boolean(receipt.replyOrForward)
    : normaliseSubject(subject) !== subject.replace(/\s+/g, ' ').trim();
  const fields = {
    sender,
    forwarded: replyOrForward ||
      Boolean(rule?.from && normaliseEmailAddress(sender) !== normaliseEmailAddress(rule.from))
  };
  if (receipt?.quotedFrom?.address) fields.originalSender = receipt.quotedFrom.address;
  if (receipt?.quotedFrom?.name) fields.originalSenderName = receipt.quotedFrom.name;
  return fields;
}

function receiptFields(receipt) {
  const fields = {
    amount: receipt.amount ?? null,
    currency: receipt.currency ?? null,
    taxAmount: receipt.taxAmount ?? null,
    amounts: receipt.amounts ? { ...receipt.amounts } : null,
    productLabel: receipt.productLabel ?? null
  };
  for (const kind of NUMBER_KINDS) {
    if (receipt[kind]) fields[kind] = receipt[kind];
  }
  return fields;
}

/**
 * Prefer the totals printed in the saved PDF over the email body's. Never
 * throws: an unreadable PDF leaves the body amounts in place with a note.
 */
async function applyPdfAmounts(entry, savedPath, rules) {
  let pdf;
  try {
    pdf = await parseReceiptPdf(await fs.promises.readFile(savedPath), rules);
  } catch (error) {
    entry.amountSource = 'body';
    entry.pdfParseNote = `saved PDF could not be parsed (${error?.message || error}); amounts are from the email body`;
    return;
  }
  const bodyTotal = entry.amount;
  if (pdf.amounts?.total !== null && pdf.amounts?.total !== undefined) {
    entry.amount = pdf.amounts.total;
    entry.currency = pdf.amounts.currency || entry.currency;
    entry.taxAmount = pdf.amounts.tax ?? null;
    entry.amounts = { ...pdf.amounts };
    entry.amountSource = 'pdf';
    if (typeof bodyTotal === 'number' && Math.abs(bodyTotal - pdf.amounts.total) > 0.005) {
      entry.amountMismatch = { body: bodyTotal, pdf: pdf.amounts.total };
    }
  } else {
    entry.amountSource = 'body';
  }
  if (!entry.productLabel && pdf.productLabel) entry.productLabel = pdf.productLabel;
  if (!entry.invoiceNumber && pdf.invoiceNumber) entry.invoiceNumber = pdf.invoiceNumber;
  if (!entry.receiptNumber && pdf.receiptNumber) entry.receiptNumber = pdf.receiptNumber;
}

function recordSaved(entry, saved, method, isFormalInvoice) {
  Object.assign(entry, {
    status: saved.action === 'skipped' ? 'skipped' : 'saved',
    savedPath: saved.savedPath,
    action: saved.action,
    method,
    isFormalInvoice,
    sha256: saved.sha256
  });
}

async function searchMessages(graphApiClient, base, filters) {
  const result = throwIfMcpError(await graphApiClient.makeRequest(`${base}/messages`, {
    filter: filters.join(' and '),
    select: MESSAGE_SELECT,
    orderby: 'receivedDateTime desc',
    top: 50
  }));
  return result.value || [];
}

/**
 * Ids of the Sent Items and Drafts folders. A folder that cannot be resolved
 * excludes nothing and leaves a run-level note.
 */
async function resolveExcludedFolderIds(graphApiClient, base, notes) {
  const ids = new Set();
  for (const name of EXCLUDED_FOLDERS) {
    try {
      const folder = throwIfMcpError(await graphApiClient.makeRequest(`${base}/mailFolders/${name}`, { select: 'id' }));
      if (!folder?.id) throw new Error('no folder id returned');
      ids.add(folder.id);
    } catch (error) {
      const text = mcpErrorText(error);
      notes.push(`Could not resolve the ${name} folder (${text}); messages in it were not excluded`);
      logCollectEvent({ stage: 'folders', folder: name, outcome: 'error', error: text });
    }
  }
  return ids;
}

/** True when the message's PDF attachment (as the save would pick it) is an Invoice-*.pdf. */
async function hasInvoiceAttachment(graphApiClient, base, messageId) {
  try {
    const list = throwIfMcpError(await graphApiClient.makeRequest(
      `${base}/messages/${messageId}/attachments`,
      { select: 'id,name,contentType,size,isInline' }
    ));
    const pdf = selectPdfAttachment(list.value || [], 'invoice');
    return Boolean(pdf && /^invoice/i.test(pdf.name || ''));
  } catch {
    return false; // only a tie-break: a failure just loses the preference
  }
}

/**
 * Representative of a group of copies: one with an Invoice-*.pdf attachment,
 * else any PDF attachment, else none; ties go to the earliest received.
 */
async function chooseRepresentative(members, graphApiClient, base) {
  const rank = new Map(members.map(m => [m, m.receipt.hasPdfAttachment ? 1 : 0]));
  const withPdf = members.filter(m => m.receipt.hasPdfAttachment);
  if (withPdf.length > 1) {
    for (const m of withPdf) {
      if (await hasInvoiceAttachment(graphApiClient, base, m.message.id)) rank.set(m, 2);
    }
  }
  const time = (m) => {
    const t = Date.parse(m.message.receivedDateTime);
    return Number.isNaN(t) ? Infinity : t;
  };
  return [...members].sort((a, b) => rank.get(b) - rank.get(a) || time(a) - time(b))[0];
}

export async function collectReceiptsTool(authManager, args, deps = {}) {
  const { periodStart, periodEnd, vendors, destDir, onExisting = 'skip', mailbox } = args || {};

  if (!periodStart || !ISO_DATE.test(periodStart)) {
    return createValidationError('periodStart', 'Required ISO date or date-time, e.g. 2026-06-01 or 2026-06-01T00:00:00Z');
  }
  if (!periodEnd || !ISO_DATE.test(periodEnd)) {
    return createValidationError('periodEnd', 'Required ISO date or date-time, e.g. 2026-06-30 or 2026-06-30T23:59:59Z');
  }
  if (!Array.isArray(vendors) || vendors.length === 0) {
    return createValidationError('vendors', 'At least one vendor rule is required');
  }
  for (const rule of vendors) {
    if (!rule?.vendor) return createValidationError('vendors', 'Each rule needs a vendor name');
    if (!rule.from && !rule.subjectContains) {
      return createValidationError('vendors', `Rule '${rule.vendor}' needs a from and/or subjectContains matcher`);
    }
    for (const field of ['productContains', 'productExcludes']) {
      const problem = productTermsError(rule[field]);
      if (problem) return createValidationError('vendors', `Rule '${rule.vendor}': ${field} ${problem}`);
    }
  }

  try {
    const rules = loadReceiptRules();
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    // One base for the whole run: searches and every per-message core call
    // must address the same mailbox.
    const base = getMailboxBase(mailbox);

    const notes = [];
    const usedNames = new Map();
    const states = vendors.map(rule => ({
      rule, found: new Set(), searchNote: null, searchFailed: false, excludedByProduct: 0, excludedByFolder: 0
    }));
    // Entries are kept per rule (an ambiguous entry under its first candidate)
    // so the manifest reads vendor by vendor.
    const buckets = vendors.map(() => []);
    const involved = new Set();
    const vendorOf = (index) => vendors[index].vendor;

    // Date filter is shared by every vendor query. Date-only bounds are
    // expanded (end becomes a half-open `lt next-day`) so the month's last
    // day is never silently dropped.
    const dateFilters = [
      `receivedDateTime ge ${toStartDateTime(periodStart)}`,
      endDateFilterClause(periodEnd)
    ];

    // --- Discovery -------------------------------------------------------
    const discovered = new Map();
    for (const [index, state] of states.entries()) {
      const { rule } = state;
      let messages;
      try {
        // from and subjectContains are signals, not a hard AND: a plus-address
        // drift or reworded subject must not silently hide a receipt.
        const matchers = [];
        if (rule.from) matchers.push(`from/emailAddress/address eq '${escapeODataString(rule.from)}'`);
        if (rule.subjectContains) matchers.push(`contains(subject,'${escapeODataString(rule.subjectContains)}')`);
        messages = await searchMessages(graphApiClient, base, [
          ...dateFilters,
          matchers.length === 1 ? matchers[0] : `(${matchers.join(' or ')})`
        ]);
      } catch (error) {
        const failure = describeFailure(error);
        state.searchFailed = true;
        buckets[index].push({
          vendor: rule.vendor,
          status: 'fetch_failed',
          attemptedRoute: 'search',
          graphError: failure.graphError,
          error: `Search failed: ${failure.text}`
        });
        logCollectEvent({ vendor: rule.vendor, stage: 'search', status: 'fetch_failed', outcome: 'error', error: failure.text });
        continue;
      }

      // Zero hits with a from rule: retry on the plus-stripped local part and
      // keep only messages whose normalised sender really is the rule's sender.
      if (messages.length === 0 && rule.from) {
        try {
          const localBase = normaliseEmailAddress(rule.from).split('@')[0];
          const retried = await searchMessages(graphApiClient, base, [
            ...dateFilters,
            `startswith(from/emailAddress/address,'${escapeODataString(localBase)}')`
          ]);
          messages = retried.filter(m => normaliseEmailAddress(senderOf(m)) === normaliseEmailAddress(rule.from));
          if (messages.length > 0) {
            state.searchNote = `matched via plus-address-normalised sender fallback (exact from '${rule.from}' found nothing)`;
            logCollectEvent({ vendor: rule.vendor, stage: 'search', outcome: 'fallback-match', count: messages.length });
          }
        } catch (error) {
          // The fallback is best-effort: a failure here degrades to not_found.
          logCollectEvent({ vendor: rule.vendor, stage: 'search-fallback', outcome: 'error', error: mcpErrorText(error) });
        }
      }

      for (const message of messages) {
        if (!discovered.has(message.id)) discovered.set(message.id, message);
        state.found.add(message.id);
      }
    }

    if (discovered.size > 0) {
      const excludedFolders = await resolveExcludedFolderIds(graphApiClient, base, notes);
      for (const [id, message] of discovered) {
        if (!message.parentFolderId || !excludedFolders.has(message.parentFolderId)) continue;
        discovered.delete(id);
        for (const state of states) {
          if (state.found.delete(id)) state.excludedByFolder++;
        }
      }
    }

    // --- Extraction and claiming -----------------------------------------
    const owned = [];
    for (const message of discovered.values()) {
      let receipt = null;
      let extractFailure = null;
      try {
        receipt = await extractReceiptCore(graphApiClient, message.id, rules, mailbox);
      } catch (error) {
        extractFailure = describeFailure(error);
      }

      const claims = [];
      for (const [index, state] of states.entries()) {
        const signals = headerSignals(state.rule, message);
        if (signals.length === 0) {
          if (!state.found.has(message.id)) continue;
          signals.push('search');
        }
        if (!receipt) {
          claims.push({ index, signals, product: { verdict: 'none', signals: [] } });
          continue;
        }
        const product = evaluateProduct(state.rule, receipt);
        if (product.verdict === 'reject') {
          state.excludedByProduct++;
          continue;
        }
        claims.push({ index, signals, product });
      }
      if (claims.length === 0) continue;
      const soleRule = claims.length === 1 ? vendors[claims[0].index] : null;

      if (extractFailure) {
        claims.forEach(c => involved.add(c.index));
        const entry = {
          vendor: soleRule ? soleRule.vendor : null,
          status: 'fetch_failed',
          attemptedRoute: 'message',
          date: message.receivedDateTime,
          sourceMessageId: message.id,
          messageId: message.id,
          ...senderFields(soleRule, message, null),
          graphError: extractFailure.graphError,
          error: extractFailure.text
        };
        if (!soleRule) entry.candidates = claims.map(c => vendorOf(c.index));
        buckets[claims[0].index].push(entry);
        logCollectEvent({
          vendor: entry.vendor, stage: 'extract', status: 'fetch_failed', messageId: message.id, error: extractFailure.text
        });
        continue;
      }

      const resolved = resolveClaims(claims, vendorOf);
      if (resolved.contenders) {
        const { contenders } = resolved;
        contenders.forEach(c => involved.add(c.index));
        const single = contenders.length === 1 ? vendors[contenders[0].index] : null;
        buckets[contenders[0].index].push({
          vendor: single ? single.vendor : null,
          status: 'ambiguous',
          candidates: contenders.map(c => vendorOf(c.index)),
          reason: resolved.reason,
          date: message.receivedDateTime,
          sourceMessageId: message.id,
          ...senderFields(single, message, receipt),
          ...receiptFields(receipt)
        });
        logCollectEvent({
          stage: 'claim', status: 'ambiguous', candidates: contenders.map(c => vendorOf(c.index)),
          messageId: message.id, reason: resolved.reason
        });
        continue;
      }

      const { winner } = resolved;
      involved.add(winner.index);
      owned.push({
        message, receipt, ruleIndex: winner.index,
        matchedBy: [...winner.signals, ...winner.product.signals]
      });
    }

    // --- Dedupe copies of one receipt -------------------------------------
    const groups = new Map();
    for (const item of owned) {
      const key = dedupeKey(item.receipt);
      item.groupId = key
        ? `${item.ruleIndex}|${key.kind}|${key.value.toLowerCase()}`
        : `message|${item.message.id}`;
      if (!groups.has(item.groupId)) groups.set(item.groupId, { kind: key?.kind, members: [] });
      groups.get(item.groupId).members.push(item);
    }
    for (const group of groups.values()) {
      group.representative = group.members.length > 1
        ? await chooseRepresentative(group.members, graphApiClient, base)
        : group.members[0];
      if (group.kind) {
        group.label = `${group.kind}:${group.representative.receipt[group.kind].trim()}`;
      }
    }

    // A billing link shared by several distinct receipts is a generic portal
    // link, not a per-receipt invoice. Copies of one receipt count once.
    const linkUse = new Map();
    for (const item of owned) {
      const url = item.receipt.billingPdfUrl;
      if (!url) continue;
      if (!linkUse.has(url)) linkUse.set(url, new Set());
      linkUse.get(url).add(item.groupId);
    }

    // --- Save -------------------------------------------------------------
    const saveOrder = [];
    for (const item of owned) {
      const { message, receipt } = item;
      const { rule, searchNote } = states[item.ruleIndex];
      const group = groups.get(item.groupId);
      const entry = {
        vendor: rule.vendor,
        status: null,
        date: message.receivedDateTime,
        sourceMessageId: message.id,
        matchedBy: item.matchedBy,
        ...senderFields(rule, message, receipt)
      };
      if (searchNote) entry.searchNote = searchNote;
      buckets[item.ruleIndex].push(entry);

      if (group.representative !== item) {
        entry.status = 'duplicate';
        entry.duplicate_of = group.label;
        entry.duplicateOfMessageId = group.representative.message.id;
        for (const kind of NUMBER_KINDS) {
          if (receipt[kind]) entry[kind] = receipt[kind];
        }
        logCollectEvent({
          vendor: rule.vendor, stage: 'dedupe', status: 'duplicate', duplicate_of: entry.duplicate_of, messageId: message.id
        });
        continue;
      }

      Object.assign(entry, receiptFields(receipt));
      let formal = null;
      if (receipt.hasPdfAttachment) {
        formal = 'attachment';
      } else if (receipt.billingPdfUrl) {
        const sharedBy = linkUse.get(receipt.billingPdfUrl).size;
        if (sharedBy > 1) {
          entry.linkSkipped = `billing link is shared by ${sharedBy} receipts in this run, so it is not a per-receipt invoice link`;
        } else {
          formal = 'link';
        }
      }

      let fileName = null;
      let saved = null;
      try {
        const template = rule.filenameTemplate || getFilenameTemplate();
        fileName = dedupeFilename(renderFilenameTemplate(template, {
          vendor: rule.vendor,
          date: message.receivedDateTime
        }), usedNames);

        if (formal === 'attachment') {
          saved = await saveAttachmentCore(graphApiClient, {
            messageId: message.id,
            prefer: rule.prefer || 'invoice',
            destDir,
            fileName,
            onExisting,
            mailbox
          }, rules);
          recordSaved(entry, saved, 'attachment', true);
          entry.attachmentId = saved.attachmentId;
          entry.attachmentName = saved.attachmentName;
          await applyPdfAmounts(entry, saved.savedPath, rules);
        } else if (formal === 'link') {
          saved = await fetchBillingPdfCore(graphApiClient, {
            url: receipt.billingPdfUrl,
            destDir,
            fileName,
            onExisting,
            mailbox
          }, deps);
          recordSaved(entry, saved, 'link', true);
          await applyPdfAmounts(entry, saved.savedPath, rules);
        } else {
          saved = await renderEmailPdfCore(graphApiClient, {
            messageId: message.id,
            destDir,
            fileName,
            onExisting,
            mailbox
          }, deps);
          recordSaved(entry, saved, 'rendered', false);
          entry.amountSource = 'body';
        }
        saveOrder.push(entry);
        logCollectEvent({
          vendor: rule.vendor, stage: 'save', status: entry.status, outcome: entry.action,
          method: entry.method, bytes: saved.size, sha256: entry.sha256, messageId: message.id
        });
      } catch (error) {
        const failure = describeFailure(error);
        Object.assign(entry, {
          status: 'fetch_failed',
          attemptedRoute: formal || 'render',
          messageId: message.id,
          graphError: failure.graphError,
          error: failure.text,
          amountSource: 'body'
        });
        if (formal === 'attachment') {
          entry.attachmentId = failure.details?.attachmentId ?? null;
          entry.attachmentName = failure.details?.attachmentName ?? null;
        }
        logCollectEvent({
          vendor: rule.vendor, stage: 'save', status: 'fetch_failed', route: entry.attemptedRoute,
          messageId: message.id, error: failure.text
        });

        // Degraded copy: a rendering of the email under its own name, so it is
        // never mistaken for (or later skipped in place of) the formal invoice.
        if (formal && fileName) {
          try {
            const rendered = await renderEmailPdfCore(graphApiClient, {
              messageId: message.id,
              destDir,
              fileName: dedupeFilename(withSuffix(fileName, ' (rendered)'), usedNames),
              onExisting,
              mailbox
            }, deps);
            entry.renderedFallback = {
              savedPath: rendered.savedPath,
              sha256: rendered.sha256,
              action: rendered.action,
              method: 'rendered',
              isFormalInvoice: false,
              reason: `the ${formal} could not be fetched; this is a rendering of the email body, not the formal invoice`
            };
            logCollectEvent({
              vendor: rule.vendor, stage: 'render-fallback', status: 'fetch_failed', outcome: rendered.action,
              sha256: rendered.sha256, messageId: message.id
            });
          } catch (renderError) {
            entry.renderedFallbackError = mcpErrorText(renderError);
            logCollectEvent({
              vendor: rule.vendor, stage: 'render-fallback', status: 'fetch_failed', outcome: 'error',
              messageId: message.id, error: entry.renderedFallbackError
            });
          }
        }
      }

      if (rule.expectedCurrency && entry.currency &&
          rule.expectedCurrency.toUpperCase() !== String(entry.currency).toUpperCase()) {
        entry.currencyMismatch = `expected ${rule.expectedCurrency}, found ${entry.currency}`;
      }
    }

    // --- Identical files saved for different messages ----------------------
    const firstBySha = new Map();
    for (const entry of saveOrder) {
      if (!entry.sha256 || (entry.status !== 'saved' && entry.status !== 'skipped')) continue;
      const first = firstBySha.get(entry.sha256);
      if (!first) {
        firstBySha.set(entry.sha256, entry);
        continue;
      }
      // Only a file this run created is removed; a pre-existing one is left alone.
      if (CREATED_ACTIONS.has(entry.action) && entry.savedPath !== first.savedPath) {
        try {
          await fs.promises.rm(entry.savedPath, { force: true });
          entry.removedPath = entry.savedPath;
          delete entry.savedPath;
        } catch (error) {
          entry.removeError = error.message;
        }
      }
      for (const field of MONEY_FIELDS) delete entry[field];
      entry.status = 'duplicate';
      entry.duplicate_of = `sha256:${entry.sha256}`;
      entry.duplicateOfMessageId = first.sourceMessageId;
      logCollectEvent({
        vendor: entry.vendor, stage: 'dedupe', status: 'duplicate', duplicate_of: entry.duplicate_of,
        messageId: entry.sourceMessageId, removed: Boolean(entry.removedPath)
      });
    }

    // --- Vendors with nothing found ----------------------------------------
    const missing = [];
    for (const [index, state] of states.entries()) {
      const { rule } = state;
      if (state.searchFailed || involved.has(index)) {
        if (state.excludedByProduct > 0) {
          notes.push(`${rule.vendor}: ${state.excludedByProduct} message(s) excluded by productContains/productExcludes`);
        }
        continue;
      }
      const entry = { vendor: rule.vendor, status: 'not_found' };
      if (state.excludedByProduct > 0) entry.excludedByProduct = state.excludedByProduct;
      if (state.excludedByFolder > 0) entry.excludedByFolder = state.excludedByFolder;
      if (state.searchNote) entry.searchNote = state.searchNote;
      buckets[index].push(entry);
      missing.push(rule.vendor);
      logCollectEvent({
        vendor: rule.vendor, stage: 'search', status: 'not_found',
        excludedByProduct: state.excludedByProduct, excludedByFolder: state.excludedByFolder
      });
    }

    const manifest = buckets.flat();
    const summary = Object.fromEntries(STATUSES.map(status => [status, 0]));
    for (const entry of manifest) summary[entry.status]++;
    // Separate copies: the response serialiser treats a repeated object as circular.
    const failed = manifest.filter(e => e.status === 'fetch_failed').map(e => structuredClone(e));

    return createSafeResponse({
      period: { start: periodStart, end: periodEnd },
      manifest,
      missing,
      failed,
      summary,
      notes
    });
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to collect receipts');
  }
}
