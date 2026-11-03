/**
 * Batch "collect receipts for period" orchestration.
 *
 * Discovery: each vendor rule searches the whole mailbox by sender/subject and
 * date. Messages in Sent Items and Drafts are dropped (a user's own reply or
 * forward quoting a receipt is never a vendor receipt) and a message found by
 * several rule searches is processed once.
 *
 * Claiming: every rule is evaluated against every discovered message (sender,
 * normalised subject, optional productContains/productExcludes). When the email
 * body names no product and a rule needs one, the product is read from the PDF
 * attachment. A message claimed by one rule belongs to it; a positive
 * productContains match beats a rule without product matchers; anything still
 * contested, or whose product cannot be determined for a rule that needs it,
 * becomes an `ambiguous` entry. A message every claiming rule rejects by product
 * is listed in unmatched[].
 *
 * Saving: copies of one receipt (sharing a receipt, invoice or order number,
 * with no two different numbers of one kind and no two different known totals
 * across the group) collapse to one representative and the rest are
 * `duplicate`. The representative's PDF comes from its attachment, else a
 * per-receipt allowlisted link, else a rendering of the email. Names that
 * collide within the run carry the receipt's own number (ordinals, oldest
 * first, for receipts without one), and a receipt already saved under the
 * plain name by an earlier run keeps that file. A formal source that fails is
 * `fetch_failed` (also listed in failed[]) and gets a rendered copy that is
 * never presented as the invoice. Amounts prefer the saved PDF's labelled
 * total over the email body, except a PDF total of zero.
 *
 * Idempotent under re-runs (onExisting: 'skip' only skips a file holding the
 * same document); one structured log line per outcome.
 */

import fs from 'fs';
import path from 'path';
import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';
import { getMailboxBase, toStartDateTime, endDateFilterClause } from '../../graph/graphHelpers.js';
import {
  getFilenameTemplate,
  renderFilenameTemplate,
  resolveDestination,
  sanitiseFilename,
  existingVersionPaths,
  findSameDocument,
  sha256Hex
} from '../../utils/receiptFiles.js';
import { fetchAttachmentContent } from '../../graph/attachmentContent.js';
import { loadReceiptRules } from './receiptRules.js';
import { extractReceiptCore } from './extractReceipt.js';
import { saveAttachmentCore } from './saveAttachment.js';
import { fetchBillingPdfCore } from './fetchBillingPdf.js';
import { renderEmailPdfCore } from './renderEmailPdf.js';
import { parseReceiptPdf } from './receiptPdfText.js';
import { normaliseSubject, selectPdfAttachment } from './receiptParser.js';

const STATUSES = ['saved', 'skipped', 'not_found', 'fetch_failed', 'ambiguous', 'duplicate'];
const MESSAGE_SELECT = 'id,subject,from,receivedDateTime,hasAttachments,parentFolderId';
const ATTACHMENT_LIST_SELECT = 'id,name,contentType,size,isInline';
const ON_EXISTING = ['skip', 'overwrite', 'version'];
const PREFER = ['invoice', 'receipt', 'first'];
// Optional rule fields that must be strings when given.
const STRING_RULE_FIELDS = ['from', 'subjectContains', 'filenameTemplate', 'expectedCurrency'];
// Same cap as the parser applies to every extracted string.
const STRING_CAP = 300;
// Well-known folders whose messages are the user's own mail, never a vendor's.
const EXCLUDED_FOLDERS = ['sentitems', 'drafts'];
const NUMBER_KINDS = ['receiptNumber', 'invoiceNumber', 'orderNumber'];
// Money fields dropped from a duplicate so totals are never counted twice.
const MONEY_FIELDS = ['amount', 'currency', 'taxAmount', 'amounts', 'amountSource', 'amountMismatch', 'currencyMismatch'];
// Save actions meaning this run wrote a new file (so removing it loses nothing).
const CREATED_ACTIONS = new Set(['saved', 'versioned']);
const MB = 1024 * 1024;
// Largest PDF attachment downloaded just to read its product before claiming.
const MAX_PRODUCT_PDF_BYTES = 15 * MB;
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

function capStr(value) {
  if (typeof value !== 'string' || value.length <= STRING_CAP) return value;
  return `${value.slice(0, STRING_CAP)}…`;
}

// Deterministic in-run dedup: 2nd same-named file becomes "name (2).pdf" etc.
// Names are compared case-insensitively, as most receipt folders' file systems do.
function dedupeFilename(fileName, usedNames) {
  const key = fileName.toLowerCase();
  const count = usedNames.get(key) || 0;
  usedNames.set(key, count + 1);
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

/** A receipt/invoice/order number reduced to characters safe in any file name ('' when none). */
function fileNameToken(value) {
  if (typeof value !== 'string') return '';
  return value.trim().replace(/[^A-Za-z0-9._-]+/g, '-').replace(/^[-.]+|[-.]+$/g, '').slice(0, 40);
}

const ISO_DATE = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2})?/;

function escapeODataString(value) {
  return String(value).replace(/'/g, "''");
}

function splitAddress(address) {
  const lower = String(address || '').toLowerCase().trim();
  const at = lower.lastIndexOf('@');
  return at === -1 ? { local: lower, domain: '' } : { local: lower.slice(0, at), domain: lower.slice(at + 1) };
}

/**
 * True when `address` is the rule's sender with extra plus tags appended, so a
 * rule for billing@example.com still matches billing+statements+acct_1@example.com.
 * Tags already present in the rule must match: payment processors send every
 * merchant's receipts from one mailbox and tell them apart only by the tag
 * (billing+statements+acct_A@ vs billing+statements+acct_B@), so stripping all
 * tags would make each merchant's rule claim the other's receipts.
 */
function isPlusExtensionOf(address, ruleFrom) {
  const sender = splitAddress(address);
  const rule = splitAddress(ruleFrom);
  if (!sender.local || !rule.local || sender.domain !== rule.domain) return false;
  return sender.local === rule.local || sender.local.startsWith(`${rule.local}+`);
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

/** Validation message for a malformed vendor rule, or null. Runs before any Graph call or file write. */
function ruleError(rule, periodStart) {
  if (!rule || typeof rule !== 'object' || Array.isArray(rule)) return 'each rule must be an object';
  if (typeof rule.vendor !== 'string' || !rule.vendor.trim()) return 'each rule needs a vendor name (a non-empty string)';
  const name = `Rule '${rule.vendor}'`;
  for (const field of STRING_RULE_FIELDS) {
    if (rule[field] !== undefined && rule[field] !== null && typeof rule[field] !== 'string') {
      return `${name}: ${field} must be a string`;
    }
  }
  if (!rule.from && !rule.subjectContains) return `${name} needs a from and/or subjectContains matcher`;
  if (rule.prefer !== undefined && rule.prefer !== null && !PREFER.includes(rule.prefer)) {
    return `${name}: prefer must be one of ${PREFER.join(', ')}`;
  }
  for (const field of ['productContains', 'productExcludes']) {
    const problem = productTermsError(rule[field]);
    if (problem) return `${name}: ${field} ${problem}`;
  }
  if (rule.filenameTemplate) {
    try {
      sanitiseFilename(renderFilenameTemplate(rule.filenameTemplate, { vendor: rule.vendor, date: periodStart }));
    } catch (error) {
      return `${name}: filenameTemplate does not give a usable file name (${error.message})`;
    }
  }
  return null;
}

/** Validation message for destDir (same containment rules as every save), or null. */
function destDirError(destDir) {
  if (destDir === undefined || destDir === null) return null;
  if (typeof destDir !== 'string') return 'must be a directory path string';
  try {
    const { dir } = resolveDestination(destDir, 'receipt.pdf');
    if (fs.existsSync(dir) && !fs.statSync(dir).isDirectory()) return `is not a directory: ${dir}`;
  } catch (error) {
    return error.message;
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
    } else if (isPlusExtensionOf(fromAddress, rule.from)) {
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

/** A receipt's receipt/invoice/order numbers, trimmed and lower-cased, by kind. */
function numbersOf(receipt) {
  const numbers = {};
  for (const kind of NUMBER_KINDS) {
    const value = typeof receipt?.[kind] === 'string' ? receipt[kind].trim().toLowerCase() : '';
    if (value) numbers[kind] = value;
  }
  return numbers;
}

/** True when two known totals cannot belong to the same charge. */
function totalsDiffer(a, b) {
  if (typeof a.amount !== 'number' || typeof b.amount !== 'number') return false;
  if (Math.abs(a.amount - b.amount) > 0.005) return true;
  return Boolean(a.currency && b.currency && a.currency !== b.currency);
}

/**
 * True when all `records` (receipts or manifest entries) can be copies of one
 * receipt: no receipt, invoice or order number has two different values among
 * them and no two known totals differ. Checked over the whole resulting group,
 * so a copy carrying fewer details (a reply with only the order number, an
 * entry with no number) can never bridge two distinct charges.
 */
function canBeOneReceipt(records) {
  const numbers = records.map(numbersOf);
  for (const kind of NUMBER_KINDS) {
    if (new Set(numbers.map(n => n[kind]).filter(Boolean)).size > 1) return false;
  }
  return records.every((a, i) => records.slice(i + 1).every(b => !totalsDiffer(a, b)));
}

function formatTotal(receipt) {
  return [receipt.amount, receipt.currency].filter(v => v !== null && v !== undefined).join(' ');
}

/**
 * Group copies of one receipt within each rule: members share a receipt,
 * invoice or order number, and the group as a whole never holds two different
 * numbers of one kind or two different known totals (two charges on one order
 * are two receipts). A receipt number shared by different totals leaves a
 * numberConflict note on both sides.
 */
function groupCopies(owned) {
  const groups = [];
  const noteConflict = (item, other) => {
    const number = item.receipt.receiptNumber.trim();
    item.numberConflicts.push(
      `receiptNumber ${number} is also on message ${other.message.id} with a different total ` +
      `(${formatTotal(other.receipt)} vs ${formatTotal(item.receipt)}); kept as separate receipts`
    );
  };
  const receiptsOf = (members) => members.map(m => m.receipt);
  for (const item of owned) {
    item.numbers = numbersOf(item.receipt);
    item.numberConflicts = item.numberConflicts || [];
    let home = null;
    for (const group of [...groups]) {
      if (group.ruleIndex !== item.ruleIndex) continue;
      const shares = NUMBER_KINDS.some(kind =>
        item.numbers[kind] && group.members.some(m => m.numbers[kind] === item.numbers[kind]));
      if (!shares) continue;
      // Once the item has a group, a second group it shares a number with is
      // merged into it, again only if the result can still be one receipt.
      if (canBeOneReceipt(receiptsOf(home ? [...home.members, ...group.members] : [...group.members, item]))) {
        if (home) {
          home.members.push(...group.members);
          groups.splice(groups.indexOf(group), 1);
        } else {
          home = group;
          group.members.push(item);
        }
        continue;
      }
      for (const other of group.members) {
        if (!item.numbers.receiptNumber || other.numbers.receiptNumber !== item.numbers.receiptNumber) continue;
        if (!totalsDiffer(other.receipt, item.receipt)) continue;
        noteConflict(item, other);
        noteConflict(other, item);
      }
    }
    if (!home) groups.push({ ruleIndex: item.ruleIndex, members: [item] });
  }
  return groups;
}

/** duplicate_of label: the number linking a copy to its representative (receipt, else invoice, else order). */
function duplicateLabel(item, representative) {
  const kind = NUMBER_KINDS.find(k => item.numbers[k] && item.numbers[k] === representative.numbers[k]) ||
    NUMBER_KINDS.find(k => representative.numbers[k]);
  return `${kind}:${representative.receipt[kind].trim()}`;
}

function receivedTime(item) {
  const t = Date.parse(item.message.receivedDateTime);
  return Number.isNaN(t) ? Infinity : t;
}

/**
 * One file name per representative, from the rule's template. Receipts whose
 * names collide in this run (one vendor, one day) are told apart by their own
 * receipt/invoice/order number (the first kind unique among them), so a re-run
 * maps each receipt to the same file whatever else arrived meanwhile; an
 * ordinal "(2)" is the last resort for receipts that carry no number, given
 * oldest first so a later arrival never takes an earlier receipt's name.
 * `plain` keeps the template name, under which an earlier run may have saved
 * the receipt before it had a same-day sibling.
 */
function assignFileNames(representatives, vendors, usedNames) {
  const names = new Map();
  const byName = new Map();
  for (const item of representatives) {
    const rule = vendors[item.ruleIndex];
    try {
      const base = renderFilenameTemplate(rule.filenameTemplate || getFilenameTemplate(), {
        vendor: rule.vendor,
        date: item.message.receivedDateTime
      });
      names.set(item, { base, plain: base });
      const key = base.toLowerCase();
      if (!byName.has(key)) byName.set(key, []);
      byName.get(key).push(item);
    } catch (error) {
      names.set(item, { error });
    }
  }
  for (const members of byName.values()) {
    if (members.length < 2) continue;
    const token = (item, kind) => fileNameToken(item.receipt[kind]).toLowerCase();
    for (const item of members) {
      const kind = NUMBER_KINDS.find(k =>
        token(item, k) && members.filter(m => token(m, k) === token(item, k)).length === 1);
      if (kind) names.get(item).base = withSuffix(names.get(item).base, ` ${fileNameToken(item.receipt[kind])}`);
    }
  }
  // Stable sort: same-time receipts keep their order.
  for (const item of [...representatives].sort((a, b) => receivedTime(a) - receivedTime(b))) {
    const naming = names.get(item);
    if (naming.base) naming.fileName = dedupeFilename(naming.base, usedNames);
  }
  return names;
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
      Boolean(rule?.from && !isPlusExtensionOf(sender, rule.from))
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
  if (receipt.productSource) fields.productSource = receipt.productSource;
  return fields;
}

/**
 * Prefer the labelled total printed in the saved PDF over the email body's.
 * Reads the file only if it still holds this receipt's bytes. Never throws:
 * an unreadable PDF leaves the body amounts in place with a note.
 */
async function applyPdfAmounts(entry, saved, rules, parsePdf = parseReceiptPdf) {
  let pdf;
  try {
    const bytes = await fs.promises.readFile(saved.savedPath);
    if (saved.sha256 && sha256Hex(bytes) !== saved.sha256) {
      throw new Error('the file no longer holds the bytes saved for this receipt');
    }
    pdf = await parsePdf(bytes, rules);
  } catch (error) {
    entry.amountSource = 'body';
    entry.pdfParseNote = `saved PDF could not be parsed (${error?.message || error}); amounts are from the email body`;
    return;
  }
  const bodyTotal = typeof entry.amount === 'number' ? entry.amount : null;
  const pdfTotal = typeof pdf.amounts?.total === 'number' ? pdf.amounts.total : null;
  // totalLabel null: no total/amount paid/amount due label, the figure is a guess.
  // A parser that predates the label (undefined) is taken as labelled.
  const guessed = pdfTotal !== null && pdf.amounts.totalLabel === null;
  if (guessed) entry.pdfTotalGuessed = true;
  if (pdfTotal !== null && bodyTotal !== null && Math.abs(bodyTotal - pdfTotal) > 0.005) {
    entry.amountMismatch = { body: bodyTotal, pdf: pdfTotal };
  }
  // A PDF total of zero ("Balance due $0.00") is what is left to pay, not the
  // charge, so it never replaces a positive email total.
  const zeroOverCharge = pdfTotal !== null && Math.abs(pdfTotal) <= 0.005 && bodyTotal !== null && bodyTotal > 0.005;
  if (pdfTotal !== null && !(guessed && bodyTotal !== null) && !zeroOverCharge) {
    entry.amount = pdfTotal;
    entry.currency = pdf.amounts.currency || entry.currency;
    entry.taxAmount = pdf.amounts.tax ?? null;
    entry.amounts = { ...pdf.amounts };
    entry.amountSource = 'pdf';
  } else {
    entry.amountSource = 'body';
  }
  if (!entry.productLabel && pdf.productLabel) {
    entry.productLabel = capStr(pdf.productLabel);
    entry.productSource = 'pdf';
  }
  if (!entry.invoiceNumber && pdf.invoiceNumber) entry.invoiceNumber = capStr(pdf.invoiceNumber);
  if (!entry.receiptNumber && pdf.receiptNumber) entry.receiptNumber = capStr(pdf.receiptNumber);
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
  if (saved.existingDiffers) entry.existingDiffers = true;
}

/**
 * Product label and line items read from the message's PDF attachment, for a
 * receipt whose email body names no product. Memoised per message (ctx) and
 * prefer; a failure comes back as { error } and is never fatal.
 */
async function pdfProduct(ctx, prefer) {
  if (ctx.byPrefer.has(prefer)) return ctx.byPrefer.get(prefer);
  let result;
  try {
    if (!ctx.attachments) {
      const list = throwIfMcpError(await ctx.graphApiClient.makeRequest(
        `${ctx.base}/messages/${ctx.messageId}/attachments`,
        { select: ATTACHMENT_LIST_SELECT }
      ));
      ctx.attachments = list.value || [];
    }
    const pdf = selectPdfAttachment(ctx.attachments, prefer);
    if (!pdf) throw new Error('no PDF attachment found');
    // Read before the message is known to belong to any rule, so it is bounded
    // by the size Graph lists rather than downloaded whatever its size.
    if (typeof pdf.size === 'number' && pdf.size > MAX_PRODUCT_PDF_BYTES) {
      throw new Error(`PDF attachment is too large to read for its product (${(pdf.size / MB).toFixed(1)} MB > 15 MB)`);
    }
    const { buffer } = await fetchAttachmentContent(
      ctx.graphApiClient, ctx.base, ctx.messageId, pdf.id, { attachmentName: pdf.name }
    );
    const parsed = await ctx.parsePdf(buffer, ctx.rules);
    result = { productLabel: capStr(parsed?.productLabel || null), lineItems: parsed?.lineItems || [] };
  } catch (error) {
    result = { error: mcpErrorText(error) };
    logCollectEvent({ stage: 'claim-pdf-product', outcome: 'error', messageId: ctx.messageId, error: result.error });
  }
  ctx.byPrefer.set(prefer, result);
  return result;
}

/**
 * A receipt whose file name gained a number or ordinal since an earlier run
 * (it now has a same-day sibling) may already be on disk under the plain
 * template name or one of its numbered versions. When the file just created
 * holds the same document as one of those (identical bytes, or a PDF differing
 * only in its render timestamps), the new file is removed and the existing one
 * is reported as skipped, so a re-run never keeps two copies of one receipt.
 * Files already reported for another receipt in this run are never reused.
 */
async function reuseEarlierName(saved, destDir, plainName, claimedPaths) {
  const candidates = existingVersionPaths(destDir, plainName)
    .filter(p => p !== saved.savedPath && !claimedPaths.has(p));
  if (candidates.length === 0) return saved;
  const match = await findSameDocument(await fs.promises.readFile(saved.savedPath), candidates);
  if (!match) return saved;
  await fs.promises.rm(saved.savedPath, { force: true });
  return { ...saved, savedPath: match.path, action: 'skipped', sha256: match.sha256, size: match.size };
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
 * else any PDF attachment, else none; ties go to the earliest received. A copy
 * with no known total (e.g. a reply quoting only the order number) is never
 * chosen over a member that carries both a PDF and the total: the copy's own
 * attachment may belong to another charge.
 */
async function chooseRepresentative(members, graphApiClient, base) {
  const hasTotal = (m) => typeof m.receipt.amount === 'number';
  const pool = members.some(m => m.receipt.hasPdfAttachment && hasTotal(m)) ? members.filter(hasTotal) : members;
  const rank = new Map(pool.map(m => [m, m.receipt.hasPdfAttachment ? 1 : 0]));
  const withPdf = pool.filter(m => m.receipt.hasPdfAttachment);
  if (withPdf.length > 1) {
    for (const m of withPdf) {
      if (await hasInvoiceAttachment(graphApiClient, base, m.message.id)) rank.set(m, 2);
    }
  }
  return [...pool].sort((a, b) => rank.get(b) - rank.get(a) || receivedTime(a) - receivedTime(b))[0];
}

export async function collectReceiptsTool(authManager, args, deps = {}) {
  const { periodStart, periodEnd, vendors, destDir, mailbox } = args || {};
  const onExisting = args?.onExisting ?? 'skip';

  if (!periodStart || !ISO_DATE.test(periodStart)) {
    return createValidationError('periodStart', 'Required ISO date or date-time, e.g. 2026-06-01 or 2026-06-01T00:00:00Z');
  }
  if (!periodEnd || !ISO_DATE.test(periodEnd)) {
    return createValidationError('periodEnd', 'Required ISO date or date-time, e.g. 2026-06-30 or 2026-06-30T23:59:59Z');
  }
  if (!Array.isArray(vendors) || vendors.length === 0) {
    return createValidationError('vendors', 'At least one vendor rule is required');
  }
  // Everything that could otherwise fail per receipt, or abort after files
  // were written, is checked here, before any Graph call.
  for (const rule of vendors) {
    const problem = ruleError(rule, periodStart);
    if (problem) return createValidationError('vendors', problem);
  }
  if (!ON_EXISTING.includes(onExisting)) {
    return createValidationError('onExisting', `must be one of ${ON_EXISTING.join(', ')}`);
  }
  const destProblem = destDirError(destDir);
  if (destProblem) return createValidationError('destDir', destProblem);

  try {
    const rules = loadReceiptRules();
    const parsePdf = deps.parsePdf || parseReceiptPdf;
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
    const unmatched = [];
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

      // Zero hits with a from rule: retry on the rule's local part as a prefix
      // and keep only senders that are the rule's address plus extra tags.
      if (messages.length === 0 && rule.from) {
        try {
          const localBase = splitAddress(rule.from).local;
          const retried = await searchMessages(graphApiClient, base, [
            ...dateFilters,
            `startswith(from/emailAddress/address,'${escapeODataString(localBase)}')`
          ]);
          messages = retried.filter(m => isPlusExtensionOf(senderOf(m), rule.from));
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
      const rejections = [];
      let pdfCtx = null;
      let pdfProductError = null;
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
        let product = evaluateProduct(state.rule, receipt);
        let fromPdf = null;
        // No product in the email body: read it from the PDF before deciding.
        if (product.verdict === 'unknown' && receipt.hasPdfAttachment) {
          pdfCtx ??= { graphApiClient, base, messageId: message.id, rules, parsePdf, attachments: null, byPrefer: new Map() };
          const read = await pdfProduct(pdfCtx, state.rule.prefer || 'invoice');
          if (read.error) {
            pdfProductError = read.error;
          } else if (read.productLabel || read.lineItems.length > 0) {
            fromPdf = read;
            product = evaluateProduct(state.rule, read);
          }
        }
        if (product.verdict === 'reject') {
          state.excludedByProduct++;
          rejections.push({ index, productLabel: fromPdf ? fromPdf.productLabel : receipt.productLabel });
          continue;
        }
        claims.push({ index, signals, product, fromPdf });
      }
      if (claims.length === 0) {
        if (rejections.length > 0) {
          // Every rule that matched the sender/subject rejected the product: reported, never dropped.
          unmatched.push({
            messageId: message.id,
            date: message.receivedDateTime,
            subject: capStr(message.subject ?? null),
            sender: senderOf(message),
            productLabel: rejections[0].productLabel ?? null,
            rejectedBy: rejections.map(r => vendorOf(r.index))
          });
          logCollectEvent({
            stage: 'claim', outcome: 'unmatched', messageId: message.id, rejectedBy: rejections.map(r => vendorOf(r.index))
          });
        }
        continue;
      }
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
        const pdfRead = contenders.find(c => c.fromPdf)?.fromPdf;
        const shown = pdfRead
          ? { ...receipt, productLabel: pdfRead.productLabel || receipt.productLabel, productSource: 'pdf' }
          : receipt;
        const ambiguous = {
          vendor: single ? single.vendor : null,
          status: 'ambiguous',
          candidates: contenders.map(c => vendorOf(c.index)),
          reason: resolved.reason,
          date: message.receivedDateTime,
          sourceMessageId: message.id,
          ...senderFields(single, message, receipt),
          ...receiptFields(shown)
        };
        if (resolved.reason === PRODUCT_UNKNOWN_REASON && pdfProductError) ambiguous.pdfProductError = pdfProductError;
        buckets[contenders[0].index].push(ambiguous);
        logCollectEvent({
          stage: 'claim', status: 'ambiguous', candidates: contenders.map(c => vendorOf(c.index)),
          messageId: message.id, reason: resolved.reason
        });
        continue;
      }

      const { winner } = resolved;
      involved.add(winner.index);
      owned.push({
        message,
        receipt: winner.fromPdf
          ? { ...receipt, productLabel: winner.fromPdf.productLabel || receipt.productLabel, productSource: 'pdf' }
          : receipt,
        ruleIndex: winner.index,
        matchedBy: [...winner.signals, ...winner.product.signals]
      });
    }

    // --- Dedupe copies of one receipt -------------------------------------
    const groups = groupCopies(owned);
    for (const group of groups) {
      group.representative = group.members.length > 1
        ? await chooseRepresentative(group.members, graphApiClient, base)
        : group.members[0];
      for (const member of group.members) member.group = group;
    }

    // A billing link shared by several distinct receipts is a generic portal
    // link, not a per-receipt invoice. Copies of one receipt count once.
    const linkUse = new Map();
    for (const item of owned) {
      const url = item.receipt.billingPdfUrl;
      if (!url) continue;
      if (!linkUse.has(url)) linkUse.set(url, new Set());
      linkUse.get(url).add(item.group);
    }

    const fileNames = assignFileNames(groups.map(g => g.representative), vendors, usedNames);

    // --- Save -------------------------------------------------------------
    const saveOrder = [];
    // Files already reported for a receipt in this run.
    const claimedPaths = new Set();
    for (const item of owned) {
      const { message, receipt } = item;
      const { rule, searchNote } = states[item.ruleIndex];
      const { group } = item;
      const entry = {
        vendor: rule.vendor,
        status: null,
        date: message.receivedDateTime,
        sourceMessageId: message.id,
        matchedBy: item.matchedBy,
        ...senderFields(rule, message, receipt)
      };
      if (searchNote) entry.searchNote = searchNote;
      if (item.numberConflicts.length > 0) entry.numberConflict = item.numberConflicts.join('; ');
      buckets[item.ruleIndex].push(entry);

      if (group.representative !== item) {
        entry.status = 'duplicate';
        entry.duplicate_of = duplicateLabel(item, group.representative);
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
        const naming = fileNames.get(item);
        if (naming.error) throw naming.error;
        fileName = naming.fileName;
        const reuse = async (result) => {
          if (onExisting !== 'skip' || !CREATED_ACTIONS.has(result.action) ||
              naming.plain.toLowerCase() === fileName.toLowerCase()) return result;
          try {
            return await reuseEarlierName(result, destDir, naming.plain, claimedPaths);
          } catch (error) {
            // Best-effort: the receipt keeps the file just written under its new name.
            logCollectEvent({ vendor: rule.vendor, stage: 'reuse-name', outcome: 'error', messageId: message.id, error: error.message });
            return result;
          }
        };

        if (formal === 'attachment') {
          saved = await reuse(await saveAttachmentCore(graphApiClient, {
            messageId: message.id,
            prefer: rule.prefer || 'invoice',
            destDir,
            fileName,
            onExisting,
            mailbox
          }, rules));
          recordSaved(entry, saved, 'attachment', true);
          entry.attachmentId = saved.attachmentId;
          entry.attachmentName = saved.attachmentName;
          await applyPdfAmounts(entry, saved, rules, parsePdf);
        } else if (formal === 'link') {
          saved = await reuse(await fetchBillingPdfCore(graphApiClient, {
            url: receipt.billingPdfUrl,
            destDir,
            fileName,
            onExisting,
            mailbox
          }, deps));
          recordSaved(entry, saved, 'link', true);
          await applyPdfAmounts(entry, saved, rules, parsePdf);
        } else {
          saved = await reuse(await renderEmailPdfCore(graphApiClient, {
            messageId: message.id,
            destDir,
            fileName,
            onExisting,
            mailbox
          }, deps));
          recordSaved(entry, saved, 'rendered', false);
          entry.amountSource = 'body';
        }
        claimedPaths.add(saved.savedPath);
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
        if (Array.isArray(failure.details?.attemptedRoutes)) entry.attemptedRoutes = [...failure.details.attemptedRoutes];
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
            const renderedName = dedupeFilename(withSuffix(fileName, ' (rendered)'), usedNames);
            const rendered = await renderEmailPdfCore(graphApiClient, {
              messageId: message.id,
              destDir,
              fileName: renderedName,
              onExisting,
              mailbox
            }, deps);
            claimedPaths.add(rendered.savedPath);
            entry.renderedFallback = {
              savedPath: rendered.savedPath,
              sha256: rendered.sha256,
              action: rendered.action,
              method: 'rendered',
              isFormalInvoice: false,
              reason: `the ${formal} could not be fetched; this is a rendering of the email body, not the formal invoice`
            };
            if (rendered.existingDiffers) entry.renderedFallback.existingDiffers = true;
            // What the run holds for this receipt is a rendering, never the invoice.
            entry.method = 'rendered';
            entry.isFormalInvoice = false;
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
    // Identical bytes mean one receipt only when the entries could be one
    // receipt by the same rule as copies (no two different numbers of a kind,
    // no two different known totals across the whole group): a vendor may
    // reuse one PDF for distinct receipts, and those are kept (sameFileAs
    // names the other). Each group's first entry is the one kept.
    const groupsBySha = new Map();
    // Numbers and total as saved: a duplicate's money fields are removed below.
    const facts = (e) => ({ ...Object.fromEntries(NUMBER_KINDS.map(k => [k, e[k]])), amount: e.amount, currency: e.currency });
    for (const entry of saveOrder) {
      if (!entry.sha256 || (entry.status !== 'saved' && entry.status !== 'skipped')) continue;
      const shaGroups = groupsBySha.get(entry.sha256) || [];
      const home = shaGroups.find(g => canBeOneReceipt([...g.members, facts(entry)]));
      if (!home) {
        if (shaGroups.length > 0) entry.sameFileAs = shaGroups[0].first.sourceMessageId;
        groupsBySha.set(entry.sha256, [...shaGroups, { first: entry, members: [facts(entry)] }]);
        continue;
      }
      home.members.push(facts(entry));
      const { first } = home;
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
    summary.unmatched = unmatched.length;
    // Separate copies: the response serialiser treats a repeated object as circular.
    const failed = manifest.filter(e => e.status === 'fetch_failed').map(e => structuredClone(e));

    return createSafeResponse({
      period: { start: periodStart, end: periodEnd },
      manifest,
      missing,
      failed,
      unmatched,
      summary,
      notes
    });
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to collect receipts');
  }
}
