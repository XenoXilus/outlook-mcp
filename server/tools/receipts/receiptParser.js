/**
 * Structured receipt extraction (FR-3).
 *
 * Parses vendor receipt emails into a compact typed summary — never returns
 * the raw HTML body. Built-in heuristics are fully generic (labelled amounts,
 * receipt/invoice numbers, "receipt from X" subjects); site-specific sender
 * and product-label rules come from RECEIPT_RULES_PATH (see receiptRules.js).
 * All extractors are heuristic and return null when unsure.
 */

import { stripHtml } from '../../utils/textUtils.js';

const CURRENCY_SYMBOLS = { '£': 'GBP', '$': 'USD', '€': 'EUR' };
const AMOUNT_PATTERN = '([£$€])\\s?(\\d[\\d,]*\\.\\d{2})';

export function detectVendor(fromAddress, subject = '', vendorSenders = []) {
  const from = (fromAddress || '').toLowerCase();
  for (const { pattern, vendor } of vendorSenders) {
    if (pattern.test(from)) return vendor;
  }
  // Payment processors send on behalf of vendors: "Your receipt from <Vendor> #1234"
  const fromSubject = subject.match(/(?:receipt|invoice) from ([A-Za-z0-9][A-Za-z0-9 .&-]*?)(?:,| #|$)/i);
  if (fromSubject) return fromSubject[1].trim();
  return null;
}

export function extractAmount(text) {
  const labelled = text.match(new RegExp(`(?:amount paid|amount due|total)[:\\s]*${AMOUNT_PATTERN}`, 'i'));
  const match = labelled || text.match(new RegExp(AMOUNT_PATTERN));
  if (!match) return { amount: null, currency: null };
  return {
    amount: parseFloat(match[2].replace(/,/g, '')),
    currency: CURRENCY_SYMBOLS[match[1]] || null
  };
}

export function extractTaxAmount(text) {
  const match = text.match(new RegExp(`(?:VAT|tax)[^\\n£$€]*${AMOUNT_PATTERN}`, 'i'));
  return match ? parseFloat(match[2].replace(/,/g, '')) : null;
}

export function extractReceiptNumber(text) {
  const match = text.match(/receipt (?:number|#)\s*:?\s*#?([A-Z0-9][\w-]+)/i);
  return match ? match[1] : null;
}

export function extractInvoiceNumber(text) {
  const match = text.match(/invoice (?:number|#)\s*:?\s*#?([A-Z0-9][\w-]+)/i);
  return match ? match[1] : null;
}

export function extractProductLabel(text, productLabels = []) {
  for (const pattern of productLabels) {
    const match = text.match(pattern);
    if (match) return match[0].trim();
  }
  return null;
}

export function extractBillingPdfUrl(html, allowlist) {
  const urls = (html || '').match(/https:\/\/[^\s"'<>)]+/g) || [];
  // Decode the two most common ampersand entities that appear in raw HTML href
  // attributes — &amp; and its numeric synonym &#38; — so multi-param signed
  // URLs (e.g. Stripe pre-signed links) are returned with valid & separators.
  const decoded = urls.map(u => u.replace(/&amp;/g, '&').replace(/&#38;/g, '&'));
  const allowed = decoded.filter(u => {
    try {
      return allowlist.includes(new URL(u).hostname.toLowerCase());
    } catch {
      return false;
    }
  });
  return allowed.find(u => /pdf/i.test(u)) || allowed[0] || null;
}

export function selectPdfAttachment(attachments, prefer = 'invoice') {
  const pdfs = (attachments || []).filter(a =>
    !a.isInline &&
    (/\.pdf$/i.test(a.name || '') || (a.contentType || '').toLowerCase() === 'application/pdf')
  );
  if (pdfs.length === 0) return null;
  if (prefer === 'first') return pdfs[0];
  const wanted = prefer === 'receipt' ? /^receipt/i : /^invoice/i;
  return pdfs.find(a => wanted.test(a.name || '')) || pdfs[0];
}

const STRING_CAP = 300;
const ELLIPSIS = '…'; // …

function capStr(s) {
  if (typeof s !== 'string' || s.length <= STRING_CAP) return s;
  return s.slice(0, STRING_CAP) + ELLIPSIS;
}

// FR-3 size budget: everything except billingPdfUrl must fit in 4096 bytes.
// billingPdfUrl is exempt — it is fetched verbatim by the FR-2 link fallback;
// truncating it corrupts pre-signed Stripe URLs (fixed in commit 57d78df).
const SIZE_BUDGET = 4096;

export function parseReceiptMessage(message, attachments = [], allowlist = [], rules = { vendorSenders: [], productLabels: [] }) {
  const html = message.body?.content || '';
  const text = message.body?.contentType?.toLowerCase() === 'html' ? stripHtml(html) : html;
  const fromAddress = message.from?.emailAddress?.address || '';
  const { amount, currency } = extractAmount(text);
  const nonInline = attachments.filter(a => !a.isInline);

  const result = {
    vendor: capStr(detectVendor(fromAddress, message.subject || '', rules.vendorSenders)),
    subject: capStr(message.subject || null),
    receivedDate: message.receivedDateTime || null,
    receiptNumber: capStr(extractReceiptNumber(text)),
    invoiceNumber: capStr(extractInvoiceNumber(text)),
    amount,
    currency,
    taxAmount: extractTaxAmount(text),
    productLabel: capStr(extractProductLabel(text, rules.productLabels)),
    billingPdfUrl: extractBillingPdfUrl(html, allowlist), // never truncate — used verbatim as fetch URL
    hasPdfAttachment: selectPdfAttachment(attachments, 'first') !== null,
    attachmentIds: nonInline.map(a => a.id).slice(0, 20)
  };

  // Overall size guard (FR-3): trim attachmentIds (re-fetchable via Graph) until
  // the result excluding billingPdfUrl fits within SIZE_BUDGET bytes.
  // truncated is scoped to the attachmentIds guard only (capStr truncations are not tracked).
  const sizeOf = () => JSON.stringify({ ...result, billingPdfUrl: undefined }).length;
  if (sizeOf() > SIZE_BUDGET) {
    while (result.attachmentIds.length > 0 && sizeOf() > SIZE_BUDGET) {
      result.attachmentIds.pop();
    }
    if (sizeOf() > SIZE_BUDGET) {
      result.attachmentIds = []; // pathological: drop all
    }
    result.truncated = true;
  }

  return result;
}
