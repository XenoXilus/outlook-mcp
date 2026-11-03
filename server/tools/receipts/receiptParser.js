/**
 * Structured receipt extraction.
 *
 * Parses vendor receipt emails into a compact typed summary — never returns
 * the raw HTML body. Built-in heuristics are fully generic (labelled amounts,
 * receipt/invoice numbers, "receipt from X" subjects); site-specific sender
 * and product-label rules come from RECEIPT_RULES_PATH (see receiptRules.js).
 * All extractors are heuristic and return null when unsure.
 *
 * The totals parser (parseReceiptTotals) works on plain text and is shared with
 * receiptPdfText.js, so email bodies and PDF text yield the same shape.
 *
 * Input is untrusted: text is capped (MAX_PARSE_CHARS) and very long lines are
 * split (MAX_LINE_CHARS) before any regex work, regexes avoid ambiguous
 * repetition, and every parsed string that escapes is length-capped.
 */

import { stripHtml } from '../../utils/textUtils.js';

// ---------------------------------------------------------------------------
// Dates (used to reject "vendors" that are really dates, and to skip period lines)
// ---------------------------------------------------------------------------

const MONTH_NAMES = 'jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?';
const WEEKDAY_NAMES = 'mon(?:day)?|tue(?:s(?:day)?)?|wed(?:nesday)?|thu(?:r(?:s(?:day)?)?)?|fri(?:day)?|sat(?:urday)?|sun(?:day)?';

/** True for strings that are only a date or date range ("21 Mar 2026", "March 21", "2026-03-21", "Feb 17–Mar 17, 2026"). */
export function looksLikeDate(value) {
  const t = String(value || '').trim();
  if (!t) return false;
  if (/^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}(?:[T\s][\d:.]+Z?)?$/.test(t)) return true;
  if (/^\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}$/.test(t)) return true;
  if (!/\d/.test(t)) return false;
  const month = new RegExp(`\\b(?:${MONTH_NAMES})\\b\\.?`, 'gi');
  if (!month.test(t)) return false;
  const rest = t
    .replace(month, ' ')
    .replace(new RegExp(`\\b(?:${WEEKDAY_NAMES})\\b\\.?`, 'gi'), ' ')
    .replace(/\b\d{1,2}(?:st|nd|rd|th)\b/gi, ' ')
    .replace(/\b(?:of|to|at|am|pm|utc|gmt)\b/gi, ' ');
  return !/[A-Za-z]/.test(rest);
}

// ---------------------------------------------------------------------------
// Amounts
// ---------------------------------------------------------------------------

const SYMBOL_CURRENCY = {
  'CA$': 'CAD', 'C$': 'CAD', 'AU$': 'AUD', 'A$': 'AUD', 'NZ$': 'NZD',
  'US$': 'USD', 'HK$': 'HKD', '$': 'USD', '£': 'GBP', '€': 'EUR'
};
const ISO_CODES = 'USD|GBP|EUR|CAD|AUD|NZD|HKD|CHF|JPY|SEK|NOK|DKK|SGD|INR|CNY|ZAR|MXN|BRL|PLN|AED';
const SIGN = '[-\\u2212\\u2013\\u2012\\u2011]'; // hyphen-minus, minus sign, en dash, figure dash, non-breaking hyphen
const NUMBER = '\\d{1,3}(?:,\\d{3})+(?:\\.\\d+)?|\\d{1,3}(?:\\.\\d{3})+,\\d{2}|\\d+,\\d{2}(?!\\d)|\\d+(?:\\.\\d+)?';
const SYMBOL = '(?:(?<![A-Za-z])(?:CA|AU|NZ|US|HK|C|A)\\$|[£€$])';
const ISO = `(?<![A-Za-z])(?:${ISO_CODES})(?![A-Za-z])`;

// A sign only counts when attached to the amount ("-$5.00"), so "Plan - $24.00" stays positive.
const AMOUNT_RE = new RegExp(
  `(?<sign1>(?<!\\w)${SIGN})?(?:` +
    `(?<sym>${SYMBOL})\\s?(?<sign2>${SIGN})?\\s?(?<n1>${NUMBER})` +
    `|(?<pre>${ISO})\\s?(?<sign3>${SIGN})?\\s?(?<n2>${NUMBER})` +
    `|(?<![\\w.,])(?<n3>${NUMBER})\\s?(?<post>${ISO}|€)` +
  `)(?:\\s?(?<suffix>${ISO}))?`,
  'g'
);

function parseNumber(s) {
  if (/^\d{1,3}(?:\.\d{3})+,\d{2}$/.test(s) || /^\d+,\d{2}$/.test(s)) {
    return parseFloat(s.replace(/\./g, '').replace(',', '.')); // European "1.234,56" / "12,50"
  }
  return parseFloat(s.replace(/,/g, ''));
}

/** Parenthesis depth after s[from..to), continuing from `depth` (a ')' never goes below 0). */
function advanceParenDepth(s, from, to, depth) {
  for (let i = from; i < to; i++) {
    const ch = s.charCodeAt(i);
    if (ch === 40) depth++; // (
    else if (ch === 41 && depth > 0) depth--; // )
  }
  return depth;
}

/** All currency amounts in a line: value (negative for -/–/(…) forms), currency code, span, inParens. */
function scanAmounts(line) {
  const out = [];
  let depth = 0;
  let depthAt = 0; // depth is the parenthesis depth of line[0..depthAt) — tracked incrementally, not re-sliced per match
  for (const m of line.matchAll(AMOUNT_RE)) {
    depth = advanceParenDepth(line, depthAt, m.index, depth);
    depthAt = m.index;
    const g = m.groups;
    const value = parseNumber(g.n1 ?? g.n2 ?? g.n3);
    if (!Number.isFinite(value)) continue;
    const end = m.index + m[0].length;
    const wrapped = line[m.index - 1] === '(' && line[end] === ')';
    const negative = Boolean(g.sign1 || g.sign2 || g.sign3) || wrapped;
    out.push({
      value: negative ? -value : value,
      currency: g.suffix || (g.post && (SYMBOL_CURRENCY[g.post] || g.post)) || g.pre || SYMBOL_CURRENCY[g.sym] || null,
      index: m.index,
      end,
      inParens: !wrapped && depth > 0
    });
  }
  return out;
}

function maskAmounts(line, amounts) {
  let out = '';
  let pos = 0;
  for (const a of amounts) {
    out += line.slice(pos, a.index) + ' ';
    pos = a.end;
  }
  return (out + line.slice(pos)).replace(/\s+/g, ' ').trim();
}

const round2 = (n) => Math.round(n * 100) / 100;

// ---------------------------------------------------------------------------
// Totals + line items
// ---------------------------------------------------------------------------

const TAX_WORDS = 'VAT|GST|HST|PST|QST|sales tax|tax';
// "(Includes VAT of £2.50)" / "Includes £2.50 VAT" — a tax figure embedded in the total line.
// The gaps are bounded so a line of repeated "includes" cannot make the scan quadratic.
const INCLUDES_TAX_RE = new RegExp(`\\(?\\s*\\bincludes?d?\\b[^()\\n]{0,100}?\\b(?:${TAX_WORDS})\\b[^()\\n]{0,100}\\)?`, 'i');
const TAX_LABEL_RE = new RegExp(`\\b(?:${TAX_WORDS})\\b`, 'i');
const NET_TOTAL_RE = /\btotal\b[^a-z]*\b(?:excl\w*|excluding|ex|before|net of|without)\b|\bnet\s+(?:total|amount)\b|\bexcl(?:uding|\.)?\s+(?:tax|vat|gst)\b/i;
const CREDIT_RE = /\b(?:credits?(?!\s*card)|discount|coupon|promo\w*|rebate|adjustments?|off)\b|% off/i;
// Lines that open the totals section; item parsing stops there.
const TOTALS_START_RE = new RegExp(
  `^(?:sub-?\\s?total\\b|total\\b|amount\\s+(?:paid|due)|balance\\s+due|includes?\\b|(?:${TAX_WORDS})\\b(?:\\s*[-:(].*)?$)`, 'i'
);
// Lines that are headings/metadata rather than a purchased item.
const NOT_AN_ITEM_RE = /^(?:(?:receipt|invoice)\b.*\b(?:from|number|no\b|#)|(?:receipt|invoice)\s*[#:]|paid\b|payment\b|download\b|questions\b|powered\b|order\s*(?:number|no\b|#|id)|description\b|item\b|qty\b|quantity\b|unit price\b|amount\b|date\b|bill(?:ed)? to\b|ship(?:ped)? to\b|your (?:receipt|invoice|order)\b)/i;
// A label that opens with one of these is a credit even when the amount is shown positive
// ("Credit $10.00", "Adjustment $2.00"), so it is never a line item. "Credit card …" is not.
const CREDIT_LEAD_RE = /^(?:credits?(?!\s*card)|discount|coupon|promo\w*|adjustments?)\b/i;
const QTY_LINE_RE = /^(?:qty|quantity)\.?\s*:?\s*(\d[\d,]*(?:\.\d+)?)$/i;

const hasLetters = (s) => /[A-Za-z]/.test(s);
const stripUnitWords = (s) => s.replace(/\/\s*[A-Za-z]+/g, ' ').replace(/\b(?:each|ea)\b/gi, ' ');

// Bounds on untrusted text: a receipt is a few KB; anything past these is noise or an attack.
const MAX_PARSE_CHARS = 200 * 1024;
const MAX_LINE_CHARS = 2000;

/** At most MAX_PARSE_CHARS of the input, as a string. */
function capParseText(text) {
  const s = String(text || '');
  return s.length > MAX_PARSE_CHARS ? s.slice(0, MAX_PARSE_CHARS) : s;
}

/** Split a (whitespace-normalised) line into chunks of at most MAX_LINE_CHARS, at a space when possible. */
function splitLongLine(line) {
  if (line.length <= MAX_LINE_CHARS) return [line];
  const parts = [];
  let rest = line;
  while (rest.length > MAX_LINE_CHARS) {
    let cut = rest.lastIndexOf(' ', MAX_LINE_CHARS);
    if (cut < MAX_LINE_CHARS / 2) cut = MAX_LINE_CHARS;
    parts.push(rest.slice(0, cut).trim());
    rest = rest.slice(cut).trim();
  }
  if (rest) parts.push(rest);
  return parts;
}

function toLines(text) {
  return capParseText(text)
    .replace(/[\u200B-\u200F\u2060\uFEFF\u034F\u00AD]/g, '')
    .replace(/[\u00A0\u2007\u202F]/g, ' ')
    .split(/\r\n|\r|\n/)
    .map(l => l.replace(/\s+/g, ' ').trim())
    .filter(Boolean)
    .flatMap(splitLongLine);
}

function scanLine(raw) {
  let line = raw;
  let includesTax = null;
  const inc = line.match(INCLUDES_TAX_RE);
  if (inc) {
    const embedded = scanAmounts(inc[0]);
    if (embedded.length) {
      includesTax = Math.abs(embedded[0].value);
      line = `${line.slice(0, inc.index)} ${line.slice(inc.index + inc[0].length)}`.trim();
    }
  }
  const amounts = scanAmounts(line);
  const masked = maskAmounts(line, amounts);
  const outside = amounts.filter(a => !a.inParens);
  return {
    line,
    amounts,
    outside,
    masked,
    includesTax,
    letters: hasLetters(masked),
    // Only a number as the whole line (plus "each", "/month"): the value half of a label/value pair.
    valueOnly: outside.length > 0 && !hasLetters(stripUnitWords(masked)),
    // Amounts only inside parentheses ("Allowance ($15.00 off)"): the real value is on the next line.
    labelOnly: amounts.length > 0 && outside.length === 0,
    consumed: false
  };
}

function classifyLabel(label) {
  const l = label.toLowerCase();
  if (/amount\s+paid|paid\s+amount/.test(l)) return 'amountPaid';
  if (/amount\s+due|balance\s+due|amount\s+payable/.test(l)) return 'amountDue';
  if (/sub-?\s?total/.test(l)) return 'subtotal';
  if (NET_TOTAL_RE.test(l)) return 'netTotal';
  if (/\btotal\b.*\bincl/.test(l)) return 'total';
  if (TAX_LABEL_RE.test(l)) return 'tax';
  if (CREDIT_RE.test(l)) return 'credit';
  if (/\btotal\b/.test(l)) return 'total';
  return 'other';
}

const TOTAL_KINDS = new Set(['subtotal', 'netTotal', 'tax', 'total', 'amountPaid', 'amountDue']);

/** A credit/discount line (never a line item): credit-classified and either negative or opening with a credit word. */
function isCreditLabel(label, value) {
  return classifyLabel(label) === 'credit' && (value < 0 || CREDIT_LEAD_RE.test(label));
}

// How the chosen total was found (amounts.totalLabel); a first-amount fallback has no label.
const TOTAL_LABELS = { amountPaid: 'amount paid', total: 'total', amountDue: 'amount due' };

function isDescription(text) {
  return hasLetters(text) && text.length <= 200 && !text.includes('@') &&
    !NOT_AN_ITEM_RE.test(text) && !TOTALS_START_RE.test(text) && !looksLikeDate(text);
}

function cleanDescription(s) {
  let d = s.replace(/^[\s•*·-]+/, '').replace(/[\s:\u2013\u2014-]+$/, '').replace(/\s+/g, ' ').trim();
  if (d.length > 120) d = `${d.slice(0, 119)}${ELLIPSIS}`;
  return d;
}

const MAX_LINE_ITEMS = 10;

function parseLineItems(scanned) {
  const items = [];
  let desc = null;
  let qty = null;
  const push = (description, quantity, amount) => {
    const d = cleanDescription(description);
    if (d && Number.isFinite(amount) && items.length < MAX_LINE_ITEMS) {
      items.push({ description: d, quantity: Number.isFinite(quantity) ? quantity : null, amount });
    }
  };

  for (let i = 0; i < scanned.length; i++) {
    const info = scanned[i];
    const text = info.masked;

    if (!info.amounts.length) {
      const q = text.match(QTY_LINE_RE);
      if (q) { qty = parseNumber(q[1]); continue; }
      if (TOTALS_START_RE.test(text)) break;
      if (looksLikeDate(text)) continue; // period heading between description and values
      desc = isDescription(text) ? text : null;
      continue;
    }

    if (info.labelOnly) {
      if (TOTALS_START_RE.test(text)) break;
      desc = isDescription(text) ? info.line : null;
      continue;
    }

    if (info.valueOnly) {
      if (/\beach\b/i.test(text)) continue; // unit price under the line amount
      const next = scanned[i + 1]?.line || '';
      if (/^(?:paid|due|payment|date)\b/i.test(next)) { desc = null; continue; } // header summary amount
      if (desc && isCreditLabel(desc, info.outside.at(-1).value)) { desc = null; qty = null; continue; } // "Credit" / "$10.00"
      const rowQty = text.match(/^\s*(\d[\d,]*(?:\.\d+)?)(?!\s*%|[\d.,])/);
      if (desc) push(desc, qty ?? (rowQty ? parseNumber(rowQty[1]) : null), info.outside.at(-1).value);
      desc = null;
      qty = null;
      continue;
    }

    // Description and amount(s) on one line.
    const kind = classifyLabel(text);
    if (TOTALS_START_RE.test(text) || TOTAL_KINDS.has(kind)) break;
    const first = info.outside[0];
    let prefix = info.line.slice(0, first.index).trim();
    const trailing = stripUnitWords(info.line.slice(info.outside.at(-1).end));
    if (!prefix || hasLetters(trailing) || !isDescription(prefix) || isCreditLabel(prefix, info.outside.at(-1).value)) {
      desc = null;
      continue;
    }
    let quantity = null;
    const withQty = info.outside.length >= 2 && prefix.match(/^(.*[A-Za-z].*?)\s+(\d[\d,]*(?:\.\d+)?)$/);
    if (withQty) {
      prefix = withQty[1];
      quantity = parseNumber(withQty[2]);
    }
    push(prefix, quantity, info.outside.at(-1).value);
    desc = null;
    qty = null;
  }
  return items;
}

/**
 * Joined, de-duplicated line-item descriptions, at most 300 characters.
 * @param {{description: string}[]} lineItems
 * @returns {string|null}
 */
export function buildProductLabel(lineItems) {
  const seen = [];
  for (const item of lineItems || []) {
    if (item?.description && !seen.includes(item.description)) seen.push(item.description);
  }
  if (seen.length === 0) return null;
  const joined = seen.join('; ');
  return joined.length > STRING_CAP ? `${joined.slice(0, STRING_CAP - 1)}${ELLIPSIS}` : joined;
}

/**
 * Parse totals and line items from receipt/invoice plain text (email body or PDF text).
 * Labels and values may share a line ("Total £72.00") or sit on consecutive lines.
 *
 * @param {string} text
 * @param {'body'|'pdf'} [source='body'] copied into amounts.source
 * amounts.totalLabel says how the total was found: 'amount paid' | 'total' | 'amount due',
 * or null when the total is only a fallback guess (the first amount in the text) or absent.
 *
 * @returns {{ amounts: { subtotal: number|null, tax: number|null, credits: number|null,
 *   total: number|null, totalLabel: 'amount paid'|'total'|'amount due'|null,
 *   amountPaid: number|null, currency: string|null, source: string },
 *   lineItems: { description: string, quantity: number|null, amount: number }[] }}
 */
export function parseReceiptTotals(text, source = 'body') {
  const scanned = toLines(text).map(scanLine);

  const entries = [];
  let seenTotals = false;
  scanned.forEach((info, i) => {
    if (!info.amounts.length || info.labelOnly) return;
    let label = info.masked;
    if (info.valueOnly) {
      const prev = scanned[i - 1];
      label = '';
      if (prev && !prev.consumed && prev.letters && (prev.amounts.length === 0 || prev.labelOnly)) {
        label = prev.masked;
        prev.consumed = true;
      }
    }
    const amount = info.outside.at(-1);
    let kind = classifyLabel(label);
    if (kind === 'credit' && !(seenTotals || CREDIT_LEAD_RE.test(label) || amount.value < 0)) kind = 'other';
    if (kind === 'other' && amount.value < 0 && label) kind = 'credit';
    if (TOTAL_KINDS.has(kind)) seenTotals = true;
    entries.push({ kind, label, value: amount.value, currency: amount.currency });
  });

  const firstOf = (kind) => entries.find(e => e.kind === kind);
  const sumOf = (kind) => {
    const seen = new Set();
    let sum = 0;
    let n = 0;
    for (const e of entries) {
      const key = `${e.label}|${e.value}`;
      if (e.kind !== kind || seen.has(key)) continue;
      seen.add(key);
      sum += kind === 'credit' ? Math.abs(e.value) : e.value;
      n++;
    }
    return n ? round2(sum) : null;
  };

  let tax = sumOf('tax');
  if (tax === null) {
    const embedded = scanned.find(s => s.includesTax !== null);
    if (embedded) tax = round2(embedded.includesTax);
  }

  const paid = firstOf('amountPaid');
  const labelled = paid || firstOf('total') || firstOf('amountDue') || null;
  const chosen = labelled || scanned.flatMap(s => s.amounts).find(a => a.value >= 0) || null;
  const net = firstOf('netTotal');

  const lineItems = parseLineItems(scanned);
  const amounts = {
    subtotal: (firstOf('subtotal') || net)?.value ?? null,
    tax,
    credits: sumOf('credit'),
    total: chosen ? chosen.value : null,
    totalLabel: labelled ? TOTAL_LABELS[labelled.kind] : null,
    amountPaid: paid ? paid.value : null,
    currency: chosen?.currency || entries.find(e => e.currency)?.currency || null,
    source
  };
  return { amounts, lineItems };
}

// ---------------------------------------------------------------------------
// Vendor, subject, forwarded-mail helpers
// ---------------------------------------------------------------------------

export function detectVendor(fromAddress, subject = '', vendorSenders = []) {
  const from = (fromAddress || '').toLowerCase();
  for (const { pattern, vendor } of vendorSenders) {
    if (pattern.test(from)) return vendor;
  }
  // Payment processors send on behalf of vendors: "Your receipt from <Vendor> #1234"
  const candidates = String(subject || '').slice(0, 1000).matchAll(/(?:receipt|invoice) from ([A-Za-z0-9][A-Za-z0-9 .&-]*?)(?:,| #|$)/gi);
  for (const [, candidate] of candidates) {
    const name = candidate.trim();
    if (name && !looksLikeDate(name)) return name; // "Order Receipt from 21 Mar 2026" is a date, not a vendor
  }
  return null;
}

const REPLY_PREFIX_RE = /^\s*(?:(?:re|fwd?|aw|wg)\s*(?:\[\d+\]|\(\d+\))?\s*:\s*)+/i;

/** Strip any number of leading Re:/RE:/Fw:/FW:/Fwd:/AW:/WG: prefixes (with optional "[2]" counters). */
export function normaliseSubject(subject) {
  return String(subject || '').replace(REPLY_PREFIX_RE, '').replace(/\s+/g, ' ').trim();
}

function hasReplyPrefix(subject) {
  return REPLY_PREFIX_RE.test(String(subject || ''));
}

function parseMailbox(raw) {
  const addr = raw.match(/[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+/);
  const address = addr ? addr[0] : null;
  let name = raw.replace(/<[^>]*>|\[mailto:[^\]]*\]|mailto:\S+/gi, ' ');
  if (address) name = name.split(address).join(' ');
  name = name.replace(/[<>[\]]/g, ' ').replace(/\(\s*\)/g, ' ').replace(/\s+/g, ' ')
    .replace(/^["'\s]+|["'\s]+$/g, '').trim();
  if (!name && !address) return null;
  return { name: name ? capStr(name.slice(0, 200)) : null, address: address ? address.slice(0, 200) : null };
}

/**
 * The innermost (last) quoted "From:" header block in a body — for a forwarded
 * receipt that is the original merchant, even behind several reply layers.
 * A From: line only counts when a Sent/Date/To/Subject/Cc line follows it.
 *
 * @param {string} text
 * @returns {{ name: string|null, address: string|null }|null}
 */
export function extractQuotedFrom(text) {
  const lines = capParseText(text).replace(/[\u200B-\u200F\u2060\uFEFF]/g, '').split(/\r\n|\r|\n/)
    .map(l => (l.length > MAX_LINE_CHARS ? l.slice(0, MAX_LINE_CHARS) : l));
  let found = null;
  for (let i = 0; i < lines.length; i++) {
    const m = lines[i].match(/^[\s>*]*From:\**(.*)$/i);
    if (!m) continue;
    let raw = m[1].trim();
    let next = i + 1;
    if (!raw) { // value on the following line
      while (next < lines.length && !lines[next].trim()) next++;
      raw = (lines[next] || '').trim();
      next++;
    }
    if (!raw) continue;
    const following = lines.slice(next, next + 8).filter(l => l.trim()).slice(0, 5);
    if (!following.some(l => /^[\s>*]*(?:Sent|Date|To|Subject|Cc):/i.test(l))) continue;
    const parsed = parseMailbox(raw);
    if (parsed) found = parsed;
  }
  return found;
}

// ---------------------------------------------------------------------------
// Receipt / invoice / order numbers
// ---------------------------------------------------------------------------

// Codes must contain a digit, so a label followed by another label ("Payment method") never matches.
// The separator is written (?:\s*:)?(?:\s*#)?\s* rather than \s*:?\s*#?\s* — the same language, but
// no two whitespace runs are adjacent, so a long run of spaces cannot cause catastrophic backtracking.
const CODE = '((?=[\\w-]*\\d)[A-Za-z0-9][\\w-]*)';
const SEPARATOR = '(?:\\s*:)?(?:\\s*#)?\\s*';
const RECEIPT_NUMBER_RE = new RegExp(`receipt\\s*(?:number|no\\.?|#)${SEPARATOR}${CODE}`, 'i');
const INVOICE_NUMBER_RE = new RegExp(`invoice\\s*(?:number|no\\.?|#)${SEPARATOR}${CODE}`, 'i');
const ORDER_NUMBER_RE = new RegExp(`\\border\\s*(?:number|no\\.?|#|id)${SEPARATOR}((?=[\\w.-]*\\d)[A-Za-z0-9](?:[\\w.-]*[A-Za-z0-9])?)`, 'i');

function matchCapped(text, re) {
  const match = capParseText(text).match(re);
  return match ? capStr(match[1]) : null;
}

export function extractReceiptNumber(text) {
  return matchCapped(text, RECEIPT_NUMBER_RE);
}

export function extractInvoiceNumber(text) {
  return matchCapped(text, INVOICE_NUMBER_RE);
}

export function extractOrderNumber(text) {
  return matchCapped(text, ORDER_NUMBER_RE);
}

/** "Invoice-TSTINV42-0019.pdf" → "TSTINV42-0019" (first non-inline attachment named that way). */
export function extractInvoiceNumberFromAttachments(attachments) {
  for (const a of attachments || []) {
    if (a.isInline) continue;
    const match = (a.name || '').match(/^invoice[-_ ]+([A-Za-z0-9]+(?:-[A-Za-z0-9]+)+)\.pdf$/i);
    if (match) return capStr(match[1]);
  }
  return null;
}

// ---------------------------------------------------------------------------
// Legacy single-field extractors (kept for callers/tests; built on the totals parser)
// ---------------------------------------------------------------------------

export function extractAmount(text) {
  const { amounts } = parseReceiptTotals(text);
  return { amount: amounts.total, currency: amounts.currency };
}

export function extractTaxAmount(text) {
  return parseReceiptTotals(text).amounts.tax;
}

export function extractProductLabel(text, productLabels = []) {
  const capped = capParseText(text);
  for (const pattern of productLabels) {
    const match = capped.match(pattern);
    if (match) return capStr(match[0].trim());
  }
  return null;
}

// ---------------------------------------------------------------------------
// Billing links
// ---------------------------------------------------------------------------

const IMAGE_EXT_RE = /\.(?:png|jpe?g|gif|svg|webp|bmp|ico)$/i;
// Text just before a URL that marks it as an image/asset reference rather than a link.
const ASSET_CONTEXT_RE = /(?:url\(\s*["']?|\b(?:src|srcset|background|poster)\s*=\s*["']?)$/i;

function decodeHtmlEntities(s) {
  // &amp; and its numeric synonym &#38; keep multi-param signed URLs (e.g. pre-signed
  // Stripe links) valid; &quot;/&#39; stop a CSS url('…') value running into its quote.
  return s
    .replace(/&quot;|&#0*34;/g, '"')
    .replace(/&apos;|&#0*39;/g, "'")
    .replace(/&amp;|&#0*38;/g, '&');
}

// Billing/receipt links longer than this are dropped: real pre-signed links are a few
// hundred characters, and an unbounded link would bypass the result size budget.
const MAX_URL_LENGTH = 2048;

/**
 * Unwrap a processor click-tracking link to the URL it redirects to:
 * https://<n>.email.stripe.com/CL0/<percent-encoded target>/1/<id>/<sig>=452 → target.
 * The target is returned normalised (new URL(target).href), so it never carries raw
 * spaces, quotes or line breaks. Anything else (or a target that does not parse as
 * https) is returned unchanged. Never fetches.
 */
export function unwrapTrackedUrl(url) {
  try {
    const parsed = new URL(url);
    if (parsed.protocol !== 'https:' || !/^[a-z0-9-]+\.email\.stripe\.com$/i.test(parsed.hostname)) return url;
    const m = parsed.pathname.match(/^\/CL\d+\/(.+?)\/\d+\/[^/]+\/[^/]*$/);
    if (!m) return url;
    const decoded = decodeURIComponent(m[1]);
    if (!/^https:\/\//i.test(decoded)) return url;
    const target = new URL(decoded);
    return target.protocol === 'https:' ? target.href : url;
  } catch {
    return url;
  }
}

function hasPdfPath(pathname) {
  return pathname.split('/').some(seg => /^pdf$/i.test(seg) || /\.pdf$/i.test(seg));
}

/**
 * Allowlisted PDF links in an email body, best first. Click-tracking wrappers are
 * unwrapped; image assets (css url(), img src, image extensions) are dropped; a
 * link must have a pdf path segment or .pdf suffix. There is deliberately no
 * "first allowlisted URL" fallback: logos and other assets share those hosts.
 * Every returned url is a normalised https href of at most MAX_URL_LENGTH characters.
 * kind: 0 = invoice PDF, 1 = other PDF, 2 = receipt PDF.
 */
function collectPdfLinks(html, allowlist) {
  const source = decodeHtmlEntities(html || '');
  const links = [];
  for (const m of source.matchAll(/https:\/\/[^\s"'<>)]+/g)) {
    if (ASSET_CONTEXT_RE.test(source.slice(Math.max(0, m.index - 80), m.index))) continue;
    const target = unwrapTrackedUrl(m[0].replace(/[.,;:!?]+$/, ''));
    let parsed;
    try {
      parsed = new URL(target);
    } catch {
      continue;
    }
    if (parsed.protocol !== 'https:' || !allowlist.includes(parsed.hostname.toLowerCase())) continue;
    if (IMAGE_EXT_RE.test(parsed.pathname) || !hasPdfPath(parsed.pathname)) continue;
    const url = parsed.href;
    if (url.length > MAX_URL_LENGTH) continue;
    const path = parsed.pathname.toLowerCase();
    const kind = /\/receipts?\b/.test(path) ? 2 : /\/invoices?\b/.test(path) ? 0 : 1;
    if (!links.some(l => l.url === url)) links.push({ url, kind });
  }
  return links;
}

/** { billingPdfUrl, receiptPdfUrl } — invoice PDF preferred for billing; receipt PDF reported separately. */
export function extractBillingLinks(html, allowlist = []) {
  const links = collectPdfLinks(html, allowlist);
  const best = [0, 1, 2].map(k => links.find(l => l.kind === k)).find(Boolean);
  const receipt = links.find(l => l.kind === 2);
  return { billingPdfUrl: best ? best.url : null, receiptPdfUrl: receipt ? receipt.url : null };
}

export function extractBillingPdfUrl(html, allowlist) {
  return extractBillingLinks(html, allowlist).billingPdfUrl;
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

// ---------------------------------------------------------------------------
// Message → compact summary
// ---------------------------------------------------------------------------

const STRING_CAP = 300;
const ELLIPSIS = '…'; // …

/** Cap a parsed string at STRING_CAP characters (plus an ellipsis); non-strings pass through. */
export function capStr(s) {
  if (typeof s !== 'string' || s.length <= STRING_CAP) return s;
  return s.slice(0, STRING_CAP) + ELLIPSIS;
}

// Size budget: everything except billingPdfUrl/receiptPdfUrl must fit in 4096 bytes.
// Those two are exempt — they are fetched verbatim by the PDF link fallback, and
// truncating one corrupts a pre-signed URL — but each is at most MAX_URL_LENGTH
// characters (longer links are dropped in collectPdfLinks, never truncated).
const SIZE_BUDGET = 4096;

export function parseReceiptMessage(message, attachments = [], allowlist = [], rules = { vendorSenders: [], productLabels: [] }) {
  const html = message.body?.content || '';
  const text = capParseText(message.body?.contentType?.toLowerCase() === 'html' ? stripHtml(html) : html);
  const fromAddress = message.from?.emailAddress?.address || '';
  const subject = message.subject || '';
  const { amounts, lineItems } = parseReceiptTotals(text);
  const nonInline = attachments.filter(a => !a.isInline);
  const { billingPdfUrl, receiptPdfUrl } = extractBillingLinks(html, allowlist);

  const result = {
    vendor: capStr(detectVendor(fromAddress, subject, rules.vendorSenders || [])),
    subject: capStr(message.subject || null),
    normalisedSubject: subject ? capStr(normaliseSubject(subject)) : null,
    replyOrForward: hasReplyPrefix(subject),
    quotedFrom: extractQuotedFrom(text),
    receivedDate: message.receivedDateTime || null,
    receiptNumber: capStr(extractReceiptNumber(text)),
    invoiceNumber: capStr(extractInvoiceNumber(text) || extractInvoiceNumberFromAttachments(attachments)),
    orderNumber: capStr(extractOrderNumber(text)),
    amount: amounts.total,
    currency: amounts.currency,
    taxAmount: amounts.tax,
    amounts,
    lineItems,
    productLabel: capStr(extractProductLabel(text, rules.productLabels || []) || buildProductLabel(lineItems)),
    billingPdfUrl, // never truncate — used verbatim as fetch URL
    receiptPdfUrl, // never truncate
    hasPdfAttachment: selectPdfAttachment(attachments, 'first') !== null,
    attachmentIds: nonInline.map(a => a.id).slice(0, 20)
  };

  // Overall size guard: trim lineItems, then attachmentIds (re-fetchable via
  // Graph), until the result excluding the two URLs fits within SIZE_BUDGET bytes.
  // truncated is scoped to these guards only (capStr truncations are not tracked).
  const sizeOf = () => JSON.stringify({ ...result, billingPdfUrl: undefined, receiptPdfUrl: undefined }).length;
  if (sizeOf() > SIZE_BUDGET) {
    result.truncated = true; // set first so the flag itself counts against the budget
    while (result.lineItems.length > 0 && sizeOf() > SIZE_BUDGET) {
      result.lineItems.pop();
    }
    while (result.attachmentIds.length > 0 && sizeOf() > SIZE_BUDGET) {
      result.attachmentIds.pop();
    }
    if (sizeOf() > SIZE_BUDGET) {
      result.attachmentIds = []; // pathological: drop all
    }
  }

  return result;
}
