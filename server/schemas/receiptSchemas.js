/**
 * Receipt MCP tool schemas: save attachments, fetch billing PDFs, extract
 * receipt summaries, render emails to PDF and collect a period's receipts.
 */

import { mailboxProperty } from './sharedSchemaFragments.js';

export const saveAttachmentSchema = {
  name: 'outlook_save_attachment',
  description: 'Save an email attachment\'s original bytes to a file (no parsing/transformation). Defaults to the Invoice-*.pdf when multiple PDFs are attached. Writes only under the configured receipts/work directory.',
  inputSchema: {
    type: 'object',
    properties: {
      messageId: { type: 'string', description: 'The ID of the email containing the attachment' },
      attachmentId: { type: 'string', description: 'Attachment ID to save. Omit to auto-select a PDF attachment.' },
      attachmentName: { type: 'string', description: 'Exact attachment filename to select (alternative to attachmentId)' },
      contentType: { type: 'string', description: 'MIME type selector (e.g. "application/pdf") — picks the first non-inline attachment with this contentType. Used when attachmentId and attachmentName are absent.' },
      prefer: { type: 'string', enum: ['invoice', 'receipt', 'first'], description: 'Which PDF to auto-select when several are attached', default: 'invoice' },
      destDir: { type: 'string', description: 'Destination directory (must be inside the receipts or work directory). Defaults to MCP_OUTLOOK_RECEIPTS_DIR.' },
      fileName: { type: 'string', description: 'Explicit target filename, e.g. "Acme 29Jun26 Invoice.pdf". Takes priority over filenameTemplate. Defaults to the attachment\'s own name when neither fileName nor filenameTemplate/vendor is supplied.' },
      filenameTemplate: { type: 'string', description: 'Template for the saved filename, e.g. "{vendor} {DDMmmYY} Invoice.pdf". Overrides RECEIPT_FILENAME_TEMPLATE. Requires vendor to be supplied or auto-detectable from the message sender.' },
      vendor: { type: 'string', description: 'Vendor label used in filenameTemplate, e.g. "Acme". When omitted, the vendor is auto-detected from the message sender.' },
      onExisting: { type: 'string', enum: ['skip', 'overwrite', 'version'], description: 'Collision policy when the target file already exists: skip keeps it only if it holds identical bytes (otherwise a new version is written and existingDiffers is set)', default: 'skip' },
      ...mailboxProperty
    },
    required: ['messageId'],
  },
};

export const fetchBillingPdfSchema = {
  name: 'outlook_fetch_billing_pdf',
  description: 'Fetch a per-receipt invoice/receipt PDF download link from an email body and save it. Only allowlisted hosts (default: Stripe) over HTTPS; the response must be a real PDF (%PDF verified) — image, logo and other asset links are never fetched. Fallback for receipts without PDF attachments. Provide url or messageId (at least one). If both are given, url takes precedence.',
  inputSchema: {
    type: 'object',
    properties: {
      messageId: { type: 'string', description: 'Email whose body contains the billing link (server extracts it)' },
      url: { type: 'string', description: 'Explicit billing-PDF URL (must be on the allowlist)' },
      destDir: { type: 'string', description: 'Destination directory (inside receipts/work dir). Defaults to MCP_OUTLOOK_RECEIPTS_DIR.' },
      fileName: { type: 'string', description: 'Target filename, e.g. "Globex 23Jun26 Invoice.pdf"' },
      onExisting: { type: 'string', enum: ['skip', 'overwrite', 'version'], description: 'Collision policy when the target file already exists: skip keeps it only if it holds identical bytes (otherwise a new version is written and existingDiffers is set)', default: 'skip' },
      ...mailboxProperty
    },
    required: ['fileName'],
    anyOf: [{ required: ['messageId'] }, { required: ['url'] }],
  },
};

export const extractReceiptSchema = {
  name: 'outlook_extract_receipt',
  description: 'Extract a compact structured summary from a receipt email instead of its large HTML body: vendor, amounts (subtotal, tax, credits, total, amount paid, currency), line items, receipt/invoice/order numbers, product label, reply/forward flag with normalised subject and quoted original sender, invoice and receipt PDF links (tracking unwrapped, allowlisted hosts only), and attachment ids.',
  inputSchema: {
    type: 'object',
    properties: {
      messageId: { type: 'string', description: 'The ID of the receipt email' },
      ...mailboxProperty
    },
    required: ['messageId'],
  },
};

export const renderEmailPdfSchema = {
  name: 'outlook_render_email_pdf',
  description: 'Render an email\'s sanitised HTML body to a PDF file (headless Chrome). Fallback for receipts with no PDF attachment and no billing link (e.g. app-store order receipts) — produces an audit-trail document, not the formal vendor invoice.',
  inputSchema: {
    type: 'object',
    properties: {
      messageId: { type: 'string', description: 'The ID of the email to render' },
      destDir: { type: 'string', description: 'Destination directory (inside receipts/work dir). Defaults to MCP_OUTLOOK_RECEIPTS_DIR.' },
      fileName: { type: 'string', description: 'Target filename, e.g. "Hooli 09Jun26 Invoice.pdf"' },
      onExisting: { type: 'string', enum: ['skip', 'overwrite', 'version'], description: 'Collision policy when the target file already exists: skip keeps it only if it holds identical bytes (otherwise a new version is written and existingDiffers is set)', default: 'skip' },
      ...mailboxProperty
    },
    required: ['messageId', 'fileName'],
  },
};

export const collectReceiptsSchema = {
  name: 'outlook_collect_receipts',
  description: 'Collect all vendor receipts for a period in one call. Discovers each vendor\'s receipt emails by sender/subject/date across the whole mailbox (Sent Items and Drafts excluded; a message found by several rules is processed once), assigns each message to exactly one rule, collapses copies of one receipt (same receipt/invoice/order number, or identical PDF bytes) and saves one PDF per receipt with deterministic names: the PDF attachment, else a per-receipt allowlisted link, else a rendering of the email. Every manifest entry has a status: saved, skipped (an identical file is already on disk; a same-named file with different bytes gets a new version flagged existingDiffers), not_found, fetch_failed, ambiguous (several rules claim it, or its product cannot be determined; candidates listed) or duplicate (duplicate_of names the receipt number or sha256). Entries carry sender, a forwarded flag and originalSender for forwards, and structured amounts (subtotal, tax, credits, total, currency) read from the saved PDF when it parses, else from the email (amountSource; amountMismatch when they disagree). When the formal invoice cannot be fetched the entry stays fetch_failed with a renderedFallback copy that is marked as not the formal invoice. Returns { period, manifest, missing (not_found vendors), failed (one object per fetch_failed entry with the attempted route, ids and Graph error), summary (count per status), unmatched (messages every product matcher rejected), notes }. Idempotent on re-runs.',
  inputSchema: {
    type: 'object',
    properties: {
      periodStart: { type: 'string', description: 'ISO start of the period, e.g. 2026-06-01T00:00:00Z' },
      periodEnd: { type: 'string', description: 'ISO end of the period, e.g. 2026-06-30T23:59:59Z' },
      vendors: {
        type: 'array',
        description: 'Vendor rules to collect',
        items: {
          type: 'object',
          properties: {
            vendor: { type: 'string', description: 'Vendor label used in filenames and the manifest, e.g. "Acme"' },
            from: { type: 'string', description: 'Sender email to match, e.g. billing@acme.example' },
            subjectContains: { type: 'string', description: 'Substring the subject must contain (Re:/Fw: prefixes are ignored)' },
            productContains: {
              oneOf: [{ type: 'string' }, { type: 'array', items: { type: 'string' } }],
              description: 'Only receipts for these products: at least one term must appear (case-insensitive substring) in the receipt\'s product label or a line-item description. Use with productExcludes to split one sender into several vendors. A rule with productContains wins over a rule without product matchers; a receipt whose product cannot be determined is reported as ambiguous, never guessed.'
            },
            productExcludes: {
              oneOf: [{ type: 'string' }, { type: 'array', items: { type: 'string' } }],
              description: 'Skip receipts for these products: none of the terms may appear (case-insensitive substring) in the receipt\'s product label or any line-item description.'
            },
            prefer: { type: 'string', enum: ['invoice', 'receipt', 'first'], description: 'Which PDF to auto-select when several are attached', default: 'invoice' },
            filenameTemplate: { type: 'string', description: 'Overrides RECEIPT_FILENAME_TEMPLATE, default "{vendor} {DDMmmYY} Invoice.pdf"' },
            expectedCurrency: { type: 'string', description: 'ISO currency to sanity-check, e.g. GBP' }
          },
          required: ['vendor'],
        }
      },
      destDir: { type: 'string', description: 'Destination directory (inside receipts/work dir). Defaults to MCP_OUTLOOK_RECEIPTS_DIR.' },
      onExisting: { type: 'string', enum: ['skip', 'overwrite', 'version'], description: 'Collision policy when the target file already exists: skip keeps it only if it holds identical bytes (otherwise a new version is written and existingDiffers is set)', default: 'skip' },
      ...mailboxProperty
    },
    required: ['periodStart', 'periodEnd', 'vendors'],
  },
};

export const receiptSchemas = [
  saveAttachmentSchema,
  fetchBillingPdfSchema,
  extractReceiptSchema,
  renderEmailPdfSchema,
  collectReceiptsSchema,
];

export const receiptSchemaMap = {
  'outlook_save_attachment': saveAttachmentSchema,
  'outlook_fetch_billing_pdf': fetchBillingPdfSchema,
  'outlook_extract_receipt': extractReceiptSchema,
  'outlook_render_email_pdf': renderEmailPdfSchema,
  'outlook_collect_receipts': collectReceiptsSchema,
};
