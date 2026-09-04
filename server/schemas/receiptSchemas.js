/**
 * Receipt/invoice-run MCP tool schemas (spec: Autonomous Invoice Submission).
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
      onExisting: { type: 'string', enum: ['skip', 'overwrite', 'version'], description: 'Collision policy when the target file already exists', default: 'skip' },
      ...mailboxProperty
    },
    required: ['messageId'],
  },
};

export const fetchBillingPdfSchema = {
  name: 'outlook_fetch_billing_pdf',
  description: 'Fetch a billing PDF from an allowlisted host (default: Stripe) linked in an email body, and save it. Fallback for receipts without PDF attachments. HTTPS-only, allowlist-only, validated as PDF. Provide url or messageId (at least one). If both are given, url takes precedence.',
  inputSchema: {
    type: 'object',
    properties: {
      messageId: { type: 'string', description: 'Email whose body contains the billing link (server extracts it)' },
      url: { type: 'string', description: 'Explicit billing-PDF URL (must be on the allowlist)' },
      destDir: { type: 'string', description: 'Destination directory (inside receipts/work dir). Defaults to MCP_OUTLOOK_RECEIPTS_DIR.' },
      fileName: { type: 'string', description: 'Target filename, e.g. "Globex 23Jun26 Invoice.pdf"' },
      onExisting: { type: 'string', enum: ['skip', 'overwrite', 'version'], description: 'Collision policy when the target file already exists', default: 'skip' }
    },
    required: ['fileName'],
    anyOf: [{ required: ['messageId'] }, { required: ['url'] }],
  },
};

export const extractReceiptSchema = {
  name: 'outlook_extract_receipt',
  description: 'Extract a compact structured summary (vendor, amount, currency, receipt/invoice numbers, product label, billing link, attachment ids) from a receipt email — instead of its large HTML body.',
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
      onExisting: { type: 'string', enum: ['skip', 'overwrite', 'version'], description: 'Collision policy when the target file already exists', default: 'skip' },
      ...mailboxProperty
    },
    required: ['messageId', 'fileName'],
  },
};

export const collectReceiptsSchema = {
  name: 'outlook_collect_receipts',
  description: 'Collect all vendor receipts for a period in one call: discovers each vendor\'s receipt emails by sender/subject/date across the whole mailbox, saves each PDF (attachment, allowlisted link, or rendered fallback) with deterministic names, and returns a manifest plus a missing[] list. Idempotent on re-runs.',
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
            subjectContains: { type: 'string', description: 'Substring the subject must contain' },
            prefer: { type: 'string', enum: ['invoice', 'receipt', 'first'], description: 'Which PDF to auto-select when several are attached', default: 'invoice' },
            filenameTemplate: { type: 'string', description: 'Overrides RECEIPT_FILENAME_TEMPLATE, default "{vendor} {DDMmmYY} Invoice.pdf"' },
            expectedCurrency: { type: 'string', description: 'ISO currency to sanity-check, e.g. GBP' }
          },
          required: ['vendor'],
        }
      },
      destDir: { type: 'string', description: 'Destination directory (inside receipts/work dir). Defaults to MCP_OUTLOOK_RECEIPTS_DIR.' },
      onExisting: { type: 'string', enum: ['skip', 'overwrite', 'version'], description: 'Collision policy when the target file already exists', default: 'skip' },
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
