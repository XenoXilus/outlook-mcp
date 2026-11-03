import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { fileURLToPath } from 'url';
import { collectReceiptsTool } from '../../tools/receipts/collectReceipts.js';

const PDF_B64 = Buffer.from('%PDF-1.4 invoice').toString('base64');

function stripeMessage(id, receivedDateTime) {
  return {
    id, receivedDateTime, hasAttachments: true,
    subject: 'Your receipt from Acme, Inc #4242',
    from: { emailAddress: { address: 'invoice+statements@mail.acme.com' } }
  };
}

const STRIPE_FULL = (id, receivedDateTime) => ({
  ...stripeMessage(id, receivedDateTime),
  body: { contentType: 'html', content: '<p>Amount paid £42.00</p><td>Pro plan - 5x</td>' }
});

const ATTACHMENT_LIST = { value: [
  { id: 'a1', name: 'Invoice-0015.pdf', contentType: 'application/pdf', isInline: false },
  { id: 'a2', name: 'Receipt-2060.pdf', contentType: 'application/pdf', isInline: false }
]};

const ATTACHMENT_FULL = { id: 'a1', name: 'Invoice-0015.pdf', contentType: 'application/pdf', isInline: false, contentBytes: PDF_B64 };

// Well-known folder lookups, made once per run after discovery finds messages.
const SENT_FOLDER = { id: 'folder-sent' };
const DRAFTS_FOLDER = { id: 'folder-drafts' };

// Fictional PDFs rendered from the .html files beside them.
const here = path.dirname(fileURLToPath(import.meta.url));
const fixturePdf = (name) => fs.readFileSync(path.join(here, '../fixtures/receipts', name));

/**
 * One message as both the search header and the full message, plus its
 * attachment list and per-attachment content (base64 string or an MCP error object).
 */
function receiptMessage({ id, date, from, subject, html, attachments = [], content = {}, parentFolderId = 'folder-inbox' }) {
  const header = {
    id, receivedDateTime: date, hasAttachments: attachments.length > 0, subject,
    from: { emailAddress: { address: from } }, parentFolderId
  };
  return { header, full: { ...header, body: { contentType: 'html', content: html } }, attachments, content };
}

const pdfAttachment = (id, name) => ({ id, name, contentType: 'application/pdf', isInline: false });

/** A minimal one-page text PDF (one line per entry) that pdfjs can read. */
function textPdf(lines) {
  const esc = (s) => s.replace(/[\\()]/g, (c) => `\\${c}`);
  const stream = ['BT', '/F1 12 Tf', '14 TL', '72 720 Td', ...lines.map(l => `(${esc(l)}) Tj T*`), 'ET'].join('\n');
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
    `<< /Length ${Buffer.byteLength(stream, 'latin1')} >>\nstream\n${stream}\nendstream`
  ];
  let out = '%PDF-1.4\n';
  const offsets = [];
  objects.forEach((body, i) => {
    offsets.push(Buffer.byteLength(out, 'latin1'));
    out += `${i + 1} 0 obj\n${body}\nendobj\n`;
  });
  const xref = Buffer.byteLength(out, 'latin1');
  out += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n` +
    offsets.map(o => `${String(o).padStart(10, '0')} 00000 n \n`).join('');
  out += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(out, 'latin1');
}

/** Rough Graph $filter emulation: from eq / startswith(from) / contains(subject), OR-ed. */
function graphSearch(messages) {
  return (filter) => messages.map(m => m.header).filter(h => {
    const from = h.from.emailAddress.address.toLowerCase();
    const subject = h.subject.toLowerCase();
    const values = (re) => [...filter.matchAll(re)].map(m => m[1].toLowerCase());
    return values(/from\/emailAddress\/address eq '([^']*)'/g).includes(from) ||
      values(/startswith\(from\/emailAddress\/address,'([^']*)'\)/g).some(p => from.startsWith(p)) ||
      values(/contains\(subject,'([^']*)'\)/g).some(s => subject.includes(s));
  });
}

/**
 * makeRequest answered by path, so call order does not matter. `search(filter)`
 * returns headers (or an MCP error object); folder lookups use `folders`.
 */
function mailboxRouter({ messages = [], search = graphSearch(messages), folders = {}, base = '/me' }) {
  const byId = Object.fromEntries(messages.map(m => [m.header.id, m]));
  const folderReplies = { sentitems: SENT_FOLDER, drafts: DRAFTS_FOLDER, ...folders };
  return async (requestPath, options = {}) => {
    if (requestPath === `${base}/messages`) {
      const result = search(options.filter || '');
      return Array.isArray(result) ? { value: result } : result;
    }
    const folder = requestPath.match(/^(.*)\/mailFolders\/([^/]+)$/);
    if (folder && folder[1] === base && folder[2] in folderReplies) return folderReplies[folder[2]];
    const parts = requestPath.slice(base.length).match(/^\/messages\/([^/]+)(\/attachments(?:\/([^/]+))?)?$/);
    const message = requestPath.startsWith(base) && parts && byId[parts[1]];
    if (message) {
      if (!parts[2]) return message.full;
      if (!parts[3]) return { value: message.attachments };
      const content = message.content[parts[3]];
      if (typeof content === 'string') {
        return { ...message.attachments.find(a => a.id === parts[3]), contentBytes: content };
      }
      if (content) return content;
    }
    throw new Error(`Unexpected Graph call: ${requestPath}`);
  };
}

function graphFailure(status, code, message) {
  return {
    isError: true,
    content: [{ type: 'text', text: `Graph API error ${status} ${code}: ${message}` }],
    _errorDetails: { graphError: { status, code, message, requestId: `req-${status}`, method: 'GET' } }
  };
}

describe('collectReceiptsTool (FR-6)', () => {
  let receiptsDir, makeRequest, getBinary, authManager, savedEnv;

  beforeEach(() => {
    receiptsDir = fs.mkdtempSync(path.join(os.tmpdir(), 'receipts-'));
    savedEnv = process.env.MCP_OUTLOOK_RECEIPTS_DIR;
    process.env.MCP_OUTLOOK_RECEIPTS_DIR = receiptsDir;
    makeRequest = vi.fn();
    getBinary = vi.fn().mockRejectedValue(new Error('raw value route unavailable'));
    authManager = {
      ensureAuthenticated: vi.fn().mockResolvedValue(true),
      getGraphApiClient: vi.fn().mockReturnValue({ makeRequest, getBinary })
    };
  });

  afterEach(() => {
    if (savedEnv === undefined) delete process.env.MCP_OUTLOOK_RECEIPTS_DIR;
    else process.env.MCP_OUTLOOK_RECEIPTS_DIR = savedEnv;
    fs.rmSync(receiptsDir, { recursive: true, force: true });
  });

  it('validates required inputs', async () => {
    expect((await collectReceiptsTool(authManager, {})).isError).toBe(true);
    expect((await collectReceiptsTool(authManager, {
      periodStart: '2026-06-01T00:00:00Z', periodEnd: '2026-06-30T23:59:59Z', vendors: []
    })).isError).toBe(true);
  });

  it('saves an attachment receipt and flags an absent vendor as not_found', async () => {
    // Searches for every vendor run first, then folder lookups, extraction and saving.
    makeRequest.mockImplementation(mailboxRouter({
      messages: [{
        header: stripeMessage('m1', '2026-06-29T07:12:00Z'),
        full: STRIPE_FULL('m1', '2026-06-29T07:12:00Z'),
        attachments: ATTACHMENT_LIST.value,
        content: { a1: PDF_B64 }
      }]
    }));

    const res = await collectReceiptsTool(authManager, {
      periodStart: '2026-06-01T00:00:00Z',
      periodEnd: '2026-06-30T23:59:59Z',
      vendors: [
        { vendor: 'Acme', from: 'invoice+statements@mail.acme.com' },
        { vendor: 'Vandelay', from: 'billing@vandelay.com' }
      ]
    });
    const out = JSON.parse(res.content[0].text);

    const saved = out.manifest.find(e => e.vendor === 'Acme');
    expect(saved.status).toBe('saved');
    expect(saved.method).toBe('attachment');
    expect(saved.amount).toBe(42.0);
    expect(saved.currency).toBe('GBP');
    expect(saved.sourceMessageId).toBe('m1');
    expect(path.basename(saved.savedPath)).toBe('Acme 29Jun26 Invoice.pdf');
    expect(fs.existsSync(saved.savedPath)).toBe(true);

    expect(out.missing).toEqual(['Vandelay']);
    expect(out.manifest.find(e => e.vendor === 'Vandelay').status).toBe('not_found');

    // search used date-bounded, sender-filtered whole-mailbox query
    const [endpoint, options] = makeRequest.mock.calls[0];
    expect(endpoint).toBe('/me/messages');
    expect(options.filter).toContain("receivedDateTime ge 2026-06-01T00:00:00Z");
    expect(options.filter).toContain("from/emailAddress/address eq 'invoice+statements@mail.acme.com'");
    expect(options.select.split(',')).toContain('parentFolderId');
  });

  // Item 1: link fallback (FR-2) — no attachment, allowlisted Stripe URL in body
  it('saves via link fallback and records method:link', async () => {
    const linkMsgHeader = {
      id: 'm-link', receivedDateTime: '2026-06-29T07:12:00Z', hasAttachments: false,
      subject: 'Initech receipt', from: { emailAddress: { address: 'billing@initech.app' } }
    };
    const linkMsgFull = {
      ...linkMsgHeader,
      body: { contentType: 'html', content: '<p>Amount paid $50.00</p><a href="https://pay.stripe.com/r/x.pdf">Invoice</a>' }
    };

    makeRequest
      .mockResolvedValueOnce({ value: [linkMsgHeader] }) // search
      .mockResolvedValueOnce(SENT_FOLDER)                // excluded-folder lookups
      .mockResolvedValueOnce(DRAFTS_FOLDER)
      .mockResolvedValueOnce(linkMsgFull);               // extractReceiptCore: message detail

    const PDF_LINK = Buffer.from('%PDF-1.4 link-invoice');
    const fetchImpl = vi.fn().mockResolvedValue({
      status: 200, ok: true,
      headers: { get: (k) => k.toLowerCase() === 'content-type' ? 'application/pdf' : null },
      arrayBuffer: async () => PDF_LINK.buffer.slice(PDF_LINK.byteOffset, PDF_LINK.byteOffset + PDF_LINK.byteLength)
    });

    const res = await collectReceiptsTool(authManager, {
      periodStart: '2026-06-01T00:00:00Z',
      periodEnd: '2026-06-30T23:59:59Z',
      vendors: [{ vendor: 'Initech', from: 'billing@initech.app' }]
    }, { fetchImpl });

    const out = JSON.parse(res.content[0].text);
    const entry = out.manifest.find(e => e.vendor === 'Initech');
    expect(entry.status).toBe('saved');
    expect(entry.method).toBe('link');
    expect(entry.sourceMessageId).toBe('m-link');
    expect(fs.existsSync(entry.savedPath)).toBe(true);
    expect(fetchImpl).toHaveBeenCalledOnce();
    // FR-6 manifest contract fields
    expect(entry.sha256).toMatch(/^[0-9a-f]{64}$/);
    expect(entry.amount).toBe(50.0);
    expect(entry.currency).toBe('USD');
    expect(entry.date).toBe('2026-06-29T07:12:00Z');
  });

  // Item 2: rendered fallback (FR-7) — no attachment, no billing link
  it('saves via rendered fallback and records method:rendered', async () => {
    const renderMsgHeader = {
      id: 'm-render', receivedDateTime: '2026-06-15T09:00:00Z', hasAttachments: false,
      subject: 'Your Hooli Store order receipt',
      from: { emailAddress: { address: 'store-noreply@hooli.example' } }
    };
    const renderMsgFull = {
      ...renderMsgHeader,
      body: { contentType: 'html', content: '<p>Amount paid $14.99</p>' }
    };

    makeRequest
      .mockResolvedValueOnce({ value: [renderMsgHeader] }) // search
      .mockResolvedValueOnce(SENT_FOLDER)                  // excluded-folder lookups
      .mockResolvedValueOnce(DRAFTS_FOLDER)
      .mockResolvedValueOnce(renderMsgFull)                // extractReceiptCore: message detail
      .mockResolvedValueOnce(renderMsgFull);               // renderEmailPdfCore: message detail

    const PDF_RENDER = Buffer.from('%PDF-1.4 rendered');
    const renderImpl = vi.fn().mockResolvedValue(PDF_RENDER);

    const res = await collectReceiptsTool(authManager, {
      periodStart: '2026-06-01T00:00:00Z',
      periodEnd: '2026-06-30T23:59:59Z',
      vendors: [{ vendor: 'Hooli Store', from: 'store-noreply@hooli.example' }]
    }, { renderImpl });

    const out = JSON.parse(res.content[0].text);
    const entry = out.manifest.find(e => e.vendor === 'Hooli Store');
    expect(entry.status).toBe('saved');
    expect(entry.method).toBe('rendered');
    expect(entry.sourceMessageId).toBe('m-render');
    expect(fs.existsSync(entry.savedPath)).toBe(true);
    expect(renderImpl).toHaveBeenCalledOnce();
    // FR-6 manifest contract fields
    expect(entry.sha256).toMatch(/^[0-9a-f]{64}$/);
    expect(entry.amount).toBe(14.99);
    expect(entry.currency).toBe('USD');
    expect(entry.date).toBe('2026-06-15T09:00:00Z');
  });

  // Items 3 + 4: multiple receipts per vendor; sha256 present on each entry
  it('creates two manifest entries for a vendor with two messages (sha256 on each)', async () => {
    const m1Header = {
      id: 'm1', receivedDateTime: '2026-06-29T07:12:00Z', hasAttachments: true,
      subject: 'Pro plan receipt',
      from: { emailAddress: { address: 'invoice+statements@mail.acme.com' } }
    };
    const m1Full = {
      ...m1Header,
      body: { contentType: 'html', content: '<p>Amount paid £42.00</p><td>Pro plan</td>' }
    };
    const m2Header = {
      id: 'm2', receivedDateTime: '2026-06-22T14:00:00Z', hasAttachments: true,
      subject: 'API recharge receipt',
      from: { emailAddress: { address: 'invoice+statements@mail.acme.com' } }
    };
    const m2Full = {
      ...m2Header,
      body: { contentType: 'html', content: '<p>Amount paid $25.00</p><td>Auto-recharge</td>' }
    };
    const attachList = { value: [
      { id: 'a1', name: 'Invoice-001.pdf', contentType: 'application/pdf', isInline: false }
    ]};
    const attachFull1 = {
      id: 'a1', name: 'Invoice-001.pdf', contentType: 'application/pdf', isInline: false,
      contentBytes: Buffer.from('%PDF-1.4 m1-invoice').toString('base64')
    };
    const attachFull2 = {
      id: 'a1', name: 'Invoice-001.pdf', contentType: 'application/pdf', isInline: false,
      contentBytes: Buffer.from('%PDF-1.4 m2-invoice').toString('base64')
    };

    // Dispatch by path so the mock is order-independent (9d hardening).
    makeRequest.mockImplementation((path) => {
      if (path === '/me/messages') return Promise.resolve({ value: [m1Header, m2Header] });
      if (path === '/me/messages/m1') return Promise.resolve(m1Full);
      if (path === '/me/messages/m1/attachments') return Promise.resolve(attachList);
      if (path === '/me/messages/m1/attachments/a1') return Promise.resolve(attachFull1);
      if (path === '/me/messages/m2') return Promise.resolve(m2Full);
      if (path === '/me/messages/m2/attachments') return Promise.resolve(attachList);
      if (path === '/me/messages/m2/attachments/a1') return Promise.resolve(attachFull2);
      return Promise.reject(new Error(`Unexpected makeRequest path in multi-receipt test: ${path}`));
    });

    const res = await collectReceiptsTool(authManager, {
      periodStart: '2026-06-01T00:00:00Z',
      periodEnd: '2026-06-30T23:59:59Z',
      vendors: [{ vendor: 'Acme', from: 'invoice+statements@mail.acme.com' }]
    });

    const out = JSON.parse(res.content[0].text);
    const entries = out.manifest.filter(e => e.vendor === 'Acme');
    expect(entries).toHaveLength(2);
    // Both saved as attachments
    expect(entries[0].method).toBe('attachment');
    expect(entries[1].method).toBe('attachment');
    // Different dates → different filenames
    const names = entries.map(e => path.basename(e.savedPath));
    expect(names[0]).not.toBe(names[1]);
    // Both files exist on disk
    expect(fs.existsSync(entries[0].savedPath)).toBe(true);
    expect(fs.existsSync(entries[1].savedPath)).toBe(true);
    // sha256 present and correctly formatted (item 4)
    expect(entries[0].sha256).toMatch(/^[0-9a-f]{64}$/);
    expect(entries[1].sha256).toMatch(/^[0-9a-f]{64}$/);
    expect(entries[0].sha256).not.toBe(entries[1].sha256);
  });

  // Item 5: idempotency — second run with onExisting:skip records action:skipped, no duplicate
  it('skips file on re-run (onExisting:skip idempotency)', async () => {
    const args = {
      periodStart: '2026-06-01T00:00:00Z',
      periodEnd: '2026-06-30T23:59:59Z',
      vendors: [{ vendor: 'Acme', from: 'invoice+statements@mail.acme.com' }]
    };
    // Each run needs: search + folder lookups + extract msg + extract attachments + save list + save full
    makeRequest
      // Run 1
      .mockResolvedValueOnce({ value: [stripeMessage('m1', '2026-06-29T07:12:00Z')] })
      .mockResolvedValueOnce(SENT_FOLDER)
      .mockResolvedValueOnce(DRAFTS_FOLDER)
      .mockResolvedValueOnce(STRIPE_FULL('m1', '2026-06-29T07:12:00Z'))
      .mockResolvedValueOnce(ATTACHMENT_LIST)
      .mockResolvedValueOnce(ATTACHMENT_LIST)
      .mockResolvedValueOnce(ATTACHMENT_FULL)
      // Run 2
      .mockResolvedValueOnce({ value: [stripeMessage('m1', '2026-06-29T07:12:00Z')] })
      .mockResolvedValueOnce(SENT_FOLDER)
      .mockResolvedValueOnce(DRAFTS_FOLDER)
      .mockResolvedValueOnce(STRIPE_FULL('m1', '2026-06-29T07:12:00Z'))
      .mockResolvedValueOnce(ATTACHMENT_LIST)
      .mockResolvedValueOnce(ATTACHMENT_LIST)
      .mockResolvedValueOnce(ATTACHMENT_FULL);

    const res1 = await collectReceiptsTool(authManager, args);
    const out1 = JSON.parse(res1.content[0].text);
    expect(out1.manifest[0].action).toBe('saved');

    const res2 = await collectReceiptsTool(authManager, args);
    const out2 = JSON.parse(res2.content[0].text);
    expect(out2.manifest[0].action).toBe('skipped');
    expect(out2.manifest[0].status).toBe('skipped'); // the file was already on disk
    expect(out2.summary.skipped).toBe(1);

    // Only one PDF file on disk — no duplicate
    const files = fs.readdirSync(receiptsDir).filter(f => f.endsWith('.pdf'));
    expect(files).toHaveLength(1);
  });

  // Item 6: currency mismatch — rule has expectedCurrency:'USD', receipt is GBP
  it('records currencyMismatch when receipt currency differs from expectedCurrency', async () => {
    makeRequest
      .mockResolvedValueOnce({ value: [stripeMessage('m1', '2026-06-29T07:12:00Z')] })
      .mockResolvedValueOnce(SENT_FOLDER)
      .mockResolvedValueOnce(DRAFTS_FOLDER)
      .mockResolvedValueOnce(STRIPE_FULL('m1', '2026-06-29T07:12:00Z'))
      .mockResolvedValueOnce(ATTACHMENT_LIST)
      .mockResolvedValueOnce(ATTACHMENT_LIST)
      .mockResolvedValueOnce(ATTACHMENT_FULL);

    const res = await collectReceiptsTool(authManager, {
      periodStart: '2026-06-01T00:00:00Z',
      periodEnd: '2026-06-30T23:59:59Z',
      vendors: [{ vendor: 'Acme', from: 'invoice+statements@mail.acme.com', expectedCurrency: 'USD' }]
    });

    const out = JSON.parse(res.content[0].text);
    const entry = out.manifest[0];
    expect(entry.currency).toBe('GBP');
    expect(entry.currencyMismatch).toBe('expected USD, found GBP');
  });

  // Item 7: shared mailbox routing via MCP_OUTLOOK_SHARED_MAILBOX
  it('routes Graph requests through shared-mailbox path when MCP_OUTLOOK_SHARED_MAILBOX is set', async () => {
    const savedMailbox = process.env.MCP_OUTLOOK_SHARED_MAILBOX;
    process.env.MCP_OUTLOOK_SHARED_MAILBOX = 'finance@example.com';
    try {
      makeRequest.mockResolvedValueOnce({ value: [] }); // vendor search returns nothing
      await collectReceiptsTool(authManager, {
        periodStart: '2026-06-01T00:00:00Z',
        periodEnd: '2026-06-30T23:59:59Z',
        vendors: [{ vendor: 'Acme', from: 'invoice+statements@mail.acme.com' }]
      });
      expect(makeRequest.mock.calls[0][0]).toBe('/users/finance%40example.com/messages');
    } finally {
      if (savedMailbox === undefined) delete process.env.MCP_OUTLOOK_SHARED_MAILBOX;
      else process.env.MCP_OUTLOOK_SHARED_MAILBOX = savedMailbox;
    }
  });

  it('hard-errors on a bad RECEIPT_RULES_PATH before any Graph call', async () => {
    const savedRulesPath = process.env.RECEIPT_RULES_PATH;
    process.env.RECEIPT_RULES_PATH = '/nonexistent/receipt-rules.json';
    try {
      const res = await collectReceiptsTool(authManager, {
        periodStart: '2026-06-01T00:00:00Z',
        periodEnd: '2026-06-30T23:59:59Z',
        vendors: [{ vendor: 'Acme', from: 'billing@acme.example' }]
      });
      expect(res.isError).toBe(true);
      expect(res.content[0].text).toContain('RECEIPT_RULES_PATH');
      expect(makeRequest).not.toHaveBeenCalled();
    } finally {
      if (savedRulesPath === undefined) delete process.env.RECEIPT_RULES_PATH;
      else process.env.RECEIPT_RULES_PATH = savedRulesPath;
    }
  });

  // --- from-matching resilience ---

  it('combines from and subjectContains with OR so a sender mismatch cannot hide a subject hit', async () => {
    makeRequest
      .mockResolvedValueOnce({ value: [stripeMessage('m1', '2026-06-29T07:12:00Z')] })
      .mockResolvedValueOnce(SENT_FOLDER)
      .mockResolvedValueOnce(DRAFTS_FOLDER)
      .mockResolvedValueOnce(STRIPE_FULL('m1', '2026-06-29T07:12:00Z'))
      .mockResolvedValueOnce(ATTACHMENT_LIST)
      .mockResolvedValueOnce(ATTACHMENT_LIST)
      .mockResolvedValueOnce(ATTACHMENT_FULL);

    const res = await collectReceiptsTool(authManager, {
      periodStart: '2026-06-01T00:00:00Z',
      periodEnd: '2026-06-30T23:59:59Z',
      vendors: [{ vendor: 'Acme', from: 'invoice+statements@mail.acme.com', subjectContains: 'receipt from Acme' }]
    });

    const [, options] = makeRequest.mock.calls[0];
    expect(options.filter).toContain(
      "(from/emailAddress/address eq 'invoice+statements@mail.acme.com' or contains(subject,'receipt from Acme'))"
    );
    const out = JSON.parse(res.content[0].text);
    expect(out.manifest[0].status).toBe('saved');
    expect(out.manifest[0].matchedBy).toContain('from');
    expect(out.manifest[0].matchedBy).toContain('subject');
  });

  it('retries a zero-hit from rule with a plus-address-normalised sender fallback', async () => {
    // Config says invoice@stripe.example; the real sender drifted to a plus tag.
    const driftedHeader = {
      id: 'm-drift', receivedDateTime: '2026-06-20T10:00:00Z', hasAttachments: true,
      subject: 'Your receipt #123',
      from: { emailAddress: { address: 'invoice+statements+acct_1@stripe.example' } }
    };
    const driftedFull = {
      ...driftedHeader,
      body: { contentType: 'html', content: '<p>Amount paid $12.53</p>' }
    };

    makeRequest
      .mockResolvedValueOnce({ value: [] })              // exact-from search: nothing
      .mockResolvedValueOnce({ value: [driftedHeader] }) // startswith fallback search
      .mockResolvedValueOnce(SENT_FOLDER)                // excluded-folder lookups
      .mockResolvedValueOnce(DRAFTS_FOLDER)
      .mockResolvedValueOnce(driftedFull)                // extract: message
      .mockResolvedValueOnce(ATTACHMENT_LIST)            // extract: attachments
      .mockResolvedValueOnce(ATTACHMENT_LIST)            // save: list
      .mockResolvedValueOnce(ATTACHMENT_FULL);           // save: full

    const res = await collectReceiptsTool(authManager, {
      periodStart: '2026-06-01T00:00:00Z',
      periodEnd: '2026-06-30T23:59:59Z',
      vendors: [{ vendor: 'Stripe', from: 'invoice@stripe.example' }]
    });

    const retryFilter = makeRequest.mock.calls[1][1].filter;
    expect(retryFilter).toContain("startswith(from/emailAddress/address,'invoice')");

    const out = JSON.parse(res.content[0].text);
    const entry = out.manifest.find(e => e.vendor === 'Stripe');
    expect(entry.status).toBe('saved');
    expect(entry.matchedBy).toContain('from-normalised');
    expect(entry.searchNote).toMatch(/plus-address/i);
    expect(out.missing).toEqual([]);
  });

  it('fallback keeps only messages whose normalised sender really matches', async () => {
    const strangerHeader = {
      id: 'm-other', receivedDateTime: '2026-06-20T10:00:00Z', hasAttachments: true,
      subject: 'Invoice attached',
      from: { emailAddress: { address: 'invoice@shady.example' } }
    };

    makeRequest
      .mockResolvedValueOnce({ value: [] })                // exact-from search
      .mockResolvedValueOnce({ value: [strangerHeader] }); // fallback finds same local part, wrong domain

    const res = await collectReceiptsTool(authManager, {
      periodStart: '2026-06-01T00:00:00Z',
      periodEnd: '2026-06-30T23:59:59Z',
      vendors: [{ vendor: 'Stripe', from: 'invoice@stripe.example' }]
    });

    const out = JSON.parse(res.content[0].text);
    expect(out.missing).toEqual(['Stripe']);
    expect(out.manifest[0].status).toBe('not_found');
  });

  // --- date windows: the last day of the month must be inside the window ---

  it('accepts date-only period bounds and expands the end to cover the whole final day', async () => {
    makeRequest
      .mockResolvedValueOnce({ value: [] })
      .mockResolvedValueOnce({ value: [] }); // sender fallback also empty

    await collectReceiptsTool(authManager, {
      periodStart: '2026-08-01',
      periodEnd: '2026-08-31',
      vendors: [{ vendor: 'Acme', from: 'billing@acme.example' }]
    });

    const [, options] = makeRequest.mock.calls[0];
    expect(options.filter).toContain('receivedDateTime ge 2026-08-01T00:00:00Z');
    expect(options.filter).toContain('receivedDateTime lt 2026-09-01T00:00:00Z');
    expect(options.filter).not.toContain('le 2026-08-31');
  });

  it('marks a vendor entry fetch_failed (with error) when reading the message fails, and continues', async () => {
    makeRequest
      .mockResolvedValueOnce({ value: [stripeMessage('m1', '2026-06-29T07:12:00Z')] })
      .mockResolvedValueOnce({ value: [] }) // second vendor still searched
      .mockResolvedValueOnce({ value: [] }) // ...and its sender fallback
      .mockResolvedValueOnce(SENT_FOLDER)
      .mockResolvedValueOnce(DRAFTS_FOLDER)
      .mockResolvedValueOnce({ content: [{ type: 'text', text: 'Graph exploded' }], isError: true }); // extract fails

    const res = await collectReceiptsTool(authManager, {
      periodStart: '2026-06-01T00:00:00Z', periodEnd: '2026-06-30T23:59:59Z',
      vendors: [
        { vendor: 'Acme', from: 'invoice+statements@mail.acme.com' },
        { vendor: 'Vandelay', from: 'billing@vandelay.com' }
      ]
    });
    const out = JSON.parse(res.content[0].text);
    const failed = out.manifest.find(e => e.vendor === 'Acme');
    expect(failed.status).toBe('fetch_failed');
    expect(failed.attemptedRoute).toBe('message');
    expect(failed.error).toContain('Graph exploded');
    expect(out.failed).toHaveLength(1);
    expect(out.missing).toEqual(['Vandelay']);
  });

  it('routes the whole collection run through a per-call mailbox', async () => {
    makeRequest
      .mockResolvedValueOnce({ value: [stripeMessage('m1', '2026-06-29T07:12:00Z')] })
      .mockResolvedValueOnce(SENT_FOLDER)
      .mockResolvedValueOnce(DRAFTS_FOLDER)
      .mockResolvedValueOnce(STRIPE_FULL('m1', '2026-06-29T07:12:00Z'))
      .mockResolvedValueOnce(ATTACHMENT_LIST)
      .mockResolvedValueOnce(ATTACHMENT_LIST)
      .mockResolvedValueOnce(ATTACHMENT_FULL);

    const res = await collectReceiptsTool(authManager, {
      periodStart: '2026-06-01T00:00:00Z', periodEnd: '2026-06-30T23:59:59Z',
      mailbox: 'finance-shared@example.com',
      vendors: [{ vendor: 'Acme', from: 'invoice+statements@mail.acme.com' }]
    });
    expect(JSON.parse(res.content[0].text).manifest[0].status).toBe('saved');
    for (const call of makeRequest.mock.calls) {
      expect(call[0].startsWith('/users/finance-shared%40example.com/')).toBe(true);
    }
  });

  // --- statuses, discovery hygiene, claiming, dedupe, save routes, amounts ---

  describe('manifest statuses and claiming', () => {
    const PERIOD = { periodStart: '2026-06-01', periodEnd: '2026-06-30' };
    const renderPdf = () => vi.fn().mockImplementation(async (doc) =>
      Buffer.from(`%PDF-1.4 rendered ${doc.length} ${Math.random()}`));

    async function run(vendors, { messages = [], search, folders, deps = {}, args = {} } = {}) {
      makeRequest.mockImplementation(mailboxRouter({ messages, search, folders }));
      const res = await collectReceiptsTool(authManager, { ...PERIOD, vendors, ...args }, deps);
      expect(res.isError).toBeFalsy();
      return JSON.parse(res.content[0].text);
    }

    const graphPaths = () => makeRequest.mock.calls.map(c => c[0]);
    const pdfFiles = () => fs.readdirSync(receiptsDir).filter(f => f.endsWith('.pdf'));

    it('reports saved, not_found, fetch_failed and ambiguous apart, with a count per status', async () => {
      const acme = receiptMessage({
        id: 'm-acme', date: '2026-06-29T07:12:00Z', from: 'billing@acme.example', subject: 'Your Acme receipt',
        html: '<p>Receipt number 1001-2001</p><p>Amount paid £42.00</p>',
        attachments: [pdfAttachment('a1', 'Invoice-0015.pdf')], content: { a1: PDF_B64 }
      });
      const umbrella = receiptMessage({
        id: 'm-umbrella', date: '2026-06-20T07:12:00Z', from: 'billing@umbrella.example', subject: 'Umbrella receipt',
        html: '<p>Amount paid $9.00</p>'
      });
      const search = (filter) => (filter.includes('globex')
        ? graphFailure(503, 'ServiceUnavailable', 'Try again later')
        : graphSearch([acme, umbrella])(filter));

      const out = await run([
        { vendor: 'Acme', from: 'billing@acme.example' },
        { vendor: 'Globex', from: 'billing@globex.example' },
        { vendor: 'Initech', from: 'billing@initech.example' },
        { vendor: 'Umbrella', from: 'billing@umbrella.example' },
        { vendor: 'Umbrella Ops', subjectContains: 'Umbrella receipt' }
      ], { messages: [acme, umbrella], search });

      const byVendor = (v) => out.manifest.find(e => e.vendor === v);
      expect(byVendor('Acme').status).toBe('saved');
      expect(byVendor('Globex').status).toBe('fetch_failed');
      expect(byVendor('Initech').status).toBe('not_found');
      const ambiguous = out.manifest.find(e => e.status === 'ambiguous');
      expect(ambiguous.candidates).toEqual(['Umbrella', 'Umbrella Ops']);

      expect(out.summary).toEqual({ saved: 1, skipped: 0, not_found: 1, fetch_failed: 1, ambiguous: 1, duplicate: 0, unmatched: 0 });
      expect(out.unmatched).toEqual([]);
      expect(out.missing).toEqual(['Initech']); // only not_found vendors
      expect(out.failed.map(f => f.vendor)).toEqual(['Globex']);
      expect(out.notes).toEqual([]);
    });

    describe('merchants sharing one payment-processor mailbox', () => {
      // Processors send every merchant's receipts from one address and tell
      // them apart only by the plus tag.
      const merchant = (id, tag, name, number) => receiptMessage({
        id, date: '2026-06-18T09:00:00Z', from: `invoice+statements+${tag}@processor.example`,
        subject: `Your receipt from ${name} #${number}`,
        html: `<p>Receipt number ${number}</p><p>Amount paid $20.00</p>`,
        attachments: [pdfAttachment(`att-${id}`, `Invoice-${number}.pdf`)], content: { [`att-${id}`]: PDF_B64 }
      });
      const rules = [
        { vendor: 'Acme', from: 'invoice+statements+acct_aaa@processor.example', subjectContains: 'receipt from Acme' },
        { vendor: 'Globex', from: 'invoice+statements+acct_bbb@processor.example', subjectContains: 'receipt from Globex' }
      ];
      const acme = merchant('m-acme', 'acct_aaa', 'Acme Corp', '1111-0001');
      const globex = merchant('m-globex', 'acct_bbb', 'Globex Ltd', '2222-0002');

      it('keeps each merchant\'s receipt with its own rule instead of making both ambiguous', async () => {
        const out = await run(rules, { messages: [acme, globex] });
        const owner = (id) => out.manifest.find(e => e.sourceMessageId === id);
        expect(owner('m-acme')).toMatchObject({ vendor: 'Acme', status: 'saved' });
        expect(owner('m-globex')).toMatchObject({ vendor: 'Globex', status: 'saved' });
        expect(out.summary.ambiguous).toBe(0);
        expect(owner('m-acme').forwarded).toBe(false);
      });

      it('does not hand another merchant\'s receipt to a rule with no receipt this month', async () => {
        const out = await run(rules, { messages: [globex] });
        expect(out.manifest.find(e => e.vendor === 'Acme').status).toBe('not_found');
        expect(out.manifest.find(e => e.sourceMessageId === 'm-globex')).toMatchObject({ vendor: 'Globex', status: 'saved' });
        expect(out.missing).toEqual(['Acme']);
        expect(out.summary.ambiguous).toBe(0);
      });
    });

    it('turns a failed vendor search into fetch_failed with the Graph error, listed in failed[]', async () => {
      const out = await run(
        [{ vendor: 'Globex', from: 'billing@globex.example' }],
        { search: () => graphFailure(503, 'ServiceUnavailable', 'Try again later') }
      );

      const entry = out.manifest[0];
      expect(entry).toMatchObject({ vendor: 'Globex', status: 'fetch_failed', attemptedRoute: 'search' });
      expect(entry.graphError).toMatchObject({ status: 503, code: 'ServiceUnavailable', requestId: 'req-503' });
      expect(entry.error).toContain('Search failed: Graph API error 503 ServiceUnavailable');
      expect(out.failed).toEqual([entry]);
      expect(out.missing).toEqual([]);
      expect(out.summary.ambiguous).toBe(0);
    });

    it('records an attachment fetch failure as fetch_failed and keeps a rendered copy that is not the invoice', async () => {
      const msg = receiptMessage({
        id: 'm1', date: '2026-06-29T07:12:00Z', from: 'billing@acme.example', subject: 'Your Acme receipt',
        html: '<p>Amount paid £42.00</p>',
        attachments: [pdfAttachment('a1', 'Invoice-0015.pdf')],
        content: { a1: graphFailure(400, 'ErrorInvalidRequest', 'Attachment content unavailable') }
      });
      const renderImpl = renderPdf();

      const out = await run([{ vendor: 'Acme', from: 'billing@acme.example' }], { messages: [msg], deps: { renderImpl } });

      const entry = out.manifest[0];
      expect(entry).toMatchObject({
        vendor: 'Acme', status: 'fetch_failed', attemptedRoute: 'attachment', messageId: 'm1',
        attachmentId: 'a1', attachmentName: 'Invoice-0015.pdf', amount: 42, currency: 'GBP', amountSource: 'body'
      });
      expect(entry.graphError).toMatchObject({ status: 400, code: 'ErrorInvalidRequest' });
      expect(entry.error).toContain('Could not fetch the content of attachment');
      expect(entry.savedPath).toBeUndefined();

      expect(entry.renderedFallback).toMatchObject({ method: 'rendered', isFormalInvoice: false });
      expect(entry.renderedFallback.reason).toMatch(/not the formal invoice/);
      expect(entry.renderedFallback.sha256).toMatch(/^[0-9a-f]{64}$/);
      expect(path.basename(entry.renderedFallback.savedPath)).toBe('Acme 29Jun26 Invoice (rendered).pdf');
      expect(pdfFiles()).toEqual(['Acme 29Jun26 Invoice (rendered).pdf']); // never under the invoice name

      expect(entry).toMatchObject({ method: 'rendered', isFormalInvoice: false, attemptedRoutes: ['json', 'value'] });

      expect(out.failed).toHaveLength(1);
      expect(out.failed[0].renderedFallback.savedPath).toBe(entry.renderedFallback.savedPath);
      expect(out.failed[0]).toMatchObject({ method: 'rendered', isFormalInvoice: false, attemptedRoutes: ['json', 'value'] });
      expect(out.summary).toMatchObject({ fetch_failed: 1, saved: 0 });
    });

    it('notes a failed degraded render without hiding the fetch failure', async () => {
      const msg = receiptMessage({
        id: 'm1', date: '2026-06-29T07:12:00Z', from: 'billing@acme.example', subject: 'Your Acme receipt',
        html: '<p>Amount paid £42.00</p>',
        attachments: [pdfAttachment('a1', 'Invoice-0015.pdf')],
        content: { a1: graphFailure(500, 'InternalServerError', 'boom') }
      });
      const renderImpl = vi.fn().mockRejectedValue(new Error('Chrome not available'));

      const out = await run([{ vendor: 'Acme', from: 'billing@acme.example' }], { messages: [msg], deps: { renderImpl } });

      expect(out.manifest[0].status).toBe('fetch_failed');
      expect(out.manifest[0].renderedFallback).toBeUndefined();
      expect(out.manifest[0].renderedFallbackError).toContain('Chrome not available');
      expect(out.manifest[0].method).toBeUndefined(); // nothing was saved
      expect(out.manifest[0].attemptedRoutes).toEqual(['json', 'value']);
      expect(pdfFiles()).toEqual([]);
    });

    it('ignores messages in Sent Items and Drafts', async () => {
      const inbox = receiptMessage({
        id: 'm-inbox', date: '2026-06-10T09:00:00Z', from: 'billing@acme.example', subject: 'Your Acme receipt',
        html: '<p>Receipt number 1001-3001</p><p>Amount paid $10.00</p>'
      });
      const sentReply = receiptMessage({
        id: 'm-sent', date: '2026-06-11T09:00:00Z', from: 'me@example.com', subject: 'RE: Your Acme receipt',
        html: '<p>Thanks</p><p>Receipt number 1001-3001</p>', parentFolderId: 'folder-sent'
      });
      const draft = receiptMessage({
        id: 'm-draft', date: '2026-06-12T09:00:00Z', from: 'me@example.com', subject: 'FW: Your Acme receipt',
        html: '<p>Receipt number 1001-3001</p>', parentFolderId: 'folder-drafts'
      });
      const globexForward = receiptMessage({
        id: 'm-globex-sent', date: '2026-06-13T09:00:00Z', from: 'me@example.com', subject: 'Fw: Globex receipt',
        html: '<p>Amount paid $5.00</p>', parentFolderId: 'folder-sent'
      });

      const out = await run([
        { vendor: 'Acme', subjectContains: 'Your Acme receipt' },
        { vendor: 'Globex', subjectContains: 'Globex receipt' }
      ], { messages: [sentReply, draft, inbox, globexForward], deps: { renderImpl: renderPdf() } });

      const acme = out.manifest.filter(e => e.vendor === 'Acme');
      expect(acme).toHaveLength(1);
      expect(acme[0]).toMatchObject({ status: 'saved', sourceMessageId: 'm-inbox' });
      expect(out.manifest.find(e => e.vendor === 'Globex')).toEqual({ vendor: 'Globex', status: 'not_found', excludedByFolder: 1 });

      expect(graphPaths()).toContain('/me/mailFolders/sentitems');
      expect(graphPaths()).toContain('/me/mailFolders/drafts');
      expect(makeRequest.mock.calls.find(c => c[0] === '/me/mailFolders/sentitems')[1]).toEqual({ select: 'id' });
      expect(graphPaths().filter(p => /m-sent|m-draft|m-globex-sent/.test(p))).toEqual([]);
    });

    it('excludes nothing and leaves a note when the Sent Items folder cannot be resolved', async () => {
      const sentReply = receiptMessage({
        id: 'm-sent', date: '2026-06-11T09:00:00Z', from: 'billing@acme.example', subject: 'Your Acme receipt',
        html: '<p>Amount paid $10.00</p>', parentFolderId: 'folder-sent'
      });

      const out = await run([{ vendor: 'Acme', from: 'billing@acme.example' }], {
        messages: [sentReply],
        folders: { sentitems: graphFailure(403, 'ErrorAccessDenied', 'Access is denied') },
        deps: { renderImpl: renderPdf() }
      });

      expect(out.manifest[0]).toMatchObject({ status: 'saved', sourceMessageId: 'm-sent' });
      expect(out.notes).toHaveLength(1);
      expect(out.notes[0]).toMatch(/sentitems folder .*Access is denied.*not excluded/);
    });

    describe('copies of one receipt', () => {
      const RULE = { vendor: 'Acme', from: 'billing@acme.example', subjectContains: 'Your Acme receipt' };
      const quoted = '<p>From: Acme Billing &lt;billing@acme.example&gt;<br>Sent: Wednesday 10 June 09:00<br>' +
        'To: me@example.com<br>Subject: Your Acme receipt</p>';
      const receiptBody = '<p>Receipt number 2060-1001</p><p>Amount paid $42.00</p>';

      const original = receiptMessage({
        id: 'm-orig', date: '2026-06-10T09:00:00Z', from: 'billing@acme.example', subject: 'Your Acme receipt',
        html: receiptBody,
        attachments: [pdfAttachment('a1', 'Invoice-0015.pdf')],
        content: { a1: Buffer.from('%PDF-1.4 acme 2060-1001').toString('base64') }
      });
      const forward = receiptMessage({
        id: 'm-fw', date: '2026-06-11T09:00:00Z', from: 'colleague@example.com', subject: 'Fw: Your Acme receipt',
        html: `<p>FYI</p>${quoted}${receiptBody}`,
        attachments: [pdfAttachment('a1', 'Invoice-0015.pdf')],
        content: { a1: Buffer.from('%PDF-1.4 acme 2060-1001').toString('base64') }
      });
      const reply = receiptMessage({
        id: 'm-re', date: '2026-06-12T09:00:00Z', from: 'colleague@example.com', subject: 'RE: Your Acme receipt',
        html: `<p>Paid, thanks</p>${quoted}${receiptBody}`
      });

      it('saves one representative and marks the Fw: and Re: copies as duplicates', async () => {
        // Graph returns newest first.
        const out = await run([RULE], { messages: [reply, forward, original] });

        const saved = out.manifest.filter(e => e.status === 'saved');
        expect(saved).toHaveLength(1);
        expect(saved[0]).toMatchObject({
          sourceMessageId: 'm-orig', method: 'attachment', isFormalInvoice: true,
          receiptNumber: '2060-1001', sender: 'billing@acme.example', forwarded: false
        });
        expect(saved[0].matchedBy).toEqual(['from', 'subject']);
        expect(saved[0].originalSender).toBeUndefined();

        const duplicates = out.manifest.filter(e => e.status === 'duplicate');
        expect(duplicates.map(d => d.sourceMessageId).sort()).toEqual(['m-fw', 'm-re']);
        for (const d of duplicates) {
          expect(d).toMatchObject({
            vendor: 'Acme', duplicate_of: 'receiptNumber:2060-1001', duplicateOfMessageId: 'm-orig',
            forwarded: true, sender: 'colleague@example.com',
            originalSender: 'billing@acme.example', originalSenderName: 'Acme Billing'
          });
          expect(d.matchedBy).toEqual(['subject']); // matched on the normalised subject
          expect(d.amount).toBeUndefined(); // never counted twice
          expect(d.savedPath).toBeUndefined();
        }
        expect(out.manifest.find(d => d.sourceMessageId === 'm-fw').date).toBe('2026-06-11T09:00:00Z');

        expect(pdfFiles()).toEqual(['Acme 10Jun26 Invoice.pdf']);
        expect(graphPaths()).not.toContain('/me/messages/m-fw/attachments/a1');
        expect(out.summary).toMatchObject({ saved: 1, duplicate: 2 });
      });

      it('prefers the copy carrying an Invoice-*.pdf as the representative', async () => {
        const receiptOnly = receiptMessage({
          id: 'm-early', date: '2026-06-09T09:00:00Z', from: 'billing@acme.example', subject: 'Your Acme receipt',
          html: receiptBody,
          attachments: [pdfAttachment('r1', 'Receipt-2060-1001.pdf')],
          content: { r1: Buffer.from('%PDF-1.4 receipt copy').toString('base64') }
        });

        const out = await run([RULE], { messages: [forward, receiptOnly] });

        const saved = out.manifest.find(e => e.status === 'saved');
        expect(saved.sourceMessageId).toBe('m-fw');
        expect(saved.attachmentName).toBe('Invoice-0015.pdf');
        expect(saved.forwarded).toBe(true);
        expect(out.manifest.find(e => e.status === 'duplicate').sourceMessageId).toBe('m-early');
      });
    });

    it('flags a receipt as forwarded when the sender is not the rule sender', async () => {
      const msg = receiptMessage({
        id: 'm1', date: '2026-06-10T09:00:00Z', from: 'accounts@example.com', subject: 'Your Acme receipt',
        html: '<p>Amount paid $42.00</p>'
      });

      const out = await run(
        [{ vendor: 'Acme', from: 'billing@acme.example', subjectContains: 'Acme receipt' }],
        { messages: [msg], deps: { renderImpl: renderPdf() } }
      );

      expect(out.manifest[0]).toMatchObject({
        status: 'saved', forwarded: true, sender: 'accounts@example.com', matchedBy: ['subject']
      });
    });

    describe('product matchers', () => {
      const FROM = 'billing@acme.example';
      const itemReceipt = (id, day, item, number) => receiptMessage({
        id, date: `2026-06-${day}T09:00:00Z`, from: FROM, subject: 'Your Acme receipt',
        html: `<p>Receipt number ${number}</p><table><tr><td>${item}</td><td>$12.00</td></tr></table><p>Total $12.00</p>`
      });
      const studio = itemReceipt('m-studio', '20', 'Studio plan - 3x', '3000-0001');
      const gadget = itemReceipt('m-gadget', '15', 'Gadget kit - Lite', '3000-0002');
      const support = itemReceipt('m-support', '10', 'Support plan', '3000-0003');
      const extractCalls = (id) => makeRequest.mock.calls
        .filter(c => c[0] === `/me/messages/${id}` && c[1]?.select?.includes('hasAttachments')).length;

      it('splits one sender into three labelled streams with productContains/productExcludes', async () => {
        const out = await run([
          { vendor: 'Acme Studio', from: FROM, productContains: 'studio' },
          { vendor: 'Acme Gadget', from: FROM, productContains: ['gadget'], productExcludes: 'Studio' },
          { vendor: 'Acme Other', from: FROM, productExcludes: ['Studio', 'Gadget'] }
        ], { messages: [studio, gadget, support], deps: { renderImpl: renderPdf() } });

        const owner = (id) => out.manifest.find(e => e.sourceMessageId === id);
        expect(owner('m-studio')).toMatchObject({ vendor: 'Acme Studio', status: 'saved', productLabel: 'Studio plan - 3x' });
        expect(owner('m-studio').matchedBy).toEqual(['from', 'product:contains:studio']);
        expect(owner('m-gadget')).toMatchObject({ vendor: 'Acme Gadget', status: 'saved', productLabel: 'Gadget kit - Lite' });
        expect(owner('m-gadget').matchedBy).toEqual(['from', 'product:contains:gadget', 'product:excludes']);
        expect(owner('m-support')).toMatchObject({ vendor: 'Acme Other', status: 'saved', productLabel: 'Support plan' });
        expect(owner('m-support').matchedBy).toEqual(['from', 'product:excludes']);
        expect(out.manifest).toHaveLength(3);
        expect(out.missing).toEqual([]);

        // Found by all three searches, but each message is read once.
        for (const id of ['m-studio', 'm-gadget', 'm-support']) expect(extractCalls(id)).toBe(1);
        expect(out.notes).toContain('Acme Studio: 2 message(s) excluded by productContains/productExcludes');
        expect(out.unmatched).toEqual([]);
      });

      it('gives a message to the rule whose productContains matched over a rule without product matchers', async () => {
        const out = await run([
          { vendor: 'Acme', from: FROM },
          { vendor: 'Acme Studio', from: FROM, productContains: 'Studio' },
          { vendor: 'Acme Plans', from: FROM, productContains: 'plan' }
        ], { messages: [studio, gadget, support], deps: { renderImpl: renderPdf() } });

        const entry = (id) => out.manifest.find(e => e.sourceMessageId === id);
        expect(entry('m-gadget')).toMatchObject({ vendor: 'Acme', status: 'saved' });     // only the catch-all claims it
        expect(entry('m-support')).toMatchObject({ vendor: 'Acme Plans', status: 'saved' }); // specific beats catch-all
        // Two specific rules both match: ambiguous between them, the catch-all is not a candidate.
        expect(entry('m-studio')).toMatchObject({ status: 'ambiguous', vendor: null, candidates: ['Acme Studio', 'Acme Plans'] });
        expect(entry('m-studio').reason).toMatch(/claimed by 2 vendor rules/);
      });

      it('reports ambiguous, never guesses, when a product matcher meets a receipt with no product', async () => {
        const noProduct = receiptMessage({
          id: 'm1', date: '2026-06-10T09:00:00Z', from: FROM, subject: 'Your Acme receipt',
          html: '<p>Amount paid $5.00</p>'
        });
        const renderImpl = renderPdf();

        const out = await run([{ vendor: 'Acme Studio', from: FROM, productContains: 'Studio' }], {
          messages: [noProduct], deps: { renderImpl }
        });

        expect(out.manifest).toHaveLength(1);
        expect(out.manifest[0]).toMatchObject({
          vendor: 'Acme Studio', status: 'ambiguous', candidates: ['Acme Studio'], sourceMessageId: 'm1'
        });
        expect(out.manifest[0].reason).toMatch(/product could not be determined/);
        expect(out.missing).toEqual([]);
        expect(renderImpl).not.toHaveBeenCalled();
        expect(pdfFiles()).toEqual([]);
      });

      it('rejects product matchers that are not a string or an array of strings', async () => {
        const call = (rule) => collectReceiptsTool(authManager, { ...PERIOD, vendors: [{ vendor: 'Acme', from: FROM, ...rule }] });

        for (const rule of [
          { productContains: 5 },
          { productContains: { term: 'Studio' } },
          { productContains: [] },
          { productContains: '   ' },
          { productExcludes: ['Studio', 3] },
          { productExcludes: true }
        ]) {
          const res = await call(rule);
          expect(res.isError).toBe(true);
          expect(res.content[0].text).toMatch(/product(Contains|Excludes)/);
        }
        expect(makeRequest).not.toHaveBeenCalled();

        makeRequest.mockImplementation(mailboxRouter({}));
        expect((await call({ productContains: 'Studio', productExcludes: ['Gadget', 'Kit'] })).isError).toBeFalsy();
      });
    });

    it('lets two rules that claim one message produce a single ambiguous entry and save nothing', async () => {
      const msg = receiptMessage({
        id: 'm1', date: '2026-06-10T09:00:00Z', from: 'billing@acme.example', subject: 'Your Acme receipt',
        html: '<p>Amount paid $42.00</p>',
        attachments: [pdfAttachment('a1', 'Invoice-0015.pdf')], content: { a1: PDF_B64 }
      });

      const out = await run([
        { vendor: 'Acme', from: 'billing@acme.example' },
        { vendor: 'Acme Billing', subjectContains: 'Acme receipt' }
      ], { messages: [msg] });

      expect(out.manifest).toHaveLength(1);
      expect(out.manifest[0]).toMatchObject({
        status: 'ambiguous', vendor: null, candidates: ['Acme', 'Acme Billing'], sourceMessageId: 'm1'
      });
      expect(out.missing).toEqual([]);
      expect(graphPaths()).not.toContain('/me/messages/m1/attachments/a1');
      expect(pdfFiles()).toEqual([]);
    });

    describe('billing links', () => {
      let savedAllowlist;
      beforeEach(() => {
        savedAllowlist = process.env.BILLING_DOMAIN_ALLOWLIST;
        process.env.BILLING_DOMAIN_ALLOWLIST = 'billing.example.com';
      });
      afterEach(() => {
        if (savedAllowlist === undefined) delete process.env.BILLING_DOMAIN_ALLOWLIST;
        else process.env.BILLING_DOMAIN_ALLOWLIST = savedAllowlist;
      });

      it('skips a link shared by several receipts and renders those instead', async () => {
        const portal = 'https://billing.example.com/portal/invoices/latest.pdf';
        const linked = (id, day, number, url, subject = 'Initech receipt') => receiptMessage({
          id, date: `2026-06-${day}T09:00:00Z`, from: 'billing@initech.example', subject,
          html: `<p>Receipt number ${number}</p><p>Amount paid $5.00</p><a href="${url}">Download invoice</a>`
        });
        const messages = [
          linked('m1', '05', '77-1', portal),
          linked('m2', '06', '77-2', portal),
          linked('m3', '07', '77-3', 'https://billing.example.com/r/77-3/invoice.pdf'),
          // Two copies of one receipt sharing a link still count as one receipt.
          linked('m4', '08', '88-1', 'https://billing.example.com/r/88-1/invoice.pdf'),
          linked('m5', '09', '88-1', 'https://billing.example.com/r/88-1/invoice.pdf', 'Fw: Initech receipt')
        ];
        const fetchImpl = vi.fn().mockImplementation(async (url) => {
          const bytes = Buffer.from(`%PDF-1.4 ${url}`);
          return {
            status: 200, ok: true,
            headers: { get: (k) => (k.toLowerCase() === 'content-type' ? 'application/pdf' : null) },
            arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)
          };
        });

        const out = await run([{ vendor: 'Initech', from: 'billing@initech.example' }], {
          messages, deps: { fetchImpl, renderImpl: renderPdf() }
        });

        const entry = (id) => out.manifest.find(e => e.sourceMessageId === id);
        for (const id of ['m1', 'm2']) {
          expect(entry(id)).toMatchObject({ status: 'saved', method: 'rendered', isFormalInvoice: false, amountSource: 'body' });
          expect(entry(id).linkSkipped).toMatch(/shared by 2 receipts/);
        }
        expect(entry('m3')).toMatchObject({ status: 'saved', method: 'link', isFormalInvoice: true });
        expect(entry('m4')).toMatchObject({ status: 'saved', method: 'link' });
        expect(entry('m4').linkSkipped).toBeUndefined();
        expect(entry('m5')).toMatchObject({ status: 'duplicate', duplicate_of: 'receiptNumber:88-1' });
        expect(fetchImpl.mock.calls.map(c => c[0]).sort()).toEqual([
          'https://billing.example.com/r/77-3/invoice.pdf',
          'https://billing.example.com/r/88-1/invoice.pdf'
        ]);
      });
    });

    it('prefers the amounts printed in the saved PDF and flags a disagreement with the email', async () => {
      const acme = receiptMessage({
        id: 'm-acme', date: '2026-06-29T09:00:00Z', from: 'billing@acme.example', subject: 'Your Acme invoice',
        html: '<p>Amount paid £86.40</p>',
        attachments: [pdfAttachment('a1', 'Invoice-0042.pdf')],
        content: { a1: fixturePdf('acme-invoice-uk.pdf').toString('base64') }
      });
      const globex = receiptMessage({
        id: 'm-globex', date: '2026-06-23T09:00:00Z', from: 'billing@globex.example', subject: 'Your Globex receipt',
        html: '<p>Amount paid $25.00</p>',
        attachments: [pdfAttachment('a1', 'Receipt-4321.pdf')],
        content: { a1: fixturePdf('globex-receipt-usage.pdf').toString('base64') }
      });

      const out = await run([
        { vendor: 'Acme', from: 'billing@acme.example', expectedCurrency: 'GBP' },
        { vendor: 'Globex', from: 'billing@globex.example', expectedCurrency: 'USD' }
      ], { messages: [acme, globex] });

      const a = out.manifest.find(e => e.vendor === 'Acme');
      expect(a).toMatchObject({
        status: 'saved', amountSource: 'pdf', amount: 86.4, currency: 'GBP', taxAmount: 14.4,
        productLabel: 'Studio plan - 3x; Gadget plan - Lite', invoiceNumber: 'ZQXV9876-0042'
      });
      expect(a.amounts).toEqual({
        subtotal: 72, tax: 14.4, credits: null, total: 86.4, amountPaid: 86.4, currency: 'GBP', source: 'pdf',
        totalLabel: 'amount paid'
      });
      expect(a.amountMismatch).toBeUndefined();
      expect(a.currencyMismatch).toBeUndefined();

      const g = out.manifest.find(e => e.vendor === 'Globex');
      expect(g).toMatchObject({ status: 'saved', amountSource: 'pdf', amount: 22, currency: 'USD', receiptNumber: '5555-0000-4321' });
      expect(g.amounts).toMatchObject({ credits: 12.5, total: 22, source: 'pdf' });
      expect(g.amountMismatch).toEqual({ body: 25, pdf: 22 });
    }, 30000);

    it('keeps body amounts and a note when the saved PDF cannot be parsed', async () => {
      const msg = receiptMessage({
        id: 'm1', date: '2026-06-29T07:12:00Z', from: 'billing@acme.example', subject: 'Your Acme receipt',
        html: '<p>Amount paid £42.00</p>',
        attachments: [pdfAttachment('a1', 'Invoice-0015.pdf')], content: { a1: PDF_B64 } // header only, no PDF body
      });
      vi.spyOn(console, 'error').mockImplementation(() => {}); // pdfjs warnings

      const out = await run([{ vendor: 'Acme', from: 'billing@acme.example' }], { messages: [msg] });

      expect(out.manifest[0]).toMatchObject({ status: 'saved', amount: 42, currency: 'GBP', amountSource: 'body' });
      expect(out.manifest[0].pdfParseNote).toMatch(/could not be parsed/);
      vi.restoreAllMocks();
    });

    it('marks an identical PDF saved for another message as a sha256 duplicate and removes the extra file', async () => {
      const bytes = Buffer.from('%PDF-1.4 same invoice bytes').toString('base64');
      const later = receiptMessage({
        id: 'm-later', date: '2026-06-20T09:00:00Z', from: 'billing@acme.example', subject: 'Your Acme receipt',
        html: '<p>Receipt number 5000-0002</p><p>Amount paid $42.00</p>',
        attachments: [pdfAttachment('a1', 'Invoice-0015.pdf')], content: { a1: bytes }
      });
      const earlier = receiptMessage({
        id: 'm-earlier', date: '2026-06-10T09:00:00Z', from: 'billing@acme.example', subject: 'Your Acme receipt',
        html: '<p>Amount paid $42.00</p>', // no number, so nothing conflicts with the sha256 match
        attachments: [pdfAttachment('a1', 'Invoice-0015.pdf')], content: { a1: bytes }
      });
      vi.spyOn(console, 'error').mockImplementation(() => {});

      const out = await run([{ vendor: 'Acme', from: 'billing@acme.example' }], { messages: [later, earlier] });
      vi.restoreAllMocks();

      const saved = out.manifest.find(e => e.status === 'saved');
      const duplicate = out.manifest.find(e => e.status === 'duplicate');
      expect(saved.sourceMessageId).toBe('m-later'); // processed first
      expect(duplicate).toMatchObject({
        sourceMessageId: 'm-earlier', duplicate_of: `sha256:${saved.sha256}`, duplicateOfMessageId: 'm-later'
      });
      expect(duplicate.amount).toBeUndefined();
      expect(duplicate.savedPath).toBeUndefined();
      expect(fs.existsSync(duplicate.removedPath)).toBe(false);
      expect(pdfFiles()).toEqual([path.basename(saved.savedPath)]);
    });

    it('reports a file already on disk as skipped and reads its amounts', async () => {
      fs.writeFileSync(path.join(receiptsDir, 'Acme 29Jun26 Invoice.pdf'), fixturePdf('acme-invoice-uk.pdf'));
      const msg = receiptMessage({
        id: 'm1', date: '2026-06-29T07:12:00Z', from: 'billing@acme.example', subject: 'Your Acme invoice',
        html: '<p>Amount paid £86.40</p>',
        attachments: [pdfAttachment('a1', 'Invoice-0015.pdf')],
        content: { a1: fixturePdf('acme-invoice-uk.pdf').toString('base64') } // the same bytes as the file on disk
      });

      const out = await run([{ vendor: 'Acme', from: 'billing@acme.example' }], { messages: [msg] });

      expect(out.manifest[0]).toMatchObject({
        status: 'skipped', action: 'skipped', method: 'attachment', amountSource: 'pdf', amount: 86.4
      });
      expect(out.summary).toMatchObject({ skipped: 1, saved: 0 });
      expect(fs.readFileSync(path.join(receiptsDir, 'Acme 29Jun26 Invoice.pdf'))).toEqual(fixturePdf('acme-invoice-uk.pdf'));
    }, 30000);

    describe('file names and re-runs', () => {
      const FROM = 'billing@acme.example';
      const RULES = [{ vendor: 'Acme', from: FROM }];
      const pdfFor = (number, amount) => textPdf([`Receipt number ${number}`, `Amount paid $${amount}.00`]);
      const sameDay = (id, hour, number, amount) => receiptMessage({
        id, date: `2026-06-10T${hour}:00:00Z`, from: FROM, subject: 'Your Acme receipt',
        html: `<p>Receipt number ${number}</p><p>Amount paid $${amount}.00</p>`,
        attachments: [pdfAttachment('a1', 'Invoice-0015.pdf')],
        content: { a1: pdfFor(number, amount).toString('base64') }
      });
      const a = sameDay('m-a', '09', '7000-0001', 10);
      const b = sameDay('m-b', '10', '7000-0002', 22);
      const c = sameDay('m-c', '11', '7000-0003', 77);

      it('keeps every same-day receipt with its own PDF and amounts when a re-run finds a new one', async () => {
        vi.spyOn(console, 'error').mockImplementation(() => {}); // pdfjs font warnings
        const first = await run(RULES, { messages: [b, a] }); // Graph returns newest first
        expect(first.manifest.map(e => e.status)).toEqual(['saved', 'saved']);

        const second = await run(RULES, { messages: [c, b, a] });
        vi.restoreAllMocks();

        const entry = (id) => second.manifest.find(e => e.sourceMessageId === id);
        expect(entry('m-c')).toMatchObject({ status: 'saved', amount: 77, amountSource: 'pdf', receiptNumber: '7000-0003' });
        expect(entry('m-b')).toMatchObject({ status: 'skipped', amount: 22, amountSource: 'pdf' });
        expect(entry('m-a')).toMatchObject({ status: 'skipped', amount: 10, amountSource: 'pdf' });
        for (const [message, number, amount] of [[c, '7000-0003', 77], [b, '7000-0002', 22], [a, '7000-0001', 10]]) {
          const saved = entry(message.header.id);
          expect(path.basename(saved.savedPath)).toBe(`Acme 10Jun26 Invoice ${number}.pdf`);
          expect(fs.readFileSync(saved.savedPath)).toEqual(pdfFor(number, amount));
        }
        expect(second.summary).toMatchObject({ saved: 1, skipped: 2, duplicate: 0 });
        expect(pdfFiles().sort()).toEqual([
          'Acme 10Jun26 Invoice 7000-0001.pdf', 'Acme 10Jun26 Invoice 7000-0002.pdf', 'Acme 10Jun26 Invoice 7000-0003.pdf'
        ]);
      }, 30000);

      it('falls back to an ordinal suffix only for same-day receipts without any number', async () => {
        const plain = (id, hour, html) => receiptMessage({
          id, date: `2026-06-10T${hour}:00:00Z`, from: FROM, subject: 'Your Acme receipt', html
        });
        const out = await run(RULES, {
          messages: [
            plain('m-x', '11', '<p>Amount paid $5.00</p>'),
            plain('m-y', '10', '<p>Amount paid $6.00</p>'),
            plain('m-z', '09', '<p>Receipt number 7100-0001</p><p>Amount paid $7.00</p>')
          ],
          deps: { renderImpl: renderPdf() }
        });

        const name = (id) => path.basename(out.manifest.find(e => e.sourceMessageId === id).savedPath);
        // Ordinals go oldest first: m-y (10:00) keeps the plain name, m-x (11:00) gets "(2)".
        expect(name('m-y')).toBe('Acme 10Jun26 Invoice.pdf');
        expect(name('m-x')).toBe('Acme 10Jun26 Invoice (2).pdf');
        expect(name('m-z')).toBe('Acme 10Jun26 Invoice 7100-0001.pdf');
      });

      it('never reports a different file under the name as skipped: the receipt is versioned and flagged', async () => {
        fs.writeFileSync(path.join(receiptsDir, 'Acme 29Jun26 Invoice.pdf'), fixturePdf('acme-invoice-uk.pdf'));
        const msg = receiptMessage({
          id: 'm1', date: '2026-06-29T07:12:00Z', from: FROM, subject: 'Your Acme receipt',
          html: '<p>Amount paid $77.00</p>',
          attachments: [pdfAttachment('a1', 'Invoice-0015.pdf')],
          content: { a1: pdfFor('7000-0009', 77).toString('base64') }
        });
        vi.spyOn(console, 'error').mockImplementation(() => {});

        const out = await run(RULES, { messages: [msg] });
        vi.restoreAllMocks();

        expect(out.manifest[0]).toMatchObject({
          status: 'saved', action: 'versioned', existingDiffers: true, amount: 77, currency: 'USD', amountSource: 'pdf'
        });
        expect(path.basename(out.manifest[0].savedPath)).toBe('Acme 29Jun26 Invoice (2).pdf');
        expect(fs.readFileSync(path.join(receiptsDir, 'Acme 29Jun26 Invoice.pdf'))).toEqual(fixturePdf('acme-invoice-uk.pdf'));
      }, 30000);

      it('skips an earlier rendering of the same email on re-run even though the render timestamps differ', async () => {
        let stamp = 0;
        const renderImpl = vi.fn().mockImplementation(async () => {
          stamp++;
          return Buffer.from(`%PDF-1.4\n/CreationDate (D:2026060112000${stamp}+00'00')\n/ModDate (D:2026060112000${stamp}+00'00')\nbody`);
        });
        const msg = receiptMessage({
          id: 'm1', date: '2026-06-10T09:00:00Z', from: FROM, subject: 'Your Acme receipt', html: '<p>Amount paid $5.00</p>'
        });

        await run(RULES, { messages: [msg], deps: { renderImpl } });
        const second = await run(RULES, { messages: [msg], deps: { renderImpl } });

        expect(renderImpl).toHaveBeenCalledTimes(2);
        expect(second.manifest[0]).toMatchObject({ status: 'skipped', method: 'rendered' });
        expect(pdfFiles()).toEqual(['Acme 10Jun26 Invoice.pdf']);
      });
    });

    it('keeps receipts with different numbers apart even when their PDF bytes are identical', async () => {
      const bytes = Buffer.from('%PDF-1.4 shared template bytes').toString('base64');
      const numbered = (id, day, number, amount) => receiptMessage({
        id, date: `2026-06-${day}T09:00:00Z`, from: 'billing@acme.example', subject: 'Your Acme receipt',
        html: `<p>Receipt number ${number}</p><p>Amount paid $${amount}.00</p>`,
        attachments: [pdfAttachment('a1', 'Invoice-0015.pdf')], content: { a1: bytes }
      });
      vi.spyOn(console, 'error').mockImplementation(() => {});

      const out = await run([{ vendor: 'Acme', from: 'billing@acme.example' }], {
        messages: [numbered('m-2222', '20', '2222', 35), numbered('m-1111', '10', '1111', 20)]
      });
      vi.restoreAllMocks();

      const entry = (id) => out.manifest.find(e => e.sourceMessageId === id);
      expect(entry('m-2222')).toMatchObject({ status: 'saved', receiptNumber: '2222', amount: 35 });
      expect(entry('m-1111')).toMatchObject({ status: 'saved', receiptNumber: '1111', amount: 20, sameFileAs: 'm-2222' });
      expect(fs.existsSync(entry('m-1111').savedPath)).toBe(true);
      expect(out.summary.duplicate).toBe(0);
    });

    describe('grouping copies by number', () => {
      const FROM = 'billing@acme.example';
      const RULES = [{ vendor: 'Acme', from: FROM }];
      const mail = (id, day, html, subject = 'Your Acme receipt') => receiptMessage({
        id, date: `2026-06-${day}T09:00:00Z`, from: FROM, subject, html
      });

      it('never merges charges that share an order number when their totals differ', async () => {
        const out = await run(RULES, {
          messages: [
            mail('m-35', '12', '<p>Order #5555</p><p>Amount paid $35.00</p>'),
            mail('m-20', '11', '<p>Order #5555</p><p>Amount paid $20.00</p>')
          ],
          deps: { renderImpl: renderPdf() }
        });

        expect(out.manifest.map(e => [e.sourceMessageId, e.status, e.amount]))
          .toEqual([['m-35', 'saved', 35], ['m-20', 'saved', 20]]);
        expect(out.manifest.filter(e => e.numberConflict)).toEqual([]);
      });

      it('groups copies sharing any receipt, invoice or order number', async () => {
        const out = await run(RULES, {
          messages: [
            mail('m-fw', '11', '<p>FYI</p><p>Invoice number INV-6001</p><p>Amount paid $20.00</p>', 'Fw: Your Acme receipt'),
            mail('m-orig', '10', '<p>Receipt number 6000-0001</p><p>Invoice number INV-6001</p><p>Amount paid $20.00</p>')
          ],
          deps: { renderImpl: renderPdf() }
        });

        const entry = (id) => out.manifest.find(e => e.sourceMessageId === id);
        expect(entry('m-orig')).toMatchObject({ status: 'saved', receiptNumber: '6000-0001' });
        expect(entry('m-fw')).toMatchObject({
          status: 'duplicate', duplicate_of: 'invoiceNumber:INV-6001', duplicateOfMessageId: 'm-orig'
        });
      });

      it('keeps receipts sharing a receipt number apart when their totals differ, noting the conflict on both', async () => {
        const out = await run(RULES, {
          messages: [
            mail('m-35', '12', '<p>Receipt number 7000-0001</p><p>Amount paid $35.00</p>'),
            mail('m-20', '11', '<p>Receipt number 7000-0001</p><p>Amount paid $20.00</p>')
          ],
          deps: { renderImpl: renderPdf() }
        });

        expect(out.manifest.map(e => e.status)).toEqual(['saved', 'saved']);
        const entry = (id) => out.manifest.find(e => e.sourceMessageId === id);
        expect(entry('m-35').numberConflict).toMatch(/receiptNumber 7000-0001.*m-20/);
        expect(entry('m-20').numberConflict).toMatch(/receiptNumber 7000-0001.*m-35/);
      });
    });

    describe('PDF totals versus email totals', () => {
      const msg = (html = '<p>Amount paid $42.00</p>') => receiptMessage({
        id: 'm1', date: '2026-06-29T07:12:00Z', from: 'billing@acme.example', subject: 'Your Acme receipt', html,
        attachments: [pdfAttachment('a1', 'Invoice-0015.pdf')], content: { a1: PDF_B64 }
      });
      const parsed = (amounts, extra = {}) => vi.fn().mockResolvedValue({
        amounts: { subtotal: null, tax: null, credits: null, amountPaid: null, currency: 'USD', source: 'pdf', ...amounts },
        lineItems: [], productLabel: null, invoiceNumber: null, receiptNumber: null, ...extra
      });
      const RULES = [{ vendor: 'Acme', from: 'billing@acme.example' }];

      it('keeps a labelled email total over a guessed (unlabelled) PDF total', async () => {
        const out = await run(RULES, { messages: [msg()], deps: { parsePdf: parsed({ total: 99, totalLabel: null }) } });
        expect(out.manifest[0]).toMatchObject({
          status: 'saved', amount: 42, currency: 'USD', amountSource: 'body', pdfTotalGuessed: true,
          amountMismatch: { body: 42, pdf: 99 }
        });
        expect(out.manifest[0].amounts.source).toBe('body');
      });

      it('keeps the email total when the real PDF parser finds only an unlabelled figure', async () => {
        const unlabelled = receiptMessage({
          id: 'm1', date: '2026-06-29T07:12:00Z', from: 'billing@acme.example', subject: 'Your Acme receipt',
          html: '<p>Amount paid $42.00</p>', attachments: [pdfAttachment('a1', 'Invoice-0015.pdf')],
          content: { a1: textPdf(['Thank you for your business', 'Reference $12.00']).toString('base64') }
        });
        vi.spyOn(console, 'error').mockImplementation(() => {});
        const out = await run(RULES, { messages: [unlabelled] });
        vi.restoreAllMocks();
        expect(out.manifest[0]).toMatchObject({ amount: 42, amountSource: 'body', pdfTotalGuessed: true });
      }, 30000);

      it('prefers a labelled PDF total', async () => {
        const out = await run(RULES, { messages: [msg()], deps: { parsePdf: parsed({ total: 99, totalLabel: 'total' }) } });
        expect(out.manifest[0]).toMatchObject({ amount: 99, amountSource: 'pdf', amountMismatch: { body: 42, pdf: 99 } });
        expect(out.manifest[0].pdfTotalGuessed).toBeUndefined();
      });

      it('uses a guessed PDF total, flagged, only when the email has none', async () => {
        const out = await run(RULES, {
          messages: [msg('<p>Thanks for your order</p>')], deps: { parsePdf: parsed({ total: 99, totalLabel: null }) }
        });
        expect(out.manifest[0]).toMatchObject({ amount: 99, amountSource: 'pdf', pdfTotalGuessed: true });
      });

      it('caps receipt and invoice numbers read from the PDF at 300 characters', async () => {
        const out = await run(RULES, {
          messages: [msg()],
          deps: { parsePdf: parsed({ total: 42, totalLabel: 'amount paid' }, { receiptNumber: 'R'.repeat(500), invoiceNumber: 'I'.repeat(400) }) }
        });
        expect(out.manifest[0].receiptNumber.startsWith('R'.repeat(300))).toBe(true);
        expect(out.manifest[0].receiptNumber.length).toBe(301); // 300 + ellipsis, as the parser caps
        expect(out.manifest[0].invoiceNumber.length).toBe(301);
      });
    });

    describe('product from the PDF and unmatched messages', () => {
      const FROM = 'billing@acme.example';
      const pdfOnly = (content) => receiptMessage({
        id: 'm1', date: '2026-06-29T09:00:00Z', from: FROM, subject: 'Your Acme invoice',
        html: '<p>Amount paid £86.40</p>', attachments: [pdfAttachment('a1', 'Invoice-0042.pdf')], content: { a1: content }
      });
      const SPLIT = [
        { vendor: 'Acme Studio', from: FROM, productContains: 'studio' },
        { vendor: 'Acme Other', from: FROM, productExcludes: 'studio' }
      ];

      it('reads the product from the PDF attachment before claiming when the email names none', async () => {
        const out = await run(SPLIT, { messages: [pdfOnly(fixturePdf('acme-invoice-uk.pdf').toString('base64'))] });

        const owned = out.manifest.filter(e => e.sourceMessageId === 'm1');
        expect(owned).toHaveLength(1);
        expect(owned[0]).toMatchObject({
          vendor: 'Acme Studio', status: 'saved', productLabel: 'Studio plan - 3x; Gadget plan - Lite', productSource: 'pdf'
        });
        expect(owned[0].matchedBy).toEqual(['from', 'product:contains:studio']);
        expect(out.manifest.find(e => e.vendor === 'Acme Other')).toEqual({ vendor: 'Acme Other', status: 'not_found', excludedByProduct: 1 });
        expect(out.unmatched).toEqual([]);
      }, 30000);

      it('stays ambiguous, with a note, when the PDF cannot be read for its product', async () => {
        const out = await run(SPLIT, { messages: [pdfOnly(graphFailure(500, 'InternalServerError', 'boom'))] });

        expect(out.manifest).toHaveLength(1);
        expect(out.manifest[0]).toMatchObject({ status: 'ambiguous', candidates: ['Acme Studio', 'Acme Other'] });
        expect(out.manifest[0].reason).toMatch(/product could not be determined/);
        expect(out.manifest[0].pdfProductError).toContain('Could not fetch the content of attachment');
      });

      it('lists messages every claiming rule rejected by product in unmatched[] and counts them', async () => {
        const item = (id, day, product, subject = 'Your Acme receipt') => receiptMessage({
          id, date: `2026-06-${day}T09:00:00Z`, from: FROM, subject,
          html: `<table><tr><td>${product}</td><td>$12.00</td></tr></table><p>Total $12.00</p>`
        });
        const longSubject = `Your Acme receipt ${'x'.repeat(400)}`;

        const out = await run([{ vendor: 'Acme Studio', from: FROM, productContains: 'studio' }], {
          messages: [item('m-studio', '20', 'Studio plan - 3x'), item('m-gadget', '15', 'Gadget kit - Lite', longSubject)],
          deps: { renderImpl: renderPdf() }
        });

        expect(out.manifest.map(e => e.sourceMessageId)).toEqual(['m-studio']);
        expect(out.unmatched).toHaveLength(1);
        expect(out.unmatched[0]).toMatchObject({
          messageId: 'm-gadget', date: '2026-06-15T09:00:00Z', sender: FROM,
          productLabel: 'Gadget kit - Lite', rejectedBy: ['Acme Studio']
        });
        expect(out.unmatched[0].subject.startsWith('Your Acme receipt xxx')).toBe(true);
        expect(out.unmatched[0].subject.length).toBeLessThanOrEqual(301);
        expect(out.summary.unmatched).toBe(1);
      });
    });

    describe('distinct charges that share a number', () => {
      const FROM = 'billing@acme.example';
      const RULES = [{ vendor: 'Acme', from: FROM, subjectContains: 'Your Acme receipt' }];
      const permutations = (items) => (items.length <= 1 ? [items] : items.flatMap((item, i) =>
        permutations([...items.slice(0, i), ...items.slice(i + 1)]).map(rest => [item, ...rest])));
      const charge = (id, day, receipt, label) => receiptMessage({
        id, date: `2026-06-${day}T09:00:00Z`, from: FROM, subject: 'Your Acme receipt',
        html: `<p>Receipt number ${receipt}</p><p>Order #ORD-77</p><p>Amount paid $10.00</p>`,
        attachments: [pdfAttachment('a1', `Receipt-${receipt}.pdf`)],
        content: { a1: Buffer.from(`%PDF-1.4 ${label}`).toString('base64') }
      });
      const r1 = charge('m-r1', '10', 'R-1001', 'charge one');
      const r2 = charge('m-r2', '11', 'R-1002', 'charge two');
      const reply = (attachments = [], content = {}) => receiptMessage({
        id: 'm-re', date: '2026-06-12T09:00:00Z', from: 'colleague@example.com', subject: 'RE: Your Acme receipt',
        html: '<p>Approved</p><p>Order #ORD-77</p>', attachments, content
      });
      const entry = (out, id) => out.manifest.find(e => e.sourceMessageId === id);
      const quiet = () => vi.spyOn(console, 'error').mockImplementation(() => {});

      it('saves both receipts when their receipt numbers differ although order number and total match', async () => {
        quiet();
        const out = await run(RULES, { messages: [r2, r1] });
        vi.restoreAllMocks();

        expect(entry(out, 'm-r1')).toMatchObject({ status: 'saved', receiptNumber: 'R-1001', amount: 10 });
        expect(entry(out, 'm-r2')).toMatchObject({ status: 'saved', receiptNumber: 'R-1002', amount: 10 });
        expect(fs.readFileSync(entry(out, 'm-r2').savedPath).toString()).toBe('%PDF-1.4 charge two');
        expect(out.summary).toMatchObject({ saved: 2, duplicate: 0 });
      });

      it.each(permutations(['r1', 'r2', 'reply']).map(order => [order.join(', '), order]))(
        'never lets a reply without a total merge or hide either charge (order: %s)', async (_label, order) => {
          const byName = { r1, r2, reply: reply() };
          quiet();
          const out = await run(RULES, { messages: order.map(name => byName[name]) });
          vi.restoreAllMocks();

          expect(entry(out, 'm-r1')).toMatchObject({ status: 'saved', amount: 10 });
          expect(entry(out, 'm-r2')).toMatchObject({ status: 'saved', amount: 10 });
          expect(fs.readFileSync(entry(out, 'm-r1').savedPath).toString()).toBe('%PDF-1.4 charge one');
          expect(fs.readFileSync(entry(out, 'm-r2').savedPath).toString()).toBe('%PDF-1.4 charge two');
          expect(entry(out, 'm-re').status).toBe('duplicate');
          expect(['m-r1', 'm-r2']).toContain(entry(out, 'm-re').duplicateOfMessageId);
          expect(out.summary).toMatchObject({ saved: 2, duplicate: 1 });
          expect(pdfFiles()).toHaveLength(2);
        });

      it.each([['reply first', true], ['reply last', false]])(
        'never makes a copy without a total the representative over a receipt carrying its PDF (%s)', async (_label, replyFirst) => {
          // The reply carries an Invoice-*.pdf, which would otherwise outrank the receipt's own Receipt-*.pdf.
          const withPdf = reply([pdfAttachment('i1', 'Invoice-0099.pdf')], {
            i1: Buffer.from('%PDF-1.4 unrelated invoice').toString('base64')
          });
          quiet();
          const out = await run(RULES, { messages: replyFirst ? [withPdf, r1] : [r1, withPdf] });
          vi.restoreAllMocks();

          expect(entry(out, 'm-r1')).toMatchObject({ status: 'saved', amount: 10, attachmentName: 'Receipt-R-1001.pdf' });
          expect(fs.readFileSync(entry(out, 'm-r1').savedPath).toString()).toBe('%PDF-1.4 charge one');
          expect(entry(out, 'm-re')).toMatchObject({ status: 'duplicate', duplicateOfMessageId: 'm-r1' });
        });

      describe('identical PDF bytes', () => {
        const same = Buffer.from('%PDF-1.4 one template for every receipt').toString('base64');
        const templated = (id, day, html) => receiptMessage({
          id, date: `2026-06-${day}T09:00:00Z`, from: FROM, subject: 'Your Acme receipt', html,
          attachments: [pdfAttachment('a1', 'Invoice-0015.pdf')], content: { a1: same }
        });
        const plain = templated('m-n', '09', '<p>Amount paid $10.00</p>');
        const one = templated('m-1', '10', '<p>Receipt number R-1001</p><p>Amount paid $10.00</p>');
        const two = templated('m-2', '11', '<p>Receipt number R-1002</p><p>Amount paid $10.00</p>');

        it.each(permutations(['plain', 'one', 'two']).map(order => [order.join(', '), order]))(
          'never lets an entry without numbers absorb two differently numbered receipts (order: %s)', async (_label, order) => {
            const byName = { plain, one, two };
            quiet();
            const out = await run(RULES, { messages: order.map(name => byName[name]) });
            vi.restoreAllMocks();

            expect(out.summary).toMatchObject({ saved: 2, duplicate: 1 });
            const kept = out.manifest.filter(e => e.status === 'saved');
            expect(kept.reduce((sum, e) => sum + e.amount, 0)).toBe(20);
            for (const [id, other] of [['m-1', 'm-2'], ['m-2', 'm-1']]) {
              const e = entry(out, id);
              if (e.status !== 'duplicate') continue;
              expect(e.duplicateOfMessageId).not.toBe(other);
              // Its representative's group never also holds the other number.
              expect(entry(out, other).duplicateOfMessageId).toBeUndefined();
            }
          });

        it('keeps entries with identical bytes apart when their known totals differ', async () => {
          quiet();
          const out = await run(RULES, {
            messages: [templated('m-20', '12', '<p>Amount paid $20.00</p>'), plain]
          });
          vi.restoreAllMocks();

          expect(entry(out, 'm-20')).toMatchObject({ status: 'saved', amount: 20 });
          expect(entry(out, 'm-n')).toMatchObject({ status: 'saved', amount: 10, sameFileAs: 'm-20' });
          expect(out.summary.duplicate).toBe(0);
        });
      });
    });

    describe('stable file names across runs', () => {
      const FROM = 'billing@acme.example';
      const RULES = [{ vendor: 'Acme', from: FROM }];
      const pdfFor = (number, amount) => textPdf([`Receipt number ${number}`, `Amount paid $${amount}.00`]);
      const numbered = (id, hour, number, amount) => receiptMessage({
        id, date: `2026-06-10T${hour}:00:00Z`, from: FROM, subject: 'Your Acme receipt',
        html: `<p>Receipt number ${number}</p><p>Amount paid $${amount}.00</p>`,
        attachments: [pdfAttachment('a1', 'Invoice-0015.pdf')],
        content: { a1: pdfFor(number, amount).toString('base64') }
      });
      const plain = (id, hour, amount) => receiptMessage({
        id, date: `2026-06-10T${hour}:00:00Z`, from: FROM, subject: 'Your Acme receipt',
        html: `<p>Amount paid $${amount}.00</p>`
      });
      // Same email, same bytes: a deterministic stand-in for Chrome.
      const stableRender = () => vi.fn().mockImplementation(async (doc) => Buffer.from(`%PDF-1.4\n${doc}`));
      const entry = (out, id) => out.manifest.find(e => e.sourceMessageId === id);

      it('reuses the plain-name file of a receipt that gains a same-day sibling in a later run', async () => {
        const a = numbered('m-a', '09', '7000-0001', 10);
        const b = numbered('m-b', '10', '7000-0002', 22);
        vi.spyOn(console, 'error').mockImplementation(() => {});
        const first = await run(RULES, { messages: [a] });
        expect(path.basename(first.manifest[0].savedPath)).toBe('Acme 10Jun26 Invoice.pdf');

        const second = await run(RULES, { messages: [b, a] });
        vi.restoreAllMocks();

        expect(entry(second, 'm-a')).toMatchObject({ status: 'skipped', action: 'skipped', amount: 10, amountSource: 'pdf' });
        expect(path.basename(entry(second, 'm-a').savedPath)).toBe('Acme 10Jun26 Invoice.pdf');
        expect(entry(second, 'm-b')).toMatchObject({ status: 'saved', amount: 22 });
        expect(path.basename(entry(second, 'm-b').savedPath)).toBe('Acme 10Jun26 Invoice 7000-0002.pdf');
        expect(pdfFiles().sort()).toEqual(['Acme 10Jun26 Invoice 7000-0002.pdf', 'Acme 10Jun26 Invoice.pdf']);
      }, 30000);

      it('gives receipts without numbers ordinals oldest-first, so a new receipt never re-saves older ones', async () => {
        const x = plain('m-x', '09', 5);
        const y = plain('m-y', '10', 6);
        const z = plain('m-z', '11', 7);
        const renderImpl = stableRender();

        await run(RULES, { messages: [x], deps: { renderImpl } });
        const second = await run(RULES, { messages: [y, x], deps: { renderImpl } }); // newest first
        expect(entry(second, 'm-x')).toMatchObject({ status: 'skipped' });
        expect(pdfFiles()).toHaveLength(2);

        const third = await run(RULES, { messages: [z, y, x], deps: { renderImpl } });
        const name = (id) => path.basename(entry(third, id).savedPath);
        expect([name('m-x'), name('m-y'), name('m-z')])
          .toEqual(['Acme 10Jun26 Invoice.pdf', 'Acme 10Jun26 Invoice (2).pdf', 'Acme 10Jun26 Invoice (3).pdf']);
        expect(third.manifest.map(e => [e.sourceMessageId, e.status]))
          .toEqual([['m-z', 'saved'], ['m-y', 'skipped'], ['m-x', 'skipped']]);
        expect(pdfFiles()).toHaveLength(3);
      });
    });

    describe('zero and guessed PDF totals', () => {
      const RULES = [{ vendor: 'Acme', from: 'billing@acme.example' }];
      const withPdf = (html, content) => receiptMessage({
        id: 'm1', date: '2026-06-29T07:12:00Z', from: 'billing@acme.example', subject: 'Your Acme receipt', html,
        attachments: [pdfAttachment('a1', 'Invoice-0015.pdf')], content: { a1: content }
      });
      const parsed = (amounts) => vi.fn().mockResolvedValue({
        amounts: { subtotal: null, tax: null, credits: null, amountPaid: null, currency: 'USD', source: 'pdf', ...amounts },
        lineItems: [], productLabel: null, invoiceNumber: null, receiptNumber: null
      });

      it('keeps a positive email total over a PDF that only shows a balance due of zero', async () => {
        const msg = withPdf('<p>Amount paid $48.00</p>', textPdf(['Invoice INV-0001', 'Balance due $0.00']).toString('base64'));
        vi.spyOn(console, 'error').mockImplementation(() => {});
        const out = await run(RULES, { messages: [msg] });
        vi.restoreAllMocks();

        expect(out.manifest[0]).toMatchObject({ amount: 48, amountSource: 'body', amountMismatch: { body: 48, pdf: 0 } });
      }, 30000);

      it('keeps a positive email total over a labelled PDF amount due of zero', async () => {
        const out = await run(RULES, {
          messages: [withPdf('<p>Amount paid $48.00</p>', PDF_B64)],
          deps: { parsePdf: parsed({ total: 0, totalLabel: 'amount due' }) }
        });
        expect(out.manifest[0]).toMatchObject({ amount: 48, amountSource: 'body', amountMismatch: { body: 48, pdf: 0 } });
        expect(out.manifest[0].amounts.source).toBe('body');
      });

      it('keeps the email figure when both totals are guesses, flagging the difference', async () => {
        const out = await run(RULES, {
          messages: [withPdf('<p>Thanks for your order</p><p>$48.00</p>', PDF_B64)],
          deps: { parsePdf: parsed({ total: 40, totalLabel: null }) }
        });
        expect(out.manifest[0]).toMatchObject({ amount: 48, amountSource: 'body', amountMismatch: { body: 48, pdf: 40 } });
      });
    });

    it('does not download an oversized PDF attachment just to read its product', async () => {
      const msg = receiptMessage({
        id: 'm1', date: '2026-06-29T09:00:00Z', from: 'billing@acme.example', subject: 'Your Acme invoice',
        html: '<p>Amount paid £86.40</p>',
        attachments: [{ ...pdfAttachment('a1', 'Invoice-0042.pdf'), size: 20 * 1024 * 1024 }],
        content: { a1: fixturePdf('acme-invoice-uk.pdf').toString('base64') }
      });

      const out = await run([
        { vendor: 'Acme Studio', from: 'billing@acme.example', productContains: 'studio' },
        { vendor: 'Acme Other', from: 'billing@acme.example', productExcludes: 'studio' }
      ], { messages: [msg] });

      expect(graphPaths()).not.toContain('/me/messages/m1/attachments/a1');
      expect(out.manifest).toHaveLength(1);
      expect(out.manifest[0]).toMatchObject({ status: 'ambiguous', candidates: ['Acme Studio', 'Acme Other'] });
      expect(out.manifest[0].pdfProductError).toMatch(/too large.*15 MB/);
    });

    describe('input validation before any Graph call', () => {
      const msg = receiptMessage({
        id: 'm1', date: '2026-06-29T07:12:00Z', from: 'billing@acme.example', subject: 'Your Acme receipt',
        html: '<p>Amount paid £42.00</p>', attachments: [pdfAttachment('a1', 'Invoice-0015.pdf')], content: { a1: PDF_B64 }
      });
      const call = (args) => {
        makeRequest.mockImplementation(mailboxRouter({ messages: [msg] }));
        return collectReceiptsTool(authManager, {
          ...PERIOD, vendors: [{ vendor: 'Acme', from: 'billing@acme.example' }], ...args
        }, { renderImpl: renderPdf() });
      };

      it('rejects a destDir outside the allowed roots, not a string, or not a directory', async () => {
        const notADir = path.join(receiptsDir, 'not-a-dir');
        fs.writeFileSync(notADir, 'x');
        for (const destDir of ['/etc', 42, notADir]) {
          const res = await call({ destDir });
          expect(res.isError).toBe(true);
          expect(res.content[0].text).toMatch(/destDir/);
        }
        expect(makeRequest).not.toHaveBeenCalled();
        expect(pdfFiles()).toEqual([]);
      });

      it('rejects malformed rule fields and onExisting', async () => {
        for (const args of [
          { vendors: [{ vendor: 'Acme', from: 'billing@acme.example', expectedCurrency: 5 }] },
          { vendors: [{ vendor: 7, from: 'billing@acme.example' }] },
          { vendors: [{ vendor: 'Acme', from: 42 }] },
          { vendors: [{ vendor: 'Acme', subjectContains: ['receipt'] }] },
          { vendors: [{ vendor: 'Acme', from: 'billing@acme.example', filenameTemplate: 3 }] },
          { vendors: [{ vendor: 'Acme', from: 'billing@acme.example', prefer: 'newest' }] },
          { vendors: ['Acme'] },
          { onExisting: 'replace' }
        ]) {
          const res = await call(args);
          expect(res.isError).toBe(true);
          expect(res.content[0].text).toMatch(/Invalid parameter '(vendors|onExisting)'/);
        }
        expect(makeRequest).not.toHaveBeenCalled();
        expect(pdfFiles()).toEqual([]);
      });
    });
  });
});
