import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
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

describe('collectReceiptsTool (FR-6)', () => {
  let receiptsDir, makeRequest, authManager, savedEnv;

  beforeEach(() => {
    receiptsDir = fs.mkdtempSync(path.join(os.tmpdir(), 'receipts-'));
    savedEnv = process.env.MCP_OUTLOOK_RECEIPTS_DIR;
    process.env.MCP_OUTLOOK_RECEIPTS_DIR = receiptsDir;
    makeRequest = vi.fn();
    authManager = {
      ensureAuthenticated: vi.fn().mockResolvedValue(true),
      getGraphApiClient: vi.fn().mockReturnValue({ makeRequest })
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

  it('saves an attachment receipt and flags an absent vendor as missing', async () => {
    makeRequest
      // vendor 1: Acme — search, then extract (message + attachments), then save (list + full)
      .mockResolvedValueOnce({ value: [stripeMessage('m1', '2026-06-29T07:12:00Z')] })
      .mockResolvedValueOnce(STRIPE_FULL('m1', '2026-06-29T07:12:00Z'))
      .mockResolvedValueOnce(ATTACHMENT_LIST)
      .mockResolvedValueOnce(ATTACHMENT_LIST)
      .mockResolvedValueOnce(ATTACHMENT_FULL)
      // vendor 2: Vandelay — search returns nothing
      .mockResolvedValueOnce({ value: [] });

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
    expect(out.manifest.find(e => e.vendor === 'Vandelay').status).toBe('missing');

    // search used date-bounded, sender-filtered whole-mailbox query
    const [endpoint, options] = makeRequest.mock.calls[0];
    expect(endpoint).toBe('/me/messages');
    expect(options.filter).toContain("receivedDateTime ge 2026-06-01T00:00:00Z");
    expect(options.filter).toContain("from/emailAddress/address eq 'invoice+statements@mail.acme.com'");
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
    // Each run needs: search + extract msg + extract attachments + save list + save full
    makeRequest
      // Run 1
      .mockResolvedValueOnce({ value: [stripeMessage('m1', '2026-06-29T07:12:00Z')] })
      .mockResolvedValueOnce(STRIPE_FULL('m1', '2026-06-29T07:12:00Z'))
      .mockResolvedValueOnce(ATTACHMENT_LIST)
      .mockResolvedValueOnce(ATTACHMENT_LIST)
      .mockResolvedValueOnce(ATTACHMENT_FULL)
      // Run 2
      .mockResolvedValueOnce({ value: [stripeMessage('m1', '2026-06-29T07:12:00Z')] })
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
    expect(out2.manifest[0].status).toBe('saved'); // status is still 'saved' even when skipped

    // Only one PDF file on disk — no duplicate
    const files = fs.readdirSync(receiptsDir).filter(f => f.endsWith('.pdf'));
    expect(files).toHaveLength(1);
  });

  // Item 6: currency mismatch — rule has expectedCurrency:'USD', receipt is GBP
  it('records currencyMismatch when receipt currency differs from expectedCurrency', async () => {
    makeRequest
      .mockResolvedValueOnce({ value: [stripeMessage('m1', '2026-06-29T07:12:00Z')] })
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
    expect(out.manifest[0].status).toBe('missing');
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

  it('marks a vendor entry ambiguous (with error) when saving fails, and continues', async () => {
    makeRequest
      .mockResolvedValueOnce({ value: [stripeMessage('m1', '2026-06-29T07:12:00Z')] })
      .mockResolvedValueOnce({ content: [{ type: 'text', text: 'Graph exploded' }], isError: true }) // extract fails
      .mockResolvedValueOnce({ value: [] }); // second vendor still processed

    const res = await collectReceiptsTool(authManager, {
      periodStart: '2026-06-01T00:00:00Z', periodEnd: '2026-06-30T23:59:59Z',
      vendors: [
        { vendor: 'Acme', from: 'invoice+statements@mail.acme.com' },
        { vendor: 'Vandelay', from: 'billing@vandelay.com' }
      ]
    });
    const out = JSON.parse(res.content[0].text);
    const failed = out.manifest.find(e => e.vendor === 'Acme');
    expect(failed.status).toBe('ambiguous');
    expect(failed.error).toContain('Graph exploded');
    expect(out.missing).toEqual(['Vandelay']);
  });
});
