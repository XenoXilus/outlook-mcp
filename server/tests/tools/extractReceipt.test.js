import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { extractReceiptTool } from '../../tools/receipts/extractReceipt.js';

describe('extractReceiptTool (FR-3)', () => {
  let makeRequest, authManager, dir, savedRulesPath, savedAllowlist;

  beforeEach(() => {
    makeRequest = vi.fn();
    authManager = {
      ensureAuthenticated: vi.fn().mockResolvedValue(true),
      getGraphApiClient: vi.fn().mockReturnValue({ makeRequest })
    };
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rules-'));
    savedRulesPath = process.env.RECEIPT_RULES_PATH;
    delete process.env.RECEIPT_RULES_PATH;
    savedAllowlist = process.env.BILLING_DOMAIN_ALLOWLIST;
    delete process.env.BILLING_DOMAIN_ALLOWLIST;
  });

  afterEach(() => {
    if (savedRulesPath === undefined) delete process.env.RECEIPT_RULES_PATH;
    else process.env.RECEIPT_RULES_PATH = savedRulesPath;
    if (savedAllowlist === undefined) delete process.env.BILLING_DOMAIN_ALLOWLIST;
    else process.env.BILLING_DOMAIN_ALLOWLIST = savedAllowlist;
    fs.rmSync(dir, { recursive: true, force: true });
  });

  it('requires messageId', async () => {
    const res = await extractReceiptTool(authManager, {});
    expect(res.isError).toBe(true);
  });

  it('returns the compact FR-3 object, never the raw HTML body', async () => {
    const bigHtml = `<html><body>${'x'.repeat(60000)}<p>Amount paid £42.00</p><p>Receipt #4242-1337-0001</p></body></html>`;
    makeRequest
      .mockResolvedValueOnce({
        id: 'm1',
        subject: 'Your receipt from Acme, Inc #4242-1337-0001',
        from: { emailAddress: { address: 'invoice+statements@stripe.com' } },
        receivedDateTime: '2026-06-29T07:12:00Z',
        hasAttachments: true,
        body: { contentType: 'html', content: bigHtml }
      })
      .mockResolvedValueOnce({ value: [
        { id: 'a1', name: 'Invoice-TSTINV42-0001.pdf', contentType: 'application/pdf', isInline: false }
      ]});

    const res = await extractReceiptTool(authManager, { messageId: 'm1' });
    const out = JSON.parse(res.content[0].text);
    expect(out.vendor).toBe('Acme'); // via the generic subject heuristic, no rules file
    expect(out.amount).toBe(42.0);
    expect(out.currency).toBe('GBP');
    expect(out.hasPdfAttachment).toBe(true);
    expect(out.attachmentIds).toEqual(['a1']);
    expect(res.content[0].text.length).toBeLessThan(4096);
  });

  it('applies sender/product rules from RECEIPT_RULES_PATH', async () => {
    const rulesFile = path.join(dir, 'rules.json');
    fs.writeFileSync(rulesFile, JSON.stringify({
      vendorSenders: [{ pattern: 'no-reply@globex\\.example$', vendor: 'Globex' }],
      productLabels: ['Team seat[^\\n<,]*']
    }));
    process.env.RECEIPT_RULES_PATH = rulesFile;

    makeRequest.mockResolvedValueOnce({
      id: 'm2', subject: 'Payment confirmation',
      from: { emailAddress: { address: 'no-reply@globex.example' } },
      receivedDateTime: '2026-06-09T10:00:00Z', hasAttachments: false,
      body: { contentType: 'html', content: '<p>Team seat x2</p><p>Total: £13.37</p>' }
    });
    const res = await extractReceiptTool(authManager, { messageId: 'm2' });
    const out = JSON.parse(res.content[0].text);
    expect(out.vendor).toBe('Globex');
    expect(out.productLabel).toMatch(/Team seat/);
    expect(out.amount).toBe(13.37);
    expect(makeRequest).toHaveBeenCalledTimes(1);
  });

  // A forwarded processor-style receipt: tracked links, a merchant-logo background on an
  // allowlisted host, and an invoice attachment. Merchant, amounts and ids are invented.
  const INVOICE_TARGET = 'https://pay.stripe.com/invoice/acct_TESTACCT01/test_YWNjdF9URVNU/pdf?s=em';
  const RECEIPT_TARGET = 'https://dashboard.stripe.com/receipts/invoices/CAcQTESTRECEIPT01/pdf?s=em';
  const LOGO = 'https://files.stripe.com/files/MDB8TESTLOGOID01';
  const wrap = (target) =>
    `https://58.email.stripe.com/CL0/${encodeURIComponent(target).replace(/%3A/gi, ':').replace(/%3D/gi, '=')}` +
    '/1/010101aabbccdd00-11112222-3333-4444-5555-666677778888-000000/TESTSIGNATURE_x1y2=452';
  const forwardedHtml = () => {
    const rows = [
      'Forwarding this one for the records.',
      'From: Acme, Inc &lt;invoice+statements@acme.example&gt;',
      'Sent: Saturday, 14 March 2026 09:12:05',
      'To: Dana Example &lt;dana@example.com&gt;',
      'Subject: Your receipt from Acme, Inc #8080-1234-5678',
      'Receipt from Acme, Inc',
      'A$550.00',
      'Paid March 14, 2026',
      `<a href="${wrap(INVOICE_TARGET)}">Download invoice</a> <a href="${wrap(RECEIPT_TARGET)}">Download receipt</a>`,
      'Receipt number 8080-1234-5678',
      'Gadget plan - Pro',
      'Qty 3',
      'A$360.00',
      'Gadget plan - Lite',
      'Qty 7',
      'A$140.00',
      'Subtotal',
      'A$500.00',
      'GST - Australia (10%)',
      'A$50.00',
      'Total',
      'A$550.00',
      'Amount paid',
      'A$550.00'
    ];
    return `<html><body><a href="https://acme.example"><div style="background:url('${LOGO}')"></div></a>` +
      rows.map(r => `<div>${r}</div>`).join('') + '</body></html>';
  };
  const mockForwarded = () => makeRequest
    .mockResolvedValueOnce({
      id: 'm3',
      subject: 'Re: Fwd: Your receipt from Acme, Inc #8080-1234-5678',
      from: { emailAddress: { address: 'dana@example.com' } },
      receivedDateTime: '2026-03-15T10:04:31Z',
      hasAttachments: true,
      body: { contentType: 'html', content: forwardedHtml() }
    })
    .mockResolvedValueOnce({ value: [
      { id: 'a1', name: 'Invoice-TSTINV77-0003.pdf', contentType: 'application/pdf', isInline: false },
      { id: 'a2', name: 'Receipt-8080-1234-5678.pdf', contentType: 'application/pdf', isInline: false }
    ]});

  it('returns amounts, line items, quoted sender and the unwrapped invoice link for a forwarded receipt', async () => {
    mockForwarded();
    const res = await extractReceiptTool(authManager, { messageId: 'm3' });
    const out = JSON.parse(res.content[0].text);

    expect(out.replyOrForward).toBe(true);
    expect(out.normalisedSubject).toBe('Your receipt from Acme, Inc #8080-1234-5678');
    expect(out.quotedFrom).toEqual({ name: 'Acme, Inc', address: 'invoice+statements@acme.example' });
    expect(out.vendor).toBe('Acme');
    expect(out.amounts).toEqual({
      subtotal: 500, tax: 65, credits: null, total: 565, amountPaid: 565, currency: 'CAD', source: 'body'
    });
    expect(out.amount).toBe(550);
    expect(out.currency).toBe('AUD');
    expect(out.taxAmount).toBe(50);
    expect(out.lineItems).toEqual([
      { description: 'Gadget plan - Pro', quantity: 3, amount: 360 },
      { description: 'Gadget plan - Lite', quantity: 7, amount: 140 }
    ]);
    expect(out.productLabel).toBe('Gadget plan - Pro; Gadget plan - Lite');
    expect(out.receiptNumber).toBe('8080-1234-5678');
    expect(out.invoiceNumber).toBe('TSTINV77-0003'); // from the attachment name
    expect(out.billingPdfUrl).toBe(INVOICE_TARGET);
    expect(out.billingPdfUrl).not.toContain('files.stripe.com');
    expect(out.receiptPdfUrl).toBeNull(); // dashboard host is not on the default allowlist
    expect(out.attachmentIds).toEqual(['a1', 'a2']);
    expect(res.content[0].text).not.toContain('<div');
  });

  it('reports receiptPdfUrl when BILLING_DOMAIN_ALLOWLIST includes the receipt host', async () => {
    process.env.BILLING_DOMAIN_ALLOWLIST = 'pay.stripe.com,dashboard.stripe.com'; // restored in afterEach
    mockForwarded();
    const out = JSON.parse((await extractReceiptTool(authManager, { messageId: 'm3' })).content[0].text);
    expect(out.billingPdfUrl).toBe(INVOICE_TARGET);
    expect(out.receiptPdfUrl).toBe(RECEIPT_TARGET);
  });

  it('hard-errors when RECEIPT_RULES_PATH points at a missing file', async () => {
    process.env.RECEIPT_RULES_PATH = path.join(dir, 'missing.json');
    const res = await extractReceiptTool(authManager, { messageId: 'm1' });
    expect(res.isError).toBe(true);
    expect(res.content[0].text).toContain('RECEIPT_RULES_PATH');
    expect(res.content[0].text).toContain('missing.json');
    expect(makeRequest).not.toHaveBeenCalled();
  });
});
