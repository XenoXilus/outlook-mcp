import { describe, it, expect } from 'vitest';
import {
  detectVendor,
  extractAmount,
  extractTaxAmount,
  extractReceiptNumber,
  extractInvoiceNumber,
  extractProductLabel,
  extractBillingPdfUrl,
  selectPdfAttachment,
  parseReceiptMessage
} from '../../tools/receipts/receiptParser.js';

const ALLOWLIST = ['pay.stripe.com', 'invoice.stripe.com', 'files.stripe.com', 'm.stripe.network'];

const RULES = {
  vendorSenders: [
    { pattern: /billing@acme\.example$/i, vendor: 'Acme' },
    { pattern: /no-reply@globex\.example$/i, vendor: 'Globex' }
  ],
  productLabels: [/Pro plan[^\n<,£$€]*/i, /Team seat[^\n<,£$€]*/i]
};

const STRIPE_HTML = `
<html><body>
<h1>Receipt from Acme, Inc</h1>
<p>Receipt #4242-1337-0001</p>
<p>Invoice #TSTINV42-0001</p>
<table><tr><td>Pro plan - 5x</td><td>&pound;42.00</td></tr></table>
<p>Amount paid &pound;42.00</p>
<a href="https://pay.stripe.com/invoice/acct_1/test_123/pdf?s=em">Download invoice</a>
<a href="https://evil.example.com/invoice.pdf">bad link</a>
</body></html>`;

describe('receiptParser', () => {
  it('detects vendors from configured sender rules', () => {
    expect(detectVendor('billing@acme.example', 'Order confirmation', RULES.vendorSenders)).toBe('Acme');
    expect(detectVendor('no-reply@globex.example', 'Your receipt', RULES.vendorSenders)).toBe('Globex');
  });

  it('sender rules win over the subject heuristic', () => {
    expect(detectVendor('billing@acme.example', 'Your receipt from Initech #1', RULES.vendorSenders)).toBe('Acme');
  });

  it('derives vendor from "receipt from X" subjects with no rules configured', () => {
    expect(detectVendor('invoice+statements+acct_1@stripe.com', 'Your receipt from Initech #2286-4084')).toBe('Initech');
    expect(detectVendor('invoice+statements@stripe.com', 'Your receipt from Acme, Inc #4242')).toBe('Acme');
  });

  it('returns null with no rules and an unmatchable subject', () => {
    expect(detectVendor('no-reply@somewhere.example', 'Payment processed')).toBeNull();
  });

  it('extracts labelled amounts with currency', () => {
    expect(extractAmount('Amount paid £42.00')).toEqual({ amount: 42.0, currency: 'GBP' });
    expect(extractAmount('Total: $13.37 charged')).toEqual({ amount: 13.37, currency: 'USD' });
    expect(extractAmount('nothing here')).toEqual({ amount: null, currency: null });
  });

  it('extracts receipt and invoice numbers', () => {
    expect(extractReceiptNumber('Receipt #4242-1337-0001')).toBe('4242-1337-0001');
    expect(extractInvoiceNumber('Invoice #TSTINV42-0001')).toBe('TSTINV42-0001');
  });

  it('extracts product labels only from configured patterns', () => {
    expect(extractProductLabel('1x Pro plan - 5x £42.00', RULES.productLabels)).toMatch(/Pro plan/);
    expect(extractProductLabel('1x Pro plan - 5x £42.00')).toBeNull(); // no rules => null
    expect(extractProductLabel('nothing known', RULES.productLabels)).toBeNull();
  });

  it('extracts only allowlisted billing links, preferring pdf URLs', () => {
    expect(extractBillingPdfUrl(STRIPE_HTML, ALLOWLIST)).toBe('https://pay.stripe.com/invoice/acct_1/test_123/pdf?s=em');
    expect(extractBillingPdfUrl('<a href="https://evil.example.com/x.pdf">x</a>', ALLOWLIST)).toBeNull();
  });

  // Fix 3: HTML-entity decode in extractBillingPdfUrl
  it('decodes &amp; HTML entity in extracted billing PDF URLs', () => {
    const html = '<a href="https://pay.stripe.com/x?a=1&amp;b=2">Invoice</a>';
    expect(extractBillingPdfUrl(html, ALLOWLIST)).toBe('https://pay.stripe.com/x?a=1&b=2');
  });

  it('single-param billing URLs pass through without modification', () => {
    const html = '<a href="https://pay.stripe.com/x?token=abc123">Invoice</a>';
    expect(extractBillingPdfUrl(html, ALLOWLIST)).toBe('https://pay.stripe.com/x?token=abc123');
  });

  it('selects the Invoice-*.pdf when a message has Invoice + Receipt PDFs', () => {
    const attachments = [
      { id: 'a1', name: 'Receipt-4242-1337-0001.pdf', contentType: 'application/pdf', isInline: false },
      { id: 'a2', name: 'Invoice-TSTINV42-0001.pdf', contentType: 'application/pdf', isInline: false },
      { id: 'a3', name: 'logo.png', contentType: 'image/png', isInline: true }
    ];
    expect(selectPdfAttachment(attachments).id).toBe('a2');
    expect(selectPdfAttachment(attachments, 'receipt').id).toBe('a1');
    expect(selectPdfAttachment(attachments, 'first').id).toBe('a1');
    expect(selectPdfAttachment([{ id: 'x', name: 'a.png', contentType: 'image/png', isInline: false }])).toBeNull();
  });

  it('extractTaxAmount: parses VAT/tax lines and returns null when absent', () => {
    expect(extractTaxAmount('VAT £20.00')).toBe(20.0);
    expect(extractTaxAmount('Tax: $5.99')).toBe(5.99);
    expect(extractTaxAmount('Amount paid £42.00')).toBeNull();
  });

  it('truncates absurdly long fields so the result stays under 4 KB', () => {
    const longStr = 'A'.repeat(20000);
    const message = {
      subject: longStr,
      from: { emailAddress: { address: 'billing@acme.example' } },
      receivedDateTime: '2026-06-29T07:12:00Z',
      body: { contentType: 'text', content: `Amount paid £42.00 ${longStr}` }
    };
    const attachments = Array.from({ length: 30 }, (_, i) => ({
      id: `a${i}`, name: `file${i}.pdf`, contentType: 'application/pdf', isInline: false
    }));
    const result = parseReceiptMessage(message, attachments, []);
    expect(JSON.stringify(result).length).toBeLessThan(4096);
    expect(result.subject.length).toBeLessThan(310);
    expect(result.attachmentIds.length).toBe(20);
  });

  it('billingPdfUrl is never truncated even when it exceeds 300 chars', () => {
    const token = 'x'.repeat(280);
    const longUrl = `https://pay.stripe.com/invoice/acct_1/pdf?s=${token}&extra=1`;
    expect(longUrl.length).toBeGreaterThan(300);
    const html = `<html><body><a href="${longUrl}">Invoice</a></body></html>`;
    const message = {
      subject: 'Your receipt from Acme, Inc #4242-1337-0001',
      from: { emailAddress: { address: 'invoice+statements@stripe.com' } },
      receivedDateTime: '2026-06-29T07:12:00Z',
      body: { contentType: 'html', content: html }
    };
    const result = parseReceiptMessage(message, [], ALLOWLIST);
    expect(result.billingPdfUrl).toBe(longUrl);
    expect(result.billingPdfUrl).not.toContain('…');
  });

  it('parses a full receipt message into the FR-3 shape using configured rules', () => {
    const message = {
      id: 'msg1',
      subject: 'Your receipt from Acme, Inc #4242-1337-0001',
      from: { emailAddress: { address: 'billing@acme.example', name: 'Acme' } },
      receivedDateTime: '2026-06-29T07:12:00Z',
      body: { contentType: 'html', content: STRIPE_HTML }
    };
    const attachments = [
      { id: 'a1', name: 'Invoice-TSTINV42-0001.pdf', contentType: 'application/pdf', isInline: false },
      { id: 'a2', name: 'Receipt-4242-1337-0001.pdf', contentType: 'application/pdf', isInline: false }
    ];
    const r = parseReceiptMessage(message, attachments, ALLOWLIST, RULES);
    expect(r.vendor).toBe('Acme');
    expect(r.amount).toBe(42.0);
    expect(r.currency).toBe('GBP');
    expect(r.receiptNumber).toBe('4242-1337-0001');
    expect(r.invoiceNumber).toBe('TSTINV42-0001');
    expect(r.productLabel).toMatch(/Pro plan/);
    expect(r.billingPdfUrl).toContain('pay.stripe.com');
    expect(r.hasPdfAttachment).toBe(true);
    expect(r.attachmentIds).toEqual(['a1', 'a2']);
    expect(JSON.stringify(r).length).toBeLessThan(4096);
  });

  it('parses without rules: vendor from subject heuristic, productLabel null', () => {
    const message = {
      id: 'msg1',
      subject: 'Your receipt from Acme, Inc #4242-1337-0001',
      from: { emailAddress: { address: 'invoice+statements@stripe.com' } },
      receivedDateTime: '2026-06-29T07:12:00Z',
      body: { contentType: 'html', content: STRIPE_HTML }
    };
    const r = parseReceiptMessage(message, [], ALLOWLIST);
    expect(r.vendor).toBe('Acme');
    expect(r.productLabel).toBeNull();
  });

  it('trims attachmentIds and sets truncated:true when total (ex billingPdfUrl) exceeds 4096 bytes', () => {
    const fatId = (i) => `AAMkADEA${i.toString().padStart(3, '0')}${'x'.repeat(190)}`;
    const attachments = Array.from({ length: 20 }, (_, i) => ({
      id: fatId(i), name: `file${i}.pdf`, contentType: 'application/pdf', isInline: false
    }));
    const message = {
      subject: 'Your receipt from Acme, Inc #4242-1337-0001',
      from: { emailAddress: { address: 'invoice+statements@stripe.com' } },
      receivedDateTime: '2026-06-29T07:12:00Z',
      body: { contentType: 'text', content: 'Amount paid £42.00' }
    };
    const result = parseReceiptMessage(message, attachments, []);
    const sizeWithoutUrl = JSON.stringify({ ...result, billingPdfUrl: undefined }).length;
    expect(sizeWithoutUrl).toBeLessThanOrEqual(4096);
    expect(result.attachmentIds.length).toBeLessThan(20);
    expect(result.truncated).toBe(true);
  });

  it('leaves attachmentIds untrimmed and truncated falsy for a normal small message', () => {
    const message = {
      subject: 'Your receipt from Acme, Inc #4242-1337-0001',
      from: { emailAddress: { address: 'invoice+statements@stripe.com' } },
      receivedDateTime: '2026-06-29T07:12:00Z',
      body: { contentType: 'html', content: STRIPE_HTML }
    };
    const attachments = [
      { id: 'a1', name: 'Invoice-TSTINV42-0001.pdf', contentType: 'application/pdf', isInline: false },
      { id: 'a2', name: 'Receipt-4242-1337-0001.pdf', contentType: 'application/pdf', isInline: false }
    ];
    const result = parseReceiptMessage(message, attachments, []);
    expect(result.attachmentIds).toEqual(['a1', 'a2']);
    expect(result.truncated).toBeFalsy();
  });
});
