import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { extractReceiptTool } from '../../tools/receipts/extractReceipt.js';

describe('extractReceiptTool (FR-3)', () => {
  let makeRequest, authManager, dir, savedRulesPath;

  beforeEach(() => {
    makeRequest = vi.fn();
    authManager = {
      ensureAuthenticated: vi.fn().mockResolvedValue(true),
      getGraphApiClient: vi.fn().mockReturnValue({ makeRequest })
    };
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rules-'));
    savedRulesPath = process.env.RECEIPT_RULES_PATH;
    delete process.env.RECEIPT_RULES_PATH;
  });

  afterEach(() => {
    if (savedRulesPath === undefined) delete process.env.RECEIPT_RULES_PATH;
    else process.env.RECEIPT_RULES_PATH = savedRulesPath;
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

  it('hard-errors when RECEIPT_RULES_PATH points at a missing file', async () => {
    process.env.RECEIPT_RULES_PATH = path.join(dir, 'missing.json');
    const res = await extractReceiptTool(authManager, { messageId: 'm1' });
    expect(res.isError).toBe(true);
    expect(res.content[0].text).toContain('RECEIPT_RULES_PATH');
    expect(res.content[0].text).toContain('missing.json');
    expect(makeRequest).not.toHaveBeenCalled();
  });
});
