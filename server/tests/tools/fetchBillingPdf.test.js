import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { fetchBillingPdfTool } from '../../tools/receipts/fetchBillingPdf.js';

const PDF_BYTES = Buffer.from('%PDF-1.4 stripe invoice');

function pdfResponse() {
  return {
    status: 200, ok: true,
    headers: { get: (k) => (k.toLowerCase() === 'content-type' ? 'application/pdf' : null) },
    arrayBuffer: async () => PDF_BYTES.buffer.slice(PDF_BYTES.byteOffset, PDF_BYTES.byteOffset + PDF_BYTES.byteLength)
  };
}

describe('fetchBillingPdfTool (FR-2)', () => {
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

  it('requires fileName and one of url/messageId', async () => {
    const noFile = await fetchBillingPdfTool(authManager, { url: 'https://pay.stripe.com/x' });
    expect(noFile.isError).toBe(true);
    const noSource = await fetchBillingPdfTool(authManager, { fileName: 'x.pdf' });
    expect(noSource.isError).toBe(true);
  });

  it('fetches an explicit allowlisted URL and saves it', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(pdfResponse());
    const res = await fetchBillingPdfTool(authManager, {
      url: 'https://pay.stripe.com/invoice/x/pdf', fileName: 'Initech 23Jun26 Invoice.pdf'
    }, { fetchImpl });
    const out = JSON.parse(res.content[0].text);
    expect(out.savedPath).toBe(path.join(receiptsDir, 'Initech 23Jun26 Invoice.pdf'));
    expect(fs.readFileSync(out.savedPath).equals(PDF_BYTES)).toBe(true);
    expect(out.sourceUrl).toContain('pay.stripe.com');
  });

  it('refuses off-allowlist URLs with a clear error and no fetch', async () => {
    const fetchImpl = vi.fn();
    const res = await fetchBillingPdfTool(authManager, {
      url: 'https://evil.example.com/x.pdf', fileName: 'x.pdf'
    }, { fetchImpl });
    expect(res.isError).toBe(true);
    expect(res.content[0].text).toMatch(/allowlist/i);
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it('extracts the billing link from the message body when only messageId is given', async () => {
    makeRequest.mockResolvedValueOnce({
      id: 'm1', subject: 'receipt',
      body: { contentType: 'html', content: '<a href="https://pay.stripe.com/invoice/a/b/pdf?s=em">Download invoice</a>' }
    });
    const fetchImpl = vi.fn().mockResolvedValue(pdfResponse());
    const res = await fetchBillingPdfTool(authManager, { messageId: 'm1', fileName: 'x.pdf' }, { fetchImpl });
    const out = JSON.parse(res.content[0].text);
    expect(out.sourceUrl).toBe('https://pay.stripe.com/invoice/a/b/pdf?s=em');
    expect(makeRequest.mock.calls[0][0]).toBe('/me/messages/m1');
  });

  it('errors when the message body has no allowlisted link, and never calls fetch', async () => {
    makeRequest.mockResolvedValueOnce({ id: 'm1', body: { contentType: 'html', content: '<p>no links</p>' } });
    const fetchImpl = vi.fn();
    const res = await fetchBillingPdfTool(authManager, { messageId: 'm1', fileName: 'x.pdf' }, { fetchImpl });
    expect(res.isError).toBe(true);
    expect(res.content[0].text).toMatch(/No allowlisted billing link/i);
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  // 6b: sha256 in the result (FR-2 requires SHA-256)
  it('includes sha256 matching a 64-char hex string in the result', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(pdfResponse());
    const res = await fetchBillingPdfTool(authManager, {
      url: 'https://pay.stripe.com/invoice/x/pdf', fileName: 'SHA256-check Invoice.pdf'
    }, { fetchImpl });
    const out = JSON.parse(res.content[0].text);
    expect(out.sha256).toMatch(/^[0-9a-f]{64}$/);
  });
});
