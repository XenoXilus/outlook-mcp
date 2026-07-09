import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { getBillingAllowlist, assertAllowlisted, fetchPdf } from '../../tools/receipts/pdfFetch.js';

const PDF_BYTES = Buffer.from('%PDF-1.4 fake');

function mockResponse({ status = 200, headers = {}, body = PDF_BYTES } = {}) {
  return {
    status,
    ok: status >= 200 && status < 300,
    headers: { get: (k) => headers[k.toLowerCase()] ?? null },
    arrayBuffer: async () => body.buffer.slice(body.byteOffset, body.byteOffset + body.byteLength)
  };
}

describe('pdfFetch', () => {
  let saved;
  beforeEach(() => { saved = process.env.BILLING_DOMAIN_ALLOWLIST; });
  afterEach(() => {
    if (saved === undefined) delete process.env.BILLING_DOMAIN_ALLOWLIST;
    else process.env.BILLING_DOMAIN_ALLOWLIST = saved;
  });

  it('defaults the allowlist to the Stripe hosts', () => {
    delete process.env.BILLING_DOMAIN_ALLOWLIST;
    expect(getBillingAllowlist()).toEqual(['pay.stripe.com', 'invoice.stripe.com', 'files.stripe.com', 'm.stripe.network']);
  });

  it('honours BILLING_DOMAIN_ALLOWLIST', () => {
    process.env.BILLING_DOMAIN_ALLOWLIST = 'a.example.com, b.example.com';
    expect(getBillingAllowlist()).toEqual(['a.example.com', 'b.example.com']);
  });

  it('rejects non-HTTPS and off-allowlist hosts', () => {
    const list = ['pay.stripe.com'];
    expect(() => assertAllowlisted('http://pay.stripe.com/x', list)).toThrow(/https/i);
    expect(() => assertAllowlisted('https://evil.example.com/x', list)).toThrow(/allowlist/i);
    expect(assertAllowlisted('https://pay.stripe.com/x', list).hostname).toBe('pay.stripe.com');
  });

  it('fetches a valid PDF', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(mockResponse({
      headers: { 'content-type': 'application/pdf', 'content-length': String(PDF_BYTES.length) }
    }));
    const res = await fetchPdf('https://pay.stripe.com/invoice/x/pdf', { fetchImpl });
    expect(res.buffer.equals(PDF_BYTES)).toBe(true);
    expect(fetchImpl).toHaveBeenCalledOnce();
  });

  it('rejects wrong content-type and non-PDF bytes', async () => {
    const html = vi.fn().mockResolvedValue(mockResponse({ headers: { 'content-type': 'text/html' } }));
    await expect(fetchPdf('https://pay.stripe.com/x', { fetchImpl: html })).rejects.toThrow(/application\/pdf/);

    const fake = vi.fn().mockResolvedValue(mockResponse({
      headers: { 'content-type': 'application/pdf' }, body: Buffer.from('<html>nope</html>')
    }));
    await expect(fetchPdf('https://pay.stripe.com/x', { fetchImpl: fake })).rejects.toThrow(/%PDF/);
  });

  it('follows redirects only within the allowlist', async () => {
    const fetchImpl = vi.fn()
      .mockResolvedValueOnce(mockResponse({ status: 302, headers: { location: 'https://files.stripe.com/f.pdf' } }))
      .mockResolvedValueOnce(mockResponse({ headers: { 'content-type': 'application/pdf' } }));
    const res = await fetchPdf('https://pay.stripe.com/x', { fetchImpl });
    expect(res.finalUrl).toBe('https://files.stripe.com/f.pdf');

    const offList = vi.fn().mockResolvedValue(mockResponse({ status: 302, headers: { location: 'https://evil.example.com/f.pdf' } }));
    await expect(fetchPdf('https://pay.stripe.com/x', { fetchImpl: offList })).rejects.toThrow(/allowlist/i);
  });

  it('enforces the max size', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(mockResponse({
      headers: { 'content-type': 'application/pdf' }, body: Buffer.concat([Buffer.from('%PDF'), Buffer.alloc(64)])
    }));
    await expect(fetchPdf('https://pay.stripe.com/x', { fetchImpl, maxBytes: 16 })).rejects.toThrow(/size/i);
  });

  // 6a: port pinning
  it('rejects non-default ports on allowlisted hosts', () => {
    const list = ['pay.stripe.com'];
    expect(() => assertAllowlisted('https://pay.stripe.com:8080/x', list)).toThrow(/port/i);
    // Explicit :443 and omitted port are both allowed
    expect(assertAllowlisted('https://pay.stripe.com:443/x', list).hostname).toBe('pay.stripe.com');
    expect(assertAllowlisted('https://pay.stripe.com/x', list).hostname).toBe('pay.stripe.com');
  });

  // 6b: timeout path
  it('reports a readable timeout error when fetch is aborted', async () => {
    const fetchImpl = (_url, { signal }) => new Promise((_resolve, reject) => {
      signal.addEventListener('abort', () => {
        const err = new Error('The operation was aborted');
        err.name = 'AbortError';
        reject(err);
      });
    });
    await expect(fetchPdf('https://pay.stripe.com/x', { fetchImpl, timeoutMs: 50 }))
      .rejects.toThrow(/timed out/i);
  });

  // 6b: content-length pre-check — body must not be read
  it('rejects early on oversized content-length without reading the body', async () => {
    const mockResp = mockResponse({
      headers: { 'content-type': 'application/pdf', 'content-length': '99999999' }
    });
    const arraySpy = vi.spyOn(mockResp, 'arrayBuffer');
    const fetchImpl = vi.fn().mockResolvedValue(mockResp);
    await expect(fetchPdf('https://pay.stripe.com/x', { fetchImpl, maxBytes: 1000 }))
      .rejects.toThrow(/size/i);
    expect(arraySpy).not.toHaveBeenCalled();
  });
});
