import { describe, it, expect, vi, afterEach } from 'vitest';
import fs from 'fs';
import path from 'path';
import { execFileSync } from 'child_process';
import { fileURLToPath } from 'url';
import { extractPdfText, parseReceiptPdf } from '../../tools/receipts/receiptPdfText.js';

// Fictional PDFs rendered from the .html files beside them (see the comment at the top of each).
const here = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(here, '../../..');
const fixture = (name) => fs.readFileSync(path.join(repoRoot, 'server/tests/fixtures/receipts', name));

describe('receiptPdfText', () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  describe('extractPdfText', () => {
    it('returns the text of a real PDF, one line per baseline', async () => {
      const text = await extractPdfText(fixture('acme-invoice-uk.pdf'));
      expect(text).toContain('Invoice number ZQXV9876-0042');
      expect(text.split('\n')).toContain('Total £86.40');
      expect(text.split('\n')).toContain('VAT - United Kingdom (20% on £72.00) £14.40');
    });

    it('keeps a label and its value on consecutive lines when the layout stacks them', async () => {
      const lines = (await extractPdfText(fixture('globex-receipt-usage.pdf'))).split('\n');
      const i = lines.indexOf('Amount paid');
      expect(i).toBeGreaterThan(-1);
      expect(lines[i + 1]).toBe('$22.00');
    });

    it('accepts a Uint8Array as well as a Buffer', async () => {
      const bytes = new Uint8Array(fixture('acme-invoice-uk.pdf'));
      expect(await extractPdfText(bytes)).toContain('ZQXV9876-0042');
    });

    it('throws a clear error for non-PDF input', async () => {
      await expect(extractPdfText(Buffer.from('<html>not a pdf</html>'))).rejects.toThrow(/not a PDF/i);
      await expect(extractPdfText(Buffer.alloc(0))).rejects.toThrow(/not a PDF/i);
      await expect(extractPdfText(undefined)).rejects.toThrow(/not a PDF/i);
      await expect(extractPdfText('%PDF-1.4 as a string')).rejects.toThrow(/not a PDF/i);
    });

    it('throws a clear error for a corrupt PDF', async () => {
      vi.spyOn(console, 'error').mockImplementation(() => {}); // pdfjs warnings
      await expect(extractPdfText(Buffer.from('%PDF-1.4\ngarbage'))).rejects.toThrow(/Cannot extract text from PDF/);
    });

    it('keeps pdfjs warnings off stdout (the MCP transport) and restores console.log', async () => {
      const logSpy = vi.spyOn(console, 'log').mockImplementation(() => {});
      const errSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
      await extractPdfText(fixture('acme-invoice-uk.pdf'));
      await expect(extractPdfText(Buffer.from('%PDF-1.4\ngarbage'))).rejects.toThrow();
      expect(logSpy).not.toHaveBeenCalled();
      expect(errSpy.mock.calls.some(args => String(args[0]).startsWith('Warning'))).toBe(true);
      expect(console.log).toBe(logSpy);
    });

    it('gives up with a clear error instead of hanging when the parser never settles', async () => {
      const officeParser = (await import('officeparser')).default;
      vi.spyOn(officeParser, 'parseOfficeAsync').mockReturnValue(new Promise(() => {}));
      const logBefore = console.log;
      await expect(extractPdfText(fixture('acme-invoice-uk.pdf'), { timeoutMs: 20 }))
        .rejects.toThrow(/Cannot extract text from PDF: timed out after 20ms/);
      expect(console.log).toBe(logBefore); // restored even on failure
    });

    it('works under the Claude Desktop (Electron utilityProcess) environment without server/index.js loading the shim first', () => {
      const out = execFileSync(process.execPath, ['--input-type=module', '-e', `
        Object.defineProperty(process.versions, 'electron', { value: '42.5.1', configurable: true });
        process.type = 'utility';
        const { extractPdfText } = await import('./server/tools/receipts/receiptPdfText.js');
        const fs = await import('fs');
        const text = await extractPdfText(fs.readFileSync('server/tests/fixtures/receipts/acme-invoice-uk.pdf'));
        console.log(text.includes('ZQXV9876-0042') ? 'PDF_TEXT_OK' : 'PDF_TEXT_MISSING');
      `], { cwd: repoRoot, encoding: 'utf8', timeout: 25000, stdio: ['ignore', 'pipe', 'pipe'] });
      expect(out).toContain('PDF_TEXT_OK');
    });
  });

  describe('parseReceiptPdf', () => {
    it('parses a table-layout invoice (label and value on one line)', async () => {
      const r = await parseReceiptPdf(fixture('acme-invoice-uk.pdf'));
      expect(r.amounts).toEqual({
        subtotal: 72, tax: 14.4, credits: null, total: 86.4, amountPaid: 86.4, currency: 'GBP', source: 'pdf'
      });
      expect(r.lineItems).toEqual([
        { description: 'Studio plan - 3x', quantity: 1, amount: 60 }, // description and values on separate lines
        { description: 'Gadget plan - Lite', quantity: 2, amount: 12 } // everything on one line
      ]);
      expect(r.productLabel).toBe('Studio plan - 3x; Gadget plan - Lite');
      expect(r.invoiceNumber).toBe('ZQXV9876-0042');
      expect(r.receiptNumber).toBeNull();
    });

    it('parses a stacked-layout receipt (label above value) with a credit and zero-amount usage lines', async () => {
      const r = await parseReceiptPdf(fixture('globex-receipt-usage.pdf'));
      expect(r.amounts).toEqual({
        subtotal: 34.5, tax: null, credits: 12.5, total: 22, amountPaid: 22, currency: 'USD', source: 'pdf'
      });
      expect(r.lineItems).toEqual([
        { description: 'Archive storage (per GB)', quantity: 0, amount: 0 },
        { description: 'Build time (per runner / min)', quantity: 12345, amount: 8.5 },
        { description: 'Data transfer', quantity: 98000, amount: 4 },
        { description: 'Starter plan', quantity: 1, amount: 22 }
      ]);
      expect(r.productLabel).toBe('Archive storage (per GB); Build time (per runner / min); Data transfer; Starter plan');
      expect(r.receiptNumber).toBe('5555-0000-4321');
      expect(r.invoiceNumber).toBeNull();
    });

    it('lets configured product-label patterns win, like the email parser', async () => {
      const r = await parseReceiptPdf(fixture('acme-invoice-uk.pdf'), { productLabels: [/Studio plan[^\n]*/i] });
      expect(r.productLabel).toBe('Studio plan - 3x');
      const none = await parseReceiptPdf(fixture('acme-invoice-uk.pdf'), { productLabels: [/no such product/i] });
      expect(none.productLabel).toBe('Studio plan - 3x; Gadget plan - Lite'); // falls back to line items
    });

    it('propagates the non-PDF error', async () => {
      await expect(parseReceiptPdf(Buffer.from('plain text'))).rejects.toThrow(/not a PDF/i);
    });
  });
});
