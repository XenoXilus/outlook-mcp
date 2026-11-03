import { describe, it, expect, vi, afterEach } from 'vitest';
import fs from 'fs';
import path from 'path';
import { execFileSync } from 'child_process';
import { fileURLToPath } from 'url';
import { extractPdfText, parseReceiptPdf, MAX_PDF_BYTES, MAX_PDF_PAGES } from '../../tools/receipts/receiptPdfText.js';
import { makePdf, hostilePdf } from '../fixtures/receipts/makePdf.js';

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
      vi.spyOn(console, 'error').mockImplementation(() => {}); // pdfjs warnings
      const logBefore = console.log;
      await expect(extractPdfText(hostilePdf(), { timeoutMs: 20 }))
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
        subtotal: 72, tax: 14.4, credits: null, total: 86.4, totalLabel: 'amount paid', amountPaid: 86.4, currency: 'GBP', source: 'pdf'
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
        subtotal: 34.5, tax: null, credits: 12.5, total: 22, totalLabel: 'amount paid', amountPaid: 22, currency: 'USD', source: 'pdf'
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

  describe('hostile PDFs', () => {
    it('a crafted PDF cannot block the event loop: the timeout is a real wall-clock bound', async () => {
      vi.spyOn(console, 'error').mockImplementation(() => {});
      const pdf = hostilePdf();
      expect(pdf.length).toBeLessThan(200 * 1024);
      let ticks = 0;
      const ticker = setInterval(() => { ticks++; }, 25);
      const started = Date.now();
      try {
        await expect(extractPdfText(pdf, { timeoutMs: 500 }))
          .rejects.toThrow(/Cannot extract text from PDF: timed out after 500ms/);
      } finally {
        clearInterval(ticker);
      }
      expect(Date.now() - started).toBeLessThan(1500);
      expect(ticks).toBeGreaterThanOrEqual(8); // the main thread kept running while the PDF was parsed
    });

    it('rejects a PDF over the byte cap without parsing it', async () => {
      expect(MAX_PDF_BYTES).toBe(15 * 1024 * 1024);
      const huge = Buffer.concat([Buffer.from('%PDF-1.4\n'), Buffer.alloc(MAX_PDF_BYTES)]);
      const started = Date.now();
      await expect(extractPdfText(huge)).rejects.toThrow(/Cannot extract text from PDF: .*15 MB/);
      expect(Date.now() - started).toBeLessThan(250);
    });

    it('reads only the first 10 pages', async () => {
      expect(MAX_PDF_PAGES).toBe(10);
      const pages = Array.from({ length: 12 }, (_, i) => [`Page marker ${i + 1}`]);
      const lines = (await extractPdfText(makePdf({ pages }))).split('\n');
      expect(lines).toContain('Page marker 1');
      expect(lines).toContain('Page marker 10');
      expect(lines).not.toContain('Page marker 11');
      expect(lines).not.toContain('Page marker 12');
    });

    it('caps invoice/receipt numbers and labels parsed from PDF text', async () => {
      vi.spyOn(console, 'error').mockImplementation(() => {});
      const pdf = makePdf({ pages: [[
        `Invoice number ZQ1${'X'.repeat(50000)}`,
        `Receipt number 77${'Y'.repeat(50000)}`,
        `Studio plan ${'Z'.repeat(5000)}`,
        'Total $5.00'
      ]] });
      const r = await parseReceiptPdf(pdf, { productLabels: [/Studio plan[^\n]*/] });
      expect(r.invoiceNumber.startsWith('ZQ1XXX')).toBe(true);
      expect(r.invoiceNumber.length).toBeLessThanOrEqual(301);
      expect(r.receiptNumber.length).toBeLessThanOrEqual(301);
      expect(r.productLabel.length).toBeLessThanOrEqual(301);
      expect(r.amounts).toMatchObject({ total: 5, totalLabel: 'total', source: 'pdf' });
    });

    it('reports totalLabel null when the PDF total is only a guess', async () => {
      const r = await parseReceiptPdf(makePdf({ pages: [['Thanks!', 'Charged $9.00 today']] }));
      expect(r.amounts).toMatchObject({ total: 9, totalLabel: null, source: 'pdf' });
    });

    it('readPdfText parses at most maxPages pages and always destroys the document', async () => {
      const { readPdfText } = await import('../../tools/receipts/receiptPdfWorker.js');
      const fakePdfjs = (failOnPage) => {
        const calls = { pages: [], destroyed: 0 };
        const pdfjs = {
          getDocument: () => ({
            promise: Promise.resolve({
              numPages: 500,
              getPage: async (n) => {
                calls.pages.push(n);
                if (n === failOnPage) throw new Error('bad page');
                return {
                  getTextContent: async () => ({ items: [{ str: `p${n}`, transform: [1, 0, 0, 1, 0, n] }] }),
                  cleanup: () => {}
                };
              }
            }),
            destroy: async () => { calls.destroyed++; }
          })
        };
        return { pdfjs, calls };
      };
      const ok = fakePdfjs();
      expect(await readPdfText(ok.pdfjs, new Uint8Array(1), { maxPages: 10 })).toBe('\np1\np2\np3\np4\np5\np6\np7\np8\np9\np10');
      expect(ok.calls.pages).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
      expect(ok.calls.destroyed).toBe(1);

      const failing = fakePdfjs(3);
      await expect(readPdfText(failing.pdfjs, new Uint8Array(1), { maxPages: 10 })).rejects.toThrow('bad page');
      expect(failing.calls.destroyed).toBe(1);
    });

    it('enforces the timeout under the Claude Desktop (Electron utilityProcess) environment too', () => {
      const out = execFileSync(process.execPath, ['--input-type=module', '-e', `
        Object.defineProperty(process.versions, 'electron', { value: '42.5.1', configurable: true });
        process.type = 'utility';
        console.error = () => {};
        const { extractPdfText } = await import('./server/tools/receipts/receiptPdfText.js');
        const { hostilePdf, makePdf } = await import('./server/tests/fixtures/receipts/makePdf.js');
        const ok = await extractPdfText(makePdf({ pages: [['Invoice number ZZ-0001']] }));
        let ticks = 0;
        const ticker = setInterval(() => { ticks++; }, 25);
        const started = Date.now();
        const outcome = await extractPdfText(hostilePdf(), { timeoutMs: 500 }).then(() => 'RESOLVED', (e) => e.message);
        clearInterval(ticker);
        console.log(JSON.stringify({ ok: ok.includes('ZZ-0001'), outcome, elapsed: Date.now() - started, ticks }));
      `], { cwd: repoRoot, encoding: 'utf8', timeout: 25000, stdio: ['ignore', 'pipe', 'pipe'] });
      const result = JSON.parse(out.trim().split('\n').pop());
      expect(result.ok).toBe(true);
      expect(result.outcome).toMatch(/timed out after 500ms/);
      expect(result.elapsed).toBeLessThan(1500);
      expect(result.ticks).toBeGreaterThanOrEqual(8);
    });
  });
});
