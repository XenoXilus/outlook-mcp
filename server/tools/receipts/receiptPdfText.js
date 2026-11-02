/**
 * Text extraction and totals parsing for receipt/invoice PDFs.
 *
 * Uses officeparser (which wraps pdfjs-dist) — the same dependency the
 * attachment tools use — so no new package is needed. Text is rebuilt from
 * pdfjs text runs, one line per baseline, so a label and its value normally
 * share a line ("Total £72.00") but may land on consecutive lines; the shared
 * totals parser (receiptParser.js) accepts both.
 *
 * Claude Desktop (Electron utilityProcess) caveats. officeparser loads pdfjs at
 * import time, which crashes there unless utils/electronCompat.js ran first;
 * server/index.js already loads the shim at startup, and this module also imports
 * it itself, lazily, so extraction works however the module was reached (the shim
 * only defines DOMMatrix when the global is missing). pdfjs also classifies that
 * process as a browser and then insists on a Web Worker it does not have — the
 * parse throws inside officeparser and the promise never settles. So under
 * Electron the pdfjs worker module is preloaded into globalThis.pdfjsWorker, which
 * makes pdfjs run its worker in-process, and every extraction has a timeout.
 */

import { isPdfBuffer } from '../../utils/receiptFiles.js';
import {
  parseReceiptTotals,
  extractReceiptNumber,
  extractInvoiceNumber,
  extractProductLabel,
  buildProductLabel
} from './receiptParser.js';

let officeParserPromise = null;

async function loadPdfEngine() {
  await import('../../utils/electronCompat.js');
  const officeParser = (await import('officeparser')).default;
  if (process.versions?.electron && !globalThis.pdfjsWorker) {
    try {
      globalThis.pdfjsWorker = await import('pdfjs-dist/legacy/build/pdf.worker.mjs');
    } catch (error) {
      console.error(`Warning: could not preload the pdfjs worker (${error.message}); PDF parsing may fail in this host`);
    }
  }
  return officeParser;
}

function loadOfficeParser() {
  if (!officeParserPromise) {
    officeParserPromise = loadPdfEngine().catch((error) => {
      officeParserPromise = null; // allow a retry rather than caching the failure
      throw error;
    });
  }
  return officeParserPromise;
}

// pdfjs reports warnings with console.log, i.e. stdout — which is the MCP
// stdio transport. Send them to stderr for the duration of a parse.
let logRedirects = 0;
let originalLog = null;

function routeLogToStderr() {
  if (logRedirects++ === 0) {
    originalLog = console.log;
    console.log = (...args) => console.error(...args);
  }
  return () => {
    if (--logRedirects === 0) {
      console.log = originalLog;
      originalLog = null;
    }
  };
}

const DEFAULT_TIMEOUT_MS = 30000;

/**
 * Extract the plain text of a PDF.
 * Image-only (scanned) PDFs yield an empty string; non-PDF input throws.
 *
 * @param {Buffer|Uint8Array} buffer
 * @param {{ timeoutMs?: number }} [options]
 * @returns {Promise<string>}
 */
export async function extractPdfText(buffer, { timeoutMs = DEFAULT_TIMEOUT_MS } = {}) {
  const bytes = buffer instanceof Uint8Array && !Buffer.isBuffer(buffer) ? Buffer.from(buffer) : buffer;
  if (!isPdfBuffer(bytes)) {
    throw new Error('Cannot extract text: input is not a PDF (missing %PDF header)');
  }
  const officeParser = await loadOfficeParser();
  const restoreLog = routeLogToStderr();
  let timer;
  try {
    const timeout = new Promise((_, reject) => {
      timer = setTimeout(() => reject(new Error(`timed out after ${timeoutMs}ms`)), timeoutMs);
    });
    const text = await Promise.race([officeParser.parseOfficeAsync(bytes), timeout]);
    return typeof text === 'string' ? text : '';
  } catch (error) {
    throw new Error(`Cannot extract text from PDF: ${error?.message || error}`);
  } finally {
    clearTimeout(timer);
    restoreLog();
  }
}

/**
 * Parse a receipt/invoice PDF into the same totals shape the email-body parser
 * returns (amounts.source is 'pdf').
 *
 * @param {Buffer|Uint8Array} buffer
 * @param {{ productLabels?: RegExp[] }} [rules] optional config product-label patterns
 */
export async function parseReceiptPdf(buffer, rules = {}) {
  const text = await extractPdfText(buffer);
  const { amounts, lineItems } = parseReceiptTotals(text, 'pdf');
  const configured = extractProductLabel(text, rules.productLabels || []);
  return {
    amounts,
    lineItems,
    productLabel: configured ? configured.slice(0, 300) : buildProductLabel(lineItems),
    invoiceNumber: extractInvoiceNumber(text),
    receiptNumber: extractReceiptNumber(text)
  };
}
