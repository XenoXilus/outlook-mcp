/**
 * Text extraction and totals parsing for receipt/invoice PDFs.
 *
 * Uses pdfjs-dist (already installed with officeparser) inside a worker thread
 * (receiptPdfWorker.js). Text is rebuilt from pdfjs text runs, one line per
 * baseline, so a label and its value normally share a line ("Total £72.00") but
 * may land on consecutive lines; the shared totals parser (receiptParser.js)
 * accepts both.
 *
 * PDFs are untrusted input, so every extraction is bounded:
 *   - bytes: input over MAX_PDF_BYTES is rejected before parsing;
 *   - pages: only the first MAX_PDF_PAGES pages are read;
 *   - time: the parse runs in a worker_threads Worker that is terminate()d when
 *     timeoutMs elapses. pdfjs runs in-process and is driven by microtasks, so an
 *     in-thread timer could not interrupt it — a crafted PDF used to block the
 *     event loop for many seconds;
 *   - memory: the worker has its own heap limit, and the pdfjs document is
 *     destroyed when done;
 *   - output: at most MAX_TEXT_CHARS characters come back.
 *
 * Claude Desktop (Electron utilityProcess): the worker mirrors the host's
 * Electron markers, loads the DOMMatrix shim (utils/electronCompat.js) and
 * preloads the pdfjs worker module, as the in-process engine did — see
 * receiptPdfWorker.js. pdfjs console output is relayed to stderr, never stdout
 * (the MCP stdio transport).
 */

import { Worker } from 'node:worker_threads';
import { isPdfBuffer } from '../../utils/receiptFiles.js';
import {
  parseReceiptTotals,
  extractReceiptNumber,
  extractInvoiceNumber,
  extractProductLabel,
  buildProductLabel,
  capStr
} from './receiptParser.js';
import { WORKER_TASK } from './receiptPdfWorker.js';

export const MAX_PDF_BYTES = 15 * 1024 * 1024;
export const MAX_PDF_PAGES = 10;
export const MAX_TEXT_CHARS = 1024 * 1024;
const WORKER_HEAP_MB = 256;
const DEFAULT_TIMEOUT_MS = 30000;
const WORKER_URL = new URL('./receiptPdfWorker.js', import.meta.url);

const toMb = (bytes) => Math.round(bytes / (1024 * 1024) * 10) / 10;

/** Run one extraction in a fresh worker; resolves with its text, rejects on error, exit or timeout. */
function extractInWorker(bytes, timeoutMs) {
  const data = new Uint8Array(bytes); // a private copy: its buffer is transferred to the worker
  return new Promise((resolve, reject) => {
    let worker = null;
    let settled = false;
    const finish = (error, text) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      if (worker) worker.terminate().catch(() => {});
      if (error) reject(error);
      else resolve(text);
    };
    const timer = setTimeout(() => finish(new Error(`timed out after ${timeoutMs}ms`)), timeoutMs);

    try {
      worker = new Worker(WORKER_URL, {
        workerData: {
          task: WORKER_TASK,
          data,
          maxPages: MAX_PDF_PAGES,
          maxChars: MAX_TEXT_CHARS,
          electron: process.versions?.electron || null,
          processType: process.type || null
        },
        transferList: [data.buffer],
        // Do not inherit the host's CLI flags: entry-point options such as --input-type or
        // --eval are invalid for a worker started from a file, and the worker needs none.
        execArgv: [],
        resourceLimits: { maxOldGenerationSizeMb: WORKER_HEAP_MB },
        stdout: true // keep anything the worker prints off our stdout (the MCP transport)
      });
    } catch (error) {
      finish(error);
      return;
    }
    worker.stdout.on('data', (chunk) => process.stderr.write(chunk));
    worker.on('message', (message) => {
      if (message?.type === 'log') console.error(message.message);
      else if (message?.type === 'result') finish(null, typeof message.text === 'string' ? message.text : '');
      else if (message?.type === 'error') finish(new Error(message.message));
    });
    worker.on('error', (error) => {
      finish(error?.code === 'ERR_WORKER_OUT_OF_MEMORY' ? new Error('the parser ran out of memory') : error);
    });
    worker.on('exit', (code) => finish(new Error(`the parser exited unexpectedly (code ${code})`)));
  });
}

/**
 * Extract the plain text of a PDF (first MAX_PDF_PAGES pages).
 * Image-only (scanned) PDFs yield an empty string; non-PDF input throws, as do
 * PDFs over MAX_PDF_BYTES and parses that outlive timeoutMs.
 *
 * @param {Buffer|Uint8Array} buffer
 * @param {{ timeoutMs?: number }} [options]
 * @returns {Promise<string>}
 */
export async function extractPdfText(buffer, { timeoutMs = DEFAULT_TIMEOUT_MS } = {}) {
  const bytes = buffer instanceof Uint8Array && !Buffer.isBuffer(buffer) ? Buffer.from(buffer) : buffer;
  if (!isPdfBuffer(bytes)) {
    throw new Error('Cannot extract text from PDF: input is not a PDF (missing %PDF header)');
  }
  if (bytes.length > MAX_PDF_BYTES) {
    throw new Error(`Cannot extract text from PDF: the file is ${toMb(bytes.length)} MB, over the ${toMb(MAX_PDF_BYTES)} MB limit`);
  }
  try {
    return await extractInWorker(bytes, timeoutMs);
  } catch (error) {
    throw new Error(`Cannot extract text from PDF: ${error?.message || error}`);
  }
}

/**
 * Parse a receipt/invoice PDF into the same totals shape the email-body parser
 * returns (amounts.source is 'pdf'). Every string field is length-capped.
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
    productLabel: capStr(configured || buildProductLabel(lineItems)),
    invoiceNumber: capStr(extractInvoiceNumber(text)),
    receiptNumber: capStr(extractReceiptNumber(text))
  };
}
