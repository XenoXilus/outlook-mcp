/**
 * Worker-thread half of receiptPdfText.js: parses one PDF with pdfjs-dist and
 * posts its text back to the main thread.
 *
 * Why a worker: pdfjs runs its own "worker" in-process here (Node, and Claude
 * Desktop's Electron utilityProcess), driven entirely by promise microtasks, so
 * a crafted PDF can keep the main thread busy for seconds — no timer gets a turn
 * and a Promise.race timeout never fires. In a worker_threads Worker the main
 * thread stays free, and worker.terminate() stops the parse at once.
 *
 * Protocol: started with workerData { task: 'receipt-pdf-text', data (Uint8Array,
 * transferred), maxPages, maxChars, electron, processType }. Posts
 * { type: 'log', message } for each pdfjs console message (pdfjs logs warnings with
 * console.log, and a worker's stdout would otherwise reach the MCP transport),
 * then exactly one { type: 'result', text } or { type: 'error', message }.
 *
 * Claude Desktop: the parent passes its process.versions.electron / process.type,
 * and the worker mirrors them when its own process object lacks them, so pdfjs
 * takes the same code path as on the host's main thread. That path needs the
 * DOMMatrix shim (utils/electronCompat.js) before pdfjs is evaluated, and the pdfjs
 * worker module preloaded into globalThis.pdfjsWorker so pdfjs runs it in-thread
 * instead of looking for a Web Worker it does not have.
 */

import { parentPort, workerData, isMainThread } from 'node:worker_threads';

export const WORKER_TASK = 'receipt-pdf-text';

/**
 * Text of the first `maxPages` pages, one line per baseline: a newline is
 * inserted whenever an item's y coordinate differs from the previous item's (the
 * layout officeparser produced) and at the start of every page. At most
 * `maxChars` characters. The pdfjs document is always destroyed, on success and
 * on failure.
 *
 * @param {object} pdfjs the pdfjs-dist module (injected so it can be faked in tests)
 * @param {Uint8Array} data
 * @param {{ maxPages: number, maxChars?: number }} limits
 * @returns {Promise<string>}
 */
export async function readPdfText(pdfjs, data, { maxPages, maxChars = Infinity }) {
  const loadingTask = pdfjs.getDocument({ data, isEvalSupported: false });
  try {
    const doc = await loadingTask.promise;
    const pageCount = Math.min(doc.numPages, maxPages);
    let text = '';
    let lastY;
    for (let n = 1; n <= pageCount && text.length < maxChars; n++) {
      const page = await doc.getPage(n);
      const content = await page.getTextContent();
      lastY = undefined; // a new page always starts a new line
      for (const item of content.items) {
        if (!('str' in item) || item.str === '') continue;
        if (item.transform[5] !== lastY) text += '\n';
        text += item.str;
        lastY = item.transform[5];
        if (text.length >= maxChars) break;
      }
      page.cleanup();
    }
    return text.length > maxChars ? text.slice(0, maxChars) : text;
  } finally {
    await loadingTask.destroy();
  }
}

function mirrorHostEnvironment({ electron, processType }) {
  if (electron && !process.versions.electron) {
    Object.defineProperty(process.versions, 'electron', { value: electron, configurable: true, enumerable: true });
  }
  if (electron && processType && !process.type) process.type = processType;
}

async function loadPdfjs(electron) {
  await import('../../utils/electronCompat.js'); // before pdfjs: it may evaluate `new DOMMatrix()` at import
  if (electron && !globalThis.pdfjsWorker) {
    try {
      globalThis.pdfjsWorker = await import('pdfjs-dist/legacy/build/pdf.worker.mjs');
    } catch (error) {
      console.log(`Warning: could not preload the pdfjs worker (${error.message}); PDF parsing may fail in this host`);
    }
  }
  return import('pdfjs-dist/legacy/build/pdf.mjs');
}

async function run(job) {
  const forward = (...args) => parentPort.postMessage({ type: 'log', message: args.map(String).join(' ') });
  console.log = forward;
  console.info = forward;
  console.warn = forward;
  console.debug = forward;
  try {
    mirrorHostEnvironment(job);
    const pdfjs = await loadPdfjs(job.electron);
    const text = await readPdfText(pdfjs, job.data, { maxPages: job.maxPages, maxChars: job.maxChars });
    parentPort.postMessage({ type: 'result', text });
  } catch (error) {
    parentPort.postMessage({ type: 'error', message: error?.message || String(error) });
  }
}

if (!isMainThread && parentPort && workerData?.task === WORKER_TASK) {
  run(workerData);
}
