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
 * transferred), maxPages, maxChars, verbosity, electron, processType }. Posts
 * { type: 'log', message } for the first MAX_RELAYED_LOGS pdfjs console messages
 * (pdfjs logs warnings with console.log, and a worker's stdout would otherwise reach
 * the MCP transport) plus one "N further ... suppressed" line, then exactly one
 * { type: 'result', text } or { type: 'error', message }.
 *
 * Hostile PDFs can make pdfjs warn millions of times; each relayed message costs the
 * main thread a postMessage plus a stderr write, which starved its timers and grew
 * its memory. So warnings are off by default (pdfjs verbosity 0 = errors only), and
 * the cap is enforced here, before anything is posted.
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
export const MAX_RELAYED_LOGS = 10;
const MAX_LOG_MESSAGE_CHARS = 500;

/**
 * Forwards at most `limit` console messages through `post`; later ones are only
 * counted, and flush() posts a single summary line for them.
 *
 * @param {(message: { type: 'log', message: string }) => void} post
 * @param {number} [limit]
 */
export function createLogRelay(post, limit = MAX_RELAYED_LOGS) {
  let relayed = 0;
  let suppressed = 0;
  return {
    log(...args) {
      if (relayed >= limit) {
        suppressed++; // not even formatted: a flood must cost the worker next to nothing
        return;
      }
      relayed++;
      const message = args.map(String).join(' ');
      post({ type: 'log', message: message.length > MAX_LOG_MESSAGE_CHARS ? `${message.slice(0, MAX_LOG_MESSAGE_CHARS)}...` : message });
    },
    flush() {
      if (suppressed === 0) return;
      post({ type: 'log', message: `Warning: ${suppressed} further pdfjs warnings suppressed` });
      suppressed = 0;
    }
  };
}

/**
 * Text of the first `maxPages` pages, one line per baseline: a newline is
 * inserted whenever an item's y coordinate differs from the previous item's (the
 * layout officeparser produced) and at the start of every page. At most
 * `maxChars` characters: pages are read as a stream of text chunks, and once the
 * cap is reached the stream is cancelled mid-page and later pages are not opened,
 * so a page with millions of text runs is never materialised. The pdfjs document is
 * always destroyed, on success and on failure.
 *
 * @param {object} pdfjs the pdfjs-dist module (injected so it can be faked in tests)
 * @param {Uint8Array} data
 * @param {{ maxPages: number, maxChars?: number, verbosity?: number }} limits
 *   verbosity is pdfjs's: 0 errors only (default), 1 warnings, 5 infos
 * @returns {Promise<string>}
 */
export async function readPdfText(pdfjs, data, { maxPages, maxChars = Infinity, verbosity = 0 }) {
  const loadingTask = pdfjs.getDocument({ data, isEvalSupported: false, verbosity });
  try {
    const doc = await loadingTask.promise;
    const pageCount = Math.min(doc.numPages, maxPages);
    let text = '';
    let lastY;
    for (let n = 1; n <= pageCount && text.length < maxChars; n++) {
      const page = await doc.getPage(n);
      const reader = page.streamTextContent().getReader();
      lastY = undefined; // a new page always starts a new line
      try {
        while (text.length < maxChars) {
          const { done, value } = await reader.read();
          if (done) break;
          for (const item of value.items) {
            if (!('str' in item) || item.str === '') continue;
            if (item.transform[5] !== lastY) text += '\n';
            text += item.str;
            lastY = item.transform[5];
            if (text.length >= maxChars) break;
          }
        }
      } finally {
        // A no-op on a finished stream; otherwise it stops pdfjs mid-page. pdfjs rejects a cancel that has no
        // Error as its reason and then keeps the page task alive, which would make destroy() below wait on
        // it forever. Not awaited: the cancel settles only once the evaluator unwinds, and destroy() waits for that.
        reader.cancel(new Error('text limit reached')).catch(() => {});
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
  const relay = createLogRelay((message) => parentPort.postMessage(message));
  console.log = relay.log;
  console.info = relay.log;
  console.warn = relay.log;
  console.debug = relay.log;
  console.error = relay.log; // pdfjs's wasm glue prints with it; keep that flood bounded too
  try {
    mirrorHostEnvironment(job);
    const pdfjs = await loadPdfjs(job.electron);
    const text = await readPdfText(pdfjs, job.data, { maxPages: job.maxPages, maxChars: job.maxChars, verbosity: job.verbosity });
    relay.flush();
    parentPort.postMessage({ type: 'result', text });
  } catch (error) {
    relay.flush();
    parentPort.postMessage({ type: 'error', message: error?.message || String(error) });
  }
}

if (!isMainThread && parentPort && workerData?.task === WORKER_TASK) {
  run(workerData);
}
