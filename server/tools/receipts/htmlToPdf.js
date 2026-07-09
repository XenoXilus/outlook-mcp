/**
 * Headless HTML→PDF rendering for the FR-7 fallback.
 *
 * Uses the system Chrome/Chromium binary (`--headless --print-to-pdf`) so no
 * heavyweight npm dependency is needed. The email body is sanitised with
 * DOMPurify before it ever reaches the renderer.
 */

import fs from 'fs';
import os from 'os';
import path from 'path';
import { spawn } from 'child_process';
import DOMPurify from 'isomorphic-dompurify';
import { isPdfBuffer } from '../../utils/receiptFiles.js';

// Strip remote URLs from attributes that trigger network fetches during render.
// Registered once at module load; applies to every sanitize() call on this instance.
DOMPurify.addHook('afterSanitizeAttributes', (node) => {
  // Remove src/poster/background if they reference a remote URL (http/https/protocol-relative).
  for (const attr of ['src', 'poster', 'background']) {
    const val = node.getAttribute?.(attr);
    if (val && /^https?:|^\/\//i.test(val.trim())) {
      node.removeAttribute(attr);
    }
  }
  // Remove srcset if ANY comma-separated token references a remote URL —
  // a mixed "local.png 1x, https://remote/img.png 2x" would bypass a start-anchored check.
  const srcset = node.getAttribute?.('srcset');
  if (srcset) {
    const hasRemote = srcset.split(',').some((token) => /https?:|\/\//i.test(token.trim()));
    if (hasRemote) node.removeAttribute('srcset');
  }
  // Drop any inline style that contains url() — safest blanket rule for remote CSS references.
  const style = node.getAttribute?.('style');
  if (style && /url\s*\(/i.test(style)) {
    node.removeAttribute('style');
  }
});

const CHROME_CANDIDATES = {
  darwin: [
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    '/Applications/Chromium.app/Contents/MacOS/Chromium',
    '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge'
  ],
  linux: [
    '/usr/bin/google-chrome',
    '/usr/bin/google-chrome-stable',
    '/usr/bin/chromium-browser',
    '/usr/bin/chromium'
  ],
  win32: [
    'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe'
  ]
};

export function sanitizeEmailHtml(html) {
  return DOMPurify.sanitize(html || '', {
    FORBID_TAGS: ['script', 'iframe', 'object', 'embed', 'form', 'link'],
    FORBID_ATTR: ['onerror', 'onload', 'onclick']
  });
}

export function buildRenderDocument({ subject, fromAddress, receivedDateTime, bodyHtml }) {
  const esc = (s) => String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body { font-family: -apple-system, 'Segoe UI', Arial, sans-serif; margin: 24px; }
  .audit-header { border-bottom: 1px solid #ccc; padding-bottom: 12px; margin-bottom: 16px; font-size: 12px; color: #444; }
  .audit-header h1 { font-size: 16px; color: #000; margin: 0 0 6px 0; }
</style>
</head>
<body>
<div class="audit-header">
  <h1>${esc(subject)}</h1>
  <div>From: ${esc(fromAddress)}</div>
  <div>Received: ${esc(receivedDateTime)}</div>
  <div>Rendered from email (fallback — not the formal vendor invoice)</div>
</div>
${sanitizeEmailHtml(bodyHtml)}
</body>
</html>`;
}

export function findChromeBinary() {
  const fromEnv = (process.env.MCP_OUTLOOK_CHROME_PATH || '').trim();
  if (fromEnv) return fromEnv;
  const candidates = CHROME_CANDIDATES[process.platform] || CHROME_CANDIDATES.linux;
  return candidates.find(p => fs.existsSync(p)) || null;
}

/**
 * Returns the Chrome CLI args for headless PDF rendering.
 * Exported so tests can assert on the arg list without spawning a process.
 *
 * @param {string} pdfPath - Absolute path to the output PDF file.
 * @param {string} htmlPath - Absolute path to the input HTML file.
 * @returns {string[]}
 */
export function buildChromeArgs(pdfPath, htmlPath) {
  return [
    '--headless',
    '--disable-gpu',
    '--no-first-run',
    '--no-default-browser-check',
    '--disable-extensions',
    '--blink-settings=imagesEnabled=false',
    '--disable-remote-fonts',
    `--print-to-pdf=${pdfPath}`,
    '--no-pdf-header-footer',
    `file://${htmlPath}`
  ];
}

export async function renderHtmlToPdfBuffer(html, options = {}) {
  const { timeoutMs = 60000, chromePath = findChromeBinary() } = options;
  if (!chromePath) {
    throw new Error('No Chrome/Chromium binary found for HTML->PDF rendering; set MCP_OUTLOOK_CHROME_PATH');
  }

  const tmpDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'outlook-render-'));
  const htmlPath = path.join(tmpDir, 'email.html');
  const pdfPath = path.join(tmpDir, 'email.pdf');

  try {
    await fs.promises.writeFile(htmlPath, html, 'utf8');

    await new Promise((resolve, reject) => {
      const child = spawn(chromePath, buildChromeArgs(pdfPath, htmlPath), { stdio: 'ignore' });

      const timer = setTimeout(() => {
        child.kill('SIGKILL');
        reject(new Error(`Chrome PDF render timed out after ${timeoutMs}ms`));
      }, timeoutMs);

      child.on('error', (error) => { clearTimeout(timer); reject(error); });
      child.on('exit', (code) => {
        clearTimeout(timer);
        if (code === 0) resolve();
        else reject(new Error(`Chrome exited with code ${code} during PDF render`));
      });
    });

    const buffer = await fs.promises.readFile(pdfPath);
    if (!isPdfBuffer(buffer)) throw new Error('Rendered output is not a valid PDF');
    return buffer;
  } finally {
    await fs.promises.rm(tmpDir, { recursive: true, force: true });
  }
}
