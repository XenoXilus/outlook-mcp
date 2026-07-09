/**
 * Receipt file utilities (FR-4, NFR-3, NFR-5)
 *
 * Deterministic vendor filenames (no timestamp suffixes), collision policy,
 * filename sanitisation, write containment to the receipts/work directories,
 * and atomic saves (temp file + rename).
 */

import fs from 'fs';
import path from 'path';
import crypto from 'crypto';
import { getWorkDirectory } from './fileOutput.js';

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

export function getReceiptsDirectory() {
  const dir = (process.env.MCP_OUTLOOK_RECEIPTS_DIR || '').trim();
  if (!dir) return getWorkDirectory();
  try {
    fs.mkdirSync(dir, { recursive: true });
    return dir;
  } catch (error) {
    console.error(`Warning: cannot create receipts directory ${dir} (${error.message}); falling back to work directory`);
    return getWorkDirectory();
  }
}

export function getFilenameTemplate() {
  return process.env.RECEIPT_FILENAME_TEMPLATE || '{vendor} {DDMmmYY} Invoice.pdf';
}

export function formatDDMmmYY(dateLike) {
  const d = new Date(dateLike);
  if (isNaN(d.getTime())) throw new Error(`Invalid date for filename: ${dateLike}`);
  const dd = String(d.getUTCDate()).padStart(2, '0');
  const yy = String(d.getUTCFullYear() % 100).padStart(2, '0');
  return `${dd}${MONTHS[d.getUTCMonth()]}${yy}`;
}

export function renderFilenameTemplate(template, { vendor, date }) {
  return template
    .replaceAll('{vendor}', vendor)
    .replaceAll('{DDMmmYY}', formatDDMmmYY(date));
}

export function sanitiseFilename(name) {
  const cleaned = String(name || '')
    .replace(/[/\\:*?"<>|\x00-\x1f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!cleaned || cleaned === '.' || cleaned === '..') {
    throw new Error(`Invalid filename: ${JSON.stringify(name)}`);
  }
  return cleaned;
}

/**
 * Extra write roots granted by configuration (comma-separated). This only
 * widens where an explicit destDir may point — it never changes where files
 * go by default, so scratch downloads stay in the work directory.
 */
export function getAllowedWriteDirs() {
  return (process.env.MCP_OUTLOOK_ALLOWED_WRITE_DIRS || '')
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);
}

export function resolveDestination(destDir, fileName) {
  const defaultRoots = [path.resolve(getReceiptsDirectory()), path.resolve(getWorkDirectory())];
  const roots = [...defaultRoots, ...getAllowedWriteDirs().map(d => path.resolve(d))];
  const dir = destDir && String(destDir).trim() ? path.resolve(destDir) : defaultRoots[0];
  const contained = roots.some(root => dir === root || dir.startsWith(root + path.sep));
  if (!contained) {
    throw new Error(
      `destDir must be inside the receipts directory (${roots[0]}), work directory (${roots[1]})` +
      ` or an allowed write directory (${getAllowedWriteDirs().join(', ') || 'none configured'}), got: ${dir}`
    );
  }
  return { dir, filePath: path.join(dir, sanitiseFilename(fileName)) };
}

export function sha256Hex(buffer) {
  return crypto.createHash('sha256').update(buffer).digest('hex');
}

export function isPdfBuffer(buffer) {
  return Buffer.isBuffer(buffer) && buffer.length >= 4 && buffer.subarray(0, 4).toString('latin1') === '%PDF';
}

function nextVersionedPath(filePath) {
  const ext = path.extname(filePath);
  const base = ext ? filePath.slice(0, -ext.length) : filePath;
  for (let i = 2; ; i++) {
    const candidate = `${base} (${i})${ext}`;
    if (!fs.existsSync(candidate)) return candidate;
  }
}

export async function saveReceiptFile(buffer, destDir, fileName, options = {}) {
  const { onExisting = 'skip' } = options;
  if (!['skip', 'overwrite', 'version'].includes(onExisting)) {
    throw new Error(`Invalid onExisting policy: ${onExisting} (expected skip|overwrite|version)`);
  }
  const { dir, filePath } = resolveDestination(destDir, fileName);
  if (filePath.toLowerCase().endsWith('.pdf') && !isPdfBuffer(buffer)) {
    throw new Error('Content is not a valid PDF (missing %PDF magic bytes)');
  }

  await fs.promises.mkdir(dir, { recursive: true });

  let target = filePath;
  let action = 'saved';
  if (fs.existsSync(filePath)) {
    if (onExisting === 'skip') {
      const existing = await fs.promises.readFile(filePath);
      return { savedPath: filePath, action: 'skipped', sha256: sha256Hex(existing), size: existing.length };
    }
    if (onExisting === 'version') {
      target = nextVersionedPath(filePath);
      action = 'versioned';
    } else {
      action = 'overwritten';
    }
  }

  const tmp = path.join(dir, `.tmp-${crypto.randomBytes(6).toString('hex')}`);
  try {
    await fs.promises.writeFile(tmp, buffer);
    await fs.promises.rename(tmp, target);
  } catch (error) {
    await fs.promises.rm(tmp, { force: true });
    throw error;
  }

  return { savedPath: target, action, sha256: sha256Hex(buffer), size: buffer.length };
}
