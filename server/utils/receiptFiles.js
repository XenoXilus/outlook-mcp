/**
 * Receipt file utilities (FR-4, NFR-3, NFR-5)
 *
 * Deterministic vendor filenames (no timestamp suffixes), collision policy,
 * filename sanitisation, write containment to the receipts/work directories,
 * and atomic saves (temp file, then an exclusive link or a rename).
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

function versionedPath(filePath, n) {
  const ext = path.extname(filePath);
  const base = ext ? filePath.slice(0, -ext.length) : filePath;
  return `${base} (${n})${ext}`;
}

function nextVersionedPath(filePath) {
  for (let i = 2; ; i++) {
    const candidate = versionedPath(filePath, i);
    if (!fs.existsSync(candidate)) return candidate;
  }
}

/** The file itself and its "name (2)", "name (3)"... versions that exist, up to the first gap. */
function existingVersionsOf(filePath) {
  const found = [];
  for (let i = 1; ; i++) {
    const candidate = i === 1 ? filePath : versionedPath(filePath, i);
    if (!fs.existsSync(candidate)) return found;
    found.push(candidate);
  }
}

/** Paths a save of fileName into destDir would compare against under onExisting 'skip'. */
export function existingVersionPaths(destDir, fileName) {
  return existingVersionsOf(resolveDestination(destDir, fileName).filePath);
}

// Info-dictionary dates (D:YYYYMMDDHHmmSS plus an optional zone) that a PDF
// renderer stamps on every run.
const PDF_TIMESTAMPS = /\/(CreationDate|ModDate)\s*\(D:[0-9Z+\-' ]{0,32}\)/g;

function withoutPdfTimestamps(buffer) {
  return buffer.toString('latin1').replace(PDF_TIMESTAMPS, '/$1()');
}

/**
 * True when `existing` holds the same document as `buffer`: identical bytes,
 * or two PDFs that differ only in their CreationDate/ModDate stamps (rendering
 * one email twice never gives identical bytes). Any other difference, or a
 * file that is not a PDF, is a different document.
 */
function sameDocument(buffer, existing, sha256) {
  if (sha256Hex(existing) === sha256) return true;
  if (!isPdfBuffer(buffer) || !isPdfBuffer(existing)) return false;
  return withoutPdfTimestamps(existing) === withoutPdfTimestamps(buffer);
}

/**
 * The first of `candidates` (existing file paths) that holds the same document
 * as `buffer`, as { path, sha256, size } of that file's own bytes, or null.
 */
export async function findSameDocument(buffer, candidates) {
  const sha256 = sha256Hex(buffer);
  for (const candidate of candidates) {
    let existing;
    try {
      existing = await fs.promises.readFile(candidate);
    } catch {
      continue; // removed meanwhile: nothing to match
    }
    if (sameDocument(buffer, existing, sha256)) {
      return { path: candidate, sha256: sha256Hex(existing), size: existing.length };
    }
  }
  return null;
}

/** Where a save goes under the policy, or { result } when an existing file already holds the document. */
async function planSave(buffer, filePath, onExisting) {
  if (!fs.existsSync(filePath)) return { target: filePath, action: 'saved' };
  if (onExisting === 'skip') {
    const match = await findSameDocument(buffer, existingVersionsOf(filePath));
    if (match) {
      const result = { savedPath: match.path, action: 'skipped', sha256: match.sha256, size: match.size };
      if (match.path !== filePath) result.existingDiffers = true;
      return { result };
    }
    return { target: nextVersionedPath(filePath), action: 'versioned', existingDiffers: true };
  }
  if (onExisting === 'version') return { target: nextVersionedPath(filePath), action: 'versioned' };
  return { target: filePath, action: 'overwritten' };
}

/**
 * Give the finished temp file the target name only if no file has it yet
 * (EEXIST otherwise). A hard link is atomic; where the file system has none,
 * an exclusive copy is used.
 */
async function createExclusive(tmp, target) {
  try {
    await fs.promises.link(tmp, target);
  } catch (error) {
    if (error.code === 'EEXIST') throw error;
    await fs.promises.copyFile(tmp, target, fs.constants.COPYFILE_EXCL);
  }
}

// Concurrent saves that keep taking the chosen name before this one do not retry forever.
const MAX_CREATE_ATTEMPTS = 20;

/**
 * Write bytes under the collision policy, never replacing a file another save
 * created meanwhile (new names are created exclusively; a name taken in the
 * meantime is re-decided under the same policy):
 * - skip: an existing file is kept only when it holds the same document
 *   (identical bytes, or a PDF differing only in its CreationDate/ModDate
 *   stamps; the result then carries that file's own sha256 and size); a file
 *   holding another document is never reported in its place, so the bytes go
 *   to the first numbered version that already holds them, else to a new one,
 *   and the result carries existingDiffers: true.
 * - version: always a new "name (n)" file when the name is taken.
 * - overwrite: replace the file (atomically, temp file + rename).
 */
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

  const tmp = path.join(dir, `.tmp-${crypto.randomBytes(6).toString('hex')}`);
  let tmpWritten = false;
  try {
    for (let attempt = 1; ; attempt++) {
      const plan = await planSave(buffer, filePath, onExisting);
      if (plan.result) return plan.result;
      if (!tmpWritten) {
        await fs.promises.writeFile(tmp, buffer);
        tmpWritten = true;
      }
      try {
        if (plan.action === 'overwritten') await fs.promises.rename(tmp, plan.target);
        else await createExclusive(tmp, plan.target);
      } catch (error) {
        if (error.code === 'EEXIST' && attempt < MAX_CREATE_ATTEMPTS) continue;
        throw error;
      }
      const result = { savedPath: plan.target, action: plan.action, sha256: sha256Hex(buffer), size: buffer.length };
      if (plan.existingDiffers) result.existingDiffers = true;
      return result;
    }
  } finally {
    if (tmpWritten) await fs.promises.rm(tmp, { force: true });
  }
}
