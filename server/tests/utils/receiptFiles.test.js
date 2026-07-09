import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import {
  getReceiptsDirectory,
  formatDDMmmYY,
  renderFilenameTemplate,
  sanitiseFilename,
  resolveDestination,
  isPdfBuffer,
  saveReceiptFile
} from '../../utils/receiptFiles.js';

const PDF = Buffer.from('%PDF-1.4 test content');

describe('receiptFiles', () => {
  let receiptsDir;
  const savedEnv = {};

  beforeEach(() => {
    receiptsDir = fs.mkdtempSync(path.join(os.tmpdir(), 'receipts-'));
    savedEnv.MCP_OUTLOOK_RECEIPTS_DIR = process.env.MCP_OUTLOOK_RECEIPTS_DIR;
    savedEnv.MCP_OUTLOOK_WORK_DIR = process.env.MCP_OUTLOOK_WORK_DIR;
    process.env.MCP_OUTLOOK_RECEIPTS_DIR = receiptsDir;
  });

  afterEach(() => {
    for (const [k, v] of Object.entries(savedEnv)) {
      if (v === undefined) delete process.env[k]; else process.env[k] = v;
    }
    fs.rmSync(receiptsDir, { recursive: true, force: true });
  });

  it('uses MCP_OUTLOOK_RECEIPTS_DIR when set', () => {
    expect(getReceiptsDirectory()).toBe(receiptsDir);
  });

  it('formats DDMmmYY from an ISO date', () => {
    expect(formatDDMmmYY('2026-06-29T07:12:00Z')).toBe('29Jun26');
    expect(formatDDMmmYY('2026-06-01T00:00:00Z')).toBe('01Jun26');
  });

  it('renders the default filename template', () => {
    const name = renderFilenameTemplate('{vendor} {DDMmmYY} Invoice.pdf', {
      vendor: 'Acme',
      date: '2026-06-29T07:12:00Z'
    });
    expect(name).toBe('Acme 29Jun26 Invoice.pdf');
  });

  it('sanitises path separators and control characters', () => {
    expect(sanitiseFilename('a/b\\c:d*e.pdf')).toBe('a b c d e.pdf');
    expect(() => sanitiseFilename('..')).toThrow();
    expect(() => sanitiseFilename('   ')).toThrow();
  });

  it('refuses destinations outside the receipts/work directories', () => {
    expect(() => resolveDestination('/etc', 'x.pdf')).toThrow(/must be inside/);
    const ok = resolveDestination(receiptsDir, 'x.pdf');
    expect(ok.filePath).toBe(path.join(receiptsDir, 'x.pdf'));
  });

  it('MCP_OUTLOOK_ALLOWED_WRITE_DIRS grants extra write roots without moving defaults', () => {
    // Simulate the DXT: work dir is scratch, no receipts env, and the OneDrive
    // area is a DISJOINT root granted only via the allowed-write-dirs setting.
    const scratchDir = fs.mkdtempSync(path.join(os.tmpdir(), 'scratch-'));
    const extraRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'areas-'));
    const saved = process.env.MCP_OUTLOOK_ALLOWED_WRITE_DIRS;
    try {
      delete process.env.MCP_OUTLOOK_RECEIPTS_DIR;
      process.env.MCP_OUTLOOK_WORK_DIR = scratchDir; // work root ≠ os.tmpdir()

      // without the grant, the disjoint root is refused
      expect(() => resolveDestination(extraRoot, 'x.pdf')).toThrow(/must be inside/);

      process.env.MCP_OUTLOOK_ALLOWED_WRITE_DIRS = extraRoot;

      // a subdir of the allowed root is now writable
      const inside = path.join(extraRoot, '1 Receipts');
      expect(resolveDestination(inside, 'x.pdf').filePath).toBe(path.join(inside, 'x.pdf'));

      // outside all roots stays refused
      expect(() => resolveDestination('/etc', 'x.pdf')).toThrow(/must be inside/);

      // defaults are NOT redirected: omitting destDir still resolves to the
      // receipts/work default (scratch), not the allowed extra root
      const def = resolveDestination(undefined, 'x.pdf');
      expect(def.dir).toBe(path.resolve(scratchDir));
    } finally {
      if (saved === undefined) delete process.env.MCP_OUTLOOK_ALLOWED_WRITE_DIRS;
      else process.env.MCP_OUTLOOK_ALLOWED_WRITE_DIRS = saved;
      fs.rmSync(extraRoot, { recursive: true, force: true });
      fs.rmSync(scratchDir, { recursive: true, force: true });
    }
  });

  it('detects PDF magic bytes', () => {
    expect(isPdfBuffer(PDF)).toBe(true);
    expect(isPdfBuffer(Buffer.from('<html>'))).toBe(false);
  });

  it('saves bytes exactly, returns sha256 and action=saved', async () => {
    const res = await saveReceiptFile(PDF, receiptsDir, 'Acme 29Jun26 Invoice.pdf');
    expect(res.action).toBe('saved');
    expect(fs.readFileSync(res.savedPath).equals(PDF)).toBe(true);
    expect(res.sha256).toMatch(/^[0-9a-f]{64}$/);
    expect(res.size).toBe(PDF.length);
  });

  it('rejects non-PDF bytes for .pdf filenames', async () => {
    await expect(saveReceiptFile(Buffer.from('<html>'), receiptsDir, 'x.pdf'))
      .rejects.toThrow(/%PDF/);
  });

  it('onExisting=skip is idempotent', async () => {
    await saveReceiptFile(PDF, receiptsDir, 'x.pdf');
    const res = await saveReceiptFile(PDF, receiptsDir, 'x.pdf');
    expect(res.action).toBe('skipped');
    expect(fs.readdirSync(receiptsDir)).toHaveLength(1);
  });

  it('onExisting=overwrite replaces the file', async () => {
    await saveReceiptFile(PDF, receiptsDir, 'x.pdf');
    const other = Buffer.from('%PDF-1.7 other');
    const res = await saveReceiptFile(other, receiptsDir, 'x.pdf', { onExisting: 'overwrite' });
    expect(res.action).toBe('overwritten');
    expect(fs.readFileSync(res.savedPath).equals(other)).toBe(true);
  });

  it('onExisting=version writes "name (2).pdf"', async () => {
    await saveReceiptFile(PDF, receiptsDir, 'x.pdf');
    const res = await saveReceiptFile(PDF, receiptsDir, 'x.pdf', { onExisting: 'version' });
    expect(res.action).toBe('versioned');
    expect(path.basename(res.savedPath)).toBe('x (2).pdf');
  });

  it('leaves no temp files behind', async () => {
    await saveReceiptFile(PDF, receiptsDir, 'x.pdf');
    expect(fs.readdirSync(receiptsDir).filter(f => f.startsWith('.tmp-'))).toHaveLength(0);
  });
});
