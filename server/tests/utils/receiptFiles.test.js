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
  saveReceiptFile,
  existingVersionPaths
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

  it('onExisting=skip reports identical bytes as skipped without existingDiffers', async () => {
    await saveReceiptFile(PDF, receiptsDir, 'x.pdf');
    const res = await saveReceiptFile(Buffer.from(PDF), receiptsDir, 'x.pdf');
    expect(res.action).toBe('skipped');
    expect(res.existingDiffers).toBeUndefined();
  });

  it('onExisting=skip never skips onto a file holding different bytes: it versions and flags existingDiffers', async () => {
    await saveReceiptFile(PDF, receiptsDir, 'x.pdf');
    const other = Buffer.from('%PDF-1.7 a different receipt');

    const res = await saveReceiptFile(other, receiptsDir, 'x.pdf');
    expect(res).toMatchObject({ action: 'versioned', existingDiffers: true });
    expect(path.basename(res.savedPath)).toBe('x (2).pdf');
    expect(fs.readFileSync(res.savedPath).equals(other)).toBe(true);
    expect(fs.readFileSync(path.join(receiptsDir, 'x.pdf')).equals(PDF)).toBe(true); // untouched

    // A re-run with the same bytes finds the earlier version instead of writing another.
    const again = await saveReceiptFile(other, receiptsDir, 'x.pdf');
    expect(again).toMatchObject({ action: 'skipped', existingDiffers: true, savedPath: res.savedPath });
    expect(fs.readdirSync(receiptsDir).sort()).toEqual(['x (2).pdf', 'x.pdf']);
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

  it('lists a name and its existing numbered versions, stopping at the first gap', async () => {
    expect(existingVersionPaths(receiptsDir, 'x.pdf')).toEqual([]);
    fs.writeFileSync(path.join(receiptsDir, 'x.pdf'), PDF);
    fs.writeFileSync(path.join(receiptsDir, 'x (2).pdf'), PDF);
    fs.writeFileSync(path.join(receiptsDir, 'x (4).pdf'), PDF);
    expect(existingVersionPaths(receiptsDir, 'x.pdf')).toEqual([
      path.join(receiptsDir, 'x.pdf'), path.join(receiptsDir, 'x (2).pdf')
    ]);
  });

  it('leaves no temp files behind', async () => {
    await saveReceiptFile(PDF, receiptsDir, 'x.pdf');
    expect(fs.readdirSync(receiptsDir).filter(f => f.startsWith('.tmp-'))).toHaveLength(0);
  });

  describe('PDF creation/modification timestamps', () => {
    const stamped = (stamp, body = 'same receipt body') => Buffer.from(
      `%PDF-1.4\n1 0 obj\n<< /Producer (Renderer) /CreationDate (D:20260601120000${stamp}+00'00') ` +
      `/ModDate (D:20260601120000${stamp}+00'00') >>\nendobj\n${body}\n%%EOF\n`, 'latin1'
    );

    it('onExisting=skip treats a PDF differing only in CreationDate/ModDate as the same file', async () => {
      const first = await saveReceiptFile(stamped('1'), receiptsDir, 'x.pdf');
      const again = await saveReceiptFile(stamped('2'), receiptsDir, 'x.pdf');

      expect(again).toMatchObject({ action: 'skipped', savedPath: first.savedPath, sha256: first.sha256, size: first.size });
      expect(again.existingDiffers).toBeUndefined();
      expect(fs.readdirSync(receiptsDir)).toEqual(['x.pdf']);
      expect(fs.readFileSync(first.savedPath).equals(stamped('1'))).toBe(true); // the earlier file is untouched
    });

    it('still versions a PDF whose content differs, whatever its timestamps', async () => {
      await saveReceiptFile(stamped('1'), receiptsDir, 'x.pdf');
      const other = await saveReceiptFile(stamped('2', 'another receipt body'), receiptsDir, 'x.pdf');
      const sameStamp = await saveReceiptFile(stamped('1', 'a third receipt body'), receiptsDir, 'x.pdf');

      expect(other).toMatchObject({ action: 'versioned', existingDiffers: true });
      expect(sameStamp).toMatchObject({ action: 'versioned', existingDiffers: true });
      expect(fs.readdirSync(receiptsDir).sort()).toEqual(['x (2).pdf', 'x (3).pdf', 'x.pdf']);
    });

    it('compares files that are not PDFs byte for byte', async () => {
      const text = (stamp) => Buffer.from(`/CreationDate (D:2026060112000${stamp})\n`);
      await saveReceiptFile(text('1'), receiptsDir, 'x.txt');
      const res = await saveReceiptFile(text('2'), receiptsDir, 'x.txt');
      expect(res).toMatchObject({ action: 'versioned', existingDiffers: true });
    });
  });

  describe('concurrent saves to one name', () => {
    it('never reports two different files as saved at the same path', async () => {
      const a = Buffer.from('%PDF-1.4 receipt A');
      const b = Buffer.from('%PDF-1.4 receipt B');
      const results = await Promise.all([
        saveReceiptFile(a, receiptsDir, 'x.pdf'),
        saveReceiptFile(b, receiptsDir, 'x.pdf')
      ]);

      expect(new Set(results.map(r => r.savedPath)).size).toBe(2);
      expect(fs.readFileSync(results[0].savedPath).equals(a)).toBe(true);
      expect(fs.readFileSync(results[1].savedPath).equals(b)).toBe(true);
      expect(results.map(r => r.action).sort()).toEqual(['saved', 'versioned']);
      expect(fs.readdirSync(receiptsDir).sort()).toEqual(['x (2).pdf', 'x.pdf']);
    });

    it('writes identical bytes once and reports the other save as skipped', async () => {
      const results = await Promise.all([
        saveReceiptFile(PDF, receiptsDir, 'x.pdf'),
        saveReceiptFile(Buffer.from(PDF), receiptsDir, 'x.pdf')
      ]);

      expect(results.map(r => r.action).sort()).toEqual(['saved', 'skipped']);
      expect(results[0].savedPath).toBe(results[1].savedPath);
      expect(fs.readdirSync(receiptsDir)).toEqual(['x.pdf']);
    });

    it('applies onExisting=version to a name taken by a concurrent save', async () => {
      const results = await Promise.all([
        saveReceiptFile(PDF, receiptsDir, 'x.pdf', { onExisting: 'version' }),
        saveReceiptFile(PDF, receiptsDir, 'x.pdf', { onExisting: 'version' })
      ]);
      expect(results.map(r => path.basename(r.savedPath)).sort()).toEqual(['x (2).pdf', 'x.pdf']);
      expect(fs.readdirSync(receiptsDir).filter(f => f.startsWith('.tmp-'))).toHaveLength(0);
    });
  });
});
