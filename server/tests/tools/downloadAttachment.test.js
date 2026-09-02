import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { downloadAttachmentTool } from '../../tools/attachments/downloadAttachment.js';

/**
 * Regression tests for large-attachment delivery: with includeContent the
 * inline base64 payload blew past the MCP tool-output token cap for any
 * attachment above roughly 25 KB, and there was no way to have the server
 * write the bytes to disk instead.
 */

const PDF_BYTES = Buffer.from('%PDF-1.4 real invoice bytes for download test');
const PDF_B64 = PDF_BYTES.toString('base64');

function fileAttachmentMock({ contentBytes = PDF_B64, name = 'Invoice-001.pdf', contentType = 'application/pdf' } = {}) {
  const meta = {
    '@odata.type': '#microsoft.graph.fileAttachment',
    id: 'a1',
    name,
    contentType,
    size: Buffer.from(contentBytes, 'base64').length,
    isInline: false,
    lastModifiedDateTime: '2026-08-29T15:48:33Z'
  };
  return vi.fn((reqPath, options = {}) => {
    if (reqPath !== '/me/messages/m1/attachments/a1') {
      return Promise.reject(new Error(`unexpected path: ${reqPath}`));
    }
    if ((options.select || '').includes('contentBytes')) {
      return Promise.resolve({ ...meta, contentBytes });
    }
    return Promise.resolve(meta);
  });
}

describe('downloadAttachmentTool (file delivery)', () => {
  let workDir, receiptsDir, savedWork, savedReceipts, authManager, makeRequest;

  beforeEach(() => {
    workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'dl-work-'));
    receiptsDir = fs.mkdtempSync(path.join(os.tmpdir(), 'dl-receipts-'));
    savedWork = process.env.MCP_OUTLOOK_WORK_DIR;
    savedReceipts = process.env.MCP_OUTLOOK_RECEIPTS_DIR;
    process.env.MCP_OUTLOOK_WORK_DIR = workDir;
    process.env.MCP_OUTLOOK_RECEIPTS_DIR = receiptsDir;
  });

  afterEach(() => {
    if (savedWork === undefined) delete process.env.MCP_OUTLOOK_WORK_DIR;
    else process.env.MCP_OUTLOOK_WORK_DIR = savedWork;
    if (savedReceipts === undefined) delete process.env.MCP_OUTLOOK_RECEIPTS_DIR;
    else process.env.MCP_OUTLOOK_RECEIPTS_DIR = savedReceipts;
    fs.rmSync(workDir, { recursive: true, force: true });
    fs.rmSync(receiptsDir, { recursive: true, force: true });
  });

  function auth(mock) {
    makeRequest = mock;
    return {
      ensureAuthenticated: vi.fn().mockResolvedValue(true),
      getGraphApiClient: vi.fn().mockReturnValue({ makeRequest: mock })
    };
  }

  it('saveToFile writes the raw bytes to disk and returns metadata only', async () => {
    authManager = auth(fileAttachmentMock());

    const res = await downloadAttachmentTool(authManager, {
      messageId: 'm1', attachmentId: 'a1', saveToFile: true
    });
    expect(res.isError).toBeUndefined();
    const out = JSON.parse(res.content[0].text);

    expect(out.savedToFile).toBe(true);
    expect(out.savedPath).toBe(path.join(receiptsDir, 'Invoice-001.pdf'));
    expect(fs.readFileSync(out.savedPath)).toEqual(PDF_BYTES);
    expect(out.sha256).toMatch(/^[0-9a-f]{64}$/);
    expect(out.size).toBe(PDF_BYTES.length);
    expect(out.contentBytes).toBeUndefined();
    expect(res.content[0].text).not.toContain(PDF_B64);
  });

  it('saveToFile honours fileName and onExisting', async () => {
    authManager = auth(fileAttachmentMock());

    const args = {
      messageId: 'm1', attachmentId: 'a1',
      saveToFile: true, fileName: 'Acme 29Aug26 Invoice.pdf'
    };
    const first = JSON.parse((await downloadAttachmentTool(authManager, args)).content[0].text);
    expect(path.basename(first.savedPath)).toBe('Acme 29Aug26 Invoice.pdf');
    expect(first.action).toBe('saved');

    const second = JSON.parse((await downloadAttachmentTool(authManager, args)).content[0].text);
    expect(second.action).toBe('skipped'); // default onExisting: skip
  });

  it('a destDir alone implies saving to file', async () => {
    authManager = auth(fileAttachmentMock());

    const res = await downloadAttachmentTool(authManager, {
      messageId: 'm1', attachmentId: 'a1', destDir: receiptsDir
    });
    const out = JSON.parse(res.content[0].text);
    expect(out.savedToFile).toBe(true);
    expect(out.savedPath).toBe(path.join(receiptsDir, 'Invoice-001.pdf'));
  });

  it('fetches content with the proven $select shape', async () => {
    authManager = auth(fileAttachmentMock());

    await downloadAttachmentTool(authManager, {
      messageId: 'm1', attachmentId: 'a1', saveToFile: true
    });
    const contentFetch = makeRequest.mock.calls.find(c => (c[1]?.select || '').includes('contentBytes'));
    expect(contentFetch[1].select).toBe('id,name,contentType,size,isInline,lastModifiedDateTime,contentBytes,@odata.type');
  });

  it('still returns small content inline when saveToFile is not requested', async () => {
    authManager = auth(fileAttachmentMock());

    const res = await downloadAttachmentTool(authManager, {
      messageId: 'm1', attachmentId: 'a1', includeContent: true, decodeContent: false
    });
    const out = JSON.parse(res.content[0].text);
    expect(out.contentBytes).toBe(PDF_B64);
    expect(out.savedToFile).toBeUndefined();
  });

  it('auto-spills to a file instead of exceeding the MCP response cap', async () => {
    // ~45 KB binary → ~60 KB base64: over the ~25k-token MCP output cap while
    // staying far under the old, never-triggering 1 MB limit.
    const bigBytes = Buffer.alloc(45000, 7);
    const bigB64 = bigBytes.toString('base64');
    authManager = auth(fileAttachmentMock({ contentBytes: bigB64, name: 'big.bin', contentType: 'application/octet-stream' }));

    const res = await downloadAttachmentTool(authManager, {
      messageId: 'm1', attachmentId: 'a1', includeContent: true, decodeContent: false
    });
    expect(res.isError).toBeUndefined();
    const text = res.content[0].text;
    expect(text.length).toBeLessThan(35000);
    const out = JSON.parse(text);
    expect(out.contentBytes).toBeUndefined();
    expect(out.contentSavedToFile).toBe(true);
    expect(fs.existsSync(out.fileOutput.filePath)).toBe(true);
    expect(fs.readFileSync(out.fileOutput.filePath)).toEqual(bigBytes);
  });

  it('saveToFile rejects non-file attachments with a clear error', async () => {
    const mock = vi.fn(() => Promise.resolve({
      '@odata.type': '#microsoft.graph.itemAttachment',
      id: 'a1', name: 'embedded message', size: 100, isInline: false
    }));
    authManager = auth(mock);

    const res = await downloadAttachmentTool(authManager, {
      messageId: 'm1', attachmentId: 'a1', saveToFile: true
    });
    expect(res.isError).toBe(true);
    expect(res.content[0].text).toMatch(/file attachment/i);
  });
});
