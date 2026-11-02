import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { downloadAttachmentTool } from '../../tools/attachments/downloadAttachment.js';
import { mockGraphErrorResult } from '../helpers/mockGraph.js';

/**
 * Regression tests for large-attachment delivery: with includeContent the
 * inline base64 payload blew past the MCP tool-output token cap for any
 * attachment above roughly 25 KB, and there was no way to have the server
 * write the bytes to disk instead.
 */

const PDF_BYTES = Buffer.from('%PDF-1.4 real invoice bytes for download test');
const PDF_B64 = PDF_BYTES.toString('base64');

// Mimics Graph: a $select narrows the returned fields (contentBytes only when
// asked for), no $select returns the whole fileAttachment, and '@odata.type' is
// not a selectable property (400).
function fileAttachmentMock({ contentBytes = PDF_B64, name = 'Invoice-001.pdf', contentType = 'application/pdf', id = 'a1' } = {}) {
  const expectedPath = `/me/messages/m1/attachments/${id}`;
  const meta = {
    '@odata.type': '#microsoft.graph.fileAttachment',
    id,
    name,
    contentType,
    size: Buffer.from(contentBytes, 'base64').length,
    isInline: false,
    lastModifiedDateTime: '2026-08-29T15:48:33Z'
  };
  return vi.fn((reqPath, options = {}) => {
    if (reqPath !== expectedPath) {
      return Promise.reject(new Error(`unexpected path: ${reqPath}`));
    }
    if (options.select) {
      if (options.select.includes('@odata.type')) {
        return Promise.resolve(mockGraphErrorResult({
          message: "Parsing OData Select and Expand failed: Term '@odata.type' is not valid in a $select or $expand expression."
        }));
      }
      return Promise.resolve(options.select.includes('contentBytes') ? { ...meta, contentBytes } : meta);
    }
    return Promise.resolve({ ...meta, contentBytes });
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

  let getBinary;
  function auth(mock, binary = vi.fn()) {
    makeRequest = mock;
    getBinary = binary;
    return {
      ensureAuthenticated: vi.fn().mockResolvedValue(true),
      getGraphApiClient: vi.fn().mockReturnValue({ makeRequest: mock, getBinary: binary })
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

  it('fetches the bytes without a $select and never selects @odata.type', async () => {
    authManager = auth(fileAttachmentMock());

    const res = await downloadAttachmentTool(authManager, {
      messageId: 'm1', attachmentId: 'a1', saveToFile: true
    });
    expect(res.isError).toBeUndefined();
    expect(makeRequest).toHaveBeenCalledTimes(1);
    expect(makeRequest.mock.calls[0][1]?.select).toBeUndefined();
    expect(getBinary).not.toHaveBeenCalled();
  });

  it('metadata-only and inline requests never put @odata.type in a $select', async () => {
    authManager = auth(fileAttachmentMock());

    await downloadAttachmentTool(authManager, { messageId: 'm1', attachmentId: 'a1' });
    await downloadAttachmentTool(authManager, { messageId: 'm1', attachmentId: 'a1', includeContent: true, decodeContent: false });
    expect(makeRequest.mock.calls.length).toBeGreaterThanOrEqual(3);
    for (const call of makeRequest.mock.calls) {
      expect(call[1]?.select || '').not.toContain('@odata.type');
    }
  });

  it.each([
    ['a dash', 'AbCd-12Ef-34'],
    ['an underscore mid-string', 'AbCd_12Ef_34'],
    ['a trailing =', 'AbCd12Ef34gH='],
    ['dash, underscore and trailing =', 'Ab-Cd_12Ef=']
  ])('saves a real PDF for an attachment id with %s, passing the id through unchanged', async (_label, id) => {
    authManager = auth(fileAttachmentMock({ id }));

    const res = await downloadAttachmentTool(authManager, {
      messageId: 'm1', attachmentId: id, saveToFile: true
    });
    expect(res.isError).toBeUndefined();
    const out = JSON.parse(res.content[0].text);
    const onDisk = fs.readFileSync(out.savedPath);
    expect(onDisk.length).toBeGreaterThan(0);
    expect(onDisk.subarray(0, 4).toString('latin1')).toBe('%PDF');
    expect(onDisk).toEqual(PDF_BYTES);
    expect(makeRequest.mock.calls.map(c => c[0])).toEqual([`/me/messages/m1/attachments/${id}`]);
  });

  it('falls back to /$value when the item endpoint fails, with identical bytes', async () => {
    // Item route rejected (e.g. 400); the raw route delivers the same file.
    const binary = vi.fn(async () => Buffer.from(PDF_BYTES));
    const mock = vi.fn((reqPath, options = {}) => {
      if (reqPath !== '/me/messages/m1/attachments/a1') return Promise.reject(new Error(`unexpected path: ${reqPath}`));
      if (options.select) {
        return Promise.resolve({
          '@odata.type': '#microsoft.graph.fileAttachment', id: 'a1', name: 'Invoice-001.pdf',
          contentType: 'application/pdf', size: PDF_BYTES.length, isInline: false
        });
      }
      return Promise.resolve(mockGraphErrorResult({ message: 'The attachment could not be returned' }));
    });
    authManager = auth(mock, binary);

    const res = await downloadAttachmentTool(authManager, {
      messageId: 'm1', attachmentId: 'a1', saveToFile: true
    });
    expect(res.isError).toBeUndefined();
    const out = JSON.parse(res.content[0].text);
    expect(binary).toHaveBeenCalledWith('/me/messages/m1/attachments/a1/$value');
    expect(fs.readFileSync(out.savedPath)).toEqual(PDF_BYTES);
    expect(out.name).toBe('Invoice-001.pdf');
  });

  it('includeContent also uses the fallback and keeps the bytes intact', async () => {
    const binary = vi.fn(async () => Buffer.from(PDF_BYTES));
    const mock = vi.fn((reqPath, options = {}) => {
      if (options.select) {
        return Promise.resolve({
          '@odata.type': '#microsoft.graph.fileAttachment', id: 'a1', name: 'Invoice-001.pdf',
          contentType: 'application/pdf', size: PDF_BYTES.length, isInline: false
        });
      }
      return Promise.resolve(mockGraphErrorResult({ status: 404, code: 'ErrorItemNotFound', message: 'not found' }));
    });
    authManager = auth(mock, binary);

    const res = await downloadAttachmentTool(authManager, {
      messageId: 'm1', attachmentId: 'a1', includeContent: true, decodeContent: false
    });
    const out = JSON.parse(res.content[0].text);
    expect(out.contentIncluded).toBe(true);
    expect(Buffer.from(out.contentBytes, 'base64')).toEqual(PDF_BYTES);
  });

  it('reports both routes verbatim, with fetchDetails, when neither delivers the bytes', async () => {
    const binary = vi.fn(async () => mockGraphErrorResult({
      status: 404, code: 'ErrorItemNotFound', message: 'The specified object was not found in the store.',
      requestId: 'req-value-route', method: 'GET', url: 'https://graph.microsoft.com/v1.0/me/messages/m1/attachments/a1/$value'
    }));
    const mock = vi.fn(() => Promise.resolve(mockGraphErrorResult({
      status: 400, code: 'BadRequest', message: 'Unable to read the attachment', requestId: 'req-json-route'
    })));
    authManager = auth(mock, binary);

    const res = await downloadAttachmentTool(authManager, {
      messageId: 'm1', attachmentId: 'a1', saveToFile: true
    });
    expect(res.isError).toBe(true);
    const text = res.content[0].text;
    expect(text).toContain('HTTP 400 BadRequest: Unable to read the attachment');
    expect(text).toContain('request-id: req-json-route');
    expect(text).toContain('HTTP 404 ErrorItemNotFound: The specified object was not found in the store.');
    expect(text).toContain('request-id: req-value-route');
    expect(res._errorDetails.fetchDetails).toMatchObject({
      attemptedRoutes: ['json', 'value'], messageId: 'm1', attachmentId: 'a1'
    });
    expect(res._errorDetails.graphError).toMatchObject({ status: 404, code: 'ErrorItemNotFound' });
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

  it('saveToFile routes through the per-call mailbox', async () => {
    const mock = fileAttachmentMock();
    // fileAttachmentMock pins the /me path; accept the shared path instead
    const shared = vi.fn((p, o) => mock(p.replace('/users/careers%40example.com', '/me'), o));
    authManager = auth(shared);
    const res = await downloadAttachmentTool(authManager, {
      messageId: 'm1', attachmentId: 'a1', saveToFile: true, mailbox: 'careers@example.com'
    });
    expect(res.isError).toBeUndefined();
    expect(shared.mock.calls.every(c => c[0].startsWith('/users/careers%40example.com/'))).toBe(true);
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
    expect(res._errorDetails.fetchDetails).toMatchObject({ messageId: 'm1', attachmentId: 'a1', attachmentName: 'embedded message' });
    // A non-file attachment is final: the raw-bytes route is not tried
    expect(getBinary).not.toHaveBeenCalled();
  });
});
