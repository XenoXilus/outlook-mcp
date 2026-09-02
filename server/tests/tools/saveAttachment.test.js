import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { saveAttachmentTool } from '../../tools/receipts/saveAttachment.js';

const PDF_B64 = Buffer.from('%PDF-1.4 real invoice bytes').toString('base64');

describe('saveAttachmentTool (FR-1)', () => {
  let receiptsDir, makeRequest, authManager, savedEnv;

  beforeEach(() => {
    receiptsDir = fs.mkdtempSync(path.join(os.tmpdir(), 'receipts-'));
    savedEnv = process.env.MCP_OUTLOOK_RECEIPTS_DIR;
    process.env.MCP_OUTLOOK_RECEIPTS_DIR = receiptsDir;
    makeRequest = vi.fn();
    authManager = {
      ensureAuthenticated: vi.fn().mockResolvedValue(true),
      getGraphApiClient: vi.fn().mockReturnValue({ makeRequest })
    };
  });

  afterEach(() => {
    if (savedEnv === undefined) delete process.env.MCP_OUTLOOK_RECEIPTS_DIR;
    else process.env.MCP_OUTLOOK_RECEIPTS_DIR = savedEnv;
    fs.rmSync(receiptsDir, { recursive: true, force: true });
  });

  it('requires messageId', async () => {
    const res = await saveAttachmentTool(authManager, {});
    expect(res.isError).toBe(true);
    expect(res.content[0].text).toContain('messageId');
  });

  it('auto-selects the Invoice-*.pdf and saves raw bytes with the given fileName', async () => {
    makeRequest
      .mockResolvedValueOnce({ value: [
        { id: 'a1', name: 'Receipt-2060.pdf', contentType: 'application/pdf', isInline: false },
        { id: 'a2', name: 'Invoice-TSTINV42-0001.pdf', contentType: 'application/pdf', isInline: false }
      ]})
      .mockResolvedValueOnce({ id: 'a2', name: 'Invoice-TSTINV42-0001.pdf', contentType: 'application/pdf', isInline: false, contentBytes: PDF_B64 });

    const res = await saveAttachmentTool(authManager, {
      messageId: 'm1', fileName: 'Acme 29Jun26 Invoice.pdf'
    });
    const out = JSON.parse(res.content[0].text);
    expect(out.attachmentName).toBe('Invoice-TSTINV42-0001.pdf');
    expect(out.action).toBe('saved');
    expect(out.savedPath).toBe(path.join(receiptsDir, 'Acme 29Jun26 Invoice.pdf'));
    expect(fs.readFileSync(out.savedPath).toString('base64')).toBe(PDF_B64);
    expect(out.sha256).toMatch(/^[0-9a-f]{64}$/);
    // messages listing hit the mailbox attachments endpoint
    expect(makeRequest.mock.calls[0][0]).toBe('/me/messages/m1/attachments');
  });

  it('errors clearly when no PDF attachment exists', async () => {
    makeRequest.mockResolvedValueOnce({ value: [{ id: 'a1', name: 'x.png', contentType: 'image/png', isInline: false }] });
    const res = await saveAttachmentTool(authManager, { messageId: 'm1', fileName: 'x.pdf' });
    expect(res.isError).toBe(true);
    expect(res.content[0].text).toMatch(/No suitable PDF attachment/i);
  });

  it('names the contentType in the error when a contentType selector finds no match', async () => {
    makeRequest.mockResolvedValueOnce({ value: [
      { id: 'a1', name: 'invoice.pdf', contentType: 'application/pdf', isInline: false }
    ]});
    const res = await saveAttachmentTool(authManager, {
      messageId: 'm1', contentType: 'application/vnd.ms-excel', fileName: 'out.xls'
    });
    expect(res.isError).toBe(true);
    expect(res.content[0].text).toMatch(/No non-inline attachment with contentType 'application\/vnd\.ms-excel'/);
    expect(res.content[0].text).toContain('m1');
  });

  it('passes through Graph MCP errors', async () => {
    makeRequest.mockResolvedValueOnce({ content: [{ type: 'text', text: 'Resource not found' }], isError: true });
    const res = await saveAttachmentTool(authManager, { messageId: 'm1', fileName: 'x.pdf' });
    expect(res.isError).toBe(true);
    expect(res.content[0].text).toContain('Resource not found');
  });

  // --- Task 3: filenameTemplate + contentType selector ---

  it('renders default template with vendor arg + receivedDate → deterministic filename', async () => {
    makeRequest
      .mockResolvedValueOnce({ value: [
        { id: 'a1', name: 'Invoice-ABC.pdf', contentType: 'application/pdf', isInline: false }
      ]})
      // message details fetch (template path)
      .mockResolvedValueOnce({
        receivedDateTime: '2026-06-29T10:00:00Z',
        from: { emailAddress: { address: 'billing@vendor.com' } },
        subject: ''
      })
      // full attachment
      .mockResolvedValueOnce({ id: 'a1', name: 'Invoice-ABC.pdf', contentType: 'application/pdf', isInline: false, contentBytes: PDF_B64 });

    const res = await saveAttachmentTool(authManager, { messageId: 'm1', vendor: 'Acme' });
    expect(res.isError).toBeUndefined();
    const out = JSON.parse(res.content[0].text);
    expect(out.savedPath).toBe(path.join(receiptsDir, 'Acme 29Jun26 Invoice.pdf'));
    expect(out.action).toBe('saved');
  });

  it('honours an explicit filenameTemplate arg', async () => {
    makeRequest
      .mockResolvedValueOnce({ value: [
        { id: 'a1', name: 'Invoice-ABC.pdf', contentType: 'application/pdf', isInline: false }
      ]})
      .mockResolvedValueOnce({
        receivedDateTime: '2026-06-15T10:00:00Z',
        from: { emailAddress: { address: 'billing@stripe.com' } },
        subject: ''
      })
      .mockResolvedValueOnce({ id: 'a1', name: 'Invoice-ABC.pdf', contentType: 'application/pdf', isInline: false, contentBytes: PDF_B64 });

    const res = await saveAttachmentTool(authManager, {
      messageId: 'm1', vendor: 'Stripe', filenameTemplate: '{vendor}-{DDMmmYY}.pdf'
    });
    expect(res.isError).toBeUndefined();
    const out = JSON.parse(res.content[0].text);
    expect(out.savedPath).toBe(path.join(receiptsDir, 'Stripe-15Jun26.pdf'));
  });

  it('auto-detects vendor from the message when vendor arg is absent', async () => {
    makeRequest
      .mockResolvedValueOnce({ value: [
        { id: 'a1', name: 'Invoice-ABC.pdf', contentType: 'application/pdf', isInline: false }
      ]})
      .mockResolvedValueOnce({
        receivedDateTime: '2026-06-29T10:00:00Z',
        from: { emailAddress: { address: 'invoice+statements@stripe.com' } },
        subject: 'Your receipt from Acme, Inc #4242'
      })
      .mockResolvedValueOnce({ id: 'a1', name: 'Invoice-ABC.pdf', contentType: 'application/pdf', isInline: false, contentBytes: PDF_B64 });

    const res = await saveAttachmentTool(authManager, {
      messageId: 'm1', filenameTemplate: '{vendor} {DDMmmYY} Invoice.pdf'
    });
    expect(res.isError).toBeUndefined();
    const out = JSON.parse(res.content[0].text);
    expect(out.savedPath).toBe(path.join(receiptsDir, 'Acme 29Jun26 Invoice.pdf'));
  });

  it('selects attachment by contentType when no attachmentId or attachmentName given', async () => {
    makeRequest
      .mockResolvedValueOnce({ value: [
        { id: 'a1', name: 'summary.docx', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', isInline: false },
        { id: 'a2', name: 'invoice.pdf', contentType: 'application/pdf', isInline: false }
      ]})
      .mockResolvedValueOnce({ id: 'a2', name: 'invoice.pdf', contentType: 'application/pdf', isInline: false, contentBytes: PDF_B64 });

    const res = await saveAttachmentTool(authManager, {
      messageId: 'm1', contentType: 'application/pdf', fileName: 'out.pdf'
    });
    expect(res.isError).toBeUndefined();
    const out = JSON.parse(res.content[0].text);
    expect(out.attachmentId).toBe('a2');
    expect(out.attachmentName).toBe('invoice.pdf');
  });

  it('skips inline PDFs during auto-selection and picks the non-inline one', async () => {
    makeRequest
      .mockResolvedValueOnce({ value: [
        { id: 'a1', name: 'inline.pdf', contentType: 'application/pdf', isInline: true },
        { id: 'a2', name: 'Invoice-real.pdf', contentType: 'application/pdf', isInline: false }
      ]})
      .mockResolvedValueOnce({ id: 'a2', name: 'Invoice-real.pdf', contentType: 'application/pdf', isInline: false, contentBytes: PDF_B64 });

    const res = await saveAttachmentTool(authManager, { messageId: 'm1', fileName: 'out.pdf' });
    expect(res.isError).toBeUndefined();
    const out = JSON.parse(res.content[0].text);
    expect(out.attachmentId).toBe('a2');
  });

  it('prefer:receipt selects Receipt-*.pdf; prefer:first selects first non-inline PDF', async () => {
    // prefer: 'receipt'
    makeRequest
      .mockResolvedValueOnce({ value: [
        { id: 'a1', name: 'Invoice-01.pdf', contentType: 'application/pdf', isInline: false },
        { id: 'a2', name: 'Receipt-2060.pdf', contentType: 'application/pdf', isInline: false }
      ]})
      .mockResolvedValueOnce({ id: 'a2', name: 'Receipt-2060.pdf', contentType: 'application/pdf', isInline: false, contentBytes: PDF_B64 });

    const resR = await saveAttachmentTool(authManager, { messageId: 'm1', prefer: 'receipt', fileName: 'out.pdf' });
    expect(resR.isError).toBeUndefined();
    expect(JSON.parse(resR.content[0].text).attachmentId).toBe('a2');

    // prefer: 'first'
    makeRequest
      .mockResolvedValueOnce({ value: [
        { id: 'b1', name: 'Receipt-first.pdf', contentType: 'application/pdf', isInline: false },
        { id: 'b2', name: 'Invoice-second.pdf', contentType: 'application/pdf', isInline: false }
      ]})
      .mockResolvedValueOnce({ id: 'b1', name: 'Receipt-first.pdf', contentType: 'application/pdf', isInline: false, contentBytes: PDF_B64 });

    const resF = await saveAttachmentTool(authManager, { messageId: 'm1', prefer: 'first', fileName: 'out2.pdf' });
    expect(resF.isError).toBeUndefined();
    expect(JSON.parse(resF.content[0].text).attachmentId).toBe('b1');
  });

  it('contentType selector drives selection over default PDF auto-pick (isolation)', async () => {
    // Two non-inline attachments: A = PDF (would win under default/prefer logic),
    // B = xlsx (selected only when contentType explicitly targets it).
    // This test proves the selector drove the choice, not the PDF fallback.
    // Uses path-dispatching mockImplementation so a wrong-path fetch gets a1's
    // PDF data and loudly fails the assertions — positional mocks would not
    // detect that bug class.
    const PDF_B64_A1 = Buffer.from('PK mock-pdf-bytes-a1-WRONG').toString('base64');
    const XLSX_B64 = Buffer.from('PK mock-xlsx-bytes').toString('base64');
    const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    makeRequest.mockImplementation((path) => {
      if (path === '/me/messages/m1/attachments')
        return Promise.resolve({ value: [
          { id: 'a1', name: 'Invoice-x.pdf', contentType: 'application/pdf', isInline: false },
          { id: 'a2', name: 'data.xlsx', contentType: XLSX_TYPE, isInline: false }
        ]});
      if (path === '/me/messages/m1/attachments/a1')
        return Promise.resolve({
          id: 'a1', name: 'Invoice-x.pdf', contentType: 'application/pdf', isInline: false,
          contentBytes: PDF_B64_A1
        });
      if (path === '/me/messages/m1/attachments/a2')
        return Promise.resolve({
          id: 'a2', name: 'data.xlsx', contentType: XLSX_TYPE, isInline: false,
          contentBytes: XLSX_B64
        });
      return Promise.reject(new Error(`unexpected path: ${path}`));
    });

    const res = await saveAttachmentTool(authManager, {
      messageId: 'm1',
      contentType: XLSX_TYPE,
      fileName: 'out.xlsx'
    });
    expect(res.isError).toBeUndefined();
    const out = JSON.parse(res.content[0].text);
    expect(out.attachmentId).toBe('a2');
    expect(out.attachmentName).toBe('data.xlsx');
    // Fetched path must target a2, not a1
    expect(makeRequest.mock.calls[1][0]).toMatch(/\/a2$/);
    // Saved bytes must be xlsx — a1's PDF data would fail this
    expect(fs.readFileSync(out.savedPath).toString('base64')).toBe(XLSX_B64);
  });

  // Fix 2: hard-error when RECEIPT_RULES_PATH is set but file is missing
  it('returns isError when RECEIPT_RULES_PATH points to a nonexistent file', async () => {
    const savedRules = process.env.RECEIPT_RULES_PATH;
    try {
      process.env.RECEIPT_RULES_PATH = '/tmp/__nonexistent_outlook_mcp_rules_xyz__.json';
      const res = await saveAttachmentTool(authManager, { messageId: 'm1', fileName: 'x.pdf' });
      expect(res.isError).toBe(true);
      expect(res.content[0].text).toContain('RECEIPT_RULES_PATH');
      expect(makeRequest).not.toHaveBeenCalled();
    } finally {
      if (savedRules === undefined) delete process.env.RECEIPT_RULES_PATH;
      else process.env.RECEIPT_RULES_PATH = savedRules;
    }
  });

  // Graph returns 400 for the attachment fetch when $select asks for
  // contentBytes without the shape download_attachment uses. The content
  // fetch must stay byte-identical to the proven-working download request.
  it('fetches attachment content with the same $select shape download_attachment uses', async () => {
    makeRequest
      .mockResolvedValueOnce({ id: 'a9', name: 'Invoice.pdf', contentType: 'application/pdf', isInline: false, contentBytes: PDF_B64 });

    const res = await saveAttachmentTool(authManager, {
      messageId: 'm1', attachmentId: 'a9', fileName: 'out.pdf'
    });
    expect(res.isError).toBeUndefined();

    const contentFetch = makeRequest.mock.calls.find(c => c[0] === '/me/messages/m1/attachments/a9');
    expect(contentFetch).toBeDefined();
    expect(contentFetch[1].select).toBe('id,name,contentType,size,isInline,lastModifiedDateTime,contentBytes,@odata.type');
  });

  it('response includes size, sha256, and absolute savedPath', async () => {
    makeRequest
      .mockResolvedValueOnce({ value: [
        { id: 'a1', name: 'Invoice.pdf', contentType: 'application/pdf', isInline: false }
      ]})
      .mockResolvedValueOnce({ id: 'a1', name: 'Invoice.pdf', contentType: 'application/pdf', isInline: false, contentBytes: PDF_B64 });

    const res = await saveAttachmentTool(authManager, { messageId: 'm1', fileName: 'out.pdf' });
    expect(res.isError).toBeUndefined();
    const out = JSON.parse(res.content[0].text);
    expect(typeof out.size).toBe('number');
    expect(out.size).toBeGreaterThan(0);
    expect(out.sha256).toMatch(/^[0-9a-f]{64}$/);
    expect(path.isAbsolute(out.savedPath)).toBe(true);
  });
});
