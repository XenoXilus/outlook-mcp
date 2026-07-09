import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { createDraftTool } from '../../tools/email/createDraft.js';

// Must match UPLOAD_CHUNK_SIZE in createDraft.js (10 × 320 KiB).
const CHUNK = 3276800;

describe('createDraftTool with attachmentPaths (FR-8)', () => {
  let tmpDir, pdfPath, postWithRetry, makeRequest, authManager;

  beforeEach(() => {
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'draft-'));
    pdfPath = path.join(tmpDir, 'Acme 29Jun26 Invoice.pdf');
    fs.writeFileSync(pdfPath, Buffer.from('%PDF-1.4 receipt'));
    postWithRetry = vi.fn().mockImplementation((p) => {
      if (typeof p === 'string' && p.endsWith('/createUploadSession')) {
        return Promise.resolve({ uploadUrl: 'https://upload.example.com/session/abc' });
      }
      return Promise.resolve({ id: 'draft-1' });
    });
    makeRequest = vi.fn().mockResolvedValue({ id: 'draft-1', webLink: 'https://outlook.office365.com/owa/?ItemID=draft-1' });
    authManager = {
      ensureAuthenticated: vi.fn().mockResolvedValue(true),
      getGraphApiClient: vi.fn().mockReturnValue({ postWithRetry, makeRequest })
    };
  });

  afterEach(() => fs.rmSync(tmpDir, { recursive: true, force: true }));

  it('attaches local files to the draft and returns the webLink — never sends', async () => {
    const res = await createDraftTool(authManager, {
      to: ['reviewer@example.com'], cc: ['cc1@example.com', 'cc2@example.com'],
      subject: 'June invoice', body: 'Please find attached.',
      preserveUserStyling: false,
      attachmentPaths: [pdfPath]
    });
    const out = JSON.parse(res.content[0].text);
    expect(out.draftId).toBe('draft-1');
    expect(out.webLink).toContain('outlook.office365.com');
    expect(out.attachments).toEqual([{ name: 'Acme 29Jun26 Invoice.pdf', size: 16 }]);

    // draft created at /me/messages, attachment posted to the draft, nothing sent
    expect(postWithRetry.mock.calls[0][0]).toBe('/me/messages');
    const [attachPath, attachBody] = postWithRetry.mock.calls[1];
    expect(attachPath).toBe('/me/messages/draft-1/attachments');
    expect(attachBody['@odata.type']).toBe('#microsoft.graph.fileAttachment');
    expect(attachBody.name).toBe('Acme 29Jun26 Invoice.pdf');
    expect(Buffer.from(attachBody.contentBytes, 'base64').toString()).toContain('%PDF');
    expect(postWithRetry.mock.calls.some(([p]) => p.includes('/send'))).toBe(false);
  });

  it('errors on a missing attachment file', async () => {
    const res = await createDraftTool(authManager, {
      to: ['t@example.com'], subject: 's', body: 'b',
      preserveUserStyling: false,
      attachmentPaths: [path.join(tmpDir, 'nope.pdf')]
    });
    expect(res.isError).toBe(true);
    expect(res.content[0].text).toContain('nope.pdf');
  });

  it('uploads a file over 3 MB via an upload session with 320-KiB-aligned chunks — never sends', async () => {
    const bigPath = path.join(tmpDir, 'big.pdf');
    const total = CHUNK + 123200; // 3,400,000 → exactly two chunks
    fs.writeFileSync(bigPath, Buffer.alloc(total, 0x41));
    const fetchImpl = vi.fn().mockResolvedValue({ ok: true, status: 200 });

    const res = await createDraftTool(authManager, {
      to: ['t@example.com'], subject: 's', body: 'b', preserveUserStyling: false,
      attachmentPaths: [bigPath]
    }, { fetchImpl });

    const out = JSON.parse(res.content[0].text);
    expect(out.webLink).toContain('outlook.office365.com');
    expect(out.attachments).toEqual([{ name: 'big.pdf', size: total }]);

    // createUploadSession called with the correct AttachmentItem name + size
    const sessionCall = postWithRetry.mock.calls.find(([p]) => p.endsWith('/createUploadSession'));
    expect(sessionCall[0]).toBe('/me/messages/draft-1/attachments/createUploadSession');
    expect(sessionCall[1]).toEqual({ AttachmentItem: { attachmentType: 'file', name: 'big.pdf', size: total } });

    // two chunked PUTs to the pre-authenticated uploadUrl, no Authorization header
    expect(fetchImpl).toHaveBeenCalledTimes(2);
    fetchImpl.mock.calls.forEach(([url, opts]) => {
      expect(url).toBe('https://upload.example.com/session/abc');
      expect(opts.method).toBe('PUT');
      expect(opts.headers.Authorization).toBeUndefined();
    });
    // Content-Range headers cover the full byte range, chunk 1 aligned to 320 KiB
    expect(fetchImpl.mock.calls[0][1].headers['Content-Range']).toBe(`bytes 0-${CHUNK - 1}/${total}`);
    expect(fetchImpl.mock.calls[0][1].headers['Content-Length']).toBe(String(CHUNK));
    expect(fetchImpl.mock.calls[1][1].headers['Content-Range']).toBe(`bytes ${CHUNK}-${total - 1}/${total}`);
    expect(fetchImpl.mock.calls[1][1].headers['Content-Length']).toBe(String(total - CHUNK));

    // never sends, and never uses the direct fileAttachment POST for the large file
    expect(postWithRetry.mock.calls.some(([p]) => p.includes('/send'))).toBe(false);
    expect(postWithRetry.mock.calls.some(([p]) => p === '/me/messages/draft-1/attachments')).toBe(false);
  });

  it('handles a mix of a small (direct) and a large (session) attachment in one call', async () => {
    const bigPath = path.join(tmpDir, 'big.pdf');
    const total = CHUNK + 1000;
    fs.writeFileSync(bigPath, Buffer.alloc(total, 0x42));
    const fetchImpl = vi.fn().mockResolvedValue({ ok: true, status: 200 });

    const res = await createDraftTool(authManager, {
      to: ['t@example.com'], subject: 's', body: 'b', preserveUserStyling: false,
      attachmentPaths: [pdfPath, bigPath]
    }, { fetchImpl });

    const out = JSON.parse(res.content[0].text);
    expect(out.attachments).toEqual([
      { name: 'Acme 29Jun26 Invoice.pdf', size: 16 },
      { name: 'big.pdf', size: total }
    ]);

    // small → direct fileAttachment POST
    const direct = postWithRetry.mock.calls.find(([p]) => p === '/me/messages/draft-1/attachments');
    expect(direct[1]['@odata.type']).toBe('#microsoft.graph.fileAttachment');
    expect(direct[1].name).toBe('Acme 29Jun26 Invoice.pdf');

    // large → upload session + chunked PUTs
    const session = postWithRetry.mock.calls.find(([p]) => p.endsWith('/createUploadSession'));
    expect(session[1].AttachmentItem.size).toBe(total);
    expect(fetchImpl).toHaveBeenCalledTimes(2);

    expect(postWithRetry.mock.calls.some(([p]) => p.includes('/send'))).toBe(false);
  });

  it('rejects a file over the 150 MB per-attachment ceiling with no Graph calls', async () => {
    const hugePath = path.join(tmpDir, 'huge.pdf');
    const fd = fs.openSync(hugePath, 'w');
    fs.ftruncateSync(fd, 150 * 1024 * 1024 + 1); // sparse file — instant, ~0 bytes on disk
    fs.closeSync(fd);
    const fetchImpl = vi.fn();

    const res = await createDraftTool(authManager, {
      to: ['t@example.com'], subject: 's', body: 'b', preserveUserStyling: false,
      attachmentPaths: [hugePath]
    }, { fetchImpl });

    expect(res.isError).toBe(true);
    expect(res.content[0].text).toMatch(/150 ?MB/i);
    expect(postWithRetry).not.toHaveBeenCalled();
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it('surfaces a clear, actionable error when a chunk PUT fails — never sends', async () => {
    const bigPath = path.join(tmpDir, 'big.pdf');
    fs.writeFileSync(bigPath, Buffer.alloc(CHUNK + 10, 0x43));
    const fetchImpl = vi.fn().mockResolvedValue({ ok: false, status: 507 });

    const res = await createDraftTool(authManager, {
      to: ['t@example.com'], subject: 's', body: 'b', preserveUserStyling: false,
      attachmentPaths: [bigPath]
    }, { fetchImpl });

    expect(res.isError).toBe(true);
    expect(res.content[0].text).toMatch(/507/);
    expect(res.content[0].text).toMatch(/big\.pdf/);
    expect(postWithRetry.mock.calls.some(([p]) => p.includes('/send'))).toBe(false);
  });

  it('includes the webLink when there are no attachments — never sends', async () => {
    const res = await createDraftTool(authManager, {
      to: ['t@example.com'], subject: 's', body: 'b', preserveUserStyling: false
    });
    const out = JSON.parse(res.content[0].text);
    expect(out.draftId).toBe('draft-1');
    expect(out.webLink).toContain('outlook.office365.com');
    expect(makeRequest).toHaveBeenCalledWith('/me/messages/draft-1', { select: 'id,webLink' });
    expect(postWithRetry.mock.calls.some(([p]) => p.includes('/send'))).toBe(false);
  });

  it('treats bodyHtml as the HTML body; rejects bodyHtml + body together', async () => {
    const res = await createDraftTool(authManager, {
      to: ['t@example.com'], subject: 's', bodyHtml: '<p>hi</p>', preserveUserStyling: false
    });
    const draft = postWithRetry.mock.calls[0][1];
    expect(draft.body.contentType).toBe('HTML');
    expect(draft.body.content).toBe('<p>hi</p>');
    expect(postWithRetry.mock.calls.some(([p]) => p.includes('/send'))).toBe(false);

    postWithRetry.mockClear();
    const bad = await createDraftTool(authManager, {
      to: ['t@example.com'], subject: 's', bodyHtml: '<p>hi</p>', body: 'plain', preserveUserStyling: false
    });
    expect(bad.isError).toBe(true);
    expect(bad.content[0].text).toMatch(/bodyHtml|body/i);
    expect(postWithRetry).not.toHaveBeenCalled();
  });
});
