import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { renderEmailPdfCore, renderEmailPdfTool } from '../../tools/receipts/renderEmailPdf.js';

const PDF_BYTES = Buffer.from('%PDF-1.4 rendered-test');

describe('renderEmailPdfCore (FR-3/FR-7)', () => {
  let receiptsDir, makeRequest, graphApiClient, savedEnv;

  beforeEach(() => {
    receiptsDir = fs.mkdtempSync(path.join(os.tmpdir(), 'render-'));
    savedEnv = process.env.MCP_OUTLOOK_RECEIPTS_DIR;
    process.env.MCP_OUTLOOK_RECEIPTS_DIR = receiptsDir;
    makeRequest = vi.fn();
    graphApiClient = { makeRequest };
  });

  afterEach(() => {
    if (savedEnv === undefined) delete process.env.MCP_OUTLOOK_RECEIPTS_DIR;
    else process.env.MCP_OUTLOOK_RECEIPTS_DIR = savedEnv;
    fs.rmSync(receiptsDir, { recursive: true, force: true });
  });

  it('renders an email as PDF and returns method:rendered, a saved .pdf path, and sha256', async () => {
    makeRequest.mockResolvedValueOnce({
      id: 'm1',
      subject: 'Hooli Store receipt',
      from: { emailAddress: { address: 'store-noreply@hooli.example' } },
      receivedDateTime: '2026-06-09T10:00:00Z',
      body: { contentType: 'html', content: '<p>Total $14.99</p>' }
    });

    const renderImpl = vi.fn().mockResolvedValue(PDF_BYTES);

    const result = await renderEmailPdfCore(graphApiClient, {
      messageId: 'm1',
      fileName: 'Hooli 09Jun26 Invoice.pdf'
    }, { renderImpl });

    expect(result.method).toBe('rendered');
    expect(result.savedPath).toMatch(/\.pdf$/i);
    expect(fs.existsSync(result.savedPath)).toBe(true);
    expect(result.sha256).toMatch(/^[0-9a-f]{64}$/);
    expect(renderImpl).toHaveBeenCalledOnce();
  });

  describe('re-rendering the same email under onExisting skip', () => {
    const message = (html) => ({
      id: 'm1',
      subject: 'Hooli Store receipt',
      from: { emailAddress: { address: 'store-noreply@hooli.example' } },
      receivedDateTime: '2026-06-09T10:00:00Z',
      body: { contentType: 'html', content: html }
    });
    // Chrome stamps each rendering with its creation time.
    const stampedRender = () => {
      let stamp = 0;
      return vi.fn().mockImplementation(async (doc) => {
        stamp++;
        return Buffer.from(`%PDF-1.4\n<< /CreationDate (D:2026060912000${stamp}Z) /ModDate (D:2026060912000${stamp}Z) >>\n${doc}`);
      });
    };
    const args = { messageId: 'm1', fileName: 'Hooli 09Jun26 Invoice.pdf' };

    it('skips the earlier rendering when only the render timestamps differ', async () => {
      makeRequest.mockResolvedValue(message('<p>Total $14.99</p>'));
      const renderImpl = stampedRender();

      const first = await renderEmailPdfCore(graphApiClient, args, { renderImpl });
      const second = await renderEmailPdfCore(graphApiClient, args, { renderImpl });
      const third = await renderEmailPdfCore(graphApiClient, args, { renderImpl });

      expect(first.action).toBe('saved');
      for (const again of [second, third]) {
        expect(again).toMatchObject({ action: 'skipped', savedPath: first.savedPath, sha256: first.sha256, method: 'rendered' });
      }
      expect(fs.readdirSync(receiptsDir)).toEqual(['Hooli 09Jun26 Invoice.pdf']);
    });

    it('keeps a rendering of changed content as a new version', async () => {
      const renderImpl = stampedRender();
      makeRequest.mockResolvedValueOnce(message('<p>Total $14.99</p>'));
      await renderEmailPdfCore(graphApiClient, args, { renderImpl });
      makeRequest.mockResolvedValueOnce(message('<p>Total $19.99</p>'));
      const changed = await renderEmailPdfCore(graphApiClient, args, { renderImpl });

      expect(changed).toMatchObject({ action: 'versioned', existingDiffers: true });
      expect(fs.readdirSync(receiptsDir).sort()).toEqual(['Hooli 09Jun26 Invoice (2).pdf', 'Hooli 09Jun26 Invoice.pdf']);
    });
  });
});

describe('renderEmailPdfTool (tool-wrapper deps threading)', () => {
  let receiptsDir, makeRequest, authManager, savedEnv;

  beforeEach(() => {
    receiptsDir = fs.mkdtempSync(path.join(os.tmpdir(), 'render-tool-'));
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

  it('forwards deps.renderImpl through the tool wrapper to the core', async () => {
    makeRequest.mockResolvedValueOnce({
      id: 'm1',
      subject: 'Test receipt',
      from: { emailAddress: { address: 'test@example.com' } },
      receivedDateTime: '2026-06-09T10:00:00Z',
      body: { contentType: 'html', content: '<p>Test</p>' }
    });

    const renderImpl = vi.fn().mockResolvedValue(PDF_BYTES);
    const res = await renderEmailPdfTool(authManager, {
      messageId: 'm1',
      fileName: 'Test 09Jun26 Invoice.pdf'
    }, { renderImpl });

    expect(res.isError).toBeUndefined();
    const out = JSON.parse(res.content[0].text);
    expect(out.method).toBe('rendered');
    expect(fs.existsSync(out.savedPath)).toBe(true);
    expect(renderImpl).toHaveBeenCalledOnce();
  });
});
