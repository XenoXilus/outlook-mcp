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
