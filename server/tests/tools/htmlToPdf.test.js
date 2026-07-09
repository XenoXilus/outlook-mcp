import { describe, it, expect } from 'vitest';
import fs from 'fs';
import {
  sanitizeEmailHtml,
  buildRenderDocument,
  findChromeBinary,
  buildChromeArgs,
  renderHtmlToPdfBuffer
} from '../../tools/receipts/htmlToPdf.js';
import { isPdfBuffer } from '../../utils/receiptFiles.js';

describe('htmlToPdf (FR-7)', () => {
  it('strips scripts, iframes and event handlers', () => {
    const dirty = '<p onclick="evil()">£13.37</p><script>evil()</script><iframe src="https://evil.example.com"></iframe>';
    const clean = sanitizeEmailHtml(dirty);
    expect(clean).toContain('£13.37');
    expect(clean).not.toContain('<script');
    expect(clean).not.toContain('<iframe');
    expect(clean).not.toContain('onclick');
  });

  it('builds a printable document with an audit header', () => {
    const doc = buildRenderDocument({
      subject: 'Your Hooli Store order receipt',
      fromAddress: 'store-noreply@hooli.example',
      receivedDateTime: '2026-06-09T10:00:00Z',
      bodyHtml: '<p>Total: £13.37</p>'
    });
    expect(doc).toContain('<!DOCTYPE html>');
    expect(doc).toContain('Your Hooli Store order receipt');
    expect(doc).toContain('store-noreply@hooli.example');
    expect(doc).toContain('Total: £13.37');
    expect(doc).toContain('Rendered from email'); // fallback provenance marker
  });

  it('prefers MCP_OUTLOOK_CHROME_PATH when set', () => {
    const saved = process.env.MCP_OUTLOOK_CHROME_PATH;
    process.env.MCP_OUTLOOK_CHROME_PATH = '/custom/chrome';
    try {
      expect(findChromeBinary()).toBe('/custom/chrome');
    } finally {
      if (saved === undefined) delete process.env.MCP_OUTLOOK_CHROME_PATH;
      else process.env.MCP_OUTLOOK_CHROME_PATH = saved;
    }
  });

  // Remote-content blocking — NFR-5 / FR-7 defense-in-depth

  it('strips remote URLs from img src (tracking pixel)', () => {
    const dirty = '<p>Receipt</p><img src="https://tracker.example/pixel.gif" alt="pixel">';
    const clean = sanitizeEmailHtml(dirty);
    expect(clean).not.toContain('https://tracker.example');
    expect(clean).not.toMatch(/src="https?:/);
  });

  it('removes <link> stylesheet references', () => {
    const dirty = '<link rel="stylesheet" href="https://cdn.example/x.css"><p>text</p>';
    const clean = sanitizeEmailHtml(dirty);
    expect(clean).not.toContain('<link');
    expect(clean).not.toContain('cdn.example');
  });

  it('strips srcset attributes containing remote URLs', () => {
    const dirty = '<img srcset="https://cdn.example/img@2x.png 2x, https://cdn.example/img.png 1x" alt="logo">';
    const clean = sanitizeEmailHtml(dirty);
    expect(clean).not.toContain('cdn.example');
    expect(clean).not.toContain('srcset=');
  });

  it('strips srcset when ANY token is a remote URL (mixed local+remote bypass)', () => {
    const dirty = '<img srcset="relative.png 1x, https://tracker.example/large.png 2x" alt="img">';
    const clean = sanitizeEmailHtml(dirty);
    expect(clean).not.toContain('tracker.example');
    expect(clean).not.toMatch(/srcset="[^"]*https?:/);
  });

  it('strips inline style attributes containing url() with a remote reference', () => {
    const dirty = '<div style="background:url(https://evil.example/bg.png)">content</div>';
    const clean = sanitizeEmailHtml(dirty);
    expect(clean).not.toContain('evil.example');
    expect(clean).not.toContain('url(');
  });

  it('preserves data: image URIs and anchor href attributes', () => {
    const input = [
      '<img src="data:image/png;base64,iVBORw0KGgo=" alt="logo">',
      '<a href="https://pay.stripe.com/invoices/123">View invoice</a>'
    ].join('');
    const clean = sanitizeEmailHtml(input);
    expect(clean).toContain('data:image/png;base64,iVBORw0KGgo=');
    expect(clean).toContain('href="https://pay.stripe.com/invoices/123"');
    expect(clean).toContain('View invoice');
  });

  it('Chrome args include --blink-settings=imagesEnabled=false and --disable-remote-fonts', () => {
    const args = buildChromeArgs('/tmp/email.pdf', '/tmp/email.html');
    expect(args).toContain('--blink-settings=imagesEnabled=false');
    expect(args).toContain('--disable-remote-fonts');
  });

  // Real render — only when a Chrome binary actually exists on this machine.
  const chrome = findChromeBinary();
  it.skipIf(!chrome || !fs.existsSync(chrome))('renders HTML to a real PDF via headless Chrome', async () => {
    const buffer = await renderHtmlToPdfBuffer('<!DOCTYPE html><html><body><h1>£13.37</h1></body></html>');
    expect(isPdfBuffer(buffer)).toBe(true);
  }, 90000);
});
