import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import crypto from 'crypto';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { fetchAttachmentContent, ATTACHMENT_META_SELECT } from '../../graph/attachmentContent.js';
import { GraphApiClient } from '../../graph/graphClient.js';
import { downloadAttachmentTool } from '../../tools/attachments/downloadAttachment.js';
import { saveAttachmentTool } from '../../tools/receipts/saveAttachment.js';
import {
  mockGraphClient, mockGraphErrorResult, installFakeGraphFetch, graphJsonResponse, graphBinaryResponse
} from '../helpers/mockGraph.js';

const PDF_BYTES = Buffer.concat([Buffer.from('%PDF-1.7\n'), Buffer.from([0xe2, 0xe3, 0xcf, 0xd3, 0x00, 0xff, 0x80]), Buffer.from('\n%%EOF')]);
const PDF_B64 = PDF_BYTES.toString('base64');

// Graph attachment ids are URL-safe base64: '-' and '_' inside, '=' padding at the end.
const ATTACHMENT_IDS = [
  ['a dash', 'AbCd-12Ef-34'],
  ['an underscore mid-string', 'AbCd_12Ef_34'],
  ['a trailing =', 'AbCd12Ef34gH='],
  ['dash, underscore and trailing =', 'Ab-Cd_12Ef=']
];

const fileMeta = (id, extra = {}) => ({
  '@odata.type': '#microsoft.graph.fileAttachment',
  id,
  name: 'Invoice-001.pdf',
  contentType: 'application/pdf',
  size: PDF_BYTES.length,
  isInline: false,
  lastModifiedDateTime: '2026-08-29T15:48:33Z',
  ...extra
});

describe('ATTACHMENT_META_SELECT', () => {
  it('lists plain properties only; @odata.type is not selectable', () => {
    expect(ATTACHMENT_META_SELECT).toBe('id,name,contentType,size,isInline,lastModifiedDateTime');
    expect(ATTACHMENT_META_SELECT).not.toContain('@odata.type');
  });
});

describe('fetchAttachmentContent', () => {
  it('route json: GETs the item with no $select and decodes contentBytes', async () => {
    const { client, makeRequest, getBinary } = mockGraphClient(() => fileMeta('a1', { contentBytes: PDF_B64 }));

    const result = await fetchAttachmentContent(client, '/me', 'm1', 'a1');

    expect(result.route).toBe('json');
    expect(result.buffer).toEqual(PDF_BYTES);
    expect(result.buffer.subarray(0, 4).toString('latin1')).toBe('%PDF');
    expect(result.attachment['@odata.type']).toBe('#microsoft.graph.fileAttachment');
    expect(result.attachment.name).toBe('Invoice-001.pdf');
    expect(result.attachment.contentBytes).toBeUndefined();
    expect(makeRequest).toHaveBeenCalledTimes(1);
    expect(makeRequest.mock.calls[0][0]).toBe('/me/messages/m1/attachments/a1');
    expect(makeRequest.mock.calls[0][1]?.select).toBeUndefined();
    expect(getBinary).not.toHaveBeenCalled();
  });

  it.each(ATTACHMENT_IDS)('passes an attachment id with %s to the request path unchanged', async (_label, id) => {
    const { client, makeRequest, getBinary } = mockGraphClient(
      (reqPath) => (reqPath.endsWith(`/attachments/${id}`)
        ? mockGraphErrorResult({ message: 'item route rejected' })
        : fileMeta(id)),
      async () => Buffer.from(PDF_BYTES)
    );

    const result = await fetchAttachmentContent(client, '/me', 'm1', id);

    expect(result.route).toBe('value');
    expect(makeRequest.mock.calls[0][0]).toBe(`/me/messages/m1/attachments/${id}`);
    expect(getBinary.mock.calls[0][0]).toBe(`/me/messages/m1/attachments/${id}/$value`);
  });

  it('uses the per-call mailbox base on both routes', async () => {
    const { client, makeRequest, getBinary } = mockGraphClient(
      () => mockGraphErrorResult({ message: 'item route rejected' }),
      async () => Buffer.from(PDF_BYTES)
    );

    await fetchAttachmentContent(client, '/users/careers%40example.com', 'm1', 'a1');

    expect(makeRequest.mock.calls[0][0]).toBe('/users/careers%40example.com/messages/m1/attachments/a1');
    expect(getBinary.mock.calls[0][0]).toBe('/users/careers%40example.com/messages/m1/attachments/a1/$value');
  });

  it('falls back to /$value when the item route fails, with a metadata read that never selects @odata.type', async () => {
    const { client, makeRequest, getBinary } = mockGraphClient(
      (reqPath, options) => (options.select ? fileMeta('a1') : mockGraphErrorResult({ message: 'item route rejected' })),
      async () => Buffer.from(PDF_BYTES)
    );

    const result = await fetchAttachmentContent(client, '/me', 'm1', 'a1');

    expect(result.route).toBe('value');
    expect(result.buffer).toEqual(PDF_BYTES);
    expect(getBinary).toHaveBeenCalledWith('/me/messages/m1/attachments/a1/$value');
    expect(result.attachment).toMatchObject({ name: 'Invoice-001.pdf', '@odata.type': '#microsoft.graph.fileAttachment' });
    for (const call of makeRequest.mock.calls) {
      expect(call[1]?.select || '').not.toContain('@odata.type');
    }
    expect(makeRequest.mock.calls[1][1].select).toBe(ATTACHMENT_META_SELECT);
  });

  it('falls back to /$value when a fileAttachment comes back without contentBytes', async () => {
    const { client, makeRequest, getBinary } = mockGraphClient(() => fileMeta('a1'), async () => Buffer.from(PDF_BYTES));

    const result = await fetchAttachmentContent(client, '/me', 'm1', 'a1');

    expect(result.route).toBe('value');
    expect(result.buffer).toEqual(PDF_BYTES);
    expect(result.attachment.name).toBe('Invoice-001.pdf');
    expect(makeRequest).toHaveBeenCalledTimes(1); // metadata already known from route 1
    expect(getBinary).toHaveBeenCalledTimes(1);
  });

  it('keeps going when the metadata read after a /$value success fails', async () => {
    const { client } = mockGraphClient(
      () => mockGraphErrorResult({ message: 'everything on the item route fails' }),
      async () => Buffer.from(PDF_BYTES)
    );

    const result = await fetchAttachmentContent(client, '/me', 'm1', 'a1');

    expect(result.buffer).toEqual(PDF_BYTES);
    expect(result.attachment).toEqual({ id: 'a1' });
  });

  it.each([
    ['#microsoft.graph.itemAttachment'],
    ['#microsoft.graph.referenceAttachment']
  ])('rejects %s clearly without trying the raw route', async (type) => {
    const { client, getBinary } = mockGraphClient(() => ({ '@odata.type': type, id: 'a1', name: 'forwarded', isInline: false }));

    const error = await fetchAttachmentContent(client, '/me', 'm1', 'a1').then(() => null, e => e);

    expect(error).toBeInstanceOf(Error);
    expect(error.message).toMatch(/is not a file attachment/);
    expect(error.message).toContain(type);
    expect(error.fetchDetails).toEqual({
      attemptedRoutes: ['json'], messageId: 'm1', attachmentId: 'a1', attachmentName: 'forwarded', graphError: null
    });
    expect(getBinary).not.toHaveBeenCalled();
  });

  it('rejects an item attachment discovered only after the raw route answered', async () => {
    const { client } = mockGraphClient(
      (reqPath, options) => (options.select
        ? { '@odata.type': '#microsoft.graph.itemAttachment', id: 'a1', name: 'forwarded', isInline: false }
        : mockGraphErrorResult({ message: 'item route rejected' })),
      async () => Buffer.from('From: someone@example.com\r\n\r\nbody')
    );

    const error = await fetchAttachmentContent(client, '/me', 'm1', 'a1').then(() => null, e => e);

    expect(error.message).toMatch(/is not a file attachment/);
    expect(error.fetchDetails.attemptedRoutes).toEqual(['json', 'value']);
  });

  it('reports both routes verbatim and the last route graphError when neither works', async () => {
    const jsonFailure = mockGraphErrorResult({
      status: 400, code: 'BadRequest', message: 'Item route said no', requestId: 'req-json',
      url: 'https://graph.microsoft.com/v1.0/me/messages/m1/attachments/a1'
    });
    const valueFailure = mockGraphErrorResult({
      status: 404, code: 'ErrorItemNotFound', message: 'Value route said no', requestId: 'req-value',
      url: 'https://graph.microsoft.com/v1.0/me/messages/m1/attachments/a1/$value'
    });
    const { client } = mockGraphClient(() => jsonFailure, async () => valueFailure);

    const error = await fetchAttachmentContent(client, '/me', 'm1', 'a1', { attachmentName: 'Invoice-001.pdf' }).then(() => null, e => e);

    expect(error).toBeInstanceOf(Error);
    expect(error.message).toContain(jsonFailure.content[0].text);
    expect(error.message).toContain(valueFailure.content[0].text);
    expect(error.message).toContain('HTTP 400 BadRequest: Item route said no');
    expect(error.message).toContain('HTTP 404 ErrorItemNotFound: Value route said no');
    expect(error.fetchDetails).toEqual({
      attemptedRoutes: ['json', 'value'],
      messageId: 'm1',
      attachmentId: 'a1',
      attachmentName: 'Invoice-001.pdf',
      graphError: valueFailure._errorDetails.graphError
    });
  });

  it('treats a thrown request error like a returned one', async () => {
    const client = {
      makeRequest: vi.fn().mockRejectedValue(new Error('socket hang up')),
      getBinary: vi.fn().mockRejectedValue(new Error('connection reset'))
    };

    const error = await fetchAttachmentContent(client, '/me', 'm1', 'a1').then(() => null, e => e);

    expect(error.message).toContain('socket hang up');
    expect(error.message).toContain('connection reset');
    expect(error.fetchDetails.attemptedRoutes).toEqual(['json', 'value']);
  });

  it('does not accept an empty body from the raw route as content', async () => {
    const { client } = mockGraphClient(() => mockGraphErrorResult({ message: 'item route rejected' }), async () => Buffer.alloc(0));

    const error = await fetchAttachmentContent(client, '/me', 'm1', 'a1').then(() => null, e => e);

    expect(error.message).toContain('Graph returned 0 bytes');
  });
});

/**
 * End to end through the real GraphApiClient and Graph SDK against a fake Graph
 * service that rejects '@odata.type' in a $select exactly as Graph does.
 */
describe('attachment bytes through the real client', () => {
  let fake, client, authManager, receiptsDir, workDir, saved;

  const SELECT_REJECTION = {
    error: {
      code: 'BadRequest',
      message: "Parsing OData Select and Expand failed: Term '@odata.type' is not valid in a $select or $expand expression.",
      innerError: { date: '2026-01-15T10:00:00', 'request-id': 'req-select', 'client-request-id': 'cid-select' }
    }
  };

  // opts: { itemRoute: 'ok'|'fail', valueRoute: 'ok'|'fail', type, pdf }
  function startFakeGraph(id, { itemRoute = 'ok', valueRoute = 'ok', type = '#microsoft.graph.fileAttachment', pdf = PDF_BYTES } = {}) {
    const listPath = '/v1.0/me/messages/m1/attachments';
    const itemPath = `${listPath}/${id}`;
    fake = installFakeGraphFetch((request, url) => {
      if (url.searchParams.get('$select')?.includes('@odata.type')) {
        return graphJsonResponse(400, SELECT_REJECTION);
      }
      if (url.pathname === listPath) {
        return graphJsonResponse(200, { value: [{ id, name: 'Invoice-001.pdf', contentType: 'application/pdf', size: pdf.length, isInline: false }] });
      }
      if (url.pathname === itemPath) {
        if (url.searchParams.has('$select')) {
          return graphJsonResponse(200, fileMeta(id, { '@odata.type': type, size: pdf.length }));
        }
        if (itemRoute === 'fail') {
          return graphJsonResponse(400, { error: { code: 'ErrorInvalidIdMalformed', message: 'The item route failed', innerError: { 'request-id': 'req-item' } } });
        }
        return graphJsonResponse(200, fileMeta(id, { '@odata.type': type, size: pdf.length, ...(type === '#microsoft.graph.fileAttachment' && { contentBytes: pdf.toString('base64') }) }));
      }
      if (url.pathname === `${itemPath}/$value`) {
        if (valueRoute === 'fail') {
          return graphJsonResponse(404, { error: { code: 'ErrorItemNotFound', message: 'The value route failed', innerError: { 'request-id': 'req-value' } } });
        }
        return graphBinaryResponse(pdf);
      }
      return graphJsonResponse(404, { error: { code: 'ErrorItemNotFound', message: `unexpected request ${url.pathname}` } });
    });
    return fake;
  }

  beforeEach(() => {
    client = new GraphApiClient({ tokenManager: { getAccessToken: async () => 'test-token-value' } });
    authManager = { ensureAuthenticated: async () => true, getGraphApiClient: () => client };
    receiptsDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ac-receipts-'));
    workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ac-work-'));
    saved = { receipts: process.env.MCP_OUTLOOK_RECEIPTS_DIR, work: process.env.MCP_OUTLOOK_WORK_DIR };
    process.env.MCP_OUTLOOK_RECEIPTS_DIR = receiptsDir;
    process.env.MCP_OUTLOOK_WORK_DIR = workDir;
  });

  afterEach(() => {
    fake?.restore();
    for (const [key, value] of [['MCP_OUTLOOK_RECEIPTS_DIR', saved.receipts], ['MCP_OUTLOOK_WORK_DIR', saved.work]]) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
    fs.rmSync(receiptsDir, { recursive: true, force: true });
    fs.rmSync(workDir, { recursive: true, force: true });
  });

  const requestedUrls = () => fake.requests.map(r => r.url);

  it('the fake service rejects @odata.type in a $select like Graph does', async () => {
    startFakeGraph('a1');
    const result = await client.makeRequest('/me/messages/m1/attachments/a1', { select: 'id,name,@odata.type' });
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain("Term '@odata.type' is not valid in a $select or $expand expression.");
  });

  it.each(ATTACHMENT_IDS)('download_attachment saves a real PDF for an id with %s', async (_label, id) => {
    startFakeGraph(id);

    const res = await downloadAttachmentTool(authManager, { messageId: 'm1', attachmentId: id, saveToFile: true });

    expect(res.isError).toBeUndefined();
    const out = JSON.parse(res.content[0].text);
    const onDisk = fs.readFileSync(out.savedPath);
    expect(onDisk.length).toBeGreaterThan(0);
    expect(onDisk.subarray(0, 4).toString('latin1')).toBe('%PDF');
    expect(onDisk).toEqual(PDF_BYTES);
    expect(out.attachmentType).toBe('#microsoft.graph.fileAttachment');
    expect(requestedUrls()).toEqual([`https://graph.microsoft.com/v1.0/me/messages/m1/attachments/${id}`]);
    expect(requestedUrls().join('\n')).not.toContain('@odata.type');
  });

  it.each(ATTACHMENT_IDS)('save_attachment saves a real PDF for an id with %s', async (_label, id) => {
    startFakeGraph(id);

    const res = await saveAttachmentTool(authManager, { messageId: 'm1', attachmentId: id, fileName: 'Acme Invoice.pdf' });

    expect(res.isError).toBeUndefined();
    const out = JSON.parse(res.content[0].text);
    expect(fs.readFileSync(out.savedPath)).toEqual(PDF_BYTES);
    expect(out.attachmentId).toBe(id);
    expect(out.attachmentName).toBe('Invoice-001.pdf');
    expect(requestedUrls().join('\n')).not.toContain('@odata.type');
  });

  it('save_attachment auto-selects from the listing and fetches without any @odata.type select', async () => {
    startFakeGraph('Ab-Cd_12Ef=');

    const res = await saveAttachmentTool(authManager, { messageId: 'm1', fileName: 'Acme Invoice.pdf' });

    expect(res.isError).toBeUndefined();
    for (const url of requestedUrls()) {
      expect(new URL(url).searchParams.get('$select') || '').not.toContain('@odata.type');
    }
    expect(fs.readFileSync(JSON.parse(res.content[0].text).savedPath)).toEqual(PDF_BYTES);
  });

  it('download_attachment inline content (includeContent) also avoids the rejected select', async () => {
    startFakeGraph('a1');

    const res = await downloadAttachmentTool(authManager, { messageId: 'm1', attachmentId: 'a1', includeContent: true, decodeContent: false });

    const out = JSON.parse(res.content[0].text);
    expect(Buffer.from(out.contentBytes, 'base64')).toEqual(PDF_BYTES);
    for (const url of requestedUrls()) {
      expect(new URL(url).searchParams.get('$select') || '').not.toContain('@odata.type');
    }
  });

  it('uses /$value when the item route returns 400, saving identical bytes', async () => {
    startFakeGraph('Ab-Cd_12Ef=', { itemRoute: 'fail' });

    const res = await downloadAttachmentTool(authManager, { messageId: 'm1', attachmentId: 'Ab-Cd_12Ef=', saveToFile: true });

    expect(res.isError).toBeUndefined();
    const out = JSON.parse(res.content[0].text);
    expect(fs.readFileSync(out.savedPath)).toEqual(PDF_BYTES);
    expect(requestedUrls().some(u => u.endsWith('/attachments/Ab-Cd_12Ef=/$value'))).toBe(true);
    expect(requestedUrls().join('\n')).not.toContain('@odata.type');
  });

  it('writes large files byte-exact server-side, with no inline-base64 ceiling', async () => {
    const big = Buffer.concat([Buffer.from('%PDF-1.7\n'), crypto.randomBytes(1024 * 1024)]);
    startFakeGraph('a1', { pdf: big });

    const res = await downloadAttachmentTool(authManager, { messageId: 'm1', attachmentId: 'a1', saveToFile: true });

    const out = JSON.parse(res.content[0].text);
    expect(out.size).toBe(big.length);
    expect(out.sha256).toBe(crypto.createHash('sha256').update(big).digest('hex'));
    expect(fs.readFileSync(out.savedPath)).toEqual(big);
    expect(res.content[0].text.length).toBeLessThan(2000);
  });

  it('large file through the /$value route is byte-exact too', async () => {
    const big = Buffer.concat([Buffer.from('%PDF-1.7\n'), crypto.randomBytes(1024 * 1024)]);
    startFakeGraph('a1', { pdf: big, itemRoute: 'fail' });

    const res = await downloadAttachmentTool(authManager, { messageId: 'm1', attachmentId: 'a1', saveToFile: true });

    expect(fs.readFileSync(JSON.parse(res.content[0].text).savedPath)).toEqual(big);
  });

  it('surfaces both routes Graph errors, with ids and URLs, when neither delivers the bytes', async () => {
    startFakeGraph('a1', { itemRoute: 'fail', valueRoute: 'fail' });

    const res = await downloadAttachmentTool(authManager, { messageId: 'm1', attachmentId: 'a1', saveToFile: true });

    expect(res.isError).toBe(true);
    const text = res.content[0].text;
    expect(text).toContain('HTTP 400 ErrorInvalidIdMalformed: The item route failed');
    expect(text).toContain('request-id: req-item');
    expect(text).toContain('GET https://graph.microsoft.com/v1.0/me/messages/m1/attachments/a1');
    expect(text).toContain('HTTP 404 ErrorItemNotFound: The value route failed');
    expect(text).toContain('request-id: req-value');
    expect(text).toContain('GET https://graph.microsoft.com/v1.0/me/messages/m1/attachments/a1/$value');
    expect(res._errorDetails.fetchDetails).toMatchObject({ attemptedRoutes: ['json', 'value'], messageId: 'm1', attachmentId: 'a1' });
    expect(res._errorDetails.graphError).toMatchObject({ status: 404, code: 'ErrorItemNotFound', requestId: 'req-value', method: 'GET' });
  });

  it('rejects an item attachment clearly and never requests its raw bytes', async () => {
    startFakeGraph('a1', { type: '#microsoft.graph.itemAttachment' });

    const res = await downloadAttachmentTool(authManager, { messageId: 'm1', attachmentId: 'a1', saveToFile: true });

    expect(res.isError).toBe(true);
    expect(res.content[0].text).toMatch(/not a file attachment/);
    expect(requestedUrls().some(u => u.endsWith('/$value'))).toBe(false);
    expect(fs.readdirSync(receiptsDir)).toEqual([]);
  });
});
