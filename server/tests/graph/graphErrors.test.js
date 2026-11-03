import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { GraphError } from '@microsoft/microsoft-graph-client';
import { GraphApiClient, readHeader, parseGraphErrorBody, describeFailedRequest } from '../../graph/graphClient.js';
import { convertErrorToToolError, formatGraphErrorText } from '../../utils/mcpErrorResponse.js';
import { installFakeGraphFetch, graphJsonResponse, graphBinaryResponse } from '../helpers/mockGraph.js';

/**
 * The Graph SDK's GraphError exposes statusCode/code/requestId/date/body (a
 * JSON *string*) and `headers` as a fetch Headers instance. These tests run the
 * real client against a fake Graph service to prove the real failure — status,
 * code, message, request ids and the exact request — reaches the MCP error.
 */

const TOKEN = 'test-token-value-do-not-leak';
const SELECT_MESSAGE = "Parsing OData Select and Expand failed: Term '@odata.type' is not valid in a $select or $expand expression.";
const ATTACHMENT_PATH = '/me/messages/m1/attachments/a1';
const ATTACHMENT_URL = 'https://graph.microsoft.com/v1.0/me/messages/m1/attachments/a1';

function newClient() {
  return new GraphApiClient({ tokenManager: { getAccessToken: async () => TOKEN } });
}

describe('GraphApiClient failure details', () => {
  let fake, client;

  afterEach(() => fake?.restore());
  beforeEach(() => { client = newClient(); });

  it('surfaces status, code, message, request ids and METHOD url verbatim for a 400', async () => {
    fake = installFakeGraphFetch(() => graphJsonResponse(400, {
      error: {
        code: 'BadRequest',
        message: SELECT_MESSAGE,
        innerError: { date: '2026-01-15T10:00:00', 'request-id': '11111111-aaaa', 'client-request-id': '22222222-bbbb' }
      }
    }, { 'request-id': '11111111-aaaa', 'client-request-id': '22222222-bbbb' }));

    const result = await client.makeRequest(ATTACHMENT_PATH, { select: 'id,name,@odata.type' });

    expect(result.isError).toBe(true);
    const text = result.content[0].text;
    expect(text).toContain('HTTP 400');
    expect(text).toContain('BadRequest');
    expect(text).toContain(SELECT_MESSAGE);
    expect(text).toContain('request-id: 11111111-aaaa');
    expect(text).toContain('client-request-id: 22222222-bbbb');
    expect(text).toContain(`GET ${ATTACHMENT_URL}?$select=id,name,@odata.type`);
    expect(text).not.toContain('Please check the request parameters');

    expect(result._errorDetails.graphError).toEqual({
      status: 400,
      code: 'BadRequest',
      message: SELECT_MESSAGE,
      requestId: '11111111-aaaa',
      clientRequestId: '22222222-bbbb',
      date: '2026-01-15T10:00:00',
      method: 'GET',
      url: `${ATTACHMENT_URL}?$select=id,name,@odata.type`
    });
  });

  it('reads request ids and date from the response headers when the body has none', async () => {
    fake = installFakeGraphFetch(() => graphJsonResponse(403, {
      error: { code: 'ErrorAccessDenied', message: 'Access is denied. Check credentials and try again.' }
    }, { 'request-id': 'hdr-req-id', 'client-request-id': 'hdr-client-id', date: 'Thu, 15 Jan 2026 10:00:00 GMT' }));

    const result = await client.makeRequest(ATTACHMENT_PATH);

    const g = result._errorDetails.graphError;
    expect(g).toMatchObject({
      status: 403,
      code: 'ErrorAccessDenied',
      message: 'Access is denied. Check credentials and try again.',
      requestId: 'hdr-req-id',
      clientRequestId: 'hdr-client-id',
      date: 'Thu, 15 Jan 2026 10:00:00 GMT',
      method: 'GET',
      url: ATTACHMENT_URL
    });
    expect(result.content[0].text).toContain('request-id: hdr-req-id');
    expect(result.content[0].text).toContain('client-request-id: hdr-client-id');
  });

  it.each([
    [403, 'ErrorAccessDenied', 'Access is denied.'],
    [404, 'ErrorItemNotFound', 'The specified object was not found in the store.'],
    [409, 'ErrorIrresolvableConflict', 'The send or update operation could not be performed.']
  ])('applies to every non-retried failure, not only 400 (%i)', async (status, code, message) => {
    fake = installFakeGraphFetch(() => graphJsonResponse(status, {
      error: { code, message, innerError: { 'request-id': `req-${status}` } }
    }));

    const result = await client.makeRequest(ATTACHMENT_PATH);
    const text = result.content[0].text;
    expect(text).toContain(`HTTP ${status}`);
    expect(text).toContain(code);
    expect(text).toContain(message);
    expect(text).toContain(`request-id: req-${status}`);
    expect(text).toContain(`GET ${ATTACHMENT_URL}`);
    expect(result._errorDetails.graphError).toMatchObject({ status, code, message, requestId: `req-${status}` });
  });

  it('falls back to the client-request-id we sent when Graph does not echo one', async () => {
    fake = installFakeGraphFetch(() => new Response('<html>Not Found</html>', {
      status: 404, headers: { 'content-type': 'text/html' }
    }));

    const result = await client.makeRequest(ATTACHMENT_PATH);

    const sent = fake.requests[0].headers.get('client-request-id');
    expect(sent).toBeTruthy();
    const g = result._errorDetails.graphError;
    expect(g.status).toBe(404);
    expect(g.clientRequestId).toBe(sent);
    expect(g.url).toBe(ATTACHMENT_URL);
    expect(result.content[0].text).toContain(`client-request-id: ${sent}`);
  });

  it('records the full query string as sent and never the Authorization header or token', async () => {
    fake = installFakeGraphFetch(() => graphJsonResponse(400, { error: { code: 'BadRequest', message: 'bad query' } }));

    const result = await client.makeRequest('/me/messages', {
      select: 'id,subject',
      filter: "contains(subject,'Acme')",
      search: '"Acme invoice"',
      orderby: 'receivedDateTime desc',
      top: 5,
      expand: 'attachments'
    });

    const { url } = result._errorDetails.graphError;
    expect(url.startsWith('https://graph.microsoft.com/v1.0/me/messages?')).toBe(true);
    for (const part of ['$select=id,subject', "$filter=contains(subject,'Acme')", '$search="Acme invoice"',
      '$orderby=receivedDateTime desc', '$top=5', '$expand=attachments']) {
      expect(url).toContain(part);
    }

    // The token was really sent, and appears nowhere in what we return
    expect(fake.requests[0].headers.get('authorization')).toBe(`Bearer ${TOKEN}`);
    const serialised = JSON.stringify(result);
    expect(serialised).not.toContain(TOKEN);
    expect(serialised.toLowerCase()).not.toContain('authorization');
    expect(serialised.toLowerCase()).not.toContain('bearer');
  });

  it('keeps retryable 5xx failures as service-unavailable but with Graph details attached', () => {
    const error = new GraphError(503, 'The service is unavailable.');
    error.code = 'ServiceNotAvailable';
    error.requestId = 'req-503';
    error.headers = new Headers({ 'client-request-id': 'cid-503' });
    error.body = JSON.stringify({ code: 'ServiceNotAvailable', message: 'The service is unavailable.' });

    const details = client.extractErrorDetails(error, 'cid', '/me/messages', 'GET', 'https://graph.microsoft.com/v1.0/me/messages');
    const result = client.handleGraphError(error, details);

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('temporarily unavailable');
    expect(result.content[0].text).toContain('HTTP 503 ServiceNotAvailable: The service is unavailable.');
    expect(result._errorDetails.graphError).toMatchObject({ status: 503, requestId: 'req-503', clientRequestId: 'cid-503' });
    expect(result._errorDetails.retryable).toBe(true);
  });

  it('honours a Retry-After header delivered as a Headers instance', () => {
    const error = new GraphError(429, 'Too many requests');
    error.headers = new Headers({ 'retry-after': '7' });
    expect(client.extractRetryAfter(error)).toBe(7000);
  });
});

describe('GraphApiClient failure log (stderr)', () => {
  let fake, client, errorSpy;

  beforeEach(() => {
    client = newClient();
    errorSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
  });
  afterEach(() => {
    fake?.restore();
    errorSpy.mockRestore();
  });

  const loggedText = () => errorSpy.mock.calls.map((args) => args.map(String).join(' ')).join('\n');
  const loggedRequest = () => {
    const call = errorSpy.mock.calls.find((args) => args[0] === 'Failed request:');
    return call ? JSON.parse(call[1]) : null;
  };
  const failWith400 = () => installFakeGraphFetch(() => graphJsonResponse(400, {
    error: { code: 'BadRequest', message: 'bad request body' }
  }));

  it('includes method, full URL, redacted headers and the JSON body, and never the token', async () => {
    fake = failWith400();
    const body = { message: { subject: 'Quarterly summary', toRecipients: [{ emailAddress: { address: 'someone@example.com' } }] } };

    const result = await client.makeRequest('/me/sendMail', { body }, 'POST');

    expect(result.isError).toBe(true);
    const logged = loggedRequest();
    expect(logged.method).toBe('POST');
    expect(logged.url).toBe('https://graph.microsoft.com/v1.0/me/sendMail');
    expect(logged.headers.Authorization).toBe('[redacted]');
    expect(logged.headers['client-request-id']).toBe(fake.requests[0].headers.get('client-request-id'));
    expect(logged.body).toBe(JSON.stringify(body));

    // The token really went out on the wire, and appears nowhere in stderr or in the returned error
    expect(fake.requests[0].headers.get('authorization')).toBe(`Bearer ${TOKEN}`);
    expect(loggedText()).not.toContain(TOKEN);
    expect(loggedText().toLowerCase()).not.toContain('bearer');
    expect(JSON.stringify(result)).not.toContain(TOKEN);
    // The request detail is a log-only addition: the MCP error text stays as before
    expect(result.content[0].text).not.toContain('Quarterly summary');
    expect(result.content[0].text.toLowerCase()).not.toContain('authorization');
  });

  it('keeps the existing error logging alongside the request', async () => {
    fake = failWith400();
    await client.makeRequest('/me/sendMail', { body: { a: 1 } }, 'POST');
    const text = loggedText();
    expect(text).toContain('Error details:');
    expect(text).toContain('Graph API Error - Full Details:');
    expect(text).toContain('Failed request:');
  });

  it.each(['POST', 'PATCH', 'PUT'])('truncates a long %s body to about 2,000 characters', async (method) => {
    fake = failWith400();
    const body = { content: 'x'.repeat(5000) };

    await client.makeRequest('/me/messages/m1', { body }, method);

    const logged = loggedRequest();
    expect(logged.method).toBe(method);
    expect(logged.body.startsWith(JSON.stringify(body).slice(0, 2000))).toBe(true);
    expect(logged.body).toContain('truncated');
    expect(logged.body).not.toContain('x'.repeat(2001));
    expect(logged.body.length).toBeLessThan(2100);
  });

  it('logs the URL and headers of a GET but no body', async () => {
    fake = failWith400();

    await client.makeRequest('/me/messages', { select: 'id,subject', top: 5 });

    const logged = loggedRequest();
    expect(logged.method).toBe('GET');
    expect(logged.url).toBe('https://graph.microsoft.com/v1.0/me/messages?$select=id,subject&$top=5');
    expect(logged.headers.Authorization).toBe('[redacted]');
    expect(logged).not.toHaveProperty('body');
    expect(loggedText()).not.toContain(TOKEN);
  });

  it('redacts any credential-like header, by name or by value', () => {
    const request = {
      _headers: {
        'client-request-id': 'cid-1',
        Cookie: 'sid=abc123',
        'X-Api-Key': 'key-abc123',
        'X-Refresh-Token': 'refresh-abc123',
        Authorization: `Bearer ${TOKEN}`,
        'X-Custom': `Bearer ${TOKEN}`
      },
      buildFullUrl: () => 'https://graph.microsoft.com/v1.0/me/messages'
    };

    const described = describeFailedRequest(request, { method: 'get', path: '/me/messages' });

    expect(described.method).toBe('GET');
    expect(described.url).toBe('https://graph.microsoft.com/v1.0/me/messages');
    expect(described.headers).toEqual({
      'client-request-id': 'cid-1',
      Cookie: '[redacted]',
      'X-Api-Key': '[redacted]',
      'X-Refresh-Token': '[redacted]',
      Authorization: '[redacted]',
      'X-Custom': '[redacted]'
    });
    expect(JSON.stringify(described)).not.toContain(TOKEN);
  });

  it('describes a binary body by size instead of dumping it', () => {
    const described = describeFailedRequest({}, { method: 'PUT', path: '/me/drive/x', body: Buffer.alloc(4096, 1) });
    expect(described.body).toBe('[binary body, 4096 bytes]');
    expect(described.url).toBe('/me/drive/x');
  });
});

describe('GraphApiClient raw-byte requests', () => {
  let fake, client;

  afterEach(() => fake?.restore());
  beforeEach(() => { client = newClient(); });

  // Bytes that are not valid UTF-8 and include NUL, so any text decoding corrupts them
  const BINARY = Buffer.concat([
    Buffer.from('%PDF-1.7\n'),
    Buffer.from([0xe2, 0xe3, 0xcf, 0xd3, 0x00, 0xff, 0xfe, 0x80, 0x0d, 0x0a, 0xc3, 0x28])
  ]);

  it('getBinary resolves to a Buffer of the exact bytes', async () => {
    fake = installFakeGraphFetch(() => graphBinaryResponse(BINARY));

    const result = await client.getBinary(`${ATTACHMENT_PATH}/$value`);

    expect(Buffer.isBuffer(result)).toBe(true);
    expect(result).toEqual(BINARY);
    expect(result.subarray(0, 4).toString('latin1')).toBe('%PDF');
    expect(fake.requests[0].url).toBe(`${ATTACHMENT_URL}/$value`);
  });

  it('getBinary failures carry the decoded Graph error, ids and URL like makeRequest', async () => {
    fake = installFakeGraphFetch(() => graphJsonResponse(404, {
      error: {
        code: 'ErrorItemNotFound',
        message: 'The specified object was not found in the store.',
        innerError: { date: '2026-01-15T10:00:00', 'request-id': 'req-bin', 'client-request-id': 'cid-bin' }
      }
    }));

    const result = await client.getBinary(`${ATTACHMENT_PATH}/$value`);

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('HTTP 404 ErrorItemNotFound: The specified object was not found in the store.');
    expect(result.content[0].text).toContain(`GET ${ATTACHMENT_URL}/$value`);
    expect(result._errorDetails.graphError).toEqual({
      status: 404,
      code: 'ErrorItemNotFound',
      message: 'The specified object was not found in the store.',
      requestId: 'req-bin',
      clientRequestId: 'cid-bin',
      date: '2026-01-15T10:00:00',
      method: 'GET',
      url: `${ATTACHMENT_URL}/$value`
    });
  });

  // Graph attachment ids are URL-safe base64: '-' and '_' inside, '=' padding at the end.
  it.each([
    ['a dash', 'AbCd-12Ef-34'],
    ['an underscore mid-string', 'AbCd_12Ef_34'],
    ['a trailing =', 'AbCd12Ef34gH='],
    ['dash, underscore and trailing =', 'Ab-Cd_12Ef=']
  ])('sends an attachment id with %s to Graph unchanged', async (_label, id) => {
    fake = installFakeGraphFetch((request) => (request.url.endsWith('/$value')
      ? graphBinaryResponse(BINARY)
      : graphJsonResponse(200, { id, name: 'Invoice.pdf', contentBytes: BINARY.toString('base64') })));

    await client.makeRequest(`/me/messages/m1/attachments/${id}`);
    await client.getBinary(`/me/messages/m1/attachments/${id}/$value`);

    expect(fake.requests.map(r => r.url)).toEqual([
      `https://graph.microsoft.com/v1.0/me/messages/m1/attachments/${id}`,
      `https://graph.microsoft.com/v1.0/me/messages/m1/attachments/${id}/$value`
    ]);
  });
});

describe('Graph error parsing helpers', () => {
  it('readHeader reads a fetch Headers instance, which bracket access cannot', () => {
    const headers = new Headers({ 'Request-Id': 'abc' });
    expect(headers['request-id']).toBeUndefined();
    expect(readHeader(headers, 'request-id')).toBe('abc');
    expect(readHeader(headers, 'client-request-id')).toBeUndefined();
  });

  it('readHeader matches plain objects case-insensitively and tolerates absence', () => {
    expect(readHeader({ 'Retry-After': '3' }, 'retry-after')).toBe('3');
    expect(readHeader(undefined, 'request-id')).toBeUndefined();
    expect(readHeader(null, 'request-id')).toBeUndefined();
  });

  it('parseGraphErrorBody decodes the SDK JSON string, a wrapped object and raw bytes', () => {
    const inner = { code: 'BadRequest', message: 'nope' };
    expect(parseGraphErrorBody(JSON.stringify(inner))).toEqual(inner);
    expect(parseGraphErrorBody({ error: inner })).toEqual(inner);
    const bytes = Buffer.from(JSON.stringify({ error: inner }));
    expect(parseGraphErrorBody(bytes)).toEqual(inner);
    expect(parseGraphErrorBody(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength))).toEqual(inner);
  });

  it('parseGraphErrorBody returns null for anything that is not a Graph error body', () => {
    expect(parseGraphErrorBody('Error: fetch failed')).toBeNull();
    expect(parseGraphErrorBody(null)).toBeNull();
    expect(parseGraphErrorBody(undefined)).toBeNull();
    expect(parseGraphErrorBody(42)).toBeNull();
  });
});

describe('MCP error text for Graph failures', () => {
  const graphError = {
    status: 400,
    code: 'BadRequest',
    message: 'Something specific went wrong.',
    requestId: 'req-1',
    clientRequestId: 'cid-1',
    date: '2026-01-15T10:00:00',
    method: 'GET',
    url: 'https://graph.microsoft.com/v1.0/me/messages?$top=5'
  };

  it('formatGraphErrorText renders one verbatim line', () => {
    expect(formatGraphErrorText(graphError)).toBe(
      'HTTP 400 BadRequest: Something specific went wrong. ' +
      '[request-id: req-1; client-request-id: cid-1; date: 2026-01-15T10:00:00] ' +
      'GET https://graph.microsoft.com/v1.0/me/messages?$top=5'
    );
  });

  it('formatGraphErrorText omits what is unknown instead of printing placeholders', () => {
    expect(formatGraphErrorText({ status: 404, method: 'GET', url: 'https://graph.microsoft.com/v1.0/x' }))
      .toBe('HTTP 404: No error message returned GET https://graph.microsoft.com/v1.0/x');
    expect(formatGraphErrorText({ message: 'fetch failed' })).toBe('No HTTP response: fetch failed');
  });

  it('convertErrorToToolError keeps graphError and fetchDetails as structured details', () => {
    const error = new Error('could not fetch');
    error.fetchDetails = { attemptedRoutes: ['json', 'value'], messageId: 'm1', attachmentId: 'a1', attachmentName: null, graphError };

    const result = convertErrorToToolError(error, 'Failed to save attachment');

    expect(result.isError).toBe(true);
    expect(result.content[0].text).toBe('Failed to save attachment: could not fetch');
    expect(result._errorDetails.graphError).toEqual(graphError);
    expect(result._errorDetails.fetchDetails.attemptedRoutes).toEqual(['json', 'value']);
  });
});
