import { vi, expect } from 'vitest';
import { createToolError, formatGraphErrorText } from '../../utils/mcpErrorResponse.js';

/**
 * Mock GraphApiClient + authManager for tool tests. Every request funnels
 * through `makeRequest` so tests can assert each Graph path in one place.
 * `responder(path, options, method)` supplies the response per request.
 * `binaryResponder(path, options)` supplies raw-byte (`getBinary`) responses;
 * without one, binary requests fail like an unmocked route would.
 */
export function mockGraphClient(responder = () => ({ value: [] }), binaryResponder) {
  const makeRequest = vi.fn(async (path, options = {}, method = 'GET') => responder(path, options, method));
  const getBinary = vi.fn(async (path, options = {}) => (binaryResponder
    ? binaryResponder(path, options)
    : mockGraphErrorResult({ status: 404, code: 'ErrorItemNotFound', message: `binary route not mocked: ${path}` })));
  const client = {
    makeRequest,
    getBinary,
    postWithRetry: vi.fn(async (path, body) => makeRequest(path, { body }, 'POST')),
    patchWithRetry: vi.fn(async (path, body) => makeRequest(path, { body }, 'PATCH')),
    deleteWithRetry: vi.fn(async (path) => makeRequest(path, {}, 'DELETE')),
    getFolderResolver: vi.fn(() => ({
      resolveFolderToId: vi.fn(async (f) => f),
      resolveFoldersToIds: vi.fn(async (list) => list),
      getFolderInfo: vi.fn(async (f) => ({ id: f, displayName: f })),
      listAllFolders: vi.fn(async () => []),
    })),
  };
  const authManager = {
    ensureAuthenticated: vi.fn().mockResolvedValue(true),
    getGraphApiClient: vi.fn().mockReturnValue(client),
  };
  return { authManager, client, makeRequest, getBinary };
}

/**
 * The MCP error object GraphApiClient returns for a failed request: verbatim
 * diagnostics in the text and the structured copy in `_errorDetails.graphError`.
 */
export function mockGraphErrorResult({
  status = 400,
  code = 'BadRequest',
  message = 'Bad request',
  requestId = 'req-0001',
  clientRequestId = 'client-0001',
  date = '2026-01-01T00:00:00',
  method = 'GET',
  url = 'https://graph.microsoft.com/v1.0/me/messages/m1/attachments/a1'
} = {}) {
  const graphError = { status, code, message, requestId, clientRequestId, date, method, url };
  return createToolError(`Microsoft Graph API: ${formatGraphErrorText(graphError)}`, { statusCode: status, code, graphError });
}

/**
 * Assert every Graph call in `makeRequest` targeted the given base path.
 * `ignore` lists exact paths to exclude — for reads that are deliberately
 * scoped to the signed-in user (personal mail settings, signature/styling)
 * and therefore never routed to a shared mailbox.
 */
export function expectAllPathsUnder(makeRequest, base, { ignore = [] } = {}) {
  const calls = makeRequest.mock.calls.filter((call) => !ignore.includes(call[0]));
  expect(calls.length).toBeGreaterThan(0);
  for (const call of calls) {
    expect(call[0].startsWith(base + '/')).toBe(true);
  }
}

/**
 * Replace global fetch with a fake Graph service so the real GraphApiClient
 * (and the Graph SDK underneath it) runs end to end. `handler(request, url)`
 * returns a fetch Response; `request` is { url, method, headers }. Call
 * `restore()` in afterEach.
 */
export function installFakeGraphFetch(handler) {
  const requests = [];
  const original = globalThis.fetch;
  globalThis.fetch = vi.fn(async (input, init = {}) => {
    const url = typeof input === 'string' ? input : input.url;
    const record = {
      url,
      method: String(init.method || (typeof input === 'string' ? 'GET' : input.method) || 'GET').toUpperCase(),
      headers: new Headers(init.headers || (typeof input === 'string' ? undefined : input.headers))
    };
    requests.push(record);
    return handler(record, new URL(url));
  });
  return { requests, restore: () => { globalThis.fetch = original; } };
}

/** A Graph JSON response (success or `{ error: {...} }` body). */
export function graphJsonResponse(status, body, headers = {}) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json', ...headers }
  });
}

/** A raw-bytes response, as served by an attachment's /$value endpoint. */
export function graphBinaryResponse(bytes, headers = {}) {
  return new Response(bytes, { status: 200, headers: { 'content-type': 'application/pdf', ...headers } });
}
