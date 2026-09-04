import { vi, expect } from 'vitest';

/**
 * Mock GraphApiClient + authManager for tool tests. Every request funnels
 * through `makeRequest` so tests can assert each Graph path in one place.
 * `responder(path, options, method)` supplies the response per request.
 */
export function mockGraphClient(responder = () => ({ value: [] })) {
  const makeRequest = vi.fn(async (path, options = {}, method = 'GET') => responder(path, options, method));
  const client = {
    makeRequest,
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
  return { authManager, client, makeRequest };
}

/** Assert every Graph call in `makeRequest` targeted the given base path. */
export function expectAllPathsUnder(makeRequest, base) {
  expect(makeRequest.mock.calls.length).toBeGreaterThan(0);
  for (const call of makeRequest.mock.calls) {
    expect(call[0].startsWith(base + '/')).toBe(true);
  }
}
