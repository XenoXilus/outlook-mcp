import { describe, it, expect, vi } from 'vitest';
import { GraphApiClient } from '../../graph/graphClient.js';

function clientWithFolders() {
  const client = new GraphApiClient({ tokenManager: { getAccessToken: async () => 't' } });
  client.makeRequest = vi.fn(async (path) => ({
    value: [{ id: `${path}-hr-id`, displayName: 'Applications', parentFolderId: null }],
  }));
  return client;
}

describe('mailbox-aware FolderResolver', () => {
  it('scopes folder listing to the given mailbox base', async () => {
    const client = clientWithFolders();
    const resolver = client.getFolderResolver('/users/careers%40example.com');
    const id = await resolver.resolveFolderToId('Applications');
    expect(client.makeRequest.mock.calls[0][0]).toBe('/users/careers%40example.com/mailFolders');
    expect(id).toBe('/users/careers%40example.com/mailFolders-hr-id');
  });

  it('keeps separate cached resolvers per mailbox base', async () => {
    const client = clientWithFolders();
    const own = client.getFolderResolver();
    const shared = client.getFolderResolver('/users/careers%40example.com');
    expect(own).not.toBe(shared);
    expect(client.getFolderResolver()).toBe(own); // same base → same instance
    await own.resolveFolderToId('Applications');
    await shared.resolveFolderToId('Applications');
    const paths = client.makeRequest.mock.calls.map(c => c[0]);
    expect(paths).toContain('/me/mailFolders');
    expect(paths).toContain('/users/careers%40example.com/mailFolders');
  });
});
