import { describe, it, expect } from 'vitest';
import { mockGraphClient, expectAllPathsUnder } from '../helpers/mockGraph.js';
import { listFoldersTool } from '../../tools/folders/listFolders.js';
import { createFolderTool } from '../../tools/folders/createFolder.js';
import { renameFolderTool } from '../../tools/folders/renameFolder.js';
import { getFolderStatsTool } from '../../tools/folders/getFolderStats.js';
import { listAttachmentsTool } from '../../tools/attachments/listAttachments.js';
import { scanAttachmentsTool } from '../../tools/attachments/scanAttachments.js';

const BASE = '/users/careers%40example.com';
const M = 'careers@example.com';

// Responder that satisfies every tool in this file: message listings return one
// message (so scanAttachments also fetches its attachments); everything else
// gets a folder shaped to exercise the childFolders follow-up in getFolderStats.
const responder = (path) => path.endsWith('/messages')
  ? { value: [{ id: 'm1', subject: 's', from: { emailAddress: { address: 'x@example.com' } } }] }
  : { value: [], id: 'f1', displayName: 'Screening', childFolderCount: 1, totalItemCount: 2, unreadItemCount: 1 };

const cases = [
  ['listFoldersTool', (a) => listFoldersTool(a, { mailbox: M })],
  ['createFolderTool', (a) => createFolderTool(a, { displayName: 'Screening', mailbox: M })],
  ['createFolderTool (nested)', (a) => createFolderTool(a, { displayName: 'Screening', parentFolderId: 'f0', mailbox: M })],
  ['renameFolderTool', (a) => renameFolderTool(a, { folderId: 'f1', newDisplayName: 'Screened', mailbox: M })],
  ['getFolderStatsTool', (a) => getFolderStatsTool(a, { folderId: 'f1', mailbox: M })],
  ['listAttachmentsTool', (a) => listAttachmentsTool(a, { messageId: 'm1', mailbox: M })],
  ['scanAttachmentsTool', (a) => scanAttachmentsTool(a, { folder: 'inbox', limit: 5, mailbox: M })],
];

describe('folder and attachment tools route to the per-call mailbox', () => {
  for (const [name, run] of cases) {
    it(`${name} targets ${BASE}`, async () => {
      const { authManager, makeRequest } = mockGraphClient(responder);
      const res = await run(authManager);
      expect(res.isError).toBeUndefined();
      expectAllPathsUnder(makeRequest, BASE);
    });
  }

  it('getFolderStatsTool routes the subfolder request too', async () => {
    const { authManager, makeRequest } = mockGraphClient(responder);
    await getFolderStatsTool(authManager, { folderId: 'f1', mailbox: M });
    const paths = makeRequest.mock.calls.map((c) => c[0]);
    expect(paths).toContain(`${BASE}/mailFolders/f1/childFolders`);
  });

  it('scanAttachmentsTool routes the per-message attachment request too', async () => {
    const { authManager, makeRequest } = mockGraphClient(responder);
    await scanAttachmentsTool(authManager, { folder: 'inbox', limit: 5, mailbox: M });
    const paths = makeRequest.mock.calls.map((c) => c[0]);
    expect(paths).toContain(`${BASE}/messages/m1/attachments`);
  });

  it('surfaces a malformed mailbox as a tool error without calling Graph', async () => {
    const { authManager, makeRequest } = mockGraphClient(responder);
    const res = await listAttachmentsTool(authManager, { messageId: 'm1', mailbox: 'nonsense' });
    expect(res.isError).toBe(true);
    expect(res.content[0].text).toMatch(/Invalid mailbox address/);
    expect(makeRequest).not.toHaveBeenCalled();
  });
});
