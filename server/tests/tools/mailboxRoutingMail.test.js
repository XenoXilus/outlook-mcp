import { describe, it, expect } from 'vitest';
import { mockGraphClient, expectAllPathsUnder } from '../helpers/mockGraph.js';
import { listEmailsTool, getEmailTool } from '../../tools/email/listEmails.js';
import {
  deleteEmailTool, moveEmailTool, markAsReadTool, flagEmailTool,
  categorizeEmailTool, archiveEmailTool, batchProcessEmailsTool
} from '../../tools/email/emailManagement.js';

const BASE = '/users/careers%40example.com';
const M = 'careers@example.com';

// Responder that satisfies every tool in this file: folder listings get an
// Archive/Deleted Items pair; message GETs get a minimal message.
const responder = (path) => path.endsWith('/mailFolders')
  ? { value: [{ id: 'arch', displayName: 'Archive' }, { id: 'del', displayName: 'Deleted Items' }] }
  : { value: [], id: 'm1', subject: 's', from: { emailAddress: { address: 'x@example.com' } } };

const cases = [
  ['listEmailsTool', (a) => listEmailsTool(a, { folder: 'inbox', mailbox: M })],
  ['getEmailTool', (a) => getEmailTool(a, { messageId: 'm1', mailbox: M })],
  ['deleteEmailTool (permanent)', (a) => deleteEmailTool(a, { messageId: 'm1', permanentDelete: true, mailbox: M })],
  ['deleteEmailTool (to Deleted Items)', (a) => deleteEmailTool(a, { messageId: 'm1', permanentDelete: false, mailbox: M })],
  ['moveEmailTool', (a) => moveEmailTool(a, { messageId: 'm1', destinationFolderId: 'Archive', mailbox: M })],
  ['markAsReadTool', (a) => markAsReadTool(a, { messageId: 'm1', isRead: true, mailbox: M })],
  ['flagEmailTool', (a) => flagEmailTool(a, { messageId: 'm1', flagStatus: 'flagged', mailbox: M })],
  ['categorizeEmailTool', (a) => categorizeEmailTool(a, { messageId: 'm1', categories: ['HR'], mailbox: M })],
  ['archiveEmailTool', (a) => archiveEmailTool(a, { messageId: 'm1', mailbox: M })],
  ['batchProcessEmailsTool (markAsRead)', (a) => batchProcessEmailsTool(a, { messageIds: ['m1'], operation: 'markAsRead', mailbox: M })],
  ['batchProcessEmailsTool (markAsUnread)', (a) => batchProcessEmailsTool(a, { messageIds: ['m1'], operation: 'markAsUnread', mailbox: M })],
  ['batchProcessEmailsTool (delete, permanent)', (a) => batchProcessEmailsTool(a, { messageIds: ['m1'], operation: 'delete', operationData: { permanentDelete: true }, mailbox: M })],
  ['batchProcessEmailsTool (delete, to Deleted Items)', (a) => batchProcessEmailsTool(a, { messageIds: ['m1'], operation: 'delete', mailbox: M })],
  ['batchProcessEmailsTool (move)', (a) => batchProcessEmailsTool(a, { messageIds: ['m1'], operation: 'move', operationData: { destinationFolderId: 'arch' }, mailbox: M })],
  ['batchProcessEmailsTool (flag)', (a) => batchProcessEmailsTool(a, { messageIds: ['m1'], operation: 'flag', mailbox: M })],
  ['batchProcessEmailsTool (categorize)', (a) => batchProcessEmailsTool(a, { messageIds: ['m1'], operation: 'categorize', operationData: { categories: ['HR'] }, mailbox: M })],
];

describe('mail read/manage tools route to the per-call mailbox', () => {
  for (const [name, run] of cases) {
    it(`${name} targets ${BASE}`, async () => {
      const { authManager, makeRequest } = mockGraphClient(responder);
      const res = await run(authManager);
      expect(res.isError).toBeUndefined();
      expectAllPathsUnder(makeRequest, BASE);
    });
  }

  it('listEmailsTool resolves folders in the target mailbox', async () => {
    const { authManager, client } = mockGraphClient(responder);
    await listEmailsTool(authManager, { folder: 'Applications', mailbox: M });
    expect(client.getFolderResolver).toHaveBeenCalledWith(BASE);
  });

  it('surfaces a malformed mailbox as a tool error without calling Graph', async () => {
    const { authManager, makeRequest } = mockGraphClient(responder);
    const res = await getEmailTool(authManager, { messageId: 'm1', mailbox: 'nonsense' });
    expect(res.isError).toBe(true);
    expect(res.content[0].text).toMatch(/Invalid mailbox address/);
    expect(makeRequest).not.toHaveBeenCalled();
  });
});
