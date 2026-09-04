import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { mockGraphClient } from '../helpers/mockGraph.js';
import { listSharedMailboxesTool } from '../../tools/email/listSharedMailboxes.js';

const SELECT = 'id,displayName,totalItemCount,unreadItemCount';

// The global setup file clears MCP_OUTLOOK_SHARED_MAILBOX per test; the
// known-mailboxes setting is this suite's own business, so save/restore it.
let savedKnown;
beforeEach(() => {
  savedKnown = process.env.MCP_OUTLOOK_KNOWN_MAILBOXES;
  delete process.env.MCP_OUTLOOK_KNOWN_MAILBOXES;
});
afterEach(() => {
  if (savedKnown === undefined) delete process.env.MCP_OUTLOOK_KNOWN_MAILBOXES;
  else process.env.MCP_OUTLOOK_KNOWN_MAILBOXES = savedKnown;
});

const parse = (res) => JSON.parse(res.content[0].text);

const inbox = (extra = {}) => ({
  id: 'inbox',
  displayName: 'Inbox',
  totalItemCount: 128,
  unreadItemCount: 7,
  ...extra,
});

const mcpError = (text) => ({ isError: true, content: [{ type: 'text', text }] });

describe('listSharedMailboxesTool', () => {
  it('probes the inbox of each candidate and reports it accessible', async () => {
    const { authManager, makeRequest } = mockGraphClient(() => inbox());

    const res = await listSharedMailboxesTool(authManager, { candidates: ['Careers@Example.com'] });

    expect(res.isError).toBeUndefined();
    expect(makeRequest).toHaveBeenCalledTimes(1);
    expect(makeRequest).toHaveBeenCalledWith(
      '/users/careers%40example.com/mailFolders/inbox',
      { select: SELECT }
    );

    const data = parse(res);
    expect(data.mailboxes[0]).toMatchObject({
      address: 'careers@example.com',
      status: 'accessible',
      inbox: { totalItemCount: 128, unreadItemCount: 7 },
    });
    expect(data.accessible).toEqual(['careers@example.com']);
  });

  it('classifies accessible / access-denied / not-found candidates', async () => {
    const { authManager } = mockGraphClient((path) => {
      if (path.includes('ok%40example.com')) return inbox({ totalItemCount: 3, unreadItemCount: 1 });
      if (path.includes('denied%40example.com')) {
        return mcpError('Microsoft Graph API: Insufficient permissions to access this mailbox');
      }
      return mcpError('Resource not found for the segment');
    });

    const res = await listSharedMailboxesTool(authManager, {
      candidates: ['ok@example.com', 'denied@example.com', 'missing@example.com'],
    });

    expect(res.isError).toBeUndefined();
    const data = parse(res);
    expect(data.mailboxes.map(m => m.status)).toEqual(['accessible', 'access-denied', 'not-found']);
    expect(data.accessible).toEqual(['ok@example.com']);
    expect(data.mailboxes[1].note).toMatch(/Full Access/);
  });

  it('marks a malformed candidate invalid without calling Graph', async () => {
    const { authManager, makeRequest } = mockGraphClient(() => inbox());

    const res = await listSharedMailboxesTool(authManager, { candidates: ['not-an-address'] });

    expect(res.isError).toBeUndefined();
    expect(makeRequest).not.toHaveBeenCalled();
    const data = parse(res);
    expect(data.mailboxes[0]).toMatchObject({ address: 'not-an-address', status: 'invalid' });
    expect(data.mailboxes[0].note).toMatch(/Invalid mailbox address/);
    expect(data.accessible).toEqual([]);
  });

  it('gathers candidates from both env settings and deduplicates case-insensitively', async () => {
    process.env.MCP_OUTLOOK_KNOWN_MAILBOXES = 'hr@example.com, HR@example.com';
    process.env.MCP_OUTLOOK_SHARED_MAILBOX = 'hr@example.com';
    const { authManager, makeRequest } = mockGraphClient(() => inbox());

    const res = await listSharedMailboxesTool(authManager, { candidates: ['hr@example.com'] });

    expect(makeRequest).toHaveBeenCalledTimes(1);
    expect(makeRequest).toHaveBeenCalledWith(
      '/users/hr%40example.com/mailFolders/inbox',
      { select: SELECT }
    );

    const data = parse(res);
    expect(data.mailboxes).toHaveLength(1);
    expect(data.sources).toEqual({ candidatesArg: 1, knownMailboxesEnv: 2, sharedMailboxEnv: 1 });
  });

  it('returns an empty, non-error result with a hint when there are no candidates', async () => {
    const { authManager, makeRequest } = mockGraphClient(() => inbox());

    const res = await listSharedMailboxesTool(authManager, {});

    expect(res.isError).toBeUndefined();
    expect(makeRequest).not.toHaveBeenCalled();
    const data = parse(res);
    expect(data.mailboxes).toEqual([]);
    expect(data.accessible).toEqual([]);
    expect(data.sources).toEqual({ candidatesArg: 0, knownMailboxesEnv: 0, sharedMailboxEnv: 0 });
    expect(data.hint).toMatch(/Known Shared Mailboxes/);
    expect(data.hint).toMatch(/candidates/);
  });
});
