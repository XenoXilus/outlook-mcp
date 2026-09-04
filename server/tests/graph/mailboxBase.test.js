import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { getMailboxBase } from '../../graph/graphHelpers.js';
import { searchEmailsTool } from '../../tools/email/searchEmails.js';

describe('getMailboxBase', () => {
  let saved;
  beforeEach(() => { saved = process.env.MCP_OUTLOOK_SHARED_MAILBOX; });
  afterEach(() => {
    if (saved === undefined) delete process.env.MCP_OUTLOOK_SHARED_MAILBOX;
    else process.env.MCP_OUTLOOK_SHARED_MAILBOX = saved;
  });

  it('defaults to /me', () => {
    delete process.env.MCP_OUTLOOK_SHARED_MAILBOX;
    expect(getMailboxBase()).toBe('/me');
  });

  it('uses /users/<mailbox> when MCP_OUTLOOK_SHARED_MAILBOX is set', () => {
    process.env.MCP_OUTLOOK_SHARED_MAILBOX = 'finance@example.com';
    expect(getMailboxBase()).toBe('/users/finance%40example.com');
  });

  it('searchEmailsTool targets the shared mailbox', async () => {
    process.env.MCP_OUTLOOK_SHARED_MAILBOX = 'finance@example.com';
    const makeRequest = vi.fn().mockResolvedValue({ value: [] });
    const authManager = {
      ensureAuthenticated: vi.fn().mockResolvedValue(true),
      getGraphApiClient: vi.fn().mockReturnValue({ makeRequest, getFolderResolver: vi.fn() })
    };
    await searchEmailsTool(authManager, { from: 'no-reply@globex.example' });
    expect(makeRequest.mock.calls[0][0]).toBe('/users/finance%40example.com/messages');
  });

  it('prefers an explicit mailbox argument over the env setting', () => {
    process.env.MCP_OUTLOOK_SHARED_MAILBOX = 'finance@example.com';
    expect(getMailboxBase('careers@example.com')).toBe('/users/careers%40example.com');
  });

  it('falls back to env, then /me, when no argument is given', () => {
    process.env.MCP_OUTLOOK_SHARED_MAILBOX = 'finance@example.com';
    expect(getMailboxBase()).toBe('/users/finance%40example.com');
    delete process.env.MCP_OUTLOOK_SHARED_MAILBOX;
    expect(getMailboxBase()).toBe('/me');
    expect(getMailboxBase('')).toBe('/me');
  });

  it('rejects a malformed explicit mailbox', () => {
    expect(() => getMailboxBase('not-an-address')).toThrow(/Invalid mailbox address/);
    expect(() => getMailboxBase('a b@example.com')).toThrow(/Invalid mailbox address/);
    expect(() => getMailboxBase('x@ex/ample.com')).toThrow(/Invalid mailbox address/);
  });
});
