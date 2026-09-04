import { describe, it, expect } from 'vitest';
import { mockGraphClient, expectAllPathsUnder } from '../helpers/mockGraph.js';
import { sendEmailTool } from '../../tools/email/sendEmail.js';
import { replyToEmailTool, replyAllTool } from '../../tools/email/replyEmail.js';
import { forwardEmailTool } from '../../tools/email/forwardEmail.js';
import { createDraftTool } from '../../tools/email/createDraft.js';
import { addAttachmentTool } from '../../tools/attachments/addAttachment.js';

const BASE = '/users/careers%40example.com';
const M = 'careers@example.com';
const responder = () => ({ id: 'd1', webLink: 'https://outlook.example/x', value: [] });

// Signature/styling lookups read the SIGNED-IN user's own mail settings
// (/me/mailboxSettings, /me sent items, /me?select=id for the cache key). Those are
// personal settings, not mailbox data, so they intentionally stay on /me and are not
// routed. Every case below passes `preserveUserStyling: false` to keep them out of the
// picture; sendEmailTool additionally reads `/me` unconditionally to key its styling
// cache, so that one path is excluded from the assertion for the send case only.
const PERSONAL_SETTINGS_PATHS = ['/me'];

const cases = [
  ['sendEmailTool', (a) => sendEmailTool(a, { to: ['p@example.com'], subject: 'Offer', body: 'Hi', preserveUserStyling: false, mailbox: M }), PERSONAL_SETTINGS_PATHS],
  ['replyToEmailTool', (a) => replyToEmailTool(a, { messageId: 'm1', body: 'Thanks', preserveUserStyling: false, mailbox: M })],
  ['replyAllTool', (a) => replyAllTool(a, { messageId: 'm1', body: 'Thanks all', preserveUserStyling: false, mailbox: M })],
  ['forwardEmailTool', (a) => forwardEmailTool(a, { messageId: 'm1', to: ['p@example.com'], preserveUserStyling: false, mailbox: M })],
  ['createDraftTool', (a) => createDraftTool(a, { to: ['p@example.com'], subject: 'Offer', body: 'Hi', preserveUserStyling: false, mailbox: M })],
  ['addAttachmentTool', (a) => addAttachmentTool(a, { messageId: 'd1', name: 'cv.pdf', contentType: 'application/pdf', contentBytes: 'JVBERg==', mailbox: M })],
];

describe('send-side tools route to the per-call mailbox', () => {
  for (const [name, run, ignore = []] of cases) {
    it(`${name} targets ${BASE}`, async () => {
      const { authManager, makeRequest } = mockGraphClient(responder);
      const res = await run(authManager);
      expect(res.isError).toBeUndefined();
      expectAllPathsUnder(makeRequest, BASE, { ignore });
    });
  }

  it('surfaces a malformed mailbox as a tool error without calling Graph', async () => {
    const { authManager, makeRequest } = mockGraphClient(responder);
    const res = await sendEmailTool(authManager, {
      to: ['p@example.com'], subject: 'Offer', body: 'Hi', preserveUserStyling: false, mailbox: 'nonsense',
    });
    expect(res.isError).toBe(true);
    expect(res.content[0].text).toMatch(/Invalid mailbox address/);
    expect(makeRequest).not.toHaveBeenCalled();
  });
});
