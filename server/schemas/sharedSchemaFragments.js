/**
 * Schema fragments shared by many tool schemas.
 */
export const mailboxProperty = {
  mailbox: {
    type: 'string',
    description: 'Optional shared mailbox to operate on (e.g. careers@yourcompany.com). Needs delegated Full Access to that mailbox. Defaults to the Shared Mailbox setting, then your own mailbox.',
  },
};
