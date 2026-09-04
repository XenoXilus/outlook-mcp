import { beforeEach } from 'vitest';

// Tests assert Graph paths against the default (/me) base, so an ambient
// MCP_OUTLOOK_SHARED_MAILBOX in the developer's shell must not leak in.
beforeEach(() => { delete process.env.MCP_OUTLOOK_SHARED_MAILBOX; });
