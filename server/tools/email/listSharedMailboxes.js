import { convertErrorToToolError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';
import { getMailboxBase } from '../../graph/graphHelpers.js';

const PROBE_SELECT = 'id,displayName,totalItemCount,unreadItemCount';

const ACCESS_DENIED = /insufficient permissions|access.*denied|forbidden/i;
const NOT_FOUND = /not found/i;

const ACCESS_DENIED_NOTE = 'You need Exchange Full Access to this mailbox, and your sign-in must have consented to the Mail.*.Shared scopes.';
const NO_CANDIDATES_HINT = 'Graph has no delegated API that lists the mailboxes you can open, so discovery probes candidates. Add addresses to the Known Shared Mailboxes setting (MCP_OUTLOOK_KNOWN_MAILBOXES), set the Shared Mailbox setting, or pass a `candidates` argument.';
const USAGE_HINT = 'Pass any accessible address as the `mailbox` argument on mail/folder/attachment/receipt tools.';
const NONE_ACCESSIBLE_HINT = `None of the candidates opened. ${ACCESS_DENIED_NOTE}`;

// Comma-separated setting -> trimmed, non-empty entries.
function splitList(value) {
  return (value || '')
    .split(',')
    .map(entry => entry.trim())
    .filter(Boolean);
}

// Text of an MCP error object returned (not thrown) by makeRequest.
function errorText(value) {
  if (!value || !Array.isArray(value.content)) return '';
  return value.content.map(part => part?.text).filter(Boolean).join(' ');
}

function isMcpError(value) {
  return Boolean(value && value.isError !== undefined && value.content);
}

function classifyFailure(address, text) {
  const note = text || 'Unknown error while probing the mailbox';
  if (ACCESS_DENIED.test(note)) return { address, status: 'access-denied', note: ACCESS_DENIED_NOTE };
  if (NOT_FOUND.test(note)) return { address, status: 'not-found', note };
  return { address, status: 'error', note };
}

/**
 * Discover which shared mailboxes the signed-in user can actually open.
 *
 * Microsoft Graph exposes no delegated API that enumerates the mailboxes a
 * user holds Full Access to, so discovery is candidate-based: collect
 * addresses from the arguments and the configured settings, then probe each
 * one with a cheap inbox read and report what came back.
 */
export async function listSharedMailboxesTool(authManager, args = {}) {
  try {
    const fromArgs = Array.isArray(args?.candidates)
      ? args.candidates.filter(entry => typeof entry === 'string').map(entry => entry.trim()).filter(Boolean)
      : [];
    const fromKnownEnv = splitList(process.env.MCP_OUTLOOK_KNOWN_MAILBOXES);
    const fromSharedEnv = splitList(process.env.MCP_OUTLOOK_SHARED_MAILBOX);

    const sources = {
      candidatesArg: fromArgs.length,
      knownMailboxesEnv: fromKnownEnv.length,
      sharedMailboxEnv: fromSharedEnv.length,
    };

    // Deduplicate case-insensitively, keeping first-seen order.
    const candidates = [];
    const seen = new Set();
    for (const entry of [...fromArgs, ...fromKnownEnv, ...fromSharedEnv]) {
      const address = entry.toLowerCase();
      if (seen.has(address)) continue;
      seen.add(address);
      candidates.push(address);
    }

    if (candidates.length === 0) {
      return createSafeResponse({
        mailboxes: [],
        accessible: [],
        sources,
        hint: NO_CANDIDATES_HINT,
      });
    }

    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const mailboxes = await Promise.all(candidates.map(async (address) => {
      let mailboxBase;
      try {
        mailboxBase = getMailboxBase(address);
      } catch (error) {
        return { address, status: 'invalid', note: error.message };
      }

      try {
        const result = await graphApiClient.makeRequest(`${mailboxBase}/mailFolders/inbox`, { select: PROBE_SELECT });

        // makeRequest RETURNS MCP error objects rather than throwing them.
        if (isMcpError(result)) return classifyFailure(address, errorText(result));

        // The probe reads the inbox folder, so its displayName is just
        // "Inbox" — reporting it as a mailbox name would mislead.
        return {
          address,
          status: 'accessible',
          inbox: {
            totalItemCount: result?.totalItemCount ?? 0,
            unreadItemCount: result?.unreadItemCount ?? 0,
          },
        };
      } catch (error) {
        return classifyFailure(address, isMcpError(error) ? errorText(error) : error.message);
      }
    }));

    const accessible = mailboxes.filter(entry => entry.status === 'accessible').map(entry => entry.address);
    // The permissions hint only makes sense when something was actually
    // probed; an all-invalid candidate list is an address problem, not an
    // access problem, and each entry's note already says so.
    const probed = mailboxes.some(entry => entry.status !== 'invalid');
    const failureHint = probed
      ? NONE_ACCESSIBLE_HINT
      : 'No valid mailbox addresses to probe — check the addresses (expected forms like careers@yourcompany.com).';

    return createSafeResponse({
      mailboxes,
      accessible,
      sources,
      hint: accessible.length > 0 ? USAGE_HINT : failureHint,
    });
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to list shared mailboxes');
  }
}
