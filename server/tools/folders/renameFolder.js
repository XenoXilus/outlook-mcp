import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { getMailboxBase } from '../../graph/graphHelpers.js';

// Rename mail folder
export async function renameFolderTool(authManager, args) {
  const { folderId, newDisplayName, mailbox } = args;

  if (!folderId) {
    return createValidationError('folderId', 'Parameter is required');
  }

  if (!newDisplayName) {
    return createValidationError('newDisplayName', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    const mailboxBase = getMailboxBase(mailbox);

    await graphApiClient.makeRequest(`${mailboxBase}/mailFolders/${folderId}`, {
      body: { displayName: newDisplayName }
    }, 'PATCH');

    return {
      content: [
        {
          type: 'text',
          text: `Folder renamed to "${newDisplayName}" successfully. Folder ID: ${folderId}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to rename folder');
  }
}