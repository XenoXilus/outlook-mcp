import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';
import { getMailboxBase } from '../../graph/graphHelpers.js';

// List mail folders
export async function listFoldersTool(authManager, args) {
  const { includeHidden = false, includeChildFolders = true, top = 100, mailbox } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    const mailboxBase = getMailboxBase(mailbox);

    const options = {
      select: 'id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount,isHidden',
      top: Math.min(top, 1000)
    };

    if (!includeHidden) {
      options.filter = 'isHidden eq false';
    }

    let endpoint = `${mailboxBase}/mailFolders`;
    if (includeChildFolders) {
      endpoint = `${mailboxBase}/mailFolders?includeNestedFolders=true`;
    }

    const result = await graphApiClient.makeRequest(endpoint, options);

    const folders = result.value?.map(folder => ({
      id: folder.id,
      name: folder.displayName,
      parentFolderId: folder.parentFolderId,
      childFolderCount: folder.childFolderCount || 0,
      unreadItemCount: folder.unreadItemCount || 0,
      totalItemCount: folder.totalItemCount || 0,
      isHidden: folder.isHidden || false
    })) || [];

    return createSafeResponse({ folders, count: folders.length });
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to list folders');
  }
}