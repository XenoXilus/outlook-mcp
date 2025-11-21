import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';

// Get folder statistics
export async function getFolderStatsTool(authManager, args) {
  const { folderId, includeSubfolders = true } = args;

  if (!folderId) {
    return createValidationError('folderId', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    // Get folder details
    const folder = await graphApiClient.makeRequest(`/me/mailFolders/${folderId}`, {
      select: 'id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount,isHidden'
    });

    const stats = {
      id: folder.id,
      name: folder.displayName,
      totalItems: folder.totalItemCount || 0,
      unreadItems: folder.unreadItemCount || 0,
      readItems: (folder.totalItemCount || 0) - (folder.unreadItemCount || 0),
      childFolders: folder.childFolderCount || 0,
      isHidden: folder.isHidden || false,
      parentFolderId: folder.parentFolderId
    };

    // Get subfolder stats if requested
    if (includeSubfolders && stats.childFolders > 0) {
      try {
        const childFolders = await graphApiClient.makeRequest(`/me/mailFolders/${folderId}/childFolders`, {
          select: 'id,displayName,unreadItemCount,totalItemCount'
        });

        stats.subfolders = childFolders.value?.map(subfolder => ({
          id: subfolder.id,
          name: subfolder.displayName,
          totalItems: subfolder.totalItemCount || 0,
          unreadItems: subfolder.unreadItemCount || 0,
          readItems: (subfolder.totalItemCount || 0) - (subfolder.unreadItemCount || 0)
        })) || [];
      } catch (error) {
        stats.subfolders = [];
        stats.subfolderError = error.message;
      }
    }

    return createSafeResponse(stats);
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to get folder stats');
  }
}