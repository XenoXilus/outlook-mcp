import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';

// Email management operations (move, delete, flag, categorize, archive, batch processing)

// Delete email (soft delete to Deleted Items or permanent delete)
export async function deleteEmailTool(authManager, args) {
  const { messageId, permanentDelete = false } = args;

  if (!messageId) {
    return createValidationError('messageId', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    if (permanentDelete) {
      // Permanently delete the email
      await graphApiClient.makeRequest(`/me/messages/${messageId}`, {}, 'DELETE');
      
      return {
        content: [
          {
            type: 'text',
            text: `Email permanently deleted successfully. Message ID: ${messageId}`,
          },
        ],
      };
    } else {
      // Move to Deleted Items folder (soft delete)
      // First get the Deleted Items folder ID
      const foldersResult = await graphApiClient.makeRequest('/me/mailFolders', {
        filter: "displayName eq 'Deleted Items'"
      });
      
      let deletedItemsFolderId = 'deleteditems'; // Default fallback
      if (foldersResult.value && foldersResult.value.length > 0) {
        deletedItemsFolderId = foldersResult.value[0].id;
      }

      // Move the message to Deleted Items
      await graphApiClient.postWithRetry(`/me/messages/${messageId}/move`, {
        destinationId: deletedItemsFolderId
      });

      return {
        content: [
          {
            type: 'text',
            text: `Email moved to Deleted Items successfully. Message ID: ${messageId}`,
          },
        ],
      };
    }
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to delete email');
  }
}

// Move email to a specific folder
export async function moveEmailTool(authManager, args) {
  const { messageId, destinationFolderId } = args;

  if (!messageId) {
    return createValidationError('messageId', 'Parameter is required');
  }

  if (!destinationFolderId) {
    return createValidationError('destinationFolderId', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const result = await graphApiClient.postWithRetry(`/me/messages/${messageId}/move`, {
      destinationId: destinationFolderId
    });

    return {
      content: [
        {
          type: 'text',
          text: `Email moved successfully. New Message ID: ${result.id}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to move email');
  }
}

// Mark email as read or unread
export async function markAsReadTool(authManager, args) {
  const { messageId, isRead = true } = args;

  if (!messageId) {
    return createValidationError('messageId', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    await graphApiClient.makeRequest(`/me/messages/${messageId}`, {
      body: { isRead: isRead }
    }, 'PATCH');

    return {
      content: [
        {
          type: 'text',
          text: `Email ${isRead ? 'marked as read' : 'marked as unread'} successfully. Message ID: ${messageId}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, `Failed to mark email as ${isRead ? 'read' : 'unread'}`);
  }
}

// Flag email
export async function flagEmailTool(authManager, args) {
  const { messageId, flagStatus = 'flagged' } = args;

  if (!messageId) {
    return createValidationError('messageId', 'Parameter is required');
  }

  if (!['notFlagged', 'complete', 'flagged'].includes(flagStatus)) {
    return createValidationError('flagStatus', 'Must be one of: notFlagged, complete, flagged');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    await graphApiClient.makeRequest(`/me/messages/${messageId}`, {
      body: {
        flag: {
          flagStatus: flagStatus
        }
      }
    }, 'PATCH');

    return {
      content: [
        {
          type: 'text',
          text: `Email flag status set to '${flagStatus}' successfully. Message ID: ${messageId}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to flag email');
  }
}

// Categorize email
export async function categorizeEmailTool(authManager, args) {
  const { messageId, categories = [] } = args;

  if (!messageId) {
    return createValidationError('messageId', 'Parameter is required');
  }

  if (!Array.isArray(categories)) {
    return createValidationError('categories', 'Must be an array');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    await graphApiClient.makeRequest(`/me/messages/${messageId}`, {
      body: { categories: categories }
    }, 'PATCH');

    return {
      content: [
        {
          type: 'text',
          text: `Email categories updated successfully. Message ID: ${messageId}, Categories: ${categories.join(', ') || 'None'}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to categorize email');
  }
}

// Archive email
export async function archiveEmailTool(authManager, args) {
  const { messageId } = args;

  if (!messageId) {
    return createValidationError('messageId', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    // First try to find the Archive folder
    const foldersResult = await graphApiClient.makeRequest('/me/mailFolders', {
      filter: "displayName eq 'Archive'"
    });
    
    let archiveFolderId = 'archive'; // Default fallback
    if (foldersResult.value && foldersResult.value.length > 0) {
      archiveFolderId = foldersResult.value[0].id;
    }

    // Move the message to Archive
    const result = await graphApiClient.postWithRetry(`/me/messages/${messageId}/move`, {
      destinationId: archiveFolderId
    });

    return {
      content: [
        {
          type: 'text',
          text: `Email archived successfully. New Message ID: ${result.id}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to archive email');
  }
}

// Batch process emails
export async function batchProcessEmailsTool(authManager, args) {
  const { messageIds, operation, operationData = {} } = args;

  if (!messageIds || !Array.isArray(messageIds) || messageIds.length === 0) {
    return createValidationError('messageIds', 'Array is required and must not be empty');
  }

  if (!operation) {
    return createValidationError('operation', 'Parameter is required');
  }

  const validOperations = ['markAsRead', 'markAsUnread', 'delete', 'move', 'flag', 'categorize'];
  if (!validOperations.includes(operation)) {
    return createValidationError('operation', `Must be one of: ${validOperations.join(', ')}`);
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const results = [];
    const errors = [];

    // Process each message (could be optimized with batch requests in the future)
    for (const messageId of messageIds) {
      try {
        let result;
        
        switch (operation) {
          case 'markAsRead':
            await graphApiClient.makeRequest(`/me/messages/${messageId}`, { body: { isRead: true } }, 'PATCH');
            result = { messageId, status: 'success', operation: 'marked as read' };
            break;
            
          case 'markAsUnread':
            await graphApiClient.makeRequest(`/me/messages/${messageId}`, { body: { isRead: false } }, 'PATCH');
            result = { messageId, status: 'success', operation: 'marked as unread' };
            break;
            
          case 'delete':
            if (operationData.permanentDelete) {
              await graphApiClient.makeRequest(`/me/messages/${messageId}`, {}, 'DELETE');
              result = { messageId, status: 'success', operation: 'permanently deleted' };
            } else {
              // Find Deleted Items folder
              const foldersResult = await graphApiClient.makeRequest('/me/mailFolders', {
                filter: "displayName eq 'Deleted Items'"
              });
              let deletedItemsFolderId = 'deleteditems';
              if (foldersResult.value && foldersResult.value.length > 0) {
                deletedItemsFolderId = foldersResult.value[0].id;
              }
              await graphApiClient.postWithRetry(`/me/messages/${messageId}/move`, {
                destinationId: deletedItemsFolderId
              });
              result = { messageId, status: 'success', operation: 'moved to deleted items' };
            }
            break;
            
          case 'move':
            if (!operationData.destinationFolderId) {
              return createValidationError('destinationFolderId', 'Required for move operation');
            }
            await graphApiClient.postWithRetry(`/me/messages/${messageId}/move`, {
              destinationId: operationData.destinationFolderId
            });
            result = { messageId, status: 'success', operation: `moved to folder ${operationData.destinationFolderId}` };
            break;
            
          case 'flag':
            const flagStatus = operationData.flagStatus || 'flagged';
            await graphApiClient.makeRequest(`/me/messages/${messageId}`, {
              body: { flag: { flagStatus } }
            }, 'PATCH');
            result = { messageId, status: 'success', operation: `flagged as ${flagStatus}` };
            break;
            
          case 'categorize':
            const categories = operationData.categories || [];
            await graphApiClient.makeRequest(`/me/messages/${messageId}`, {
              body: { categories }
            }, 'PATCH');
            result = { messageId, status: 'success', operation: `categorized as ${categories.join(', ')}` };
            break;
        }
        
        results.push(result);
      } catch (error) {
        errors.push({ messageId, error: error.message });
      }
    }

    const summary = {
      totalProcessed: messageIds.length,
      successful: results.length,
      failed: errors.length,
      operation,
      results,
      errors
    };

    return createSafeResponse(summary);
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to batch process emails');
  }
}