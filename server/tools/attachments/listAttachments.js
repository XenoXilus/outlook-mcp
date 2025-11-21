import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';

// Helper function to format file size
function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// List attachments for a message
export async function listAttachmentsTool(authManager, args) {
  const { messageId } = args;

  if (!messageId) {
    return createValidationError('messageId', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const result = await graphApiClient.makeRequest(`/me/messages/${messageId}/attachments`, {
      select: 'id,name,contentType,size,isInline,lastModifiedDateTime'
    });

    const attachments = result.value?.map(attachment => ({
      id: attachment.id,
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size,
      sizeFormatted: formatFileSize(attachment.size),
      isInline: attachment.isInline || false,
      lastModifiedDateTime: attachment.lastModifiedDateTime
    })) || [];

    const summary = {
      messageId,
      totalAttachments: attachments.length,
      totalSize: attachments.reduce((sum, att) => sum + (att.size || 0), 0),
      attachments
    };

    summary.totalSizeFormatted = formatFileSize(summary.totalSize);

    return createSafeResponse(summary);
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to list attachments');
  }
}