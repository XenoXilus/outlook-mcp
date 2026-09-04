import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { getMailboxBase } from '../../graph/graphHelpers.js';

// Helper function to format file size
function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// Add attachment to message
export async function addAttachmentTool(authManager, args) {
  const { messageId, name, contentType, contentBytes, mailbox } = args;

  if (!messageId) {
    return createValidationError('messageId', 'Parameter is required');
  }

  if (!name) {
    return createValidationError('name', 'Parameter is required');
  }

  if (!contentType) {
    return createValidationError('contentType', 'Parameter is required');
  }

  if (!contentBytes) {
    return createValidationError('contentBytes', 'Parameter is required (base64 encoded)');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    const mailboxBase = getMailboxBase(mailbox);

    const attachmentData = {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: name,
      contentType: contentType,
      contentBytes: contentBytes
    };

    const result = await graphApiClient.postWithRetry(`${mailboxBase}/messages/${messageId}/attachments`, attachmentData);

    // Calculate approximate size from base64 content
    const estimatedSize = Math.floor(contentBytes.length * 0.75); // Base64 is ~33% larger than original

    return {
      content: [
        {
          type: 'text',
          text: `Attachment "${name}" added successfully. Attachment ID: ${result.id}. Estimated size: ${formatFileSize(estimatedSize)}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to add attachment');
  }
}