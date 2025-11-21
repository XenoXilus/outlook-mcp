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

// Scan attachments for security risks
export async function scanAttachmentsTool(authManager, args) {
  const { 
    folder = 'inbox', 
    maxSizeMB = 10, 
    suspiciousTypes = ['exe', 'bat', 'cmd', 'scr', 'vbs', 'js'],
    limit = 100,
    daysBack = 30
  } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    // Calculate date filter
    const sinceDate = new Date();
    sinceDate.setDate(sinceDate.getDate() - daysBack);

    const options = {
      select: 'id,subject,from,receivedDateTime,hasAttachments',
      filter: `hasAttachments eq true and receivedDateTime ge ${sinceDate.toISOString()}`,
      top: Math.min(limit, 1000),
      orderby: 'receivedDateTime desc'
    };

    const emailsResult = await graphApiClient.makeRequest(`/me/mailFolders/${folder}/messages`, options);

    const suspiciousEmails = [];
    const largeAttachments = [];
    const scanSummary = {
      totalEmailsScanned: emailsResult.value?.length || 0,
      suspiciousAttachments: 0,
      largeAttachments: 0,
      totalAttachments: 0,
      folder,
      daysBack,
      maxSizeMB,
      suspiciousTypes: suspiciousTypes.join(', ')
    };

    // Process each email with attachments
    for (const email of emailsResult.value || []) {
      try {
        const attachmentsResult = await graphApiClient.makeRequest(`/me/messages/${email.id}/attachments`, {
          select: 'id,name,contentType,size,isInline'
        });

        const attachments = attachmentsResult.value || [];
        scanSummary.totalAttachments += attachments.length;

        for (const attachment of attachments) {
          const sizeInMB = (attachment.size || 0) / (1024 * 1024);
          
          // Check for suspicious file types
          const extension = attachment.name?.split('.').pop()?.toLowerCase();
          if (extension && suspiciousTypes.includes(extension)) {
            suspiciousEmails.push({
              messageId: email.id,
              subject: email.subject,
              from: email.from?.emailAddress?.address || 'Unknown',
              receivedDateTime: email.receivedDateTime,
              attachment: {
                id: attachment.id,
                name: attachment.name,
                contentType: attachment.contentType,
                size: attachment.size,
                sizeFormatted: formatFileSize(attachment.size),
                extension,
                reason: 'Suspicious file type'
              }
            });
            scanSummary.suspiciousAttachments++;
          }

          // Check for large attachments
          if (sizeInMB > maxSizeMB) {
            largeAttachments.push({
              messageId: email.id,
              subject: email.subject,
              from: email.from?.emailAddress?.address || 'Unknown',
              receivedDateTime: email.receivedDateTime,
              attachment: {
                id: attachment.id,
                name: attachment.name,
                contentType: attachment.contentType,
                size: attachment.size,
                sizeFormatted: formatFileSize(attachment.size),
                sizeMB: sizeInMB.toFixed(2),
                reason: `Exceeds ${maxSizeMB}MB limit`
              }
            });
            scanSummary.largeAttachments++;
          }
        }
      } catch (error) {
        console.warn(`Could not scan attachments for message ${email.id}:`, error.message);
      }
    }

    const scanResults = {
      summary: scanSummary,
      suspiciousAttachments: suspiciousEmails,
      largeAttachments: largeAttachments
    };

    return createSafeResponse(scanResults);
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to scan attachments');
  }
}