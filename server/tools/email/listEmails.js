// List emails from a specific folder
import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';

export async function listEmailsTool(authManager, args) {
  const { folder = 'inbox', limit = 10, filter } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    const folderResolver = graphApiClient.getFolderResolver();
    
    // Resolve folder name to ID
    let folderId;
    try {
      folderId = await folderResolver.resolveFolderToId(folder);
    } catch (folderError) {
      return createValidationError('folder', folderError.message);
    }
    
    const options = {
      select: 'subject,from,receivedDateTime,bodyPreview,isRead',
      top: limit,
      orderby: 'receivedDateTime desc',
    };

    if (filter) {
      options.filter = filter;
    }

    const result = await graphApiClient.makeRequest(`/me/mailFolders/${folderId}/messages`, options);

    // Handle MCP error responses from makeRequest
    if (result.content && result.isError !== undefined) {
      return result;
    }

    const emails = result.value?.map(email => ({
      id: email.id,
      subject: email.subject,
      from: email.from?.emailAddress?.address || 'Unknown',
      fromName: email.from?.emailAddress?.name || 'Unknown',
      receivedDateTime: email.receivedDateTime,
      preview: email.bodyPreview,
      isRead: email.isRead,
    })) || [];

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ 
            folder: {
              name: folder,
              id: folderId
            },
            emails, 
            count: emails.length 
          }, null, 2),
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to list emails');
  }
}

// Get detailed information about a specific email
export async function getEmailTool(authManager, args) {
  console.error(`DEBUG getEmailTool: Called with args:`, JSON.stringify(args, null, 2));
  const { messageId } = args;

  if (!messageId) {
    console.error(`DEBUG getEmailTool: Missing messageId parameter`);
    return createValidationError('messageId', 'Parameter is required');
  }

  try {
    console.error(`DEBUG getEmailTool: Starting authentication for messageId: ${messageId}`);
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    
    const options = {
      select: 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,body,bodyPreview,importance,isRead,hasAttachments,attachments,conversationId'
    };

    console.error(`DEBUG getEmailTool: Making Graph API request for ${messageId}`);
    const email = await graphApiClient.makeRequest(`/me/messages/${messageId}`, options);
    console.error(`DEBUG getEmailTool: Got email response with subject: ${email?.subject || 'NO SUBJECT'}`);
    
    // Check if the response is already an MCP error
    if (email && email.content && email.isError !== undefined) {
      console.error(`DEBUG getEmailTool: Graph API returned MCP error:`, email);
      return email;
    }

    const emailData = {
      id: email.id,
      subject: email.subject,
      from: {
        address: email.from?.emailAddress?.address || 'Unknown',
        name: email.from?.emailAddress?.name || 'Unknown'
      },
      toRecipients: email.toRecipients?.map(r => ({
        address: r.emailAddress?.address,
        name: r.emailAddress?.name
      })) || [],
      ccRecipients: email.ccRecipients?.map(r => ({
        address: r.emailAddress?.address,
        name: r.emailAddress?.name
      })) || [],
      bccRecipients: email.bccRecipients?.map(r => ({
        address: r.emailAddress?.address,
        name: r.emailAddress?.name
      })) || [],
      receivedDateTime: email.receivedDateTime,
      sentDateTime: email.sentDateTime,
      body: {
        contentType: email.body?.contentType || 'Text',
        content: email.body?.content || ''
      },
      bodyPreview: email.bodyPreview,
      importance: email.importance,
      isRead: email.isRead,
      hasAttachments: email.hasAttachments,
      attachments: email.attachments?.map(a => ({
        id: a.id,
        name: a.name,
        contentType: a.contentType,
        size: a.size
      })) || [],
      conversationId: email.conversationId
    };

    console.error(`DEBUG getEmailTool: Built emailData structure, returning response`);
    const response = {
      content: [
        {
          type: 'text',
          text: JSON.stringify(emailData, null, 2),
        },
      ],
    };
    console.error(`DEBUG getEmailTool: Final response length: ${response.content[0].text.length} chars`);
    return response;
  } catch (error) {
    console.error(`DEBUG getEmailTool: Caught error:`, error);
    return convertErrorToToolError(error, 'Failed to get email');
  }
}