import { applyUserStyling } from '../common/sharedUtils.js';
import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { getMailboxBase } from '../../graph/graphHelpers.js';

// Reply to an email
export async function replyToEmailTool(authManager, args) {
  const { messageId, body, bodyType = 'text', comment = '', preserveUserStyling = true, mailbox } = args;

  if (!messageId) {
    return createValidationError('messageId', 'Parameter is required');
  }

  if (!body && !comment) {
    return createValidationError('body/comment', 'Either body or comment is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    const mailboxBase = getMailboxBase(mailbox);

    const replyPayload = {};

    // Use body or comment as the reply message text
    const replyText = body || comment;
    if (replyText) {
      if (preserveUserStyling) {
        const styledBody = await applyUserStyling(graphApiClient, replyText, bodyType);
        replyPayload.message = {
          body: {
            contentType: styledBody.type === 'html' ? 'HTML' : 'Text',
            content: styledBody.content,
          },
        };
      } else {
        replyPayload.message = {
          body: {
            contentType: bodyType === 'html' ? 'HTML' : 'Text',
            content: replyText,
          },
        };
      }
    }

    const result = await graphApiClient.postWithRetry(`${mailboxBase}/messages/${messageId}/reply`, replyPayload);

    return {
      content: [
        {
          type: 'text',
          text: `Reply created successfully. Reply ID: ${result.id || 'N/A'}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to reply to email');
  }
}

// Reply all to an email
export async function replyAllTool(authManager, args) {
  const { messageId, body, bodyType = 'text', comment = '', preserveUserStyling = true, mailbox } = args;

  if (!messageId) {
    return createValidationError('messageId', 'Parameter is required');
  }

  if (!body && !comment) {
    return createValidationError('body/comment', 'Either body or comment is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    const mailboxBase = getMailboxBase(mailbox);

    const replyPayload = {};

    // Use body or comment as the reply message text
    const replyText = body || comment;
    if (replyText) {
      if (preserveUserStyling) {
        const styledBody = await applyUserStyling(graphApiClient, replyText, bodyType);
        replyPayload.message = {
          body: {
            contentType: styledBody.type === 'html' ? 'HTML' : 'Text',
            content: styledBody.content,
          },
        };
      } else {
        replyPayload.message = {
          body: {
            contentType: bodyType === 'html' ? 'HTML' : 'Text',
            content: replyText,
          },
        };
      }
    }

    const result = await graphApiClient.postWithRetry(`${mailboxBase}/messages/${messageId}/replyAll`, replyPayload);

    return {
      content: [
        {
          type: 'text',
          text: `Reply all created successfully. Reply ID: ${result.id || 'N/A'}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to reply all to email');
  }
}