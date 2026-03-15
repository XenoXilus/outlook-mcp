import { applyUserStyling } from '../common/sharedUtils.js';
import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';

// Create draft email with user styling
export async function createDraftTool(authManager, args) {
  const { to, subject, body, bodyType = 'text', cc = [], bcc = [], importance = 'normal', preserveUserStyling = true, replyToMessageId } = args;

  // When not replying, to and subject are required
  if (!replyToMessageId) {
    if (!to || to.length === 0) {
      return createValidationError('to', 'At least one recipient is required');
    }
    if (!subject) {
      return createValidationError('subject', 'Subject is required');
    }
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    // Apply user styling if enabled
    let finalBody = body || '';
    let finalBodyType = bodyType;

    if (preserveUserStyling && finalBody) {
      const styledBody = await applyUserStyling(graphApiClient, finalBody, bodyType);
      finalBody = styledBody.content;
      finalBodyType = styledBody.type;
    }

    let result;

    if (replyToMessageId) {
      // Create a reply draft preserving thread/conversation context
      const replyPayload = {};
      if (finalBody) {
        replyPayload.message = {
          body: {
            contentType: finalBodyType === 'html' ? 'HTML' : 'Text',
            content: finalBody,
          },
        };
        // Allow overriding additional recipients if explicitly provided
        if (to && to.length > 0) {
          replyPayload.message.toRecipients = to.map(email => ({
            emailAddress: { address: email },
          }));
        }
        if (cc.length > 0) {
          replyPayload.message.ccRecipients = cc.map(email => ({
            emailAddress: { address: email },
          }));
        }
        if (bcc.length > 0) {
          replyPayload.message.bccRecipients = bcc.map(email => ({
            emailAddress: { address: email },
          }));
        }
      }
      result = await graphApiClient.postWithRetry(`/me/messages/${replyToMessageId}/createReply`, replyPayload);
    } else {
      // Create a brand-new draft message
      const draft = {
        subject,
        body: {
          contentType: finalBodyType === 'html' ? 'HTML' : 'Text',
          content: finalBody,
        },
        toRecipients: to.map(email => ({
          emailAddress: { address: email },
        })),
        importance,
      };

      if (cc.length > 0) {
        draft.ccRecipients = cc.map(email => ({
          emailAddress: { address: email },
        }));
      }

      if (bcc.length > 0) {
        draft.bccRecipients = bcc.map(email => ({
          emailAddress: { address: email },
        }));
      }

      result = await graphApiClient.postWithRetry('/me/messages', draft);
    }

    return {
      content: [
        {
          type: 'text',
          text: `Draft created successfully. Draft ID: ${result.id}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to create draft');
  }
}