import { applyUserStyling, clearStylingCache } from '../common/sharedUtils.js';
import { convertErrorToToolError } from '../../utils/mcpErrorResponse.js';
import { getMailboxBase } from '../../graph/graphHelpers.js';

// Send email with user styling
export async function sendEmailTool(authManager, args) {
  const { to, subject, body, bodyType = 'text', cc = [], bcc = [], preserveUserStyling = true, mailbox } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    const mailboxBase = getMailboxBase(mailbox);

    let finalBody = body;
    let finalBodyType = bodyType;

    // If preserving user styling, get user's default styling and signature
    if (preserveUserStyling) {
      const styledBody = await applyUserStyling(graphApiClient, body, bodyType);
      finalBody = styledBody.content;
      finalBodyType = styledBody.type;
    }

    const message = {
      subject,
      body: {
        contentType: finalBodyType === 'html' ? 'HTML' : 'Text',
        content: finalBody,
      },
      toRecipients: to.map(email => ({
        emailAddress: { address: email },
      })),
    };

    if (cc.length > 0) {
      message.ccRecipients = cc.map(email => ({
        emailAddress: { address: email },
      }));
    }

    if (bcc.length > 0) {
      message.bccRecipients = bcc.map(email => ({
        emailAddress: { address: email },
      }));
    }

    await graphApiClient.postWithRetry(`${mailboxBase}/sendMail`, {
      message,
      saveToSentItems: true,
    });

    // Invalidate styling cache after sending email (user might have changed styling)
    // Don't invalidate signature cache as frequently since signatures change less often.
    // The styling cache is keyed on the SIGNED-IN user, whose own mail settings and
    // signature supplied the styling, so this read stays on /me even when the message
    // was sent from a shared mailbox.
    try {
      const userInfo = await graphApiClient.makeRequest('/me', { select: 'id' });
      clearStylingCache(userInfo.id);
    } catch (error) {
      console.warn('Could not invalidate styling cache:', error.message);
    }

    return {
      content: [
        {
          type: 'text',
          text: `Email sent successfully to ${to.join(', ')}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to send email');
  }
}