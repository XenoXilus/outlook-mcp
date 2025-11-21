import { applyUserStyling } from '../common/sharedUtils.js';
import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';

// Create calendar event with Teams meeting support
export async function createEventTool(authManager, args) {
  const { subject, start, end, body = '', bodyType = 'text', location = '', attendees = [], isOnlineMeeting = false, onlineMeetingProvider = 'teamsForBusiness', recurrence, preserveUserStyling = true } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    // Apply user styling if enabled and body is provided
    let finalBody = body;
    let finalBodyType = bodyType;

    if (preserveUserStyling && finalBody) {
      const styledBody = await applyUserStyling(graphApiClient, finalBody, bodyType);
      finalBody = styledBody.content;
      finalBodyType = styledBody.type;
    }

    const event = {
      subject,
      start,
      end,
      body: {
        contentType: finalBodyType === 'html' ? 'HTML' : 'Text',
        content: finalBody,
      },
    };

    if (location) {
      event.location = {
        displayName: location,
      };
    }

    if (attendees.length > 0) {
      event.attendees = attendees.map(email => ({
        emailAddress: { address: email },
        type: 'required',
      }));
    }

    // Add Teams meeting support
    if (isOnlineMeeting) {
      event.isOnlineMeeting = true;
      event.onlineMeetingProvider = onlineMeetingProvider;
    }

    // Add recurrence support with validation
    if (recurrence) {
      if (!recurrence.pattern || !recurrence.range) {
        return createValidationError('recurrence', 'Recurrence must include both pattern and range');
      }

      const validPatternTypes = ['daily', 'weekly', 'absoluteMonthly', 'relativeMonthly', 'absoluteYearly', 'relativeYearly'];
      const validRangeTypes = ['endDate', 'noEnd', 'numbered'];

      if (!validPatternTypes.includes(recurrence.pattern.type)) {
        return createValidationError('recurrence.pattern.type', `Must be one of: ${validPatternTypes.join(', ')}`);
      }

      if (!validRangeTypes.includes(recurrence.range.type)) {
        return createValidationError('recurrence.range.type', `Must be one of: ${validRangeTypes.join(', ')}`);
      }

      event.recurrence = recurrence;
    }

    const result = await graphApiClient.postWithRetry('/me/events', event);

    const isRecurring = recurrence ? true : false;
    const meetingType = isOnlineMeeting ? 'Teams meeting' : 'Event';
    const recurrenceInfo = isRecurring ? ' (recurring)' : '';

    const successMessage = `${meetingType} "${subject}"${recurrenceInfo} created successfully. Event ID: ${result.id}` +
      (isOnlineMeeting && result.onlineMeeting?.joinUrl ? ` Join URL: ${result.onlineMeeting.joinUrl}` : '');

    return {
      content: [
        {
          type: 'text',
          text: successMessage,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to create event');
  }
}