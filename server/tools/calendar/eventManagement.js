import { applyUserStyling } from '../common/sharedUtils.js';
import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse, safeStringify } from '../../utils/jsonUtils.js';

// Get detailed event information
export async function getEventTool(authManager, args) {
  const { eventId, calendarId } = args;

  if (!eventId) {
    return createValidationError('eventId', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    
    const endpoint = calendarId ? 
      `/me/calendars/${calendarId}/events/${eventId}` : 
      `/me/events/${eventId}`;
    
    const options = {
      select: 'id,subject,start,end,location,attendees,body,bodyPreview,organizer,isAllDay,showAs,sensitivity,importance,recurrence,reminderMinutesBeforeStart,responseRequested,allowNewTimeProposals,onlineMeeting,isOnlineMeeting,onlineMeetingProvider,categories,createdDateTime,lastModifiedDateTime'
    };

    const event = await graphApiClient.makeRequest(endpoint, options);

    const eventData = {
      id: event.id,
      subject: event.subject,
      start: event.start,
      end: event.end,
      location: event.location || {},
      attendees: event.attendees?.map(a => ({
        emailAddress: a.emailAddress,
        status: a.status,
        type: a.type
      })) || [],
      body: {
        contentType: event.body?.contentType || 'Text',
        content: event.body?.content || ''
      },
      bodyPreview: event.bodyPreview,
      organizer: event.organizer,
      isAllDay: event.isAllDay,
      showAs: event.showAs,
      sensitivity: event.sensitivity,
      importance: event.importance,
      recurrence: event.recurrence,
      reminderMinutesBeforeStart: event.reminderMinutesBeforeStart,
      responseRequested: event.responseRequested,
      allowNewTimeProposals: event.allowNewTimeProposals,
      onlineMeeting: event.onlineMeeting,
      isOnlineMeeting: event.isOnlineMeeting,
      onlineMeetingProvider: event.onlineMeetingProvider,
      categories: event.categories || [],
      createdDateTime: event.createdDateTime,
      lastModifiedDateTime: event.lastModifiedDateTime
    };

    return createSafeResponse(eventData);
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to get event');
  }
}

// Update event
export async function updateEventTool(authManager, args) {
  const { eventId, subject, start, end, body, bodyType = 'text', location, attendees, isOnlineMeeting, onlineMeetingProvider, recurrence, preserveUserStyling = true } = args;

  if (!eventId) {
    return createValidationError('eventId', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const updateData = {};

    if (subject !== undefined) {
      updateData.subject = subject;
    }

    if (start !== undefined) {
      updateData.start = start;
    }

    if (end !== undefined) {
      updateData.end = end;
    }

    if (body !== undefined) {
      let finalBody = body;
      let finalBodyType = bodyType;
      
      if (preserveUserStyling && finalBody) {
        const styledBody = await applyUserStyling(graphApiClient, finalBody, bodyType);
        finalBody = styledBody.content;
        finalBodyType = styledBody.type;
      }

      updateData.body = {
        contentType: finalBodyType === 'html' ? 'HTML' : 'Text',
        content: finalBody,
      };
    }

    if (location !== undefined) {
      updateData.location = {
        displayName: location,
      };
    }

    if (attendees !== undefined) {
      updateData.attendees = attendees.map(email => ({
        emailAddress: { address: email },
        type: 'required',
      }));
    }

    if (isOnlineMeeting !== undefined) {
      updateData.isOnlineMeeting = isOnlineMeeting;
      if (isOnlineMeeting && onlineMeetingProvider) {
        updateData.onlineMeetingProvider = onlineMeetingProvider;
      }
    }

    if (recurrence !== undefined) {
      updateData.recurrence = recurrence;
    }

    const result = await graphApiClient.makeRequest(`/me/events/${eventId}`, {
      body: updateData
    }, 'PATCH');

    return {
      content: [
        {
          type: 'text',
          text: `Event updated successfully. Event ID: ${result.id}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to update event');
  }
}

// Delete event
export async function deleteEventTool(authManager, args) {
  const { eventId, calendarId } = args;

  if (!eventId) {
    return createValidationError('eventId', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const endpoint = calendarId ? 
      `/me/calendars/${calendarId}/events/${eventId}` : 
      `/me/events/${eventId}`;

    await graphApiClient.makeRequest(endpoint, {}, 'DELETE');

    return {
      content: [
        {
          type: 'text',
          text: `Event deleted successfully. Event ID: ${eventId}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to delete event');
  }
}

// Respond to event invitation
export async function respondToInviteTool(authManager, args) {
  const { eventId, response = 'accept', comment = '' } = args;

  if (!eventId) {
    return createValidationError('eventId', 'Parameter is required');
  }

  if (!['accept', 'tentativelyAccept', 'decline'].includes(response)) {
    return createValidationError('response', 'Must be one of: accept, tentativelyAccept, decline');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const payload = {
      comment,
      sendResponse: true
    };

    const result = await graphApiClient.postWithRetry(`/me/events/${eventId}/${response}`, payload);

    return {
      content: [
        {
          type: 'text',
          text: `Response "${response}" sent successfully for event ${eventId}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to respond to invite');
  }
}

// Validate event date times
export async function validateEventDateTimesTool(authManager, args) {
  const { startDateTime, endDateTime, timeZone = 'UTC' } = args;

  if (!startDateTime) {
    return createValidationError('startDateTime', 'Parameter is required');
  }

  if (!endDateTime) {
    return createValidationError('endDateTime', 'Parameter is required');
  }

  try {
    const start = new Date(startDateTime);
    const end = new Date(endDateTime);

    if (isNaN(start.getTime())) {
      return createValidationError('startDateTime', `Invalid date format: ${startDateTime}`);
    }

    if (isNaN(end.getTime())) {
      return createValidationError('endDateTime', `Invalid date format: ${endDateTime}`);
    }

    if (start >= end) {
      return createValidationError('dateRange', 'Start date/time must be before end date/time');
    }

    const duration = end.getTime() - start.getTime();
    const durationMinutes = Math.floor(duration / (1000 * 60));
    const durationHours = Math.floor(durationMinutes / 60);
    const remainingMinutes = durationMinutes % 60;

    const validation = {
      valid: true,
      startDateTime: {
        original: startDateTime,
        parsed: start.toISOString(),
        formatted: start.toLocaleString(),
        timeZone: timeZone
      },
      endDateTime: {
        original: endDateTime,
        parsed: end.toISOString(),
        formatted: end.toLocaleString(),
        timeZone: timeZone
      },
      duration: {
        totalMinutes: durationMinutes,
        hours: durationHours,
        minutes: remainingMinutes,
        formatted: `${durationHours}h ${remainingMinutes}m`
      }
    };

    return createSafeResponse(validation);
  } catch (error) {
    return createSafeResponse({
      valid: false,
      error: error.message
    });
  }
}