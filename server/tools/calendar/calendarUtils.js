import { applyUserStyling } from '../common/sharedUtils.js';
import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse, safeStringify } from '../../utils/jsonUtils.js';

// Create recurring event
export async function createRecurringEventTool(authManager, args) {
  const { subject, start, end, recurrencePattern, body = '', bodyType = 'text', location = '', attendees = [], isOnlineMeeting = false, preserveUserStyling = true } = args;

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
      recurrence: recurrencePattern,
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

    if (isOnlineMeeting) {
      event.isOnlineMeeting = true;
      event.onlineMeetingProvider = 'teamsForBusiness';
    }

    const result = await graphApiClient.postWithRetry('/me/events', event);

    const meetingType = isOnlineMeeting ? 'Teams meeting' : 'Event';
    const successMessage = `Recurring ${meetingType} "${subject}" created successfully. Event ID: ${result.id}` +
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
    return convertErrorToToolError(error, 'Failed to create recurring event');
  }
}

// Find meeting times
export async function findMeetingTimesTool(authManager, args) {
  const { attendees = [], timeConstraint, maxCandidates = 20, meetingDuration = 60 } = args;

  if (!attendees || attendees.length === 0) {
    return createValidationError('attendees', 'At least one attendee is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const requestBody = {
      schedules: attendees,
      startTime: timeConstraint?.start,
      endTime: timeConstraint?.end,
      maxCandidates,
      meetingDuration,
    };

    const result = await graphApiClient.postWithRetry('/me/calendar/getSchedule', requestBody);

    return createSafeResponse(result);
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to find meeting times');
  }
}

// Check availability
export async function checkAvailabilityTool(authManager, args) {
  const { schedules = [], startTime, endTime, availabilityViewInterval = 60 } = args;

  if (!schedules || schedules.length === 0) {
    return createValidationError('schedules', 'At least one schedule is required');
  }

  if (!startTime) {
    return createValidationError('startTime', 'Parameter is required');
  }

  if (!endTime) {
    return createValidationError('endTime', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const requestBody = {
      schedules,
      startTime: {
        dateTime: startTime,
        timeZone: 'UTC'
      },
      endTime: {
        dateTime: endTime,
        timeZone: 'UTC'
      },
      availabilityViewInterval
    };

    const result = await graphApiClient.postWithRetry('/me/calendar/getSchedule', requestBody);

    return createSafeResponse(result);
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to check availability');
  }
}

// Schedule online meeting
export async function scheduleOnlineMeetingTool(authManager, args) {
  const { subject, startTime, endTime, attendees = [], body = '', bodyType = 'text', meetingProvider = 'teamsForBusiness', preserveUserStyling = true } = args;

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
      start: {
        dateTime: startTime,
        timeZone: 'UTC'
      },
      end: {
        dateTime: endTime,
        timeZone: 'UTC'
      },
      body: {
        contentType: finalBodyType === 'html' ? 'HTML' : 'Text',
        content: finalBody,
      },
      isOnlineMeeting: true,
      onlineMeetingProvider: meetingProvider,
    };

    if (attendees.length > 0) {
      event.attendees = attendees.map(email => ({
        emailAddress: { address: email },
        type: 'required',
      }));
    }

    const result = await graphApiClient.postWithRetry('/me/events', event);

    const successMessage = `Online meeting "${subject}" scheduled successfully. Event ID: ${result.id}` +
      (result.onlineMeeting?.joinUrl ? ` Join URL: ${result.onlineMeeting.joinUrl}` : '');

    return {
      content: [
        {
          type: 'text',
          text: successMessage,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to schedule online meeting');
  }
}

// List calendars
export async function listCalendarsTool(authManager, args) {
  const { includeSharedCalendars = false, top = 100 } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const options = {
      select: 'id,name,color,isDefaultCalendar,canShare,canViewPrivateItems,canEdit,owner',
      top: Math.min(top, 1000)
    };

    const result = await graphApiClient.makeRequest('/me/calendars', options);

    const calendars = result.value?.map(calendar => ({
      id: calendar.id,
      name: calendar.name,
      color: calendar.color,
      isDefaultCalendar: calendar.isDefaultCalendar,
      canShare: calendar.canShare,
      canViewPrivateItems: calendar.canViewPrivateItems,
      canEdit: calendar.canEdit,
      owner: calendar.owner
    })) || [];

    return createSafeResponse({ calendars, count: calendars.length });
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to list calendars');
  }
}

// Get calendar view
export async function getCalendarViewTool(authManager, args) {
  const { startDateTime, endDateTime, calendarId, top = 100 } = args;

  if (!startDateTime) {
    return createValidationError('startDateTime', 'Parameter is required');
  }

  if (!endDateTime) {
    return createValidationError('endDateTime', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const endpoint = calendarId ? `/me/calendars/${calendarId}/calendarView` : '/me/calendarView';
    const options = {
      startDateTime,
      endDateTime,
      select: 'id,subject,start,end,location,attendees,bodyPreview,organizer,isAllDay,showAs,importance,sensitivity,categories,webLink',
      top: Math.min(top, 1000),
      orderby: 'start/dateTime'
    };

    const result = await graphApiClient.makeRequest(endpoint, options);

    const events = result.value?.map(event => ({
      id: event.id,
      subject: event.subject,
      start: event.start,
      end: event.end,
      location: event.location?.displayName || 'No location',
      attendees: event.attendees?.map(a => a.emailAddress?.address) || [],
      preview: event.bodyPreview,
      organizer: event.organizer?.emailAddress?.address || 'Unknown',
      isAllDay: event.isAllDay,
      showAs: event.showAs,
      importance: event.importance,
      sensitivity: event.sensitivity,
      categories: event.categories || [],
      webLink: event.webLink
    })) || [];

    return createSafeResponse({ events, count: events.length });
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to get calendar view');
  }
}

// Get busy times
export async function getBusyTimesTool(authManager, args) {
  const { schedules = [], startTime, endTime, availabilityViewInterval = 60 } = args;

  if (!schedules || schedules.length === 0) {
    return createValidationError('schedules', 'At least one schedule is required');
  }

  if (!startTime) {
    return createValidationError('startTime', 'Parameter is required');
  }

  if (!endTime) {
    return createValidationError('endTime', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const requestBody = {
      schedules,
      startTime: {
        dateTime: startTime,
        timeZone: 'UTC'
      },
      endTime: {
        dateTime: endTime,
        timeZone: 'UTC'
      },
      availabilityViewInterval
    };

    const result = await graphApiClient.postWithRetry('/me/calendar/getSchedule', requestBody);

    // Process to extract busy times
    const busyTimes = [];
    
    result.value?.forEach((schedule, index) => {
      const userSchedule = schedules[index];
      const busySchedule = {
        user: userSchedule,
        busyTimes: []
      };

      schedule.busyViewData?.forEach((busyStatus, intervalIndex) => {
        if (busyStatus === '2') { // '2' indicates busy
          const intervalStart = new Date(startTime);
          intervalStart.setMinutes(intervalStart.getMinutes() + (intervalIndex * availabilityViewInterval));
          
          const intervalEnd = new Date(intervalStart);
          intervalEnd.setMinutes(intervalEnd.getMinutes() + availabilityViewInterval);
          
          busySchedule.busyTimes.push({
            start: intervalStart.toISOString(),
            end: intervalEnd.toISOString()
          });
        }
      });

      busyTimes.push(busySchedule);
    });

    return createSafeResponse({ busyTimes, rawData: result });
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to get busy times');
  }
}

// Build recurrence pattern
export async function buildRecurrencePatternTool(authManager, args) {
  const { 
    patternType = 'daily', 
    interval = 1, 
    daysOfWeek = [], 
    dayOfMonth = 1, 
    monthOfYear = 1, 
    index = 'first',
    rangeType = 'noEnd',
    numberOfOccurrences = 10,
    recurrenceTimeZone = 'UTC',
    rangeStartDate,
    rangeEndDate
  } = args;

  try {
    const validPatternTypes = ['daily', 'weekly', 'absoluteMonthly', 'relativeMonthly', 'absoluteYearly', 'relativeYearly'];
    const validRangeTypes = ['endDate', 'noEnd', 'numbered'];
    const validIndexes = ['first', 'second', 'third', 'fourth', 'last'];
    const validDaysOfWeek = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];

    if (!validPatternTypes.includes(patternType)) {
      return createValidationError('patternType', `Must be one of: ${validPatternTypes.join(', ')}`);
    }

    if (!validRangeTypes.includes(rangeType)) {
      return createValidationError('rangeType', `Must be one of: ${validRangeTypes.join(', ')}`);
    }

    const recurrencePattern = {
      pattern: {
        type: patternType,
        interval: interval
      },
      range: {
        type: rangeType,
        recurrenceTimeZone: recurrenceTimeZone
      }
    };

    // Add pattern-specific fields
    if (patternType === 'weekly' || patternType === 'relativeMonthly' || patternType === 'relativeYearly') {
      if (daysOfWeek.length > 0) {
        const invalidDays = daysOfWeek.filter(day => !validDaysOfWeek.includes(day));
        if (invalidDays.length > 0) {
          return createValidationError('daysOfWeek', `Invalid days: ${invalidDays.join(', ')}. Must be: ${validDaysOfWeek.join(', ')}`);
        }
        recurrencePattern.pattern.daysOfWeek = daysOfWeek;
      }
    }

    if (patternType === 'absoluteMonthly' || patternType === 'absoluteYearly') {
      recurrencePattern.pattern.dayOfMonth = dayOfMonth;
    }

    if (patternType === 'relativeMonthly' || patternType === 'relativeYearly') {
      if (!validIndexes.includes(index)) {
        return createValidationError('index', `Must be one of: ${validIndexes.join(', ')}`);
      }
      recurrencePattern.pattern.index = index;
    }

    if (patternType === 'absoluteYearly' || patternType === 'relativeYearly') {
      recurrencePattern.pattern.month = monthOfYear;
    }

    // Add range-specific fields
    if (rangeType === 'endDate') {
      if (!rangeEndDate) {
        return createValidationError('rangeEndDate', 'Required when rangeType is endDate');
      }
      recurrencePattern.range.endDate = rangeEndDate;
    }

    if (rangeType === 'numbered') {
      recurrencePattern.range.numberOfOccurrences = numberOfOccurrences;
    }

    if (rangeStartDate) {
      recurrencePattern.range.startDate = rangeStartDate;
    }

    const patternInfo = {
      recurrencePattern,
      description: generateRecurrenceDescription(recurrencePattern),
      validation: {
        valid: true,
        patternType,
        rangeType,
        interval
      }
    };

    return createSafeResponse(patternInfo);
  } catch (error) {
    return createSafeResponse({
      valid: false,
      error: error.message
    });
  }
}

// Helper function to generate recurrence description
function generateRecurrenceDescription(recurrencePattern) {
  const { pattern, range } = recurrencePattern;
  let description = '';

  // Pattern description
  switch (pattern.type) {
    case 'daily':
      description = pattern.interval === 1 ? 'Every day' : `Every ${pattern.interval} days`;
      break;
    case 'weekly':
      const days = pattern.daysOfWeek ? pattern.daysOfWeek.join(', ') : 'every week';
      description = pattern.interval === 1 ? `Weekly on ${days}` : `Every ${pattern.interval} weeks on ${days}`;
      break;
    case 'absoluteMonthly':
      description = pattern.interval === 1 ? 
        `Monthly on day ${pattern.dayOfMonth}` : 
        `Every ${pattern.interval} months on day ${pattern.dayOfMonth}`;
      break;
    case 'relativeMonthly':
      const monthlyDays = pattern.daysOfWeek ? pattern.daysOfWeek.join(', ') : 'weekday';
      description = pattern.interval === 1 ? 
        `Monthly on the ${pattern.index} ${monthlyDays}` : 
        `Every ${pattern.interval} months on the ${pattern.index} ${monthlyDays}`;
      break;
    case 'absoluteYearly':
      description = `Yearly on day ${pattern.dayOfMonth} of month ${pattern.month}`;
      break;
    case 'relativeYearly':
      const yearlyDays = pattern.daysOfWeek ? pattern.daysOfWeek.join(', ') : 'weekday';
      description = `Yearly on the ${pattern.index} ${yearlyDays} of month ${pattern.month}`;
      break;
  }

  // Range description
  switch (range.type) {
    case 'endDate':
      description += ` until ${range.endDate}`;
      break;
    case 'numbered':
      description += ` for ${range.numberOfOccurrences} occurrences`;
      break;
    case 'noEnd':
      description += ' (no end date)';
      break;
  }

  return description;
}

// Create recurrence helper
export async function createRecurrenceHelperTool(authManager, args) {
  const { 
    eventTitle = 'Recurring Event',
    startDateTime,
    endDateTime,
    recurrenceType = 'daily',
    endAfter = 'occurrences',
    occurrences = 10,
    endDate,
    weekDays = [],
    monthDay = 1,
    monthWeek = 'first',
    yearMonth = 1,
    timeZone = 'UTC',
    location = '',
    attendees = [],
    body = '',
    bodyType = 'text',
    preserveUserStyling = true
  } = args;

  if (!startDateTime) {
    return createValidationError('startDateTime', 'Parameter is required');
  }

  if (!endDateTime) {
    return createValidationError('endDateTime', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    // Build recurrence pattern based on type
    let recurrencePattern;
    
    switch (recurrenceType) {
      case 'daily':
        recurrencePattern = {
          pattern: {
            type: 'daily',
            interval: 1
          },
          range: {
            type: endAfter === 'date' ? 'endDate' : 'numbered',
            recurrenceTimeZone: timeZone
          }
        };
        break;

      case 'weekly':
        recurrencePattern = {
          pattern: {
            type: 'weekly',
            interval: 1,
            daysOfWeek: weekDays.length > 0 ? weekDays : ['monday']
          },
          range: {
            type: endAfter === 'date' ? 'endDate' : 'numbered',
            recurrenceTimeZone: timeZone
          }
        };
        break;

      case 'monthly':
        recurrencePattern = {
          pattern: {
            type: 'absoluteMonthly',
            interval: 1,
            dayOfMonth: monthDay
          },
          range: {
            type: endAfter === 'date' ? 'endDate' : 'numbered',
            recurrenceTimeZone: timeZone
          }
        };
        break;

      case 'yearly':
        recurrencePattern = {
          pattern: {
            type: 'absoluteYearly',
            interval: 1,
            dayOfMonth: monthDay,
            month: yearMonth
          },
          range: {
            type: endAfter === 'date' ? 'endDate' : 'numbered',
            recurrenceTimeZone: timeZone
          }
        };
        break;

      default:
        return createValidationError('recurrenceType', 'Must be one of: daily, weekly, monthly, yearly');
    }

    // Set range details
    if (endAfter === 'date') {
      if (!endDate) {
        return createValidationError('endDate', 'Required when endAfter is "date"');
      }
      recurrencePattern.range.endDate = endDate;
    } else {
      recurrencePattern.range.numberOfOccurrences = occurrences;
    }

    // Apply user styling if enabled and body is provided
    let finalBody = body;
    let finalBodyType = bodyType;
    
    if (preserveUserStyling && finalBody) {
      const styledBody = await applyUserStyling(graphApiClient, finalBody, bodyType);
      finalBody = styledBody.content;
      finalBodyType = styledBody.type;
    }

    // Create the event
    const event = {
      subject: eventTitle,
      start: {
        dateTime: startDateTime,
        timeZone: timeZone
      },
      end: {
        dateTime: endDateTime,
        timeZone: timeZone
      },
      body: {
        contentType: finalBodyType === 'html' ? 'HTML' : 'Text',
        content: finalBody,
      },
      recurrence: recurrencePattern
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

    const result = await graphApiClient.postWithRetry('/me/events', event);

    const description = generateRecurrenceDescription(recurrencePattern);
    
    return {
      content: [
        {
          type: 'text',
          text: `Recurring event "${eventTitle}" created successfully.\nEvent ID: ${result.id}\nRecurrence: ${description}`,
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to create recurring event');
  }
}

// Check calendar permissions
export async function checkCalendarPermissionsTool(authManager, args) {
  const { calendarId } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const endpoint = calendarId ? `/me/calendars/${calendarId}` : '/me/calendar';
    const options = {
      select: 'id,name,canEdit,canShare,canViewPrivateItems,owner,permissions'
    };

    const calendar = await graphApiClient.makeRequest(endpoint, options);

    const permissions = {
      calendarId: calendar.id,
      calendarName: calendar.name,
      canEdit: calendar.canEdit,
      canShare: calendar.canShare,
      canViewPrivateItems: calendar.canViewPrivateItems,
      owner: calendar.owner,
      permissions: calendar.permissions || []
    };

    return createSafeResponse(permissions);
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to check calendar permissions');
  }
}