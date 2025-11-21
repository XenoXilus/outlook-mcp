/**
 * Calendar-related MCP tool schemas
 * 
 * This module contains all JSON schemas for calendar operations in the Outlook MCP server.
 * Includes support for event creation, recurring meetings, and Teams integration.
 */

export const listEventsSchema = {
  name: 'outlook_list_events',
  description: 'List calendar events from Outlook',
  inputSchema: {
    type: 'object',
    properties: {
      startDateTime: {
        type: 'string',
        description: 'Start date/time in ISO 8601 format',
      },
      endDateTime: {
        type: 'string',
        description: 'End date/time in ISO 8601 format',
      },
      limit: {
        type: 'number',
        description: 'Maximum number of events to return',
        default: 10,
      },
      calendar: {
        type: 'string',
        description: 'Calendar ID (default: primary calendar)',
      },
    },
  },
};

export const createEventSchema = {
  name: 'outlook_create_event',
  description: 'Create a new calendar event in Outlook with optional Teams meeting integration',
  inputSchema: {
    type: 'object',
    properties: {
      subject: {
        type: 'string',
        description: 'Event subject/title',
      },
      start: {
        type: 'object',
        description: 'Event start date and time configuration',
        properties: {
          dateTime: {
            type: 'string',
            description: 'Start date/time in ISO 8601 format',
          },
          timeZone: {
            type: 'string',
            description: 'Time zone (e.g., "Pacific Standard Time")',
          },
        },
        required: ['dateTime', 'timeZone'],
      },
      end: {
        type: 'object',
        description: 'Event end date and time configuration',
        properties: {
          dateTime: {
            type: 'string',
            description: 'End date/time in ISO 8601 format',
          },
          timeZone: {
            type: 'string',
            description: 'Time zone (e.g., "Pacific Standard Time")',
          },
        },
        required: ['dateTime', 'timeZone'],
      },
      body: {
        type: 'string',
        description: 'Event description',
      },
      location: {
        type: 'string',
        description: 'Event location',
      },
      attendees: {
        type: 'array',
        items: { type: 'string' },
        description: 'Attendee email addresses',
      },
      isOnlineMeeting: {
        type: 'boolean',
        description: 'Whether to create this as a Teams meeting (default: false)',
      },
      onlineMeetingProvider: {
        type: 'string',
        enum: ['teamsForBusiness', 'skypeForBusiness'],
        description: 'Online meeting provider (default: "teamsForBusiness")',
      },
      recurrence: {
        type: 'object',
        description: 'Recurrence pattern for recurring meetings',
        properties: {
          pattern: {
            type: 'object',
            description: 'The recurrence pattern',
            properties: {
              type: {
                type: 'string',
                enum: ['daily', 'weekly', 'absoluteMonthly', 'relativeMonthly', 'absoluteYearly', 'relativeYearly'],
                description: 'The recurrence pattern type',
              },
              interval: {
                type: 'integer',
                minimum: 1,
                description: 'Number of units between occurrences',
              },
              daysOfWeek: {
                type: 'array',
                items: {
                  type: 'string',
                  enum: ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'],
                },
                description: 'Days of the week (required for weekly, relativeMonthly, relativeYearly)',
              },
              dayOfMonth: {
                type: 'integer',
                minimum: 1,
                maximum: 31,
                description: 'Day of the month (required for absoluteMonthly, absoluteYearly)',
              },
              month: {
                type: 'integer',
                minimum: 1,
                maximum: 12,
                description: 'Month of the year (required for absoluteYearly, relativeYearly)',
              },
              index: {
                type: 'string',
                enum: ['first', 'second', 'third', 'fourth', 'last'],
                description: 'Instance of the allowed days (for relativeMonthly, relativeYearly)',
              },
              firstDayOfWeek: {
                type: 'string',
                enum: ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'],
                description: 'First day of the week (for weekly patterns, default: sunday)',
              },
            },
            required: ['type', 'interval'],
          },
          range: {
            type: 'object',
            description: 'The recurrence range',
            properties: {
              type: {
                type: 'string',
                enum: ['numbered', 'endDate', 'noEnd'],
                description: 'The recurrence range type',
              },
              startDate: {
                type: 'string',
                format: 'date',
                description: 'Start date of the recurrence (YYYY-MM-DD)',
              },
              endDate: {
                type: 'string',
                format: 'date',
                description: 'End date of the recurrence (YYYY-MM-DD, required for endDate type)',
              },
              numberOfOccurrences: {
                type: 'integer',
                minimum: 1,
                description: 'Number of occurrences (required for numbered type)',
              },
            },
            required: ['type', 'startDate'],
          },
        },
        required: ['pattern', 'range'],
      },
    },
    required: ['subject', 'start', 'end'],
  },
};

export const getEventSchema = {
  name: 'outlook_get_event',
  description: 'Get a specific calendar event',
  inputSchema: {
    type: 'object',
    properties: {
      eventId: {
        type: 'string',
        description: 'The ID of the event to retrieve',
      },
    },
    required: ['eventId'],
  },
};

export const updateEventSchema = {
  name: 'outlook_update_event',
  description: 'Update an existing calendar event',
  inputSchema: {
    type: 'object',
    properties: {
      eventId: {
        type: 'string',
        description: 'The ID of the event to update',
      },
      subject: {
        type: 'string',
        description: 'New subject',
      },
      body: {
        type: 'string',
        description: 'New body content',
      },
      location: {
        type: 'string',
        description: 'New location',
      },
      start: {
        type: 'object',
        description: 'New start time',
        properties: {
          dateTime: { type: 'string' },
          timeZone: { type: 'string' },
        },
      },
      end: {
        type: 'object',
        description: 'New end time',
        properties: {
          dateTime: { type: 'string' },
          timeZone: { type: 'string' },
        },
      },
      attendees: {
        type: 'array',
        items: { type: 'string' },
        description: 'New list of attendees',
      },
    },
    required: ['eventId'],
  },
};

export const deleteEventSchema = {
  name: 'outlook_delete_event',
  description: 'Delete a calendar event',
  inputSchema: {
    type: 'object',
    properties: {
      eventId: {
        type: 'string',
        description: 'The ID of the event to delete',
      },
    },
    required: ['eventId'],
  },
};

export const respondToInviteSchema = {
  name: 'outlook_respond_to_invite',
  description: 'Respond to a meeting invitation',
  inputSchema: {
    type: 'object',
    properties: {
      eventId: {
        type: 'string',
        description: 'The ID of the event to respond to',
      },
      response: {
        type: 'string',
        enum: ['accept', 'decline', 'tentativelyAccept'],
        description: 'Response type',
      },
      comment: {
        type: 'string',
        description: 'Optional comment',
      },
      sendResponse: {
        type: 'boolean',
        description: 'Whether to send a response email (default: true)',
      },
    },
    required: ['eventId', 'response'],
  },
};

export const validateEventDateTimesSchema = {
  name: 'outlook_validate_event_datetimes',
  description: 'Validate event start and end times',
  inputSchema: {
    type: 'object',
    properties: {
      start: {
        type: 'object',
        description: 'Start time to validate',
        properties: {
          dateTime: { type: 'string' },
          timeZone: { type: 'string' },
        },
        required: ['dateTime', 'timeZone'],
      },
      end: {
        type: 'object',
        description: 'End time to validate',
        properties: {
          dateTime: { type: 'string' },
          timeZone: { type: 'string' },
        },
        required: ['dateTime', 'timeZone'],
      },
    },
    required: ['start', 'end'],
  },
};

export const createRecurringEventSchema = {
  name: 'outlook_create_recurring_event',
  description: 'Create a recurring calendar event',
  inputSchema: {
    type: 'object',
    properties: {
      subject: {
        type: 'string',
        description: 'Event subject',
      },
      start: {
        type: 'object',
        description: 'Start time',
        properties: {
          dateTime: { type: 'string' },
          timeZone: { type: 'string' },
        },
        required: ['dateTime', 'timeZone'],
      },
      end: {
        type: 'object',
        description: 'End time',
        properties: {
          dateTime: { type: 'string' },
          timeZone: { type: 'string' },
        },
        required: ['dateTime', 'timeZone'],
      },
      recurrencePattern: {
        type: 'object',
        description: 'Recurrence pattern object',
      },
      body: {
        type: 'string',
        description: 'Event body content',
      },
      location: {
        type: 'string',
        description: 'Event location',
      },
      attendees: {
        type: 'array',
        items: { type: 'string' },
        description: 'List of attendees',
      },
      isOnlineMeeting: {
        type: 'boolean',
        description: 'Whether to make this an online meeting',
      },
    },
    required: ['subject', 'start', 'end', 'recurrencePattern'],
  },
};

export const findMeetingTimesSchema = {
  name: 'outlook_find_meeting_times',
  description: 'Find optimal meeting times for attendees',
  inputSchema: {
    type: 'object',
    properties: {
      attendees: {
        type: 'array',
        items: { type: 'string' },
        description: 'List of attendees to check availability for',
      },
      timeConstraint: {
        type: 'object',
        description: 'Time range to search within',
        properties: {
          start: { type: 'object' },
          end: { type: 'object' },
        },
      },
      meetingDuration: {
        type: 'string',
        description: 'Duration of the meeting (ISO 8601 duration)',
      },
      maxCandidates: {
        type: 'integer',
        description: 'Maximum number of time slots to return',
      },
    },
    required: ['attendees'],
  },
};

export const checkAvailabilitySchema = {
  name: 'outlook_check_availability',
  description: 'Check availability for users',
  inputSchema: {
    type: 'object',
    properties: {
      schedules: {
        type: 'array',
        items: { type: 'string' },
        description: 'List of email addresses to check',
      },
      startTime: {
        type: 'string',
        description: 'Start of the time range',
      },
      endTime: {
        type: 'string',
        description: 'End of the time range',
      },
      availabilityViewInterval: {
        type: 'integer',
        description: 'Interval in minutes for availability view',
      },
    },
    required: ['schedules', 'startTime', 'endTime'],
  },
};

export const scheduleOnlineMeetingSchema = {
  name: 'outlook_schedule_online_meeting',
  description: 'Schedule an online meeting (Teams/Skype)',
  inputSchema: {
    type: 'object',
    properties: {
      subject: {
        type: 'string',
        description: 'Meeting subject',
      },
      startTime: {
        type: 'string',
        description: 'Start time',
      },
      endTime: {
        type: 'string',
        description: 'End time',
      },
      attendees: {
        type: 'array',
        items: { type: 'string' },
        description: 'List of attendees',
      },
      meetingProvider: {
        type: 'string',
        enum: ['teamsForBusiness', 'skypeForBusiness'],
        description: 'Online meeting provider',
      },
    },
    required: ['subject', 'startTime', 'endTime'],
  },
};

export const listCalendarsSchema = {
  name: 'outlook_list_calendars',
  description: 'List available calendars',
  inputSchema: {
    type: 'object',
    properties: {
      includeSharedCalendars: {
        type: 'boolean',
        description: 'Whether to include shared calendars',
      },
      top: {
        type: 'integer',
        description: 'Number of calendars to return',
      },
    },
  },
};

export const getCalendarViewSchema = {
  name: 'outlook_get_calendar_view',
  description: 'Get a view of a calendar for a specific time range',
  inputSchema: {
    type: 'object',
    properties: {
      startDateTime: {
        type: 'string',
        description: 'Start of the time range',
      },
      endDateTime: {
        type: 'string',
        description: 'End of the time range',
      },
      calendarId: {
        type: 'string',
        description: 'ID of the calendar to view',
      },
      top: {
        type: 'integer',
        description: 'Number of events to return',
      },
    },
    required: ['startDateTime', 'endDateTime'],
  },
};

export const getBusyTimesSchema = {
  name: 'outlook_get_busy_times',
  description: 'Get busy times for users',
  inputSchema: {
    type: 'object',
    properties: {
      schedules: {
        type: 'array',
        items: { type: 'string' },
        description: 'List of email addresses to check',
      },
      startTime: {
        type: 'string',
        description: 'Start of the time range',
      },
      endTime: {
        type: 'string',
        description: 'End of the time range',
      },
      availabilityViewInterval: {
        type: 'integer',
        description: 'Interval in minutes for availability view',
      },
    },
    required: ['schedules', 'startTime', 'endTime'],
  },
};

export const buildRecurrencePatternSchema = {
  name: 'outlook_build_recurrence_pattern',
  description: 'Build a recurrence pattern object',
  inputSchema: {
    type: 'object',
    properties: {
      patternType: {
        type: 'string',
        description: 'Type of recurrence pattern',
      },
      interval: {
        type: 'integer',
        description: 'Interval between occurrences',
      },
      daysOfWeek: {
        type: 'array',
        items: { type: 'string' },
        description: 'Days of the week for the pattern',
      },
      dayOfMonth: {
        type: 'integer',
        description: 'Day of the month',
      },
      monthOfYear: {
        type: 'integer',
        description: 'Month of the year',
      },
      index: {
        type: 'string',
        description: 'Index for relative patterns (e.g., "first")',
      },
      rangeType: {
        type: 'string',
        description: 'Type of recurrence range',
      },
      numberOfOccurrences: {
        type: 'integer',
        description: 'Number of occurrences',
      },
      rangeStartDate: {
        type: 'string',
        description: 'Start date of the range',
      },
      rangeEndDate: {
        type: 'string',
        description: 'End date of the range',
      },
    },
    required: ['patternType', 'rangeType'],
  },
};

export const createRecurrenceHelperSchema = {
  name: 'outlook_create_recurrence_helper',
  description: 'Helper to create a recurring event with simplified inputs',
  inputSchema: {
    type: 'object',
    properties: {
      eventTitle: {
        type: 'string',
        description: 'Title of the event',
      },
      startDateTime: {
        type: 'string',
        description: 'Start date and time',
      },
      endDateTime: {
        type: 'string',
        description: 'End date and time',
      },
      recurrenceType: {
        type: 'string',
        description: 'Type of recurrence (daily, weekly, etc.)',
      },
      endAfter: {
        type: 'string',
        description: 'When to end the recurrence (date or occurrences)',
      },
      occurrences: {
        type: 'integer',
        description: 'Number of occurrences',
      },
      endDate: {
        type: 'string',
        description: 'End date',
      },
    },
    required: ['eventTitle', 'startDateTime', 'endDateTime', 'recurrenceType'],
  },
};

export const checkCalendarPermissionsSchema = {
  name: 'outlook_check_calendar_permissions',
  description: 'Check permissions for a calendar',
  inputSchema: {
    type: 'object',
    properties: {
      calendarId: {
        type: 'string',
        description: 'ID of the calendar to check',
      },
    },
  },
};

// Export all calendar schemas as an array for easy iteration
export const calendarSchemas = [
  listEventsSchema,
  createEventSchema,
  getEventSchema,
  updateEventSchema,
  deleteEventSchema,
  respondToInviteSchema,
  validateEventDateTimesSchema,
  createRecurringEventSchema,
  findMeetingTimesSchema,
  checkAvailabilitySchema,
  scheduleOnlineMeetingSchema,
  listCalendarsSchema,
  getCalendarViewSchema,
  getBusyTimesSchema,
  buildRecurrencePatternSchema,
  createRecurrenceHelperSchema,
  checkCalendarPermissionsSchema,
];

// Export mapping for quick lookup
export const calendarSchemaMap = {
  'outlook_list_events': listEventsSchema,
  'outlook_create_event': createEventSchema,
  'outlook_get_event': getEventSchema,
  'outlook_update_event': updateEventSchema,
  'outlook_delete_event': deleteEventSchema,
  'outlook_respond_to_invite': respondToInviteSchema,
  'outlook_validate_event_datetimes': validateEventDateTimesSchema,
  'outlook_create_recurring_event': createRecurringEventSchema,
  'outlook_find_meeting_times': findMeetingTimesSchema,
  'outlook_check_availability': checkAvailabilitySchema,
  'outlook_schedule_online_meeting': scheduleOnlineMeetingSchema,
  'outlook_list_calendars': listCalendarsSchema,
  'outlook_get_calendar_view': getCalendarViewSchema,
  'outlook_get_busy_times': getBusyTimesSchema,
  'outlook_build_recurrence_pattern': buildRecurrencePatternSchema,
  'outlook_create_recurrence_helper': createRecurrenceHelperSchema,
  'outlook_check_calendar_permissions': checkCalendarPermissionsSchema,
};