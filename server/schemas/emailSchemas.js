/**
 * Email-related MCP tool schemas
 * 
 * This module contains all JSON schemas for email operations in the Outlook MCP server.
 * Schemas are organized by functionality and include comprehensive validation rules.
 */

export const listEmailsSchema = {
  name: 'outlook_list_emails',
  description: 'List emails from Outlook inbox or specified folder',
  inputSchema: {
    type: 'object',
    properties: {
      folder: {
        type: 'string',
        description: 'Folder to list emails from (default: inbox)',
        default: 'inbox',
      },
      limit: {
        type: 'number',
        description: 'Maximum number of emails to return',
        default: 10,
      },
      filter: {
        type: 'string',
        description: 'OData filter query for emails',
      },
    },
  },
};

export const sendEmailSchema = {
  name: 'outlook_send_email',
  description: 'Send an email through Outlook',
  inputSchema: {
    type: 'object',
    properties: {
      to: {
        type: 'array',
        items: { type: 'string' },
        description: 'Recipient email addresses',
      },
      subject: {
        type: 'string',
        description: 'Email subject',
      },
      body: {
        type: 'string',
        description: 'Email body content',
      },
      bodyType: {
        type: 'string',
        enum: ['text', 'html'],
        default: 'text',
        description: 'Body content type',
      },
      cc: {
        type: 'array',
        items: { type: 'string' },
        description: 'CC recipients',
      },
      bcc: {
        type: 'array',
        items: { type: 'string' },
        description: 'BCC recipients',
      },
      preserveUserStyling: {
        type: 'boolean',
        description: 'Apply user\'s default Outlook styling, font preferences, and signature',
        default: true,
      },
    },
    required: ['to', 'subject', 'body'],
  },
};

export const getEmailSchema = {
  name: 'outlook_get_email',
  description: 'Get detailed information about a specific email',
  inputSchema: {
    type: 'object',
    properties: {
      messageId: {
        type: 'string',
        description: 'The ID of the email message to retrieve',
      },
      truncate: {
        type: 'boolean',
        description: 'Truncate long email bodies (default: true)',
        default: true,
      },
      maxLength: {
        type: 'number',
        description: 'Maximum length for truncated body (default: 1000)',
        default: 1000,
      },
      format: {
        type: 'string',
        enum: ['text', 'html'],
        description: 'Format of the body content (default: text)',
        default: 'text',
      },
    },
    required: ['messageId'],
  },
};

export const searchEmailsSchema = {
  name: 'outlook_search_emails',
  description: 'Search emails across all folders with advanced filters for analysis',
  inputSchema: {
    type: 'object',
    properties: {
      query: {
        type: 'string',
        description: 'Free-text search query across email content',
      },
      subject: {
        type: 'string',
        description: 'Search emails with specific subject text',
      },
      from: {
        type: 'string',
        description: 'Filter emails from specific sender',
      },
      startDate: {
        type: 'string',
        description: 'Start date for email search (ISO 8601 format)',
      },
      endDate: {
        type: 'string',
        description: 'End date for email search (ISO 8601 format)',
      },
      folders: {
        type: 'array',
        items: { type: 'string' },
        description: 'Specific folders to search in',
      },
      limit: {
        type: 'number',
        description: 'Maximum number of emails to return (max 1000)',
        default: 100,
      },
      includeBody: {
        type: 'boolean',
        description: 'Include full email body content for analysis',
        default: false,
      },
      truncate: {
        type: 'boolean',
        description: 'Truncate long email bodies (default: true)',
        default: true,
      },
      maxLength: {
        type: 'number',
        description: 'Maximum length for truncated body (default: 1000)',
        default: 1000,
      },
      format: {
        type: 'string',
        enum: ['text', 'html'],
        description: 'Format of the body content (default: text)',
        default: 'text',
      },
      orderBy: {
        type: 'string',
        description: 'Sort order (e.g., "receivedDateTime desc")',
        default: 'receivedDateTime desc',
      },
    },
  },
};

export const createDraftSchema = {
  name: 'outlook_create_draft',
  description: 'Create an email draft without sending',
  inputSchema: {
    type: 'object',
    properties: {
      to: {
        type: 'array',
        items: { type: 'string' },
        description: 'Recipient email addresses',
      },
      subject: {
        type: 'string',
        description: 'Email subject',
      },
      body: {
        type: 'string',
        description: 'Email body content',
      },
      bodyType: {
        type: 'string',
        enum: ['text', 'html'],
        default: 'text',
        description: 'Body content type',
      },
      cc: {
        type: 'array',
        items: { type: 'string' },
        description: 'CC recipients',
      },
      bcc: {
        type: 'array',
        items: { type: 'string' },
        description: 'BCC recipients',
      },
      importance: {
        type: 'string',
        enum: ['low', 'normal', 'high'],
        default: 'normal',
        description: 'Email importance level',
      },
    },
    required: ['to', 'subject'],
  },
};

export const replyToEmailSchema = {
  name: 'outlook_reply_to_email',
  description: 'Reply to an existing email',
  inputSchema: {
    type: 'object',
    properties: {
      messageId: {
        type: 'string',
        description: 'The ID of the email to reply to',
      },
      body: {
        type: 'string',
        description: 'Reply message body',
      },
      bodyType: {
        type: 'string',
        enum: ['text', 'html'],
        default: 'text',
        description: 'Body content type',
      },
      comment: {
        type: 'string',
        description: 'Optional comment for the reply',
      },
    },
    required: ['messageId'],
  },
};

export const replyAllSchema = {
  name: 'outlook_reply_all',
  description: 'Reply to all recipients of an existing email',
  inputSchema: {
    type: 'object',
    properties: {
      messageId: {
        type: 'string',
        description: 'The ID of the email to reply all to',
      },
      body: {
        type: 'string',
        description: 'Reply message body',
      },
      bodyType: {
        type: 'string',
        enum: ['text', 'html'],
        default: 'text',
        description: 'Body content type',
      },
      comment: {
        type: 'string',
        description: 'Optional comment for the reply',
      },
    },
    required: ['messageId'],
  },
};

export const forwardEmailSchema = {
  name: 'outlook_forward_email',
  description: 'Forward an existing email to new recipients',
  inputSchema: {
    type: 'object',
    properties: {
      messageId: {
        type: 'string',
        description: 'The ID of the email to forward',
      },
      to: {
        type: 'array',
        items: { type: 'string' },
        description: 'Recipient email addresses for forwarding',
      },
      body: {
        type: 'string',
        description: 'Additional message body for the forward',
      },
      bodyType: {
        type: 'string',
        enum: ['text', 'html'],
        default: 'text',
        description: 'Body content type',
      },
      comment: {
        type: 'string',
        description: 'Optional comment for the forward',
      },
    },
    required: ['messageId', 'to'],
  },
};

export const deleteEmailSchema = {
  name: 'outlook_delete_email',
  description: 'Delete an email (move to Deleted Items or permanently delete)',
  inputSchema: {
    type: 'object',
    properties: {
      messageId: {
        type: 'string',
        description: 'The ID of the email to delete',
      },
      permanentDelete: {
        type: 'boolean',
        description: 'Whether to permanently delete (true) or move to Deleted Items (false)',
        default: false,
      },
    },
    required: ['messageId'],
  },
};

export const moveEmailSchema = {
  name: 'outlook_move_email',
  description: 'Move an email to a different folder',
  inputSchema: {
    type: 'object',
    properties: {
      messageId: {
        type: 'string',
        description: 'The ID of the email to move',
      },
      destinationFolderId: {
        type: 'string',
        description: 'The ID of the destination folder',
      },
    },
    required: ['messageId', 'destinationFolderId'],
  },
};

export const markAsReadSchema = {
  name: 'outlook_mark_as_read',
  description: 'Mark an email as read or unread',
  inputSchema: {
    type: 'object',
    properties: {
      messageId: {
        type: 'string',
        description: 'The ID of the email to mark',
      },
      isRead: {
        type: 'boolean',
        description: 'Whether to mark as read (true) or unread (false)',
        default: true,
      },
    },
    required: ['messageId'],
  },
};

export const flagEmailSchema = {
  name: 'outlook_flag_email',
  description: 'Flag or unflag an email for follow-up',
  inputSchema: {
    type: 'object',
    properties: {
      messageId: {
        type: 'string',
        description: 'The ID of the email to flag',
      },
      flagStatus: {
        type: 'string',
        enum: ['notFlagged', 'complete', 'flagged'],
        description: 'The flag status to set',
        default: 'flagged',
      },
    },
    required: ['messageId'],
  },
};

export const categorizeEmailSchema = {
  name: 'outlook_categorize_email',
  description: 'Apply categories to an email',
  inputSchema: {
    type: 'object',
    properties: {
      messageId: {
        type: 'string',
        description: 'The ID of the email to categorize',
      },
      categories: {
        type: 'array',
        items: { type: 'string' },
        description: 'List of category names to apply',
        default: [],
      },
    },
    required: ['messageId'],
  },
};

export const archiveEmailSchema = {
  name: 'outlook_archive_email',
  description: 'Archive an email (move to Archive folder)',
  inputSchema: {
    type: 'object',
    properties: {
      messageId: {
        type: 'string',
        description: 'The ID of the email to archive',
      },
    },
    required: ['messageId'],
  },
};

export const batchProcessEmailsSchema = {
  name: 'outlook_batch_process_emails',
  description: 'Perform bulk operations on multiple emails',
  inputSchema: {
    type: 'object',
    properties: {
      messageIds: {
        type: 'array',
        items: { type: 'string' },
        description: 'Array of email IDs to process',
      },
      operation: {
        type: 'string',
        enum: ['markAsRead', 'markAsUnread', 'delete', 'move', 'flag', 'categorize'],
        description: 'The operation to perform on all emails',
      },
      operationData: {
        type: 'object',
        description: 'Additional data for the operation (e.g., destinationFolderId for move)',
        properties: {
          destinationFolderId: {
            type: 'string',
            description: 'Destination folder ID for move operations'
          },
          flagStatus: {
            type: 'string',
            description: 'Flag status for flag operations'
          },
          categories: {
            type: 'array',
            items: { type: 'string' },
            description: 'Categories for categorize operations'
          },
          permanentDelete: {
            type: 'boolean',
            description: 'Whether to permanently delete for delete operations'
          },
        },
      },
    },
    required: ['messageIds', 'operation'],
  },
};

// Export all email schemas as an array for easy iteration
export const emailSchemas = [
  listEmailsSchema,
  sendEmailSchema,
  getEmailSchema,
  searchEmailsSchema,
  createDraftSchema,
  replyToEmailSchema,
  replyAllSchema,
  forwardEmailSchema,
  deleteEmailSchema,
  moveEmailSchema,
  markAsReadSchema,
  flagEmailSchema,
  categorizeEmailSchema,
  archiveEmailSchema,
  batchProcessEmailsSchema,
];

// Export mapping for quick lookup
export const emailSchemaMap = {
  'outlook_list_emails': listEmailsSchema,
  'outlook_send_email': sendEmailSchema,
  'outlook_get_email': getEmailSchema,
  'outlook_search_emails': searchEmailsSchema,
  'outlook_create_draft': createDraftSchema,
  'outlook_reply_to_email': replyToEmailSchema,
  'outlook_reply_all': replyAllSchema,
  'outlook_forward_email': forwardEmailSchema,
  'outlook_delete_email': deleteEmailSchema,
  'outlook_move_email': moveEmailSchema,
  'outlook_mark_as_read': markAsReadSchema,
  'outlook_flag_email': flagEmailSchema,
  'outlook_categorize_email': categorizeEmailSchema,
  'outlook_archive_email': archiveEmailSchema,
  'outlook_batch_process_emails': batchProcessEmailsSchema,
};