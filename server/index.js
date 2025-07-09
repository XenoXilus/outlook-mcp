#!/usr/bin/env node

import 'dotenv/config';
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { ListToolsRequestSchema, CallToolRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import { OutlookAuthManager } from './auth/auth.js';
import { 
  listEmailsTool,
  sendEmailTool,
  listEventsTool,
  createEventTool,
  getEmailTool,
  searchEmailsTool,
  createDraftTool,
  replyToEmailTool,
  replyAllTool,
  forwardEmailTool,
  deleteEmailTool,
  // Email Management Tools
  moveEmailTool,
  markAsReadTool,
  flagEmailTool,
  categorizeEmailTool,
  archiveEmailTool,
  batchProcessEmailsTool,
  // Folder Management Tools
  listFoldersTool,
  createFolderTool,
  renameFolderTool,
  getFolderStatsTool,
  // Attachment Tools
  listAttachmentsTool,
  downloadAttachmentTool,
  addAttachmentTool,
  scanAttachmentsTool
} from './tools/index.js';

const server = new Server(
  {
    name: 'outlook-mcp',
    version: '1.0.0',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

const authManager = new OutlookAuthManager(
  process.env.AZURE_CLIENT_ID,
  process.env.AZURE_TENANT_ID
);

server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
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
      },
      {
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
      },
      {
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
      },
      {
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
      },
      {
        name: 'outlook_get_email',
        description: 'Get detailed information about a specific email',
        inputSchema: {
          type: 'object',
          properties: {
            messageId: {
              type: 'string',
              description: 'The ID of the email message to retrieve',
            },
          },
          required: ['messageId'],
        },
      },
      {
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
              default: true,
            },
            orderBy: {
              type: 'string',
              description: 'Sort order (e.g., "receivedDateTime desc")',
              default: 'receivedDateTime desc',
            },
          },
        },
      },
      {
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
      },
      {
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
      },
      {
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
      },
      {
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
      },
      {
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
      },
      // Email Management Tools
      {
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
      },
      {
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
      },
      {
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
      },
      {
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
      },
      {
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
      },
      {
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
                destinationFolderId: { type: 'string' },
                flagStatus: { type: 'string' },
                categories: { type: 'array', items: { type: 'string' } },
                permanentDelete: { type: 'boolean' },
              },
            },
          },
          required: ['messageIds', 'operation'],
        },
      },
      // Folder Management Tools
      {
        name: 'outlook_list_folders',
        description: 'List all email folders',
        inputSchema: {
          type: 'object',
          properties: {
            includeHidden: {
              type: 'boolean',
              description: 'Include hidden folders',
              default: false,
            },
            includeChildFolders: {
              type: 'boolean',
              description: 'Include nested child folders',
              default: true,
            },
            top: {
              type: 'number',
              description: 'Maximum number of folders to return',
              default: 100,
            },
          },
        },
      },
      {
        name: 'outlook_create_folder',
        description: 'Create a new email folder',
        inputSchema: {
          type: 'object',
          properties: {
            displayName: {
              type: 'string',
              description: 'Name of the new folder',
            },
            parentFolderId: {
              type: 'string',
              description: 'ID of parent folder (optional, creates at root level if not specified)',
            },
          },
          required: ['displayName'],
        },
      },
      {
        name: 'outlook_rename_folder',
        description: 'Rename an existing email folder',
        inputSchema: {
          type: 'object',
          properties: {
            folderId: {
              type: 'string',
              description: 'ID of the folder to rename',
            },
            newDisplayName: {
              type: 'string',
              description: 'New name for the folder',
            },
          },
          required: ['folderId', 'newDisplayName'],
        },
      },
      {
        name: 'outlook_get_folder_stats',
        description: 'Get statistics for a specific folder',
        inputSchema: {
          type: 'object',
          properties: {
            folderId: {
              type: 'string',
              description: 'ID of the folder to get stats for',
            },
            includeSubfolders: {
              type: 'boolean',
              description: 'Include statistics for subfolders',
              default: true,
            },
          },
          required: ['folderId'],
        },
      },
      // Attachment Tools
      {
        name: 'outlook_list_attachments',
        description: 'List all attachments for a specific email',
        inputSchema: {
          type: 'object',
          properties: {
            messageId: {
              type: 'string',
              description: 'The ID of the email to list attachments for',
            },
          },
          required: ['messageId'],
        },
      },
      {
        name: 'outlook_download_attachment',
        description: 'Download a specific email attachment',
        inputSchema: {
          type: 'object',
          properties: {
            messageId: {
              type: 'string',
              description: 'The ID of the email containing the attachment',
            },
            attachmentId: {
              type: 'string',
              description: 'The ID of the attachment to download',
            },
            includeContent: {
              type: 'boolean',
              description: 'Whether to include the base64-encoded file content',
              default: false,
            },
          },
          required: ['messageId', 'attachmentId'],
        },
      },
      {
        name: 'outlook_add_attachment',
        description: 'Add an attachment to an email draft',
        inputSchema: {
          type: 'object',
          properties: {
            messageId: {
              type: 'string',
              description: 'The ID of the email (draft) to add attachment to',
            },
            name: {
              type: 'string',
              description: 'Name of the attachment file',
            },
            contentType: {
              type: 'string',
              description: 'MIME type of the attachment',
            },
            contentBytes: {
              type: 'string',
              description: 'Base64-encoded content of the attachment',
            },
          },
          required: ['messageId', 'name', 'contentType', 'contentBytes'],
        },
      },
      {
        name: 'outlook_scan_attachments',
        description: 'Scan emails for large or suspicious attachments',
        inputSchema: {
          type: 'object',
          properties: {
            folder: {
              type: 'string',
              description: 'Folder to scan (default: inbox)',
              default: 'inbox',
            },
            maxSizeMB: {
              type: 'number',
              description: 'Maximum attachment size in MB to flag as large',
              default: 10,
            },
            suspiciousTypes: {
              type: 'array',
              items: { type: 'string' },
              description: 'File extensions to flag as suspicious',
              default: ['exe', 'bat', 'cmd', 'scr', 'vbs', 'js'],
            },
            limit: {
              type: 'number',
              description: 'Maximum number of emails to scan',
              default: 100,
            },
            daysBack: {
              type: 'number',
              description: 'How many days back to scan',
              default: 30,
            },
          },
        },
      },
    ],
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    switch (name) {
      case 'outlook_list_emails':
        return await listEmailsTool(authManager, args);
      
      case 'outlook_send_email':
        return await sendEmailTool(authManager, args);
      
      case 'outlook_list_events':
        return await listEventsTool(authManager, args);
      
      case 'outlook_create_event':
        return await createEventTool(authManager, args);
      
      case 'outlook_get_email':
        return await getEmailTool(authManager, args);
      
      case 'outlook_search_emails':
        return await searchEmailsTool(authManager, args);
      
      case 'outlook_create_draft':
        return await createDraftTool(authManager, args);
      
      case 'outlook_reply_to_email':
        return await replyToEmailTool(authManager, args);
      
      case 'outlook_reply_all':
        return await replyAllTool(authManager, args);
      
      case 'outlook_forward_email':
        return await forwardEmailTool(authManager, args);
      
      case 'outlook_delete_email':
        return await deleteEmailTool(authManager, args);
      
      // Email Management Tools
      case 'outlook_move_email':
        return await moveEmailTool(authManager, args);
      
      case 'outlook_mark_as_read':
        return await markAsReadTool(authManager, args);
      
      case 'outlook_flag_email':
        return await flagEmailTool(authManager, args);
      
      case 'outlook_categorize_email':
        return await categorizeEmailTool(authManager, args);
      
      case 'outlook_archive_email':
        return await archiveEmailTool(authManager, args);
      
      case 'outlook_batch_process_emails':
        return await batchProcessEmailsTool(authManager, args);
      
      // Folder Management Tools
      case 'outlook_list_folders':
        return await listFoldersTool(authManager, args);
      
      case 'outlook_create_folder':
        return await createFolderTool(authManager, args);
      
      case 'outlook_rename_folder':
        return await renameFolderTool(authManager, args);
      
      case 'outlook_get_folder_stats':
        return await getFolderStatsTool(authManager, args);
      
      // Attachment Tools
      case 'outlook_list_attachments':
        return await listAttachmentsTool(authManager, args);
      
      case 'outlook_download_attachment':
        return await downloadAttachmentTool(authManager, args);
      
      case 'outlook_add_attachment':
        return await addAttachmentTool(authManager, args);
      
      case 'outlook_scan_attachments':
        return await scanAttachmentsTool(authManager, args);
      
      default:
        throw new Error(`Unknown tool: ${name}`);
    }
  } catch (error) {
    return {
      error: {
        code: 'TOOL_ERROR',
        message: error.message,
      },
    };
  }
});

async function main() {
  if (!process.env.AZURE_CLIENT_ID || !process.env.AZURE_TENANT_ID) {
    console.error('Error: AZURE_CLIENT_ID and AZURE_TENANT_ID environment variables are required.');
    console.error('Please set these in your MCP server configuration.');
    console.error('Note: This server uses OAuth 2.0 with PKCE for secure delegated authentication.');
    process.exit(1);
  }

  console.error('Starting Outlook MCP Server...');
  console.error('Authenticating with Microsoft Office 365...');
  console.error('You will be prompted to select or add an Office 365 account.');
  
  try {
    const authResult = await authManager.authenticate();
    
    if (!authResult.success) {
      console.error('\nAuthentication failed:', authResult.error || 'Unknown error');
      console.error('Please disconnect and reconnect this MCP server to try again.');
      process.exit(1);
    }
    
    console.error(`\nSuccessfully authenticated as: ${authResult.user.mail || authResult.user.displayName}`);
    console.error('To use a different account, disconnect and reconnect this MCP server.\n');
    
  } catch (error) {
    console.error('\nAuthentication error:', error.message);
    console.error('Please disconnect and reconnect this MCP server to try again.');
    process.exit(1);
  }

  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Outlook MCP server is ready');
}

main().catch((error) => {
  console.error('Server error:', error);
  process.exit(1);
});