/**
 * Attachment-related MCP tool schemas
 * 
 * This module contains all JSON schemas for attachment operations in the Outlook MCP server.
 * Includes attachment management, download, scanning, and security functionality.
 */

export const listAttachmentsSchema = {
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
};

export const downloadAttachmentSchema = {
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
        description: 'Whether to include the file content',
        default: false,
      },
      decodeContent: {
        type: 'boolean',
        description: 'Whether to decode Base64 content to readable format (text files) or provide summary (binary files)',
        default: true,
      },
    },
    required: ['messageId', 'attachmentId'],
  },
};

export const addAttachmentSchema = {
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
};

export const scanAttachmentsSchema = {
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
};

// Export all attachment schemas as an array for easy iteration
export const attachmentSchemas = [
  listAttachmentsSchema,
  downloadAttachmentSchema,
  addAttachmentSchema,
  scanAttachmentsSchema,
];

// Export mapping for quick lookup
export const attachmentSchemaMap = {
  'outlook_list_attachments': listAttachmentsSchema,
  'outlook_download_attachment': downloadAttachmentSchema,
  'outlook_add_attachment': addAttachmentSchema,
  'outlook_scan_attachments': scanAttachmentsSchema,
};