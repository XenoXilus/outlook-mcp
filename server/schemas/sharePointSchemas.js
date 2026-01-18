/**
 * SharePoint-related MCP tool schemas
 * 
 * This module contains all JSON schemas for SharePoint file operations using the same
 * authenticated session as Outlook.
 */

export const getSharePointFileSchema = {
  name: 'outlook_get_sharepoint_file',
  description: 'Fetch a SharePoint file using the same authenticated session as Outlook. Handles sharing links from emails. Either sharePointUrl OR fileId must be provided.',
  inputSchema: {
    type: 'object',
    properties: {
      sharePointUrl: {
        type: 'string',
        description: 'SharePoint sharing URL from email (e.g., https://company.sharepoint.com/:w:/s/sitename/...). Required if fileId is not provided.',
      },
      fileId: {
        type: 'string',
        description: 'Direct file ID if known. Required if sharePointUrl is not provided.',
      },
      driveId: {
        type: 'string',
        description: 'Drive ID for direct file access (defaults to user\'s OneDrive)',
      },
      downloadContent: {
        type: 'boolean',
        description: 'Whether to download and include file content as base64 (max 50MB)',
        default: false,
      },
    },
  },
};

export const listSharePointFilesSchema = {
  name: 'outlook_list_sharepoint_files',
  description: 'List files in SharePoint sites or OneDrive folders using the same authenticated session',
  inputSchema: {
    type: 'object',
    properties: {
      siteId: {
        type: 'string',
        description: 'SharePoint site ID (optional)',
      },
      driveId: {
        type: 'string',
        description: 'Drive ID (defaults to user\'s OneDrive if not specified)',
      },
      folderId: {
        type: 'string',
        description: 'Specific folder ID to list contents of',
      },
      limit: {
        type: 'number',
        description: 'Maximum number of files to return',
        default: 50,
        minimum: 1,
        maximum: 200,
      },
      orderBy: {
        type: 'string',
        description: 'Field to order results by',
        default: 'name',
        enum: ['name', 'lastModifiedDateTime', 'size', 'createdDateTime'],
      },
    },
  },
};

export const resolveSharePointLinkSchema = {
  name: 'outlook_resolve_sharepoint_link',
  description: 'Resolve SharePoint sharing links from emails to get file metadata without downloading',
  inputSchema: {
    type: 'object',
    properties: {
      sharePointUrl: {
        type: 'string',
        description: 'SharePoint sharing URL to resolve',
      },
      includePermissions: {
        type: 'boolean',
        description: 'Whether to include sharing permissions information',
        default: false,
      },
    },
    required: ['sharePointUrl'],
  },
};

// Export all SharePoint schemas as an array for easy iteration
export const sharePointSchemas = [
  getSharePointFileSchema,
  listSharePointFilesSchema,
  resolveSharePointLinkSchema,
];

// Export mapping for quick lookup
export const sharePointSchemaMap = {
  'outlook_get_sharepoint_file': getSharePointFileSchema,
  'outlook_list_sharepoint_files': listSharePointFilesSchema,
  'outlook_resolve_sharepoint_link': resolveSharePointLinkSchema,
};
