/**
 * Folder-related MCP tool schemas
 * 
 * This module contains all JSON schemas for folder operations in the Outlook MCP server.
 * Includes folder management, statistics, and organization functionality.
 */
import { mailboxProperty } from './sharedSchemaFragments.js';

export const listFoldersSchema = {
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
      ...mailboxProperty,
    },
  },
};

export const createFolderSchema = {
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
      ...mailboxProperty,
    },
    required: ['displayName'],
  },
};

export const renameFolderSchema = {
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
      ...mailboxProperty,
    },
    required: ['folderId', 'newDisplayName'],
  },
};

export const getFolderStatsSchema = {
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
      ...mailboxProperty,
    },
    required: ['folderId'],
  },
};

// Export all folder schemas as an array for easy iteration
export const folderSchemas = [
  listFoldersSchema,
  createFolderSchema,
  renameFolderSchema,
  getFolderStatsSchema,
];

// Export mapping for quick lookup
export const folderSchemaMap = {
  'outlook_list_folders': listFoldersSchema,
  'outlook_create_folder': createFolderSchema,
  'outlook_rename_folder': renameFolderSchema,
  'outlook_get_folder_stats': getFolderStatsSchema,
};