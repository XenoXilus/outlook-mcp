/**
 * Folder-related MCP tool schemas
 * 
 * This module contains all JSON schemas for folder operations in the Outlook MCP server.
 * Includes folder management, statistics, and organization functionality.
 */

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
    },
    required: ['folderId'],
  },
};

export const listRulesSchema = {
  name: 'outlook_list_rules',
  description: 'List all inbox message rules',
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

export const createRuleSchema = {
  name: 'outlook_create_rule',
  description: 'Create an inbox message rule to automatically move emails matching sender criteria to a specified folder',
  inputSchema: {
    type: 'object',
    properties: {
      displayName: {
        type: 'string',
        description: 'Name of the rule',
      },
      senderContains: {
        type: 'array',
        items: { type: 'string' },
        description: 'List of strings to match against sender email address (e.g. ["bizreach"])',
      },
      moveToFolder: {
        type: 'string',
        description: 'ID of the destination folder',
      },
      isEnabled: {
        type: 'boolean',
        description: 'Whether the rule is enabled (default: true)',
        default: true,
      },
      sequence: {
        type: 'number',
        description: 'Order in which rule is applied (default: 1)',
        default: 1,
      },
    },
    required: ['displayName', 'senderContains', 'moveToFolder'],
  },
};

export const deleteRuleSchema = {
  name: 'outlook_delete_rule',
  description: 'Delete an inbox message rule by ID',
  inputSchema: {
    type: 'object',
    properties: {
      ruleId: {
        type: 'string',
        description: 'ID of the rule to delete',
      },
    },
    required: ['ruleId'],
  },
};

// Export all folder schemas as an array for easy iteration
export const folderSchemas = [
  listFoldersSchema,
  createFolderSchema,
  renameFolderSchema,
  getFolderStatsSchema,
  listRulesSchema,
  createRuleSchema,
  deleteRuleSchema,
];

// Export mapping for quick lookup
export const folderSchemaMap = {
  'outlook_list_folders': listFoldersSchema,
  'outlook_create_folder': createFolderSchema,
  'outlook_rename_folder': renameFolderSchema,
  'outlook_get_folder_stats': getFolderStatsSchema,
  'outlook_list_rules': listRulesSchema,
  'outlook_create_rule': createRuleSchema,
  'outlook_delete_rule': deleteRuleSchema,
};