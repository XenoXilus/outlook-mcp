/**
 * Centralized tool schemas for Outlook MCP Server
 * 
 * This module exports all MCP tool schemas in a unified format for easy consumption
 * by the server. Schemas are organized by category and include validation rules.
 */

import { 
  emailSchemas, 
  emailSchemaMap 
} from './emailSchemas.js';

import { 
  calendarSchemas, 
  calendarSchemaMap 
} from './calendarSchemas.js';

import { 
  folderSchemas, 
  folderSchemaMap 
} from './folderSchemas.js';

import { 
  attachmentSchemas, 
  attachmentSchemaMap 
} from './attachmentSchemas.js';

import { 
  sharePointSchemas, 
  sharePointSchemaMap 
} from './sharePointSchemas.js';

/**
 * Complete array of all tool schemas
 */
export const allToolSchemas = [
  ...emailSchemas,
  ...calendarSchemas,
  ...folderSchemas,
  ...attachmentSchemas,
  ...sharePointSchemas,
];

/**
 * Complete mapping of tool names to schemas
 */
export const allToolSchemaMap = {
  ...emailSchemaMap,
  ...calendarSchemaMap,
  ...folderSchemaMap,
  ...attachmentSchemaMap,
  ...sharePointSchemaMap,
};

/**
 * Schemas organized by category
 */
export const schemasByCategory = {
  email: emailSchemas,
  calendar: calendarSchemas,
  folder: folderSchemas,
  attachment: attachmentSchemas,
  sharepoint: sharePointSchemas,
};

/**
 * Schema maps organized by category
 */
export const schemaMaps = {
  email: emailSchemaMap,
  calendar: calendarSchemaMap,
  folder: folderSchemaMap,
  attachment: attachmentSchemaMap,
  sharepoint: sharePointSchemaMap,
};

/**
 * Get schema by tool name
 */
export function getSchemaByName(toolName) {
  return allToolSchemaMap[toolName];
}

/**
 * Get all schemas for a specific category
 */
export function getSchemasByCategory(category) {
  return schemasByCategory[category] || [];
}

/**
 * Get tool names by category
 */
export function getToolNamesByCategory(category) {
  const schemas = getSchemasByCategory(category);
  return schemas.map(schema => schema.name);
}

/**
 * Get all tool names
 */
export function getAllToolNames() {
  return allToolSchemas.map(schema => schema.name);
}

/**
 * Validate that all schemas have required properties
 */
export function validateSchemas() {
  const errors = [];
  
  allToolSchemas.forEach(schema => {
    if (!schema.name) {
      errors.push(`Schema missing name: ${JSON.stringify(schema)}`);
    }
    
    if (!schema.description) {
      errors.push(`Schema ${schema.name} missing description`);
    }
    
    if (!schema.inputSchema) {
      errors.push(`Schema ${schema.name} missing inputSchema`);
    }
    
    if (schema.inputSchema && !schema.inputSchema.type) {
      errors.push(`Schema ${schema.name} inputSchema missing type`);
    }
  });
  
  return errors;
}

/**
 * Get schema statistics
 */
export function getSchemaStats() {
  return {
    totalSchemas: allToolSchemas.length,
    schemasByCategory: Object.keys(schemasByCategory).map(category => ({
      category,
      count: schemasByCategory[category].length
    })),
    requiredParameters: allToolSchemas.reduce((acc, schema) => {
      const required = schema.inputSchema?.required || [];
      return acc + required.length;
    }, 0),
    optionalParameters: allToolSchemas.reduce((acc, schema) => {
      const properties = schema.inputSchema?.properties || {};
      const required = schema.inputSchema?.required || [];
      const optional = Object.keys(properties).filter(prop => !required.includes(prop));
      return acc + optional.length;
    }, 0)
  };
}

// Export individual schema modules for direct access
export { 
  emailSchemas, 
  emailSchemaMap 
} from './emailSchemas.js';

export { 
  calendarSchemas, 
  calendarSchemaMap 
} from './calendarSchemas.js';

export { 
  folderSchemas, 
  folderSchemaMap 
} from './folderSchemas.js';

export { 
  attachmentSchemas, 
  attachmentSchemaMap 
} from './attachmentSchemas.js';

export { 
  sharePointSchemas, 
  sharePointSchemaMap 
} from './sharePointSchemas.js';
