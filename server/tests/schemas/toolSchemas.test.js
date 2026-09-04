import { describe, it, expect } from 'vitest';
import {
  allToolSchemas,
  allToolSchemaMap,
  schemasByCategory,
  getSchemaByName,
  getSchemasByCategory,
  getToolNamesByCategory,
  getAllToolNames,
  validateSchemas,
  getSchemaStats
} from '../../schemas/toolSchemas.js';

describe('Tool Schemas', () => {
  describe('Schema Structure', () => {
    it('should export all tool schemas', () => {
      expect(allToolSchemas).toBeDefined();
      expect(Array.isArray(allToolSchemas)).toBe(true);
      expect(allToolSchemas.length).toBeGreaterThan(0);
    });

    it('should export schema map', () => {
      expect(allToolSchemaMap).toBeDefined();
      expect(typeof allToolSchemaMap).toBe('object');
      expect(Object.keys(allToolSchemaMap).length).toBeGreaterThan(0);
    });

    it('should organize schemas by category', () => {
      expect(schemasByCategory).toBeDefined();
      expect(schemasByCategory.email).toBeDefined();
      expect(schemasByCategory.calendar).toBeDefined();
      expect(schemasByCategory.folder).toBeDefined();
      expect(schemasByCategory.attachment).toBeDefined();
    });

    it('should have consistent schema structure', () => {
      allToolSchemas.forEach(schema => {
        expect(schema).toHaveProperty('name');
        expect(schema).toHaveProperty('description');
        expect(schema).toHaveProperty('inputSchema');
        expect(schema.inputSchema).toHaveProperty('type');
        expect(schema.inputSchema).toHaveProperty('properties');
        expect(typeof schema.name).toBe('string');
        expect(typeof schema.description).toBe('string');
        expect(schema.inputSchema.type).toBe('object');
      });
    });
  });

  describe('Schema Validation', () => {
    it('should validate all schemas successfully', () => {
      const errors = validateSchemas();
      expect(errors).toEqual([]);
    });

    it('should have valid email schemas', () => {
      const emailSchemas = getSchemasByCategory('email');
      expect(emailSchemas.length).toBeGreaterThan(0);
      
      emailSchemas.forEach(schema => {
        expect(schema.name).toMatch(/^outlook_/);
        expect(schema.description).toBeTruthy();
        expect(schema.inputSchema.type).toBe('object');
      });
    });

    it('should have valid calendar schemas', () => {
      const calendarSchemas = getSchemasByCategory('calendar');
      expect(calendarSchemas.length).toBeGreaterThan(0);
      
      calendarSchemas.forEach(schema => {
        expect(schema.name).toMatch(/^outlook_/);
        expect(schema.description).toBeTruthy();
        expect(schema.inputSchema.type).toBe('object');
      });
    });

    it('should have valid folder schemas', () => {
      const folderSchemas = getSchemasByCategory('folder');
      expect(folderSchemas.length).toBeGreaterThan(0);
      
      folderSchemas.forEach(schema => {
        expect(schema.name).toMatch(/^outlook_/);
        expect(schema.description).toBeTruthy();
        expect(schema.inputSchema.type).toBe('object');
      });
    });

    it('should have valid attachment schemas', () => {
      const attachmentSchemas = getSchemasByCategory('attachment');
      expect(attachmentSchemas.length).toBeGreaterThan(0);
      
      attachmentSchemas.forEach(schema => {
        expect(schema.name).toMatch(/^outlook_/);
        expect(schema.description).toBeTruthy();
        expect(schema.inputSchema.type).toBe('object');
      });
    });
  });

  describe('Schema Lookup Functions', () => {
    it('should get schema by name', () => {
      const schema = getSchemaByName('outlook_send_email');
      expect(schema).toBeDefined();
      expect(schema.name).toBe('outlook_send_email');
      expect(schema.description).toBeTruthy();
    });

    it('should return undefined for non-existent schema', () => {
      const schema = getSchemaByName('non_existent_tool');
      expect(schema).toBeUndefined();
    });

    it('should get schemas by category', () => {
      const emailSchemas = getSchemasByCategory('email');
      expect(emailSchemas).toBeDefined();
      expect(Array.isArray(emailSchemas)).toBe(true);
      expect(emailSchemas.length).toBeGreaterThan(0);
    });

    it('should return empty array for non-existent category', () => {
      const schemas = getSchemasByCategory('non_existent');
      expect(schemas).toEqual([]);
    });

    it('should get tool names by category', () => {
      const emailToolNames = getToolNamesByCategory('email');
      expect(emailToolNames).toBeDefined();
      expect(Array.isArray(emailToolNames)).toBe(true);
      expect(emailToolNames.length).toBeGreaterThan(0);
      expect(emailToolNames).toContain('outlook_send_email');
    });

    it('should get all tool names', () => {
      const allToolNames = getAllToolNames();
      expect(allToolNames).toBeDefined();
      expect(Array.isArray(allToolNames)).toBe(true);
      expect(allToolNames.length).toBeGreaterThan(0);
      expect(allToolNames).toContain('outlook_send_email');
      expect(allToolNames).toContain('outlook_list_events');
    });
  });

  describe('Schema Statistics', () => {
    it('should provide schema statistics', () => {
      const stats = getSchemaStats();
      expect(stats).toBeDefined();
      expect(stats.totalSchemas).toBeGreaterThan(0);
      expect(stats.schemasByCategory).toBeDefined();
      expect(Array.isArray(stats.schemasByCategory)).toBe(true);
      expect(stats.requiredParameters).toBeGreaterThan(0);
      expect(stats.optionalParameters).toBeGreaterThan(0);
    });

    it('should have correct schema counts', () => {
      const stats = getSchemaStats();
      const totalFromCategories = stats.schemasByCategory.reduce((sum, cat) => sum + cat.count, 0);
      expect(totalFromCategories).toBe(stats.totalSchemas);
    });
  });

  describe('Specific Schema Tests', () => {
    it('should have correct send email schema structure', () => {
      const schema = getSchemaByName('outlook_send_email');
      expect(schema).toBeDefined();
      expect(schema.inputSchema.required).toContain('to');
      expect(schema.inputSchema.required).toContain('subject');
      expect(schema.inputSchema.required).toContain('body');
      expect(schema.inputSchema.properties.to.type).toBe('array');
      expect(schema.inputSchema.properties.bodyType.enum).toContain('text');
      expect(schema.inputSchema.properties.bodyType.enum).toContain('html');
    });

    it('should have correct create event schema structure', () => {
      const schema = getSchemaByName('outlook_create_event');
      expect(schema).toBeDefined();
      expect(schema.inputSchema.required).toContain('subject');
      expect(schema.inputSchema.required).toContain('start');
      expect(schema.inputSchema.required).toContain('end');
      expect(schema.inputSchema.properties.start.type).toBe('object');
      expect(schema.inputSchema.properties.start.required).toContain('dateTime');
      expect(schema.inputSchema.properties.start.required).toContain('timeZone');
    });

    it('should have correct list attachments schema structure', () => {
      const schema = getSchemaByName('outlook_list_attachments');
      expect(schema).toBeDefined();
      expect(schema.inputSchema.required).toContain('messageId');
      expect(schema.inputSchema.properties.messageId.type).toBe('string');
    });

    it('should have correct folder creation schema structure', () => {
      const schema = getSchemaByName('outlook_create_folder');
      expect(schema).toBeDefined();
      expect(schema.inputSchema.required).toContain('displayName');
      expect(schema.inputSchema.properties.displayName.type).toBe('string');
      expect(schema.inputSchema.properties.parentFolderId).toBeDefined();
    });
  });

  describe('Schema Completeness', () => {
    it('should cover all expected email operations', () => {
      const emailToolNames = getToolNamesByCategory('email');
      const expectedEmailTools = [
        'outlook_list_emails',
        'outlook_send_email',
        'outlook_get_email',
        'outlook_search_emails',
        'outlook_create_draft',
        'outlook_reply_to_email',
        'outlook_reply_all',
        'outlook_forward_email',
        'outlook_delete_email',
        'outlook_move_email',
        'outlook_mark_as_read',
        'outlook_flag_email',
        'outlook_categorize_email',
        'outlook_archive_email',
        'outlook_batch_process_emails'
      ];

      expectedEmailTools.forEach(tool => {
        expect(emailToolNames).toContain(tool);
      });
    });

    it('should cover all expected calendar operations', () => {
      const calendarToolNames = getToolNamesByCategory('calendar');
      const expectedCalendarTools = [
        'outlook_list_events',
        'outlook_create_event'
      ];

      expectedCalendarTools.forEach(tool => {
        expect(calendarToolNames).toContain(tool);
      });
    });

    it('should cover all expected folder operations', () => {
      const folderToolNames = getToolNamesByCategory('folder');
      const expectedFolderTools = [
        'outlook_list_folders',
        'outlook_create_folder',
        'outlook_rename_folder',
        'outlook_get_folder_stats'
      ];

      expectedFolderTools.forEach(tool => {
        expect(folderToolNames).toContain(tool);
      });
    });

    it('should cover all expected attachment operations', () => {
      const attachmentToolNames = getToolNamesByCategory('attachment');
      const expectedAttachmentTools = [
        'outlook_list_attachments',
        'outlook_download_attachment',
        'outlook_add_attachment',
        'outlook_scan_attachments'
      ];

      expectedAttachmentTools.forEach(tool => {
        expect(attachmentToolNames).toContain(tool);
      });
    });
  });

  describe('Schema Consistency', () => {
    it('should have consistent naming convention', () => {
      const allToolNames = getAllToolNames();
      allToolNames.forEach(name => {
        expect(name).toMatch(/^outlook_[a-z_]+$/);
      });
    });

    it('should have array and map consistency', () => {
      const arrayCount = allToolSchemas.length;
      const mapCount = Object.keys(allToolSchemaMap).length;
      expect(arrayCount).toBe(mapCount);
    });

    it('should have schema names match map keys', () => {
      allToolSchemas.forEach(schema => {
        expect(allToolSchemaMap[schema.name]).toBeDefined();
        expect(allToolSchemaMap[schema.name]).toBe(schema);
      });
    });

    // Calendar and SharePoint are deliberately personal-only in v1.3, so they
    // are excluded. Every other category is mailbox-aware and must advertise it.
    const MAILBOX_AWARE_CATEGORIES = ['email', 'folder', 'attachment', 'receipts'];

    // Discovery tools enumerate mailboxes instead of operating inside one, so they take no `mailbox`.
    const MAILBOX_ENUMERATION_TOOLS = new Set(['outlook_list_shared_mailboxes']);

    it.each(MAILBOX_AWARE_CATEGORIES)('every %s schema declares a mailbox property', (category) => {
      const schemas = getSchemasByCategory(category);
      expect(schemas.length).toBeGreaterThan(0);

      const missing = schemas
        .filter(schema => !MAILBOX_ENUMERATION_TOOLS.has(schema.name))
        .filter(schema => !schema.inputSchema?.properties?.mailbox)
        .map(schema => schema.name);
      expect(missing, `${category} schemas missing a mailbox property — add ...mailboxProperty`).toEqual([]);
    });
  });
});