// Barrel exports for all tools (replaces the monolithic file)

// Common utilities
export { 
  clearStylingCache, 
  clearSignatureCache, 
  getStylingCacheStats,
  applyUserStyling,
  stylingCache,
  signatureCache
} from './common/sharedUtils.js';

// Email tools
export { listEmailsTool, getEmailTool } from './email/listEmails.js';
export { sendEmailTool } from './email/sendEmail.js';
export { searchEmailsTool } from './email/searchEmails.js';
export { createDraftTool } from './email/createDraft.js';
export { replyToEmailTool, replyAllTool } from './email/replyEmail.js';
export { forwardEmailTool } from './email/forwardEmail.js';
export { 
  deleteEmailTool, 
  moveEmailTool, 
  markAsReadTool, 
  flagEmailTool, 
  categorizeEmailTool, 
  archiveEmailTool, 
  batchProcessEmailsTool 
} from './email/emailManagement.js';

// Calendar tools
export { listEventsTool } from './calendar/listEvents.js';
export { createEventTool } from './calendar/createEvent.js';
export { 
  getEventTool, 
  updateEventTool, 
  deleteEventTool, 
  respondToInviteTool, 
  validateEventDateTimesTool 
} from './calendar/eventManagement.js';
export { 
  createRecurringEventTool,
  findMeetingTimesTool,
  checkAvailabilityTool,
  scheduleOnlineMeetingTool,
  listCalendarsTool,
  getCalendarViewTool,
  getBusyTimesTool,
  buildRecurrencePatternTool,
  createRecurrenceHelperTool,
  checkCalendarPermissionsTool
} from './calendar/calendarUtils.js';

// Folder tools
export { listFoldersTool } from './folders/listFolders.js';
export { createFolderTool } from './folders/createFolder.js';
export { renameFolderTool } from './folders/renameFolder.js';
export { getFolderStatsTool } from './folders/getFolderStats.js';

// Attachment tools
export { listAttachmentsTool } from './attachments/listAttachments.js';
export { downloadAttachmentTool } from './attachments/downloadAttachment.js';
export { addAttachmentTool } from './attachments/addAttachment.js';
export { scanAttachmentsTool } from './attachments/scanAttachments.js';

// Utility tools
export { getRateLimitMetricsTool, resetRateLimitMetricsTool } from './common/rateLimitUtils.js';

// SharePoint tools
export { getSharePointFileTool, listSharePointFilesTool, resolveSharePointLinkTool } from './sharepoint/getSharePointFile.js';
