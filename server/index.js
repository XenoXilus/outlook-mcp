#!/usr/bin/env node

// Add global error handlers for debugging
process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled Rejection at:', promise, 'reason:', reason);
  console.error('Stack trace:', reason.stack || 'No stack trace available');
  // Don't exit immediately - let the MCP server handle errors gracefully
});

process.on('uncaughtException', (error) => {
  console.error('Uncaught Exception:', error);
  console.error('Stack trace:', error.stack);
  // Only exit on truly fatal errors
  if (error.code === 'MODULE_NOT_FOUND' || error.name === 'SyntaxError') {
    process.exit(1);
  }
});

// Main initialization function using IIFE to handle async imports
(async function initializeServer() {
  console.error('Debug: Script starting...');

  try {
    console.error('Debug: Loading dotenv...');
    await import('dotenv/config');

    console.error('Debug: Loading MCP SDK...');
    const { Server } = await import('@modelcontextprotocol/sdk/server/index.js');
    const { StdioServerTransport } = await import('@modelcontextprotocol/sdk/server/stdio.js');
    const {
      ListToolsRequestSchema,
      CallToolRequestSchema,
      InitializeRequestSchema,
      InitializedNotificationSchema,
      ListPromptsRequestSchema,
      GetPromptRequestSchema
    } = await import('@modelcontextprotocol/sdk/types.js');

    console.error('Debug: Loading auth manager...');
    const { OutlookAuthManager } = await import('./auth/auth.js');

    console.error('Debug: Loading MCP error utilities...');
    const { createToolError, createProtocolError, ErrorCodes, convertErrorToToolError } = await import('./utils/mcpErrorResponse.js');

    console.error('Debug: Loading tool schemas...');
    const { allToolSchemas } = await import('./schemas/toolSchemas.js');
    console.error(`Debug: Loaded ${allToolSchemas.length} tool schemas`);

    console.error('Debug: Loading tools...');
    const tools = await import('./tools/index.js');
    console.error('Debug: Tools imported, available:', Object.keys(tools).length);

    console.error('Debug: Loading prompts...');
    const { promptList, getPrompt } = await import('./prompts/index.js');

    // Extract the specific tools we need
    const {
      listEmailsTool,
      sendEmailTool,
      listEventsTool,
      createEventTool,
      getEventTool,
      updateEventTool,
      deleteEventTool,
      respondToInviteTool,
      validateEventDateTimesTool,
      createRecurringEventTool,
      findMeetingTimesTool,
      checkAvailabilityTool,
      scheduleOnlineMeetingTool,
      listCalendarsTool,
      getCalendarViewTool,
      getBusyTimesTool,
      buildRecurrencePatternTool,
      createRecurrenceHelperTool,
      checkCalendarPermissionsTool,
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
      scanAttachmentsTool,
      // SharePoint Tools
      getSharePointFileTool,
      listSharePointFilesTool,
      resolveSharePointLinkTool
    } = tools;

    console.error('Debug: All required tools extracted successfully');
    console.error('Debug: All imports successful');

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

    server.setRequestHandler(InitializeRequestSchema, async (request) => {
      console.error('Debug: Handling MCP initialization...');
      console.error('Debug: Initialize request:', JSON.stringify(request, null, 2));

      const response = {
        protocolVersion: '2025-06-18',
        capabilities: {
          tools: {},
          prompts: {},
        },
        serverInfo: {
          name: 'outlook-mcp',
          version: '1.0.0',
        },
      };

      console.error('Debug: Initialize response:', JSON.stringify(response, null, 2));
      return response;
    });

    server.setNotificationHandler(InitializedNotificationSchema, async () => {
      console.error('Debug: Client initialized');
    });

    server.setRequestHandler(ListToolsRequestSchema, async () => {
      console.error(`Debug: Returning ${allToolSchemas.length} tools to client`);
      return {
        tools: allToolSchemas,
      };
    });

    server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;
      console.error(`DEBUG Tool Dispatch: Called tool '${name}' with args:`, JSON.stringify(args, null, 2));

      try {
        switch (name) {
          case 'outlook_list_emails':
            console.error(`DEBUG: Calling listEmailsTool`);
            return await listEmailsTool(authManager, args);

          case 'outlook_send_email':
            console.error(`DEBUG: Calling sendEmailTool`);
            return await sendEmailTool(authManager, args);

          case 'outlook_list_events':
            console.error(`DEBUG: Calling listEventsTool`);
            return await listEventsTool(authManager, args);

          case 'outlook_create_event':
            console.error(`DEBUG: Calling createEventTool`);
            return await createEventTool(authManager, args);

          case 'outlook_get_event':
            return await getEventTool(authManager, args);

          case 'outlook_update_event':
            return await updateEventTool(authManager, args);

          case 'outlook_delete_event':
            return await deleteEventTool(authManager, args);

          case 'outlook_respond_to_invite':
            return await respondToInviteTool(authManager, args);

          case 'outlook_validate_event_datetimes':
            return await validateEventDateTimesTool(authManager, args);

          case 'outlook_create_recurring_event':
            return await createRecurringEventTool(authManager, args);

          case 'outlook_find_meeting_times':
            return await findMeetingTimesTool(authManager, args);

          case 'outlook_check_availability':
            return await checkAvailabilityTool(authManager, args);

          case 'outlook_schedule_online_meeting':
            return await scheduleOnlineMeetingTool(authManager, args);

          case 'outlook_list_calendars':
            return await listCalendarsTool(authManager, args);

          case 'outlook_get_calendar_view':
            return await getCalendarViewTool(authManager, args);

          case 'outlook_get_busy_times':
            return await getBusyTimesTool(authManager, args);

          case 'outlook_build_recurrence_pattern':
            return await buildRecurrencePatternTool(authManager, args);

          case 'outlook_create_recurrence_helper':
            return await createRecurrenceHelperTool(authManager, args);

          case 'outlook_check_calendar_permissions':
            return await checkCalendarPermissionsTool(authManager, args);

          case 'outlook_get_email':
            console.error(`DEBUG: Calling getEmailTool`);
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

          case 'outlook_list_folders':
            return await listFoldersTool(authManager, args);

          case 'outlook_create_folder':
            return await createFolderTool(authManager, args);

          case 'outlook_rename_folder':
            return await renameFolderTool(authManager, args);

          case 'outlook_get_folder_stats':
            return await getFolderStatsTool(authManager, args);

          case 'outlook_list_attachments':
            return await listAttachmentsTool(authManager, args);

          case 'outlook_download_attachment':
            return await downloadAttachmentTool(authManager, args);

          case 'outlook_add_attachment':
            return await addAttachmentTool(authManager, args);

          case 'outlook_scan_attachments':
            return await scanAttachmentsTool(authManager, args);

          case 'outlook_get_sharepoint_file':
            return await getSharePointFileTool(authManager, args);

          case 'outlook_list_sharepoint_files':
            return await listSharePointFilesTool(authManager, args);

          case 'outlook_resolve_sharepoint_link':
            return await resolveSharePointLinkTool(authManager, args);

          default:
            return createProtocolError(
              ErrorCodes.METHOD_NOT_FOUND,
              `Unknown tool: ${name}`,
              { availableTools: Object.keys(tools).filter(key => key.endsWith('Tool')) }
            );
        }
      } catch (error) {
        console.error('Unexpected error in tool handler:', error);

        // If it's already an MCP error response, return it as-is
        if (error.content && error.isError !== undefined) {
          return error;
        }

        // Convert other errors to MCP tool error format
        return convertErrorToToolError(error, 'Tool execution failed');
      }
    });

    server.setRequestHandler(ListPromptsRequestSchema, async () => {
      console.error(`Debug: Returning ${promptList.length} prompts to client`);
      return {
        prompts: promptList,
      };
    });

    server.setRequestHandler(GetPromptRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;
      console.error(`Debug: Getting prompt '${name}' with args:`, JSON.stringify(args, null, 2));
      try {
        return await getPrompt(name, args);
      } catch (error) {
        console.error(`Error getting prompt ${name}:`, error);
        throw convertErrorToToolError(error, `Failed to get prompt ${name}`);
      }
    });

    // Start the server
    async function main() {
      console.error('Debug: Starting main function...');
      console.error(`Debug: AZURE_CLIENT_ID = ${process.env.AZURE_CLIENT_ID ? 'SET' : 'NOT SET'}`);
      console.error(`Debug: AZURE_TENANT_ID = ${process.env.AZURE_TENANT_ID ? 'SET' : 'NOT SET'}`);

      if (!process.env.AZURE_CLIENT_ID || !process.env.AZURE_TENANT_ID) {
        console.error('Error: AZURE_CLIENT_ID and AZURE_TENANT_ID environment variables are required.');
        console.error('Please set these in your MCP server configuration.');
        console.error('Note: This server uses OAuth 2.0 with PKCE for secure delegated authentication.');
        process.exit(1);
      }

      console.error('Starting Outlook MCP Server...');
      console.error('Authentication will be performed when first tool is called.');

      try {
        console.error('Debug: Creating StdioServerTransport...');
        const transport = new StdioServerTransport();

        console.error('Debug: Connecting server to transport...');
        await server.connect(transport);

        console.error('Outlook MCP server is ready and connected');

      } catch (error) {
        console.error('Error during server connection:', error);
        throw error;
      }
    }

    await main();

  } catch (error) {
    console.error('Server error:', error);
    process.exit(1);
  }
})();
