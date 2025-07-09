// Cache for user styling information to avoid repeated API calls
const stylingCache = new Map();
const signatureCache = new Map();
const CACHE_DURATION = 30 * 60 * 1000; // 30 minutes in milliseconds
const SIGNATURE_CACHE_DURATION = 60 * 60 * 1000; // 1 hour for signatures (they change less frequently)

// Function to clear styling cache for a specific user or all users
export function clearStylingCache(userId = null) {
  if (userId) {
    const cacheKey = `styling_${userId}`;
    stylingCache.delete(cacheKey);
    console.log(`Cleared styling cache for user ${userId}`);
  } else {
    stylingCache.clear();
    console.log('Cleared all styling cache');
  }
}

// Function to clear signature cache for a specific user or all users
export function clearSignatureCache(userId = null) {
  if (userId) {
    const cacheKey = `signature_${userId}`;
    signatureCache.delete(cacheKey);
    console.log(`Cleared signature cache for user ${userId}`);
  } else {
    signatureCache.clear();
    console.log('Cleared all signature cache');
  }
}

// Function to get cache statistics
export function getStylingCacheStats() {
  const stats = {
    totalEntries: stylingCache.size,
    entries: []
  };
  
  for (const [key, value] of stylingCache) {
    stats.entries.push({
      key,
      age: Date.now() - value.timestamp,
      isExpired: (Date.now() - value.timestamp) > CACHE_DURATION
    });
  }
  
  return stats;
}

// Function to clean up expired cache entries
function cleanupExpiredCache() {
  const now = Date.now();
  
  // Clean up styling cache
  for (const [key, value] of stylingCache) {
    if ((now - value.timestamp) > CACHE_DURATION) {
      stylingCache.delete(key);
      console.log(`Cleaned up expired styling cache entry: ${key}`);
    }
  }
  
  // Clean up signature cache
  for (const [key, value] of signatureCache) {
    if ((now - value.timestamp) > SIGNATURE_CACHE_DURATION) {
      signatureCache.delete(key);
      console.log(`Cleaned up expired signature cache entry: ${key}`);
    }
  }
}

// Run cleanup every 10 minutes
setInterval(cleanupExpiredCache, 10 * 60 * 1000);

export async function listEmailsTool(authManager, args) {
  const { folder = 'inbox', limit = 10, filter } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    
    const options = {
      select: 'subject,from,receivedDateTime,bodyPreview,isRead',
      top: limit,
      orderby: 'receivedDateTime desc',
    };

    if (filter) {
      options.filter = filter;
    }

    const result = await graphApiClient.makeRequest(`/me/mailFolders/${folder}/messages`, options);

    const emails = result.value.map(email => ({
      id: email.id,
      subject: email.subject,
      from: email.from?.emailAddress?.address || 'Unknown',
      fromName: email.from?.emailAddress?.name || 'Unknown',
      receivedDateTime: email.receivedDateTime,
      preview: email.bodyPreview,
      isRead: email.isRead,
    }));

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ emails, count: emails.length }, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to list emails: ${error.message}`);
  }
}

export async function sendEmailTool(authManager, args) {
  const { to, subject, body, bodyType = 'text', cc = [], bcc = [], preserveUserStyling = true } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    let finalBody = body;
    let finalBodyType = bodyType;

    // If preserving user styling, get user's default styling and signature
    if (preserveUserStyling) {
      const styledBody = await applyUserStyling(graphApiClient, body, bodyType);
      finalBody = styledBody.content;
      finalBodyType = styledBody.type;
    }

    const message = {
      subject,
      body: {
        contentType: finalBodyType === 'html' ? 'HTML' : 'Text',
        content: finalBody,
      },
      toRecipients: to.map(email => ({
        emailAddress: { address: email },
      })),
    };

    if (cc.length > 0) {
      message.ccRecipients = cc.map(email => ({
        emailAddress: { address: email },
      }));
    }

    if (bcc.length > 0) {
      message.bccRecipients = bcc.map(email => ({
        emailAddress: { address: email },
      }));
    }

    await graphApiClient.postWithRetry('/me/sendMail', {
      message,
      saveToSentItems: true,
    });

    // Invalidate styling cache after sending email (user might have changed styling)
    // Don't invalidate signature cache as frequently since signatures change less often
    try {
      const userInfo = await graphApiClient.makeRequest('/me', { select: 'id' });
      clearStylingCache(userInfo.id);
    } catch (error) {
      console.warn('Could not invalidate styling cache:', error.message);
    }

    return {
      content: [
        {
          type: 'text',
          text: `Email sent successfully to ${to.join(', ')}`,
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to send email: ${error.message}`);
  }
}

export async function listEventsTool(authManager, args) {
  const { startDateTime, endDateTime, limit = 10, calendar } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const endpoint = calendar ? `/me/calendars/${calendar}/events` : '/me/events';
    const options = {
      select: 'subject,start,end,location,attendees,bodyPreview',
      top: limit,
      orderby: 'start/dateTime',
    };

    if (startDateTime && endDateTime) {
      options.filter = `start/dateTime ge '${startDateTime}' and end/dateTime le '${endDateTime}'`;
    }

    const result = await graphApiClient.makeRequest(endpoint, options);

    const events = result.value.map(event => ({
      id: event.id,
      subject: event.subject,
      start: event.start,
      end: event.end,
      location: event.location?.displayName || 'No location',
      attendees: event.attendees?.map(a => a.emailAddress?.address) || [],
      preview: event.bodyPreview,
    }));

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ events, count: events.length }, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to list events: ${error.message}`);
  }
}

export async function createEventTool(authManager, args) {
  const { subject, start, end, body = '', location = '', attendees = [] } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const event = {
      subject,
      start,
      end,
      body: {
        contentType: 'Text',
        content: body,
      },
    };

    if (location) {
      event.location = {
        displayName: location,
      };
    }

    if (attendees.length > 0) {
      event.attendees = attendees.map(email => ({
        emailAddress: { address: email },
        type: 'required',
      }));
    }

    const result = await graphApiClient.postWithRetry('/me/events', event);

    return {
      content: [
        {
          type: 'text',
          text: `Event "${subject}" created successfully. Event ID: ${result.id}`,
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to create event: ${error.message}`);
  }
}

export async function getEmailTool(authManager, args) {
  const { messageId } = args;

  if (!messageId) {
    throw new Error('messageId is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    
    const options = {
      select: 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,sentDateTime,body,bodyPreview,importance,isRead,hasAttachments,attachments,conversationId'
    };

    const email = await graphApiClient.makeRequest(`/me/messages/${messageId}`, options);

    const emailData = {
      id: email.id,
      subject: email.subject,
      from: {
        address: email.from?.emailAddress?.address || 'Unknown',
        name: email.from?.emailAddress?.name || 'Unknown'
      },
      toRecipients: email.toRecipients?.map(r => ({
        address: r.emailAddress?.address,
        name: r.emailAddress?.name
      })) || [],
      ccRecipients: email.ccRecipients?.map(r => ({
        address: r.emailAddress?.address,
        name: r.emailAddress?.name
      })) || [],
      bccRecipients: email.bccRecipients?.map(r => ({
        address: r.emailAddress?.address,
        name: r.emailAddress?.name
      })) || [],
      receivedDateTime: email.receivedDateTime,
      sentDateTime: email.sentDateTime,
      body: {
        contentType: email.body?.contentType || 'Text',
        content: email.body?.content || ''
      },
      bodyPreview: email.bodyPreview,
      importance: email.importance,
      isRead: email.isRead,
      hasAttachments: email.hasAttachments,
      attachments: email.attachments?.map(a => ({
        id: a.id,
        name: a.name,
        contentType: a.contentType,
        size: a.size
      })) || [],
      conversationId: email.conversationId
    };

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(emailData, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to get email: ${error.message}`);
  }
}

export async function searchEmailsTool(authManager, args) {
  const { 
    query,
    subject,
    from,
    startDate,
    endDate,
    folders = [],
    limit = 100,
    includeBody = true,
    orderBy = 'receivedDateTime desc'
  } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const options = {
      top: Math.min(limit, 1000) // Cap at 1000 for performance
      // orderby will be added conditionally after determining if search is used
    };

    // Build select based on includeBody parameter
    if (includeBody) {
      options.select = 'id,subject,from,toRecipients,receivedDateTime,sentDateTime,body,bodyPreview,importance,isRead,hasAttachments,conversationId';
    } else {
      options.select = 'id,subject,from,toRecipients,receivedDateTime,sentDateTime,bodyPreview,importance,isRead,hasAttachments,conversationId';
    }

    // Determine search strategy and endpoint
    let endpoint = '/me/messages';
    let useKQLSearch = false;
    let useODataFilters = false;
    const isSpecificFolder = folders.length === 1;
    
    if (isSpecificFolder) {
      // Single folder search
      endpoint = `/me/mailFolders/${folders[0]}/messages`;
      useODataFilters = true;
    } else if (folders.length > 1) {
      // Multiple folders - we'll need to make separate requests and combine
      // For now, fall back to all folders search
      endpoint = '/me/messages';
      useODataFilters = true;
    } else {
      // All folders search (folders.length === 0)
      endpoint = '/me/messages';
      
      // Decide between KQL search and OData filters
      if (query) {
        // Use KQL search for text queries (more efficient for content search)
        useKQLSearch = true;
      } else {
        // Use OData filters for sender/subject/date searches (more comprehensive)
        useODataFilters = true;
      }
    }

    // Build search parameters based on chosen strategy
    if (useODataFilters) {
      // Use $filter for reliable, comprehensive searches
      const filterConditions = [];
      
      // For Microsoft Graph API compatibility with $orderby, we need receivedDateTime in $filter
      // when using receivedDateTime in $orderby. Add it first to match orderby priority.
      if (orderBy && orderBy.includes('receivedDateTime')) {
        if (startDate) {
          filterConditions.push(`receivedDateTime ge ${startDate}`);
        } else {
          // Add a broad receivedDateTime filter to satisfy API requirements
          filterConditions.push(`receivedDateTime ge 1900-01-01T00:00:00Z`);
        }
        
        if (endDate) {
          filterConditions.push(`receivedDateTime le ${endDate}`);
        }
      } else {
        // Add date filters normally if not using receivedDateTime orderby
        if (startDate) {
          filterConditions.push(`receivedDateTime ge ${startDate}`);
        }
        
        if (endDate) {
          filterConditions.push(`receivedDateTime le ${endDate}`);
        }
      }
      
      if (from) {
        filterConditions.push(`from/emailAddress/address eq '${from.replace(/'/g, "''")}'`);
      }
      
      if (subject) {
        filterConditions.push(`contains(subject,'${subject.replace(/'/g, "''")}')`);
      }
      
      if (query) {
        // Use contains for general text search
        filterConditions.push(`contains(subject,'${query.replace(/'/g, "''")}') or contains(body/content,'${query.replace(/'/g, "''")}')`);
      }
      
      if (filterConditions.length > 0) {
        options.filter = filterConditions.join(' and ');
      }
      
      // Add orderby for OData filter searches
      options.orderby = orderBy;
      
    } else if (useKQLSearch) {
      // Use KQL search for text-based queries (combines text search with other filters)
      const kqlTerms = [];
      const filterConditions = [];

      // General text search using KQL (more efficient for content search)
      if (query) {
        kqlTerms.push(`"${query.replace(/"/g, '\\"')}"`);
      }

      // For sender/subject/date filters, we'll use KQL when possible, OData filters as fallback
      if (from) {
        kqlTerms.push(`"from:${from.replace(/"/g, '\\"')}"`);
      }

      if (subject) {
        kqlTerms.push(`"subject:${subject.replace(/"/g, '\\"')}"`);
      }

      // Date range using KQL format
      if (startDate && endDate) {
        const startFormatted = new Date(startDate).toLocaleDateString('en-US');
        const endFormatted = new Date(endDate).toLocaleDateString('en-US');
        kqlTerms.push(`"received:${startFormatted}..${endFormatted}"`);
      } else if (startDate) {
        const startFormatted = new Date(startDate).toLocaleDateString('en-US');
        kqlTerms.push(`"received>=${startFormatted}"`);
      } else if (endDate) {
        const endFormatted = new Date(endDate).toLocaleDateString('en-US');
        kqlTerms.push(`"received<=${endFormatted}"`);
      }

      // Combine all KQL terms with AND
      if (kqlTerms.length > 0) {
        options.search = kqlTerms.join(' AND ');
      }

      // Only add orderby if not using search (since search has its own sorting)
      if (!options.search) {
        options.orderby = orderBy;
      }
    }

    // Make the request using chosen search strategy
    const result = await graphApiClient.makeRequest(endpoint, options);

    const emails = result.value.map(email => {
      const emailData = {
        id: email.id,
        subject: email.subject,
        from: {
          address: email.from?.emailAddress?.address || 'Unknown',
          name: email.from?.emailAddress?.name || 'Unknown'
        },
        toRecipients: email.toRecipients?.map(r => ({
          address: r.emailAddress?.address,
          name: r.emailAddress?.name
        })) || [],
        receivedDateTime: email.receivedDateTime,
        sentDateTime: email.sentDateTime,
        bodyPreview: email.bodyPreview,
        importance: email.importance,
        isRead: email.isRead,
        hasAttachments: email.hasAttachments,
        conversationId: email.conversationId
      };

      // Include full body if requested
      if (includeBody && email.body) {
        emailData.body = {
          contentType: email.body.contentType || 'Text',
          content: email.body.content || ''
        };
      }

      return emailData;
    });

    const searchSummary = {
      searchApproach: useKQLSearch ? 'KQL (Keyword Query Language)' : 'OData $filter',
      kqlQuery: options.search || null,
      filterQuery: options.filter || null,
      endpoint: endpoint,
      folders: folders.length > 0 ? folders : ['All folders'],
      parameters: {
        generalSearch: query || null,
        sender: from ? (useKQLSearch ? `KQL: from:${from}` : `Filter: from/emailAddress/address eq '${from}'`) : null,
        subject: subject ? (useKQLSearch ? `KQL: subject:${subject}` : `Filter: contains(subject,'${subject}')`) : null,
        dateRange: startDate && endDate ? (useKQLSearch ? `KQL: received:${new Date(startDate).toLocaleDateString('en-US')}..${new Date(endDate).toLocaleDateString('en-US')}` : `Filter: receivedDateTime ge ${startDate} and receivedDateTime le ${endDate}`) : 
                   startDate ? (useKQLSearch ? `KQL: received>=${new Date(startDate).toLocaleDateString('en-US')}` : `Filter: receivedDateTime ge ${startDate}`) :
                   endDate ? (useKQLSearch ? `KQL: received<=${new Date(endDate).toLocaleDateString('en-US')}` : `Filter: receivedDateTime le ${endDate}`) : null
      },
      totalResults: emails.length,
      includesFullBody: includeBody,
      optimization: useKQLSearch ? 'Using KQL for text-based search (efficient for content search)' : 
                   isSpecificFolder ? 'Using $filter for specific folder search (comprehensive)' : 
                   'Using $filter for all-folders search (comprehensive across all folders)'
    };

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({ 
            searchSummary,
            emails 
          }, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to search emails: ${error.message}`);
  }
}

export async function createDraftTool(authManager, args) {
  const { to, subject, body, bodyType = 'text', cc = [], bcc = [], importance = 'normal' } = args;

  if (!to || to.length === 0) {
    throw new Error('At least one recipient is required');
  }

  if (!subject) {
    throw new Error('Subject is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const draft = {
      subject,
      body: {
        contentType: bodyType === 'html' ? 'HTML' : 'Text',
        content: body || '',
      },
      toRecipients: to.map(email => ({
        emailAddress: { address: email },
      })),
      importance,
    };

    if (cc.length > 0) {
      draft.ccRecipients = cc.map(email => ({
        emailAddress: { address: email },
      }));
    }

    if (bcc.length > 0) {
      draft.bccRecipients = bcc.map(email => ({
        emailAddress: { address: email },
      }));
    }

    const result = await graphApiClient.postWithRetry('/me/messages', draft);

    return {
      content: [
        {
          type: 'text',
          text: `Draft created successfully. Draft ID: ${result.id}`,
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to create draft: ${error.message}`);
  }
}

export async function replyToEmailTool(authManager, args) {
  const { messageId, body, bodyType = 'text', comment = '' } = args;

  if (!messageId) {
    throw new Error('messageId is required');
  }

  if (!body && !comment) {
    throw new Error('Either body or comment is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const replyPayload = {};

    // If body is provided, create the reply message structure
    if (body) {
      replyPayload.message = {
        body: {
          contentType: bodyType === 'html' ? 'HTML' : 'Text',
          content: body,
        },
      };
    }

    // Add comment if provided (this is for the reply action)
    if (comment) {
      replyPayload.comment = comment;
    }

    const result = await graphApiClient.postWithRetry(`/me/messages/${messageId}/reply`, replyPayload);

    return {
      content: [
        {
          type: 'text',
          text: `Reply created successfully. Reply ID: ${result.id || 'N/A'}`,
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to reply to email: ${error.message}`);
  }
}

export async function replyAllTool(authManager, args) {
  const { messageId, body, bodyType = 'text', comment = '' } = args;

  if (!messageId) {
    throw new Error('messageId is required');
  }

  if (!body && !comment) {
    throw new Error('Either body or comment is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const replyPayload = {};

    // If body is provided, create the reply message structure
    if (body) {
      replyPayload.message = {
        body: {
          contentType: bodyType === 'html' ? 'HTML' : 'Text',
          content: body,
        },
      };
    }

    // Add comment if provided (this is for the reply action)
    if (comment) {
      replyPayload.comment = comment;
    }

    const result = await graphApiClient.postWithRetry(`/me/messages/${messageId}/replyAll`, replyPayload);

    return {
      content: [
        {
          type: 'text',
          text: `Reply all created successfully. Reply ID: ${result.id || 'N/A'}`,
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to reply all to email: ${error.message}`);
  }
}

export async function forwardEmailTool(authManager, args) {
  const { messageId, to, body = '', bodyType = 'text', comment = '' } = args;

  if (!messageId) {
    throw new Error('messageId is required');
  }

  if (!to || to.length === 0) {
    throw new Error('At least one recipient is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const forwardPayload = {
      toRecipients: to.map(email => ({
        emailAddress: { address: email },
      })),
    };

    // If body is provided, create the forward message structure
    if (body) {
      forwardPayload.message = {
        body: {
          contentType: bodyType === 'html' ? 'HTML' : 'Text',
          content: body,
        },
      };
    }

    // Add comment if provided (this is for the forward action)
    if (comment) {
      forwardPayload.comment = comment;
    }

    const result = await graphApiClient.postWithRetry(`/me/messages/${messageId}/forward`, forwardPayload);

    return {
      content: [
        {
          type: 'text',
          text: `Email forwarded successfully to ${to.join(', ')}. Forward ID: ${result.id || 'N/A'}`,
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to forward email: ${error.message}`);
  }
}

export async function deleteEmailTool(authManager, args) {
  const { messageId, permanentDelete = false } = args;

  if (!messageId) {
    throw new Error('messageId is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    if (permanentDelete) {
      // Permanently delete the email
      await graphApiClient.makeRequest(`/me/messages/${messageId}`, {}, 'DELETE');
      
      return {
        content: [
          {
            type: 'text',
            text: `Email permanently deleted successfully. Message ID: ${messageId}`,
          },
        ],
      };
    } else {
      // Move to Deleted Items folder (soft delete)
      // First get the Deleted Items folder ID
      const foldersResult = await graphApiClient.makeRequest('/me/mailFolders', {
        filter: "displayName eq 'Deleted Items'"
      });
      
      let deletedItemsFolderId = 'deleteditems'; // Default fallback
      if (foldersResult.value && foldersResult.value.length > 0) {
        deletedItemsFolderId = foldersResult.value[0].id;
      }

      // Move the message to Deleted Items
      await graphApiClient.postWithRetry(`/me/messages/${messageId}/move`, {
        destinationId: deletedItemsFolderId
      });

      return {
        content: [
          {
            type: 'text',
            text: `Email moved to Deleted Items successfully. Message ID: ${messageId}`,
          },
        ],
      };
    }
  } catch (error) {
    throw new Error(`Failed to delete email: ${error.message}`);
  }
}

// Enhanced Calendar Integration Tools

export async function getEventTool(authManager, args) {
  const { eventId, calendarId } = args;

  if (!eventId) {
    throw new Error('eventId is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    
    const endpoint = calendarId ? 
      `/me/calendars/${calendarId}/events/${eventId}` : 
      `/me/events/${eventId}`;
    
    const options = {
      select: 'id,subject,start,end,location,attendees,body,bodyPreview,organizer,isAllDay,showAs,sensitivity,importance,recurrence,reminderMinutesBeforeStart,responseRequested,allowNewTimeProposals,onlineMeeting,isOnlineMeeting,onlineMeetingProvider,categories,createdDateTime,lastModifiedDateTime'
    };

    const event = await graphApiClient.makeRequest(endpoint, options);

    const eventData = {
      id: event.id,
      subject: event.subject,
      start: event.start,
      end: event.end,
      location: event.location || {},
      attendees: event.attendees?.map(a => ({
        emailAddress: a.emailAddress,
        status: a.status,
        type: a.type
      })) || [],
      body: {
        contentType: event.body?.contentType || 'Text',
        content: event.body?.content || ''
      },
      bodyPreview: event.bodyPreview,
      organizer: event.organizer,
      isAllDay: event.isAllDay,
      showAs: event.showAs,
      sensitivity: event.sensitivity,
      importance: event.importance,
      recurrence: event.recurrence,
      reminderMinutesBeforeStart: event.reminderMinutesBeforeStart,
      responseRequested: event.responseRequested,
      allowNewTimeProposals: event.allowNewTimeProposals,
      onlineMeeting: event.onlineMeeting,
      isOnlineMeeting: event.isOnlineMeeting,
      onlineMeetingProvider: event.onlineMeetingProvider,
      categories: event.categories || [],
      createdDateTime: event.createdDateTime,
      lastModifiedDateTime: event.lastModifiedDateTime
    };

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(eventData, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to get event: ${error.message}`);
  }
}

export async function updateEventTool(authManager, args) {
  const { eventId, calendarId, subject, start, end, body, location, attendees, isAllDay, showAs, importance, reminderMinutesBeforeStart, categories } = args;

  if (!eventId) {
    throw new Error('eventId is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const endpoint = calendarId ? 
      `/me/calendars/${calendarId}/events/${eventId}` : 
      `/me/events/${eventId}`;

    const updateData = {};

    if (subject !== undefined) updateData.subject = subject;
    if (start !== undefined) updateData.start = start;
    if (end !== undefined) updateData.end = end;
    if (isAllDay !== undefined) updateData.isAllDay = isAllDay;
    if (showAs !== undefined) updateData.showAs = showAs;
    if (importance !== undefined) updateData.importance = importance;
    if (reminderMinutesBeforeStart !== undefined) updateData.reminderMinutesBeforeStart = reminderMinutesBeforeStart;
    if (categories !== undefined) updateData.categories = categories;

    if (body !== undefined) {
      updateData.body = {
        contentType: 'Text',
        content: body,
      };
    }

    if (location !== undefined) {
      updateData.location = typeof location === 'string' ? 
        { displayName: location } : location;
    }

    if (attendees !== undefined) {
      updateData.attendees = attendees.map(attendee => {
        if (typeof attendee === 'string') {
          return {
            emailAddress: { address: attendee },
            type: 'required',
          };
        }
        return attendee;
      });
    }

    const result = await graphApiClient.makeRequest(endpoint, updateData, 'PATCH');

    return {
      content: [
        {
          type: 'text',
          text: `Event "${result.subject}" updated successfully. Event ID: ${result.id}`,
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to update event: ${error.message}`);
  }
}

export async function deleteEventTool(authManager, args) {
  const { eventId, calendarId } = args;

  if (!eventId) {
    throw new Error('eventId is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const endpoint = calendarId ? 
      `/me/calendars/${calendarId}/events/${eventId}` : 
      `/me/events/${eventId}`;

    await graphApiClient.makeRequest(endpoint, {}, 'DELETE');

    return {
      content: [
        {
          type: 'text',
          text: `Event deleted successfully. Event ID: ${eventId}`,
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to delete event: ${error.message}`);
  }
}

export async function createRecurringEventTool(authManager, args) {
  const { 
    subject, 
    start, 
    end, 
    body = '', 
    location = '', 
    attendees = [],
    recurrencePattern,
    recurrenceRange,
    calendarId
  } = args;

  if (!subject || !start || !end || !recurrencePattern) {
    throw new Error('subject, start, end, and recurrencePattern are required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const endpoint = calendarId ? 
      `/me/calendars/${calendarId}/events` : 
      '/me/events';

    const event = {
      subject,
      start,
      end,
      body: {
        contentType: 'Text',
        content: body,
      },
      recurrence: {
        pattern: recurrencePattern,
        range: recurrenceRange || {
          type: 'noEnd',
          startDate: start.dateTime.split('T')[0]
        }
      }
    };

    if (location) {
      event.location = {
        displayName: location,
      };
    }

    if (attendees.length > 0) {
      event.attendees = attendees.map(email => ({
        emailAddress: { address: email },
        type: 'required',
      }));
    }

    const result = await graphApiClient.postWithRetry(endpoint, event);

    return {
      content: [
        {
          type: 'text',
          text: `Recurring event "${subject}" created successfully. Event ID: ${result.id}`,
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to create recurring event: ${error.message}`);
  }
}

// Meeting Scheduling Tools

export async function findMeetingTimesTool(authManager, args) {
  const { 
    attendees = [], 
    timeConstraint,
    meetingDuration = 30,
    maxCandidates = 20,
    isOrganizerOptional = false,
    returnSuggestionReasons = true,
    activityDomain = 'work'
  } = args;

  if (!timeConstraint || !timeConstraint.timeslots || timeConstraint.timeslots.length === 0) {
    throw new Error('timeConstraint with timeslots is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const requestBody = {
      attendees: attendees.map(email => ({
        emailAddress: {
          address: typeof email === 'string' ? email : email.address,
          name: typeof email === 'string' ? email : email.name
        }
      })),
      timeConstraint,
      meetingDuration,
      maxCandidates,
      isOrganizerOptional,
      returnSuggestionReasons,
      activityDomain
    };

    const result = await graphApiClient.postWithRetry('/me/calendar/findMeetingTimes', requestBody);

    const suggestions = result.meetingTimeSuggestions?.map(suggestion => ({
      confidence: suggestion.confidence,
      organizerAvailability: suggestion.organizerAvailability,
      suggestionReason: suggestion.suggestionReason,
      meetingTimeSlot: {
        start: suggestion.meetingTimeSlot.start,
        end: suggestion.meetingTimeSlot.end
      },
      attendeeAvailability: suggestion.attendeeAvailability?.map(availability => ({
        attendee: availability.attendee,
        availability: availability.availability
      }))
    })) || [];

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            emptySuggestionsReason: result.emptySuggestionsReason,
            suggestions,
            totalSuggestions: suggestions.length
          }, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to find meeting times: ${error.message}`);
  }
}

export async function checkAvailabilityTool(authManager, args) {
  const { attendees = [], startTime, endTime, availabilityViewInterval = 60 } = args;

  if (!attendees || attendees.length === 0) {
    throw new Error('At least one attendee is required');
  }

  if (!startTime || !endTime) {
    throw new Error('startTime and endTime are required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const requestBody = {
      schedules: attendees.map(email => typeof email === 'string' ? email : email.address),
      startTime: {
        dateTime: startTime,
        timeZone: 'UTC'
      },
      endTime: {
        dateTime: endTime,
        timeZone: 'UTC'
      },
      availabilityViewInterval
    };

    const result = await graphApiClient.postWithRetry('/me/calendar/getSchedule', requestBody);

    const schedules = result.value?.map((schedule, index) => ({
      scheduleId: attendees[index],
      availabilityView: schedule.availabilityView,
      busyTimes: schedule.busyTimes?.map(busyTime => ({
        start: busyTime.start,
        end: busyTime.end
      })) || [],
      workingHours: schedule.workingHours
    })) || [];

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            timeframe: {
              startTime,
              endTime,
              availabilityViewInterval
            },
            schedules
          }, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to check availability: ${error.message}`);
  }
}

export async function scheduleOnlineMeetingTool(authManager, args) {
  const { 
    subject, 
    start, 
    end, 
    body = '', 
    location = '', 
    attendees = [],
    onlineMeetingProvider = 'teamsForBusiness',
    allowNewTimeProposals = true,
    responseRequested = true
  } = args;

  if (!subject || !start || !end) {
    throw new Error('subject, start, and end are required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const event = {
      subject,
      start,
      end,
      body: {
        contentType: 'Text',
        content: body,
      },
      isOnlineMeeting: true,
      onlineMeetingProvider,
      allowNewTimeProposals,
      responseRequested
    };

    if (location) {
      event.location = {
        displayName: location,
      };
    }

    if (attendees.length > 0) {
      event.attendees = attendees.map(attendee => {
        if (typeof attendee === 'string') {
          return {
            emailAddress: { address: attendee },
            type: 'required',
          };
        }
        return {
          emailAddress: {
            address: attendee.address,
            name: attendee.name
          },
          type: attendee.type || 'required'
        };
      });
    }

    const result = await graphApiClient.postWithRetry('/me/events', event);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            eventId: result.id,
            subject: result.subject,
            start: result.start,
            end: result.end,
            onlineMeeting: result.onlineMeeting,
            joinUrl: result.onlineMeeting?.joinUrl,
            conferenceId: result.onlineMeeting?.conferenceId,
            attendees: result.attendees?.length || 0
          }, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to schedule online meeting: ${error.message}`);
  }
}

export async function respondToInviteTool(authManager, args) {
  const { eventId, response, comment = '', sendResponse = true } = args;

  if (!eventId) {
    throw new Error('eventId is required');
  }

  if (!response || !['accept', 'decline', 'tentativelyAccept'].includes(response)) {
    throw new Error('response must be one of: accept, decline, tentativelyAccept');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const requestBody = {
      comment,
      sendResponse
    };

    const endpoint = `/me/events/${eventId}/${response}`;
    
    await graphApiClient.postWithRetry(endpoint, requestBody);

    return {
      content: [
        {
          type: 'text',
          text: `Successfully ${response === 'tentativelyAccept' ? 'tentatively accepted' : response + 'ed'} the meeting invitation. Event ID: ${eventId}`,
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to respond to meeting invitation: ${error.message}`);
  }
}

// Calendar Management Tools

export async function listCalendarsTool(authManager, args) {
  const { includeShared = false, top = 50 } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const options = {
      select: 'id,name,color,isDefaultCalendar,canShare,canViewPrivateItems,canEdit,allowedOnlineMeetingProviders,defaultOnlineMeetingProvider,isTallyingResponses,isRemovable,owner',
      top
    };

    let endpoint = '/me/calendars';
    if (includeShared) {
      endpoint = '/me/calendarGroups/MyCalendars/calendars'; // This includes shared calendars
    }

    const result = await graphApiClient.makeRequest(endpoint, options);

    const calendars = result.value?.map(calendar => ({
      id: calendar.id,
      name: calendar.name,
      color: calendar.color,
      isDefaultCalendar: calendar.isDefaultCalendar,
      canShare: calendar.canShare,
      canViewPrivateItems: calendar.canViewPrivateItems,
      canEdit: calendar.canEdit,
      allowedOnlineMeetingProviders: calendar.allowedOnlineMeetingProviders || [],
      defaultOnlineMeetingProvider: calendar.defaultOnlineMeetingProvider,
      isTallyingResponses: calendar.isTallyingResponses,
      isRemovable: calendar.isRemovable,
      owner: calendar.owner
    })) || [];

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            calendars,
            totalCount: calendars.length,
            includesShared: includeShared
          }, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to list calendars: ${error.message}`);
  }
}

export async function getCalendarViewTool(authManager, args) {
  const { 
    startDateTime, 
    endDateTime, 
    calendarId,
    top = 100,
    orderBy = 'start/dateTime',
    includeRecurrences = true 
  } = args;

  if (!startDateTime || !endDateTime) {
    throw new Error('startDateTime and endDateTime are required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const startDate = new Date(startDateTime).toISOString();
    const endDate = new Date(endDateTime).toISOString();

    let endpoint = '/me/calendarView';
    if (calendarId) {
      endpoint = `/me/calendars/${calendarId}/calendarView`;
    }

    const options = {
      startDateTime: startDate,
      endDateTime: endDate,
      select: 'id,subject,start,end,location,attendees,organizer,bodyPreview,importance,showAs,sensitivity,isAllDay,isCancelled,isOrganizer,recurrence,seriesMasterId,type',
      top,
      orderby: orderBy
    };

    if (!includeRecurrences) {
      options.filter = "type eq 'singleInstance' or type eq 'seriesMaster'";
    }

    const result = await graphApiClient.makeRequest(endpoint, options);

    const events = result.value?.map(event => ({
      id: event.id,
      subject: event.subject,
      start: event.start,
      end: event.end,
      location: event.location?.displayName || 'No location',
      attendees: event.attendees?.map(a => ({
        emailAddress: a.emailAddress,
        status: a.status,
        type: a.type
      })) || [],
      organizer: event.organizer,
      bodyPreview: event.bodyPreview,
      importance: event.importance,
      showAs: event.showAs,
      sensitivity: event.sensitivity,
      isAllDay: event.isAllDay,
      isCancelled: event.isCancelled,
      isOrganizer: event.isOrganizer,
      recurrence: event.recurrence,
      seriesMasterId: event.seriesMasterId,
      type: event.type
    })) || [];

    const summary = {
      dateRange: {
        startDateTime: startDate,
        endDateTime: endDate
      },
      totalEvents: events.length,
      eventsByType: {
        singleInstance: events.filter(e => e.type === 'singleInstance').length,
        occurrence: events.filter(e => e.type === 'occurrence').length,
        seriesMaster: events.filter(e => e.type === 'seriesMaster').length
      },
      calendarId: calendarId || 'default'
    };

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            summary,
            events
          }, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to get calendar view: ${error.message}`);
  }
}

export async function getBusyTimesTool(authManager, args) {
  const { 
    attendees = [], 
    startTime, 
    endTime, 
    availabilityViewInterval = 30,
    includeWorkingHours = true 
  } = args;

  if (!attendees || attendees.length === 0) {
    throw new Error('At least one attendee is required');
  }

  if (!startTime || !endTime) {
    throw new Error('startTime and endTime are required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const requestBody = {
      schedules: attendees.map(email => typeof email === 'string' ? email : email.address),
      startTime: {
        dateTime: startTime,
        timeZone: 'UTC'
      },
      endTime: {
        dateTime: endTime,
        timeZone: 'UTC'
      },
      availabilityViewInterval
    };

    const result = await graphApiClient.postWithRetry('/me/calendar/getSchedule', requestBody);

    const busyTimesSummary = result.value?.map((schedule, index) => {
      const attendeeEmail = typeof attendees[index] === 'string' ? attendees[index] : attendees[index].address;
      
      // Parse availability view to identify free/busy periods
      const availabilityView = schedule.availabilityView || '';
      const freeBusyPeriods = [];
      
      for (let i = 0; i < availabilityView.length; i++) {
        const status = availabilityView[i];
        const periodStart = new Date(startTime);
        periodStart.setMinutes(periodStart.getMinutes() + (i * availabilityViewInterval));
        
        const periodEnd = new Date(periodStart);
        periodEnd.setMinutes(periodEnd.getMinutes() + availabilityViewInterval);
        
        if (status !== '0') { // Not free
          freeBusyPeriods.push({
            start: periodStart.toISOString(),
            end: periodEnd.toISOString(),
            status: status === '1' ? 'tentative' : 
                   status === '2' ? 'busy' : 
                   status === '3' ? 'oof' : 
                   status === '4' ? 'workingElsewhere' : 'unknown'
          });
        }
      }

      return {
        attendee: attendeeEmail,
        busyTimes: schedule.busyTimes || [],
        freeBusyPeriods,
        workingHours: includeWorkingHours ? schedule.workingHours : undefined,
        availabilityView: schedule.availabilityView
      };
    }) || [];

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            timeframe: {
              startTime,
              endTime,
              availabilityViewInterval
            },
            summary: {
              totalAttendees: busyTimesSummary.length,
              intervalMinutes: availabilityViewInterval
            },
            busyTimesSummary
          }, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to get busy times: ${error.message}`);
  }
}

// Recurring Event Support and Pattern Tools

export async function buildRecurrencePatternTool(authManager, args) {
  const { 
    type, 
    interval = 1, 
    daysOfWeek = [], 
    dayOfMonth, 
    weekOfMonth, 
    month,
    rangeType = 'noEnd',
    startDate,
    endDate,
    numberOfOccurrences
  } = args;

  if (!type) {
    throw new Error('Recurrence type is required (daily, weekly, absoluteMonthly, relativeMonthly, absoluteYearly, relativeYearly)');
  }

  if (!['daily', 'weekly', 'absoluteMonthly', 'relativeMonthly', 'absoluteYearly', 'relativeYearly'].includes(type)) {
    throw new Error('Invalid recurrence type');
  }

  try {
    const pattern = {
      type,
      interval
    };

    // Add type-specific properties
    switch (type) {
      case 'weekly':
        if (daysOfWeek.length === 0) {
          throw new Error('daysOfWeek is required for weekly recurrence');
        }
        pattern.daysOfWeek = daysOfWeek;
        break;
      
      case 'absoluteMonthly':
        if (!dayOfMonth) {
          throw new Error('dayOfMonth is required for absoluteMonthly recurrence');
        }
        pattern.dayOfMonth = dayOfMonth;
        break;
      
      case 'relativeMonthly':
        if (daysOfWeek.length === 0 || !weekOfMonth) {
          throw new Error('daysOfWeek and weekOfMonth are required for relativeMonthly recurrence');
        }
        pattern.daysOfWeek = daysOfWeek;
        pattern.index = weekOfMonth; // first, second, third, fourth, last
        break;
      
      case 'absoluteYearly':
        if (!dayOfMonth || !month) {
          throw new Error('dayOfMonth and month are required for absoluteYearly recurrence');
        }
        pattern.dayOfMonth = dayOfMonth;
        pattern.month = month;
        break;
      
      case 'relativeYearly':
        if (daysOfWeek.length === 0 || !weekOfMonth || !month) {
          throw new Error('daysOfWeek, weekOfMonth, and month are required for relativeYearly recurrence');
        }
        pattern.daysOfWeek = daysOfWeek;
        pattern.index = weekOfMonth;
        pattern.month = month;
        break;
    }

    // Build range
    const range = {
      type: rangeType,
      startDate: startDate || new Date().toISOString().split('T')[0]
    };

    switch (rangeType) {
      case 'endDate':
        if (!endDate) {
          throw new Error('endDate is required when rangeType is endDate');
        }
        range.endDate = endDate;
        break;
      
      case 'numbered':
        if (!numberOfOccurrences) {
          throw new Error('numberOfOccurrences is required when rangeType is numbered');
        }
        range.numberOfOccurrences = numberOfOccurrences;
        break;
    }

    const recurrence = { pattern, range };

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            recurrence,
            summary: {
              type,
              interval,
              rangeType,
              description: this.getRecurrenceDescription(recurrence)
            }
          }, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to build recurrence pattern: ${error.message}`);
  }
}

// Helper function to generate human-readable recurrence description
function getRecurrenceDescription(recurrence) {
  const { pattern, range } = recurrence;
  let description = '';

  // Build pattern description
  switch (pattern.type) {
    case 'daily':
      description = pattern.interval === 1 ? 'Daily' : `Every ${pattern.interval} days`;
      break;
    
    case 'weekly':
      const days = pattern.daysOfWeek.join(', ');
      description = pattern.interval === 1 
        ? `Weekly on ${days}` 
        : `Every ${pattern.interval} weeks on ${days}`;
      break;
    
    case 'absoluteMonthly':
      description = pattern.interval === 1 
        ? `Monthly on day ${pattern.dayOfMonth}` 
        : `Every ${pattern.interval} months on day ${pattern.dayOfMonth}`;
      break;
    
    case 'relativeMonthly':
      const dayName = pattern.daysOfWeek[0];
      description = pattern.interval === 1 
        ? `Monthly on the ${pattern.index} ${dayName}` 
        : `Every ${pattern.interval} months on the ${pattern.index} ${dayName}`;
      break;
    
    case 'absoluteYearly':
      description = `Yearly on ${pattern.month} ${pattern.dayOfMonth}`;
      break;
    
    case 'relativeYearly':
      description = `Yearly on the ${pattern.index} ${pattern.daysOfWeek[0]} of ${pattern.month}`;
      break;
  }

  // Add range description
  switch (range.type) {
    case 'endDate':
      description += `, until ${range.endDate}`;
      break;
    
    case 'numbered':
      description += `, for ${range.numberOfOccurrences} occurrences`;
      break;
    
    case 'noEnd':
      description += ', indefinitely';
      break;
  }

  return description;
}

export async function createRecurrenceHelperTool(authManager, args) {
  const { 
    frequency, // 'daily', 'weekly', 'monthly', 'yearly'
    interval = 1,
    specificDays = [], // For weekly: ['monday', 'wednesday', 'friday']
    monthlyType = 'date', // 'date' or 'day' (e.g., "15th" vs "2nd Tuesday")
    dayOfMonth = 1,
    weekOfMonth = 'first', // 'first', 'second', 'third', 'fourth', 'last'
    endType = 'never', // 'never', 'date', 'count'
    endDate,
    occurrenceCount,
    startDate
  } = args;

  if (!frequency || !['daily', 'weekly', 'monthly', 'yearly'].includes(frequency)) {
    throw new Error('frequency must be one of: daily, weekly, monthly, yearly');
  }

  try {
    let recurrenceType;
    let pattern = { interval };
    
    // Map friendly names to Microsoft Graph types
    switch (frequency) {
      case 'daily':
        recurrenceType = 'daily';
        break;
      
      case 'weekly':
        recurrenceType = 'weekly';
        if (specificDays.length === 0) {
          throw new Error('specificDays is required for weekly recurrence');
        }
        pattern.daysOfWeek = specificDays;
        break;
      
      case 'monthly':
        if (monthlyType === 'date') {
          recurrenceType = 'absoluteMonthly';
          pattern.dayOfMonth = dayOfMonth;
        } else {
          recurrenceType = 'relativeMonthly';
          if (specificDays.length === 0) {
            throw new Error('specificDays is required for relative monthly recurrence');
          }
          pattern.daysOfWeek = specificDays;
          pattern.index = weekOfMonth;
        }
        break;
      
      case 'yearly':
        if (monthlyType === 'date') {
          recurrenceType = 'absoluteYearly';
          pattern.dayOfMonth = dayOfMonth;
          pattern.month = new Date().getMonth() + 1; // Default to current month
        } else {
          recurrenceType = 'relativeYearly';
          if (specificDays.length === 0) {
            throw new Error('specificDays is required for relative yearly recurrence');
          }
          pattern.daysOfWeek = specificDays;
          pattern.index = weekOfMonth;
          pattern.month = new Date().getMonth() + 1; // Default to current month
        }
        break;
    }

    // Build range
    let rangeType;
    const range = {
      startDate: startDate || new Date().toISOString().split('T')[0]
    };

    switch (endType) {
      case 'never':
        rangeType = 'noEnd';
        break;
      
      case 'date':
        if (!endDate) {
          throw new Error('endDate is required when endType is date');
        }
        rangeType = 'endDate';
        range.endDate = endDate;
        break;
      
      case 'count':
        if (!occurrenceCount) {
          throw new Error('occurrenceCount is required when endType is count');
        }
        rangeType = 'numbered';
        range.numberOfOccurrences = occurrenceCount;
        break;
      
      default:
        throw new Error('endType must be one of: never, date, count');
    }

    range.type = rangeType;
    pattern.type = recurrenceType;

    const recurrence = { pattern, range };

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            recurrence,
            friendlyDescription: getRecurrenceDescription(recurrence),
            summary: {
              frequency,
              interval,
              endType,
              monthlyType: frequency === 'monthly' || frequency === 'yearly' ? monthlyType : undefined
            }
          }, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to create recurrence helper: ${error.message}`);
  }
}

// Enhanced Error Handling and Validation for Calendar Operations

export async function validateEventDateTimesTool(authManager, args) {
  const { startDateTime, endDateTime, timeZone, isAllDay = false } = args;

  if (!startDateTime || !endDateTime) {
    throw new Error('Both startDateTime and endDateTime are required');
  }

  try {
    const errors = [];
    const warnings = [];

    // Parse dates
    let startDate, endDate;
    
    try {
      startDate = new Date(startDateTime);
      endDate = new Date(endDateTime);
    } catch (error) {
      errors.push('Invalid date format. Use ISO 8601 format (YYYY-MM-DDTHH:mm:ss)');
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({ valid: false, errors, warnings }, null, 2),
          },
        ],
      };
    }

    // Check if dates are valid
    if (isNaN(startDate.getTime())) {
      errors.push('Invalid startDateTime');
    }
    if (isNaN(endDate.getTime())) {
      errors.push('Invalid endDateTime');
    }

    if (errors.length > 0) {
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({ valid: false, errors, warnings }, null, 2),
          },
        ],
      };
    }

    // Validate date logic
    if (startDate >= endDate) {
      errors.push('Start date must be before end date');
    }

    // Check if start date is in the past
    const now = new Date();
    if (startDate < now) {
      warnings.push('Start date is in the past');
    }

    // Check for reasonable duration
    const durationHours = (endDate - startDate) / (1000 * 60 * 60);
    if (durationHours > 24 && !isAllDay) {
      warnings.push('Event duration is longer than 24 hours');
    }
    if (durationHours < 0.25 && !isAllDay) {
      warnings.push('Event duration is less than 15 minutes');
    }

    // Validate timezone if provided
    if (timeZone) {
      // Import the graphHelpers to use timezone validation
      try {
        const { graphHelpers } = await import('../graph/graphHelpers.js');
        const normalizedTz = graphHelpers.timezone.normalizeTimezone(timeZone);
        if (normalizedTz === 'UTC' && timeZone !== 'UTC') {
          warnings.push(`Timezone '${timeZone}' was normalized to UTC. Consider using a more specific timezone.`);
        }
      } catch (error) {
        warnings.push('Could not validate timezone');
      }
    }

    const isValid = errors.length === 0;

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            valid: isValid,
            errors,
            warnings,
            duration: {
              hours: durationHours,
              minutes: Math.round(durationHours * 60)
            },
            summary: isValid ? 'Event dates are valid' : 'Event dates have validation errors'
          }, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to validate event dates: ${error.message}`);
  }
}

export async function checkCalendarPermissionsTool(authManager, args) {
  const { calendarId } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const permissions = {
      canRead: false,
      canWrite: false,
      canShare: false,
      canViewPrivateItems: false,
      isOwner: false,
      errors: []
    };

    try {
      // Try to read calendar information
      const endpoint = calendarId ? `/me/calendars/${calendarId}` : '/me/calendar';
      const calendar = await graphApiClient.makeRequest(endpoint, {
        select: 'id,name,canEdit,canShare,canViewPrivateItems,owner,isDefaultCalendar'
      });

      permissions.canRead = true;
      permissions.canWrite = calendar.canEdit || false;
      permissions.canShare = calendar.canShare || false;
      permissions.canViewPrivateItems = calendar.canViewPrivateItems || false;
      permissions.isOwner = calendar.owner?.address === (await graphApiClient.makeRequest('/me', { select: 'mail' })).mail;
      permissions.isDefaultCalendar = calendar.isDefaultCalendar || false;
      permissions.calendarName = calendar.name;

    } catch (error) {
      if (error.status === 403) {
        permissions.errors.push('Insufficient permissions to access this calendar');
      } else if (error.status === 404) {
        permissions.errors.push('Calendar not found');
      } else {
        permissions.errors.push(`Error accessing calendar: ${error.message}`);
      }
    }

    // Test write permissions by attempting to create a dummy event (then delete it)
    if (permissions.canRead && permissions.canWrite) {
      try {
        const testEvent = {
          subject: 'Permission Test - Will Be Deleted',
          start: {
            dateTime: new Date(Date.now() + 60000).toISOString(), // 1 minute from now
            timeZone: 'UTC'
          },
          end: {
            dateTime: new Date(Date.now() + 120000).toISOString(), // 2 minutes from now
            timeZone: 'UTC'
          },
          body: {
            contentType: 'Text',
            content: 'This is a test event to verify write permissions'
          }
        };

        const createEndpoint = calendarId ? `/me/calendars/${calendarId}/events` : '/me/events';
        const createdEvent = await graphApiClient.postWithRetry(createEndpoint, testEvent);
        
        // Immediately delete the test event
        const deleteEndpoint = calendarId ? 
          `/me/calendars/${calendarId}/events/${createdEvent.id}` : 
          `/me/events/${createdEvent.id}`;
        await graphApiClient.makeRequest(deleteEndpoint, {}, 'DELETE');

        permissions.writeTestSuccessful = true;
      } catch (error) {
        permissions.canWrite = false;
        permissions.writeTestSuccessful = false;
        permissions.errors.push(`Write permission test failed: ${error.message}`);
      }
    }

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            calendarId: calendarId || 'default',
            permissions,
            summary: `Calendar access: ${permissions.canRead ? 'readable' : 'not readable'}, ${permissions.canWrite ? 'writable' : 'not writable'}`
          }, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to check calendar permissions: ${error.message}`);
  }
}

// Enhanced error handler for calendar operations
export function handleCalendarError(error, operation) {
  const errorDetails = {
    operation,
    statusCode: error.status || error.statusCode,
    code: error.code,
    message: error.message,
    suggestion: null,
    retryable: false
  };

  // Handle specific Microsoft Graph calendar errors
  switch (error.status) {
    case 400:
      errorDetails.suggestion = 'Check request parameters and datetime formats';
      break;
    
    case 401:
      errorDetails.suggestion = 'Token may be expired. Try re-authenticating';
      errorDetails.retryable = true;
      break;
    
    case 403:
      if (error.message?.includes('calendar')) {
        errorDetails.suggestion = 'Insufficient calendar permissions. Check that Calendars.ReadWrite is granted';
      } else {
        errorDetails.suggestion = 'Access denied. Check permissions';
      }
      break;
    
    case 404:
      if (operation.includes('event')) {
        errorDetails.suggestion = 'Event not found. It may have been deleted or moved';
      } else if (operation.includes('calendar')) {
        errorDetails.suggestion = 'Calendar not found. Check calendar ID';
      }
      break;
    
    case 409:
      errorDetails.suggestion = 'Conflict detected. This may be due to conflicting updates or time conflicts';
      break;
    
    case 429:
      errorDetails.suggestion = 'Rate limit exceeded. Wait before retrying';
      errorDetails.retryable = true;
      errorDetails.retryAfter = error.retryAfter || 60;
      break;
    
    case 503:
      errorDetails.suggestion = 'Service temporarily unavailable. Retry in a few minutes';
      errorDetails.retryable = true;
      break;

    default:
      if (error.message?.includes('timezone')) {
        errorDetails.suggestion = 'Invalid timezone. Use Microsoft Graph timezone format';
      } else if (error.message?.includes('recurrence')) {
        errorDetails.suggestion = 'Invalid recurrence pattern. Check pattern and range properties';
      } else if (error.message?.includes('attendee')) {
        errorDetails.suggestion = 'Invalid attendee format. Use valid email addresses';
      } else {
        errorDetails.suggestion = 'Unknown error. Check request format and try again';
      }
  }

  return errorDetails;
}

// Rate Limit Monitoring Tools

export async function getRateLimitMetricsTool(authManager, args) {
  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    
    // Get current metrics
    const metrics = graphApiClient.getRateLimitMetrics();
    
    // Run health check
    const healthAlerts = graphApiClient.checkRateLimitHealth();
    
    const summary = {
      status: healthAlerts.length === 0 ? 'healthy' : 'warnings',
      metrics,
      healthAlerts,
      recommendations: []
    };

    // Add recommendations based on metrics
    if (metrics.rateLimitHits > 0) {
      summary.recommendations.push('Monitor request patterns to avoid rate limiting');
    }
    
    if (metrics.averageRequestDuration > 2000) {
      summary.recommendations.push('Consider optimizing requests with $select parameters to reduce response size');
    }
    
    if (metrics.activeRequests > 2) {
      summary.recommendations.push('High concurrent request usage detected - monitor for performance impact');
    }

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(summary, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to get rate limit metrics: ${error.message}`);
  }
}

export async function resetRateLimitMetricsTool(authManager, args) {
  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    
    // Get metrics before reset for reporting
    const metricsBeforeReset = graphApiClient.getRateLimitMetrics();
    
    // Reset metrics
    graphApiClient.resetMetrics();
    
    const resetSummary = {
      action: 'metrics_reset',
      timestamp: new Date().toISOString(),
      previousMetrics: {
        rateLimitHits: metricsBeforeReset.rateLimitHits,
        totalRetries: metricsBeforeReset.totalRetries,
        backoffTime: metricsBeforeReset.backoffTime,
        averageRequestDuration: metricsBeforeReset.averageRequestDuration
      },
      newMetrics: graphApiClient.getRateLimitMetrics()
    };

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(resetSummary, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to reset rate limit metrics: ${error.message}`);
  }
}

// Email Management Tools

export async function moveEmailTool(authManager, args) {
  const { messageId, destinationFolderId } = args;

  if (!messageId) {
    throw new Error('messageId is required');
  }

  if (!destinationFolderId) {
    throw new Error('destinationFolderId is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const result = await graphApiClient.postWithRetry(`/me/messages/${messageId}/move`, {
      destinationId: destinationFolderId
    });

    return {
      content: [
        {
          type: 'text',
          text: `Email moved successfully. New Message ID: ${result.id}`,
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to move email: ${error.message}`);
  }
}

export async function markAsReadTool(authManager, args) {
  const { messageId, isRead = true } = args;

  if (!messageId) {
    throw new Error('messageId is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    await graphApiClient.makeRequest(`/me/messages/${messageId}`, {
      isRead: isRead
    }, 'PATCH');

    return {
      content: [
        {
          type: 'text',
          text: `Email ${isRead ? 'marked as read' : 'marked as unread'} successfully. Message ID: ${messageId}`,
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to mark email as ${isRead ? 'read' : 'unread'}: ${error.message}`);
  }
}

export async function flagEmailTool(authManager, args) {
  const { messageId, flagStatus = 'flagged' } = args;

  if (!messageId) {
    throw new Error('messageId is required');
  }

  if (!['notFlagged', 'complete', 'flagged'].includes(flagStatus)) {
    throw new Error('flagStatus must be one of: notFlagged, complete, flagged');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    await graphApiClient.makeRequest(`/me/messages/${messageId}`, {
      flag: {
        flagStatus: flagStatus
      }
    }, 'PATCH');

    return {
      content: [
        {
          type: 'text',
          text: `Email flag status set to '${flagStatus}' successfully. Message ID: ${messageId}`,
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to flag email: ${error.message}`);
  }
}

export async function categorizeEmailTool(authManager, args) {
  const { messageId, categories = [] } = args;

  if (!messageId) {
    throw new Error('messageId is required');
  }

  if (!Array.isArray(categories)) {
    throw new Error('categories must be an array');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    await graphApiClient.makeRequest(`/me/messages/${messageId}`, {
      categories: categories
    }, 'PATCH');

    return {
      content: [
        {
          type: 'text',
          text: `Email categories updated successfully. Message ID: ${messageId}, Categories: ${categories.join(', ') || 'None'}`,
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to categorize email: ${error.message}`);
  }
}

export async function archiveEmailTool(authManager, args) {
  const { messageId } = args;

  if (!messageId) {
    throw new Error('messageId is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    // First try to find the Archive folder
    const foldersResult = await graphApiClient.makeRequest('/me/mailFolders', {
      filter: "displayName eq 'Archive'"
    });
    
    let archiveFolderId = 'archive'; // Default fallback
    if (foldersResult.value && foldersResult.value.length > 0) {
      archiveFolderId = foldersResult.value[0].id;
    }

    // Move the message to Archive
    const result = await graphApiClient.postWithRetry(`/me/messages/${messageId}/move`, {
      destinationId: archiveFolderId
    });

    return {
      content: [
        {
          type: 'text',
          text: `Email archived successfully. New Message ID: ${result.id}`,
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to archive email: ${error.message}`);
  }
}

export async function batchProcessEmailsTool(authManager, args) {
  const { messageIds, operation, operationData = {} } = args;

  if (!messageIds || !Array.isArray(messageIds) || messageIds.length === 0) {
    throw new Error('messageIds array is required and must not be empty');
  }

  if (!operation) {
    throw new Error('operation is required');
  }

  const validOperations = ['markAsRead', 'markAsUnread', 'delete', 'move', 'flag', 'categorize'];
  if (!validOperations.includes(operation)) {
    throw new Error(`operation must be one of: ${validOperations.join(', ')}`);
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const results = [];
    const errors = [];

    // Process each message (could be optimized with batch requests in the future)
    for (const messageId of messageIds) {
      try {
        let result;
        
        switch (operation) {
          case 'markAsRead':
            await graphApiClient.makeRequest(`/me/messages/${messageId}`, { isRead: true }, 'PATCH');
            result = { messageId, status: 'success', operation: 'marked as read' };
            break;
            
          case 'markAsUnread':
            await graphApiClient.makeRequest(`/me/messages/${messageId}`, { isRead: false }, 'PATCH');
            result = { messageId, status: 'success', operation: 'marked as unread' };
            break;
            
          case 'delete':
            if (operationData.permanentDelete) {
              await graphApiClient.makeRequest(`/me/messages/${messageId}`, {}, 'DELETE');
              result = { messageId, status: 'success', operation: 'permanently deleted' };
            } else {
              // Find Deleted Items folder
              const foldersResult = await graphApiClient.makeRequest('/me/mailFolders', {
                filter: "displayName eq 'Deleted Items'"
              });
              let deletedItemsFolderId = 'deleteditems';
              if (foldersResult.value && foldersResult.value.length > 0) {
                deletedItemsFolderId = foldersResult.value[0].id;
              }
              await graphApiClient.postWithRetry(`/me/messages/${messageId}/move`, {
                destinationId: deletedItemsFolderId
              });
              result = { messageId, status: 'success', operation: 'moved to deleted items' };
            }
            break;
            
          case 'move':
            if (!operationData.destinationFolderId) {
              throw new Error('destinationFolderId is required for move operation');
            }
            await graphApiClient.postWithRetry(`/me/messages/${messageId}/move`, {
              destinationId: operationData.destinationFolderId
            });
            result = { messageId, status: 'success', operation: `moved to folder ${operationData.destinationFolderId}` };
            break;
            
          case 'flag':
            const flagStatus = operationData.flagStatus || 'flagged';
            await graphApiClient.makeRequest(`/me/messages/${messageId}`, {
              flag: { flagStatus }
            }, 'PATCH');
            result = { messageId, status: 'success', operation: `flagged as ${flagStatus}` };
            break;
            
          case 'categorize':
            const categories = operationData.categories || [];
            await graphApiClient.makeRequest(`/me/messages/${messageId}`, {
              categories
            }, 'PATCH');
            result = { messageId, status: 'success', operation: `categorized as ${categories.join(', ')}` };
            break;
        }
        
        results.push(result);
      } catch (error) {
        errors.push({ messageId, error: error.message });
      }
    }

    const summary = {
      totalProcessed: messageIds.length,
      successful: results.length,
      failed: errors.length,
      operation,
      results,
      errors
    };

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(summary, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to batch process emails: ${error.message}`);
  }
}

// Folder Management Tools

export async function listFoldersTool(authManager, args) {
  const { includeHidden = false, includeChildFolders = true, top = 100 } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const options = {
      select: 'id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount,isHidden',
      top: Math.min(top, 1000)
    };

    if (!includeHidden) {
      options.filter = 'isHidden eq false';
    }

    let endpoint = '/me/mailFolders';
    if (includeChildFolders) {
      endpoint = '/me/mailFolders?includeNestedFolders=true';
    }

    const result = await graphApiClient.makeRequest(endpoint, options);

    const folders = result.value?.map(folder => ({
      id: folder.id,
      name: folder.displayName,
      parentFolderId: folder.parentFolderId,
      childFolderCount: folder.childFolderCount || 0,
      unreadItemCount: folder.unreadItemCount || 0,
      totalItemCount: folder.totalItemCount || 0,
      isHidden: folder.isHidden || false
    })) || [];

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            folders,
            totalCount: folders.length,
            includesHidden: includeHidden,
            includesChildFolders: includeChildFolders
          }, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to list folders: ${error.message}`);
  }
}

export async function createFolderTool(authManager, args) {
  const { displayName, parentFolderId } = args;

  if (!displayName) {
    throw new Error('displayName is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const folderData = {
      displayName: displayName
    };

    let endpoint = '/me/mailFolders';
    if (parentFolderId) {
      endpoint = `/me/mailFolders/${parentFolderId}/childFolders`;
    }

    const result = await graphApiClient.postWithRetry(endpoint, folderData);

    return {
      content: [
        {
          type: 'text',
          text: `Folder "${displayName}" created successfully. Folder ID: ${result.id}`,
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to create folder: ${error.message}`);
  }
}

export async function renameFolderTool(authManager, args) {
  const { folderId, newDisplayName } = args;

  if (!folderId) {
    throw new Error('folderId is required');
  }

  if (!newDisplayName) {
    throw new Error('newDisplayName is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    await graphApiClient.makeRequest(`/me/mailFolders/${folderId}`, {
      displayName: newDisplayName
    }, 'PATCH');

    return {
      content: [
        {
          type: 'text',
          text: `Folder renamed to "${newDisplayName}" successfully. Folder ID: ${folderId}`,
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to rename folder: ${error.message}`);
  }
}

export async function getFolderStatsTool(authManager, args) {
  const { folderId, includeSubfolders = true } = args;

  if (!folderId) {
    throw new Error('folderId is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    // Get folder details
    const folder = await graphApiClient.makeRequest(`/me/mailFolders/${folderId}`, {
      select: 'id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount,isHidden'
    });

    const stats = {
      id: folder.id,
      name: folder.displayName,
      totalItems: folder.totalItemCount || 0,
      unreadItems: folder.unreadItemCount || 0,
      readItems: (folder.totalItemCount || 0) - (folder.unreadItemCount || 0),
      childFolders: folder.childFolderCount || 0,
      isHidden: folder.isHidden || false,
      parentFolderId: folder.parentFolderId
    };

    // Get subfolder stats if requested
    if (includeSubfolders && stats.childFolders > 0) {
      try {
        const childFolders = await graphApiClient.makeRequest(`/me/mailFolders/${folderId}/childFolders`, {
          select: 'id,displayName,childFolderCount,unreadItemCount,totalItemCount'
        });

        stats.subfolders = childFolders.value?.map(subfolder => ({
          id: subfolder.id,
          name: subfolder.displayName,
          totalItems: subfolder.totalItemCount || 0,
          unreadItems: subfolder.unreadItemCount || 0,
          childFolders: subfolder.childFolderCount || 0
        })) || [];

        // Calculate totals including subfolders
        stats.totalItemsIncludingSubfolders = stats.totalItems + 
          stats.subfolders.reduce((sum, sf) => sum + sf.totalItems, 0);
        stats.unreadItemsIncludingSubfolders = stats.unreadItems + 
          stats.subfolders.reduce((sum, sf) => sum + sf.unreadItems, 0);
      } catch (error) {
        stats.subfolderError = `Could not fetch subfolder stats: ${error.message}`;
      }
    }

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(stats, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to get folder stats: ${error.message}`);
  }
}

// Attachment Tools

export async function listAttachmentsTool(authManager, args) {
  const { messageId } = args;

  if (!messageId) {
    throw new Error('messageId is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const result = await graphApiClient.makeRequest(`/me/messages/${messageId}/attachments`, {
      select: 'id,name,contentType,size,isInline,lastModifiedDateTime'
    });

    const attachments = result.value?.map(attachment => ({
      id: attachment.id,
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size,
      sizeFormatted: formatFileSize(attachment.size),
      isInline: attachment.isInline || false,
      lastModifiedDateTime: attachment.lastModifiedDateTime
    })) || [];

    const summary = {
      messageId,
      totalAttachments: attachments.length,
      totalSize: attachments.reduce((sum, att) => sum + (att.size || 0), 0),
      attachments
    };

    summary.totalSizeFormatted = formatFileSize(summary.totalSize);

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(summary, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to list attachments: ${error.message}`);
  }
}

export async function downloadAttachmentTool(authManager, args) {
  const { messageId, attachmentId, includeContent = false } = args;

  if (!messageId) {
    throw new Error('messageId is required');
  }

  if (!attachmentId) {
    throw new Error('attachmentId is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const selectFields = includeContent 
      ? 'id,name,contentType,size,contentBytes,isInline,lastModifiedDateTime'
      : 'id,name,contentType,size,isInline,lastModifiedDateTime';

    const attachment = await graphApiClient.makeRequest(`/me/messages/${messageId}/attachments/${attachmentId}`, {
      select: selectFields
    });

    const attachmentInfo = {
      id: attachment.id,
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size,
      sizeFormatted: formatFileSize(attachment.size),
      isInline: attachment.isInline || false,
      lastModifiedDateTime: attachment.lastModifiedDateTime
    };

    if (includeContent && attachment.contentBytes) {
      attachmentInfo.contentBytes = attachment.contentBytes;
      attachmentInfo.note = 'Content is base64 encoded. Decode before saving to file.';
    } else {
      attachmentInfo.note = 'Content not included. Set includeContent=true to retrieve file content.';
    }

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(attachmentInfo, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to download attachment: ${error.message}`);
  }
}

export async function addAttachmentTool(authManager, args) {
  const { messageId, name, contentType, contentBytes } = args;

  if (!messageId) {
    throw new Error('messageId is required');
  }

  if (!name) {
    throw new Error('name is required');
  }

  if (!contentType) {
    throw new Error('contentType is required');
  }

  if (!contentBytes) {
    throw new Error('contentBytes is required (base64 encoded)');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    const attachmentData = {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: name,
      contentType: contentType,
      contentBytes: contentBytes
    };

    const result = await graphApiClient.postWithRetry(`/me/messages/${messageId}/attachments`, attachmentData);

    return {
      content: [
        {
          type: 'text',
          text: `Attachment "${name}" added successfully. Attachment ID: ${result.id}`,
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to add attachment: ${error.message}`);
  }
}

export async function scanAttachmentsTool(authManager, args) {
  const { 
    folder = 'inbox', 
    maxSizeMB = 10, 
    suspiciousTypes = ['exe', 'bat', 'cmd', 'scr', 'vbs', 'js'],
    limit = 100,
    daysBack = 30
  } = args;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    // Calculate date filter
    const sinceDate = new Date();
    sinceDate.setDate(sinceDate.getDate() - daysBack);

    const options = {
      select: 'id,subject,from,receivedDateTime,hasAttachments',
      filter: `hasAttachments eq true and receivedDateTime ge ${sinceDate.toISOString()}`,
      top: Math.min(limit, 1000),
      orderby: 'receivedDateTime desc'
    };

    const emailsResult = await graphApiClient.makeRequest(`/me/mailFolders/${folder}/messages`, options);

    const suspiciousEmails = [];
    const largeAttachments = [];
    const scanSummary = {
      totalEmailsScanned: emailsResult.value?.length || 0,
      suspiciousAttachments: 0,
      largeAttachments: 0,
      maxSizeMB: maxSizeMB,
      suspiciousFileTypes: suspiciousTypes
    };

    // Scan each email's attachments
    for (const email of emailsResult.value || []) {
      try {
        const attachmentsResult = await graphApiClient.makeRequest(`/me/messages/${email.id}/attachments`, {
          select: 'id,name,contentType,size'
        });

        for (const attachment of attachmentsResult.value || []) {
          const sizeMB = (attachment.size || 0) / (1024 * 1024);
          const fileExtension = attachment.name?.split('.').pop()?.toLowerCase() || '';

          // Check for large attachments
          if (sizeMB > maxSizeMB) {
            largeAttachments.push({
              emailId: email.id,
              emailSubject: email.subject,
              emailFrom: email.from?.emailAddress?.address,
              attachmentId: attachment.id,
              attachmentName: attachment.name,
              sizeMB: Math.round(sizeMB * 100) / 100,
              contentType: attachment.contentType
            });
            scanSummary.largeAttachments++;
          }

          // Check for suspicious file types
          if (suspiciousTypes.includes(fileExtension)) {
            suspiciousEmails.push({
              emailId: email.id,
              emailSubject: email.subject,
              emailFrom: email.from?.emailAddress?.address,
              attachmentId: attachment.id,
              attachmentName: attachment.name,
              fileExtension: fileExtension,
              sizeMB: Math.round(sizeMB * 100) / 100,
              risk: 'Potentially executable file type'
            });
            scanSummary.suspiciousAttachments++;
          }
        }
      } catch (error) {
        // Skip emails where we can't access attachments
        continue;
      }
    }

    const scanResults = {
      scanSummary,
      largeAttachments,
      suspiciousEmails,
      scannedAt: new Date().toISOString(),
      folder: folder
    };

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(scanResults, null, 2),
        },
      ],
    };
  } catch (error) {
    throw new Error(`Failed to scan attachments: ${error.message}`);
  }
}

// Helper function for file size formatting
function formatFileSize(bytes) {
  if (!bytes) return '0 Bytes';
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  if (bytes === 0) return '0 Bytes';
  const i = Math.floor(Math.log(bytes) / Math.log(1024));
  return Math.round(bytes / Math.pow(1024, i) * 100) / 100 + ' ' + sizes[i];
}

// Helper function to apply user's default styling and signature to email content
async function applyUserStyling(graphApiClient, content, bodyType) {
  try {
    // Try to get user's mail settings (requires MailboxSettings.Read permission)
    let mailSettings = null;
    try {
      mailSettings = await graphApiClient.makeRequest('/me/mailboxSettings');
      console.log('Retrieved mailbox settings:', JSON.stringify(mailSettings, null, 2));
    } catch (error) {
      if (error.message.includes('403') || error.message.includes('ErrorAccessDenied')) {
        console.warn('MailboxSettings.Read permission not granted, using basic styling');
        // Use basic styling without accessing mailbox settings
        return await applyBasicStyling(content, bodyType, graphApiClient);
      }
      throw error;
    }

    // Get user's default signature (if available)
    let signature = '';
    try {
      signature = await getUserSignature(graphApiClient);
    } catch (error) {
      console.warn('Could not retrieve signature:', error.message);
    }

    // Try to get actual styling from recent emails
    let actualStyling = null;
    try {
      actualStyling = await extractUserStylingFromEmails(graphApiClient);
      console.log('Extracted styling from emails:', actualStyling);
    } catch (error) {
      console.warn('Could not extract styling from emails:', error.message);
    }

    // Convert content to HTML if it's not already
    let htmlContent = content;
    if (bodyType === 'text') {
      // Convert plain text to HTML with basic formatting
      htmlContent = convertTextToHtml(content);
    }

    // Apply user's styling preferences
    const styledContent = await applyOutlookStyling(htmlContent, signature, mailSettings, actualStyling);

    return {
      content: styledContent,
      type: 'html'
    };
  } catch (error) {
    console.warn('Could not apply user styling, using original content:', error.message);
    return {
      content: content,
      type: bodyType
    };
  }
}

// Apply basic styling when mailbox settings are not accessible
async function applyBasicStyling(content, bodyType, graphApiClient) {
  try {
    // Get signature from sent emails (this works with existing permissions)
    const signature = await getUserSignature(graphApiClient);

    // Convert content to HTML if it's not already
    let htmlContent = content;
    if (bodyType === 'text') {
      htmlContent = convertTextToHtml(content);
    }

    // Apply basic Outlook-like styling without accessing mailbox settings
    const styledContent = await applyOutlookStyling(htmlContent, signature, null, null);

    return {
      content: styledContent,
      type: 'html'
    };
  } catch (error) {
    console.warn('Could not apply basic styling:', error.message);
    return {
      content: content,
      type: bodyType
    };
  }
}

// Extract actual styling from user's recent emails
async function extractUserStylingFromEmails(graphApiClient) {
  try {
    // Get user identifier for cache key
    const userInfo = await graphApiClient.makeRequest('/me', { select: 'id' });
    const cacheKey = `styling_${userInfo.id}`;
    
    // Check cache first
    const cachedData = stylingCache.get(cacheKey);
    if (cachedData && (Date.now() - cachedData.timestamp) < CACHE_DURATION) {
      console.log('Using cached styling information');
      return cachedData.styling;
    }
    
    console.log('Fetching fresh styling information from sent emails');
    
    // Get recent sent emails to analyze styling patterns
    const sentItems = await graphApiClient.makeRequest('/me/mailFolders/sentitems/messages', {
      top: 10,
      select: 'body',
      orderby: 'sentDateTime desc'
    });

    // Analyze HTML content for font styling patterns
    const stylingPatterns = {
      fontFamily: new Map(),
      fontSize: new Map(),
      fontColor: new Map()
    };

    for (const message of sentItems.value) {
      if (message.body?.contentType === 'HTML' && message.body.content) {
        const extracted = extractStylingFromHtml(message.body.content);
        console.log('Extracted styling from email:', extracted);
        if (extracted.fontFamily) {
          stylingPatterns.fontFamily.set(extracted.fontFamily, 
            (stylingPatterns.fontFamily.get(extracted.fontFamily) || 0) + 1);
        }
        if (extracted.fontSize) {
          stylingPatterns.fontSize.set(extracted.fontSize, 
            (stylingPatterns.fontSize.get(extracted.fontSize) || 0) + 1);
        }
        if (extracted.fontColor) {
          stylingPatterns.fontColor.set(extracted.fontColor, 
            (stylingPatterns.fontColor.get(extracted.fontColor) || 0) + 1);
        }
      }
    }

    // Get most common styling
    const mostCommonStyling = {
      fontFamily: getMostCommon(stylingPatterns.fontFamily),
      fontSize: getMostCommon(stylingPatterns.fontSize),
      fontColor: getMostCommon(stylingPatterns.fontColor)
    };

    console.log('Font family patterns found:', Array.from(stylingPatterns.fontFamily.entries()));
    console.log('Most common styling selected:', mostCommonStyling);

    // Cache the results
    stylingCache.set(cacheKey, {
      styling: mostCommonStyling,
      timestamp: Date.now()
    });

    return mostCommonStyling;
  } catch (error) {
    console.warn('Could not extract styling from emails:', error.message);
    return null;
  }
}

// Extract styling information from HTML content
function extractStylingFromHtml(htmlContent) {
  const styling = {};
  
  // Look for font-family in style attributes and CSS - try multiple patterns
  const fontFamilyPatterns = [
    /font-family:\s*([^;,}]+)/gi,
    /font-family\s*=\s*["']([^"']+)["']/gi,
    /<font[^>]*face\s*=\s*["']([^"']+)["']/gi
  ];
  
  for (const pattern of fontFamilyPatterns) {
    const matches = [...htmlContent.matchAll(pattern)];
    for (const match of matches) {
      if (match[1]) {
        const fontFamily = match[1].replace(/["']/g, '').trim();
        // Filter out generic fallbacks and prioritize specific fonts
        if (fontFamily && !fontFamily.includes('sans-serif') && !fontFamily.includes('serif') && !fontFamily.includes('monospace')) {
          styling.fontFamily = fontFamily;
          break;
        }
      }
    }
    if (styling.fontFamily) break;
  }

  // Look for font-size
  const fontSizePatterns = [
    /font-size:\s*([^;,}]+)/gi,
    /<font[^>]*size\s*=\s*["']([^"']+)["']/gi
  ];
  
  for (const pattern of fontSizePatterns) {
    const match = htmlContent.match(pattern);
    if (match) {
      styling.fontSize = match[1].trim();
      break;
    }
  }

  // Look for color
  const colorPatterns = [
    /color:\s*([^;,}]+)/gi,
    /<font[^>]*color\s*=\s*["']([^"']+)["']/gi
  ];
  
  for (const pattern of colorPatterns) {
    const match = htmlContent.match(pattern);
    if (match) {
      styling.fontColor = match[1].trim();
      break;
    }
  }

  // Also check for body tag attributes
  const bodyMatch = htmlContent.match(/<body[^>]*style="([^"]*)"[^>]*>/i);
  if (bodyMatch) {
    const bodyStyle = bodyMatch[1];
    if (!styling.fontFamily) {
      const bodyFontMatch = bodyStyle.match(/font-family:\s*([^;,}]+)/i);
      if (bodyFontMatch) {
        const fontFamily = bodyFontMatch[1].replace(/["']/g, '').trim();
        if (fontFamily && !fontFamily.includes('sans-serif') && !fontFamily.includes('serif')) {
          styling.fontFamily = fontFamily;
        }
      }
    }
  }

  return styling;
}

// Get most common value from a Map
function getMostCommon(map) {
  if (map.size === 0) return null;
  
  let mostCommon = null;
  let maxCount = 0;
  
  for (const [value, count] of map) {
    if (count > maxCount) {
      maxCount = count;
      mostCommon = value;
    }
  }
  
  return mostCommon;
}

// Get user's signature from their mailbox settings
async function getUserSignature(graphApiClient) {
  try {
    // Get user identifier for cache key
    const userInfo = await graphApiClient.makeRequest('/me', { select: 'id' });
    const cacheKey = `signature_${userInfo.id}`;
    
    // Check cache first
    const cachedSignature = signatureCache.get(cacheKey);
    if (cachedSignature && (Date.now() - cachedSignature.timestamp) < SIGNATURE_CACHE_DURATION) {
      console.log('Using cached signature');
      return cachedSignature.signature;
    }
    
    console.log('Searching for user signature in sent emails');
    
    // Search through more sent emails to find genuine signatures (not MCP-generated)
    const sentItems = await graphApiClient.makeRequest('/me/mailFolders/sentitems/messages', {
      top: 50, // Look through more emails
      select: 'body,subject',
      orderby: 'sentDateTime desc'
    });

    let foundSignature = '';
    
    // Look for emails that don't appear to be MCP-generated
    for (const message of sentItems.value) {
      if (message.body?.contentType === 'HTML' && message.body.content) {
        // Skip emails that look like they were generated by our MCP tool
        if (isMcpGeneratedEmail(message.body.content)) {
          continue;
        }
        
        const signature = extractSignatureFromHtml(message.body.content);
        if (signature && signature.trim().length > 10) { // Must be substantial
          foundSignature = signature;
          console.log('Found genuine signature in email');
          break;
        }
      }
    }
    
    // If no signature found in recent emails, try to get from mailbox settings
    if (!foundSignature) {
      try {
        const mailSettings = await graphApiClient.makeRequest('/me/mailboxSettings');
        if (mailSettings?.automaticRepliesSettings?.internalReplyMessage) {
          // Sometimes signature info is in automatic replies
          const autoReplySignature = extractSignatureFromHtml(mailSettings.automaticRepliesSettings.internalReplyMessage);
          if (autoReplySignature) {
            foundSignature = autoReplySignature;
            console.log('Found signature from mailbox settings');
          }
        }
      } catch (error) {
        console.warn('Could not access mailbox settings for signature:', error.message);
      }
    }
    
    // Cache the result (even if empty)
    signatureCache.set(cacheKey, {
      signature: foundSignature,
      timestamp: Date.now()
    });
    
    return foundSignature;
  } catch (error) {
    console.warn('Could not retrieve user signature:', error.message);
    return '';
  }
}

// Check if an email was generated by our MCP tool
function isMcpGeneratedEmail(htmlContent) {
  // Look for indicators that this email was generated by our MCP tool
  const mcpIndicators = [
    // Our styled HTML structure
    /<html>\s*<head>\s*<meta charset="UTF-8">\s*<style>/,
    // Our CSS classes
    /class="email-content"/,
    /class="signature"/,
    // Our specific styling patterns
    /font-family:\s*[^,]+,\s*sans-serif/,
    // Complete HTML structure we generate
    /<html>[\s\S]*<head>[\s\S]*<style>[\s\S]*\.email-content/
  ];
  
  for (const indicator of mcpIndicators) {
    if (indicator.test(htmlContent)) {
      return true;
    }
  }
  
  return false;
}

// Extract signature from HTML content
function extractSignatureFromHtml(htmlContent) {
  // Common signature patterns - ordered by reliability
  const signaturePatterns = [
    // Explicit signature divs
    /<div[^>]*id[^>]*signature[^>]*>.*?<\/div>/is,
    /<div[^>]*class[^>]*signature[^>]*>.*?<\/div>/is,
    
    // Outlook signature patterns
    /<div[^>]*id="Signature"[^>]*>.*?<\/div>/is,
    /<div[^>]*class="OutlookMessageHeader"[^>]*>.*?<\/div>/is,
    
    // Common separators followed by signature content
    /--\s*<br[^>]*>.*$/is,
    /<hr[^>]*>.*$/is,
    
    // Table-based signatures (common in corporate emails)
    /<table[^>]*>.*?<\/table>/is,
    
    // Look for patterns like "Best regards," "Sincerely," etc. followed by contact info
    /(?:Best regards|Sincerely|Thanks|Regards|Cheers|Best|Thank you)[\s\S]*?(?:<br[^>]*>|<\/p>|<\/div>)[\s\S]*?(?:@|phone|tel:|mobile|cell)/is,
    
    // Look for contact information patterns (email, phone, etc.)
    /.*?(?:@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}|(?:\+?1[-.\s]?)?\(?[0-9]{3}\)?[-.\s]?[0-9]{3}[-.\s]?[0-9]{4})[\s\S]*$/is,
    
    // Footer-like content at the end
    /<div[^>]*>.*?(?:@|phone|tel:|mobile|cell|linkedin|twitter).*?<\/div>/is
  ];

  for (const pattern of signaturePatterns) {
    const match = htmlContent.match(pattern);
    if (match && match[0]) {
      const signature = match[0].trim();
      // Filter out very short or generic signatures
      if (signature.length > 20 && !signature.includes('Sent from') && !signature.includes('Get Outlook')) {
        return signature;
      }
    }
  }

  return '';
}

// Convert plain text to HTML with basic formatting
function convertTextToHtml(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
    .replace(/\r?\n/g, '<br>')
    .replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;');
}

// Apply Outlook-like styling to HTML content
async function applyOutlookStyling(htmlContent, signature, mailSettings, actualStyling) {
  // Get user's preferred font and styling from actualStyling (extracted from emails), 
  // then mailSettings, then defaults
  const fontFamily = actualStyling?.fontFamily || mailSettings?.defaultFontName || 'Calibri';
  const fontSize = actualStyling?.fontSize || mailSettings?.defaultFontSize || '11pt';
  const fontColor = actualStyling?.fontColor || mailSettings?.defaultFontColor || '#000000';

  // Create a complete HTML email with proper styling
  const styledHtml = `
    <html>
      <head>
        <meta charset="UTF-8">
        <style>
          body {
            font-family: ${fontFamily}, sans-serif;
            font-size: ${fontSize};
            color: ${fontColor};
            margin: 0;
            padding: 0;
            line-height: 1.4;
          }
          .email-content {
            margin: 0;
            padding: 0;
          }
          .signature {
            margin-top: 20px;
            border-top: 1px solid #e0e0e0;
            padding-top: 10px;
          }
        </style>
      </head>
      <body>
        <div class="email-content">
          ${htmlContent}
        </div>
        ${signature ? `<div class="signature">${signature}</div>` : ''}
      </body>
    </html>
  `;

  return styledHtml;
}