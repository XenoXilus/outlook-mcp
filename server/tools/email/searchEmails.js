import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';
import { stripHtml, truncateText } from '../../utils/textUtils.js';
import { getMailboxBase, toStartDateTime, endDateFilterClause } from '../../graph/graphHelpers.js';

// Search emails with intelligent KQL/OData strategy selection
export async function searchEmailsTool(authManager, args) {
  const {
    query,
    subject,
    from,
    startDate,
    endDate,
    folders = [],
    limit = 25,
    includeBody = false, // Changed default to false
    truncate = true,
    maxLength = 1000,
    format = 'text',
    orderBy = 'receivedDateTime desc',
    mailbox
  } = args;

  // Cap limit at 5 when includeBody is true to prevent context overflow
  const effectiveLimit = includeBody ? Math.min(limit, 5) : limit;
  const limitWasCapped = includeBody && limit > 5;

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    const mailboxBase = getMailboxBase(mailbox);

    const options = {
      top: Math.min(effectiveLimit, 1000) // Cap at 1000 for performance
      // orderby will be added conditionally after determining if search is used
    };

    // Build select based on includeBody parameter
    if (includeBody) {
      options.select = 'id,subject,from,toRecipients,receivedDateTime,sentDateTime,body,bodyPreview,importance,isRead,hasAttachments,conversationId';
    } else {
      options.select = 'id,subject,from,toRecipients,receivedDateTime,sentDateTime,bodyPreview,importance,isRead,hasAttachments,conversationId';
    }

    // Resolve folder names to IDs if provided
    let resolvedFolderIds = [];
    if (folders.length > 0) {
      try {
        const folderResolver = graphApiClient.getFolderResolver(mailboxBase);
        resolvedFolderIds = await folderResolver.resolveFoldersToIds(folders);
      } catch (folderError) {
        return createValidationError('folders', folderError.message);
      }
    }

    // Determine search strategy and endpoint
    let endpoint = `${mailboxBase}/messages`;
    let useKQLSearch = false;
    let useODataFilters = false;
    const isSpecificFolder = resolvedFolderIds.length === 1;

    if (isSpecificFolder) {
      // Single folder search
      endpoint = `${mailboxBase}/mailFolders/${resolvedFolderIds[0]}/messages`;
      useODataFilters = true;
    } else if (resolvedFolderIds.length > 1) {
      // Multiple folders - we'll need to make separate requests and combine
      // For now, fall back to all folders search
      endpoint = `${mailboxBase}/messages`;
      useODataFilters = true;
    } else {
      // All folders search (folders.length === 0)
      endpoint = `${mailboxBase}/messages`;

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
      // Date-only bounds are expanded so the window covers the whole final day:
      // a bare `le 2026-08-31` means midnight and silently drops the 31st.
      if (orderBy && orderBy.includes('receivedDateTime')) {
        if (startDate) {
          filterConditions.push(`receivedDateTime ge ${toStartDateTime(startDate)}`);
        } else {
          // Add a broad receivedDateTime filter to satisfy API requirements
          filterConditions.push(`receivedDateTime ge 1900-01-01T00:00:00Z`);
        }

        if (endDate) {
          filterConditions.push(endDateFilterClause(endDate));
        }
      } else {
        // Add date filters normally if not using receivedDateTime orderby
        if (startDate) {
          filterConditions.push(`receivedDateTime ge ${toStartDateTime(startDate)}`);
        }

        if (endDate) {
          filterConditions.push(endDateFilterClause(endDate));
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

    // Handle MCP error responses from makeRequest
    if (result.content && result.isError !== undefined) {
      return result;
    }

    const emails = result.value?.map(email => {
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
        let processedContent = email.body.content || '';
        let contentType = email.body.contentType || 'Text';

        // Strip HTML if requested (default: true, unless format is explicitly 'html')
        if (format === 'text' && contentType === 'html') {
          processedContent = stripHtml(processedContent);
          contentType = 'text';
        }

        // Truncate if requested (default: true)
        if (truncate) {
          processedContent = truncateText(processedContent, maxLength);
          emailData.truncated = true;
        }

        emailData.body = {
          contentType: contentType,
          content: processedContent
        };
      }

      return emailData;
    }) || [];

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
      limitWasCapped: limitWasCapped,
      optimization: useKQLSearch ? 'Using KQL for text-based search (efficient for content search)' :
        isSpecificFolder ? 'Using $filter for specific folder search (comprehensive)' :
          'Using $filter for all-folders search (comprehensive across all folders)'
    };

    return createSafeResponse({
      searchSummary,
      emails
    });
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to search emails');
  }
}