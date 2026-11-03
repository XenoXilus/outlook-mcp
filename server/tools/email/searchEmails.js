import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';
import { stripHtml, truncateText } from '../../utils/textUtils.js';
import {
  getMailboxBase,
  toStartDateTime,
  toExclusiveEndDateTime,
  endDateFilterClause
} from '../../graph/graphHelpers.js';

const DATE_ONLY = /^\d{4}-\d{2}-\d{2}$/;
const HAS_UTC_OFFSET = /(?:Z|[+-]\d{2}:?\d{2})$/i;
// Whitespace separates KQL terms; quotes, parentheses, ':' and comparison
// characters are KQL syntax. A value containing any of them has to be quoted.
const KQL_NEEDS_QUOTES = /[\s"():<>=]/;

const DAY_MS = 24 * 60 * 60 * 1000;
const MAX_PAGE_SIZE = 1000;

const text = (value) => String(value ?? '').trim();

// Epoch millis for a date or date-time. Date-only and zone-less input is read
// as UTC, never local time. Returns null when the value cannot be parsed.
function toUtcMillis(value) {
  const input = text(value);
  const iso = DATE_ONLY.test(input) ? `${input}T00:00:00Z` : (HAS_UTC_OFFSET.test(input) ? input : `${input}Z`);
  const millis = Date.parse(iso);
  return Number.isNaN(millis) ? null : millis;
}

// KQL date operand: the UTC calendar date as YYYY-MM-DD, independent of the
// host's locale and timezone (toLocaleDateString is both, and M/D/YYYY is ambiguous).
// `offsetDays` shifts the calendar date (UTC days are a fixed 24 hours).
function toKqlDate(value, offsetDays = 0) {
  return new Date(toUtcMillis(value) + offsetDays * DAY_MS).toISOString().slice(0, 10);
}

// A KQL value is quoted only when it needs to be. Inner quotes are dropped:
// the whole expression is quoted again for $search, and a nested quote cannot
// be expressed unambiguously inside a quoted value.
function toKqlValue(value) {
  return KQL_NEEDS_QUOTES.test(value) ? `"${value.replace(/"/g, '')}"` : value;
}

// KQL clauses for each requested restriction. Property restrictions stay
// unquoted (`from:value`, `received:2026-09-30..2026-11-01`): KQL reads a
// quoted `received:...` as a literal phrase, not a property restriction, and
// only the value after the colon may be quoted.
//
// Whether the service treats the end day of a `received:` range as inclusive,
// and in which time zone it evaluates the dates, is not something we can rely
// on. The date restriction is therefore widened by one day on each side (start
// minus one day, end plus one day) so no message in the requested window can
// be missed; the client-side guard trims the result to the exact window.
function buildKqlClauses({ queryText, fromText, subjectText, startBound, endBound }) {
  let received = null;
  if (startBound && endBound) {
    received = `received:${toKqlDate(startBound, -1)}..${toKqlDate(endBound, 1)}`;
  } else if (startBound) {
    received = `received>=${toKqlDate(startBound, -1)}`;
  } else if (endBound) {
    received = `received<=${toKqlDate(endBound, 1)}`;
  }

  return {
    text: queryText ? toKqlValue(queryText) : null,
    from: fromText ? `from:${toKqlValue(fromText)}` : null,
    subject: subjectText ? `subject:${toKqlValue(subjectText)}` : null,
    received
  };
}

// Predicate over raw Graph messages that re-checks every requested restriction
// client-side, so a response the service failed to filter is never returned as-is.
function buildRestrictionGuard({ fromText, subjectText, startBound, endBound }) {
  const startMs = startBound ? toUtcMillis(toStartDateTime(startBound)) : null;
  // End semantics mirror endDateFilterClause: a date-only end covers the whole
  // day (exclusive next midnight), an explicit date-time end is inclusive.
  const exclusiveEnd = endBound ? toExclusiveEndDateTime(endBound) : null;
  const endMs = endBound ? toUtcMillis(exclusiveEnd || endBound) : null;
  const fromNeedle = fromText.toLowerCase();
  const subjectNeedle = subjectText.toLowerCase();
  const contains = (haystack, needle) => String(haystack ?? '').toLowerCase().includes(needle);

  return (message) => {
    if (startMs !== null || endMs !== null) {
      const receivedMs = toUtcMillis(message.receivedDateTime);
      // A message whose date cannot be checked against the window is dropped
      if (receivedMs === null) return false;
      if (startMs !== null && receivedMs < startMs) return false;
      if (endMs !== null && (exclusiveEnd ? receivedMs >= endMs : receivedMs > endMs)) return false;
    }

    if (fromNeedle) {
      const sender = message.from?.emailAddress;
      if (!contains(sender?.address, fromNeedle) && !contains(sender?.name, fromNeedle)) return false;
    }

    if (subjectNeedle && !contains(message.subject, subjectNeedle)) return false;

    return true;
  };
}

// Merge order for results gathered from several folder requests: follows the
// first $orderby clause (e.g. 'receivedDateTime desc').
function compareByOrderBy(orderBy) {
  const [field = 'receivedDateTime', direction = 'asc'] = text(orderBy).split(',')[0].trim().split(/\s+/);
  const sign = direction.toLowerCase() === 'desc' ? -1 : 1;
  const read = (message) => field.split('/').reduce((value, key) => value?.[key], message) ?? '';
  return (a, b) => {
    const left = read(a);
    const right = read(b);
    if (left === right) return 0;
    return left < right ? -sign : sign;
  };
}

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

  const queryText = text(query);
  const fromText = text(from);
  const subjectText = text(subject);
  const startBound = text(startDate);
  const endBound = text(endDate);

  // Cap limit at 5 when includeBody is true to prevent context overflow
  const effectiveLimit = includeBody ? Math.min(limit, 5) : limit;
  const limitWasCapped = includeBody && limit > 5;

  try {
    for (const [name, bound] of [['startDate', startBound], ['endDate', endBound]]) {
      if (bound && toUtcMillis(bound) === null) {
        return createValidationError(name, `'${bound}' is not a valid ISO 8601 date or date-time`);
      }
    }

    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    const mailboxBase = getMailboxBase(mailbox);

    // Number of messages the caller asked for (capped at 1000 for performance)
    const resultLimit = Math.min(effectiveLimit, MAX_PAGE_SIZE);
    const options = {
      top: resultLimit
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
    const folderIds = [...new Set(resolvedFolderIds)];

    // Determine search strategy and endpoint(s)
    let endpoints = [`${mailboxBase}/messages`];
    let useKQLSearch = false;
    let useODataFilters = false;
    let kql = null;

    if (folderIds.length > 0) {
      // A folder restriction is bound by the endpoint itself: one request per
      // folder (results are merged below), never a fall-back to all folders.
      endpoints = folderIds.map(folderId => `${mailboxBase}/mailFolders/${folderId}/messages`);
      useODataFilters = true;
    } else if (queryText) {
      // Use KQL search for text queries (more efficient for content search)
      useKQLSearch = true;
    } else {
      // Use OData filters for sender/subject/date searches (more comprehensive)
      useODataFilters = true;
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
        if (startBound) {
          filterConditions.push(`receivedDateTime ge ${toStartDateTime(startBound)}`);
        } else {
          // Add a broad receivedDateTime filter to satisfy API requirements
          filterConditions.push(`receivedDateTime ge 1900-01-01T00:00:00Z`);
        }

        if (endBound) {
          filterConditions.push(endDateFilterClause(endBound));
        }
      } else {
        // Add date filters normally if not using receivedDateTime orderby
        if (startBound) {
          filterConditions.push(`receivedDateTime ge ${toStartDateTime(startBound)}`);
        }

        if (endBound) {
          filterConditions.push(endDateFilterClause(endBound));
        }
      }

      if (fromText) {
        filterConditions.push(`from/emailAddress/address eq '${fromText.replace(/'/g, "''")}'`);
      }

      if (subjectText) {
        filterConditions.push(`contains(subject,'${subjectText.replace(/'/g, "''")}')`);
      }

      if (queryText) {
        // Use contains for general text search. Parenthesised: an unparenthesised
        // `or` would split the whole conjunction and let body matches bypass
        // every other restriction.
        const escapedQuery = queryText.replace(/'/g, "''");
        filterConditions.push(`(contains(subject,'${escapedQuery}') or contains(body/content,'${escapedQuery}'))`);
      }

      if (filterConditions.length > 0) {
        options.filter = filterConditions.join(' and ');
      }

      // Add orderby for OData filter searches
      options.orderby = orderBy;

    } else if (useKQLSearch) {
      // Graph does not allow $filter (or $orderby) together with $search on message
      // collections, so a date window cannot move to $filter here: it travels inside
      // the KQL, and the results are re-checked against it below.
      kql = buildKqlClauses({ queryText, fromText, subjectText, startBound, endBound });
      if (kql.received) {
        // The widened window also matches the margin-day messages the guard will drop
        // again, and they use up page slots; ask for more than the limit so they
        // cannot crowd out in-window messages (the result is trimmed to the limit below).
        options.top = Math.min(Math.max(resultLimit * 2, resultLimit + 25), MAX_PAGE_SIZE);
      }
      // Graph expects the whole KQL expression as one double-quoted $search value;
      // quotes inside it are backslash-escaped.
      const expression = [kql.text, kql.from, kql.subject, kql.received].filter(Boolean).join(' AND ');
      options.search = `"${expression.replace(/"/g, '\\"')}"`;
    }

    // Make the request(s) using chosen search strategy
    const fetched = [];
    // A request that came back with a full page may have been cut off by the service
    let receivedFullPage = false;
    for (const endpoint of endpoints) {
      const result = await graphApiClient.makeRequest(endpoint, options);

      // Handle MCP error responses from makeRequest
      if (result.content && result.isError !== undefined) {
        return result;
      }

      const page = result.value || [];
      if (page.length >= options.top) receivedFullPage = true;
      fetched.push(...page);
    }

    // Defence in depth: whatever the service did with the query, never return a
    // message that violates a requested date window, sender, or subject.
    const matchesRestrictions = buildRestrictionGuard({ fromText, subjectText, startBound, endBound });
    let messages = fetched.filter(matchesRestrictions);
    const droppedOutsideFilters = fetched.length - messages.length;

    // A widened KQL window returns the margin days by design; only drops the
    // widening does not explain mean the service ignored a restriction.
    const kqlWindowWidened = useKQLSearch && Boolean(kql.received);
    let unexpectedDrops = droppedOutsideFilters;
    if (kqlWindowWidened) {
      const withinMargin = buildRestrictionGuard({
        fromText,
        subjectText,
        startBound: startBound ? toKqlDate(startBound, -1) : '',
        endBound: endBound ? toKqlDate(endBound, 1) : ''
      });
      unexpectedDrops = fetched.filter(message => !withinMargin(message)).length;
    }

    // The guard may have left more than the caller asked for (a widened window
    // requests extra, and a merge of several folders can exceed it).
    if (endpoints.length > 1) {
      messages = messages.sort(compareByOrderBy(orderBy));
    }
    messages = messages.slice(0, resultLimit);

    // A full page with anything dropped means in-window messages may have been
    // cut off by the page size before the guard could see them.
    const possiblyTruncated = receivedFullPage && droppedOutsideFilters > 0;
    const warnings = [];
    if (unexpectedDrops > 0) {
      warnings.push(
        `The service returned ${unexpectedDrops} message${unexpectedDrops === 1 ? '' : 's'} outside the requested filters; ` +
        `${unexpectedDrops === 1 ? 'it was' : 'they were'} removed, so the result may be smaller than the requested limit ` +
        '(results possibly incomplete - narrow the date window or raise limit).'
      );
    }
    if (possiblyTruncated) {
      warnings.push(
        `The service returned a full page of ${options.top} messages and ${droppedOutsideFilters} of them fell outside the requested filters, ` +
        'so matching messages may have been cut off ' +
        '(results possibly incomplete - narrow the date window or raise limit).'
      );
    }

    const emails = messages.map(email => {
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
    });

    let dateRange = null;
    if (useKQLSearch) {
      dateRange = kql.received ? `KQL: ${kql.received}` : null;
    } else if (startBound && endBound) {
      dateRange = `Filter: receivedDateTime ge ${startBound} and receivedDateTime le ${endBound}`;
    } else if (startBound) {
      dateRange = `Filter: receivedDateTime ge ${startBound}`;
    } else if (endBound) {
      dateRange = `Filter: receivedDateTime le ${endBound}`;
    }

    const searchSummary = {
      searchApproach: useKQLSearch ? 'KQL (Keyword Query Language)' : 'OData $filter',
      kqlQuery: options.search || null,
      filterQuery: options.filter || null,
      endpoint: endpoints.length === 1 ? endpoints[0] : endpoints,
      folders: folders.length > 0 ? folders : ['All folders'],
      parameters: {
        generalSearch: queryText || null,
        sender: fromText ? (useKQLSearch ? `KQL: ${kql.from}` : `Filter: from/emailAddress/address eq '${fromText}'`) : null,
        subject: subjectText ? (useKQLSearch ? `KQL: ${kql.subject}` : `Filter: contains(subject,'${subjectText}')`) : null,
        dateRange
      },
      totalResults: emails.length,
      droppedOutsideFilters,
      ...(kqlWindowWidened && { kqlWindowWidened: true }),
      ...(warnings.length > 0 && { warning: warnings.join(' ') }),
      includesFullBody: includeBody,
      limitWasCapped: limitWasCapped,
      optimization: useKQLSearch ? 'Using KQL for text-based search (efficient for content search)' :
        folderIds.length === 1 ? 'Using $filter for specific folder search (comprehensive)' :
          folderIds.length > 1 ? 'Using $filter with one request per folder, merged (comprehensive)' :
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
