import { Buffer } from 'buffer';
import { Client, ResponseType } from '@microsoft/microsoft-graph-client';
import { authConfig } from '../auth/config.js';
import { convertErrorToToolError, createServiceUnavailableError, createRateLimitError, createValidationError, formatGraphErrorText } from '../utils/mcpErrorResponse.js';
import { FolderResolver } from './folderResolver.js';
import { safeStringify } from '../utils/jsonUtils.js';

/**
 * Percent-encode the characters that corrupt an OData query value on the wire.
 * The Graph SDK appends $filter/$search values to the URL verbatim, and the
 * service decodes a literal '+' in a query string as a space — so a filter on a
 * plus-addressed sender (invoice+statements@...) silently matches nothing, and
 * a literal '&' or '#' truncates the query. Spaces/quotes are already handled
 * by WHATWG URL normalisation, so only these four need escaping.
 */
export function encodeODataQueryValue(value) {
  return String(value)
    .replace(/%/g, '%25')
    .replace(/\+/g, '%2B')
    .replace(/&/g, '%26')
    .replace(/#/g, '%23');
}

/**
 * Read one response header from whatever the SDK attached to the error. The
 * SDK copies the fetch `Headers` instance onto GraphError.headers, which has
 * no enumerable properties — bracket access on it always yields undefined, so
 * `.get()` must be used when present. Plain objects (tests, other transports)
 * are matched case-insensitively.
 */
export function readHeader(headers, name) {
  if (!headers) return undefined;
  if (typeof headers.get === 'function') {
    const value = headers.get(name);
    return value === null || value === undefined ? undefined : value;
  }
  const wanted = name.toLowerCase();
  for (const key of Object.keys(headers)) {
    if (key.toLowerCase() === wanted) return headers[key];
  }
  return undefined;
}

/**
 * Decode GraphError.body into the Graph `error` object ({ code, message,
 * innerError }). The SDK stores it as a JSON *string* of that inner object;
 * for requests made with an ARRAYBUFFER response type the failure body is the
 * undecoded bytes instead. Returns null when the body is not a Graph error.
 */
export function parseGraphErrorBody(body) {
  let parsed = body;
  if (body instanceof ArrayBuffer) parsed = Buffer.from(body).toString('utf8');
  else if (ArrayBuffer.isView(body)) parsed = Buffer.from(body.buffer, body.byteOffset, body.byteLength).toString('utf8');
  if (typeof parsed === 'string') {
    try {
      parsed = JSON.parse(parsed);
    } catch {
      return null;
    }
  }
  if (!parsed || typeof parsed !== 'object') return null;
  return parsed.error && typeof parsed.error === 'object' ? parsed.error : parsed;
}

/**
 * Fill in the fields the SDK leaves empty when a failure body could not be
 * decoded at throw time (binary requests): code and message come from the
 * parsed body so downstream checks such as `code === 'TooManyRequests'` still work.
 */
function normalizeGraphError(error) {
  if (!error || typeof error !== 'object') return error;
  const body = parseGraphErrorBody(error.body);
  if (body) {
    if (!error.code && body.code) error.code = body.code;
    if (!error.message && body.message) error.message = body.message;
  }
  return error;
}

function toBinaryBuffer(value) {
  if (Buffer.isBuffer(value)) return value;
  if (value instanceof ArrayBuffer) return Buffer.from(value);
  if (ArrayBuffer.isView(value)) return Buffer.from(value.buffer, value.byteOffset, value.byteLength);
  return Buffer.alloc(0);
}

// Full URL as sent (host, version, path, query string). buildFullUrl is
// private in the SDK typings but is the only place the final query string is
// assembled, so fall back to the bare path if it ever disappears.
function describeRequestUrl(request, fallback) {
  try {
    if (request && typeof request.buildFullUrl === 'function') return request.buildFullUrl();
  } catch {
    // fall through
  }
  return fallback;
}

const REQUEST_BODY_LOG_LIMIT = 1000;
const REDACTED = '[redacted]';
const BODY_METHODS = new Set(['POST', 'PATCH', 'PUT']);
// Header names that carry credentials; a value that looks like a scheme + credential is covered too.
const SENSITIVE_HEADER_NAME = /auth|key|credential|secret|token|cookie|session/i;
const CREDENTIAL_VALUE = /^\s*(?:bearer|basic)\s/i;
// Body keys whose value is never logged (attachment bytes, anything credential-like).
const SENSITIVE_BODY_KEY = /^contentBytes$|token|secret|password|key|credential|auth/i;
// Short strings under these keys are protocol enumerations, not user content, and stay readable.
const LITERAL_STRING_KEYS = new Set(['@odata.type', 'contentType', 'importance', 'sensitivity', 'showAs', 'type']);
const LITERAL_STRING_MAX = 80;
const BODY_SUMMARY_MAX_DEPTH = 6;
const BODY_SUMMARY_ARRAY_ITEMS = 3;

function truncateForLog(value, limit = REQUEST_BODY_LOG_LIMIT) {
  return value.length > limit ? `${value.slice(0, limit)}... [truncated, ${value.length} chars total]` : value;
}

/**
 * A request body reduced to what is safe to write to a log: the JSON keys and
 * the type/length of each value, never user content. Strings (recipients,
 * subjects, mail text, base64) appear as "<string, N chars>"; only short values
 * under a few protocol-enumeration keys (contentType, importance, ...) are shown
 * as-is. contentBytes and credential-named keys are "[redacted]".
 */
function summariseBodyValue(value, key, depth) {
  if (key !== undefined && SENSITIVE_BODY_KEY.test(key)) return REDACTED;
  if (typeof value === 'string') {
    return LITERAL_STRING_KEYS.has(key) && value.length <= LITERAL_STRING_MAX ? value : `<string, ${value.length} chars>`;
  }
  if (value === null || typeof value === 'number' || typeof value === 'boolean') return value;
  if (Array.isArray(value)) {
    if (depth >= BODY_SUMMARY_MAX_DEPTH) return `<array, ${value.length} items>`;
    const items = value.slice(0, BODY_SUMMARY_ARRAY_ITEMS).map(item => summariseBodyValue(item, undefined, depth + 1));
    if (value.length > BODY_SUMMARY_ARRAY_ITEMS) items.push(`<${value.length - BODY_SUMMARY_ARRAY_ITEMS} more items>`);
    return items;
  }
  if (typeof value === 'object') {
    if (depth >= BODY_SUMMARY_MAX_DEPTH) return '<object>';
    return Object.fromEntries(Object.entries(value).map(([name, inner]) => [name, summariseBodyValue(inner, name, depth + 1)]));
  }
  return `<${typeof value}>`;
}

function summariseBodyForLog(body) {
  return truncateForLog(JSON.stringify(summariseBodyValue(body ?? {}, undefined, 0)));
}

/**
 * The request as it should appear in a failure log: method, full URL, headers
 * and (for POST/PATCH/PUT) a redacted summary of the JSON body (keys and value
 * types/lengths, at most ~1,000 characters; see summariseBodyValue). Credential
 * headers are shown as "[redacted]" and the token is never read. The SDK adds
 * Authorization when the request is sent, so it is listed explicitly (redacted)
 * rather than looked up.
 */
export function describeFailedRequest(request, { method, path, body } = {}) {
  const headers = { Authorization: REDACTED };
  const own = request && typeof request === 'object' ? request._headers : null;
  if (own && typeof own === 'object') {
    for (const [name, value] of Object.entries(own)) {
      headers[name] = SENSITIVE_HEADER_NAME.test(name) || CREDENTIAL_VALUE.test(String(value)) ? REDACTED : String(value);
    }
  }

  const verb = String(method || 'GET').toUpperCase();
  const described = { method: verb, url: describeRequestUrl(request, path), headers };
  if (BODY_METHODS.has(verb)) {
    if (Buffer.isBuffer(body) || body instanceof ArrayBuffer || ArrayBuffer.isView(body)) {
      described.body = `[binary body, ${body.byteLength} bytes]`;
    } else {
      try {
        described.body = summariseBodyForLog(body);
      } catch {
        described.body = '[body could not be summarised]';
      }
    }
  }
  return described;
}

export class GraphApiClient {
  constructor(authManager) {
    this.authManager = authManager;
    this.client = null;
    this.requestCount = 0;
    this.requestWindow = [];
    this.maxConcurrentRequests = 4; // Per mailbox limit from Graph API
    this.activeRequests = 0;
    this.folderResolvers = new Map(); // FolderResolver per mailbox base
    
    // Rate limiting and monitoring metrics
    this.rateLimitMetrics = {
      rateLimitHits: 0,
      totalRetries: 0,
      backoffTime: 0,
      requestsInWindow: 0,
      lastRateLimitHit: null,
      averageRequestDuration: 0,
      requestDurations: []
    };
  }

  async initialize() {
    if (this.client) return this.client;

    const authProvider = {
      getAccessToken: async () => {
        const tokenManager = this.authManager.tokenManager;
        try {
          return await tokenManager.getAccessToken();
        } catch (error) {
          if (error.message.includes('needs refresh')) {
            await this.authManager.refreshAccessToken();
            return await tokenManager.getAccessToken();
          }
          throw error;
        }
      },
    };

    this.client = Client.init({
      authProvider: (done) => {
        authProvider.getAccessToken()
          .then(token => done(null, token))
          .catch(error => done(error, null));
      },
      defaultVersion: 'v1.0',
      debugLogging: process.env.NODE_ENV === 'development',
    });

    this.setupMiddleware();
    return this.client;
  }

  setupMiddleware() {
    // Note: Microsoft Graph SDK handles middleware differently
    // We'll implement rate limiting and retry logic in our makeRequest method instead
    console.error('Graph client initialized with rate limiting and retry logic');
  }

  async enforceRateLimit() {
    // Remove requests older than 1 minute
    const oneMinuteAgo = Date.now() - 60000;
    this.requestWindow = this.requestWindow.filter(time => time > oneMinuteAgo);

    // Wait if we're at the concurrent request limit
    while (this.activeRequests >= this.maxConcurrentRequests) {
      await new Promise(resolve => setTimeout(resolve, 100));
    }

    this.activeRequests++;
    this.requestWindow.push(Date.now());
  }


  generateCorrelationId() {
    return `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
  }

  async sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  extractRetryAfter(error) {
    // Check for Retry-After header (fetch Headers instance or plain object)
    const retryAfter = readHeader(error?.headers, 'retry-after');
    if (retryAfter) {
      const seconds = parseInt(retryAfter);
      return !isNaN(seconds) ? seconds * 1000 : null; // Convert to milliseconds
    }

    // Check in error response body
    const innerError = parseGraphErrorBody(error?.body)?.innerError;
    if (innerError && innerError['retry-after-ms']) {
      return parseInt(innerError['retry-after-ms']);
    }

    return null;
  }

  extractErrorDetails(error, clientRequestId, path, method, url) {
    normalizeGraphError(error);
    const body = parseGraphErrorBody(error?.body);
    const headers = error?.headers;
    // GraphError carries `statusCode` (-1 when the request never got a response)
    const rawStatus = error?.statusCode ?? error?.status;
    const statusCode = Number.isInteger(rawStatus) && rawStatus > 0 ? rawStatus : undefined;
    const bodyInner = body?.innerError || {};

    const microsoftCorrelationIds = {
      requestId: readHeader(headers, 'request-id') || readHeader(headers, 'x-ms-request-id') || error?.requestId || bodyInner['request-id'],
      clientRequestId: readHeader(headers, 'client-request-id') || bodyInner['client-request-id'],
      agsId: readHeader(headers, 'x-ms-ags-diagnostic'),
      correlationId: readHeader(headers, 'x-ms-correlation-id'),
      activityId: readHeader(headers, 'x-ms-activity-id')
    };
    if (bodyInner['request-id']) microsoftCorrelationIds.innerRequestId = bodyInner['request-id'];
    const date = bodyInner['date'] || readHeader(headers, 'date');
    if (date) microsoftCorrelationIds.date = date;

    const errorDetails = {
      timestamp: new Date().toISOString(),
      method,
      path,
      url: url || path,
      clientRequestId,
      statusCode,
      code: body?.code || error?.code,
      message: body?.message || error?.message,
      microsoftCorrelationIds,
      retryAfter: this.extractRetryAfter(error),
      innerError: null
    };

    // Extract detailed error information from response body
    if (body) {
      errorDetails.innerError = {
        code: body.code,
        message: body.message,
        target: body.target,
        details: body.details,
        innerError: body.innerError
      };
    }

    // Structured copy for callers (never contains the Authorization header or token).
    // We always send our own client-request-id, so it stays known even when Graph does not echo it.
    errorDetails.graphError = {
      status: statusCode ?? null,
      code: errorDetails.code || null,
      message: errorDetails.message || null,
      requestId: microsoftCorrelationIds.requestId || null,
      clientRequestId: microsoftCorrelationIds.clientRequestId || clientRequestId || null,
      date: date || null,
      method: method || null,
      url: errorDetails.url || null
    };

    return errorDetails;
  }

  // stderr only (never part of the MCP error returned to the caller)
  logFailedRequest(request, method, path, options = {}) {
    try {
      const details = describeFailedRequest(request, { method, path, body: options.body || {} });
      console.error('Failed request:', JSON.stringify(details, null, 2));
    } catch (logError) {
      console.error('Failed request could not be described:', logError.message);
    }
  }

  async makeRequest(path, options = {}, method = 'GET') {
    await this.initialize();
    
    // Generate correlation ID for this request
    const clientRequestId = this.generateCorrelationId();
    const requestStartTime = Date.now();
    
    // Enforce rate limiting before making request
    await this.enforceRateLimit();
    
    const maxRetries = authConfig.retry.maxAttempts;
    let retryCount = 0;
    let delay = authConfig.retry.initialDelay; // Start at exactly 1 second

    while (retryCount <= maxRetries) {
      let request = null;
      try {
        request = this.client.api(path);

        // Add correlation ID header
        request = request.header('client-request-id', clientRequestId);

        // Apply common query parameters
        if (options.select) {
          request = request.select(options.select);
        }
        if (options.top) {
          request = request.top(options.top);
        }
        if (options.filter) {
          request = request.filter(encodeODataQueryValue(options.filter));
        }
        if (options.orderby) {
          request = request.orderby(options.orderby);
        }
        if (options.expand) {
          request = request.expand(options.expand);
        }
        if (options.search) {
          request = request.search(encodeODataQueryValue(options.search));
        }
        if (options.startDateTime && options.endDateTime) {
          request = request.query({
            startDateTime: options.startDateTime,
            endDateTime: options.endDateTime
          });
        }
        if (options.responseType) {
          request = request.responseType(options.responseType);
        }

        // Log request attempt
        if (retryCount > 0) {
          console.error(`Graph API retry attempt ${retryCount} for ${method} ${path} [correlation: ${clientRequestId}]`);
        }

        let response;
        // Execute the appropriate method
        switch (method.toUpperCase()) {
          case 'POST':
            response = await request.post(options.body || {});
            break;
          case 'PATCH':
            response = await request.patch(options.body || {});
            break;
          case 'PUT':
            response = await request.put(options.body || {});
            break;
          case 'DELETE':
            response = await request.delete();
            break;
          default:
            response = await request.get();
        }

        // Success - decrement active requests and return
        this.activeRequests--;
        
        const requestDuration = Date.now() - requestStartTime;
        this.updateMetrics(requestDuration);
        console.error(`Graph API success: ${method} ${path} (${requestDuration}ms) [correlation: ${clientRequestId}]`);
        
        return response;

      } catch (error) {
        this.activeRequests--;
        
        // Extract detailed error information including Microsoft's correlation IDs
        const errorDetails = this.extractErrorDetails(error, clientRequestId, path, method, describeRequestUrl(request, path));
        
        // Handle rate limiting (429 responses)
        if (error.code === 'TooManyRequests' || error.status === 429) {
          const retryAfter = this.extractRetryAfter(error);
          const waitTime = retryAfter || delay;
          
          // Update rate limit metrics
          this.rateLimitMetrics.rateLimitHits++;
          this.rateLimitMetrics.lastRateLimitHit = new Date().toISOString();
          this.rateLimitMetrics.backoffTime += waitTime;
          
          console.warn(`Rate limited on ${method} ${path}. Waiting ${waitTime}ms before retry ${retryCount + 1}/${maxRetries} [correlation: ${clientRequestId}]`);
          console.warn('Rate limit details:', safeStringify(errorDetails, 2));
          console.warn('Rate limit metrics:', safeStringify(this.getRateLimitMetrics(), 2));
          
          if (retryCount < maxRetries) {
            await this.sleep(waitTime);
            retryCount++;
            this.rateLimitMetrics.totalRetries++;
            // Use exponential backoff for next attempt if no Retry-After header
            if (!retryAfter) {
              delay = Math.min(delay * authConfig.retry.backoffMultiplier, authConfig.retry.maxDelay);
            }
            continue;
          } else {
            this.logFailedRequest(request, method, path, options);
            // Return rate limit error instead of throwing
            return createRateLimitError(Math.ceil(waitTime / 1000));
          }
        }
        
        // Handle server errors (5xx) with exponential backoff
        if (error.status >= 500 && error.status < 600) {
          console.warn(`Server error ${error.status} on ${method} ${path}. Retry ${retryCount + 1}/${maxRetries} after ${delay}ms [correlation: ${clientRequestId}]`);
          console.warn('Server error details:', safeStringify(errorDetails, 2));
          
          if (retryCount < maxRetries) {
            await this.sleep(delay);
            retryCount++;
            delay = Math.min(delay * authConfig.retry.backoffMultiplier, authConfig.retry.maxDelay);
            continue;
          } else {
            this.logFailedRequest(request, method, path, options);
            // Return service unavailable error instead of throwing
            return createServiceUnavailableError('Microsoft Graph API');
          }
        }
        
        // Handle authentication errors
        if (error.status === 401 || error.code === 'InvalidAuthenticationToken') {
          console.warn(`Authentication error on ${method} ${path}. Attempting token refresh [correlation: ${clientRequestId}]`);
          try {
            await this.authManager.refreshAccessToken();
            if (retryCount < maxRetries) {
              retryCount++;
              continue; // Retry with new token
            }
          } catch (refreshError) {
            console.error('Token refresh failed:', refreshError.message);
          }
        }
        
        // Log final error and return MCP error
        console.error(`Graph API error: ${method} ${errorDetails.url} [correlation: ${clientRequestId}]`);
        console.error('Error details:', safeStringify(errorDetails, 2));
        this.logFailedRequest(request, method, path, options);
        
        return this.handleGraphError(error, errorDetails);
      }
    }

    // If we get here, all retries have been exhausted
    return createServiceUnavailableError(`Microsoft Graph API (after ${maxRetries} retry attempts)`);
  }

  async makeBatchRequest(requests) {
    if (requests.length > 20) {
      return createValidationError('requests', 'Batch requests are limited to 20 operations');
    }

    await this.initialize();

    const batchContent = {
      requests: requests.map((req, index) => ({
        id: String(index + 1),
        method: req.method || 'GET',
        url: req.url,
        body: req.body,
        headers: req.headers,
      })),
    };

    let request = null;
    try {
      request = this.client.api('/$batch');
      const response = await request.post(batchContent);
      return response.responses;
    } catch (error) {
      this.logFailedRequest(request, 'POST', '/$batch', { body: batchContent });
      return this.handleGraphError(error, this.extractErrorDetails(error, null, '/$batch', 'POST'));
    }
  }

  handleGraphError(error, enhancedErrorDetails = null) {
    // Use enhanced error details if provided, otherwise extract them from the error alone
    const errorDetails = enhancedErrorDetails || this.extractErrorDetails(error);

    // Log comprehensive error details for debugging
    console.error('Graph API Error - Full Details:', safeStringify(errorDetails, 2));

    const correlationIds = errorDetails.microsoftCorrelationIds || {};
    const graphError = errorDetails.graphError;

    // Short human prefix; the verbatim Graph diagnostics follow it (see formatGraphErrorText)
    let prefix = '';
    switch (errorDetails.statusCode) {
      case 400:
        prefix = 'Bad request.';
        break;
      case 401:
        prefix = 'Authentication failed. Please re-authenticate.';
        break;
      case 403:
        prefix = 'Insufficient permissions. Please check your app permissions in Azure AD.';
        break;
      case 404:
        prefix = 'Resource not found. The requested item may have been deleted or moved.';
        break;
      case 409:
        prefix = 'Conflict detected. This may be due to concurrent updates or scheduling conflicts.';
        break;
      case 429: {
        // Return specific rate limit error
        const retryAfterMs = this.extractRetryAfter(error) || 60000;
        return createRateLimitError(Math.ceil(retryAfterMs / 1000));
      }
      case 500:
      case 502:
      case 503:
      case 504: {
        // Keep the retryable service-unavailable shape but do not drop what Graph said
        const unavailable = createServiceUnavailableError('Microsoft Graph API');
        unavailable.content[0].text += ` ${formatGraphErrorText(graphError)}`;
        unavailable._errorDetails = { ...unavailable._errorDetails, graphError };
        return unavailable;
      }
      default:
        if (errorDetails.code === 'InvalidAuthenticationToken') {
          prefix = 'Invalid or expired authentication token. Please re-authenticate.';
        } else {
          prefix = 'Graph API error.';
        }
    }

    // Create error with additional properties for MCP error conversion
    const finalError = new Error(`${prefix} ${formatGraphErrorText(graphError)}`);
    finalError.originalError = error;
    finalError.correlationIds = correlationIds;
    finalError.statusCode = errorDetails.statusCode;
    if (errorDetails.code) finalError.code = errorDetails.code;
    finalError.graphError = graphError;
    finalError.retryable = this.isRetryableError(errorDetails.statusCode);

    return convertErrorToToolError(finalError, 'Microsoft Graph API');
  }

  isRetryableError(statusCode) {
    // Define which errors are retryable
    return [401, 429, 500, 502, 503, 504].includes(statusCode);
  }

  // Get FolderResolver for a mailbox base (lazy, cached per base)
  getFolderResolver(mailboxBase = '/me') {
    if (!this.folderResolvers.has(mailboxBase)) {
      this.folderResolvers.set(mailboxBase, new FolderResolver(this, mailboxBase));
    }
    return this.folderResolvers.get(mailboxBase);
  }

  // Utility methods for common operations
  async getWithSelect(path, fields) {
    return this.makeRequest(path, { select: fields.join(',') }, 'GET');
  }

  async postWithRetry(path, body) {
    return this.makeRequest(path, { body }, 'POST');
  }

  async patchWithRetry(path, body) {
    return this.makeRequest(path, { body }, 'PATCH');
  }

  async deleteWithRetry(path) {
    return this.makeRequest(path, {}, 'DELETE');
  }

  /**
   * GET a resource as raw bytes (e.g. an attachment's /$value). Shares
   * makeRequest's retry and error handling: resolves to a Buffer on success
   * and to the same MCP error object makeRequest returns on failure.
   */
  async getBinary(path, options = {}) {
    const result = await this.makeRequest(path, { ...options, responseType: ResponseType.ARRAYBUFFER }, 'GET');
    if (result && result.isError !== undefined && result.content) return result;
    return toBinaryBuffer(result);
  }

  // Pagination helper
  async *iterateAllPages(path, options = {}) {
    let nextLink = null;
    
    do {
      const response = nextLink 
        ? await this.client.api(nextLink).get()
        : await this.makeRequest(path, options, 'GET');
      
      yield response.value || [];
      nextLink = response['@odata.nextLink'];
    } while (nextLink);
  }

  // Rate limit monitoring and metrics
  updateMetrics(requestDuration) {
    // Track request durations for average calculation
    this.rateLimitMetrics.requestDurations.push(requestDuration);
    
    // Keep only last 100 request durations for rolling average
    if (this.rateLimitMetrics.requestDurations.length > 100) {
      this.rateLimitMetrics.requestDurations.shift();
    }
    
    // Calculate average request duration
    const sum = this.rateLimitMetrics.requestDurations.reduce((a, b) => a + b, 0);
    this.rateLimitMetrics.averageRequestDuration = Math.round(sum / this.rateLimitMetrics.requestDurations.length);
    
    // Update requests in current window
    this.rateLimitMetrics.requestsInWindow = this.requestWindow.length;
  }

  getRateLimitMetrics() {
    return {
      ...this.rateLimitMetrics,
      activeRequests: this.activeRequests,
      requestsInWindow: this.requestWindow.length,
      lastUpdated: new Date().toISOString()
    };
  }

  resetMetrics() {
    this.rateLimitMetrics = {
      rateLimitHits: 0,
      totalRetries: 0,
      backoffTime: 0,
      requestsInWindow: 0,
      lastRateLimitHit: null,
      averageRequestDuration: 0,
      requestDurations: []
    };
  }

  // Monitoring alerts
  checkRateLimitHealth() {
    const metrics = this.getRateLimitMetrics();
    const alerts = [];

    // Alert if rate limit hits are frequent
    if (metrics.rateLimitHits > 5) {
      alerts.push({
        level: 'warning',
        message: `High rate limit hits detected: ${metrics.rateLimitHits}`,
        suggestion: 'Consider implementing request queuing or reducing request frequency'
      });
    }

    // Alert if average request duration is high
    if (metrics.averageRequestDuration > 5000) {
      alerts.push({
        level: 'warning',
        message: `High average request duration: ${metrics.averageRequestDuration}ms`,
        suggestion: 'Graph API performance may be degraded'
      });
    }

    // Alert if too many active requests
    if (metrics.activeRequests >= this.maxConcurrentRequests) {
      alerts.push({
        level: 'error',
        message: `Maximum concurrent requests reached: ${metrics.activeRequests}/${this.maxConcurrentRequests}`,
        suggestion: 'Requests are being queued due to concurrency limits'
      });
    }

    return alerts;
  }
}