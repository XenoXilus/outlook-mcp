import { Client } from '@microsoft/microsoft-graph-client';
import { authConfig } from '../auth/config.js';
import { convertErrorToToolError, createServiceUnavailableError, createRateLimitError, createValidationError } from '../utils/mcpErrorResponse.js';
import { FolderResolver } from './folderResolver.js';
import { safeStringify } from '../utils/jsonUtils.js';

export class GraphApiClient {
  constructor(authManager) {
    this.authManager = authManager;
    this.client = null;
    this.requestCount = 0;
    this.requestWindow = [];
    this.maxConcurrentRequests = 4; // Per mailbox limit from Graph API
    this.activeRequests = 0;
    this.folderResolver = null; // Will be initialized when needed
    
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
    // Check for Retry-After header in various formats
    if (error.headers) {
      const retryAfter = error.headers['retry-after'] || error.headers['Retry-After'];
      if (retryAfter) {
        const seconds = parseInt(retryAfter);
        return !isNaN(seconds) ? seconds * 1000 : null; // Convert to milliseconds
      }
    }
    
    // Check in error response body
    if (error.body && error.body.error) {
      const innerError = error.body.error.innerError;
      if (innerError && innerError['retry-after-ms']) {
        return parseInt(innerError['retry-after-ms']);
      }
    }
    
    return null;
  }

  extractErrorDetails(error, clientRequestId, path, method) {
    const errorDetails = {
      timestamp: new Date().toISOString(),
      method,
      path,
      clientRequestId,
      statusCode: error.status || error.statusCode,
      code: error.code,
      message: error.message,
      microsoftCorrelationIds: {},
      retryAfter: this.extractRetryAfter(error),
      innerError: null
    };

    // Extract Microsoft's correlation IDs from headers
    if (error.headers) {
      errorDetails.microsoftCorrelationIds = {
        requestId: error.headers['request-id'] || error.headers['x-ms-request-id'],
        clientRequestId: error.headers['client-request-id'],
        agsId: error.headers['x-ms-ags-diagnostic'],
        correlationId: error.headers['x-ms-correlation-id'],
        activityId: error.headers['x-ms-activity-id']
      };
    }

    // Extract detailed error information from response body
    if (error.body && error.body.error) {
      const graphError = error.body.error;
      errorDetails.innerError = {
        code: graphError.code,
        message: graphError.message,
        target: graphError.target,
        details: graphError.details,
        innerError: graphError.innerError
      };

      // Extract additional correlation IDs from inner error
      if (graphError.innerError) {
        if (graphError.innerError['request-id']) {
          errorDetails.microsoftCorrelationIds.innerRequestId = graphError.innerError['request-id'];
        }
        if (graphError.innerError['date']) {
          errorDetails.microsoftCorrelationIds.date = graphError.innerError['date'];
        }
      }
    }

    return errorDetails;
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
      try {
        let request = this.client.api(path);

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
          request = request.filter(options.filter);
        }
        if (options.orderby) {
          request = request.orderby(options.orderby);
        }
        if (options.expand) {
          request = request.expand(options.expand);
        }
        if (options.search) {
          request = request.search(options.search);
        }
        if (options.startDateTime && options.endDateTime) {
          request = request.query({
            startDateTime: options.startDateTime,
            endDateTime: options.endDateTime
          });
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
        const errorDetails = this.extractErrorDetails(error, clientRequestId, path, method);
        
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
        console.error(`Graph API error: ${method} ${path} [correlation: ${clientRequestId}]`);
        console.error('Error details:', safeStringify(errorDetails, 2));
        
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

    try {
      const response = await this.client.api('/$batch').post(batchContent);
      return response.responses;
    } catch (error) {
      return this.handleGraphError(error);
    }
  }

  handleGraphError(error, enhancedErrorDetails = null) {
    // Use enhanced error details if provided, otherwise extract basic details
    const errorDetails = enhancedErrorDetails || {
      timestamp: new Date().toISOString(),
      statusCode: error.status || error.statusCode,
      code: error.code,
      message: error.message,
      microsoftCorrelationIds: {
        requestId: error.headers?.['request-id'] || 'unknown'
      }
    };

    // Log comprehensive error details for debugging
    console.error('Graph API Error - Full Details:', safeStringify(errorDetails, 2));

    // Create user-friendly error message with correlation IDs for support
    let userMessage = '';
    let supportInfo = '';

    // Build support information with correlation IDs
    const correlationIds = errorDetails.microsoftCorrelationIds || {};
    const supportCorrelationIds = Object.entries(correlationIds)
      .filter(([key, value]) => value && value !== 'unknown')
      .map(([key, value]) => `${key}: ${value}`)
      .join(', ');

    if (supportCorrelationIds) {
      supportInfo = ` [Support reference: ${supportCorrelationIds}]`;
    }

    // Enhanced error messages for common scenarios
    switch (errorDetails.statusCode) {
      case 400:
        userMessage = 'Bad request. Please check the request parameters and format.';
        break;
      case 401:
        userMessage = 'Authentication failed. Please re-authenticate.';
        break;
      case 403:
        userMessage = 'Insufficient permissions. Please check your app permissions in Azure AD.';
        break;
      case 404:
        userMessage = 'Resource not found. The requested item may have been deleted or moved.';
        break;
      case 409:
        userMessage = 'Conflict detected. This may be due to concurrent updates or scheduling conflicts.';
        break;
      case 429:
        // Return specific rate limit error
        const retryAfterMs = this.extractRetryAfter(error) || 60000;
        return createRateLimitError(Math.ceil(retryAfterMs / 1000));
      case 500:
      case 502:
      case 503:
      case 504:
        return createServiceUnavailableError('Microsoft Graph API');
      default:
        if (error.code === 'InvalidAuthenticationToken') {
          userMessage = 'Invalid or expired authentication token. Please re-authenticate.';
        } else {
          userMessage = `Graph API error: ${errorDetails.message || 'Unknown error occurred'}`;
        }
    }

    // Include specific error details if available
    if (errorDetails.innerError && errorDetails.innerError.message) {
      userMessage += ` Details: ${errorDetails.innerError.message}`;
    }

    // Append support information
    userMessage += supportInfo;

    // Create error with additional properties for MCP error conversion
    const finalError = new Error(userMessage);
    finalError.originalError = error;
    finalError.correlationIds = correlationIds;
    finalError.statusCode = errorDetails.statusCode;
    finalError.retryable = this.isRetryableError(errorDetails.statusCode);

    return convertErrorToToolError(finalError, 'Microsoft Graph API');
  }

  isRetryableError(statusCode) {
    // Define which errors are retryable
    return [401, 429, 500, 502, 503, 504].includes(statusCode);
  }

  // Get FolderResolver instance (lazy initialization)
  getFolderResolver() {
    if (!this.folderResolver) {
      this.folderResolver = new FolderResolver(this);
    }
    return this.folderResolver;
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