/**
 * **ultrathink** This error handler implements a comprehensive error management system 
 * for the Outlook MCP server. The complexity lies in:
 * 1. Multi-layered error classification (Graph API, MCP, Auth, Generic)
 * 2. Context-aware error formatting for different consumers
 * 3. Retry logic with exponential backoff and respect for Retry-After headers
 * 4. Sensitive data sanitization for security
 * 5. Metrics collection for monitoring and debugging
 * 
 * The design follows the principle of fail-fast with detailed diagnostics while
 * maintaining security by sanitizing sensitive information.
 */

import { formatGraphErrorText } from './mcpErrorResponse.js';

// Custom error classes for type-safe error handling
export class MCPError extends Error {
  constructor(message, code, details = {}) {
    super(message);
    this.name = 'MCPError';
    this.code = code;
    this.details = details;
  }
}

export class GraphError extends Error {
  constructor(message, statusCode, code, correlationIds = {}) {
    super(message);
    this.name = 'GraphError';
    this.statusCode = statusCode;
    this.code = code;
    this.correlationIds = correlationIds;
  }
}

export class AuthError extends Error {
  constructor(message, code, retryable = false) {
    super(message);
    this.name = 'AuthError';
    this.code = code;
    this.retryable = retryable;
  }
}

export class ErrorHandler {
  constructor(logger = console) {
    this.logger = logger;
    this.metrics = this.initializeMetrics();
    
    // Sensitive data patterns to redact
    this.sensitivePatterns = [
      /token[s]?\s*[=:]\s*[^\s]+/gi,
      /password[s]?\s*[=:]\s*[^\s]+/gi,
      /secret[s]?\s*[=:]\s*[^\s]+/gi,
      /key[s]?\s*[=:]\s*[^\s]+/gi,
      /bearer\s+[^\s]+/gi,
      /auth[a-z]*\s*[=:]\s*[^\s]+/gi,
      // Token in error messages
      /with\s+token\s+[^\s]+/gi,
      // Common token patterns
      /[a-zA-Z0-9]{20,}/g,
      // Email patterns in error messages
      /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g
    ];
  }

  initializeMetrics() {
    return {
      totalErrors: 0,
      errorsByType: {
        graph: 0,
        auth: 0,
        mcp: 0,
        generic: 0
      },
      errorsByCategory: {
        rate_limit: 0,
        server_error: 0,
        client_error: 0,
        auth_error: 0,
        validation: 0,
        network: 0,
        unknown: 0
      },
      errorsBySeverity: {
        low: 0,
        medium: 0,
        high: 0,
        critical: 0
      },
      retryableErrors: 0,
      lastErrorTime: null,
      errorFrequency: []
    };
  }

  /**
   * Classify error into type, category, severity, and retryability
   */
  classifyError(error) {
    const classification = {
      type: 'generic',
      category: 'unknown',
      severity: 'medium',
      retryable: false,
      statusCode: error.statusCode || error.status || null,
      code: error.code || null
    };

    // Authentication errors (check first to prioritize over graph errors)
    if (this.isAuthError(error)) {
      classification.type = 'auth';
      classification.category = 'auth_error';
      classification.severity = 'high';
      
      if (error.code === 'TokenExpired' || error.code === 'InvalidAuthenticationToken') {
        classification.category = 'invalid_token';
        classification.retryable = true;
      } else if (error.code === 'InsufficientPermissions') {
        classification.category = 'permissions';
        classification.retryable = false;
      } else {
        classification.retryable = error.retryable || false;
      }
    }
    // Graph API errors
    else if (this.isGraphError(error)) {
      classification.type = 'graph';
      classification.retryable = this.isRetryableStatusCode(error.statusCode);
      
      switch (error.statusCode) {
        case 429:
          classification.category = 'rate_limit';
          classification.severity = 'medium';
          classification.retryable = true;
          break;
        case 401:
          classification.category = 'auth_error';
          classification.severity = 'high';
          classification.retryable = true;
          break;
        case 400:
          classification.category = 'client_error';
          classification.severity = 'medium';
          classification.retryable = false;
          break;
        case 403:
          classification.category = 'client_error';
          classification.severity = 'high';
          classification.retryable = false;
          break;
        case 404:
          classification.category = 'client_error';
          classification.severity = 'low';
          classification.retryable = false;
          break;
        case 409:
          classification.category = 'client_error';
          classification.severity = 'medium';
          classification.retryable = false;
          break;
        case 500:
        case 502:
        case 503:
        case 504:
          classification.category = 'server_error';
          classification.severity = 'high';
          classification.retryable = true;
          break;
        default:
          classification.category = 'unknown';
      }
    }
    // MCP errors
    else if (this.isMCPError(error)) {
      classification.type = 'mcp';
      classification.category = 'validation';
      classification.severity = 'low';
      classification.retryable = false;
    }
    // Network errors
    else if (this.isNetworkError(error)) {
      classification.type = 'generic';
      classification.category = 'network';
      classification.severity = 'medium';
      classification.retryable = true;
    }

    return classification;
  }

  /**
   * Main error handling method
   */
  handleError(error, operation = 'unknown', context = {}) {
    const classification = this.classifyError(error);
    const errorId = this.generateErrorId();
    const timestamp = new Date().toISOString();
    
    // Update metrics
    this.updateMetrics(classification);
    
    // Sanitize error for logging
    const sanitizedError = this.sanitizeError(error);
    
    // Log error with appropriate level
    const logLevel = this.getLogLevel(classification.severity);
    this.logger[logLevel](`Error in ${operation} [${errorId}]`, {
      ...sanitizedError,
      classification,
      context,
      timestamp
    });
    
    // Format response
    const response = {
      success: false,
      error: {
        id: errorId,
        message: this.getUserFriendlyMessage(error, classification),
        code: this.getMCPErrorCode(classification),
        type: classification.type,
        severity: classification.severity,
        retryable: classification.retryable,
        timestamp,
        details: this.getSafeErrorDetails(error),
        correlationIds: error.correlationIds || null
      }
    };
    
    return response;
  }

  /**
   * Check if error is retryable
   */
  isRetryableError(error) {
    return this.classifyError(error).retryable;
  }

  /**
   * Calculate retry delay with exponential backoff
   */
  calculateRetryDelay(attempt, error = null) {
    // Check for Retry-After header
    const retryAfter = this.extractRetryAfter(error);
    if (retryAfter) {
      return retryAfter;
    }
    
    // Exponential backoff: 1s, 2s, 4s, 8s, 16s, 30s (max)
    const baseDelay = 1000;
    const maxDelay = 30000;
    const delay = Math.min(baseDelay * Math.pow(2, attempt - 1), maxDelay);
    
    // Add jitter to prevent thundering herd
    const jitter = Math.random() * 0.1 * delay;
    return Math.floor(delay + jitter);
  }

  /**
   * Format error for MCP response
   */
  formatForMCP(error) {
    const classification = this.classifyError(error);
    
    return {
      error: {
        code: this.getMCPErrorCode(classification),
        message: this.getUserFriendlyMessage(error, classification),
        data: {
          type: classification.type,
          severity: classification.severity,
          retryable: classification.retryable,
          correlationIds: error.correlationIds || null
        }
      }
    };
  }

  /**
   * Format error for user display
   */
  formatForUser(error) {
    const classification = this.classifyError(error);
    
    return {
      title: this.getErrorTitle(classification),
      message: this.getUserFriendlyMessage(error, classification),
      severity: classification.severity,
      actionable: classification.retryable,
      suggestions: this.getErrorSuggestions(classification)
    };
  }

  /**
   * Get error metrics
   */
  getErrorMetrics() {
    return {
      ...this.metrics,
      errorRate: this.calculateErrorRate(),
      topErrors: this.getTopErrors()
    };
  }

  /**
   * Reset metrics
   */
  resetMetrics() {
    this.metrics = this.initializeMetrics();
  }

  // Private helper methods

  isGraphError(error) {
    return error.name === 'GraphError' || 
           error.statusCode || 
           error.code === 'TooManyRequests' ||
           error.message?.includes('Graph');
  }

  isAuthError(error) {
    return error.name === 'AuthError' ||
           error.code === 'InvalidAuthenticationToken' ||
           error.code === 'TokenExpired' ||
           (error.statusCode === 401 && error.code === 'InvalidAuthenticationToken');
  }

  isMCPError(error) {
    return error.name === 'MCPError' ||
           error.code === 'INVALID_PARAMS' ||
           error.code === 'TOOL_ERROR';
  }

  isNetworkError(error) {
    return error.code === 'ECONNRESET' ||
           error.code === 'ECONNREFUSED' ||
           error.code === 'ETIMEDOUT' ||
           error.code === 'ENOTFOUND';
  }

  isRetryableStatusCode(statusCode) {
    return [401, 429, 500, 502, 503, 504].includes(statusCode);
  }

  sanitizeError(error) {
    const sanitized = {
      message: this.sanitizeMessage(error.message),
      stack: error.stack,
      name: error.name,
      code: error.code,
      statusCode: error.statusCode
    };
    
    // Remove sensitive properties
    const sensitiveProps = ['token', 'password', 'secret', 'key', 'authorization'];
    for (const prop of sensitiveProps) {
      if (error[prop]) {
        sanitized[prop] = '[REDACTED]';
      }
    }
    
    return sanitized;
  }

  sanitizeMessage(message) {
    if (!message) return message;
    
    let sanitized = message;
    for (const pattern of this.sensitivePatterns) {
      sanitized = sanitized.replace(pattern, '[REDACTED]');
    }
    
    return sanitized;
  }

  extractRetryAfter(error) {
    if (!error?.headers) return null;
    
    const retryAfter = error.headers['retry-after'] || error.headers['Retry-After'];
    if (!retryAfter) return null;
    
    const seconds = parseInt(retryAfter);
    return !isNaN(seconds) ? seconds * 1000 : null;
  }

  getUserFriendlyMessage(error, classification) {
    const baseMessage = this.sanitizeMessage(error.message);

    // A rejected Graph request keeps what Graph said (status, code, message,
    // request ids, method + URL). The structured copy holds no credentials,
    // and sanitising would mangle the ids it exists to preserve.
    if (error.graphError && (classification.category === 'client_error' || classification.category === 'unknown')) {
      const prefix = classification.category === 'client_error' ? 'Invalid request.' : 'Graph API error.';
      return `${prefix} ${formatGraphErrorText(error.graphError)}`;
    }
    
    switch (classification.category) {
      case 'rate_limit':
        return 'Request rate limit exceeded. Please try again in a moment.';
      case 'auth_error':
        return 'Authentication failed. Please re-authenticate.';
      case 'invalid_token':
        return 'Authentication token is invalid or expired. Please re-authenticate.';
      case 'permissions':
        return 'Insufficient permissions to perform this operation.';
      case 'server_error':
        return 'Microsoft service is temporarily unavailable. Please try again.';
      case 'client_error':
        return 'Invalid request. Please check your parameters.';
      case 'network':
        return 'Network connection failed. Please check your connection.';
      case 'validation':
        return baseMessage || 'Invalid input parameters.';
      default:
        return baseMessage || 'An unexpected error occurred.';
    }
  }

  getMCPErrorCode(classification) {
    // Check status code first for specific cases
    if (classification.statusCode === 404) {
      return 'RESOURCE_NOT_FOUND';
    }
    
    switch (classification.category) {
      case 'rate_limit':
        return 'RATE_LIMIT_EXCEEDED';
      case 'auth_error':
      case 'invalid_token':
        return 'AUTHENTICATION_FAILED';
      case 'permissions':
        return 'INSUFFICIENT_PERMISSIONS';
      case 'server_error':
        return 'SERVICE_UNAVAILABLE';
      case 'client_error':
        return 'INVALID_REQUEST';
      case 'network':
        return 'NETWORK_ERROR';
      case 'validation':
        return 'INVALID_PARAMS';
      default:
        return 'UNKNOWN_ERROR';
    }
  }

  getErrorTitle(classification) {
    switch (classification.category) {
      case 'rate_limit':
        return 'Rate Limit Exceeded';
      case 'auth_error':
      case 'invalid_token':
        return 'Authentication Failed';
      case 'permissions':
        return 'Permission Denied';
      case 'server_error':
        return 'Service Unavailable';
      case 'client_error':
        return 'Invalid Request';
      case 'network':
        return 'Network Error';
      case 'validation':
        return 'Validation Error';
      default:
        return 'Error';
    }
  }

  getErrorSuggestions(classification) {
    switch (classification.category) {
      case 'rate_limit':
        return ['Wait a moment before retrying', 'Reduce request frequency'];
      case 'auth_error':
      case 'invalid_token':
        return ['Re-authenticate your account', 'Check your credentials'];
      case 'permissions':
        return ['Contact your administrator', 'Check app permissions in Azure AD'];
      case 'server_error':
        return ['Try again later', 'Check Microsoft service status'];
      case 'client_error':
        return ['Check your input parameters', 'Verify the resource exists'];
      case 'network':
        return ['Check your internet connection', 'Try again in a moment'];
      case 'validation':
        return ['Check input parameters', 'Refer to API documentation'];
      default:
        return ['Try again', 'Contact support if the issue persists'];
    }
  }

  getLogLevel(severity) {
    switch (severity) {
      case 'critical':
        return 'error';
      case 'high':
        return 'error';
      case 'medium':
        return 'warn';
      case 'low':
        return 'info';
      default:
        return 'error';
    }
  }

  getSafeErrorDetails(error) {
    const details = {};
    
    if (error.statusCode) details.statusCode = error.statusCode;
    if (error.code) details.code = error.code;
    if (error.name) details.name = error.name;
    if (error.graphError) details.graphError = error.graphError;
    
    return details;
  }

  generateErrorId() {
    return `err_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  updateMetrics(classification) {
    this.metrics.totalErrors++;
    this.metrics.errorsByType[classification.type]++;
    this.metrics.errorsByCategory[classification.category]++;
    this.metrics.errorsBySeverity[classification.severity]++;
    
    if (classification.retryable) {
      this.metrics.retryableErrors++;
    }
    
    this.metrics.lastErrorTime = new Date().toISOString();
    this.metrics.errorFrequency.push(Date.now());
    
    // Keep only last 100 error timestamps
    if (this.metrics.errorFrequency.length > 100) {
      this.metrics.errorFrequency.shift();
    }
  }

  calculateErrorRate() {
    if (this.metrics.errorFrequency.length < 2) return 0;
    
    const now = Date.now();
    const fiveMinutesAgo = now - 5 * 60 * 1000;
    const recentErrors = this.metrics.errorFrequency.filter(timestamp => timestamp > fiveMinutesAgo);
    
    return recentErrors.length / 5; // errors per minute
  }

  getTopErrors() {
    const sorted = Object.entries(this.metrics.errorsByCategory)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5);
    
    return sorted.map(([category, count]) => ({ category, count }));
  }
}