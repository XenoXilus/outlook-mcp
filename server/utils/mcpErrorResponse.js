/**
 * MCP Error Response Utilities
 * 
 * Helper functions to create standardized MCP error responses that comply
 * with the JSON-RPC 2.0 specification and MCP protocol requirements.
 */

/**
 * Creates a tool execution error response in MCP format
 * @param {string} message - The error message to display
 * @param {Object} details - Additional error details (optional)
 * @returns {Object} MCP-compliant tool error response
 */
export function createToolError(message, details = {}) {
  return {
    content: [
      {
        type: 'text',
        text: message
      }
    ],
    isError: true,
    ...(Object.keys(details).length > 0 && { _errorDetails: details })
  };
}

/**
 * Creates a protocol-level error response in JSON-RPC 2.0 format
 * @param {number} code - JSON-RPC error code (e.g., -32601 for MethodNotFound)
 * @param {string} message - Short error description
 * @param {*} data - Additional error data (optional)
 * @returns {Object} JSON-RPC 2.0 error response
 */
export function createProtocolError(code, message, data = null) {
  const error = {
    jsonrpc: "2.0",
    error: {
      code,
      message
    }
  };

  if (data !== null) {
    error.error.data = data;
  }

  return error;
}

/**
 * Standard JSON-RPC error codes used by MCP
 */
export const ErrorCodes = {
  PARSE_ERROR: -32700,
  INVALID_REQUEST: -32600,
  METHOD_NOT_FOUND: -32601,
  INVALID_PARAMS: -32602,
  INTERNAL_ERROR: -32603,
  CONNECTION_CLOSED: -32000,
  REQUEST_TIMEOUT: -32001
};

/**
 * Converts a standard Error object to MCP tool error format
 * @param {Error} error - The error object to convert
 * @param {string} context - Context about where the error occurred
 * @returns {Object} MCP-compliant tool error response
 */
export function convertErrorToToolError(error, context = '') {
  // If it's already an MCP error, return it as-is
  if (error && error.isError !== undefined) {
    return error;
  }

  // Handle cases where error might not be an Error object
  const errorMessage = error?.message || error?.toString() || 'Unknown error';

  const message = context
    ? `${context}: ${errorMessage}`
    : errorMessage;

  const details = {
    name: error?.name || 'Error',
    stack: error?.stack?.split('\n')[0] // First line only for brevity
  };

  // Preserve common error properties
  if (error?.statusCode) details.statusCode = error.statusCode;
  if (error?.code) details.code = error.code;
  if (error?.correlationId) details.correlationId = error.correlationId;
  if (error?.correlationIds) details.correlationIds = error.correlationIds;
  if (error?.retryAfter) details.retryAfter = error.retryAfter;
  // Structured Graph failure (status/code/message/request ids/method/url) so
  // callers can propagate it without re-parsing the message text.
  const graphError = error?.graphError || error?.fetchDetails?.graphError;
  if (graphError) details.graphError = graphError;
  if (error?.fetchDetails) details.fetchDetails = error.fetchDetails;

  return createToolError(message, details);
}

/**
 * Renders a structured Graph failure as one line of verbatim diagnostics:
 * HTTP status, Graph error code and message, the request ids Graph needs to
 * trace the call, and the method + full URL (query string included) as sent.
 * @param {Object} graphError - { status, code, message, requestId, clientRequestId, date, method, url }
 * @returns {string}
 */
export function formatGraphErrorText(graphError) {
  const g = graphError || {};
  const status = g.status ? `HTTP ${g.status}` : 'No HTTP response';
  const code = g.code ? ` ${g.code}` : '';
  const message = g.message || 'No error message returned';
  const ids = [];
  if (g.requestId) ids.push(`request-id: ${g.requestId}`);
  if (g.clientRequestId) ids.push(`client-request-id: ${g.clientRequestId}`);
  if (g.date) ids.push(`date: ${g.date}`);
  const idPart = ids.length > 0 ? ` [${ids.join('; ')}]` : '';
  const request = g.url ? ` ${g.method ? `${g.method} ` : ''}${g.url}` : '';
  return `${status}${code}: ${message}${idPart}${request}`;
}

/**
 * Creates a validation error response for invalid parameters
 * @param {string} paramName - Name of the invalid parameter
 * @param {string} reason - Why the parameter is invalid
 * @returns {Object} MCP-compliant tool error response
 */
export function createValidationError(paramName, reason) {
  return createToolError(`Invalid parameter '${paramName}': ${reason}`, {
    type: 'validation',
    parameter: paramName
  });
}

/**
 * Creates an authentication error response
 * @param {string} message - Authentication error message
 * @param {boolean} retryable - Whether the error is retryable
 * @returns {Object} MCP-compliant tool error response
 */
export function createAuthError(message, retryable = false) {
  return createToolError(`Authentication failed: ${message}`, {
    type: 'authentication',
    retryable,
    suggestion: retryable
      ? 'Please re-authenticate your account'
      : 'Check your credentials and try again'
  });
}

/**
 * Creates a rate limit error response
 * @param {number} retryAfter - Seconds to wait before retrying
 * @returns {Object} MCP-compliant tool error response
 */
export function createRateLimitError(retryAfter) {
  return createToolError(
    `Rate limit exceeded. Please wait ${retryAfter} seconds before retrying.`,
    {
      type: 'rate_limit',
      retryAfter,
      retryable: true
    }
  );
}

/**
 * Creates a service unavailable error response
 * @param {string} service - Name of the unavailable service
 * @returns {Object} MCP-compliant tool error response
 */
export function createServiceUnavailableError(service = 'Microsoft Graph API') {
  return createToolError(
    `${service} is temporarily unavailable. Please try again later.`,
    {
      type: 'service_unavailable',
      retryable: true,
      suggestion: 'Check Microsoft service status or try again in a few minutes'
    }
  );
}