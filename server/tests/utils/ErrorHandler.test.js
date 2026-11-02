import { describe, it, expect, beforeEach, vi } from 'vitest';
import { ErrorHandler, MCPError, GraphError, AuthError } from '../../utils/ErrorHandler.js';

describe('ErrorHandler', () => {
  let errorHandler;
  let mockLogger;

  beforeEach(() => {
    mockLogger = {
      error: vi.fn(),
      warn: vi.fn(),
      info: vi.fn(),
      debug: vi.fn()
    };
    errorHandler = new ErrorHandler(mockLogger);
  });

  describe('Error Classification', () => {
    it('should classify Graph API errors correctly', () => {
      const graphError = new Error('Graph API error');
      graphError.statusCode = 429;
      graphError.code = 'TooManyRequests';
      
      const classified = errorHandler.classifyError(graphError);
      
      expect(classified.type).toBe('graph');
      expect(classified.category).toBe('rate_limit');
      expect(classified.severity).toBe('medium');
      expect(classified.retryable).toBe(true);
    });

    it('should classify authentication errors correctly', () => {
      const authError = new Error('Invalid token');
      authError.statusCode = 401;
      authError.code = 'InvalidAuthenticationToken';
      
      const classified = errorHandler.classifyError(authError);
      
      expect(classified.type).toBe('auth');
      expect(classified.category).toBe('invalid_token');
      expect(classified.severity).toBe('high');
      expect(classified.retryable).toBe(true);
    });

    it('should classify MCP errors correctly', () => {
      const mcpError = new Error('Invalid tool parameters');
      mcpError.code = 'INVALID_PARAMS';
      
      const classified = errorHandler.classifyError(mcpError);
      
      expect(classified.type).toBe('mcp');
      expect(classified.category).toBe('validation');
      expect(classified.severity).toBe('low');
      expect(classified.retryable).toBe(false);
    });

    it('should classify unknown errors as generic', () => {
      const unknownError = new Error('Something went wrong');
      
      const classified = errorHandler.classifyError(unknownError);
      
      expect(classified.type).toBe('generic');
      expect(classified.category).toBe('unknown');
      expect(classified.severity).toBe('medium');
      expect(classified.retryable).toBe(false);
    });
  });

  describe('Error Handling', () => {
    it('should handle errors with proper logging', () => {
      const error = new Error('Test error');
      error.statusCode = 500;
      
      const result = errorHandler.handleError(error, 'test-operation');
      
      expect(mockLogger.error).toHaveBeenCalledWith(
        expect.stringContaining('Error in test-operation'),
        expect.objectContaining({
          message: 'Test error',
          statusCode: 500
        })
      );
      
      expect(result).toMatchObject({
        success: false,
        error: expect.objectContaining({
          message: expect.stringContaining('Microsoft service is temporarily unavailable')
        })
      });
    });

    it('should sanitize sensitive information', () => {
      const error = new Error('Authentication failed with token abc123');
      error.token = 'secret-token';
      error.password = 'secret-password';
      
      const result = errorHandler.handleError(error, 'auth-operation');
      
      expect(result.error.message).not.toContain('abc123');
      expect(result.error.message).toContain('[REDACTED]');
      expect(result.error.details).not.toHaveProperty('token');
      expect(result.error.details).not.toHaveProperty('password');
    });

    it('should include correlation IDs when available', () => {
      const error = new Error('Graph API error');
      error.correlationIds = {
        requestId: 'req-123',
        clientRequestId: 'client-456'
      };
      
      const result = errorHandler.handleError(error, 'graph-operation');
      
      expect(result.error.correlationIds).toEqual({
        requestId: 'req-123',
        clientRequestId: 'client-456'
      });
    });
  });

  describe('Graph failure details', () => {
    const graphError = {
      status: 400,
      code: 'BadRequest',
      message: "Parsing OData Select and Expand failed: Term '@odata.type' is not valid in a $select or $expand expression.",
      requestId: 'req-1111',
      clientRequestId: 'client-2222',
      date: '2026-01-15T10:00:00',
      method: 'GET',
      url: 'https://graph.microsoft.com/v1.0/me/messages/m1/attachments/a1?$select=id,name'
    };

    function rejected(status, overrides = {}) {
      const error = new Error(graphError.message);
      error.statusCode = status;
      error.code = graphError.code;
      error.graphError = { ...graphError, status, ...overrides };
      return error;
    }

    it('classifies a 400 as a non-retryable client error', () => {
      const classified = errorHandler.classifyError(rejected(400));
      expect(classified.category).toBe('client_error');
      expect(classified.retryable).toBe(false);
    });

    it.each([400, 403, 404, 409])('keeps status, code, message, request ids and METHOD url in the user message (%i)', (status) => {
      const result = errorHandler.handleError(rejected(status), 'graph-operation');
      const message = result.error.message;
      expect(message).toContain(`HTTP ${status}`);
      expect(message).toContain('BadRequest');
      expect(message).toContain(graphError.message);
      expect(message).toContain('request-id: req-1111');
      expect(message).toContain('client-request-id: client-2222');
      expect(message).toContain(`GET ${graphError.url}`);
      expect(message).not.toContain('Please check your parameters');
    });

    it('exposes the structured graphError in the safe error details', () => {
      const result = errorHandler.handleError(rejected(400), 'graph-operation');
      expect(result.error.details.graphError).toEqual({ ...graphError, status: 400 });
    });

    it('leaves the generic message alone when no structured Graph error is attached', () => {
      const error = new Error('Graph rejected it');
      error.statusCode = 404;
      const result = errorHandler.handleError(error, 'graph-operation');
      expect(result.error.message).toBe('Invalid request. Please check your parameters.');
      expect(result.error.details).not.toHaveProperty('graphError');
    });
  });

  describe('Retry Logic', () => {
    it('should determine if error is retryable', () => {
      const retryableError = new Error('Rate limited');
      retryableError.statusCode = 429;
      
      const nonRetryableError = new Error('Bad request');
      nonRetryableError.statusCode = 400;
      
      expect(errorHandler.isRetryableError(retryableError)).toBe(true);
      expect(errorHandler.isRetryableError(nonRetryableError)).toBe(false);
    });

    it('should calculate retry delay with exponential backoff', () => {
      const delay1 = errorHandler.calculateRetryDelay(1);
      const delay2 = errorHandler.calculateRetryDelay(2);
      const delay3 = errorHandler.calculateRetryDelay(3);
      
      expect(delay2).toBeGreaterThan(delay1);
      expect(delay3).toBeGreaterThan(delay2);
      expect(delay3).toBeLessThanOrEqual(30000); // Max delay
    });

    it('should respect Retry-After header', () => {
      const error = new Error('Rate limited');
      error.headers = { 'retry-after': '5' };
      
      const delay = errorHandler.calculateRetryDelay(1, error);
      expect(delay).toBe(5000); // 5 seconds in ms
    });
  });

  describe('Error Formatting', () => {
    it('should format errors for MCP responses', () => {
      const error = new Error('Test error');
      error.statusCode = 404;
      
      const formatted = errorHandler.formatForMCP(error);
      
      expect(formatted).toMatchObject({
        error: {
          code: 'RESOURCE_NOT_FOUND',
          message: expect.stringContaining('Invalid request')
        }
      });
    });

    it('should format errors for user display', () => {
      const error = new Error('Graph API error');
      error.statusCode = 403;
      error.code = 'Forbidden';
      
      const formatted = errorHandler.formatForUser(error);
      
      expect(formatted).toMatchObject({
        title: 'Invalid Request',
        message: expect.stringContaining('Invalid request'),
        severity: 'high',
        actionable: false
      });
    });
  });

  describe('Error Metrics', () => {
    it('should track error metrics', () => {
      const error1 = new Error('Error 1');
      error1.statusCode = 429;
      
      const error2 = new Error('Error 2');
      error2.statusCode = 500;
      
      errorHandler.handleError(error1, 'operation1');
      errorHandler.handleError(error2, 'operation2');
      
      const metrics = errorHandler.getErrorMetrics();
      
      expect(metrics.totalErrors).toBe(2);
      expect(metrics.errorsByType.graph).toBe(2);
      expect(metrics.errorsByCategory.rate_limit).toBe(1);
      expect(metrics.errorsByCategory.server_error).toBe(1);
    });

    it('should reset error metrics', () => {
      const error = new Error('Test error');
      errorHandler.handleError(error, 'operation');
      
      let metrics = errorHandler.getErrorMetrics();
      expect(metrics.totalErrors).toBe(1);
      
      errorHandler.resetMetrics();
      
      metrics = errorHandler.getErrorMetrics();
      expect(metrics.totalErrors).toBe(0);
    });
  });
});

describe('Custom Error Classes', () => {
  describe('MCPError', () => {
    it('should create MCP error with proper structure', () => {
      const error = new MCPError('Invalid parameters', 'INVALID_PARAMS', { param: 'value' });
      
      expect(error.message).toBe('Invalid parameters');
      expect(error.code).toBe('INVALID_PARAMS');
      expect(error.details).toEqual({ param: 'value' });
      expect(error.name).toBe('MCPError');
    });
  });

  describe('GraphError', () => {
    it('should create Graph error with correlation IDs', () => {
      const correlationIds = { requestId: 'req-123' };
      const error = new GraphError('Graph error', 429, 'TooManyRequests', correlationIds);
      
      expect(error.message).toBe('Graph error');
      expect(error.statusCode).toBe(429);
      expect(error.code).toBe('TooManyRequests');
      expect(error.correlationIds).toEqual(correlationIds);
      expect(error.name).toBe('GraphError');
    });
  });

  describe('AuthError', () => {
    it('should create Auth error with retry information', () => {
      const error = new AuthError('Token expired', 'TOKEN_EXPIRED', true);
      
      expect(error.message).toBe('Token expired');
      expect(error.code).toBe('TOKEN_EXPIRED');
      expect(error.retryable).toBe(true);
      expect(error.name).toBe('AuthError');
    });
  });
});