// Rate limiting utilities
import { convertErrorToToolError } from '../../utils/mcpErrorResponse.js';
import { createSafeResponse } from '../../utils/jsonUtils.js';

// Get rate limit metrics
export async function getRateLimitMetricsTool(authManager, args) {
  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    // Get rate limit info from the graph client
    const metrics = graphApiClient.getRateLimitMetrics ? graphApiClient.getRateLimitMetrics() : {
      requestCount: 0,
      throttleCount: 0,
      lastResetTime: new Date().toISOString(),
      currentWindow: {
        remainingRequests: 'unknown',
        resetTime: 'unknown'
      }
    };

    return createSafeResponse(metrics);
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to get rate limit metrics');
  }
}

// Reset rate limit metrics
export async function resetRateLimitMetricsTool(authManager, args) {
  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    // Reset rate limit metrics if the method exists
    if (graphApiClient.resetRateLimitMetrics) {
      graphApiClient.resetRateLimitMetrics();
    }

    return {
      content: [
        {
          type: 'text',
          text: 'Rate limit metrics reset successfully',
        },
      ],
    };
  } catch (error) {
    return convertErrorToToolError(error, 'Failed to reset rate limit metrics');
  }
}