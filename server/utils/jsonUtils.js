/**
 * JSON Utilities for safe serialization that prevents crashes
 * Handles circular references, undefined values, and other edge cases
 */

/**
 * Safely stringify an object, handling circular references and undefined values
 * @param {*} obj - Object to stringify
 * @param {number} space - Number of spaces for indentation (default: 2)
 * @returns {string} Safely stringified JSON
 */
export function safeStringify(obj, space = 2) {
  const seen = new WeakSet();
  
  try {
    return JSON.stringify(obj, (key, value) => {
      // Handle undefined values
      if (value === undefined) {
        return null;
      }
      
      // Handle circular references
      if (typeof value === 'object' && value !== null) {
        if (seen.has(value)) {
          return '[Circular Reference]';
        }
        seen.add(value);
      }
      
      // Handle functions (shouldn't be in API responses, but just in case)
      if (typeof value === 'function') {
        return '[Function]';
      }
      
      // Handle symbols
      if (typeof value === 'symbol') {
        return value.toString();
      }
      
      // Handle BigInt
      if (typeof value === 'bigint') {
        return value.toString();
      }
      
      return value;
    }, space);
  } catch (error) {
    console.error('Safe stringify failed, falling back to basic representation:', error);
    
    // Last resort - create a basic representation
    try {
      return JSON.stringify({
        error: 'Failed to serialize response',
        type: typeof obj,
        isArray: Array.isArray(obj),
        hasToString: obj && typeof obj.toString === 'function',
        basicInfo: obj && typeof obj === 'object' ? Object.keys(obj).slice(0, 10) : String(obj).slice(0, 200),
        serializationError: error.message
      }, null, space);
    } catch (fallbackError) {
      // Ultimate fallback
      return `{"error": "Complete serialization failure", "message": "${fallbackError.message}"}`;
    }
  }
}

/**
 * Clean an object to remove potentially problematic properties
 * @param {*} obj - Object to clean
 * @param {number} maxDepth - Maximum depth to traverse (default: 10)
 * @returns {*} Cleaned object
 */
export function cleanObject(obj, maxDepth = 10) {
  const seen = new WeakSet();
  
  function clean(value, depth = 0) {
    // Prevent infinite recursion
    if (depth > maxDepth) {
      return '[Max depth exceeded]';
    }
    
    // Handle null and undefined
    if (value === null || value === undefined) {
      return null;
    }
    
    // Handle primitive types
    if (typeof value !== 'object') {
      return value;
    }
    
    // Handle circular references
    if (seen.has(value)) {
      return '[Circular Reference]';
    }
    seen.add(value);
    
    // Handle arrays
    if (Array.isArray(value)) {
      return value.map(item => clean(item, depth + 1));
    }
    
    // Handle objects
    const cleaned = {};
    for (const [key, val] of Object.entries(value)) {
      // Skip certain problematic properties
      if (key.startsWith('_') || key === 'constructor' || key === '__proto__') {
        continue;
      }
      
      try {
        cleaned[key] = clean(val, depth + 1);
      } catch (error) {
        cleaned[key] = `[Error: ${error.message}]`;
      }
    }
    
    return cleaned;
  }
  
  try {
    return clean(obj);
  } catch (error) {
    console.error('Object cleaning failed:', error);
    return {
      error: 'Failed to clean object',
      message: error.message,
      type: typeof obj
    };
  }
}

/**
 * Create a safe MCP response with proper JSON serialization
 * @param {*} data - Data to include in the response
 * @param {Object} options - Options for serialization
 * @returns {Object} MCP-compliant response object
 */
export function createSafeResponse(data, options = {}) {
  const { space = 2, clean = true } = options;
  
  try {
    // Clean the data if requested
    const cleanedData = clean ? cleanObject(data) : data;
    
    // Safely stringify
    const jsonText = safeStringify(cleanedData, space);
    
    return {
      content: [
        {
          type: 'text',
          text: jsonText,
        },
      ],
    };
  } catch (error) {
    console.error('Failed to create safe response:', error);
    
    // Return error response
    return {
      content: [
        {
          type: 'text',
          text: safeStringify({
            error: 'Failed to serialize response data',
            message: error.message,
            timestamp: new Date().toISOString()
          }),
        },
      ],
      isError: true
    };
  }
}
