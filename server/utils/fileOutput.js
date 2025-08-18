/**
 * File Output Utility for MCP Large File Handling
 * 
 * This utility handles automatic file saving when content exceeds MCP's 1MB response limit.
 * Instead of failing or truncating, large files are saved to disk and file paths are returned.
 */

import fs from 'fs';
import path from 'path';
import { Buffer } from 'buffer';
import os from 'os';

/**
 * Get the work directory for file outputs
 * Uses MCP_OUTLOOK_WORK_DIR environment variable or falls back to system temp
 */
export function getWorkDirectory() {
  const workDir = process.env.MCP_OUTLOOK_WORK_DIR || os.tmpdir();
  
  // Ensure directory exists
  try {
    if (!fs.existsSync(workDir)) {
      fs.mkdirSync(workDir, { recursive: true });
    }
  } catch (error) {
    console.warn(`Warning: Could not create work directory ${workDir}, falling back to system temp:`, error.message);
    return os.tmpdir();
  }
  
  return workDir;
}

/**
 * Generate a unique filename with appropriate extension
 */
export function generateUniqueFilename(originalName, prefix = 'outlook') {
  const timestamp = Date.now();
  const random = Math.random().toString(36).substring(2, 8);
  
  if (originalName) {
    // Extract extension from original name
    const ext = path.extname(originalName);
    const base = path.basename(originalName, ext);
    return `${prefix}_${base}_${timestamp}_${random}${ext}`;
  }
  
  return `${prefix}_file_${timestamp}_${random}`;
}

/**
 * Save file content to disk and return file info
 */
export async function saveFileToDisc(content, filename, options = {}) {
  const {
    encoding = 'base64',
    prefix = 'outlook',
    mimeType = null,
    originalSize = null
  } = options;
  
  try {
    const workDir = getWorkDirectory();
    const uniqueFilename = generateUniqueFilename(filename, prefix);
    const filePath = path.join(workDir, uniqueFilename);
    
    // Write file based on encoding
    let buffer;
    if (encoding === 'base64') {
      buffer = Buffer.from(content, 'base64');
    } else if (encoding === 'utf8' || encoding === 'text') {
      buffer = Buffer.from(content, 'utf8');
    } else {
      // Assume it's already a buffer or binary content
      buffer = Buffer.isBuffer(content) ? content : Buffer.from(content);
    }
    
    await fs.promises.writeFile(filePath, buffer);
    
    // Get file stats
    const stats = await fs.promises.stat(filePath);
    
    const fileInfo = {
      success: true,
      filePath,
      filename: uniqueFilename,
      originalFilename: filename,
      size: stats.size,
      sizeFormatted: formatFileSize(stats.size),
      originalSize: originalSize || stats.size,
      mimeType,
      encoding,
      createdAt: new Date().toISOString(),
      workDirectory: workDir,
      note: `File saved due to MCP 1MB response limit. Use the file path to access the content.`
    };
    
    console.log(`File saved to disk: ${filePath} (${fileInfo.sizeFormatted})`);
    return fileInfo;
    
  } catch (error) {
    console.error('Failed to save file to disk:', error);
    return {
      success: false,
      error: `Failed to save file: ${error.message}`,
      filename,
      encoding
    };
  }
}

/**
 * Save Base64 content as a file
 */
export async function saveBase64File(base64Content, filename, mimeType = null) {
  return await saveFileToDisc(base64Content, filename, {
    encoding: 'base64',
    mimeType,
    originalSize: Math.round(base64Content.length * 0.75) // Approximate decoded size
  });
}

/**
 * Save text content as a file
 */
export async function saveTextFile(textContent, filename, encoding = 'utf8') {
  return await saveFileToDisc(textContent, filename, {
    encoding,
    mimeType: 'text/plain'
  });
}

/**
 * Check if content size would exceed MCP limits
 */
export function shouldSaveToFile(content, maxSize = 1048576) { // 1MB default
  let contentSize;
  
  if (typeof content === 'string') {
    contentSize = Buffer.byteLength(content, 'utf8');
  } else if (Buffer.isBuffer(content)) {
    contentSize = content.length;
  } else {
    // Estimate JSON size
    contentSize = Buffer.byteLength(JSON.stringify(content), 'utf8');
  }
  
  return contentSize > maxSize;
}

/**
 * Decide whether to return content inline or save to file
 */
export async function handleLargeContent(content, filename, options = {}) {
  const {
    maxSize = 1048576, // 1MB
    encoding = 'base64',
    mimeType = null,
    forceFile = false
  } = options;
  
  if (forceFile || shouldSaveToFile(content, maxSize)) {
    // Save to file and return file info
    const fileInfo = await saveFileToDisc(content, filename, { encoding, mimeType });
    return {
      savedToFile: true,
      ...fileInfo
    };
  } else {
    // Return content inline
    return {
      savedToFile: false,
      content,
      size: typeof content === 'string' ? Buffer.byteLength(content, 'utf8') : content.length,
      encoding
    };
  }
}

/**
 * Format file size for display
 */
function formatFileSize(bytes) {
  if (!bytes || bytes === 0) return '0 Bytes';
  if (isNaN(bytes)) return 'Unknown size';
  
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

/**
 * Clean up old files in work directory (optional maintenance function)
 */
export async function cleanupOldFiles(maxAge = 24 * 60 * 60 * 1000) { // 24 hours default
  try {
    const workDir = getWorkDirectory();
    const files = await fs.promises.readdir(workDir);
    const now = Date.now();
    let cleaned = 0;
    
    for (const file of files) {
      if (!file.startsWith('outlook_')) continue; // Only clean our files
      
      const filePath = path.join(workDir, file);
      try {
        const stats = await fs.promises.stat(filePath);
        if (now - stats.mtime.getTime() > maxAge) {
          await fs.promises.unlink(filePath);
          cleaned++;
        }
      } catch (error) {
        // File might have been deleted already, ignore
      }
    }
    
    if (cleaned > 0) {
      console.log(`Cleaned up ${cleaned} old files from work directory`);
    }
    
    return { cleaned, workDir };
  } catch (error) {
    console.warn('Failed to cleanup old files:', error.message);
    return { cleaned: 0, error: error.message };
  }
}

/**
 * Get configuration info for debugging
 */
export function getConfigInfo() {
  return {
    workDirectory: getWorkDirectory(),
    environmentVariable: 'MCP_OUTLOOK_WORK_DIR',
    currentValue: process.env.MCP_OUTLOOK_WORK_DIR || '(not set, using system temp)',
    maxResponseSize: '1MB (1048576 bytes)',
    supportedEncodings: ['base64', 'utf8', 'text', 'binary']
  };
}
