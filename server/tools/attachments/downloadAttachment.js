import { convertErrorToToolError, createValidationError } from '../../utils/mcpErrorResponse.js';
import { Buffer } from 'buffer';
import * as XLSX from 'xlsx';
import officeParser from 'officeparser';
import { handleLargeContent, saveBase64File } from '../../utils/fileOutput.js';

// Helper function to format file size
function formatFileSize(bytes) {
  if (!bytes || bytes === 0) return '0 Bytes';
  if (isNaN(bytes)) return 'Unknown size';
  
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// Helper function to determine if content should be treated as text
function isTextContent(contentType, filename, contentBytes = null) {
  console.error(`Debug: isTextContent check - contentType: "${contentType}", filename: "${filename}"`);
  
  const textTypes = [
    'text/',
    'application/json',
    'application/xml',
    'application/javascript',
    'application/typescript', 
    'application/x-python',
    'application/x-sh',
    'application/sql'
  ];
  
  const textExtensions = [
    '.txt', '.md', '.csv', '.log', '.ini', '.cfg', '.conf',
    '.html', '.htm', '.xml', '.json', '.js', '.ts', '.py',
    '.sh', '.bash', '.sql', '.css', '.scss', '.less',
    '.yaml', '.yml', '.toml', '.properties', '.env'
  ];
  
  // Check content type first
  if (contentType) {
    const lowerContentType = contentType.toLowerCase();
    if (textTypes.some(type => lowerContentType.startsWith(type))) {
      console.error(`Debug: Detected as text by contentType: ${contentType}`);
      return true;
    }
  }
  
  // Check file extension
  if (filename) {
    const ext = filename.toLowerCase().substring(filename.lastIndexOf('.'));
    if (textExtensions.includes(ext)) {
      console.error(`Debug: Detected as text by extension: ${ext}`);
      return true;
    }
  }
  
  // If contentType is null/empty, try to detect from content
  if ((!contentType || contentType.trim() === '') && contentBytes) {
    try {
      const sampleContent = Buffer.from(contentBytes, 'base64').toString('utf8', 0, 200);
      if (sampleContent.includes('<!DOCTYPE html>') || 
          sampleContent.includes('<html>') ||
          sampleContent.includes('<?xml') ||
          sampleContent.startsWith('{') ||
          sampleContent.startsWith('[')) {
        console.error(`Debug: Detected as text by content analysis`);
        return true;
      }
    } catch (error) {
      console.error(`Debug: Content analysis failed: ${error.message}`);
    }
  }
  
  console.error(`Debug: Detected as binary`);
  return false;
}

// Helper function to check if file is an Excel file
function isExcelFile(contentType, filename) {
  const excelMimeTypes = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
    'application/vnd.ms-excel', // .xls
    'application/vnd.openxmlformats-officedocument.spreadsheetml.template', // .xltx
    'application/vnd.ms-excel.sheet.macroEnabled.12', // .xlsm
    'application/vnd.ms-excel.template.macroEnabled.12', // .xltm
    'application/vnd.ms-excel.addin.macroEnabled.12', // .xlam
    'application/vnd.ms-excel.sheet.binary.macroEnabled.12' // .xlsb
  ];
  
  const excelExtensions = ['.xlsx', '.xls', '.xlsm', '.xltx', '.xltm', '.xlam', '.xlsb'];
  
  // Check content type
  if (contentType && excelMimeTypes.includes(contentType.toLowerCase())) {
    return true;
  }
  
  // Check file extension
  if (filename) {
    const ext = filename.toLowerCase().substring(filename.lastIndexOf('.'));
    return excelExtensions.includes(ext);
  }
  
  return false;
}

// Helper function to parse Excel files
function parseExcelContent(contentBytes, filename, maxSheets = 10, maxRowsPerSheet = 1000) {
  try {
    console.error(`Debug: Parsing Excel file: ${filename}`);
    
    // Decode Base64 to buffer
    const buffer = Buffer.from(contentBytes, 'base64');
    
    // Parse Excel file
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    
    const result = {
      type: 'excel',
      filename: filename,
      sheets: [],
      summary: {
        totalSheets: workbook.SheetNames.length,
        sheetNames: workbook.SheetNames
      }
    };
    
    // Process up to maxSheets sheets
    const sheetsToProcess = workbook.SheetNames.slice(0, maxSheets);
    
    for (const sheetName of sheetsToProcess) {
      console.error(`Debug: Processing sheet: ${sheetName}`);
      
      const worksheet = workbook.Sheets[sheetName];
      
      // Get sheet range
      const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
      const totalRows = range.e.r - range.s.r + 1;
      const totalCols = range.e.c - range.s.c + 1;
      
      // Limit rows to prevent overwhelming output
      const rowsToProcess = Math.min(totalRows, maxRowsPerSheet);
      
      // Convert to JSON with limited rows
      const jsonData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1, // Use array format instead of object
        range: rowsToProcess < totalRows ? `${worksheet['!ref'].split(':')[0]}:${XLSX.utils.encode_cell({r: range.s.r + rowsToProcess - 1, c: range.e.c})}` : undefined
      });
      
      const sheetInfo = {
        name: sheetName,
        dimensions: {
          rows: totalRows,
          columns: totalCols,
          range: worksheet['!ref'] || 'A1:A1'
        },
        data: jsonData,
        truncated: rowsToProcess < totalRows,
        displayedRows: jsonData.length,
        note: rowsToProcess < totalRows ? `Sheet truncated to ${maxRowsPerSheet} rows (total: ${totalRows})` : undefined
      };
      
      result.sheets.push(sheetInfo);
    }
    
    if (workbook.SheetNames.length > maxSheets) {
      result.summary.note = `Only first ${maxSheets} sheets displayed (total: ${workbook.SheetNames.length})`;
    }
    
    console.error(`Debug: Successfully parsed Excel file with ${result.sheets.length} sheets`);
    return result;
    
  } catch (error) {
    console.error(`Debug: Excel parsing failed: ${error.message}`);
    return {
      type: 'excel_error',
      error: `Failed to parse Excel file: ${error.message}`,
      note: 'File may be corrupted or in an unsupported Excel format'
    };
  }
}

// Helper function to check if file is an office document
function isOfficeDocument(contentType, filename) {
  const officeMimeTypes = [
    // PDF
    'application/pdf',
    // Word documents
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // .docx
    'application/msword', // .doc
    'application/vnd.openxmlformats-officedocument.wordprocessingml.template', // .dotx
    'application/vnd.ms-word.document.macroEnabled.12', // .docm
    'application/vnd.ms-word.template.macroEnabled.12', // .dotm
    // PowerPoint documents
    'application/vnd.openxmlformats-officedocument.presentationml.presentation', // .pptx
    'application/vnd.ms-powerpoint', // .ppt
    'application/vnd.openxmlformats-officedocument.presentationml.template', // .potx
    'application/vnd.openxmlformats-officedocument.presentationml.slideshow', // .ppsx
    'application/vnd.ms-powerpoint.addin.macroEnabled.12', // .ppam
    'application/vnd.ms-powerpoint.presentation.macroEnabled.12', // .pptm
    'application/vnd.ms-powerpoint.template.macroEnabled.12', // .potm
    'application/vnd.ms-powerpoint.slideshow.macroEnabled.12', // .ppsm
    // OpenDocument formats
    'application/vnd.oasis.opendocument.text', // .odt
    'application/vnd.oasis.opendocument.presentation', // .odp
    'application/vnd.oasis.opendocument.spreadsheet', // .ods
    // RTF
    'application/rtf',
    'text/rtf'
  ];
  
  const officeExtensions = [
    '.pdf',
    '.doc', '.docx', '.docm', '.dotx', '.dotm',
    '.ppt', '.pptx', '.pptm', '.potx', '.potm', '.ppsx', '.ppsm', '.ppam',
    '.odt', '.odp', '.ods',
    '.rtf'
  ];
  
  // Check content type
  if (contentType && officeMimeTypes.includes(contentType.toLowerCase())) {
    return true;
  }
  
  // Check file extension
  if (filename) {
    const ext = filename.toLowerCase().substring(filename.lastIndexOf('.'));
    return officeExtensions.includes(ext);
  }
  
  return false;
}

// Helper function to parse office documents using officeParser
function parseOfficeDocument(contentBytes, filename, maxTextLength = 50000) {
  try {
    console.error(`Debug: Parsing office document: ${filename}`);
    
    // Decode Base64 to buffer
    const buffer = Buffer.from(contentBytes, 'base64');
    
    // Parse office document using officeParser
    return new Promise((resolve) => {
      officeParser.parseOffice(buffer, (data, err) => {
        if (err) {
          console.error(`Debug: Office parsing failed: ${err}`);
          resolve({
            type: 'office_error',
            error: `Failed to parse office document: ${err}`,
            note: 'File may be corrupted, password-protected, or in an unsupported format'
          });
          return;
        }
        
        // Extract and process the text content
        const extractedText = data || '';
        const textLength = extractedText.length;
        const truncated = textLength > maxTextLength;
        const displayText = truncated ? extractedText.substring(0, maxTextLength) + '...' : extractedText;
        
        const result = {
          type: 'office_document',
          filename: filename,
          content: {
            text: displayText,
            extractedLength: textLength,
            truncated: truncated,
            truncatedLength: truncated ? maxTextLength : undefined,
            note: truncated ? `Text truncated to ${maxTextLength} characters (total: ${textLength})` : undefined
          },
          metadata: {
            originalSize: buffer.length,
            textLength: textLength,
            hasContent: textLength > 0
          }
        };
        
        console.error(`Debug: Successfully parsed office document with ${textLength} characters of text`);
        resolve(result);
      });
    });
    
  } catch (error) {
    console.error(`Debug: Office parsing failed: ${error.message}`);
    return Promise.resolve({
      type: 'office_error',
      error: `Failed to parse office document: ${error.message}`,
      note: 'File may be corrupted, password-protected, or in an unsupported format'
    });
  }
}

// Helper function to decode Base64 content appropriately
async function decodeAttachmentContent(contentBytes, contentType, filename, maxTextSize = 1024 * 1024) {
  try {
    const buffer = Buffer.from(contentBytes, 'base64');
    const decodedSize = buffer.length;
    
    console.error(`Debug: decodeAttachmentContent - size: ${decodedSize}, contentType: "${contentType}", filename: "${filename}"`);
    
    // For text content, decode to string if not too large
    if (isTextContent(contentType, filename, contentBytes)) {
      if (decodedSize <= maxTextSize) {
        const textContent = buffer.toString('utf8');
        return {
          type: 'text',
          content: textContent,
          size: decodedSize,
          encoding: 'utf8'
        };
      } else {
        return {
          type: 'text',
          content: `[Text file too large to display: ${formatFileSize(decodedSize)}]`,
          contentBytes: contentBytes, // Keep original for external processing
          size: decodedSize,
          encoding: 'base64_preserved',
          note: 'File exceeds display limit, use contentBytes for full content'
        };
      }
    } else if (isExcelFile(contentType, filename)) {
      // For Excel files, parse and extract data
      console.error(`Debug: Detected Excel file, attempting to parse`);
      const excelData = parseExcelContent(contentBytes, filename);
      
      return {
        type: 'excel',
        content: excelData,
        size: decodedSize,
        encoding: 'parsed',
        contentBytes: contentBytes, // Keep original for external processing
        note: 'Excel file parsed and data extracted. Use contentBytes for raw file access.'
      };
    } else if (isOfficeDocument(contentType, filename)) {
      // For office documents (PDF, Word, PowerPoint), parse and extract text
      console.error(`Debug: Detected office document, attempting to parse`);
      const officeData = await parseOfficeDocument(contentBytes, filename);
      
      return {
        type: 'office',
        content: officeData,
        size: decodedSize,
        encoding: 'parsed',
        contentBytes: contentBytes, // Keep original for external processing
        note: 'Office document parsed and text extracted. Use contentBytes for raw file access.'
      };
    } else {
      // For binary content, provide summary and keep Base64
      return {
        type: 'binary',
        content: `[Binary file: ${contentType || 'unknown type'}, ${formatFileSize(decodedSize)}]`,
        contentBytes: contentBytes,
        size: decodedSize,
        encoding: 'base64',
        note: 'Binary file preserved as Base64, decode with Buffer.from(contentBytes, "base64") if needed'
      };
    }
  } catch (error) {
    return {
      type: 'error',
      content: `[Failed to decode content: ${error.message}]`,
      contentBytes: contentBytes,
      encoding: 'base64_fallback',
      error: error.message
    };
  }
}

// Download attachment
export async function downloadAttachmentTool(authManager, args) {
  const { messageId, attachmentId, includeContent = false, decodeContent = true } = args;

  if (!messageId) {
    return createValidationError('messageId', 'Parameter is required');
  }

  if (!attachmentId) {
    return createValidationError('attachmentId', 'Parameter is required');
  }

  try {
    await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();

    console.error(`Debug: Downloading attachment ${attachmentId} from message ${messageId}`);

    // First, get attachment metadata and type
    const attachment = await graphApiClient.makeRequest(`/me/messages/${messageId}/attachments/${attachmentId}`, {
      select: 'id,name,contentType,size,isInline,lastModifiedDateTime,@odata.type'
    });

    console.error(`Debug: Attachment type: ${attachment['@odata.type']}, size: ${attachment.size}, contentType: "${attachment.contentType}"`);

    const attachmentInfo = {
      id: attachment.id,
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size || 0,
      sizeFormatted: formatFileSize(attachment.size),
      isInline: attachment.isInline || false,
      lastModifiedDateTime: attachment.lastModifiedDateTime,
      attachmentType: attachment['@odata.type']
    };

    if (includeContent) {
      try {
        console.error('Debug: Attempting to download attachment content...');
        
        // Try different approaches based on attachment type
        if (attachment['@odata.type'] === '#microsoft.graph.fileAttachment') {
          // Standard file attachment - request with contentBytes
          const fullAttachment = await graphApiClient.makeRequest(`/me/messages/${messageId}/attachments/${attachmentId}`, {
            select: 'id,name,contentType,size,isInline,lastModifiedDateTime,contentBytes,@odata.type'
          });
          
          if (fullAttachment.contentBytes) {
            if (decodeContent) {
              // Decode the Base64 content appropriately
              const decodedContent = await decodeAttachmentContent(
                fullAttachment.contentBytes, 
                attachment.contentType, 
                attachment.name
              );
              
              // Add decoded content info
              attachmentInfo.content = decodedContent.content;
              attachmentInfo.decodedContentType = decodedContent.type;
              attachmentInfo.encoding = decodedContent.encoding;
              attachmentInfo.contentIncluded = true;
              
              // Keep raw Base64 for binary files or when needed
              if (decodedContent.contentBytes) {
                attachmentInfo.contentBytes = decodedContent.contentBytes;
              }
              
              // Add any additional info
              if (decodedContent.note) {
                attachmentInfo.note = decodedContent.note;
              }
              
              if (decodedContent.error) {
                attachmentInfo.decodingError = decodedContent.error;
              }
              
              console.error(`Debug: Successfully downloaded and decoded content (type: ${decodedContent.type}, size: ${decodedContent.size} bytes)`);
            } else {
              // Return raw Base64 content
              attachmentInfo.contentBytes = fullAttachment.contentBytes;
              attachmentInfo.contentIncluded = true;
              attachmentInfo.encoding = 'base64';
              attachmentInfo.note = 'Raw Base64 content (set decodeContent: true to decode)';
              console.error(`Debug: Successfully downloaded raw content (${fullAttachment.contentBytes.length} Base64 characters)`);
            }
          } else {
            attachmentInfo.contentIncluded = false;
            attachmentInfo.contentError = 'No content bytes returned from API';
          }
          
        } else if (attachment['@odata.type'] === '#microsoft.graph.itemAttachment') {
          // Item attachment (embedded message/calendar item)
          const fullAttachment = await graphApiClient.makeRequest(`/me/messages/${messageId}/attachments/${attachmentId}`, {
            expand: 'item'
          });
          
          if (fullAttachment.item) {
            attachmentInfo.itemContent = fullAttachment.item;
            attachmentInfo.contentIncluded = true;
            attachmentInfo.encoding = 'json';
            console.error('Debug: Successfully downloaded item attachment content');
          } else {
            attachmentInfo.contentIncluded = false;
            attachmentInfo.contentError = 'No item content available for item attachment';
          }
          
        } else if (attachment['@odata.type'] === '#microsoft.graph.referenceAttachment') {
          // Reference attachment (link to SharePoint/OneDrive)
          const fullAttachment = await graphApiClient.makeRequest(`/me/messages/${messageId}/attachments/${attachmentId}`);
          
          attachmentInfo.sourceUrl = fullAttachment.sourceUrl;
          attachmentInfo.providerType = fullAttachment.providerType;
          attachmentInfo.thumbnailUrl = fullAttachment.thumbnailUrl;
          attachmentInfo.previewUrl = fullAttachment.previewUrl;
          attachmentInfo.permission = fullAttachment.permission;
          attachmentInfo.isFolder = fullAttachment.isFolder;
          attachmentInfo.contentIncluded = false;
          attachmentInfo.contentError = 'Reference attachment - use sourceUrl to access the linked resource';
          console.error('Debug: Reference attachment processed, sourceUrl:', fullAttachment.sourceUrl);
          
        } else {
          // Unknown attachment type - try the standard approach
          console.error('Debug: Unknown attachment type, trying standard contentBytes approach');
          const fullAttachment = await graphApiClient.makeRequest(`/me/messages/${messageId}/attachments/${attachmentId}`);
          
          if (fullAttachment.contentBytes) {
            if (decodeContent) {
              // Decode the Base64 content for unknown types too
              const decodedContent = await decodeAttachmentContent(
                fullAttachment.contentBytes, 
                attachment.contentType, 
                attachment.name
              );
              
              attachmentInfo.content = decodedContent.content;
              attachmentInfo.decodedContentType = decodedContent.type;
              attachmentInfo.encoding = decodedContent.encoding;
              attachmentInfo.contentIncluded = true;
              
              if (decodedContent.contentBytes) {
                attachmentInfo.contentBytes = decodedContent.contentBytes;
              }
              
              if (decodedContent.note) {
                attachmentInfo.note = decodedContent.note;
              }
            } else {
              // Return raw Base64 content
              attachmentInfo.contentBytes = fullAttachment.contentBytes;
              attachmentInfo.contentIncluded = true;
              attachmentInfo.encoding = 'base64';
            }
          } else {
            attachmentInfo.contentIncluded = false;
            attachmentInfo.contentError = `Unsupported attachment type: ${attachment['@odata.type']}`;
            // Include the full response for debugging
            attachmentInfo.debugInfo = {
              availableFields: Object.keys(fullAttachment),
              odataType: attachment['@odata.type']
            };
          }
        }
        
      } catch (contentError) {
        console.error('Debug: Error downloading attachment content:', contentError);
        attachmentInfo.contentIncluded = false;
        attachmentInfo.contentError = `Failed to download content: ${contentError.message}`;
        attachmentInfo.errorDetails = {
          statusCode: contentError.statusCode,
          code: contentError.code
        };
      }
    } else {
      attachmentInfo.contentIncluded = false;
      attachmentInfo.contentError = 'Content download not requested (set includeContent: true to download)';
    }

    // Handle large content by saving to file if needed
    const responseText = JSON.stringify(attachmentInfo, null, 2);
    const maxMcpResponseSize = 1048576; // 1MB MCP limit
    
    if (responseText.length > maxMcpResponseSize && attachmentInfo.contentBytes) {
      console.log(`Response size (${formatFileSize(responseText.length)}) exceeds MCP limit, saving to file...`);
      
      // Save the Base64 content to file
      const fileResult = await saveBase64File(
        attachmentInfo.contentBytes, 
        attachmentInfo.name, 
        attachmentInfo.contentType
      );
      
      if (fileResult.success) {
        // Replace contentBytes with file info
        const fileResponseInfo = {
          ...attachmentInfo,
          contentSavedToFile: true,
          fileOutput: fileResult,
          note: `Attachment content saved to file: ${fileResult.filePath}. Use the file path to access the full content.`,
          usage: {
            filePath: 'Use fileOutput.filePath to access the saved file',
            originalContent: 'Large content automatically saved due to MCP 1MB limit',
            decoding: attachmentInfo.encoding === 'parsed' ? 'Content was parsed before saving' : 'Raw file saved as downloaded'
          }
        };
        
        // Remove the large contentBytes from response
        delete fileResponseInfo.contentBytes;
        // Keep parsed content if it's small enough
        if (attachmentInfo.content && typeof attachmentInfo.content === 'object') {
          const contentSize = JSON.stringify(attachmentInfo.content).length;
          if (contentSize > maxMcpResponseSize / 2) { // If parsed content is also large
            delete fileResponseInfo.content;
            fileResponseInfo.parsedContentTruncated = true;
            fileResponseInfo.note += ' Parsed content also truncated due to size.';
          }
        }
        
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(fileResponseInfo, null, 2),
            },
          ],
        };
      } else {
        // Fall back to truncation if file saving failed
        const largeContentInfo = {
          ...attachmentInfo,
          contentTruncated: true,
          contentSize: responseText.length,
          contentSizeFormatted: formatFileSize(responseText.length),
          mcpLimitExceeded: true,
          fileSaveError: fileResult.error,
          note: `Response size (${formatFileSize(responseText.length)}) exceeds MCP limit. File save failed: ${fileResult.error}`,
          alternatives: {
            suggestion: 'Use decodeContent: true to get parsed text content instead of raw Base64',
            rawAccess: 'Content is available but too large to return in MCP response'
          }
        };
        
        delete largeContentInfo.contentBytes;
        delete largeContentInfo.content;
        
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(largeContentInfo, null, 2),
            },
          ],
        };
      }
    }

    return {
      content: [
        {
          type: 'text',
          text: responseText,
        },
      ],
    };
  } catch (error) {
    console.error('Debug: Error in downloadAttachmentTool:', error);
    return convertErrorToToolError(error, 'Failed to download attachment');
  }
}
