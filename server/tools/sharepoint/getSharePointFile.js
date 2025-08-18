/**
 * SharePoint File Access Tool
 * 
 * Fetches files from SharePoint using the same authenticated session as Outlook.
 * Handles various SharePoint URL formats and sharing links.
 */

import { createValidationError, createToolError, convertErrorToToolError } from '../../utils/mcpErrorResponse.js';
import { Buffer } from 'buffer';
import { graphHelpers } from '../../graph/graphHelpers.js';
import * as XLSX from 'xlsx';
import officeParser from 'officeparser';
import { handleLargeContent, saveBase64File } from '../../utils/fileOutput.js';

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
async function decodeSharePointContent(contentBytes, contentType, filename, maxTextSize = 1024 * 1024) {
  try {
    const buffer = Buffer.from(contentBytes, 'base64');
    const decodedSize = buffer.length;
    
    console.error(`Debug: decodeSharePointContent - size: ${decodedSize}, contentType: "${contentType}", filename: "${filename}"`);
    
    // For text content, decode to string if not too large
    if (isTextContent(contentType, filename, contentBytes)) {
      if (decodedSize <= maxTextSize) {
        const textContent = buffer.toString('utf8');
        return {
          type: 'text',
          content: textContent,
          size: decodedSize,
          sizeFormatted: graphHelpers.general.formatFileSize(decodedSize),
          encoding: 'utf8'
        };
      } else {
        return {
          type: 'text',
          content: `[Text file too large to display: ${graphHelpers.general.formatFileSize(decodedSize)}]`,
          contentBytes: contentBytes, // Keep original for external processing
          size: decodedSize,
          sizeFormatted: graphHelpers.general.formatFileSize(decodedSize),
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
        sizeFormatted: graphHelpers.general.formatFileSize(decodedSize),
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
        sizeFormatted: graphHelpers.general.formatFileSize(decodedSize),
        encoding: 'parsed',
        contentBytes: contentBytes, // Keep original for external processing
        note: 'Office document parsed and text extracted. Use contentBytes for raw file access.'
      };
    } else {
      // For binary content, provide summary and keep Base64
      return {
        type: 'binary',
        content: `[Binary file: ${contentType || 'unknown type'}, ${graphHelpers.general.formatFileSize(decodedSize)}]`,
        contentBytes: contentBytes,
        size: decodedSize,
        sizeFormatted: graphHelpers.general.formatFileSize(decodedSize),
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

/**
 * Enhanced SharePoint URL parser with comprehensive pattern matching
 * @param {string} sharePointUrl - The SharePoint URL from the email
 * @returns {object} Parsed URL components
 */
function parseSharePointUrl(sharePointUrl) {
  try {
    console.error(`Debug: Parsing SharePoint URL: ${sharePointUrl}`);
    const url = new URL(sharePointUrl);
    const hostname = url.hostname.toLowerCase();
    const pathname = url.pathname;
    const searchParams = Object.fromEntries(url.searchParams);
    
    console.error(`Debug: URL components - hostname: ${hostname}, pathname: ${pathname}`);
    console.error(`Debug: Search params:`, searchParams);
    
    // Check if it's a SharePoint domain
    if (!hostname.includes('sharepoint.com')) {
      throw new Error(`Not a SharePoint URL: ${hostname}`);
    }
    
    // Pattern 1: SharePoint sharing links with format /:x:/r/personal/ or /:w:/r/sites/
    const sharingLinkPattern = /^\/(:[a-z]:)?\/([gr])\/(.+)$/i;
    const sharingMatch = pathname.match(sharingLinkPattern);
    
    if (sharingMatch) {
      const [, docType, accessType, resourcePath] = sharingMatch;
      console.error(`Debug: Detected sharing link - docType: ${docType}, accessType: ${accessType}, resourcePath: ${resourcePath}`);
      
      return {
        type: 'sharing_link',
        hostname,
        originalUrl: sharePointUrl,
        docType: docType || ':x:', // Default to Excel if not specified
        accessType, // 'r' for read, 'g' for guest
        resourcePath,
        searchParams,
        isPersonal: resourcePath.startsWith('personal/'),
        isSite: resourcePath.startsWith('sites/')
      };
    }
    
    // Pattern 2: Direct OneDrive for Business URLs
    if (pathname.includes('/personal/')) {
      const personalMatch = pathname.match(/\/personal\/([^\/]+)/);  
      if (personalMatch) {
        console.error(`Debug: Detected OneDrive personal folder: ${personalMatch[1]}`);
        return {
          type: 'onedrive_personal',
          hostname,
          userFolder: personalMatch[1],
          fullPath: pathname,
          searchParams
        };
      }
    }
    
    // Pattern 3: Team site URLs
    if (pathname.includes('/sites/')) {
      const siteMatch = pathname.match(/\/sites\/([^\/]+)/);
      if (siteMatch) {
        console.error(`Debug: Detected team site: ${siteMatch[1]}`);
        return {
          type: 'team_site',
          hostname,
          siteName: siteMatch[1],
          fullPath: pathname,
          searchParams
        };
      }
    }
    
    // Pattern 4: Check for any sharing parameters
    const hasShareParams = searchParams.d || searchParams.e || searchParams.share || searchParams.guestaccess;
    if (hasShareParams) {
      console.error(`Debug: Detected sharing parameters`);
      return {
        type: 'sharing_with_params',
        hostname,
        fullPath: pathname,
        searchParams,
        hasShareParams: true
      };
    }
    
    // Fallback: Generic SharePoint URL
    console.error(`Debug: Falling back to generic SharePoint URL`);
    return {
      type: 'generic_sharepoint',
      hostname,
      fullPath: pathname,
      searchParams
    };
    
  } catch (error) {
    console.error(`Debug: URL parsing failed: ${error.message}`);
    throw new Error(`Invalid SharePoint URL: ${error.message}`);
  }
}

/**
 * Enhanced sharing link resolver with multiple strategies
 * @param {object} graphClient - Authenticated Graph API client
 * @param {string} sharingUrl - SharePoint sharing URL
 * @param {object} urlInfo - Parsed URL information
 * @returns {object} Sharing information
 */
async function resolveSharedFile(graphClient, sharingUrl, urlInfo) {
  const strategies = [
    // Strategy 1: Direct Graph API shares endpoint
    async () => {
      console.error(`Debug: Trying Graph API shares endpoint`);
      const encodedUrl = Buffer.from(sharingUrl).toString('base64')
        .replace(/\+/g, '-')
        .replace(/\//g, '_')
        .replace(/=/g, '');
      
      return await graphClient.makeRequest(`/shares/u!${encodedUrl}/driveItem`, {
        select: 'id,name,size,createdDateTime,lastModifiedDateTime,webUrl,file,folder,@microsoft.graph.downloadUrl,parentReference'
      });
    },
    
    // Strategy 2: Try to extract drive and item info from URL structure
    async () => {
      console.error(`Debug: Trying URL structure parsing`);
      if ((urlInfo.type === 'sharing_link' || urlInfo.type === 'sharing_with_params') && urlInfo.searchParams.d) {
        // The 'd' parameter in SharePoint URLs is typically a shortened reference
        // Let's try different approaches to extract useful information
        const dParam = decodeURIComponent(urlInfo.searchParams.d);
        console.error(`Debug: Found 'd' parameter: ${dParam}`);
        
        // Strategy 2a: Try to use the 'd' parameter as a sharing token with Graph API
        try {
          // Some SharePoint URLs use the 'd' parameter as a sharing token
          const sharingTokenUrl = `${urlInfo.originalUrl.split('?')[0]}?d=${encodeURIComponent(dParam)}`;
          const encodedSharingUrl = Buffer.from(sharingTokenUrl).toString('base64')
            .replace(/\+/g, '-')
            .replace(/\//g, '_')
            .replace(/=/g, '');
          
          console.error(`Debug: Trying 'd' parameter as sharing token`);
          const result = await graphClient.makeRequest(`/shares/u!${encodedSharingUrl}/driveItem`, {
            select: 'id,name,size,createdDateTime,lastModifiedDateTime,webUrl,file,folder,@microsoft.graph.downloadUrl,parentReference'
          });
          
          if (result && result.id) {
            console.error(`Debug: Successfully resolved using 'd' parameter as sharing token`);
            return result;
          }
        } catch (sharingTokenError) {
          console.error(`Debug: 'd' parameter as sharing token failed: ${sharingTokenError.message}`);
        }
        
        // Strategy 2b: Try to extract file ID patterns from 'd' parameter
        const fileIdPatterns = [
          /([A-Z0-9]{20,})/gi,  // Long alphanumeric strings
          /([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})/gi, // GUIDs
          /([A-Z0-9_-]{15,})/gi // SharePoint-style IDs
        ];
        
        for (const pattern of fileIdPatterns) {
          const matches = dParam.match(pattern);
          if (matches) {
            for (const potentialFileId of matches) {
              console.error(`Debug: Trying potential file ID: ${potentialFileId}`);
              
              // Try different drive contexts
              const driveContexts = ['me', 'root'];
              
              for (const driveContext of driveContexts) {
                try {
                  const result = await graphClient.makeRequest(`/drives/${driveContext}/items/${potentialFileId}`, {
                    select: 'id,name,size,createdDateTime,lastModifiedDateTime,webUrl,file,folder,@microsoft.graph.downloadUrl,parentReference'
                  });
                  console.error(`Debug: Successfully found file with ID ${potentialFileId} in drive context: ${driveContext}`);
                  return result;
                } catch (driveError) {
                  console.error(`Debug: Drive context ${driveContext} with ID ${potentialFileId} failed: ${driveError.message}`);
                }
              }
            }
          }
        }
      }
      throw new Error('Could not extract usable file information from URL parameters');
    },
    
    // Strategy 3: Try SharePoint REST API endpoint construction
    async () => {
      console.error(`Debug: Trying SharePoint REST API approach`);
      if (urlInfo.type === 'sharing_link' || urlInfo.type === 'sharing_with_params') {
        // Extract site collection and construct direct API call
        const siteUrl = `https://${urlInfo.hostname}`;
        
        // Try to get the site information first
        const siteResponse = await graphClient.makeRequest(`/sites/${urlInfo.hostname}:/`, {
          select: 'id,displayName,webUrl'
        });
        
        if (siteResponse && siteResponse.id) {
          console.error(`Debug: Found site ID: ${siteResponse.id}`);
          // This would require additional parsing to get to the specific file
          // For now, return site info as fallback
          return {
            id: siteResponse.id,
            name: 'Site Root',
            size: 0,
            isFolder: true,
            webUrl: siteResponse.webUrl,
            note: 'Resolved to site root - specific file resolution needs additional implementation'
          };
        }
      }
      throw new Error('SharePoint REST API approach not applicable');
    }
  ];
  
  let lastError = null;
  
  // Try each strategy in sequence
  for (const [index, strategy] of strategies.entries()) {
    try {
      console.error(`Debug: Attempting resolution strategy ${index + 1}`);
      const result = await strategy();
      if (result && result.id) {
        console.error(`Debug: Strategy ${index + 1} succeeded`);
        return result;
      }
    } catch (error) {
      console.error(`Debug: Strategy ${index + 1} failed: ${error.message}`);
      lastError = error;
    }
  }
  
  // All strategies failed
  throw new Error(`Failed to resolve shared file after trying ${strategies.length} strategies. Last error: ${lastError?.message || 'Unknown error'}`);
}

/**
 * Get file content from SharePoint using Graph API
 * @param {object} graphClient - Authenticated Graph API client
 * @param {string} driveId - Drive ID (site, user, etc.)
 * @param {string} itemId - File item ID
 * @param {boolean} downloadContent - Whether to download file content
 * @returns {object} File information and optionally content
 */
async function getFileFromDrive(graphClient, driveId, itemId, downloadContent = false) {
  try {
    // Get file metadata
    const fileInfo = await graphClient.makeRequest(`/drives/${driveId}/items/${itemId}`, {
      select: 'id,name,size,createdDateTime,lastModifiedDateTime,webUrl,file,folder,@microsoft.graph.downloadUrl'
    });
    
    const result = {
      id: fileInfo.id,
      name: fileInfo.name,
      size: fileInfo.size,
      createdDateTime: fileInfo.createdDateTime,
      lastModifiedDateTime: fileInfo.lastModifiedDateTime,
      webUrl: fileInfo.webUrl,
      isFolder: !!fileInfo.folder,
      mimeType: fileInfo.file?.mimeType,
      downloadUrl: fileInfo['@microsoft.graph.downloadUrl']
    };
    
    // Download content if requested and file is not too large
    if (downloadContent && !fileInfo.folder) {
      const maxSize = 50 * 1024 * 1024; // 50MB limit
      
      if (fileInfo.size > maxSize) {
        result.contentError = `File too large to download inline (${Math.round(fileInfo.size / 1024 / 1024)}MB > 50MB). Use downloadUrl for direct download.`;
      } else {
        try {
          const contentResponse = await fetch(fileInfo['@microsoft.graph.downloadUrl']);
          if (contentResponse.ok) {
            const contentBuffer = await contentResponse.arrayBuffer();
            const contentBytes = Buffer.from(contentBuffer).toString('base64');
            const contentType = contentResponse.headers.get('content-type') || fileInfo.file?.mimeType;
            
            // Decode content intelligently based on type
            const decodedContent = await decodeSharePointContent(contentBytes, contentType, fileInfo.name);
            
            // Add decoded content info
            result.content = decodedContent.content;
            result.decodedContentType = decodedContent.type;
            result.encoding = decodedContent.encoding;
            result.contentType = contentType;
            result.contentSize = decodedContent.size;
            result.sizeFormatted = decodedContent.sizeFormatted;
            
            // Keep raw Base64 for binary files or when needed
            if (decodedContent.contentBytes) {
              result.contentBytes = decodedContent.contentBytes;
            }
            
            // Add any additional info
            if (decodedContent.note) {
              result.note = decodedContent.note;
            }
            
            if (decodedContent.error) {
              result.decodingError = decodedContent.error;
            }
          }
        } catch (downloadError) {
          result.contentError = `Failed to download content: ${downloadError.message}`;
        }
      }
    }
    
    return result;
  } catch (error) {
    console.error('Error getting file from drive:', error);
    throw new Error(`Failed to get file: ${error.message}`);
  }
}

/**
 * Main tool function to get SharePoint file
 * @param {object} authManager - Outlook authentication manager (same session)
 * @param {object} args - Tool arguments
 * @returns {object} MCP tool response
 */
export async function getSharePointFileTool(authManager, args) {
  try {
    // Input validation
    if (!args.sharePointUrl && !args.fileId) {
      return createValidationError('sharePointUrl or fileId', 'Either SharePoint URL or file ID is required');
    }
    
    // Ensure authentication
    const graphClient = await authManager.ensureAuthenticated();
    const graphApiClient = authManager.getGraphApiClient();
    
    let fileResult;
    
    if (args.sharePointUrl) {
      console.error(`Fetching SharePoint file from URL: ${args.sharePointUrl}`);
      
      // Parse the SharePoint URL
      const urlInfo = parseSharePointUrl(args.sharePointUrl);
      console.error('Parsed URL info:', JSON.stringify(urlInfo, null, 2));
      
      // Standard SharePoint sharing URLs from emails should be processed directly
      // They follow the format: /:x:/r/personal/ or /:w:/r/sites/ with d= and e= parameters
      if (urlInfo.type === 'sharing_link' || 
          urlInfo.type === 'sharing_with_params' || 
          urlInfo.hasShareParams ||
          // Check for standard sharing parameters from email links
          (urlInfo.searchParams && (urlInfo.searchParams.d || urlInfo.searchParams.e))) {
        
        console.error('Debug: Detected standard SharePoint sharing URL from email, attempting resolution');
        try {
          fileResult = await resolveSharedFile(graphApiClient, args.sharePointUrl, urlInfo);
          console.error(`Debug: Successfully resolved file: ${fileResult.name}`);
          
          // Handle content download if requested
          if (args.downloadContent && !fileResult.folder) {
            console.error('Debug: Content download requested for resolved sharing link');
            const maxSize = 50 * 1024 * 1024; // 50MB limit
            
            if (fileResult.size > maxSize) {
              fileResult.contentError = `File too large to download inline (${Math.round(fileResult.size / 1024 / 1024)}MB > 50MB). Use downloadUrl for direct download.`;
            } else {
              try {
                const downloadUrl = fileResult['@microsoft.graph.downloadUrl'];
                if (!downloadUrl) {
                  console.error('Debug: No download URL available, trying to fetch it');
                  // Try to get download URL using the file ID and parent reference
                  if (fileResult.id && fileResult.parentReference?.driveId) {
                    const freshFileInfo = await graphApiClient.makeRequest(`/drives/${fileResult.parentReference.driveId}/items/${fileResult.id}`, {
                      select: '@microsoft.graph.downloadUrl'
                    });
                    fileResult['@microsoft.graph.downloadUrl'] = freshFileInfo['@microsoft.graph.downloadUrl'];
                  } else {
                    throw new Error('No download URL available and insufficient metadata to fetch it');
                  }
                }
                
                const actualDownloadUrl = fileResult['@microsoft.graph.downloadUrl'];
                console.error(`Debug: Using download URL: ${actualDownloadUrl ? 'URL available' : 'URL missing'}`);
                
                if (actualDownloadUrl) {
                  const contentResponse = await fetch(actualDownloadUrl);
                  if (contentResponse.ok) {
                    const contentBuffer = await contentResponse.arrayBuffer();
                    const contentBytes = Buffer.from(contentBuffer).toString('base64');
                    const contentType = contentResponse.headers.get('content-type') || fileResult.mimeType;
                    
                    // Decode content intelligently based on type
                    const decodedContent = await decodeSharePointContent(contentBytes, contentType, fileResult.name);
                    
                    // Add decoded content info
                    fileResult.content = decodedContent.content;
                    fileResult.decodedContentType = decodedContent.type;
                    fileResult.encoding = decodedContent.encoding;
                    fileResult.contentType = contentType;
                    fileResult.contentSize = decodedContent.size;
                    fileResult.sizeFormatted = decodedContent.sizeFormatted;
                    
                    // Keep raw Base64 for binary files or when needed
                    if (decodedContent.contentBytes) {
                      fileResult.contentBytes = decodedContent.contentBytes;
                    }
                    
                    // Add any additional info
                    if (decodedContent.note) {
                      fileResult.note = decodedContent.note;
                    }
                    
                    if (decodedContent.error) {
                      fileResult.decodingError = decodedContent.error;
                    }
                    
                    console.error(`Debug: Successfully downloaded and decoded content (type: ${decodedContent.type}, size: ${decodedContent.size} bytes)`);
                  } else {
                    fileResult.contentError = `Failed to download content: HTTP ${contentResponse.status} ${contentResponse.statusText}`;
                  }
                } else {
                  fileResult.contentError = 'No download URL available for content download';
                }
              } catch (downloadError) {
                console.error(`Debug: Content download failed: ${downloadError.message}`);
                fileResult.contentError = `Failed to download content: ${downloadError.message}`;
              }
            }
          }
        } catch (resolveError) {
          console.error(`Debug: Resolution failed: ${resolveError.message}`);
          return createToolError(
            `Failed to resolve SharePoint sharing link: ${resolveError.message}`,
            'RESOLUTION_FAILED',
            {
              originalUrl: args.sharePointUrl,
              urlType: urlInfo.type,
              parsedInfo: urlInfo,
              resolutionAttempted: true,
              detailedError: resolveError.message,
              troubleshooting: {
                checkPermissions: 'Ensure you have access to the shared file',
                checkUrl: 'Verify the sharing link is valid and not expired',
                tryFileId: 'If you have the file ID, try using it directly'
              }
            }
          );
        }
        
      } else if (urlInfo.type === 'onedrive_personal' || urlInfo.type === 'team_site') {
        // For direct site URLs, try to provide more helpful guidance
        return createToolError(
          `Direct ${urlInfo.type} URLs require file-specific sharing links. Please use a sharing link to the specific file.`,
          'DIRECT_SITE_URL_UNSUPPORTED',
          {
            suggestion: 'Right-click the file in SharePoint/OneDrive and select "Copy link" to get a sharing link',
            urlType: urlInfo.type,
            parsedInfo: urlInfo,
            supportedFormats: [
              'https://company.sharepoint.com/:w:/r/sites/...',
              'https://company.sharepoint.com/:x:/g/personal/...',
              'https://company-my.sharepoint.com/:b:/personal/...'
            ]
          }
        );
        
      } else {
        // Generic SharePoint URL
        return createToolError(
          `Unsupported SharePoint URL format. Please use a file sharing link.`,
          'UNSUPPORTED_URL_FORMAT',
          {
            suggestion: 'Generate a sharing link from SharePoint by right-clicking the file and selecting "Copy link"',
            urlType: urlInfo.type,
            parsedInfo: urlInfo,
            supportedFormats: [
              'File sharing links: https://company.sharepoint.com/:w:/r/...',
              'OneDrive sharing links: https://company-my.sharepoint.com/:x:/g/...'
            ]
          }
        );
      }
    } else if (args.fileId) {
      // Direct file access using Graph API
      const driveId = args.driveId || 'me'; // Default to user's OneDrive
      fileResult = await getFileFromDrive(graphApiClient, driveId, args.fileId, args.downloadContent);
    }
    
    const response = {
      success: true,
      file: fileResult,
      message: `Successfully retrieved ${fileResult.isFolder ? 'folder' : 'file'}: ${fileResult.name}`,
      usage: {
        downloadUrl: 'Use the downloadUrl for direct file download',
        content: fileResult.content ? 'File content included' : 'Content not downloaded (use downloadContent: true to include)',
        webUrl: 'Use webUrl to view file in SharePoint/OneDrive'
      }
    };
    
    // Handle large content by saving to file if necessary
    const finalResponse = await handleLargeContent(response, ['file.contentBytes', 'file.content'], {
      filenameSuffix: fileResult.name ? `_${fileResult.name}` : '_sharepoint_file',
      contextInfo: {
        toolName: 'sharepoint_get_file',
        fileName: fileResult.name,
        fileSize: fileResult.size,
        originalUrl: args.sharePointUrl || 'Direct file access'
      }
    });
    
    if (finalResponse.savedToFile) {
      // Add helpful context when content was saved to file
      finalResponse.file = {
        ...finalResponse.file,
        contentAccessInfo: {
          savedToFile: true,
          reason: 'File content exceeded MCP response size limit (1MB)',
          alternatives: {
            localFile: 'Content saved to local file (see savedFiles)',
            downloadUrl: finalResponse.file.downloadUrl || 'Use downloadUrl for direct download',
            webUrl: finalResponse.file.webUrl || 'Use webUrl to view in SharePoint/OneDrive'
          }
        }
      };
    }

    return {
      content: [{
        type: 'text',
        text: JSON.stringify(finalResponse, null, 2)
      }]
    };
    
  } catch (error) {
    console.error('SharePoint file access error:', error);
    
    if (error.isError) {
      return error; // Already an MCP error
    }
    
    return convertErrorToToolError(error, 'SharePoint file access failed');
  }
}

/**
 * Tool to resolve SharePoint sharing links without downloading
 * @param {object} authManager - Authentication manager
 * @param {object} args - Tool arguments
 * @returns {object} MCP tool response
 */
export async function resolveSharePointLinkTool(authManager, args) {
  try {
    if (!args.sharePointUrl) {
      return createValidationError('sharePointUrl', 'SharePoint URL is required');
    }
    
    const graphApiClient = authManager.getGraphApiClient();
    
    // Parse the URL first
    const urlInfo = parseSharePointUrl(args.sharePointUrl);
    console.error('Resolve SharePoint link - Parsed URL info:', JSON.stringify(urlInfo, null, 2));
    
    // Resolve the sharing link to get metadata only
    const fileInfo = await resolveSharedFile(graphApiClient, args.sharePointUrl, urlInfo);
    
    const result = {
      id: fileInfo.id,
      name: fileInfo.name,
      size: fileInfo.size,
      type: fileInfo.folder ? 'folder' : 'file',
      mimeType: fileInfo.file?.mimeType,
      createdDateTime: fileInfo.createdDateTime,
      lastModifiedDateTime: fileInfo.lastModifiedDateTime,
      webUrl: fileInfo.webUrl,
      downloadUrl: fileInfo['@microsoft.graph.downloadUrl'],
      sharing: {
        originalUrl: args.sharePointUrl,
        resolved: true,
        accessible: true
      }
    };
    
    // Add permissions info if requested
    if (args.includePermissions) {
      try {
        const permissions = await graphApiClient.makeRequest(`/drives/items/${fileInfo.id}/permissions`);
        result.sharing.permissions = permissions.value || [];
      } catch (permError) {
        result.sharing.permissionsError = 'Could not retrieve permissions information';
      }
    }
    
    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          success: true,
          file: result,
          message: `Successfully resolved ${result.type}: ${result.name}`,
          usage: {
            downloadUrl: 'Use downloadUrl for direct download without re-authentication',
            webUrl: 'Use webUrl to view in browser',
            fileId: 'Use id with outlook_get_sharepoint_file for content download'
          }
        }, null, 2)
      }]
    };
    
  } catch (error) {
    console.error('SharePoint link resolution error:', error);
    return convertErrorToToolError(error, 'SharePoint link resolution failed');
  }
}

/**
 * Tool to list files in a SharePoint site or folder
 * @param {object} authManager - Authentication manager
 * @param {object} args - Tool arguments
 * @returns {object} MCP tool response
 */
export async function listSharePointFilesTool(authManager, args) {
  try {
    const graphApiClient = authManager.getGraphApiClient();
    
    let listPath;
    if (args.siteId && args.driveId) {
      listPath = `/sites/${args.siteId}/drives/${args.driveId}/root/children`;
    } else if (args.driveId) {
      listPath = `/drives/${args.driveId}/root/children`;
    } else {
      listPath = '/me/drive/root/children'; // Default to user's OneDrive
    }
    
    if (args.folderId) {
      listPath = `/drives/${args.driveId || 'me'}/items/${args.folderId}/children`;
    }
    
    const response = await graphApiClient.makeRequest(listPath, {
      select: 'id,name,size,createdDateTime,lastModifiedDateTime,webUrl,file,folder',
      top: args.limit || 50,
      orderby: args.orderBy || 'name'
    });
    
    const files = (response.value || []).map(item => ({
      id: item.id,
      name: item.name,
      size: item.size,
      type: item.folder ? 'folder' : 'file',
      mimeType: item.file?.mimeType,
      createdDateTime: item.createdDateTime,
      lastModifiedDateTime: item.lastModifiedDateTime,
      webUrl: item.webUrl
    }));
    
    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          success: true,
          files: files,
          count: files.length,
          message: `Found ${files.length} items`
        }, null, 2)
      }]
    };
    
  } catch (error) {
    console.error('SharePoint list files error:', error);
    return convertErrorToToolError(error, 'SharePoint file listing failed');
  }
}
