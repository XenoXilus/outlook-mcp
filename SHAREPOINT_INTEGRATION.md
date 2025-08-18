# SharePoint Integration with Outlook MCP

## Overview
Your Outlook MCP server now supports fetching SharePoint files using the **same authenticated session** as Outlook. This means you can seamlessly access SharePoint documents linked in emails without additional authentication.

## How It Works

### Shared Authentication Session
- The same Microsoft Graph API access token used for Outlook also works for SharePoint
- Added SharePoint-specific OAuth scopes: `Sites.Read.All`, `Sites.ReadWrite.All`, `Files.Read.All`, `Files.ReadWrite.All`
- No additional login required - if you're authenticated for Outlook, you can access SharePoint

### SharePoint URL Support
The system handles various SharePoint URL patterns:

1. **Sharing Links** (Recommended)
   - `https://company.sharepoint.com/:w:/s/sitename/EwABC123...`
   - `https://company-my.sharepoint.com/:b:/g/personal/user_company_com/EfDEF456...`
   - These work best with the Graph API `/shares` endpoint

2. **OneDrive for Business Links**
   - Personal workspace links are supported
   - Automatic detection of OneDrive vs SharePoint sites

3. **Direct Site URLs**
   - Currently limited support, will suggest using sharing links
   - Future enhancement could add full site navigation

## Available Tools

### 1. `outlook_get_sharepoint_file`
Fetch and optionally download SharePoint files.

**Parameters:**
- `sharePointUrl`: SharePoint sharing URL from email
- `fileId`: Direct file ID (alternative to URL)
- `driveId`: Drive ID for direct access
- `downloadContent`: Download file content as base64 (max 50MB)

**Example Usage:**
```json
{
  "sharePointUrl": "https://company.sharepoint.com/:w:/s/project/EwABC123...",
  "downloadContent": true
}
```

**Response:**
- File metadata (name, size, timestamps, etc.)
- Download URL for direct access
- Optional base64 content for small files
- Web URL for browser viewing

### 2. `outlook_list_sharepoint_files`
List files in SharePoint sites or OneDrive folders.

**Parameters:**
- `siteId`: SharePoint site ID (optional)
- `driveId`: Drive ID (defaults to user's OneDrive)
- `folderId`: Specific folder to browse
- `limit`: Max files to return (default 50)
- `orderBy`: Sort field (name, lastModifiedDateTime, etc.)

### 3. `outlook_resolve_sharepoint_link`
Resolve SharePoint links to get metadata without downloading.

**Parameters:**
- `sharePointUrl`: URL to resolve
- `includePermissions`: Include sharing permissions info

**Use Case:** Quickly check if a file exists and get basic info before downloading.

## Authentication Requirements

### Azure AD App Permissions
Your Azure AD application needs these permissions:
- **Mail.Read**, Mail.ReadWrite, Mail.Send (existing)
- **Sites.Read.All** - Read all SharePoint sites ✅ Added
- **Files.Read.All** - Read all files user can access ✅ Added
- **Sites.ReadWrite.All** - Read/write SharePoint sites ✅ Added  
- **Files.ReadWrite.All** - Read/write files ✅ Added

### Re-authentication Required
Since new scopes were added, existing users will need to re-authenticate once to grant the additional permissions.

## Integration Workflow

### Typical Email → SharePoint Workflow

1. **Receive Email with SharePoint Link**
   ```
   User gets email: "Please review the document: https://company.sharepoint.com/:w:/s/project/EwABC123..."
   ```

2. **Extract and Resolve Link**
   ```json
   {
     "tool": "outlook_resolve_sharepoint_link",
     "args": {
       "sharePointUrl": "https://company.sharepoint.com/:w:/s/project/EwABC123..."
     }
   }
   ```

3. **Download File if Needed**
   ```json
   {
     "tool": "outlook_get_sharepoint_file", 
     "args": {
       "fileId": "01ABC123DEF456789",
       "downloadContent": true
     }
   }
   ```

### Error Handling
- **Permission Errors**: User doesn't have access to the file/site
- **Invalid URLs**: Malformed or unsupported SharePoint URLs  
- **File Size Limits**: Files >50MB can't be downloaded inline
- **Network Issues**: Automatic retry with exponential backoff

### Security Features
- Same PKCE OAuth 2.0 flow as Outlook
- Encrypted token storage
- Respect SharePoint permissions (user can only access what they normally can)
- Audit logging of all file access attempts

## Implementation Details

### Graph API Endpoints Used
- `/shares/u!{encoded-url}/driveItem` - Resolve sharing links
- `/drives/{drive-id}/items/{item-id}` - Direct file access
- `/me/drive/root/children` - List OneDrive files
- `/sites/{site-id}/drives/{drive-id}` - Site-specific access

### File Content Handling
- **Small files (<50MB)**: Can download as base64 for immediate use
- **Large files**: Provide download URL for external handling
- **Folders**: List contents, not downloadable
- **Binary files**: Proper MIME type detection

### Rate Limiting
- Same rate limiting as other Graph API calls
- Respects SharePoint throttling limits
- Automatic retry with backoff

## Benefits

1. **Seamless Integration**: No additional authentication needed
2. **Email Context Preservation**: Access files directly from email context
3. **Comprehensive File Info**: Metadata, permissions, download options
4. **Security**: Respects existing SharePoint permissions
5. **Performance**: Efficient Graph API usage with caching

## Limitations & Future Enhancements

### Current Limitations
- Direct site navigation not fully implemented
- 50MB download limit for inline content
- Read-only focus (write operations available but not emphasized)

### Planned Enhancements
- Full site browsing and navigation
- Document preview/thumbnail support
- Collaborative editing status
- Version history access
- Advanced search across SharePoint sites

## Testing

The system is now ready for testing. Try:
1. Finding an email with a SharePoint link
2. Using `outlook_resolve_sharepoint_link` to test basic access
3. Using `outlook_get_sharepoint_file` to download content
4. Using `outlook_list_sharepoint_files` to browse OneDrive

Your existing authentication session will work seamlessly with SharePoint!
