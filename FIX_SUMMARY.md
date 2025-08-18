# Fix Summary: Outlook MCP Tool - Folder-Specific Email Retrieval

## Problem Description
The `outlook_list_emails` and `outlook_search_emails` functions were failing when folder parameters were provided, throwing "Cannot read properties of undefined (reading 'map')" errors. The root cause was that folder names (like "BASB") were being used directly as folder IDs in Microsoft Graph API calls, but the API requires folder IDs (GUIDs/base64-encoded strings), not display names.

## Root Cause Analysis
1. **Microsoft Graph API Requirement**: The `/me/mailFolders/{id}/messages` endpoint requires folder IDs, not folder display names.
2. **Folder ID Format**: Microsoft Graph folder IDs are base64-encoded strings (e.g., `BBMkAGE1M2IyNGNmLWI3MjgtNDIwNS05YjB=`), not human-readable names.
3. **Missing Name-to-ID Resolution**: The existing code assumed folder parameters were already valid IDs.

## Solution Implemented

### 1. Created FolderResolver Utility (`server/graph/folderResolver.js`)
A new utility class that handles folder name-to-ID resolution:

**Features:**
- **Caching**: Stores folder information in memory with 5-minute expiry
- **Case-Insensitive Lookup**: Handles folder names in any case
- **Special Folder Support**: Handles "inbox" as a special Microsoft Graph folder name
- **ID Pass-through**: Recognizes when a folder ID is already provided
- **Multiple Folder Resolution**: Can resolve arrays of folder names/IDs
- **Error Handling**: Provides helpful error messages with available folder names

**Key Methods:**
- `resolveFolderToId(folderNameOrId)`: Converts folder name to ID
- `resolveFoldersToIds(folderNamesOrIds)`: Handles multiple folders
- `listAllFolders()`: Returns all available folders
- `getFolderInfo(folderNameOrId)`: Gets detailed folder information

### 2. Integrated FolderResolver into GraphApiClient (`server/graph/graphClient.js`)
- Added lazy initialization of FolderResolver instance
- Added `getFolderResolver()` method for accessing the resolver

### 3. Updated Email Tools to Use FolderResolver

#### Modified `listEmailsTool` (`server/tools/email/listEmails.js`)
- Added folder name-to-ID resolution before API calls
- Enhanced error handling for invalid folder names
- Added folder information to response output

#### Modified `searchEmailsTool` (`server/tools/email/searchEmails.js`)
- Added folder resolution for single and multiple folder searches
- Maintained backward compatibility with existing search logic
- Enhanced error handling and validation

## Key Benefits

1. **Backward Compatibility**: Existing code using folder IDs continues to work
2. **User-Friendly**: Users can now use readable folder names instead of cryptic IDs
3. **Robust Error Handling**: Clear error messages when folders don't exist
4. **Performance**: Caching reduces API calls for folder resolution
5. **Case Insensitive**: Works with folder names in any case

## Test Results
Created comprehensive unit tests that verify:
- ✅ Folder name resolution (BASB → `BBMkAGE1M2IyNGNmLWI3MjgtNDIwNS05YjB=`)
- ✅ Case-insensitive resolution (basb → same ID)
- ✅ Special inbox handling (inbox → "inbox")
- ✅ Folder ID pass-through (existing IDs work unchanged)
- ✅ Multiple folder resolution
- ✅ Error handling for non-existent folders
- ✅ Folder listing functionality

## Files Modified

1. **New File**: `server/graph/folderResolver.js` - Core folder resolution logic
2. **Modified**: `server/graph/graphClient.js` - Added FolderResolver integration
3. **Modified**: `server/tools/email/listEmails.js` - Added folder resolution
4. **Modified**: `server/tools/email/searchEmails.js` - Added folder resolution

## Example Usage

### Before Fix (Would Fail):
```javascript
// This would fail because "BASB" is not a valid folder ID
await listEmailsTool(authManager, { folder: "BASB" });
```

### After Fix (Works):
```javascript
// Now works - "BASB" gets resolved to actual folder ID
await listEmailsTool(authManager, { folder: "BASB" });
await listEmailsTool(authManager, { folder: "basb" }); // Case insensitive
await searchEmailsTool(authManager, { folders: ["BASB", "Inbox"] }); // Multiple folders
```

## Resolution Verification
The fix resolves the original issue:
- ❌ **Before**: `outlook_list_emails` with `folder="BASB"` → "Cannot read properties of undefined"
- ✅ **After**: `outlook_list_emails` with `folder="BASB"` → Returns emails from BASB folder
- ❌ **Before**: `outlook_search_emails` with `folders=["BASB"]` → "Cannot read properties of undefined"  
- ✅ **After**: `outlook_search_emails` with `folders=["BASB"]` → Returns emails from BASB folder

The MCP tool now properly handles folder-specific email retrieval using both folder names and folder IDs, providing a much better user experience.
