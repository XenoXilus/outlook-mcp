# Microsoft Outlook MCP Server

A Model Context Protocol (MCP) server that enables AI assistants to interact with Microsoft Outlook email, calendar, contacts, and tasks through the Microsoft Graph API.

## Features

- **Email Operations**: Read, search, send, reply to emails and download attachments
- **SharePoint Integration**: Access SharePoint files via sharing links or direct file IDs
- **Calendar Management**: View and manage calendar events and appointments
- **Contact Management**: Access and manage Outlook contacts
- **Task Management**: Create and manage Outlook tasks
- **Office Document Processing**: Parse PDF, Word, PowerPoint, and Excel files with extracted text content
- **Large File Support**: Automatic handling of files that exceed MCP response size limits

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd outlook-mcp
```

2. Install dependencies:
```bash
npm install
```

3. Set up Azure AD Application:
   - Register an application in Azure AD
   - Configure required Microsoft Graph API permissions
   - Note the Application (client) ID for configuration

## Configuration

### Environment Variables

The server supports the following environment variables:

- `AZURE_CLIENT_ID`: Your Azure AD application client ID (required)
- `MCP_OUTLOOK_WORK_DIR`: Directory for saving large files that exceed MCP response limits (optional)

#### Large File Handling

When downloading large attachments or SharePoint files, the server automatically detects when the response would exceed the MCP 1MB limit and saves the content to local files instead.

**MCP_OUTLOOK_WORK_DIR Configuration:**
- If set, large files are saved to this directory
- If not set, files are saved to the system temp directory
- Files are automatically named with timestamps to avoid conflicts
- Old files are periodically cleaned up to manage disk space

Example configuration:
```bash
export MCP_OUTLOOK_WORK_DIR="/path/to/your/work/directory"
export AZURE_CLIENT_ID="your-azure-app-id"
```

### Claude Desktop Integration

Add to your Claude Desktop MCP settings:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["/path/to/outlook-mcp/server/index.js"],
      "env": {
        "AZURE_CLIENT_ID": "your-azure-app-id",
        "MCP_OUTLOOK_WORK_DIR": "/path/to/your/work/directory"
      }
    }
  }
}
```

## Usage

### Email Operations

```javascript
// Search for emails
outlook_search_emails({
  query: "from:user@example.com",
  maxResults: 10,
  folder: "inbox"
})

// Download email attachment
outlook_download_attachment({
  messageId: "message-id",
  attachmentId: "attachment-id",
  decodeContent: true  // Automatically parse office documents
})
```

### SharePoint File Access

```javascript
// Access SharePoint file via sharing link
outlook_get_sharepoint_file({
  sharePointUrl: "https://company.sharepoint.com/:w:/r/sites/...",
  downloadContent: true
})

// Direct file access
outlook_get_sharepoint_file({
  fileId: "file-id",
  driveId: "drive-id",
  downloadContent: true
})
```

### Office Document Processing

The server automatically detects and processes:
- **PDF files**: Extract text content
- **Word documents** (.doc, .docx): Extract text content
- **PowerPoint presentations** (.ppt, .pptx): Extract text content
- **Excel spreadsheets** (.xls, .xlsx): Parse and extract data in structured format

## Authentication

The server uses OAuth 2.0 with PKCE for secure authentication:

1. First run will open a browser for Microsoft authentication
2. Tokens are securely stored using the OS keychain (keytar)
3. Automatic token refresh for long-term usage
4. No sensitive data stored in plain text

## Required Permissions

Your Azure AD application needs these Microsoft Graph API permissions:

- `Mail.ReadWrite`: Email access
- `Calendars.ReadWrite`: Calendar access
- `Contacts.ReadWrite`: Contact access
- `Tasks.ReadWrite`: Task management
- `Files.Read.All`: SharePoint file access
- `Sites.Read.All`: SharePoint site access

## File Size Limits and Handling

### MCP Response Size Limit
- MCP responses are limited to 1MB
- Large files automatically trigger file output mode
- Content is saved locally and file paths are returned instead

### File Output Behavior
1. **Automatic Detection**: Server checks if response size would exceed 1MB
2. **Local Saving**: Large content is saved to `MCP_OUTLOOK_WORK_DIR` or system temp
3. **Metadata Response**: Returns file metadata with local file paths
4. **Alternative Access**: Provides download URLs and web URLs for direct access

### Supported File Types for Parsing
- **Text Files**: .txt, .md, .csv, .log, .json, .xml, .html, .js, .py, etc.
- **Office Documents**: .pdf, .doc/.docx, .ppt/.pptx, .xls/.xlsx
- **Binary Files**: Preserved as Base64 for external processing

## Development

### Project Structure
```
outlook-mcp/
├── server/
│   ├── index.js              # Main MCP server
│   ├── auth/                 # Authentication management
│   ├── graph/                # Microsoft Graph API client
│   ├── tools/                # MCP tool implementations
│   │   ├── email/            # Email-related tools
│   │   ├── sharepoint/       # SharePoint tools
│   │   ├── calendar/         # Calendar tools
│   │   └── contacts/         # Contact tools
│   └── utils/                # Utility modules
│       ├── fileOutput.js     # Large file handling
│       └── mcpErrorResponse.js
├── docs/                     # Documentation
└── package.json
```

### Running Tests
```bash
npm test                    # Run all tests
npm run test:watch          # Watch mode
npm run test:benchmark      # Performance benchmarks
```

### Debugging
```bash
npm run test:graph          # Test Graph API connection
```

## Troubleshooting

### Large File Issues
- **Problem**: "Result exceeds maximum length" error
- **Solution**: Ensure `MCP_OUTLOOK_WORK_DIR` is set and writable
- **Alternative**: Files automatically save to system temp if work dir not configured

### Authentication Issues
- **Problem**: Authentication failures
- **Solution**: Verify Azure AD app permissions and client ID
- **Reset**: Clear stored tokens and re-authenticate

### SharePoint Access Issues
- **Problem**: Cannot access SharePoint files
- **Solution**: Ensure sharing links are valid and user has access permissions
- **Alternative**: Use direct file ID access if available

## License

ISC License

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make changes with tests
4. Submit a pull request

For detailed development information, see the documentation in the `docs/` directory.
