export const authConfig = {
  oauth: {
    authorizeUrl: (tenantId) => 
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`,
    tokenUrl: (tenantId) => 
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    scope: [
      'Mail.Read',
      'Mail.ReadWrite', 
      'Mail.Send',
      'Calendars.Read',
      'Calendars.ReadWrite',
      'Contacts.Read',
      'Contacts.ReadWrite',
      'Tasks.Read',
      'Tasks.ReadWrite',
      'User.Read',
      'MailboxSettings.Read',
      // SharePoint and OneDrive access
      'Sites.Read.All',         // Read all SharePoint sites
      'Sites.ReadWrite.All',    // Read/write all SharePoint sites
      'Files.Read.All',         // Read all files user can access
      'Files.ReadWrite.All',    // Read/write all files user can access
      'offline_access',         // Required for refresh tokens
    ].join(' '),
    redirectUri: 'http://localhost:8080/callback',
  },
  
  token: {
    accessTokenTTL: 60 * 60 * 1000, // 60 minutes in milliseconds
    refreshThreshold: 55 * 60 * 1000, // Refresh at 55 minutes
    refreshTokenTTL: 90 * 24 * 60 * 60 * 1000, // 90 days
  },
  
  retry: {
    maxAttempts: 3,
    initialDelay: 1000, // 1 second
    maxDelay: 30000, // 30 seconds
    backoffMultiplier: 2,
  },
  
  security: {
    usePKCE: true,        // PKCE ensures secure authentication without client secrets
    encryptTokens: true,  // Tokens are encrypted in storage
    auditLogging: true,   // All authentication events are logged
  },
};