// OAuth 2.0 authentication is handled manually with PKCE flow
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenManager } from './tokenManager.js';
import { authConfig } from './config.js';
import { GraphApiClient } from '../graph/graphClient.js';
import { createAuthError, convertErrorToToolError } from '../utils/mcpErrorResponse.js';
import http from 'http';
import url from 'url';
import crypto from 'crypto';
import { exec } from 'child_process';

export class OutlookAuthManager {
  constructor(clientId, tenantId) {
    this.clientId = clientId;
    this.tenantId = tenantId;
    this.tokenManager = new TokenManager(clientId);
    this.graphClient = null;
    this.graphApiClient = null;
    this.isAuthenticated = false;
    this.isAuthenticated = false;
    this.authenticationRecord = null;
    this.lastUsedPort = null;
  }

  openBrowser(url) {
    const platform = process.platform;
    let command;

    switch (platform) {
      case 'darwin': // macOS
        command = `open "${url}"`;
        break;
      case 'win32': // Windows
        command = `start "" "${url}"`;
        break;
      default: // Linux and others
        command = `xdg-open "${url}"`;
        break;
    }

    exec(command, (error) => {
      if (error) {
        // Silent fail - URL is already displayed above
      }
    });
  }

  async authenticate() {
    try {
      const isTokenValid = await this.tokenManager.isAuthenticated();

      if (isTokenValid) {
        await this.initializeGraphClient();
        return await this.validateAuthentication();
      }

      // Use interactive authentication with PKCE for delegated access
      return await this.authenticateInteractive();
    } catch (error) {
      console.error('Authentication error:', error);
      this.isAuthenticated = false;
      if (error.isError) {
        // Already an MCP error, return as-is
        return {
          success: false,
          error: error,
        };
      }
      return {
        success: false,
        error: createAuthError(error.message, true),
      };
    }
  }

  async authenticateInteractive() {
    const codeVerifier = this.tokenManager.generateCodeVerifier();
    const codeChallenge = this.tokenManager.generateCodeChallenge(codeVerifier);
    await this.tokenManager.storePKCEVerifier(codeVerifier);

    const authorizationCode = await this.getAuthorizationCode(codeChallenge);

    if (!authorizationCode) {
      return {
        success: false,
        error: createAuthError('Failed to get authorization code', true),
      };
    }

    const tokenResponse = await this.exchangeCodeForToken(authorizationCode);

    await this.tokenManager.storeTokens(
      tokenResponse.access_token,
      tokenResponse.refresh_token,
      tokenResponse.expires_in
    );

    await this.initializeGraphClient();
    return await this.validateAuthentication();
  }

  async getAuthorizationCode(codeChallenge) {
    return new Promise((resolve, reject) => {
      const state = crypto.randomBytes(16).toString('hex');
      const authUrl = new URL(authConfig.oauth.authorizeUrl(this.tenantId));

      authUrl.searchParams.append('client_id', this.clientId);
      authUrl.searchParams.append('response_type', 'code');
      authUrl.searchParams.append('redirect_uri', `http://localhost:${server.address().port}/callback`);
      authUrl.searchParams.append('scope', authConfig.oauth.scope);
      authUrl.searchParams.append('state', state);
      authUrl.searchParams.append('code_challenge', codeChallenge);
      authUrl.searchParams.append('code_challenge_method', 'S256');
      authUrl.searchParams.append('prompt', 'select_account');

      console.error(`\nOpening your browser for Microsoft account selection...`);
      console.error(`If the browser doesn't open automatically, please visit:`);
      console.error(authUrl.toString());

      // Attempt to open the browser automatically
      this.openBrowser(authUrl.toString());

      const server = http.createServer(async (req, res) => {
        const parsedUrl = url.parse(req.url, true);

        if (parsedUrl.pathname === '/callback') {
          const code = parsedUrl.query.code;
          const returnedState = parsedUrl.query.state;

          if (returnedState !== state) {
            res.writeHead(400, { 'Content-Type': 'text/html' });
            res.end(`
              <html>
                <head>
                  <title>Authentication Error</title>
                  <style>
                    body { 
                      font-family: 'Segoe UI', Arial, sans-serif; 
                      text-align: center; 
                      padding: 50px;
                      background-color: #f3f2f1;
                      margin: 0;
                    }
                    .container {
                      background: white;
                      border-radius: 8px;
                      padding: 40px;
                      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                      max-width: 400px;
                      margin: 0 auto;
                    }
                    h1 { 
                      color: #d83b01; 
                      margin-bottom: 20px;
                    }
                    .error-icon {
                      width: 80px;
                      height: 80px;
                      margin: 0 auto 20px;
                      background-color: #d83b01;
                      border-radius: 50%;
                      display: flex;
                      align-items: center;
                      justify-content: center;
                    }
                    .error-icon svg {
                      width: 50px;
                      height: 50px;
                      fill: white;
                    }
                    .instructions {
                      color: #605e5c;
                      font-size: 14px;
                      margin-top: 20px;
                    }
                  </style>
                </head>
                <body>
                  <div class="container">
                    <div class="error-icon">
                      <svg viewBox="0 0 24 24">
                        <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/>
                      </svg>
                    </div>
                    <h1>Security Error</h1>
                    <p>The authentication request could not be verified.</p>
                    <p class="instructions">Please disconnect and reconnect the MCP server to try again.</p>
                  </div>
                </body>
              </html>
            `);
            server.close();
            reject(createAuthError('State mismatch - possible CSRF attack', false));
            return;
          }

          if (code) {
            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.end(`
              <html>
                <head>
                  <title>Authentication Successful</title>
                  <style>
                    body { 
                      font-family: 'Segoe UI', Arial, sans-serif; 
                      text-align: center; 
                      padding: 50px;
                      background-color: #f3f2f1;
                      margin: 0;
                    }
                    .container {
                      background: white;
                      border-radius: 8px;
                      padding: 40px;
                      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                      max-width: 400px;
                      margin: 0 auto;
                    }
                    h1 { 
                      color: #0078d4; 
                      margin-bottom: 20px;
                    }
                    .checkmark {
                      width: 80px;
                      height: 80px;
                      margin: 0 auto 20px;
                      background-color: #107c10;
                      border-radius: 50%;
                      display: flex;
                      align-items: center;
                      justify-content: center;
                      animation: scaleIn 0.3s ease-in-out;
                    }
                    .checkmark svg {
                      width: 50px;
                      height: 50px;
                      fill: white;
                    }
                    @keyframes scaleIn {
                      from { transform: scale(0); opacity: 0; }
                      to { transform: scale(1); opacity: 1; }
                    }
                    .countdown {
                      color: #605e5c;
                      font-size: 14px;
                      margin-top: 20px;
                    }
                    #timer {
                      font-weight: bold;
                      color: #0078d4;
                    }
                    .manual-close {
                      font-size: 12px;
                      color: #a19f9d;
                      margin-top: 10px;
                    }
                  </style>
                </head>
                <body>
                  <div class="container">
                    <div class="checkmark">
                      <svg viewBox="0 0 24 24">
                        <path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41L9 16.17z"/>
                      </svg>
                    </div>
                    <h1>Authentication Successful!</h1>
                    <p>The Outlook MCP server has been configured with your selected account.</p>
                    <p class="countdown">This window will close in <span id="timer">5</span> seconds...</p>
                    <p class="manual-close">If the window doesn't close automatically, you can close it manually.</p>
                  </div>
                  <script>
                    let countdown = 5;
                    const timerElement = document.getElementById('timer');
                    
                    const interval = setInterval(() => {
                      countdown--;
                      timerElement.textContent = countdown;
                      
                      if (countdown <= 0) {
                        clearInterval(interval);
                        // Try to close the window
                        window.close();
                        // If window.close() doesn't work (blocked by browser), update the message
                        setTimeout(() => {
                          document.querySelector('.countdown').textContent = 'You can now close this window.';
                        }, 500);
                      }
                    }, 1000);
                  </script>
                </body>
              </html>
            `);
            server.close();
            resolve(code);
          } else {
            res.writeHead(400, { 'Content-Type': 'text/html' });
            res.end(`
              <html>
                <head>
                  <title>Authentication Failed</title>
                  <style>
                    body { 
                      font-family: 'Segoe UI', Arial, sans-serif; 
                      text-align: center; 
                      padding: 50px;
                      background-color: #f3f2f1;
                      margin: 0;
                    }
                    .container {
                      background: white;
                      border-radius: 8px;
                      padding: 40px;
                      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                      max-width: 400px;
                      margin: 0 auto;
                    }
                    h1 { 
                      color: #d83b01; 
                      margin-bottom: 20px;
                    }
                    .error-icon {
                      width: 80px;
                      height: 80px;
                      margin: 0 auto 20px;
                      background-color: #d83b01;
                      border-radius: 50%;
                      display: flex;
                      align-items: center;
                      justify-content: center;
                    }
                    .error-icon svg {
                      width: 50px;
                      height: 50px;
                      fill: white;
                    }
                    .instructions {
                      color: #605e5c;
                      font-size: 14px;
                      margin-top: 20px;
                    }
                  </style>
                </head>
                <body>
                  <div class="container">
                    <div class="error-icon">
                      <svg viewBox="0 0 24 24">
                        <path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12 19 6.41z"/>
                      </svg>
                    </div>
                    <h1>Authentication Failed</h1>
                    <p>The authentication process was cancelled or failed.</p>
                    <p class="instructions">Please disconnect and reconnect the MCP server to try again.</p>
                  </div>
                </body>
              </html>
            `);
            server.close();
            reject(createAuthError('No authorization code received', true));
          }
        }
      });

      server.listen(0, () => {
        const port = server.address().port;
        this.lastUsedPort = port;
        console.error(`\nListening for authentication callback on port ${port}...`);

        // Update redirect URI with actual port
        authUrl.searchParams.set('redirect_uri', `http://localhost:${port}/callback`);

        console.error(`\nOpening your browser for Microsoft account selection...`);
        console.error(`If the browser doesn't open automatically, please visit:`);
        console.error(authUrl.toString());

        // Attempt to open the browser automatically
        this.openBrowser(authUrl.toString());
      });

      setTimeout(() => {
        server.close();
        reject(createAuthError('Authentication timeout - please try again', true));
      }, 5 * 60 * 1000); // 5 minute timeout
    });
  }

  async exchangeCodeForToken(code) {
    const codeVerifier = await this.tokenManager.getPKCEVerifier();

    const tokenUrl = authConfig.oauth.tokenUrl(this.tenantId);
    const params = new URLSearchParams({
      client_id: this.clientId,
      scope: authConfig.oauth.scope,
      code: code,
      redirect_uri: `http://localhost:${this.lastUsedPort}/callback`,
      grant_type: 'authorization_code',
      code_verifier: codeVerifier,
    });

    const response = await fetch(tokenUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: params.toString(),
    });

    if (!response.ok) {
      const error = await response.text();
      throw createAuthError(`Token exchange failed: ${error}`, true);
    }

    return await response.json();
  }

  async refreshAccessToken() {
    try {
      const refreshToken = await this.tokenManager.getRefreshToken();
      const tokenUrl = authConfig.oauth.tokenUrl(this.tenantId);

      const params = new URLSearchParams({
        client_id: this.clientId,
        scope: authConfig.oauth.scope,
        refresh_token: refreshToken,
        grant_type: 'refresh_token',
      });

      const response = await fetch(tokenUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: params.toString(),
      });

      if (!response.ok) {
        const error = await response.text();
        throw createAuthError(`Token refresh failed: ${error}`, true);
      }

      const tokenResponse = await response.json();

      await this.tokenManager.storeTokens(
        tokenResponse.access_token,
        tokenResponse.refresh_token || refreshToken,
        tokenResponse.expires_in
      );

      await this.initializeGraphClient();
      return true;
    } catch (error) {
      console.error('Token refresh failed:', error);
      await this.tokenManager.clearTokens();
      if (error.isError) {
        // Already an MCP error, re-throw as-is
        throw error;
      }
      throw convertErrorToToolError(error, 'Token refresh failed');
    }
  }

  async initializeGraphClient() {
    const authProvider = {
      getAccessToken: async () => {
        try {
          return await this.tokenManager.getAccessToken();
        } catch (error) {
          if (error.message.includes('needs refresh')) {
            await this.refreshAccessToken();
            return await this.tokenManager.getAccessToken();
          }
          throw error;
        }
      },
    };

    this.graphClient = Client.init({
      authProvider: (done) => {
        authProvider.getAccessToken()
          .then(token => done(null, token))
          .catch(error => done(error, null));
      },
      defaultVersion: 'v1.0',
    });

    // Initialize the enhanced GraphApiClient
    this.graphApiClient = new GraphApiClient(this);
    await this.graphApiClient.initialize();
  }

  async validateAuthentication() {
    try {
      const user = await this.graphClient.api('/me').get();
      this.isAuthenticated = true;

      return {
        success: true,
        user: {
          id: user.id,
          displayName: user.displayName,
          mail: user.mail || user.userPrincipalName,
        },
      };
    } catch (error) {
      this.isAuthenticated = false;
      if (error.isError) {
        // Already an MCP error, re-throw as-is
        throw error;
      }
      throw convertErrorToToolError(error, 'User validation failed');
    }
  }

  async ensureAuthenticated() {
    if (!this.isAuthenticated || !this.graphClient) {
      const result = await this.authenticate();
      if (!result.success) {
        if (result.error.isError) {
          // Already an MCP error, re-throw as-is
          throw result.error;
        }
        throw createAuthError(`Authentication failed: ${result.error}`, true);
      }
    }

    try {
      await this.tokenManager.getAccessToken();
    } catch (error) {
      if (error.isError) {
        // Handle MCP errors from token manager
        if (error._errorDetails && error._errorDetails.needsRefresh) {
          await this.refreshAccessToken();
        } else {
          throw error;
        }
      } else if (error.message.includes('needs refresh')) {
        await this.refreshAccessToken();
      } else {
        throw convertErrorToToolError(error, 'Token validation failed');
      }
    }

    return this.graphClient;
  }

  getGraphClient() {
    if (!this.graphClient) {
      throw createAuthError('Not authenticated. Call authenticate() first.', true);
    }
    return this.graphClient;
  }

  getGraphApiClient() {
    if (!this.graphApiClient) {
      throw createAuthError('Not authenticated. Call authenticate() first.', true);
    }
    return this.graphApiClient;
  }

  async logout() {
    await this.tokenManager.clearTokens();
    this.graphClient = null;
    this.graphApiClient = null;
    this.isAuthenticated = false;
    this.authenticationRecord = null;
  }
}