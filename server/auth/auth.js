// OAuth 2.0 authentication is handled manually with PKCE flow
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenManager } from './tokenManager.js';
import { authConfig } from './config.js';
import { GraphApiClient } from '../graph/graphClient.js';
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
    this.authenticationRecord = null;
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
      return {
        success: false,
        error: error.message,
      };
    }
  }

  async authenticateInteractive() {
    const codeVerifier = this.tokenManager.generateCodeVerifier();
    const codeChallenge = this.tokenManager.generateCodeChallenge(codeVerifier);
    await this.tokenManager.storePKCEVerifier(codeVerifier);

    const authorizationCode = await this.getAuthorizationCode(codeChallenge);
    
    if (!authorizationCode) {
      throw new Error('Failed to get authorization code');
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
      authUrl.searchParams.append('redirect_uri', authConfig.oauth.redirectUri);
      authUrl.searchParams.append('scope', authConfig.oauth.scope);
      authUrl.searchParams.append('state', state);
      authUrl.searchParams.append('code_challenge', codeChallenge);
      authUrl.searchParams.append('code_challenge_method', 'S256');
      authUrl.searchParams.append('prompt', 'select_account');

      console.log(`\nOpening your browser for Microsoft account selection...`);
      console.log(`If the browser doesn't open automatically, please visit:`);
      console.log(authUrl.toString());
      
      // Attempt to open the browser automatically
      this.openBrowser(authUrl.toString());

      const server = http.createServer(async (req, res) => {
        const parsedUrl = url.parse(req.url, true);
        
        if (parsedUrl.pathname === '/callback') {
          const code = parsedUrl.query.code;
          const returnedState = parsedUrl.query.state;
          
          if (returnedState !== state) {
            res.writeHead(400, { 'Content-Type': 'text/html' });
            res.end('<h1>Error: State mismatch</h1>');
            server.close();
            reject(new Error('State mismatch - possible CSRF attack'));
            return;
          }

          if (code) {
            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.end(`
              <html>
                <head>
                  <title>Authentication Successful</title>
                  <style>
                    body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
                    h1 { color: #0078d4; }
                  </style>
                </head>
                <body>
                  <h1>Authentication Successful!</h1>
                  <p>You can now close this window and return to your terminal.</p>
                  <p>The Outlook MCP server is being configured with your selected account.</p>
                </body>
              </html>
            `);
            server.close();
            resolve(code);
          } else {
            res.writeHead(400, { 'Content-Type': 'text/html' });
            res.end('<h1>Authentication failed</h1>');
            server.close();
            reject(new Error('No authorization code received'));
          }
        }
      });

      server.listen(8080, () => {
        console.log('\nWaiting for authentication callback...');
      });

      setTimeout(() => {
        server.close();
        reject(new Error('Authentication timeout'));
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
      redirect_uri: authConfig.oauth.redirectUri,
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
      throw new Error(`Token exchange failed: ${error}`);
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
        throw new Error(`Token refresh failed: ${error}`);
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
      throw error;
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
      throw error;
    }
  }

  async ensureAuthenticated() {
    if (!this.isAuthenticated || !this.graphClient) {
      const result = await this.authenticate();
      if (!result.success) {
        throw new Error(`Authentication failed: ${result.error}`);
      }
    }

    try {
      await this.tokenManager.getAccessToken();
    } catch (error) {
      if (error.message.includes('needs refresh')) {
        await this.refreshAccessToken();
      } else {
        throw error;
      }
    }

    return this.graphClient;
  }

  getGraphClient() {
    if (!this.graphClient) {
      throw new Error('Not authenticated. Call authenticate() first.');
    }
    return this.graphClient;
  }

  getGraphApiClient() {
    if (!this.graphApiClient) {
      throw new Error('Not authenticated. Call authenticate() first.');
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