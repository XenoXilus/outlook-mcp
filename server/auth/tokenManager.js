// Import keytar as optional dependency - it may not be available in all environments
let keytar;
let keytarImportPromise;

// Lazy import keytar when needed
const getKeytar = async () => {
  if (keytar === undefined) {
    if (!keytarImportPromise) {
      keytarImportPromise = (async () => {
        try {
          const keytarModule = await import('keytar');
          return keytarModule.default;
        } catch (error) {
          console.error('Keytar not available - using fallback token storage:', error.message);
          return null;
        }
      })();
    }
    keytar = await keytarImportPromise;
  }
  return keytar;
};
import storage from 'node-persist';
import crypto from 'crypto';
import path from 'path';
import { fileURLToPath } from 'url';
import { createAuthError, convertErrorToToolError } from '../utils/mcpErrorResponse.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SERVICE_NAME = 'outlook-mcp';
const ENCRYPTION_KEY_ACCOUNT = 'encryption-key';
const ACCESS_TOKEN_ACCOUNT = 'access-token';
const REFRESH_TOKEN_ACCOUNT = 'refresh-token';
const TOKEN_METADATA_KEY = 'token-metadata';

export class TokenManager {
  constructor(clientId) {
    this.clientId = clientId;
    this.storageInitialized = false;
    this.encryptionKey = null;
  }

  async initialize() {
    if (this.storageInitialized) return;

    await storage.init({
      dir: (process.env.MCP_OUTLOOK_REFRESH_TOKEN_PATH || '').trim() || path.join(__dirname, '../../.tokens'),
      logging: false,
    });

    this.encryptionKey = await this.getOrCreateEncryptionKey();
    this.storageInitialized = true;
  }

  async getOrCreateEncryptionKey() {
    try {
      const keytarInstance = await getKeytar();
      if (keytarInstance) {
        const existingKey = await keytarInstance.getPassword(SERVICE_NAME, ENCRYPTION_KEY_ACCOUNT);
        if (existingKey) {
          return Buffer.from(existingKey, 'base64');
        }

        const newKey = crypto.randomBytes(32);
        await keytarInstance.setPassword(SERVICE_NAME, ENCRYPTION_KEY_ACCOUNT, newKey.toString('base64'));
        return newKey;
      }
    } catch (error) {
      // Fall through to fallback
    }
    
    // Fallback for environments without keytar (containers, MCP servers, etc.)
    // Tokens will still be encrypted and stored securely in the file system
    const fallbackKey = crypto.createHash('sha256')
      .update(this.clientId + (process.env.AZURE_TENANT_ID || 'default'))
      .digest();
    return fallbackKey;
  }

  encrypt(text) {
    const iv = crypto.randomBytes(16);
    const cipher = crypto.createCipheriv('aes-256-cbc', this.encryptionKey, iv);
    let encrypted = cipher.update(text, 'utf8', 'hex');
    encrypted += cipher.final('hex');
    return iv.toString('hex') + ':' + encrypted;
  }

  decrypt(encryptedText) {
    const parts = encryptedText.split(':');
    const iv = Buffer.from(parts.shift(), 'hex');
    const encrypted = parts.join(':');
    const decipher = crypto.createDecipheriv('aes-256-cbc', this.encryptionKey, iv);
    let decrypted = decipher.update(encrypted, 'hex', 'utf8');
    decrypted += decipher.final('utf8');
    return decrypted;
  }

  async storeTokens(accessToken, refreshToken, expiresIn = 3600) {
    await this.initialize();

    let usingFallback = false;
    try {
      const keytarInstance = await getKeytar();
      if (keytarInstance) {
        await keytarInstance.setPassword(SERVICE_NAME, ACCESS_TOKEN_ACCOUNT, this.encrypt(accessToken));
        if (refreshToken) {
          await keytarInstance.setPassword(SERVICE_NAME, REFRESH_TOKEN_ACCOUNT, this.encrypt(refreshToken));
        }
      } else {
        usingFallback = true;
        await storage.setItem('fallback_access_token', this.encrypt(accessToken));
        if (refreshToken) {
          await storage.setItem('fallback_refresh_token', this.encrypt(refreshToken));
        }
      }
    } catch (error) {
      // Keytar not available in this environment, using secure file storage instead
      usingFallback = true;
      await storage.setItem('fallback_access_token', this.encrypt(accessToken));
      if (refreshToken) {
        await storage.setItem('fallback_refresh_token', this.encrypt(refreshToken));
      }
    }

    const metadata = {
      accessTokenExpiry: Date.now() + (expiresIn * 1000),
      refreshTokenExpiry: Date.now() + (90 * 24 * 60 * 60 * 1000), // 90 days
      lastRefresh: Date.now(),
    };
    await storage.setItem(TOKEN_METADATA_KEY, metadata);

    if (usingFallback) {
      console.error('Tokens stored securely in encrypted file storage');
    }
  }

  async getAccessToken() {
    try {
      await this.initialize();

      const metadata = await storage.getItem(TOKEN_METADATA_KEY);
      if (!metadata) {
        throw createAuthError('No token metadata found', true);
      }

      const refreshThreshold = 55 * 60 * 1000; // 55 minutes
      const shouldRefresh = Date.now() > (metadata.accessTokenExpiry - refreshThreshold);

      if (shouldRefresh) {
        const error = createAuthError('Access token needs refresh', true);
        error._errorDetails = { ...error._errorDetails, needsRefresh: true };
        throw error;
      }

      const keytarInstance = await getKeytar();
      if (keytarInstance) {
        try {
          const encryptedToken = await keytarInstance.getPassword(SERVICE_NAME, ACCESS_TOKEN_ACCOUNT);
          if (encryptedToken) {
            return this.decrypt(encryptedToken);
          }
        } catch (error) {
          // Fall through to fallback
        }
      }
      
      const fallbackToken = await storage.getItem('fallback_access_token');
      if (fallbackToken) {
        return this.decrypt(fallbackToken);
      }

      throw createAuthError('No access token found', true);
    } catch (error) {
      if (error.isError) {
        // Already an MCP error, re-throw as-is
        throw error;
      }
      throw convertErrorToToolError(error, 'Failed to retrieve access token');
    }
  }

  async getRefreshToken() {
    try {
      await this.initialize();

      const metadata = await storage.getItem(TOKEN_METADATA_KEY);
      if (!metadata) {
        throw createAuthError('No token metadata found', true);
      }

      if (Date.now() > metadata.refreshTokenExpiry) {
        throw createAuthError('Refresh token has expired', true);
      }

      const keytarInstance = await getKeytar();
      if (keytarInstance) {
        try {
          const encryptedToken = await keytarInstance.getPassword(SERVICE_NAME, REFRESH_TOKEN_ACCOUNT);
          if (encryptedToken) {
            return this.decrypt(encryptedToken);
          }
        } catch (error) {
          // Fall through to fallback
        }
      }
      
      const fallbackToken = await storage.getItem('fallback_refresh_token');
      if (fallbackToken) {
        return this.decrypt(fallbackToken);
      }

      throw createAuthError('No refresh token found', true);
    } catch (error) {
      if (error.isError) {
        // Already an MCP error, re-throw as-is
        throw error;
      }
      throw convertErrorToToolError(error, 'Failed to retrieve refresh token');
    }
  }

  async clearTokens() {
    await this.initialize();

    const keytarInstance = await getKeytar();
    if (keytarInstance) {
      try {
        await keytarInstance.deletePassword(SERVICE_NAME, ACCESS_TOKEN_ACCOUNT);
        await keytarInstance.deletePassword(SERVICE_NAME, REFRESH_TOKEN_ACCOUNT);
      } catch (error) {
        // Silently continue - keytar might not be available
      }
    }

    await storage.removeItem('fallback_access_token');
    await storage.removeItem('fallback_refresh_token');
    await storage.removeItem(TOKEN_METADATA_KEY);
  }

  generateCodeVerifier() {
    return crypto.randomBytes(32).toString('base64url');
  }

  generateCodeChallenge(verifier) {
    return crypto.createHash('sha256')
      .update(verifier)
      .digest('base64url');
  }

  async storePKCEVerifier(verifier) {
    await this.initialize();
    await storage.setItem('pkce_verifier', verifier);
  }

  async getPKCEVerifier() {
    try {
      await this.initialize();
      const verifier = await storage.getItem('pkce_verifier');
      await storage.removeItem('pkce_verifier');
      if (!verifier) {
        throw createAuthError('PKCE verifier not found or expired', true);
      }
      return verifier;
    } catch (error) {
      if (error.isError) {
        // Already an MCP error, re-throw as-is
        throw error;
      }
      throw convertErrorToToolError(error, 'Failed to retrieve PKCE verifier');
    }
  }

  async isAuthenticated() {
    try {
      await this.getAccessToken();
      return true;
    } catch (error) {
      // For isAuthenticated, we just return false instead of throwing
      // as this is used for checking authentication status
      return false;
    }
  }

  async hasRefreshToken() {
    try {
      await this.getRefreshToken();
      return true;
    } catch {
      return false;
    }
  }

  async getTokenMetadata() {
    await this.initialize();
    return await storage.getItem(TOKEN_METADATA_KEY);
  }
}