// Folder resolution utilities for Microsoft Graph API
import { convertErrorToToolError, createValidationError } from '../utils/mcpErrorResponse.js';

/**
 * Resolves folder names to folder IDs for Microsoft Graph API calls
 */
export class FolderResolver {
  constructor(graphApiClient) {
    this.graphApiClient = graphApiClient;
    this.foldersByName = new Map(); // Cache folders by display name (case insensitive)
    this.foldersById = new Map(); // Cache folders by ID
    this.foldersList = []; // Master list of all folders
    this.cacheExpiry = 5 * 60 * 1000; // 5 minutes
    this.lastCacheUpdate = 0;
  }

  /**
   * Get all folders and cache them
   */
  async refreshFolderCache() {
    try {
      const result = await this.graphApiClient.makeRequest('/me/mailFolders', {
        select: 'id,displayName,parentFolderId',
        top: 1000 // Get up to 1000 folders
      });

      this.foldersByName.clear();
      this.foldersById.clear();
      this.foldersList = [];
      
      if (result.value) {
        result.value.forEach(folder => {
          const folderInfo = {
            id: folder.id,
            displayName: folder.displayName,
            parentFolderId: folder.parentFolderId
          };
          
          // Store by display name (case insensitive)
          this.foldersByName.set(folder.displayName.toLowerCase(), folderInfo);
          // Store by ID (case insensitive)
          this.foldersById.set(folder.id.toLowerCase(), folderInfo);
          // Add to master list
          this.foldersList.push(folderInfo);
        });
      }

      this.lastCacheUpdate = Date.now();
      console.error(`Folder cache refreshed with ${this.foldersList.length} folders`);
      
    } catch (error) {
      console.error('Failed to refresh folder cache:', error);
      throw error;
    }
  }

  /**
   * Check if folder cache needs refreshing
   */
  shouldRefreshCache() {
    return (Date.now() - this.lastCacheUpdate) > this.cacheExpiry || this.foldersList.length === 0;
  }

  /**
   * Resolve folder name or ID to folder ID
   * @param {string} folderNameOrId - Folder name or ID
   * @returns {Promise<string>} - Folder ID
   */
  async resolveFolderToId(folderNameOrId) {
    if (!folderNameOrId) {
      throw new Error('Folder name or ID is required');
    }

    // Handle special case for 'inbox'
    if (folderNameOrId.toLowerCase() === 'inbox') {
      return 'inbox'; // Microsoft Graph accepts 'inbox' as a special folder name
    }

    // Check if it's already a folder ID (Microsoft Graph uses base64-like strings)
    // Folder IDs are typically long alphanumeric strings with + and = characters
    const folderIdRegex = /^[A-Za-z0-9+/]+=*$/;
    if (folderIdRegex.test(folderNameOrId) && folderNameOrId.length > 20) {
      return folderNameOrId; // It's already a valid folder ID
    }

    // Refresh cache if needed
    if (this.shouldRefreshCache()) {
      await this.refreshFolderCache();
    }

    // Look up by display name (case insensitive)
    const folderInfo = this.foldersByName.get(folderNameOrId.toLowerCase());
    if (folderInfo) {
      return folderInfo.id;
    }

    // If not found, try refreshing cache once more in case folder was recently created
    await this.refreshFolderCache();
    const refreshedFolderInfo = this.foldersByName.get(folderNameOrId.toLowerCase());
    if (refreshedFolderInfo) {
      return refreshedFolderInfo.id;
    }

    // Folder not found
    throw new Error(`Folder '${folderNameOrId}' not found. Available folders: ${
      this.foldersList
        .map(f => f.displayName)
        .join(', ')
    }`);
  }

  /**
   * Resolve multiple folder names/IDs to folder IDs
   * @param {string[]} folderNamesOrIds - Array of folder names or IDs
   * @returns {Promise<string[]>} - Array of folder IDs
   */
  async resolveFoldersToIds(folderNamesOrIds) {
    if (!Array.isArray(folderNamesOrIds) || folderNamesOrIds.length === 0) {
      return [];
    }

    const resolvedIds = [];
    for (const folderNameOrId of folderNamesOrIds) {
      try {
        const folderId = await this.resolveFolderToId(folderNameOrId);
        resolvedIds.push(folderId);
      } catch (error) {
        throw new Error(`Failed to resolve folder '${folderNameOrId}': ${error.message}`);
      }
    }

    return resolvedIds;
  }

  /**
   * Get folder info by name or ID
   * @param {string} folderNameOrId - Folder name or ID
   * @returns {Promise<object>} - Folder info object
   */
  async getFolderInfo(folderNameOrId) {
    const folderId = await this.resolveFolderToId(folderNameOrId);
    
    // If we already have it in cache, return it
    const cachedInfo = this.foldersById.get(folderId.toLowerCase());
    if (cachedInfo) {
      return cachedInfo;
    }

    // Otherwise make direct API call
    try {
      const folderData = await this.graphApiClient.makeRequest(`/me/mailFolders/${folderId}`, {
        select: 'id,displayName,parentFolderId,totalItemCount,unreadItemCount'
      });

      return {
        id: folderData.id,
        displayName: folderData.displayName,
        parentFolderId: folderData.parentFolderId,
        totalItemCount: folderData.totalItemCount,
        unreadItemCount: folderData.unreadItemCount
      };
    } catch (error) {
      throw new Error(`Failed to get folder info for '${folderNameOrId}': ${error.message}`);
    }
  }

  /**
   * List all available folders with their names and IDs
   * @returns {Promise<object[]>} - Array of folder info objects
   */
  async listAllFolders() {
    if (this.shouldRefreshCache()) {
      await this.refreshFolderCache();
    }

    return [...this.foldersList]
      .sort((a, b) => a.displayName.localeCompare(b.displayName));
  }
}
