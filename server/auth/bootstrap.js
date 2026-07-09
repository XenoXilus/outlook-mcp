#!/usr/bin/env node
/**
 * One-time interactive bootstrap for headless runs (NFR-1).
 *
 * Runs the browser PKCE flow once and stores the encrypted refresh token, so
 * scheduled runs with MCP_OUTLOOK_AUTH_MODE=headless can refresh silently.
 * Usage: npm run auth:bootstrap  (AZURE_CLIENT_ID / AZURE_TENANT_ID must be set)
 */

import 'dotenv/config';
import { OutlookAuthManager } from './auth.js';

const clientId = process.env.AZURE_CLIENT_ID;
const tenantId = process.env.AZURE_TENANT_ID;

if (!clientId || !tenantId) {
  console.error('AZURE_CLIENT_ID and AZURE_TENANT_ID must be set (env or .env file).');
  process.exit(1);
}

const manager = new OutlookAuthManager(clientId, tenantId);
const result = await manager.authenticate();

if (result.success) {
  console.error(`Authenticated as ${result.user.displayName} <${result.user.mail}>.`);
  console.error('Refresh token stored. Headless runs (MCP_OUTLOOK_AUTH_MODE=headless) will now refresh silently.');
  if (process.env.MCP_OUTLOOK_REFRESH_TOKEN_PATH) {
    console.error(`Token store: ${process.env.MCP_OUTLOOK_REFRESH_TOKEN_PATH}`);
  }
  process.exit(0);
} else {
  console.error('Bootstrap failed:', result.error?.content?.[0]?.text || result.error);
  process.exit(1);
}
