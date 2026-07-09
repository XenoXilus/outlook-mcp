import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { OutlookAuthManager } from '../../auth/auth.js';

describe('headless auth mode (NFR-1)', () => {
  let manager, savedMode;

  beforeEach(() => {
    savedMode = process.env.MCP_OUTLOOK_AUTH_MODE;
    manager = new OutlookAuthManager('client-id', 'tenant-id');
    // Never allow a real browser/interactive flow in tests
    manager.authenticateInteractive = vi.fn().mockResolvedValue({ success: true, user: {} });
    manager.initializeGraphClient = vi.fn().mockResolvedValue(undefined);
    manager.validateAuthentication = vi.fn().mockResolvedValue({ success: true, user: { displayName: 'V' } });
  });

  afterEach(() => {
    if (savedMode === undefined) delete process.env.MCP_OUTLOOK_AUTH_MODE;
    else process.env.MCP_OUTLOOK_AUTH_MODE = savedMode;
  });

  it('defaults to interactive mode', () => {
    delete process.env.MCP_OUTLOOK_AUTH_MODE;
    expect(manager.getAuthMode()).toBe('interactive');
  });

  it('silently refreshes instead of going interactive when a refresh token exists', async () => {
    manager.tokenManager.isAuthenticated = vi.fn().mockResolvedValue(false);
    manager.tokenManager.hasRefreshToken = vi.fn().mockResolvedValue(true);
    manager.refreshAccessToken = vi.fn().mockResolvedValue(true);

    const result = await manager.authenticate();
    expect(result.success).toBe(true);
    expect(manager.refreshAccessToken).toHaveBeenCalled();
    expect(manager.authenticateInteractive).not.toHaveBeenCalled();
  });

  it('headless mode fails fast with an actionable error when no refresh token exists', async () => {
    process.env.MCP_OUTLOOK_AUTH_MODE = 'headless';
    manager.tokenManager.isAuthenticated = vi.fn().mockResolvedValue(false);
    manager.tokenManager.hasRefreshToken = vi.fn().mockResolvedValue(false);

    const result = await manager.authenticate();
    expect(result.success).toBe(false);
    expect(manager.authenticateInteractive).not.toHaveBeenCalled();
    expect(result.error.content[0].text).toMatch(/auth:bootstrap/);
  });

  it('headless mode fails fast (no browser) when the silent refresh fails', async () => {
    process.env.MCP_OUTLOOK_AUTH_MODE = 'headless';
    manager.tokenManager.isAuthenticated = vi.fn().mockResolvedValue(false);
    manager.tokenManager.hasRefreshToken = vi.fn().mockResolvedValue(true);
    manager.refreshAccessToken = vi.fn().mockRejectedValue(new Error('invalid_grant'));

    const result = await manager.authenticate();
    expect(result.success).toBe(false);
    expect(manager.authenticateInteractive).not.toHaveBeenCalled();
  });

  it('interactive mode still falls through to the browser flow', async () => {
    delete process.env.MCP_OUTLOOK_AUTH_MODE;
    manager.tokenManager.isAuthenticated = vi.fn().mockResolvedValue(false);
    manager.tokenManager.hasRefreshToken = vi.fn().mockResolvedValue(false);

    await manager.authenticate();
    expect(manager.authenticateInteractive).toHaveBeenCalled();
  });
});
