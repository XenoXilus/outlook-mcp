import { describe, it, expect, vi } from 'vitest';
import http from 'http';
import { OutlookAuthManager } from '../../auth/auth.js';

/**
 * The OAuth redirect arrives with the browser's entire localhost cookie jar —
 * cookies are domain-scoped, not port-scoped — so session cookies from any
 * local dev server ride along and can exceed Node's 16 KB default header
 * budget. Node then rejects the callback with 431 before our handler runs,
 * and consent never completes.
 */
describe('OAuth callback server header budget', () => {
  it('accepts a callback carrying an oversized localhost cookie jar', async () => {
    const manager = new OutlookAuthManager('client-id', 'tenant-id');
    manager.openBrowser = vi.fn();

    const pending = manager.getAuthorizationCode('test-challenge');
    pending.catch(() => {}); // settled below; avoid unhandled-rejection noise

    await vi.waitFor(() => {
      if (!manager.lastUsedPort) throw new Error('callback server not listening yet');
    });

    // 32 KB of cookies: over Node's 16 KB default, under the widened budget.
    const bigCookie = 'session=' + 'x'.repeat(32 * 1024);
    const status = await new Promise((resolve, reject) => {
      const req = http.request({
        host: '127.0.0.1',
        port: manager.lastUsedPort,
        path: '/callback?code=abc&state=wrong-on-purpose',
        headers: { Cookie: bigCookie },
      }, res => { res.resume(); resolve(res.statusCode); });
      req.on('error', reject);
      req.end();
    });

    // Reaching the handler yields the state-mismatch 400 (which also closes
    // the server); the pre-fix failure mode is Node's own 431 rejection.
    expect(status).toBe(400);
    await expect(pending).rejects.toSatisfy(
      e => JSON.stringify(e).includes('State mismatch')
    );
  });
});
