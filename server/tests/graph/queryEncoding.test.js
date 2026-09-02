import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { GraphApiClient } from '../../graph/graphClient.js';

/**
 * Regression tests for OData query-value encoding: the Graph SDK appends
 * $filter/$search values to the URL without percent-encoding, and Graph decodes
 * a literal '+' in the query string as a space. A filter on a plus-addressed
 * sender (invoice+statements@...) therefore silently matched nothing.
 */
describe('GraphApiClient query encoding', () => {
  let client, captured, originalFetch;

  beforeEach(() => {
    captured = [];
    originalFetch = globalThis.fetch;
    globalThis.fetch = vi.fn(async (input) => {
      captured.push(typeof input === 'string' ? input : input.url);
      return new Response(JSON.stringify({ value: [] }), {
        status: 200,
        headers: { 'content-type': 'application/json' }
      });
    });
    const authManager = { tokenManager: { getAccessToken: async () => 'test-token' } };
    client = new GraphApiClient(authManager);
  });

  afterEach(() => {
    globalThis.fetch = originalFetch;
  });

  it('percent-encodes + in $filter so plus-addressed senders survive the wire', async () => {
    await client.makeRequest('/me/messages', {
      filter: "from/emailAddress/address eq 'invoice+statements@billing.example.com'"
    });
    expect(captured).toHaveLength(1);
    expect(captured[0]).toContain('invoice%2Bstatements');
    expect(captured[0]).not.toContain('invoice+statements');
  });

  it('percent-encodes & and # in $filter values', async () => {
    await client.makeRequest('/me/messages', {
      filter: "contains(subject,'Smith & Co #42')"
    });
    expect(captured[0]).toContain('%26');
    expect(captured[0]).toContain('%23');
    // The & must not truncate the query string: $top must survive intact
    const url = new URL(captured[0]);
    expect(url.searchParams.get('$filter')).toBe("contains(subject,'Smith & Co #42')");
  });

  it('percent-encodes % in $filter values without double-encoding', async () => {
    await client.makeRequest('/me/messages', {
      filter: "contains(subject,'100% cotton')"
    });
    const url = new URL(captured[0]);
    expect(url.searchParams.get('$filter')).toBe("contains(subject,'100% cotton')");
  });

  it('percent-encodes + in $search values', async () => {
    await client.makeRequest('/me/messages', { search: '"from:invoice+statements@payments.example.com"' });
    expect(captured[0]).toContain('%2B');
    expect(captured[0]).not.toContain('invoice+statements');
  });

  it('leaves ordinary filters readable and functional', async () => {
    await client.makeRequest('/me/messages', {
      filter: "receivedDateTime ge 2026-08-01T00:00:00Z and contains(subject,'Vandelay')",
      orderby: 'receivedDateTime desc',
      top: 5
    });
    const url = new URL(captured[0]);
    expect(url.searchParams.get('$filter'))
      .toBe("receivedDateTime ge 2026-08-01T00:00:00Z and contains(subject,'Vandelay')");
    expect(url.searchParams.get('$top')).toBe('5');
  });
});
