import { describe, it, expect } from 'vitest';
import { safeStringify } from '../../utils/jsonUtils.js';

describe('safeStringify', () => {
  describe('Secret key redaction', () => {
    it('redacts top-level access_token value', () => {
      const result = safeStringify({ access_token: 'abc' });
      expect(result).not.toContain('abc');
      expect(result).toContain('[REDACTED]');
    });

    it('redacts nested Authorization header value', () => {
      const result = safeStringify({ headers: { Authorization: 'Bearer xyz' } });
      expect(result).not.toContain('xyz');
      expect(result).toContain('[REDACTED]');
    });

    it('redacts keys case-insensitively — Authorization and REFRESH_TOKEN', () => {
      const obj = {
        Authorization: 'Bearer token123',
        REFRESH_TOKEN: 'some-refresh-value'
      };
      const result = safeStringify(obj);
      expect(result).not.toContain('token123');
      expect(result).not.toContain('some-refresh-value');
      expect(result).toContain('[REDACTED]');
    });

    it('redacts password at arbitrary nesting depth', () => {
      const obj = { user: { profile: { password: 'hunter2' } } };
      const result = safeStringify(obj);
      expect(result).not.toContain('hunter2');
      expect(result).toContain('[REDACTED]');
    });

    it('redacts client_secret value', () => {
      const result = safeStringify({ client_secret: 'my-secret-value' });
      expect(result).not.toContain('my-secret-value');
      expect(result).toContain('[REDACTED]');
    });

    it('redacts cookie value', () => {
      const result = safeStringify({ cookie: 'session=abc123; path=/' });
      expect(result).not.toContain('session=abc123');
      expect(result).toContain('[REDACTED]');
    });

    it('redacts api_key value', () => {
      const result = safeStringify({ api_key: 'key-12345' });
      expect(result).not.toContain('key-12345');
      expect(result).toContain('[REDACTED]');
    });

    it('redacts object values under a secret key wholesale', () => {
      const obj = { token: { inner: 'sensitive', other: 'also-sensitive' } };
      const result = safeStringify(obj);
      expect(result).not.toContain('sensitive');
      expect(result).not.toContain('also-sensitive');
      expect(result).toContain('[REDACTED]');
    });

    it('redacts hyphenated api-key (api[-_]?key regex)', () => {
      const result = safeStringify({ 'api-key': 'k' });
      expect(result).not.toContain('"k"');
      expect(result).toContain('[REDACTED]');
    });

    it('redacts array values under a secret key wholesale (no recursion into the array)', () => {
      const result = safeStringify({ tokens: ['a', 'b'] });
      expect(result).not.toContain('"a"');
      expect(result).not.toContain('"b"');
      expect(result).toContain('[REDACTED]');
    });
  });

  describe('Non-secret keys are NOT redacted', () => {
    it('does not redact error.code diagnostic strings', () => {
      const obj = {
        error: {
          code: 'ErrorInvalidRequest',
          message: 'The request is invalid'
        }
      };
      const result = safeStringify(obj);
      expect(result).toContain('ErrorInvalidRequest');
      expect(result).not.toContain('[REDACTED]');
    });

    it('does not redact statusCode', () => {
      const obj = { statusCode: 400, message: 'Bad request' };
      const result = safeStringify(obj);
      expect(result).toContain('400');
      expect(result).not.toContain('[REDACTED]');
    });

    it('does not redact errorCode', () => {
      const obj = { errorCode: 'INVALID_PARAMS', detail: 'Missing field' };
      const result = safeStringify(obj);
      expect(result).toContain('INVALID_PARAMS');
      expect(result).not.toContain('[REDACTED]');
    });

    it('passes non-secret values through unchanged', () => {
      const obj = {
        name: 'John Doe',
        email: 'john@example.com',
        requestId: 'req-123',
        statusCode: 200
      };
      const result = safeStringify(obj);
      expect(result).toContain('John Doe');
      expect(result).toContain('john@example.com');
      expect(result).toContain('req-123');
      expect(result).toContain('200');
      expect(result).not.toContain('[REDACTED]');
    });
  });

  describe('Existing behaviors preserved', () => {
    it('handles circular references', () => {
      const obj = { name: 'test' };
      obj.self = obj;
      const result = safeStringify(obj);
      expect(result).toContain('[Circular Reference]');
    });

    it('converts undefined values to null', () => {
      const obj = { defined: 'value', missing: undefined };
      const result = safeStringify(obj);
      const parsed = JSON.parse(result);
      expect(parsed.defined).toBe('value');
      expect(parsed.missing).toBeNull();
    });

    it('serializes functions as [Function]', () => {
      const obj = { fn: () => 'hello', name: 'test' };
      const result = safeStringify(obj);
      expect(result).toContain('[Function]');
    });

    it('serializes BigInt as string', () => {
      // Use a BigInt literal (n suffix) to avoid JS Number precision loss
      const obj = { big: 9007199254740993n };
      const result = safeStringify(obj);
      expect(result).toContain('9007199254740993');
    });

    it('accepts a space parameter for indentation', () => {
      const obj = { key: 'value' };
      const result0 = safeStringify(obj, 0);
      const result4 = safeStringify(obj, 4);
      expect(result0).toBe('{"key":"value"}');
      expect(result4).toMatch(/\n {4}/);
    });

    it('handles null values without issue', () => {
      const obj = { a: null, b: 'ok' };
      const result = safeStringify(obj);
      const parsed = JSON.parse(result);
      expect(parsed.a).toBeNull();
      expect(parsed.b).toBe('ok');
    });
  });
});
