import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { loadReceiptRules } from '../../tools/receipts/receiptRules.js';

describe('loadReceiptRules', () => {
  let dir, saved;

  beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'rules-'));
    saved = process.env.RECEIPT_RULES_PATH;
  });

  afterEach(() => {
    if (saved === undefined) delete process.env.RECEIPT_RULES_PATH;
    else process.env.RECEIPT_RULES_PATH = saved;
    fs.rmSync(dir, { recursive: true, force: true });
  });

  function writeRules(content) {
    const p = path.join(dir, 'rules.json');
    fs.writeFileSync(p, typeof content === 'string' ? content : JSON.stringify(content));
    process.env.RECEIPT_RULES_PATH = p;
    return p;
  }

  it('returns empty rules when RECEIPT_RULES_PATH is unset', () => {
    delete process.env.RECEIPT_RULES_PATH;
    expect(loadReceiptRules()).toEqual({ vendorSenders: [], productLabels: [] });
  });

  it('returns empty rules when RECEIPT_RULES_PATH is blank', () => {
    process.env.RECEIPT_RULES_PATH = '   ';
    expect(loadReceiptRules()).toEqual({ vendorSenders: [], productLabels: [] });
  });

  it('loads and compiles a valid rules file (case-insensitive regexes)', () => {
    writeRules({
      vendorSenders: [{ pattern: 'billing@acme\\.example$', vendor: 'Acme' }],
      productLabels: ['Pro plan[^\\n<,]*']
    });
    const rules = loadReceiptRules();
    expect(rules.vendorSenders).toHaveLength(1);
    expect(rules.vendorSenders[0].vendor).toBe('Acme');
    expect(rules.vendorSenders[0].pattern.test('BILLING@ACME.EXAMPLE')).toBe(true);
    expect(rules.productLabels[0].test('pro plan - 5x')).toBe(true);
  });

  it('treats omitted sections as empty', () => {
    writeRules({ vendorSenders: [{ pattern: 'x', vendor: 'X' }] });
    expect(loadReceiptRules().productLabels).toEqual([]);
    writeRules({ productLabels: ['y'] });
    expect(loadReceiptRules().vendorSenders).toEqual([]);
  });

  it('throws naming the path when the file is missing', () => {
    process.env.RECEIPT_RULES_PATH = path.join(dir, 'nope.json');
    expect(() => loadReceiptRules()).toThrow(/RECEIPT_RULES_PATH \(.*nope\.json\).*cannot read/i);
  });

  it('throws on invalid JSON', () => {
    writeRules('{ not json');
    expect(() => loadReceiptRules()).toThrow(/invalid JSON/i);
  });

  it('throws on wrong top-level or section shapes', () => {
    writeRules('[]');
    expect(() => loadReceiptRules()).toThrow(/must be a JSON object/i);
    writeRules({ vendorSenders: 'nope' });
    expect(() => loadReceiptRules()).toThrow(/vendorSenders must be an array/i);
    writeRules({ vendorSenders: [{ vendor: 'X' }] });
    expect(() => loadReceiptRules()).toThrow(/pattern.*vendor.*strings/i);
    writeRules({ productLabels: [42] });
    expect(() => loadReceiptRules()).toThrow(/productLabels must be an array of strings/i);
  });

  it('throws on an invalid regex, naming the pattern', () => {
    writeRules({ vendorSenders: [{ pattern: '(unclosed', vendor: 'X' }] });
    expect(() => loadReceiptRules()).toThrow(/invalid regex.*\(unclosed/i);
  });
});
