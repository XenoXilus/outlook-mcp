/**
 * Receipt vendor-rules config loader.
 *
 * RECEIPT_RULES_PATH points at a JSON file supplying site-specific matching
 * rules; the parser's built-in heuristics stay fully generic. Schema:
 *   { "vendorSenders": [{"pattern": "<regex>", "vendor": "<name>"}],
 *     "productLabels": ["<regex>", ...] }
 * Unset/blank env => empty rules (generic heuristics only). Set but
 * missing/unreadable/malformed => hard error: a scheduled run silently
 * downgrading to weaker rules would produce a wrong manifest.
 */

import fs from 'fs';

function fail(rulesPath, detail) {
  throw new Error(`RECEIPT_RULES_PATH (${rulesPath}): ${detail}`);
}

function compile(rulesPath, source) {
  try {
    return new RegExp(source, 'i');
  } catch (error) {
    fail(rulesPath, `invalid regex '${source}' — ${error.message}`);
  }
}

export function loadReceiptRules(env = process.env) {
  const rulesPath = (env.RECEIPT_RULES_PATH || '').trim();
  if (!rulesPath) return { vendorSenders: [], productLabels: [] };

  let raw;
  try {
    raw = fs.readFileSync(rulesPath, 'utf8');
  } catch (error) {
    fail(rulesPath, `cannot read file — ${error.message}`);
  }

  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (error) {
    fail(rulesPath, `invalid JSON — ${error.message}`);
  }

  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    fail(rulesPath, 'must be a JSON object with optional vendorSenders/productLabels arrays');
  }

  const { vendorSenders = [], productLabels = [] } = parsed;
  if (!Array.isArray(vendorSenders)) fail(rulesPath, 'vendorSenders must be an array');
  if (!Array.isArray(productLabels) || productLabels.some(p => typeof p !== 'string')) {
    fail(rulesPath, 'productLabels must be an array of strings');
  }

  const senders = vendorSenders.map((entry, i) => {
    if (typeof entry?.pattern !== 'string' || typeof entry?.vendor !== 'string') {
      fail(rulesPath, `vendorSenders[${i}] needs pattern and vendor strings`);
    }
    return { pattern: compile(rulesPath, entry.pattern), vendor: entry.vendor };
  });

  return {
    vendorSenders: senders,
    productLabels: productLabels.map(p => compile(rulesPath, p))
  };
}
