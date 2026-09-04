import { describe, it, expect } from 'vitest';
import fs from 'fs';
import path from 'path';

// Resolved against the repo root, not the CWD, so the guard scans the same
// files however the suite is invoked.
const REPO_ROOT = path.resolve(import.meta.dirname, '../../..');

// Calendar and SharePoint stay personal in v1.3 by design.
const GUARDED_DIRS = ['server/tools/email', 'server/tools/folders', 'server/tools/attachments', 'server/tools/receipts', 'server/graph'];
const ALLOWED = new Set(['graphHelpers.js']); // owns the '/me' fallback itself

describe('no hardcoded /me mail paths in mailbox-aware tool files', () => {
  for (const dir of GUARDED_DIRS) {
    const absDir = path.join(REPO_ROOT, dir);
    for (const file of fs.readdirSync(absDir).filter(f => f.endsWith('.js'))) {
      if (ALLOWED.has(file)) continue;
      it(`${dir}/${file}`, () => {
        const src = fs.readFileSync(path.join(absDir, file), 'utf8');
        const hits = src.match(/['"`]\/me\//g) || [];
        expect(hits, `hardcoded /me path in ${file} — use getMailboxBase(args.mailbox)`).toEqual([]);
      });
    }
  }
});
