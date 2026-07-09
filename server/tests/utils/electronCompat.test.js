import { describe, it, expect } from 'vitest';
import { execFileSync } from 'child_process';
import { fileURLToPath } from 'url';
import path from 'path';

// repo root (tests live in server/tests/utils)
const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../../..');

// Claude Desktop runs DXT servers in Electron's utilityProcess, where
// process.versions.electron is set and process.type === 'utility'. pdfjs-dist
// (via officeparser) then takes its "browser" path and dereferences the
// DOMMatrix global at import time. Fake that environment in a plain-node
// subprocess to reproduce.
function runInFakeElectron(code) {
  return execFileSync(process.execPath, ['--input-type=module', '-e', `
    Object.defineProperty(process.versions, 'electron', { value: '42.5.1', configurable: true });
    process.type = 'utility';
    ${code}
  `], { cwd: repoRoot, encoding: 'utf8', timeout: 25000 });
}

describe('electronCompat (Claude Desktop utilityProcess)', () => {
  it('repro: officeparser/pdfjs dies on DOMMatrix under the Electron env without the shim', () => {
    let stderr = '';
    try {
      runInFakeElectron(`await import('officeparser'); console.log('loaded');`);
    } catch (error) {
      stderr = String(error.stderr || error.message);
    }
    expect(stderr).toMatch(/DOMMatrix is not defined/);
  });

  it('boots the full tool graph under the Electron env once the shim is loaded first', () => {
    const out = runInFakeElectron(`
      await import('./server/utils/electronCompat.js');
      await import('./server/tools/index.js');
      console.log('TOOLS_LOADED');
    `);
    expect(out).toContain('TOOLS_LOADED');
  });

  it('provides a working minimal DOMMatrix', async () => {
    const { DOMMatrixShim } = await import('../../utils/electronCompat.js');
    const identity = new DOMMatrixShim();
    expect([identity.a, identity.b, identity.c, identity.d, identity.e, identity.f]).toEqual([1, 0, 0, 1, 0, 0]);

    const moved = new DOMMatrixShim().translate(10, 5).scale(2);
    expect([moved.a, moved.d, moved.e, moved.f]).toEqual([2, 2, 10, 5]);

    const point = moved.transformPoint({ x: 1, y: 1 });
    expect([point.x, point.y]).toEqual([12, 7]);
  });
});
