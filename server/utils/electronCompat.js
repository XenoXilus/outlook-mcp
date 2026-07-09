/**
 * Compatibility shims for Electron utilityProcess (Claude Desktop's built-in
 * Node host for DXT extensions).
 *
 * pdfjs-dist (pulled in by officeparser) classifies any Electron process as a
 * browser — its isNodeJS check fails when process.versions.electron is set and
 * process.type !== 'browser' — and then evaluates `new DOMMatrix()` at module
 * scope. Electron utility processes have no DOMMatrix global, so the whole
 * server dies at startup with "ReferenceError: DOMMatrix is not defined"
 * (a fatal unhandledRejection in the MCP node host).
 *
 * This module must be imported before anything that imports officeparser.
 * It defines a minimal 2D-affine DOMMatrix only when the global is missing;
 * real browsers and plain Node (where pdfjs takes its own Node path) are
 * untouched.
 */

export class DOMMatrixShim {
  constructor(init) {
    this.a = 1; this.b = 0; this.c = 0; this.d = 1; this.e = 0; this.f = 0;
    if (Array.isArray(init) && init.length === 6) {
      [this.a, this.b, this.c, this.d, this.e, this.f] = init;
    } else if (init && typeof init === 'object') {
      const { a = 1, b = 0, c = 0, d = 1, e = 0, f = 0 } = init;
      Object.assign(this, { a, b, c, d, e, f });
    }
  }

  multiply(other) {
    const m = other instanceof DOMMatrixShim ? other : new DOMMatrixShim(other);
    return new DOMMatrixShim([
      this.a * m.a + this.c * m.b,
      this.b * m.a + this.d * m.b,
      this.a * m.c + this.c * m.d,
      this.b * m.c + this.d * m.d,
      this.a * m.e + this.c * m.f + this.e,
      this.b * m.e + this.d * m.f + this.f
    ]);
  }

  multiplySelf(other) {
    const result = this.multiply(other);
    Object.assign(this, result);
    return this;
  }

  translate(tx = 0, ty = 0) {
    return this.multiply(new DOMMatrixShim([1, 0, 0, 1, tx, ty]));
  }

  scale(sx = 1, sy = sx) {
    return this.multiply(new DOMMatrixShim([sx, 0, 0, sy, 0, 0]));
  }

  invertSelf() {
    const det = this.a * this.d - this.b * this.c;
    if (det === 0) {
      Object.assign(this, { a: NaN, b: NaN, c: NaN, d: NaN, e: NaN, f: NaN });
      return this;
    }
    const { a, b, c, d, e, f } = this;
    Object.assign(this, {
      a: d / det,
      b: -b / det,
      c: -c / det,
      d: a / det,
      e: (c * f - d * e) / det,
      f: (b * e - a * f) / det
    });
    return this;
  }

  transformPoint(point = { x: 0, y: 0 }) {
    const { x = 0, y = 0 } = point;
    return {
      x: this.a * x + this.c * y + this.e,
      y: this.b * x + this.d * y + this.f,
      z: 0,
      w: 1
    };
  }
}

if (typeof globalThis.DOMMatrix === 'undefined') {
  globalThis.DOMMatrix = DOMMatrixShim;
}
