/**
 * Builds small, valid PDFs in memory for tests (fictional content only).
 *
 * makePdf({ pages: [['line 1', 'line 2'], ['page 2 line']] }) lays each string out
 * on its own baseline in Helvetica. makePdf({ pages: [...], streams: [rawOps] })
 * uses the given raw content-stream operators for those pages instead, so a test
 * can build pathological documents (huge or highly compressible streams).
 * Streams are Flate-compressed by default; a correct xref table is always written.
 */

import zlib from 'zlib';

const escapePdfString = (s) => String(s).replace(/[\\()]/g, (c) => `\\${c}`);

/** Content-stream operators that draw each line on its own baseline. */
export function textOps(lines, { top = 760, leading = 14 } = {}) {
  const rows = lines.map((line, i) =>
    `1 0 0 1 50 ${top - i * leading} Tm (${escapePdfString(line)}) Tj`);
  return `BT /F1 10 Tf\n${rows.join('\n')}\nET`;
}

/**
 * @param {{ pages?: string[][], streams?: (string|Buffer)[], compress?: boolean }} spec
 * @returns {Buffer}
 */
export function makePdf({ pages = [], streams = [], compress = true } = {}) {
  const contents = streams.length ? streams : pages.map((lines) => textOps(lines));
  const objects = []; // index i holds object number i + 1
  const add = (body) => { objects.push(body); return objects.length; };

  const catalog = add(null);
  const pageTree = add(null);
  const font = add(Buffer.from('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'));
  const kids = [];
  for (const content of contents) {
    const raw = Buffer.isBuffer(content) ? content : Buffer.from(content, 'latin1');
    const data = compress ? zlib.deflateSync(raw, { level: 9 }) : raw;
    const filter = compress ? ' /Filter /FlateDecode' : '';
    const stream = add(Buffer.concat([
      Buffer.from(`<< /Length ${data.length}${filter} >>\nstream\n`, 'latin1'),
      data,
      Buffer.from('\nendstream', 'latin1')
    ]));
    kids.push(add(Buffer.from(
      `<< /Type /Page /Parent ${pageTree} 0 R /MediaBox [0 0 612 792] ` +
      `/Resources << /Font << /F1 ${font} 0 R >> >> /Contents ${stream} 0 R >>`, 'latin1')));
  }
  objects[catalog - 1] = Buffer.from(`<< /Type /Catalog /Pages ${pageTree} 0 R >>`, 'latin1');
  objects[pageTree - 1] = Buffer.from(
    `<< /Type /Pages /Kids [${kids.map((k) => `${k} 0 R`).join(' ')}] /Count ${kids.length} >>`, 'latin1');

  const chunks = [Buffer.from('%PDF-1.4\n%\xE2\xE3\xCF\xD3\n', 'latin1')];
  let offset = chunks[0].length;
  const offsets = [];
  objects.forEach((body, i) => {
    const chunk = Buffer.concat([Buffer.from(`${i + 1} 0 obj\n`, 'latin1'), body, Buffer.from('\nendobj\n', 'latin1')]);
    offsets.push(offset);
    offset += chunk.length;
    chunks.push(chunk);
  });
  const xref = [
    'xref',
    `0 ${objects.length + 1}`,
    '0000000000 65535 f ',
    ...offsets.map((o) => `${String(o).padStart(10, '0')} 00000 n `),
    'trailer',
    `<< /Size ${objects.length + 1} /Root ${catalog} 0 R >>`,
    'startxref',
    String(offset),
    '%%EOF',
    ''
  ].join('\n');
  chunks.push(Buffer.from(xref, 'latin1'));
  return Buffer.concat(chunks);
}

/**
 * A ~100 KB PDF whose single page inflates to ~33 MB of text operators
 * (`repeats` pairs of moves and shows). Parsing it keeps pdfjs busy for seconds
 * and allocates hundreds of MB — the shape of a hostile attachment.
 */
export function hostilePdf(repeats = 1000000) {
  const ops = `BT /F1 10 Tf\n50 700 Td\n${'0 -12 Td (AB) Tj\n0 12 Td (CD) Tj\n'.repeat(repeats)}ET`;
  return makePdf({ streams: [ops] });
}
