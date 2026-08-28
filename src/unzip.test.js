import { describe, expect, it } from 'vitest';
import zlib from 'node:zlib';
import { inflateRaw, readZip, docxText, pptxText, xlsxText, officeText } from './unzip.js';

/**
 * The decompressor is tested against Node's own zlib rather than against
 * fixtures I generated with it and then froze. zlib produces the compressed
 * bytes, this file decodes them, and the two are compared — so the test fails
 * if my decoder is wrong, not if my fixture is stale.
 *
 * zlib is available here and blocked inside n8n's Code node, which is the whole
 * reason this module exists.
 */

/** @param {string|Uint8Array} s @param {number} [level] */
const raw = (s, level) =>
  new Uint8Array(zlib.deflateRawSync(typeof s === 'string' ? Buffer.from(s, 'utf8') : Buffer.from(s), { level }));

/** @param {Uint8Array} u8 */
const text = (u8) => new TextDecoder('utf-8').decode(u8);

describe('inflateRaw decodes what zlib produces', () => {
  it('handles a short string', () => {
    expect(text(inflateRaw(raw('hello crossarm workbook')))).toBe('hello crossarm workbook');
  });

  it('handles a stored block (level 0, no compression)', () => {
    const s = 'Momentum Room, Thursday.';
    expect(text(inflateRaw(raw(s, 0)))).toBe(s);
  });

  it('handles fixed Huffman (tiny input, level 9)', () => {
    const s = 'a';
    expect(text(inflateRaw(raw(s, 9)))).toBe(s);
  });

  // Long repetitive text forces dynamic Huffman plus long back-references,
  // which is where a hand-written inflate usually breaks.
  it('handles dynamic Huffman with long back-references', () => {
    const s = ('SAIFI CAIDI ETR O&M crossarm capitalization. ').repeat(400);
    expect(text(inflateRaw(raw(s, 9)))).toBe(s);
  });

  it('handles binary data with every byte value', () => {
    const bytes = new Uint8Array(256 * 40);
    for (let i = 0; i < bytes.length; i += 1) bytes[i] = i % 256;
    expect(Array.from(inflateRaw(raw(bytes)))).toEqual(Array.from(bytes));
  });

  it('handles data larger than the initial output buffer', () => {
    const s = 'x'.repeat(200_000);
    expect(inflateRaw(raw(s)).length).toBe(200_000);
  });

  it('handles an empty payload', () => {
    expect(inflateRaw(raw('')).length).toBe(0);
  });

  it('round-trips a hundred pseudo-random payloads', () => {
    // Deterministic PRNG so a failure is reproducible.
    let seed = 12345;
    const next = () => {
      seed = (seed * 1103515245 + 12345) & 0x7fffffff;
      return seed;
    };
    for (let n = 0; n < 100; n += 1) {
      const len = next() % 3000;
      const bytes = new Uint8Array(len);
      for (let i = 0; i < len; i += 1) bytes[i] = next() % 256;
      expect(Array.from(inflateRaw(raw(bytes)))).toEqual(Array.from(bytes));
    }
  });

  it('throws rather than hanging on corrupt input', () => {
    expect(() => inflateRaw(new Uint8Array([0xff, 0xff, 0xff, 0xff]))).toThrow();
  });
});

/**
 * Build a real ZIP in memory. Office files are ZIPs, so if this reader handles
 * a genuine archive it handles them.
 * @param {Array<{name: string, body: string, store?: boolean}>} entries
 */
function makeZip(entries) {
  const chunks = [];
  const central = [];
  let offset = 0;

  for (const e of entries) {
    const nameBuf = Buffer.from(e.name, 'utf8');
    const rawBuf = Buffer.from(e.body, 'utf8');
    const comp = e.store ? rawBuf : zlib.deflateRawSync(rawBuf);
    const method = e.store ? 0 : 8;

    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(method, 8);
    local.writeUInt32LE(0, 14); // crc, not checked by the reader
    local.writeUInt32LE(comp.length, 18);
    local.writeUInt32LE(rawBuf.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);

    const cd = Buffer.alloc(46);
    cd.writeUInt32LE(0x02014b50, 0);
    cd.writeUInt16LE(20, 4);
    cd.writeUInt16LE(20, 6);
    cd.writeUInt16LE(0, 8);
    cd.writeUInt16LE(method, 10);
    cd.writeUInt32LE(0, 16);
    cd.writeUInt32LE(comp.length, 20);
    cd.writeUInt32LE(rawBuf.length, 24);
    cd.writeUInt16LE(nameBuf.length, 28);
    cd.writeUInt32LE(offset, 42);
    central.push(Buffer.concat([cd, nameBuf]));

    chunks.push(local, nameBuf, comp);
    offset += local.length + nameBuf.length + comp.length;
  }

  const cdBuf = Buffer.concat(central);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(cdBuf.length, 12);
  eocd.writeUInt32LE(offset, 16);

  return new Uint8Array(Buffer.concat([...chunks, cdBuf, eocd]));
}

describe('readZip walks a real archive', () => {
  it('reads only the entries asked for', () => {
    const zip = makeZip([
      { name: 'word/document.xml', body: '<w:t>kept</w:t>' },
      { name: 'word/media/image1.png', body: 'x'.repeat(50_000) },
    ]);
    const got = readZip(zip, (n) => n === 'word/document.xml');
    expect(Object.keys(got)).toEqual(['word/document.xml']);
  });

  it('reads stored entries as well as deflated ones', () => {
    const zip = makeZip([{ name: 'a.txt', body: 'stored content', store: true }]);
    expect(text(readZip(zip, () => true)['a.txt'] || new Uint8Array())).toBe('stored content');
  });

  it('rejects something that is not a ZIP', () => {
    expect(() => readZip(new Uint8Array(100), () => true)).toThrow(/not a ZIP/);
  });
});

describe('Office documents', () => {
  /** @param {string} body */
  const docx = (body) => makeZip([{ name: 'word/document.xml', body }]);

  it('reads paragraphs out of a .docx', () => {
    const bytes = docx(
      '<?xml version="1.0"?><w:document><w:body>' +
        '<w:p><w:r><w:t>Crossarm capitalization</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t>2025 spend was </w:t></w:r><w:r><w:t>$6.4M</w:t></w:r></w:p>' +
        '</w:body></w:document>',
    );
    const out = docxText(bytes);
    expect(out).toContain('Crossarm capitalization');
    // Runs split mid-sentence by formatting must rejoin, or figures break apart.
    expect(out).toContain('2025 spend was $6.4M');
  });

  it('decodes XML entities rather than leaking them into the prompt', () => {
    const out = docxText(docx('<w:p><w:r><w:t>O&amp;M &lt;5% &quot;target&quot;</w:t></w:r></w:p>'));
    expect(out).toBe('O&M <5% "target"');
  });

  it('reads slides from a .pptx in slide order, not archive order', () => {
    /** @param {number} n @param {string} t */
    const slide = (n, t) => ({
      name: 'ppt/slides/slide' + n + '.xml',
      body: '<p:sld><p:cSld><a:p><a:r><a:t>' + t + '</a:t></a:r></a:p></p:cSld></p:sld>',
    });
    // Deliberately out of order in the archive.
    const bytes = makeZip([slide(10, 'Tenth'), slide(2, 'Second'), slide(1, 'First')]);
    const out = pptxText(bytes);
    expect(out.indexOf('First')).toBeLessThan(out.indexOf('Second'));
    expect(out.indexOf('Second')).toBeLessThan(out.indexOf('Tenth'));
    expect(out).toContain('[Slide 1]');
  });

  it('routes by extension and returns empty for anything else', () => {
    const bytes = docx('<w:p><w:r><w:t>hi</w:t></w:r></w:p>');
    expect(officeText(bytes, 'Report.DOCX')).toBe('hi');
    expect(officeText(bytes, 'notes.txt')).toBe('');
  });

  it('returns empty rather than throwing when the document part is missing', () => {
    expect(docxText(makeZip([{ name: 'word/settings.xml', body: '<x/>' }]))).toBe('');
    expect(pptxText(makeZip([{ name: 'ppt/presentation.xml', body: '<x/>' }]))).toBe('');
  });
});

describe('spreadsheets', () => {
  /** A minimal but genuine .xlsx: shared strings, workbook, one sheet. */
  /** @param {string} rows @param {string} [sheetName] */
  const book = (rows, sheetName) =>
    makeZip([
      {
        name: 'xl/sharedStrings.xml',
        body: '<sst>' + ['Feeder', 'CAIDI', 'Crossarm &amp; pole'].map((t) => '<si><t>' + t + '</t></si>').join('') + '</sst>',
      },
      { name: 'xl/workbook.xml', body: '<workbook><sheets><sheet name="' + (sheetName || 'Reliability') + '" sheetId="1"/></sheets></workbook>' },
      { name: 'xl/worksheets/sheet1.xml', body: '<worksheet><sheetData>' + rows + '</sheetData></worksheet>' },
    ]);

  it('keeps rows intact so a figure stays with its label', () => {
    const out = xlsxText(
      book(
        '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>' +
          '<row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>88.4</v></c></row>',
      ),
    );
    expect(out).toContain('Feeder | CAIDI');
    // The number must stay on the same line as what it describes.
    expect(out).toContain('Crossarm & pole | 88.4');
  });

  it('names the sheet', () => {
    expect(xlsxText(book('<row r="1"><c r="A1" t="s"><v>0</v></c></row>', 'Q3 Tracker'))).toContain('[Q3 Tracker]');
  });

  it('reads inline strings as well as the shared table', () => {
    const out = xlsxText(book('<row r="1"><c r="A1" t="inlineStr"><is><t>Inline value</t></is></c></row>'));
    expect(out).toContain('Inline value');
  });

  it('drops rows that are only formatting', () => {
    const out = xlsxText(
      book('<row r="1"><c r="A1" t="s"><v>0</v></c></row><row r="2"><c r="A2"/><c r="B2"/></row>'),
    );
    expect(out.split('\n').filter((l) => l && !l.startsWith('['))).toHaveLength(1);
  });

  it('stops after the row cap rather than dumping a whole tracker', () => {
    const rows = Array.from({ length: 500 }, (_, i) => '<row r="' + (i + 1) + '"><c r="A' + (i + 1) + '"><v>' + i + '</v></c></row>').join('');
    const out = xlsxText(book(rows), { maxRows: 10 });
    expect(out.split('\n').filter((l) => l && !l.startsWith('['))).toHaveLength(10);
  });

  it('is reachable through officeText by extension', () => {
    const bytes = book('<row r="1"><c r="A1" t="s"><v>0</v></c></row>');
    expect(officeText(bytes, 'Workbook.XLSX')).toContain('Feeder');
  });
});
