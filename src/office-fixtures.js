/**
 * Office, PDF and image fixtures, built rather than pasted.
 *
 * The seeder needs attachments to put in front of WF-D, and the obvious way to
 * get them is to paste base64 of a real file into the workflow. That was tried:
 * a .docx went in with the right length and a valid ZIP signature and inflated
 * to nothing, and it read exactly like a broken decoder rather than a broken
 * fixture. Base64 in a source file is unreadable, unreviewable, and wrong in
 * ways nothing catches.
 *
 * These are built instead — a few hundred bytes each, from text you can read.
 * Every one is round-tripped through the real extractors in
 * `src/office-fixtures.test.js`, so a fixture that would not parse fails the
 * suite rather than the sandbox.
 *
 * Office files are ZIPs. These use the stored method (no compression), which
 * `readZip` handles, so no compressor is needed to produce one.
 */

const CRC_TABLE = (() => {
  const table = new Int32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c;
  }
  return table;
})();

/**
 * @param {Uint8Array} bytes
 * @returns {number} unsigned CRC-32
 */
function crc32(bytes) {
  let c = 0 ^ -1;
  for (let i = 0; i < bytes.length; i += 1) {
    c = (c >>> 8) ^ (CRC_TABLE[(c ^ (bytes[i] ?? 0)) & 0xff] ?? 0);
  }
  return (c ^ -1) >>> 0;
}

/** @param {string} text @returns {Uint8Array} */
function utf8(text) {
  return new TextEncoder().encode(text);
}

/**
 * A ZIP archive with every entry stored uncompressed.
 *
 * An entry carries either `text`, which is encoded as UTF-8, or `bytes` for a
 * part that is not text at all — a `.pbix` nests a whole second ZIP inside its
 * DataMashup part, and that cannot be expressed as a string.
 *
 * @param {Array<{ name: string, text?: string, bytes?: Uint8Array }>} entries
 * @returns {Uint8Array}
 */
export function zipStored(entries) {
  const parts = [];
  const central = [];
  let offset = 0;

  for (const entry of entries) {
    const name = utf8(entry.name);
    const body = entry.bytes ? entry.bytes : utf8(entry.text || '');
    const sum = crc32(body);

    const local = new Uint8Array(30 + name.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true); // version needed
    lv.setUint16(6, 0, true); // flags
    lv.setUint16(8, 0, true); // method 0 = stored
    lv.setUint16(10, 0, true); // mod time
    lv.setUint16(12, 0, true); // mod date
    lv.setUint32(14, sum, true);
    lv.setUint32(18, body.length, true);
    lv.setUint32(22, body.length, true);
    lv.setUint16(26, name.length, true);
    lv.setUint16(28, 0, true); // extra
    local.set(name, 30);

    const head = new Uint8Array(46 + name.length);
    const hv = new DataView(head.buffer);
    hv.setUint32(0, 0x02014b50, true);
    hv.setUint16(4, 20, true); // version made by
    hv.setUint16(6, 20, true); // version needed
    hv.setUint16(8, 0, true);
    hv.setUint16(10, 0, true); // stored
    hv.setUint16(12, 0, true);
    hv.setUint16(14, 0, true);
    hv.setUint32(16, sum, true);
    hv.setUint32(20, body.length, true);
    hv.setUint32(24, body.length, true);
    hv.setUint16(28, name.length, true);
    hv.setUint16(30, 0, true); // extra
    hv.setUint16(32, 0, true); // comment
    hv.setUint16(34, 0, true); // disk
    hv.setUint16(36, 0, true); // internal attrs
    hv.setUint32(38, 0, true); // external attrs
    hv.setUint32(42, offset, true);
    head.set(name, 46);

    parts.push(local, body);
    central.push(head);
    offset += local.length + body.length;
  }

  const centralSize = central.reduce((n, c) => n + c.length, 0);
  const end = new Uint8Array(22);
  const ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(4, 0, true);
  ev.setUint16(6, 0, true);
  ev.setUint16(8, entries.length, true);
  ev.setUint16(10, entries.length, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, offset, true);
  ev.setUint16(20, 0, true);

  const all = [...parts, ...central, end];
  const total = all.reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(total);
  let at = 0;
  for (const p of all) {
    out.set(p, at);
    at += p.length;
  }
  return out;
}

/** @param {string} value @returns {string} */
function xmlEscape(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * A .docx whose body is the paragraphs given. Carries [Content_Types].xml so it
 * is a real package rather than only what our own reader happens to need.
 *
 * @param {string[]} paragraphs
 * @returns {Uint8Array}
 */
export function docxFixture(paragraphs) {
  const body = paragraphs
    .map((p) => '<w:p><w:r><w:t>' + xmlEscape(p) + '</w:t></w:r></w:p>')
    .join('');
  return zipStored([
    {
      name: '[Content_Types].xml',
      text:
        '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="xml" ContentType="application/xml"/></Types>',
    },
    {
      name: 'word/document.xml',
      text:
        '<?xml version="1.0" encoding="UTF-8"?>' +
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:body>' + body + '</w:body></w:document>',
    },
  ]);
}

/**
 * A .pptx, one entry per slide, in order.
 *
 * @param {string[][]} slides lines per slide
 * @returns {Uint8Array}
 */
export function pptxFixture(slides) {
  const entries = slides.map((lines, i) => ({
    name: 'ppt/slides/slide' + (i + 1) + '.xml',
    text:
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      lines.map((l) => '<a:p><a:r><a:t>' + xmlEscape(l) + '</a:t></a:r></a:p>').join('') +
      '</p:sld>',
  }));
  return zipStored(entries);
}

/**
 * An .xlsx using shared strings and t="s" cells, which is what Excel actually
 * writes — so the fixture exercises the shared-string lookup rather than the
 * inline shortcut nothing real uses.
 *
 * @param {string} sheetName
 * @param {string[][]} rows
 * @returns {Uint8Array}
 */
export function xlsxFixture(sheetName, rows) {
  /** @type {string[]} */
  const shared = [];
  /** @param {string} value */
  const index = (value) => {
    const at = shared.indexOf(value);
    if (at !== -1) return at;
    shared.push(value);
    return shared.length - 1;
  };

  const COLUMNS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const sheetRows = rows
    .map((cells, r) => {
      const cellXml = cells
        .map((value, c) => {
          const ref = (COLUMNS[c] ?? 'A') + (r + 1);
          return '<c r="' + ref + '" t="s"><v>' + index(value) + '</v></c>';
        })
        .join('');
      return '<row r="' + (r + 1) + '">' + cellXml + '</row>';
    })
    .join('');

  return zipStored([
    {
      name: 'xl/workbook.xml',
      text:
        '<?xml version="1.0" encoding="UTF-8"?>' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<sheets><sheet name="' + xmlEscape(sheetName) + '" sheetId="1" r:id="rId1"/></sheets></workbook>',
    },
    {
      name: 'xl/sharedStrings.xml',
      text:
        '<?xml version="1.0" encoding="UTF-8"?>' +
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' +
        shared.length + '" uniqueCount="' + shared.length + '">' +
        shared.map((s) => '<si><t>' + xmlEscape(s) + '</t></si>').join('') +
        '</sst>',
    },
    {
      name: 'xl/worksheets/sheet1.xml',
      text:
        '<?xml version="1.0" encoding="UTF-8"?>' +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<sheetData>' + sheetRows + '</sheetData></worksheet>',
    },
  ]);
}

/**
 * A structurally complete single-page PDF with a real cross-reference table.
 *
 * PDFs are one of only two kinds that leave the server for OpenAI, so the
 * fixture has to be a file a reader will actually accept rather than something
 * shaped like one.
 *
 * @param {string[]} lines
 * @returns {Uint8Array}
 */
export function pdfFixture(lines) {
  const text = lines
    .map((l, i) => '(' + String(l).replace(/([()\\])/g, '\\$1') + ') Tj 0 -18 Td')
    .join('\n');
  const stream = 'BT /F1 12 Tf 56 760 Td\n' + text + '\nET';

  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
    '<< /Length ' + utf8(stream).length + ' >>\nstream\n' + stream + '\nendstream',
  ];

  let pdf = '%PDF-1.4\n';
  /** @type {number[]} */
  const offsets = [];
  objects.forEach((body, i) => {
    offsets.push(utf8(pdf).length);
    pdf += i + 1 + ' 0 obj\n' + body + '\nendobj\n';
  });

  const xrefAt = utf8(pdf).length;
  pdf += 'xref\n0 ' + (objects.length + 1) + '\n0000000000 65535 f \n';
  for (const at of offsets) pdf += String(at).padStart(10, '0') + ' 00000 n \n';
  pdf +=
    'trailer\n<< /Size ' + (objects.length + 1) + ' /Root 1 0 R >>\nstartxref\n' + xrefAt + '\n%%EOF\n';

  return utf8(pdf);
}

/**
 * A 1x1 PNG. Deliberately tiny: at 68 bytes it sits below the reading
 * threshold, so it proves the triage skips a signature logo instead of paying
 * a model call to be told what a logo is.
 *
 * @returns {Uint8Array}
 */
export function pngFixture() {
  return new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82,
  ]);
}

const B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

/**
 * Base64, written out rather than taken from Buffer, so this module runs
 * anywhere the rest of src/ does.
 *
 * @param {Uint8Array} bytes
 * @returns {string}
 */
export function toBase64(bytes) {
  let out = '';
  for (let i = 0; i < bytes.length; i += 3) {
    const a = bytes[i] ?? 0;
    const b = bytes[i + 1] ?? 0;
    const c = bytes[i + 2] ?? 0;
    const has = bytes.length - i;
    out += B64[a >> 2];
    out += B64[((a & 3) << 4) | (b >> 4)];
    out += has > 1 ? B64[((b & 15) << 2) | (c >> 6)] : '=';
    out += has > 2 ? B64[c & 63] : '=';
  }
  return out;
}

/**
 * A `.pbix` carrying the M queries a real report would carry.
 *
 * The DataMashup part is not a ZIP on its own: it is four bytes of version,
 * four of package length, and then the package — which *is* a ZIP. Building it
 * for real is the only way to prove the reader unwraps it for real.
 *
 * @param {string[]} queries M query lines, e.g. File.Contents("C:\\x.xlsx")
 * @param {{ connections?: string }} [extra]
 * @returns {Uint8Array}
 */
export function pbixFixture(queries, extra) {
  const section = 'section Section1;\n' + (queries || []).join('\n') + '\n';
  const pkg = zipStored([{ name: 'Formulas/Section1.m', text: section }]);

  const mashup = new Uint8Array(8 + pkg.length);
  const view = new DataView(mashup.buffer);
  view.setUint32(0, 0, true); // version
  view.setUint32(4, pkg.length, true); // package length
  mashup.set(pkg, 8);

  const entries = [
    { name: '[Content_Types].xml', text: '<?xml version="1.0"?><Types/>' },
    { name: 'DataMashup', bytes: mashup },
    { name: 'Report/Layout', text: '{"config":"{}"}' },
  ];
  if (extra && extra.connections) entries.push({ name: 'Connections', text: extra.connections });
  return zipStored(entries);
}
