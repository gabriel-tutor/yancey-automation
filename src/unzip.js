/**
 * A ZIP reader and DEFLATE decompressor, in plain JavaScript.
 *
 * This exists because `.docx` and `.pptx` are ZIP archives and n8n's Code node
 * refuses `require('zlib')` — probed on 2026-08-29, verbatim: "Module 'zlib' is
 * disallowed". `Buffer` is available, everything else is not.
 *
 * Without this, attachment ingestion would cover PDFs, spreadsheets, images and
 * plain text but silently miss decks and Word documents, which in a consulting
 * inbox is most of what actually gets sent. Installing tooling on the client's
 * VPS to avoid ~250 lines of well-understood algorithm would be the worse
 * trade: more attack surface, more to break, on a server running their live
 * mail.
 *
 * The inflate is the classic canonical-Huffman decoder (the "puff" shape). It
 * is deliberately written for clarity over speed — attachments are a handful a
 * day, not a hot loop.
 */

/** Bit-reversal table for building canonical Huffman codes. */
const MAX_BITS = 15;

/**
 * A canonical Huffman decoding table: `counts[n]` is how many symbols use n
 * bits, `symbols` is the symbol list ordered by code.
 * @typedef {{ counts: Int32Array, symbols: Int32Array }} Huffman
 */

/**
 * Read one element of a byte or int array as a number.
 *
 * Every index in this file is bounded by the DEFLATE format itself — table
 * sizes are derived from the header immediately before they are read. TypeScript
 * cannot see that, and `noUncheckedIndexedAccess` would otherwise scatter `| 0`
 * through a decoder where clarity matters more than anywhere else in the
 * codebase. The claim is made once, here, and the 100 randomised round-trips in
 * unzip.test.js are what actually holds it up.
 *
 * @param {Uint8Array | Int32Array} arr
 * @param {number} i
 * @returns {number}
 */
function at(arr, i) {
  const v = arr[i];
  return v === undefined ? 0 : v;
}

/**
 * @param {Uint8Array} lengths code length per symbol
 * @returns {Huffman}
 */
function buildHuffman(lengths) {
  const counts = new Int32Array(MAX_BITS + 1);
  for (let i = 0; i < lengths.length; i += 1) counts[at(lengths, i)] = at(counts, at(lengths, i)) + 1;
  counts[0] = 0;

  const offsets = new Int32Array(MAX_BITS + 1);
  for (let i = 1; i <= MAX_BITS; i += 1) offsets[i] = at(offsets, i - 1) + at(counts, i - 1);

  const symbols = new Int32Array(lengths.length);
  for (let i = 0; i < lengths.length; i += 1) {
    const len = at(lengths, i);
    if (len !== 0) {
      symbols[at(offsets, len)] = i;
      offsets[len] = at(offsets, len) + 1;
    }
  }
  return { counts, symbols };
}

/** Reader state for the bit stream. */
class BitReader {
  /** @param {Uint8Array} bytes */
  constructor(bytes) {
    this.bytes = bytes;
    this.pos = 0;
    this.bitBuf = 0;
    this.bitCnt = 0;
  }

  /**
   * @param {number} need
   * @returns {number}
   */
  bits(need) {
    let val = this.bitBuf;
    while (this.bitCnt < need) {
      if (this.pos >= this.bytes.length) throw new Error('unexpected end of deflate stream');
      val |= at(this.bytes, this.pos) << this.bitCnt;
      this.pos += 1;
      this.bitCnt += 8;
    }
    this.bitBuf = val >>> need;
    this.bitCnt -= need;
    return val & ((1 << need) - 1);
  }

  /**
   * Walk the code tree one bit at a time until a symbol is complete.
   * @param {Huffman} h
   * @returns {number}
   */
  decode(h) {
    let code = 0;
    let first = 0;
    let index = 0;
    for (let len = 1; len <= MAX_BITS; len += 1) {
      code |= this.bits(1);
      const count = at(h.counts, len);
      if (code - count < first) return at(h.symbols, index + (code - first));
      index += count;
      first += count;
      first <<= 1;
      code <<= 1;
    }
    throw new Error('invalid Huffman code');
  }
}

const LENGTH_BASE = [
  3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131,
  163, 195, 227, 258,
];
const LENGTH_EXTRA = [
  0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0,
];
const DIST_BASE = [
  1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049,
  3073, 4097, 6145, 8193, 12289, 16385, 24577,
];
const DIST_EXTRA = [
  0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13,
];
const CLEN_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

/** @type {Huffman | null} */
let fixedLit = null;
/** @type {Huffman | null} */
let fixedDist = null;

function fixedTables() {
  if (fixedLit && fixedDist) return { lit: fixedLit, dist: fixedDist };
  const lengths = new Uint8Array(288);
  for (let i = 0; i < 144; i += 1) lengths[i] = 8;
  for (let i = 144; i < 256; i += 1) lengths[i] = 9;
  for (let i = 256; i < 280; i += 1) lengths[i] = 7;
  for (let i = 280; i < 288; i += 1) lengths[i] = 8;
  fixedLit = buildHuffman(lengths);
  fixedDist = buildHuffman(new Uint8Array(30).fill(5));
  return { lit: fixedLit, dist: fixedDist };
}

/**
 * Decompress a raw DEFLATE stream (no zlib or gzip wrapper).
 *
 * @param {Uint8Array} input
 * @param {number} [expectedSize] hint used to size the output buffer
 * @returns {Uint8Array}
 */
function inflateRaw(input, expectedSize) {
  const reader = new BitReader(input);
  let out = new Uint8Array(expectedSize && expectedSize > 0 ? expectedSize : 1024);
  let outLen = 0;

  /** @param {number} extra */
  function ensure(extra) {
    if (outLen + extra <= out.length) return;
    let size = out.length * 2;
    while (size < outLen + extra) size *= 2;
    const bigger = new Uint8Array(size);
    bigger.set(out.subarray(0, outLen));
    out = bigger;
  }

  let final = 0;
  do {
    final = reader.bits(1);
    const type = reader.bits(2);

    if (type === 0) {
      // Stored: discard the partial byte, then a length/complement pair.
      reader.bitBuf = 0;
      reader.bitCnt = 0;
      if (reader.pos + 4 > input.length) throw new Error('truncated stored block');
      const len = at(input, reader.pos) | (at(input, reader.pos + 1) << 8);
      reader.pos += 4;
      if (reader.pos + len > input.length) throw new Error('stored block overruns input');
      ensure(len);
      out.set(input.subarray(reader.pos, reader.pos + len), outLen);
      outLen += len;
      reader.pos += len;
      continue;
    }

    /** @type {Huffman} */
    let lit;
    /** @type {Huffman} */
    let dist;

    if (type === 1) {
      const t = fixedTables();
      lit = t.lit;
      dist = t.dist;
    } else if (type === 2) {
      const nlen = reader.bits(5) + 257;
      const ndist = reader.bits(5) + 1;
      const ncode = reader.bits(4) + 4;
      const clens = new Uint8Array(19);
      for (let i = 0; i < ncode; i += 1) clens[CLEN_ORDER[i] || 0] = reader.bits(3);
      const clenTable = buildHuffman(clens);

      const lengths = new Uint8Array(nlen + ndist);
      let i = 0;
      while (i < nlen + ndist) {
        const sym = reader.decode(clenTable);
        if (sym < 16) {
          lengths[i] = sym;
          i += 1;
        } else if (sym === 16) {
          if (i === 0) throw new Error('repeat with no previous length');
          const prev = at(lengths, i - 1);
          let repeat = 3 + reader.bits(2);
          while (repeat > 0 && i < lengths.length) {
            lengths[i] = prev;
            i += 1;
            repeat -= 1;
          }
        } else if (sym === 17) {
          i += 3 + reader.bits(3);
        } else {
          i += 11 + reader.bits(7);
        }
      }
      lit = buildHuffman(lengths.subarray(0, nlen));
      dist = buildHuffman(lengths.subarray(nlen));
    } else {
      throw new Error('invalid deflate block type');
    }

    for (;;) {
      const sym = reader.decode(lit);
      if (sym === 256) break;
      if (sym < 256) {
        ensure(1);
        out[outLen] = sym;
        outLen += 1;
        continue;
      }
      const lenIdx = sym - 257;
      if (lenIdx >= LENGTH_BASE.length) throw new Error('invalid length symbol');
      const length = (LENGTH_BASE[lenIdx] || 0) + reader.bits(LENGTH_EXTRA[lenIdx] || 0);
      const distSym = reader.decode(dist);
      if (distSym >= DIST_BASE.length) throw new Error('invalid distance symbol');
      const distance = (DIST_BASE[distSym] || 0) + reader.bits(DIST_EXTRA[distSym] || 0);
      if (distance > outLen) throw new Error('distance runs before the start of output');
      ensure(length);
      let from = outLen - distance;
      for (let n = 0; n < length; n += 1) {
        out[outLen] = at(out, from);
        outLen += 1;
        from += 1;
      }
    }
  } while (!final);

  return out.subarray(0, outLen);
}

/**
 * Read a ZIP archive's entries, walking the central directory rather than
 * scanning local headers — local headers may omit sizes when the streaming bit
 * is set, and the central directory always has them.
 *
 * Only entries matching `wanted` are decompressed, so opening a 40MB deck to
 * read its slide text does not inflate the embedded images.
 *
 * @param {Uint8Array} bytes
 * @param {(name: string) => boolean} wanted
 * @returns {Record<string, Uint8Array>}
 */
function readZip(bytes, wanted) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  // End of central directory: scan back from the tail for the signature.
  let eocd = -1;
  const lowest = Math.max(0, bytes.length - 66_000);
  for (let i = bytes.length - 22; i >= lowest; i -= 1) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd === -1) throw new Error('not a ZIP archive (no end-of-central-directory)');

  const count = view.getUint16(eocd + 10, true);
  let ptr = view.getUint32(eocd + 16, true);

  /** @type {Record<string, Uint8Array>} */
  const out = {};

  for (let n = 0; n < count; n += 1) {
    if (ptr + 46 > bytes.length || view.getUint32(ptr, true) !== 0x02014b50) break;

    const method = view.getUint16(ptr + 10, true);
    const compSize = view.getUint32(ptr + 20, true);
    const rawSize = view.getUint32(ptr + 24, true);
    const nameLen = view.getUint16(ptr + 28, true);
    const extraLen = view.getUint16(ptr + 30, true);
    const commentLen = view.getUint16(ptr + 32, true);
    const localOff = view.getUint32(ptr + 42, true);

    const name = new TextDecoder('utf-8').decode(bytes.subarray(ptr + 46, ptr + 46 + nameLen));
    ptr += 46 + nameLen + extraLen + commentLen;

    if (!wanted(name)) continue;
    if (localOff + 30 > bytes.length || view.getUint32(localOff, true) !== 0x04034b50) continue;

    const lNameLen = view.getUint16(localOff + 26, true);
    const lExtraLen = view.getUint16(localOff + 28, true);
    const start = localOff + 30 + lNameLen + lExtraLen;
    const body = bytes.subarray(start, start + compSize);

    if (method === 0) out[name] = body;
    else if (method === 8) out[name] = inflateRaw(body, rawSize);
    // Any other method (bzip2, lzma) is vanishingly rare in Office files and is
    // skipped rather than guessed at.
  }

  return out;
}

/**
 * Strip XML tags to readable text, keeping paragraph and cell boundaries.
 * @param {string} xml
 * @param {string} textTag the tag holding runs of text, e.g. `w:t`
 * @param {string} breakTag the tag that ends a line, e.g. `w:p`
 * @returns {string}
 */
function xmlText(xml, textTag, breakTag) {
  const runs = [];
  const re = new RegExp('<' + textTag + '(?:\\s[^>]*)?>([\\s\\S]*?)</' + textTag + '>|<' + breakTag + '(?:\\s[^>]*)?(?:/)?>', 'g');
  let m = re.exec(xml);
  while (m !== null) {
    if (m[1] === undefined) runs.push('\n');
    else runs.push(m[1]);
    m = re.exec(xml);
  }
  return runs
    .join('')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;|&apos;/g, "'")
    .replace(/&amp;/g, '&')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/**
 * Text of a .docx — body only. Headers, footers and footnotes are deliberately
 * skipped: they are boilerplate in every document this will ever see, and
 * including them would push real content out of the character budget.
 *
 * @param {Uint8Array} bytes
 * @returns {string}
 */
function docxText(bytes) {
  const zip = readZip(bytes, (n) => n === 'word/document.xml');
  const doc = zip['word/document.xml'];
  if (!doc) return '';
  return xmlText(new TextDecoder('utf-8').decode(doc), 'w:t', 'w:p');
}

/**
 * Text of a .pptx, slide by slide and in slide order — the order matters,
 * because a deck read out of sequence reads as nonsense.
 *
 * @param {Uint8Array} bytes
 * @returns {string}
 */
function pptxText(bytes) {
  const zip = readZip(bytes, (n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
  const names = Object.keys(zip).sort((a, b) => {
    const na = Number((a.match(/slide(\d+)\.xml$/) || [])[1] || 0);
    const nb = Number((b.match(/slide(\d+)\.xml$/) || [])[1] || 0);
    return na - nb;
  });
  const parts = [];
  for (const name of names) {
    const n = (name.match(/slide(\d+)\.xml$/) || [])[1];
    const text = xmlText(new TextDecoder('utf-8').decode(zip[name]), 'a:t', 'a:p');
    if (text) parts.push('[Slide ' + n + '] ' + text);
  }
  return parts.join('\n\n');
}

/**
 * Text of an .xlsx, as rows.
 *
 * Spreadsheets are the dominant attachment in this inbox, so it is worth doing
 * properly rather than dumping the string table: a list of every distinct
 * string in a workbook tells you the vocabulary but not which column a number
 * sits under, and "the crossarm figure is $6.4M" only survives if the row stays
 * intact.
 *
 * Only the first `maxRows` rows of each sheet are read. Beyond that a tracker
 * repeats itself, and the character budget is better spent on the other
 * attachments.
 *
 * @param {Uint8Array} bytes
 * @param {{ maxRows?: number, maxSheets?: number }} [limits]
 * @returns {string}
 */
function xlsxText(bytes, limits) {
  const opts = limits || {};
  const maxRows = opts.maxRows === undefined ? 40 : opts.maxRows;
  const maxSheets = opts.maxSheets === undefined ? 5 : opts.maxSheets;

  const zip = readZip(
    bytes,
    (n) => n === 'xl/sharedStrings.xml' || n === 'xl/workbook.xml' || /^xl\/worksheets\/sheet\d+\.xml$/.test(n),
  );
  const dec = new TextDecoder('utf-8');

  /** Shared strings are referenced by index from every cell that holds text. */
  const shared = [];
  if (zip['xl/sharedStrings.xml']) {
    const xml = dec.decode(zip['xl/sharedStrings.xml']);
    const siRe = /<si>([\s\S]*?)<\/si>/g;
    let si = siRe.exec(xml);
    while (si !== null) {
      // Rich text splits one value across several <t> runs; join them.
      const runs = (si[1] || '').match(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g) || [];
      shared.push(
        runs
          .map((r) => r.replace(/<[^>]+>/g, ''))
          .join('')
          .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
          .replace(/&#39;|&apos;/g, "'").replace(/&amp;/g, '&'),
      );
      si = siRe.exec(xml);
    }
  }

  /** Sheet display names, in workbook order. */
  const names = [];
  if (zip['xl/workbook.xml']) {
    const wb = dec.decode(zip['xl/workbook.xml']);
    const re = /<sheet[^>]*name="([^"]*)"/g;
    let m = re.exec(wb);
    while (m !== null) {
      names.push((m[1] || '').replace(/&amp;/g, '&'));
      m = re.exec(wb);
    }
  }

  const sheetFiles = Object.keys(zip)
    .filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))
    .sort((a, b) => {
      const na = Number((a.match(/sheet(\d+)\.xml$/) || [])[1] || 0);
      const nb = Number((b.match(/sheet(\d+)\.xml$/) || [])[1] || 0);
      return na - nb;
    })
    .slice(0, maxSheets);

  const out = [];
  for (let i = 0; i < sheetFiles.length; i += 1) {
    const sheetName = sheetFiles[i] || '';
    const xml = dec.decode(zip[sheetName]);
    const rows = [];
    const rowRe = /<row[^>]*>([\s\S]*?)<\/row>/g;
    let row = rowRe.exec(xml);
    while (row !== null && rows.length < maxRows) {
      const rowXml = row[1] || '';
      const cells = [];
      const cellRe = /<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b[^>]*\/>/g;
      let cell = cellRe.exec(rowXml);
      while (cell !== null) {
        const attrs = cell[1] || '';
        const inner = cell[2] || '';
        const t = (attrs.match(/\bt="([^"]*)"/) || [])[1] || '';
        let value = '';
        if (t === 's') {
          const idx = Number((inner.match(/<v>([\s\S]*?)<\/v>/) || [])[1]);
          value = Number.isFinite(idx) && shared[idx] !== undefined ? shared[idx] : '';
        } else if (t === 'inlineStr') {
          value = (inner.match(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/) || [])[1] || '';
        } else {
          value = (inner.match(/<v>([\s\S]*?)<\/v>/) || [])[1] || '';
        }
        cells.push(
          String(value)
            .replace(/<[^>]+>/g, '')
            .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
            .replace(/&#39;|&apos;/g, "'").replace(/&amp;/g, '&')
            .trim(),
        );
        cell = cellRe.exec(rowXml);
      }
      // A row of nothing but empty cells is formatting, not data.
      if (cells.some((c) => c.length > 0)) rows.push(cells.join(' | '));
      row = rowRe.exec(xml);
    }
    if (rows.length > 0) {
      const title = names[i] || 'Sheet ' + (i + 1);
      out.push('[' + title + ']\n' + rows.join('\n'));
    }
  }
  return out.join('\n\n');
}

/**
 * Text of an Office Open XML file, chosen by extension.
 * @param {Uint8Array} bytes
 * @param {string} filename
 * @returns {string}
 */
function officeText(bytes, filename) {
  const lower = String(filename || '').toLowerCase();
  if (lower.endsWith('.docx')) return docxText(bytes);
  if (lower.endsWith('.pptx')) return pptxText(bytes);
  if (lower.endsWith('.xlsx') || lower.endsWith('.xlsm')) return xlsxText(bytes);
  return '';
}

export { inflateRaw, readZip, docxText, pptxText, xlsxText, officeText, xmlText };
