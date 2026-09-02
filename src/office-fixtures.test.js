import { describe, expect, it } from 'vitest';
import { crc32 as nodeCrc32, inflateRawSync } from 'node:zlib';
import { docxFixture, pdfFixture, pngFixture, pptxFixture, toBase64, xlsxFixture, zipStored } from './office-fixtures.js';
import { docxText, officeText, pptxText, readZip, xlsxText } from './unzip.js';

/**
 * A fixture that does not parse is worse than no fixture: it fails in the
 * sandbox and reads exactly like a broken reader. That happened once with a
 * pasted .docx — right length, valid ZIP signature, inflated to nothing.
 *
 * So every fixture here is round-tripped through the real extractor, and the
 * text has to come back out.
 */
describe('the fixtures parse with the readers that will read them', () => {
  it('builds a .docx whose paragraphs come back', () => {
    const bytes = docxFixture(['Crossarm capitalization, Q3', 'Two hundred poles surveyed.']);
    const text = docxText(bytes);
    expect(text).toContain('Crossarm capitalization, Q3');
    expect(text).toContain('Two hundred poles surveyed.');
  });

  it('builds a .pptx and keeps the slides in order', () => {
    const bytes = pptxFixture([['Momentum Room'], ['CAIDI trend'], ['Next steps']]);
    const text = pptxText(bytes);
    expect(text).toContain('[Slide 1] Momentum Room');
    expect(text).toContain('[Slide 3] Next steps');
    expect(text.indexOf('[Slide 1]')).toBeLessThan(text.indexOf('[Slide 2]'));
  });

  it('builds an .xlsx that reads through the shared-string table', () => {
    const bytes = xlsxFixture('Crossarms', [
      ['District', 'Poles', 'Status'],
      ['Northern', '412', 'In progress'],
    ]);
    const text = xlsxText(bytes);
    expect(text).toContain('[Crossarms]');
    expect(text).toContain('District | Poles | Status');
    expect(text).toContain('Northern | 412 | In progress');
  });

  it('routes each one by extension, the way WF-D will', () => {
    expect(officeText(docxFixture(['hello']), 'a.docx')).toContain('hello');
    expect(officeText(pptxFixture([['hello']]), 'a.pptx')).toContain('hello');
    expect(officeText(xlsxFixture('S', [['hello']]), 'a.xlsx')).toContain('hello');
  });

  // Escaping is the quiet failure: a stray ampersand produces a ZIP that opens
  // and XML that does not, so the text comes back empty rather than wrong.
  it('survives characters that would break the XML', () => {
    const text = docxText(docxFixture(['Costs & fees <urgent> "quoted"']));
    expect(text).toContain('Costs & fees <urgent> "quoted"');
  });
});

describe('the archives are real ZIPs, not shapes that look like one', () => {
  it('writes a central directory a reader can walk', () => {
    const bytes = zipStored([{ name: 'a.txt', text: 'one' }, { name: 'b/c.txt', text: 'two' }]);
    const zip = readZip(bytes, () => true);
    expect(new TextDecoder().decode(zip['a.txt'])).toBe('one');
    expect(new TextDecoder().decode(zip['b/c.txt'])).toBe('two');
  });

  it('records a correct CRC and size in the local header', () => {
    const bytes = zipStored([{ name: 'x.txt', text: 'payload' }]);
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    expect(view.getUint32(0, true)).toBe(0x04034b50);
    expect(view.getUint16(8, true)).toBe(0); // stored
    expect(view.getUint32(18, true)).toBe(7); // compressed size
    expect(view.getUint32(22, true)).toBe(7); // uncompressed size
    // Checked against Node's own CRC-32 rather than a constant copied in by
    // hand: a wrong table corrupts every file it writes and nothing else here
    // would notice.
    expect(view.getUint32(14, true)).toBe(nodeCrc32(Buffer.from('payload')));
  });

  /**
   * Our own reader could agree with our own writer while both are wrong. Node's
   * zlib is the outside opinion: a stored entry is raw deflate-free bytes, so
   * inflating it must fail rather than quietly succeed on nonsense.
   */
  it('stores rather than deflates, as the header claims', () => {
    const bytes = zipStored([{ name: 'x.txt', text: 'payload' }]);
    const body = bytes.slice(30 + 5, 30 + 5 + 7);
    expect(new TextDecoder().decode(body)).toBe('payload');
    expect(() => inflateRawSync(Buffer.from(body))).toThrow();
  });
});

describe('the checksum, against an outside implementation', () => {
  it.each(['', 'a', 'payload', 'the quick brown fox', '\u00e9\u00e8\u00ea — accented', 'x'.repeat(1000)])(
    'agrees with node zlib on %o',
    (input) => {
      const bytes = zipStored([{ name: 'f', text: input }]);
      const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      expect(view.getUint32(14, true)).toBe(nodeCrc32(Buffer.from(input, 'utf8')));
    },
  );
});

describe('the PDF and the image', () => {
  it('builds a PDF with a header, a trailer and a resolvable xref', () => {
    const bytes = pdfFixture(['Paychex invoice', 'Total due: see attached']);
    const text = new TextDecoder().decode(bytes);
    expect(text.startsWith('%PDF-1.4')).toBe(true);
    expect(text.trimEnd().endsWith('%%EOF')).toBe(true);
    expect(text).toContain('/Type /Catalog');
    expect(text).toContain('(Paychex invoice) Tj');

    // startxref must point at the byte where "xref" actually begins, or a
    // reader rejects the file.
    const startxref = Number((text.match(/startxref\n(\d+)/) || [])[1]);
    expect(text.slice(startxref, startxref + 4)).toBe('xref');
  });

  it('escapes parentheses, which would otherwise close the text object early', () => {
    const text = new TextDecoder().decode(pdfFixture(['Fee (estimated) \\ pending']));
    expect(text).toContain('(Fee \\(estimated\\) \\\\ pending) Tj');
  });

  it('builds a PNG small enough that the triage skips it', () => {
    const png = pngFixture();
    expect(Array.from(png.slice(0, 8))).toEqual([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    // Below TINY_IMAGE_BYTES in src/attachments.js, so it is skipped unread.
    expect(png.length).toBeLessThan(12 * 1024);
  });
});

describe('base64', () => {
  it.each([
    ['', ''],
    ['f', 'Zg=='],
    ['fo', 'Zm8='],
    ['foo', 'Zm9v'],
    ['foob', 'Zm9vYg=='],
    ['fooba', 'Zm9vYmE='],
    ['foobar', 'Zm9vYmFy'],
  ])('encodes %o the way the RFC says', (input, expected) => {
    expect(toBase64(new TextEncoder().encode(input))).toBe(expected);
  });

  it('agrees with Node on bytes that are not text', () => {
    const bytes = pngFixture();
    expect(toBase64(bytes)).toBe(Buffer.from(bytes).toString('base64'));
  });
});
