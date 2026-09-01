import { describe, expect, it } from 'vitest';
import {
  triageAll,
  triageAttachment,
  describeSize,
  condense,
  summariseAttachments,
  BY_EXTENSION,
  KNOWN_UNREADABLE,
} from './attachments.js';

/** @param {Partial<Record<string, any>>} over */
const att = (over) =>
  Object.assign(
    { name: 'file.pdf', contentType: 'application/pdf', size: 100_000, isInline: false, '@odata.type': '#microsoft.graph.fileAttachment' },
    over,
  );

describe('triage — which attachments get opened', () => {
  it('routes the file types this inbox actually receives', () => {
    const cases = [
      ['Q3 deck.pptx', 'office'],
      ['scope.docx', 'office'],
      // .xlsx is a ZIP like .docx/.pptx, so it goes down the same in-process
      // unzip path rather than needing a separate binary conversion.
      ['crossarm workbook.xlsx', 'office'],
      // .xls is OLE2, not a ZIP — named honestly instead of half-read.
      ['legacy.xls', 'unreadable'],
      ['SVP review.pdf', 'pdf'],
      ['screenshot.png', 'image'],
      ['photo.JPEG', 'image'],
      ['notes.txt', 'text'],
      ['export.csv', 'text'],
    ];
    for (const [name, route] of cases) {
      expect(triageAttachment(att({ name, contentType: '' })).route, name).toBe(route);
    }
  });

  // Outlook labels plenty of ordinary files application/octet-stream, so the
  // extension has to win.
  it('trusts the extension when Outlook reports a generic content type', () => {
    const v = triageAttachment(att({ name: 'Vector CI.pptx', contentType: 'application/octet-stream' }));
    expect(v.route).toBe('office');
  });

  it('falls back to content type when there is no extension', () => {
    expect(triageAttachment(att({ name: 'scan', contentType: 'image/png' })).route).toBe('image');
    expect(triageAttachment(att({ name: 'doc', contentType: 'application/pdf' })).route).toBe('pdf');
    expect(
      triageAttachment(att({ name: 'sheet', contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })).route,
    ).toBe('office');
  });

  it('names what a Power BI file is rather than shrugging at it', () => {
    const v = triageAttachment(att({ name: 'Electrical_Ops.pbix', contentType: '' }));
    expect(v.route).toBe('unreadable');
    expect(v.reason).toContain('Power BI');
  });

  it('skips inline signatures and logos', () => {
    const logo = triageAttachment(att({ name: 'logo.png', contentType: 'image/png', size: 9_000, isInline: true }));
    expect(logo.route).toBe('skip');
    expect(logo.reason).toContain('signature/logo');
  });

  // A screenshot pasted into the body is inline too, and is the whole point of
  // the message. Size is what separates it from a logo.
  it('still opens a real inline screenshot', () => {
    const shot = triageAttachment(att({ name: 'chart.png', contentType: 'image/png', size: 400_000, isInline: true }));
    expect(shot.route).toBe('image');
  });

  it('refuses anything over the size limit', () => {
    const v = triageAttachment(att({ name: 'huge.pdf', size: 40 * 1024 * 1024 }));
    expect(v.route).toBe('skip');
    expect(v.reason).toContain('limit');
  });

  // The ceiling used to be 8 MB, which refused a real 22 MB PDF that turned out
  // to read fine in 21 seconds. Pinned so it is not quietly lowered back to a
  // number nobody measured.
  it('opens a PDF the size of a real scanned deck', () => {
    const v = triageAttachment(att({ name: 'Upperform Partners.pdf', size: 22 * 1024 * 1024 }));
    expect(v.route).toBe('pdf');
  });

  it('says how big the limit was when it refuses', () => {
    const v = triageAttachment(att({ name: 'huge.pdf', size: 40 * 1024 * 1024 }));
    expect(v.reason).toContain('25 MB');
  });

  it('skips Graph attachment kinds that have no bytes', () => {
    const item = triageAttachment(att({ name: 'Fwd', '@odata.type': '#microsoft.graph.itemAttachment' }));
    expect(item.route).toBe('skip');
    expect(item.reason).toContain('embedded Outlook item');

    const ref = triageAttachment(att({ name: 'Deck', '@odata.type': '#microsoft.graph.referenceAttachment' }));
    expect(ref.route).toBe('skip');
    expect(ref.reason).toContain('cloud storage');
  });

  it('does not claim to read an unknown binary', () => {
    const v = triageAttachment(att({ name: 'thing.qqq', contentType: 'application/x-weird' }));
    expect(v.route).toBe('unreadable');
    expect(v.reason).toContain('.qqq');
  });

  // A signature logo that arrives without an inline marker — which is exactly
  // what forwarding produces — used to cost a vision call to be told it was a
  // logo, and ate part of the prompt budget saying so.
  it('skips a signature logo that arrived without an inline marker', () => {
    const v = triageAttachment(att({ name: 'image001.png', contentType: 'image/png', size: 18 * 1024, isInline: false }));
    expect(v.route).toBe('skip');
    expect(v.reason).toContain('signature/logo');
  });

  it('skips the logo Outlook itself attaches', () => {
    const v = triageAttachment(att({ name: 'Outlook-Logo  Desc.png', contentType: 'image/png', size: 18 * 1024, isInline: false }));
    expect(v.route).toBe('skip');
  });

  // The whole point of the name check is that it must not swallow a real one.
  it('still opens a small screenshot that has a real name', () => {
    const v = triageAttachment(att({ name: 'Screenshot 2026-09-02 at 10.14.21.png', contentType: 'image/png', size: 30 * 1024, isInline: false }));
    expect(v.route).toBe('image');
  });

  it('never spends a model call on a tracking pixel', () => {
    const v = triageAttachment(att({ name: 'beacon.gif', contentType: 'image/gif', size: 43, isInline: false }));
    expect(v.route).toBe('skip');
  });

  it('has no extension mapped to two different routes', () => {
    for (const ext of Object.keys(KNOWN_UNREADABLE)) {
      expect(BY_EXTENSION[ext], ext + ' is both readable and unreadable').toBeUndefined();
    }
  });
});

/**
 * Caps that only mean anything across a whole message: how many attachments
 * are opened at all, and how many may cost a model call. A 22 MB PDF takes
 * about 21 seconds, and WF-A waits for the answer before it classifies, so an
 * unbounded number of them holds up the poller.
 */
describe('triageAll — what one message is allowed to cost', () => {
  /** @param {string} name @param {number} [size] */
  const pdf = (name, size) => att({ name, contentType: 'application/pdf', size: size || 900_000 });

  it('opens only the first few files that need the model', () => {
    const out = triageAll([pdf('a.pdf'), pdf('b.pdf'), pdf('c.pdf'), pdf('d.pdf')], { maxModelCalls: 2 });
    expect(out.map((v) => v.route)).toEqual(['pdf', 'pdf', 'skip', 'skip']);
  });

  it('names the ones it did not open rather than dropping them', () => {
    const out = triageAll([pdf('a.pdf'), pdf('b.pdf')], { maxModelCalls: 1 });
    expect(out[1]?.reason).toContain('only the first 1');
    expect(out[1]?.label).toContain('b.pdf');
  });

  // Office files are unzipped in process and cost milliseconds, so they must
  // not consume a budget that exists to bound model calls.
  it('does not charge office files against the model budget', () => {
    const out = triageAll(
      [att({ name: 'deck.pptx', contentType: '' }), att({ name: 'book.xlsx', contentType: '' }), pdf('late.pdf')],
      { maxModelCalls: 1 },
    );
    expect(out.map((v) => v.route)).toEqual(['office', 'office', 'pdf']);
  });

  // Nor should something that was skipped anyway.
  it('does not charge a skipped logo against the model budget', () => {
    const out = triageAll(
      [att({ name: 'image001.png', contentType: 'image/png', size: 9000 }), pdf('real.pdf')],
      { maxModelCalls: 1 },
    );
    expect(out.map((v) => v.route)).toEqual(['skip', 'pdf']);
  });

  it('still refuses everything past the attachment cap', () => {
    const out = triageAll([pdf('a.pdf'), pdf('b.pdf'), pdf('c.pdf')], { maxAttachments: 2, maxModelCalls: 9 });
    expect(out[2]?.route).toBe('skip');
    expect(out[2]?.reason).toContain('beyond the first 2');
  });

  it('reports which of them need their bytes fetched', () => {
    const out = triageAll([pdf('a.pdf'), att({ name: 'thing.pbix', contentType: '' })]);
    expect(out.map((v) => v.needsBytes)).toEqual([true, false]);
  });

  it('survives a message with nothing attached', () => {
    expect(triageAll([])).toEqual([]);
  });
});

describe('sizes read the way a person would say them', () => {
  it.each([
    [512, '512 bytes'],
    [2048, '2 KB'],
    [1024 * 1024, '1 MB'],
    [(2.5 * 1024 * 1024) | 0, '2.5 MB'],
  ])('%i -> %s', (bytes, expected) => {
    expect(describeSize(bytes)).toBe(expected);
  });
});

describe('condense', () => {
  it('leaves short text alone', () => {
    expect(condense('Crossarm workbook attached.')).toBe('Crossarm workbook attached.');
  });

  it('cuts on a word boundary and says it truncated', () => {
    const out = condense('word '.repeat(2000), 100);
    expect(out.length).toBeLessThan(140);
    expect(out).toContain('truncated');
    expect(out).not.toMatch(/wor$/);
  });

  it('collapses the whitespace a spreadsheet dump arrives with', () => {
    expect(condense('a\r\n\n\n\n b     c')).toBe('a\n\nb c');
  });
});

describe('summary — what the model is told about the attachments', () => {
  it('is empty when there are none', () => {
    expect(summariseAttachments([]).text).toBe('');
  });

  it('names each attachment and quotes what was read', () => {
    const out = summariseAttachments([
      { label: 'workbook.xlsx (240 KB)', route: 'sheet', text: 'Feeder,CAIDI\nF12,88' },
    ]);
    expect(out.text).toContain('workbook.xlsx (240 KB)');
    expect(out.text).toContain('F12,88');
    expect(out.read).toBe(1);
  });

  // An attachment silently ignored is worse than one that says it was skipped:
  // Yancey can act on the second and cannot see the first.
  it('says so when something was deliberately not read', () => {
    const out = summariseAttachments([
      { label: 'Electrical_Ops.pbix (12 MB)', route: 'unreadable', reason: 'Power BI report, cannot be read as text' },
    ]);
    expect(out.text).toContain('not read: Power BI report');
    expect(out.skipped).toBe(1);
  });

  it('reports a failure without pretending it succeeded', () => {
    const out = summariseAttachments([{ label: 'broken.pdf', route: 'pdf', error: 'password protected' }]);
    expect(out.text).toContain('could not be read (password protected)');
    expect(out.failed).toBe(1);
  });

  it('flags an attachment that opened but held nothing readable', () => {
    const out = summariseAttachments([{ label: 'blank.pdf', route: 'pdf', text: '   ' }]);
    expect(out.text).toContain('no readable text');
  });

  // One 200-page deck must not crowd three other attachments out of the prompt.
  it('shares the budget so no single attachment starves the others', () => {
    const big = 'x'.repeat(50_000);
    const out = summariseAttachments(
      [
        { label: 'a.pdf', route: 'pdf', text: big },
        { label: 'b.pdf', route: 'pdf', text: big },
        { label: 'c.pdf', route: 'pdf', text: big },
      ],
      { totalBudget: 3000 },
    );
    expect(out.read).toBe(3);
    expect(out.text).toContain('a.pdf');
    expect(out.text).toContain('b.pdf');
    expect(out.text).toContain('c.pdf');
    expect(out.text.length).toBeLessThan(4200);
  });

  it('stays within the total budget even with many attachments', () => {
    const items = Array.from({ length: 12 }, (_, i) => ({
      label: 'f' + i + '.pdf',
      route: 'pdf',
      text: 'y'.repeat(9000),
    }));
    const out = summariseAttachments(items, { totalBudget: 6000 });
    const quoted = out.text.replace(/- f\d+\.pdf:?/g, '').length;
    expect(quoted).toBeLessThan(9000);
  });
});
