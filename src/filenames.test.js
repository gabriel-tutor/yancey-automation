import { describe, expect, it } from 'vitest';
import { STATUSES, diagnose, parseFileName, proposeFileName } from './filenames.js';

/**
 * Yancey's naming convention, sent 28 August 2026 and again on the 31st:
 *
 *   [Project]_[Document Name]_[YYYY-MM-DD]_[Status]_v#.#
 *   SCE_PM_Ride-Along-Plan_2026-09-01_Draft_v0.1.docx
 *
 * His four worked examples are the specification. Every one of them has to
 * parse, and the rules he wrote out underneath them are the failures worth
 * naming back to him.
 */

const HIS_EXAMPLES = [
  'SCE_PM_Ride-Along-Plan_2026-09-01_Draft_v0.1.docx',
  'NIPSCO_CAIDI_Training-Guide_2026-08-29_Review_v0.4.pptx',
  'SCE_PM_Opportunity-List_2026-09-02_Draft_v0.2.xlsx',
  'Upperform_Momentum-Room_One-Pager_2026-08-29_Final_v1.0.pdf',
];

describe('the convention Yancey actually wrote', () => {
  it.each(HIS_EXAMPLES)('accepts %s', (name) => {
    expect(parseFileName(name).ok, name).toBe(true);
    expect(diagnose(name), name).toEqual([]);
  });

  // His template reads [Project][Document Name][date], but every example puts
  // two or three sections before the date. Everything ahead of the date is the
  // name; the date, status and version are the part that has to be exact.
  it('takes everything before the date as the name, however many sections', () => {
    expect(parseFileName('SCE_PM_Ride-Along-Plan_2026-09-01_Draft_v0.1.docx').name).toBe('SCE_PM_Ride-Along-Plan');
    expect(parseFileName('Upperform_Momentum-Room_One-Pager_2026-08-29_Final_v1.0.pdf').name).toBe(
      'Upperform_Momentum-Room_One-Pager',
    );
  });

  it('reads the parts back out', () => {
    const p = parseFileName('NIPSCO_CAIDI_Training-Guide_2026-08-29_Review_v0.4.pptx');
    expect(p.date).toBe('2026-08-29');
    expect(p.status).toBe('Review');
    expect(p.version).toEqual({ major: 0, minor: 4 });
    expect(p.extension).toBe('pptx');
  });

  it('offers the four statuses he listed and no others', () => {
    expect(STATUSES).toEqual(['Draft', 'Review', 'Final', 'Archived']);
  });
});

describe('what it says is wrong, in his words', () => {
  it('rejects a status he did not list', () => {
    expect(diagnose('SCE_PM_Plan_2026-09-01_Approved_v1.0.docx').join(' ')).toContain('Approved');
  });

  it('names the date format rather than just failing', () => {
    const problems = diagnose('SCE_PM_Plan_09-01-2026_Draft_v0.1.docx').join(' ');
    expect(problems).toContain('YYYY-MM-DD');
  });

  it('catches a date that is formatted right but is not a real day', () => {
    expect(diagnose('SCE_PM_Plan_2026-02-31_Draft_v0.1.docx').join(' ')).toContain('not a real date');
  });

  it('catches spaces', () => {
    expect(diagnose('SCE PM Plan_2026-09-01_Draft_v0.1.docx').join(' ')).toContain('space');
  });

  it('catches the special characters he asked us to avoid', () => {
    expect(diagnose('SCE_PM_Plan(2)_2026-09-01_Draft_v0.1.docx').join(' ')).toContain('special character');
  });

  // He named these three himself.
  it.each(['SCE_PM_Plan_Final_Final_2026-09-01_Final_v1.0.docx', 'SCE_PM_Latest_2026-09-01_Draft_v0.1.docx', 'SCE_PM_New_2026-09-01_Draft_v0.1.docx'])(
    'rejects the names he called out: %s',
    (name) => {
      expect(diagnose(name).join(' ')).toMatch(/Final_Final|Latest|New/);
    },
  );

  it('insists on a version even when the document is Final', () => {
    expect(diagnose('SCE_PM_Plan_2026-09-01_Final.docx').join(' ')).toContain('version');
  });

  // v1.0 is his "first approved version", so Final at v0.x contradicts itself.
  it('points out a Final still numbered as a working draft', () => {
    const problems = diagnose('SCE_PM_Plan_2026-09-01_Final_v0.4.docx').join(' ');
    expect(problems).toContain('v1.0');
  });

  it('says nothing about a Draft numbered v0.x, which is correct', () => {
    expect(diagnose('SCE_PM_Plan_2026-09-01_Draft_v0.4.docx')).toEqual([]);
  });

  it('handles a name with no extension without throwing', () => {
    expect(() => diagnose('SCE_PM_Plan_2026-09-01_Draft_v0.1')).not.toThrow();
  });

  it('handles an empty name without throwing', () => {
    expect(parseFileName('').ok).toBe(false);
    expect(diagnose('').length).toBeGreaterThan(0);
  });
});

describe('building a compliant name', () => {
  it('rebuilds his own example from its parts', () => {
    expect(
      proposeFileName({
        project: 'SCE',
        document: 'PM_Ride-Along-Plan',
        date: '2026-09-01',
        status: 'Draft',
        version: { major: 0, minor: 1 },
        extension: 'docx',
      }),
    ).toBe('SCE_PM_Ride-Along-Plan_2026-09-01_Draft_v0.1.docx');
  });

  it('turns spaces in a document name into hyphens, the way he asked', () => {
    const out = proposeFileName({
      project: 'NIPSCO',
      document: 'Training Guide',
      date: '2026-08-29',
      status: 'Review',
      version: { major: 0, minor: 4 },
      extension: 'pptx',
    });
    expect(out).toBe('NIPSCO_Training-Guide_2026-08-29_Review_v0.4.pptx');
  });

  it('drops characters he asked us to avoid rather than passing them through', () => {
    const out = proposeFileName({
      project: 'SCE',
      document: 'Plan (2) — copy',
      date: '2026-09-01',
      status: 'Draft',
      version: { major: 0, minor: 1 },
      extension: 'docx',
    });
    expect(diagnose(out)).toEqual([]);
    expect(out).not.toMatch(/[()—]/);
  });

  it('produces something that always satisfies its own checker', () => {
    const out = proposeFileName({
      project: 'Upperform',
      document: 'Momentum Room One Pager',
      date: '2026-08-29',
      status: 'Final',
      version: { major: 1, minor: 0 },
      extension: 'pdf',
    });
    expect(diagnose(out)).toEqual([]);
    expect(parseFileName(out).ok).toBe(true);
  });

  it('refuses to invent a status it was not given', () => {
    expect(() =>
      proposeFileName({ project: 'SCE', document: 'Plan', date: '2026-09-01', status: 'Approved', version: { major: 1, minor: 0 }, extension: 'docx' }),
    ).toThrow(/status/i);
  });
});
