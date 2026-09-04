import { describe, expect, it } from 'vitest';
import { STATUSES, diagnose, parseFileName, planRename, planRenames, proposeFileName } from './filenames.js';

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

/**
 * The real contents of Yancey's Task Management library, read on 2 September
 * 2026. Renaming is destructive — a saved link, a Teams tab or a Power BI data
 * source stops resolving — so the decision about what to rename is made here,
 * where it can be tested, rather than inside a workflow node where it cannot.
 */
const HIS_LIBRARY = [
  { name: 'Task Management Channel Guide_YD_04042025.docx', path: '/sites/TaskManagement/Shared Documents/a.docx', modified: '2025-04-04T10:00:00Z' },
  { name: 'TCO Pilot Program Oversight Dashboard_Review_KV_14042025.pbix', path: '/sites/TaskManagement/Shared Documents/b.pbix', modified: '2025-04-14T10:00:00Z' },
  { name: 'Screenshot 2025-04-06 at 6.42.39 PM.png', path: '/sites/TaskManagement/Shared Documents/c.png', modified: '2025-04-06T10:00:00Z' },
  { name: 'Working TimesKV.xlsx', path: '/sites/TaskManagement/Shared Documents/d.xlsx', modified: '2025-05-01T10:00:00Z' },
];

describe('planRename — what to do with one legacy file', () => {
  /** @param {string} name @param {string} [modified] */
  const plan = (name, modified) =>
    planRename({ name, path: '/x/' + name, modified: modified || '2025-04-04T10:00:00Z' }, {});

  it('leaves a compliant file alone', () => {
    const out = plan('SCE_PM_Ride-Along-Plan_2026-09-01_Draft_v0.1.docx');
    expect(out.compliant).toBe(true);
    expect(out.proposed).toBe('');
  });

  // Status is a claim about whether a document is approved. Draft is the
  // weakest such claim, so assuming it under-states rather than over-states.
  it('assumes Draft when no status is in the name, and says so', () => {
    const out = plan('Working TimesKV.xlsx', '2025-05-01T10:00:00Z');
    expect(out.proposed).toContain('_Draft_');
    expect(out.assumptions.join(' ')).toContain('status');
  });

  it('never assumes a status when the name already carries one', () => {
    const out = plan('Oversight Dashboard_Review_v0.2.pbix', '2025-04-14T10:00:00Z');
    expect(out.proposed).toContain('_Review_');
    expect(out.assumptions.join(' ')).not.toContain('status');
  });

  it('reads a year-first date out of the name rather than assuming', () => {
    const out = plan('Screenshot 2025-04-06 at 6.42.39 PM.png');
    expect(out.proposed).toContain('2025-04-06');
    expect(out.assumptions.join(' ')).not.toContain('date');
  });

  // The file's own timestamp stands in for a date only when the name carries
  // none at all. Where the name has digits that cannot be read, see the
  // refusals below — substituting would leave two dates in one name.
  it('falls back to the file timestamp when the name carries no date at all', () => {
    const out = plan('Working TimesKV.xlsx', '2025-05-01T10:00:00Z');
    expect(out.assumptions.join(' ')).toContain('date');
    expect(out.proposed).toContain('2025-05-01');
  });

  it('produces a name its own checker accepts', () => {
    for (const f of HIS_LIBRARY) {
      const out = planRename(f, {});
      if (!out.proposed) continue;
      expect(diagnose(out.proposed), f.name + ' -> ' + out.proposed).toEqual([]);
    }
  });

  it('refuses when there is no document name left to keep', () => {
    const out = plan('v0.1.docx');
    expect(out.proposed).toBe('');
    expect(out.blocked).toBeTruthy();
  });

  it('can be told not to assume anything', () => {
    const out = planRename(
      { name: 'Working TimesKV.xlsx', path: '/x', modified: '2025-05-01T10:00:00Z' },
      { assumeStatus: '', assumeVersion: null },
    );
    expect(out.proposed).toBe('');
    expect(out.blocked).toContain('status');
  });
});

describe('planRenames — what to do with a whole library', () => {
  it('refuses both files when two would land on the same name', () => {
    const out = planRenames(
      [
        { name: 'Report Draft v0.1.docx', path: '/x/1', modified: '2025-04-04T10:00:00Z' },
        { name: 'Report_Draft_v0.1.docx', path: '/x/2', modified: '2025-04-04T10:00:00Z' },
      ],
      {},
    );
    const blocked = out.filter((r) => r.blocked.indexOf('same name') !== -1);
    expect(blocked.length).toBe(2);
    expect(out.every((r) => r.proposed === '')).toBe(true);
  });

  // A suffix that means something is kept, so these two do not collide.
  it('keeps a distinguishing suffix rather than forcing a collision', () => {
    const out = planRenames(
      [
        { name: 'Report_Draft_v0.1 copy.docx', path: '/x/1', modified: '2025-04-04T10:00:00Z' },
        { name: 'Report_Draft_v0.1 (2).docx', path: '/x/2', modified: '2025-04-04T10:00:00Z' },
      ],
      {},
    );
    expect(out[0]?.proposed).toBe('Report-copy_2025-04-04_Draft_v0.1.docx');
    expect(out[1]?.proposed).toBe('Report-2_2025-04-04_Draft_v0.1.docx');
    expect(out.every((r) => r.blocked === '')).toBe(true);
  });

  it('refuses to rename onto a name already in the library', () => {
    const out = planRenames(
      [
        { name: 'Guide_2025-04-04_Draft_v0.1.docx', path: '/x/taken', modified: '2025-04-04T10:00:00Z' },
        { name: 'Guide Draft v0.1 2025-04-04.docx', path: '/x/other', modified: '2025-04-04T10:00:00Z' },
      ],
      {},
    );
    const collided = out.filter((r) => r.blocked);
    expect(collided.length).toBeGreaterThan(0);
  });

  it('leaves the already-correct file untouched and unblocked', () => {
    const out = planRenames(
      [{ name: 'SCE_PM_Plan_2026-09-01_Draft_v0.1.docx', path: '/x/1', modified: '2026-09-01T10:00:00Z' }],
      {},
    );
    expect(out[0]?.compliant).toBe(true);
    expect(out[0]?.blocked).toBe('');
  });

  it('handles the real library without throwing and proposes only valid names', () => {
    const out = planRenames(HIS_LIBRARY, {});
    expect(out.length).toBe(HIS_LIBRARY.length);
    for (const row of out) {
      if (row.proposed) expect(diagnose(row.proposed), row.proposed).toEqual([]);
    }
  });
});

describe('a name that already carries something unreadable', () => {
  /** @param {string} name */
  const plan = (name) => planRename({ name, path: '/x/' + name, modified: '2025-04-07T10:00:00Z' }, {});

  // Refusing to read 04042025 while still leaving it in the name produced
  // Task-Management-Channel-Guide_YD_04042025_2025-04-07_Draft_v0.1.docx —
  // two dates that disagree. Carrying two is worse than either choice.
  it('refuses a name holding an eight-digit date it cannot read', () => {
    const out = plan('Task Management Channel Guide_YD_04042025.docx');
    expect(out.proposed).toBe('');
    expect(out.blocked).toContain('04042025');
  });

  it('refuses when the digits are genuinely ambiguous between conventions', () => {
    const out = plan('TCO Pilot Program Oversight Dashboard_Review_KV_14042025.pbix');
    expect(out.proposed).toBe('');
    expect(out.blocked).toContain('14042025');
  });

  // Trackingv1 has no separator, so it is not read as a version — but appending
  // _v0.1 would leave the file carrying two.
  it('refuses a version glued to a word rather than adding a second one', () => {
    const out = plan('Daily_Employee_Project_Time_Trackingv1.xlsx');
    expect(out.proposed).toBe('');
    expect(out.blocked).toContain('version');
  });

  it('still renames a name with nothing unreadable in it', () => {
    const out = plan('Working TimesKV.xlsx');
    expect(out.proposed).toBe('Working-TimesKV_2025-04-07_Draft_v0.1.xlsx');
  });

  // A time of day is not an eight-digit date and must not be mistaken for one.
  it('is not tripped by a timestamp in a screenshot name', () => {
    const out = plan('Screenshot 2025-04-06 at 6.42.39 PM.png');
    expect(out.proposed).toContain('2025-04-06');
    expect(out.blocked).toBe('');
  });
});

/**
 * The Task Management library is a Power BI working area — its top folder is
 * literally "Power BI Task Management", and four .pbix reports sit among the
 * workbooks they read from. A .pbix stores its data-source paths inside the
 * file, so renaming a workbook breaks the report's refresh, and nothing on the
 * outside of either file says which workbooks a given report reads.
 *
 * The tool cannot see inside a .pbix. So where a report is present it declines,
 * and says which report and why, rather than renaming and finding out later.
 */
describe('planRenames — files a Power BI report may be reading', () => {
  /** @param {string} folder @param {string} name */
  const at = (folder, name) => ({
    name,
    path: '/sites/TaskManagement/Shared Documents/' + folder + '/' + name,
    modified: '2026-07-09T10:00:00Z',
  });

  it('refuses to rename a report and its like-named workbook, naming the counterpart', () => {
    const out = planRenames(
      [at('PD alignment', 'Performance Management System.pbix'), at('PD alignment', 'Performance Management System.xlsx')],
      {},
    );
    expect(out.every((r) => r.proposed === '')).toBe(true);
    expect(out.every((r) => r.blocked !== '')).toBe(true);
    const workbook = out.find((r) => r.current.endsWith('.xlsx'));
    expect(workbook?.blocked).toContain('Performance Management System.pbix');
  });

  it('refuses a workbook sharing a report folder, even with an unrelated name', () => {
    const out = planRenames(
      [at('2. In Progress', 'NiSource Metric Dashboard.pbix'), at('2. In Progress', 'NIPSCO_Master Roster_2025.xlsx')],
      {},
    );
    const roster = out.find((r) => r.current.indexOf('Roster') !== -1);
    expect(roster?.proposed).toBe('');
    expect(roster?.blocked).toContain('NiSource Metric Dashboard.pbix');
  });

  // The pairing that folder proximity misses: the data sits at the top of the
  // library and the report two folders away, but the names say they belong together.
  it('refuses a workbook whose name opens the same way as a report elsewhere', () => {
    const out = planRenames(
      [
        at('4. Completed Tasks', 'TCO Pilot Program Oversight Dashboard_Review_KV_14042025.pbix'),
        at('', 'TCO Pilot Program Oversight Dashboard Data.xlsx'),
      ],
      {},
    );
    const data = out.find((r) => r.current.indexOf('Data.xlsx') !== -1);
    expect(data?.proposed).toBe('');
    expect(data?.blocked).toContain('TCO Pilot Program Oversight Dashboard_Review_KV_14042025.pbix');
  });

  it('still renames a screenshot beside a report, which no report reads from', () => {
    const out = planRenames(
      [at('Task 2', 'DIMP E-1 Performance Measures.pbix'), at('Task 2', 'Screenshot 2025-04-06 at 6.34.18 PM.png')],
      {},
    );
    const shot = out.find((r) => r.current.endsWith('.png'));
    expect(shot?.blocked).toBe('');
    expect(shot?.proposed).toBe('Screenshot_2025-04-06_Draft_v0.1.png');
  });

  it('leaves a workbook alone when no report is anywhere in the library', () => {
    const out = planRenames([at('Resources', 'Working TimesKV.xlsx')], {});
    expect(out[0]?.blocked).toBe('');
    expect(out[0]?.proposed).toBe('Working-TimesKV_2026-07-09_Draft_v0.1.xlsx');
  });

  it('does not treat a merely similar name as a pairing', () => {
    const out = planRenames(
      [at('2. In Progress', 'NiSource Metric Dashboard.pbix'), at('Task 2', 'Metrics .xlsx')],
      {},
    );
    const metrics = out.find((r) => r.current.indexOf('Metrics') !== -1);
    expect(metrics?.blocked).toBe('');
  });
});

describe('planRename — things that are not documents', () => {
  it('refuses a folder rather than renaming everything beneath it', () => {
    const out = planRename(
      { name: '1. Assigned Tasks', path: '/lib/1. Assigned Tasks', modified: '2025-04-04T10:00:00Z' },
      {},
    );
    expect(out.proposed).toBe('');
    expect(out.blocked).toContain('no file extension');
  });
});

/**
 * "Keep the document name short but descriptive" is one of his six rules, and
 * it is the one a generator can most easily break. A screenshot name carries a
 * date, which belongs in the date section, and a time, which belongs nowhere —
 * it says when the file was made, not what it holds.
 */
describe('short but descriptive', () => {
  /** @param {string} folder @param {string} name @param {string} [modified] */
  const at = (folder, name, modified) => ({
    name,
    path: '/lib/' + folder + '/' + name,
    modified: modified || '2025-04-06T18:42:00Z',
  });

  it('drops the time of day from a screenshot name', () => {
    const out = planRename(at('Task 1', 'Screenshot 2025-04-06 at 6.42.39 PM.png'), {});
    expect(out.proposed).toBe('Screenshot_2025-04-06_Draft_v0.1.png');
  });

  it('drops a 24-hour time written with an am marker', () => {
    const out = planRename(at('Task 2', 'Screenshot 2025-04-14 at 9.06.08 AM.png'), {});
    expect(out.proposed).toBe('Screenshot_2025-04-14_Draft_v0.1.png');
  });

  // A standard number is not a time. Guessing wrong here deletes meaning.
  it('keeps a dotted standard number that only looks like a time', () => {
    const out = planRename(at('DIMP Folder', 'DEL - asme.b31.8s.2004.pdf', '2025-03-19T10:00:00Z'), {});
    expect(out.proposed).toBe('DEL-asme-b31-8s-2004_2025-03-19_Draft_v0.1.pdf');
  });

  it('keeps a bare number with no time marker at all', () => {
    const out = planRename(at('x', 'Meeting 10.30 Notes.docx', '2025-05-01T10:00:00Z'), {});
    expect(out.proposed).toBe('Meeting-10-30-Notes_2025-05-01_Draft_v0.1.docx');
  });

  it('says the project prefix his format opens with was not supplied', () => {
    const out = planRename(at('x', 'Metrics .xlsx', '2025-04-15T10:00:00Z'), {});
    expect(out.assumptions).toContain('no project prefix, the library does not say which project this belongs to');
  });

  it('uses the project when one is given, as in his examples', () => {
    const out = planRename(at('x', 'Ride Along Plan.docx', '2026-09-01T10:00:00Z'), { project: 'SCE_PM' });
    expect(out.proposed).toBe('SCE_PM_Ride-Along-Plan_2026-09-01_Draft_v0.1.docx');
    expect(out.assumptions).not.toContain('no project prefix, the library does not say which project this belongs to');
  });
});

describe('a name only has to be unique inside its own folder', () => {
  /** @param {string} folder @param {string} name */
  const shot = (folder, name) => ({
    name,
    path: '/lib/' + folder + '/' + name,
    modified: '2025-04-06T18:00:00Z',
  });

  it('refuses same-day screenshots that share a folder', () => {
    const out = planRenames(
      [
        shot('Task 2', 'Screenshot 2025-04-06 at 6.34.18 PM.png'),
        shot('Task 2', 'Screenshot 2025-04-06 at 6.34.50 PM.png'),
      ],
      {},
    );
    expect(out.every((r) => r.proposed === '')).toBe(true);
    expect(out.every((r) => r.blocked.indexOf('same name') !== -1)).toBe(true);
  });

  it('renames same-day screenshots that sit in different folders', () => {
    const out = planRenames(
      [
        shot('Task 1', 'Screenshot 2025-04-06 at 6.42.39 PM.png'),
        shot('Task 2', 'Screenshot 2025-04-06 at 6.34.18 PM.png'),
      ],
      {},
    );
    expect(out.every((r) => r.proposed === 'Screenshot_2025-04-06_Draft_v0.1.png')).toBe(true);
    expect(out.every((r) => r.blocked === '')).toBe(true);
  });
});

describe('the Power BI report switch', () => {
  const FILES = [{ name: 'NiSource Metric Dashboard.pbix', path: '/lib/NiSource Metric Dashboard.pbix', modified: '2026-04-15T10:00:00Z' }];

  it('refuses a report by default, and says why', () => {
    const rows = planRenames(FILES, {});
    expect(rows[0]?.proposed).toBe('');
    expect(rows[0]?.blocked).toContain('published or set to refresh');
  });

  it('proposes a name for one once renameReports is on', () => {
    const rows = planRenames(FILES, { renameReports: true });
    expect(rows[0]?.blocked).toBe('');
    expect(rows[0]?.proposed).toContain('.pbix');
  });
});

/**
 * A proposal its own checker then refuses.
 *
 * Found on a live production run of 2026-09-04. `PD_Reliability_Preloaded_Final.xlsx`
 * carries a status in its name and no version, so proposeFileName took `Final`
 * from the name and `v0.1` from assumeVersion, and offered
 * `PD_Reliability_Preloaded_2026-09-04_Final_v0.1.xlsx` — which diagnose refuses
 * on the next line down, because "v1.0 — First approved version" is in Yancey's
 * own version guide and a document marked Final while numbered v0.x contradicts
 * itself.
 *
 * Nothing was written, so the guard held. What broke is the review: it put a
 * change in front of a person, they approved it, and the run then refused it.
 * A page that offers something it will not do is worse than one that offers
 * nothing, because somebody has already made a decision about it.
 */
describe('a proposal has to satisfy the convention it is proposing under', () => {
  /** @param {string} current */
  const propose = (current) =>
    planRename({ name: current, path: '/x/' + current, modified: '2026-09-04T00:00:00Z' }, {});

  it('does not offer Final at v0.1, which its own diagnose refuses', () => {
    const out = propose('PD_Reliability_Preloaded_Final.xlsx');
    expect(out.proposed, 'nothing to check').toBeTruthy();
    expect(diagnose(out.proposed).length, 'offered ' + out.proposed + ', which diagnose refuses')
      .toBe(0);
  });

  it('assumes the first approved version when the name says Final', () => {
    const out = propose('PD_Reliability_Preloaded_Final.xlsx');
    expect(out.proposed).toContain('_Final_v1.0.');
    expect(out.assumptions.join(' ')).toContain('v1.0');
  });

  it('still assumes v0.1 when the name does not claim to be approved', () => {
    const out = propose('Some_Working_Draft.xlsx');
    expect(out.proposed).toContain('_v0.1.');
  });

  /**
   * The rule is about a version the name did not state. A name that says v0.1
   * AND says Final is a document contradicting itself, and that is a problem to
   * report rather than one to paper over.
   */
  it('does not quietly promote a version the name actually states', () => {
    const out = propose('Thing_2026-09-04_Final_v0.1.xlsx');
    expect(out.compliant, 'this name is wrong and should be reported as wrong').toBe(false);
    expect(out.problems.join(' ')).toContain('an approved version starts at v1.0');
    expect(out.blocked, 'and the reason has to reach the page').toContain('an approved version starts at v1.0');
  });

  /**
   * Whatever a proposal is built from, it has to survive the check. This is the
   * general form of the bug rather than the one case that found it.
   */
  it.each([
    'PD_Reliability_Preloaded_Final.xlsx',
    'Board_Pack_FINAL.docx',
    'Rates_Case_final.xlsx',
    'Notes_Draft.docx',
    'Filing_Review.pdf',
    'Old_Thing_Archived.xlsx',
  ])('offers a name that passes diagnose for %s', (current) => {
    const out = propose(current);
    if (!out.proposed) return;
    expect(diagnose(out.proposed), 'offered ' + out.proposed).toEqual([]);
  });
});
