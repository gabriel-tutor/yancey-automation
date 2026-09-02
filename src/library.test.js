import { describe, expect, it } from 'vitest';
import {
  categorise,
  categoriseAll,
  folderOf,
  isReport,
  planPlacements,
  reportReading,
  reportsIn,
  splitExtension,
  stageFolderOf,
} from './library.js';

/**
 * The Task Management library in miniature, keeping the shapes that matter:
 * numbered stage folders, a per-person working folder, a resources folder,
 * reports sitting beside the workbooks they read, and two files loose at the
 * top with no folder at all.
 */
const ROOT = '/sites/TaskManagement/Shared Documents/Power BI Task Management';

/** @param {string} folder @param {string} name */
const at = (folder, name) => ({ name, path: folder ? ROOT + '/' + folder + '/' + name : ROOT + '/' + name });

const LIBRARY = [
  at('', 'Daily_Employee_Project_Time_Trackingv1.xlsx'),
  at('', 'TCO Pilot Program Oversight Dashboard Data.xlsx'),
  at('1. Assigned Tasks', 'Power BI TCO Pilot_YD_04072025.xlsx'),
  at('1. Assigned Tasks/DIMP Folder', 'DEL - asme.b31.8s.2004.pdf'),
  at('1. Assigned Tasks/Task 2', 'Screenshot 2025-04-06 at 6.34.18 PM.png'),
  at('1. Assigned Tasks/Task 2', 'Metrics .xlsx'),
  at('2. In Progress', 'DIMP E-1 Performance Measures.pbix'),
  at('2. In Progress', 'DIMP E-1.xlsx'),
  at('2. In Progress', 'NIPSCO_Master Roster_2025.xlsx'),
  at('4. Completed Tasks', 'TCO Pilot Program Oversight Dashboard_Review_KV_14042025.pbix'),
  at('Resources and References', 'Task Management Channel Guide_YD_04042025.docx'),
];

/** @param {ReturnType<typeof planPlacements>} rows @param {string} needle */
const row = (rows, needle) => rows.find((r) => r.name.indexOf(needle) !== -1);

describe('splitExtension', () => {
  it('reads an ordinary extension', () => {
    expect(splitExtension('Report.docx')).toEqual({ stem: 'Report', extension: 'docx' });
  });

  // The dot in a version is not an extension, or every compliant name would
  // be read as ending in ".1".
  it('does not mistake a version for an extension', () => {
    expect(splitExtension('Plan_2026-09-01_Draft_v0.1')).toEqual({
      stem: 'Plan_2026-09-01_Draft_v0.1',
      extension: '',
    });
  });

  it('ignores a trailing dot', () => {
    expect(splitExtension('Notes.')).toEqual({ stem: 'Notes.', extension: '' });
  });
});

describe('folderOf', () => {
  it('drops the file from the path', () => {
    expect(folderOf('/a/b/c.docx')).toBe('/a/b');
  });

  it('returns empty for a bare name', () => {
    expect(folderOf('c.docx')).toBe('');
  });
});

describe('categorise', () => {
  const reports = reportsIn(LIBRARY);

  it('finds the reports', () => {
    expect(reports.map((r) => r.name)).toEqual([
      'DIMP E-1 Performance Measures.pbix',
      'TCO Pilot Program Oversight Dashboard_Review_KV_14042025.pbix',
    ]);
    expect(isReport('x.pbix')).toBe(true);
    expect(isReport('x.xlsx')).toBe(false);
  });

  it('calls a .pbix a Power BI report', () => {
    const out = categorise(at('2. In Progress', 'NiSource Metric Dashboard.pbix'), reports);
    expect(out.category).toBe('Power BI report');
  });

  it('calls a workbook a report reads report data, and says which report', () => {
    const out = categorise(at('2. In Progress', 'DIMP E-1.xlsx'), reports);
    expect(out.category).toBe('report data');
    expect(out.evidence).toContain('DIMP E-1 Performance Measures.pbix');
  });

  it('calls a workbook nothing reads a spreadsheet', () => {
    const out = categorise(at('1. Assigned Tasks/Task 2', 'Metrics .xlsx'), reports);
    expect(out.category).toBe('spreadsheet');
  });

  it('calls a screen capture a screenshot rather than just an image', () => {
    const out = categorise(at('1. Assigned Tasks/Task 2', 'Screenshot 2025-04-06 at 6.34.18 PM.png'), reports);
    expect(out.category).toBe('screenshot');
  });

  it('calls a photograph an image', () => {
    expect(categorise(at('x', 'crossarm-damage.jpg'), reports).category).toBe('image');
  });

  it('calls a pdf a document', () => {
    expect(categorise(at('x', 'DEL - asme.b31.8s.2004.pdf'), reports).category).toBe('document');
  });

  it('admits when it does not recognise the type', () => {
    const out = categorise(at('x', 'model.qvw'), reports);
    expect(out.category).toBe('other');
    expect(out.evidence).toContain('qvw');
  });

  // An image cannot be a Power BI source, so proximity to a report means
  // nothing for it.
  it('does not call a screenshot report data just for sharing a folder', () => {
    const out = categorise(at('2. In Progress', 'Screenshot 2025-04-06 at 6.34.18 PM.png'), reports);
    expect(out.category).toBe('screenshot');
  });

  it('categorises a whole listing in one pass', () => {
    const all = categoriseAll(LIBRARY);
    expect(all.length).toBe(LIBRARY.length);
    expect(all.filter((r) => r.category === 'Power BI report').length).toBe(2);
  });
});

describe('reportReading', () => {
  const reports = reportsIn(LIBRARY);

  it('pairs a workbook to a report by a shared opening across folders', () => {
    const found = reportReading(
      'TCO Pilot Program Oversight Dashboard Data.xlsx',
      ROOT + '/TCO Pilot Program Oversight Dashboard Data.xlsx',
      reports,
    );
    expect(found?.because).toBe('name');
    expect(found?.report).toContain('TCO Pilot Program Oversight Dashboard');
  });

  it('pairs a workbook to a report by a shared folder', () => {
    const found = reportReading('NIPSCO_Master Roster_2025.xlsx', ROOT + '/2. In Progress/x.xlsx', reports);
    expect(found?.because).toBe('folder');
  });

  it('leaves an unrelated workbook alone', () => {
    const found = reportReading('Metrics .xlsx', ROOT + '/1. Assigned Tasks/Task 2/Metrics .xlsx', reports);
    expect(found).toBeNull();
  });

  it('never says a report reads itself', () => {
    expect(reportReading('DIMP E-1 Performance Measures.pbix', ROOT + '/2. In Progress/x.pbix', reports)).toBeNull();
  });
});

describe('stageFolderOf', () => {
  it('finds a numbered stage folder at any depth', () => {
    expect(stageFolderOf(ROOT + '/1. Assigned Tasks/Task 2/Metrics .xlsx')).toBe('1. Assigned Tasks');
  });

  it('does not treat an ordinary folder as a stage', () => {
    expect(stageFolderOf(ROOT + '/Resources and References/Guide.docx')).toBe('');
  });

  // "2026 Rate Case" opens with digits but is not a step in a sequence.
  it('needs the number to be followed by a separator, not just digits', () => {
    expect(stageFolderOf(ROOT + '/2026 Rate Case/Guide.docx')).toBe('2026 Rate Case');
    expect(stageFolderOf(ROOT + '/2026Rate/Guide.docx')).toBe('');
  });
});

describe('planPlacements', () => {
  const rows = planPlacements(LIBRARY);

  it('leaves a report where it is', () => {
    const r = row(rows, 'DIMP E-1 Performance Measures.pbix');
    expect(r?.verdict).toBe('settled');
    expect(r?.reason).toContain('referenced by path');
  });

  it('leaves a workbook a report reads, for the same reason a rename is refused', () => {
    const r = row(rows, 'TCO Pilot Program Oversight Dashboard Data.xlsx');
    expect(r?.verdict).toBe('settled');
    expect(r?.reason).toContain('a move changes the path exactly as a rename changes the name');
  });

  it('never moves anything out of a numbered stage folder', () => {
    const r = row(rows, 'Metrics .xlsx');
    expect(r?.verdict).toBe('settled');
    expect(r?.reason).toContain('1. Assigned Tasks');
    expect(r?.reason).toContain('nothing about a file says which stage');
  });

  it('flags a file sitting above the filing with folders of documents beneath it', () => {
    const r = row(rows, 'Daily_Employee_Project_Time_Trackingv1.xlsx');
    expect(r?.verdict).toBe('unfiled');
    expect(r?.reason).toContain('above the filing');
  });

  it('is content with a file in an ordinary leaf folder', () => {
    const r = row(rows, 'Task Management Channel Guide');
    expect(r?.verdict).toBe('settled');
    expect(r?.reason).toBe('filed under Resources and References');
  });

  it('proposes no moves at all, and that is the point', () => {
    expect(rows.every((r) => r.verdict === 'settled' || r.verdict === 'unfiled')).toBe(true);
    expect(rows.filter((r) => r.verdict === 'unfiled').length).toBe(1);
  });
});

describe('what counts as evidence', () => {
  // A caller who gives names but no paths has said nothing about folders. Two
  // files with no folder do not share one, and refusing a rename on that basis
  // would be refusing it for no reason.
  it('does not pair a file to a report just because neither has a path', () => {
    const reports = [{ name: 'NiSource Metric Dashboard.pbix' }];
    expect(reportReading('Unrelated Roster.xlsx', '', reports)).toBeNull();
  });

  it('still pairs by folder when there really is one', () => {
    const reports = [{ name: 'NiSource Metric Dashboard.pbix', path: '/lib/2. In Progress/x.pbix' }];
    const found = reportReading('Unrelated Roster.xlsx', '/lib/2. In Progress/Unrelated Roster.xlsx', reports);
    expect(found?.because).toBe('folder');
  });
});
