import { describe, expect, it } from 'vitest';
import { applyProposals, buildNamingRequest, namingCandidates } from './naming.js';
import { planRenames } from './filenames.js';
import { reportsIn } from './library.js';

/**
 * The model is not under test here; the checking around it is.
 *
 * A model will occasionally answer with a project it invented, a status that is
 * not one of the four, a name that collides with another, or nothing useful at
 * all. None of those may reach the library. The proposals below are written to
 * be wrong in each of those ways on purpose.
 */
const ROOT = '/lib/Programme';
/** @param {string} folder @param {string} name */
const at = (folder, name) => ({
  name,
  path: folder ? ROOT + '/' + folder + '/' + name : ROOT + '/' + name,
  modified: '2026-04-15T10:00:00Z',
});

const FILES = [
  at('Task 2', 'Metrics .xlsx'),
  at('Task 2', 'Screenshot 2025-04-06 at 6.34.18 PM.png'),
  at('Task 2', 'Screenshot 2025-04-06 at 6.34.50 PM.png'),
  at('2. In Progress', 'NiSource Metric Dashboard.pbix'),
  at('2. In Progress', 'NiSource Metric Data.xlsx'),
];

/** @returns {any[]} */
const freshRows = () => planRenames(FILES, { assumeStatus: 'Draft', assumeVersion: { major: 0, minor: 1 } });
const REPORTS = reportsIn(FILES);

describe('which files a model is allowed to name', () => {
  it('never offers it a Power BI report', () => {
    const names = namingCandidates(freshRows(), REPORTS).map((r) => r.current);
    expect(names).not.toContain('NiSource Metric Dashboard.pbix');
  });

  it('never offers it a workbook a report reads', () => {
    const names = namingCandidates(freshRows(), REPORTS).map((r) => r.current);
    expect(names).not.toContain('NiSource Metric Data.xlsx');
  });

  it('does offer it the files a regex could not name', () => {
    const names = namingCandidates(freshRows(), REPORTS).map((r) => r.current);
    expect(names).toContain('Metrics .xlsx');
    expect(names).toContain('Screenshot 2025-04-06 at 6.34.18 PM.png');
  });
});

describe('turning the model’s words into a name', () => {
  it('writes a real project and document name where a regex could write neither', () => {
    const rows = freshRows();
    applyProposals(
      rows,
      [{ file: 'Metrics .xlsx', project: 'NIPSCO', document: 'Idling-Cost-Analysis', status: 'Review', date: '2026-03-02', version: 'v0.4', category: 'analysis', why: 'header says NIPSCO idling cost review' }],
      { files: FILES, reports: REPORTS },
    );
    const row = rows.find((r) => r.current === 'Metrics .xlsx');
    expect(row?.proposed).toBe('NIPSCO_Idling-Cost-Analysis_2026-03-02_Review_v0.4.xlsx');
    expect(row?.namedBy).toBe('model');
  });

  it('falls back to the file timestamp and Draft, and says both were assumed', () => {
    const rows = freshRows();
    applyProposals(
      rows,
      [{ file: 'Metrics .xlsx', project: '', document: 'Idling-Cost-Analysis', status: '', date: '', version: '', category: '', why: '' }],
      { files: FILES, reports: REPORTS },
    );
    const row = rows.find((r) => r.current === 'Metrics .xlsx');
    expect(row?.proposed).toBe('Idling-Cost-Analysis_2026-04-15_Draft_v0.1.xlsx');
    expect(row?.assumptions).toContain('date taken from the file timestamp, not the document');
    expect(row?.assumptions).toContain('no project code was evidenced, so none was added');
  });

  // His guide says v1.0 is the first approved version, so assuming v0.1 under
  // Final would produce a name his own rules call wrong.
  it('starts an approved document at v1.0 rather than v0.1', () => {
    const rows = freshRows();
    applyProposals(
      rows,
      [{ file: 'Metrics .xlsx', project: 'SCE', document: 'Rate-Case-Exhibit', status: 'Final', date: '2026-02-02', version: '', category: '', why: '' }],
      { files: FILES, reports: REPORTS },
    );
    expect(rows.find((r) => r.current === 'Metrics .xlsx')?.proposed).toBe(
      'SCE_Rate-Case-Exhibit_2026-02-02_Final_v1.0.xlsx',
    );
  });

  it('moves the evidence line when the model reclassifies a document', () => {
    const rows = freshRows();
    for (const row of rows) row.categoryWhy = 'from the .xlsx extension';
    applyProposals(
      rows,
      [{ file: 'Metrics .xlsx', project: '', document: 'Idling-Cost', status: '', date: '', version: '', category: 'Metric Tracker', why: 'it tracks metric deliverables' }],
      { files: FILES, reports: REPORTS },
    );
    const row = rows.find((r) => r.current === 'Metrics .xlsx');
    expect(row?.category).toBe('Metric Tracker');
    expect(row?.categoryWhy).toBe('from reading the document, not its extension');
  });

  it('gives two same-day screenshots names that tell them apart', () => {
    const rows = freshRows();
    applyProposals(
      rows,
      [
        { file: 'Screenshot 2025-04-06 at 6.34.18 PM.png', project: 'NIPSCO', document: 'CAIDI-Trend-Chart', status: '', date: '', version: '', category: 'screenshot', why: 'shows a CAIDI trend' },
        { file: 'Screenshot 2025-04-06 at 6.34.50 PM.png', project: 'NIPSCO', document: 'Outage-Cause-Breakdown', status: '', date: '', version: '', category: 'screenshot', why: 'shows outage causes' },
      ],
      { files: FILES, reports: REPORTS },
    );
    const one = rows.find((r) => r.current.indexOf('6.34.18') !== -1);
    const two = rows.find((r) => r.current.indexOf('6.34.50') !== -1);
    expect(one?.proposed).toContain('NIPSCO_CAIDI-Trend-Chart');
    expect(two?.proposed).toContain('NIPSCO_Outage-Cause-Breakdown');
    expect(one?.blocked).toBe('');
  });
});

describe('what the checking refuses', () => {
  it('refuses to rename a Power BI report even if the model answers for one', () => {
    const rows = freshRows();
    applyProposals(
      rows,
      [{ file: 'NiSource Metric Dashboard.pbix', project: 'NIPSCO', document: 'Metric-Dashboard', status: 'Final', date: '2026-01-01', version: 'v1.0', category: 'report', why: 'it is a dashboard' }],
      { files: FILES, reports: REPORTS },
    );
    const row = rows.find((r) => r.current.endsWith('.pbix'));
    expect(row?.proposed).toBe('');
    expect(row?.blocked).toContain('Power BI report');
  });

  it('refuses to rename a workbook a report reads', () => {
    const rows = freshRows();
    applyProposals(
      rows,
      [{ file: 'NiSource Metric Data.xlsx', project: 'NIPSCO', document: 'Metric-Source', status: '', date: '', version: '', category: '', why: '' }],
      { files: FILES, reports: REPORTS },
    );
    expect(rows.find((r) => r.current === 'NiSource Metric Data.xlsx')?.proposed).toBe('');
  });

  it('drops a status that is not one of the four rather than passing it through', () => {
    const rows = freshRows();
    applyProposals(
      rows,
      [{ file: 'Metrics .xlsx', project: '', document: 'Idling-Cost', status: 'WIP', date: '', version: '', category: '', why: '' }],
      { files: FILES, reports: REPORTS },
    );
    expect(rows.find((r) => r.current === 'Metrics .xlsx')?.proposed).toBe(
      'Idling-Cost_2026-04-15_Draft_v0.1.xlsx',
    );
  });

  it('ignores a date that is not a real YYYY-MM-DD', () => {
    const rows = freshRows();
    applyProposals(
      rows,
      [{ file: 'Metrics .xlsx', project: '', document: 'Idling-Cost', status: '', date: '04042025', version: '', category: '', why: '' }],
      { files: FILES, reports: REPORTS },
    );
    expect(rows.find((r) => r.current === 'Metrics .xlsx')?.proposed).toContain('_2026-04-15_');
  });

  it('leaves a file alone when the model returns no document name', () => {
    const rows = freshRows();
    const before = rows.find((r) => r.current === 'Metrics .xlsx')?.proposed;
    applyProposals(
      rows,
      [{ file: 'Metrics .xlsx', project: 'NIPSCO', document: '   ', status: '', date: '', version: '', category: '', why: '' }],
      { files: FILES, reports: REPORTS },
    );
    expect(rows.find((r) => r.current === 'Metrics .xlsx')?.proposed).toBe(before);
  });

  it('refuses both when the model lands two files on one name', () => {
    const rows = freshRows();
    applyProposals(
      rows,
      [
        { file: 'Screenshot 2025-04-06 at 6.34.18 PM.png', project: 'NIPSCO', document: 'Chart', status: '', date: '2026-01-01', version: '', category: '', why: '' },
        { file: 'Screenshot 2025-04-06 at 6.34.50 PM.png', project: 'NIPSCO', document: 'Chart', status: '', date: '2026-01-01', version: '', category: '', why: '' },
      ],
      { files: FILES, reports: REPORTS },
    );
    for (const row of rows.filter((r) => r.current.indexOf('Screenshot') !== -1)) {
      expect(row.proposed).toBe('');
      expect(row.blocked).toContain('same name');
    }
  });

  it('marks a file compliant rather than renaming it to what it already is', () => {
    const already = [{ name: 'SCE_Plan_2026-04-15_Draft_v0.1.docx', path: ROOT + '/Task 2/SCE_Plan_2026-04-15_Draft_v0.1.docx', modified: '2026-04-15T10:00:00Z' }];
    const rows = planRenames(already, {});
    applyProposals(
      rows,
      [{ file: String(already[0]?.name), project: 'SCE', document: 'Plan', status: 'Draft', date: '2026-04-15', version: 'v0.1', category: '', why: '' }],
      { files: already, reports: [] },
    );
    expect(rows[0]?.compliant).toBe(true);
    expect(rows[0]?.proposed).toBe('');
  });
});

describe('the request sent to the model', () => {
  it('carries his convention and his examples', () => {
    const req = buildNamingRequest([{ name: 'a.docx', path: '/x/a.docx', folder: 'x' }]);
    expect(req.instructions).toContain('[Project]_[Document Name]_[YYYY-MM-DD]_[Status]_v#.#');
    expect(req.instructions).toContain('SCE_PM_Ride-Along-Plan_2026-09-01_Draft_v0.1.docx');
    expect(req.instructions).toContain('An invented project code is worse than none');
  });

  it('asks for parts rather than a file name, so the shape is never its job', () => {
    const req = buildNamingRequest([]);
    const props = req.text.format.schema.properties.proposals.items.properties;
    expect(Object.keys(props).sort()).toEqual(
      ['category', 'date', 'document', 'file', 'project', 'status', 'version', 'why'].sort(),
    );
    expect(req.text.format.schema.properties.proposals.items.properties.status.enum).toContain('Draft');
  });

  it('sends an image to be looked at and text to be read', () => {
    const req = buildNamingRequest([
      { name: 'notes.docx', path: '/x/notes.docx', text: 'Crossarm survey, northern district.' },
      { name: 'shot.png', path: '/x/shot.png', dataUri: 'data:image/png;base64,AAAA', isImage: true },
      { name: 'report.pdf', path: '/x/report.pdf', dataUri: 'data:application/pdf;base64,BBBB' },
    ]);
    const content = req.input[0]?.content ?? [];
    const types = content.map((c) => c.type);
    expect(types).toContain('input_image');
    expect(types).toContain('input_file');
    expect(content[0]?.text).toContain('Crossarm survey, northern district.');
  });

  it('tells the model a file was not read rather than that it could not be read', () => {
    const req = buildNamingRequest([
      { name: 'a.docx', path: '/x/a.docx', contentNote: 'not read - reading contents is switched off' },
    ]);
    expect(req.input[0]?.content[0]?.text).toContain('reading contents is switched off');
  });

  it('caps how much of each document it sends', () => {
    const req = buildNamingRequest([{ name: 'big.docx', path: '/x/big.docx', text: 'x'.repeat(9000) }], { textPerFile: 100 });
    expect(String(req.input[0]?.content[0]?.text).length).toBeLessThan(600);
  });

  it('does not store the request on the provider', () => {
    expect(buildNamingRequest([]).store).toBe(false);
  });
});
