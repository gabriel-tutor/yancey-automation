import { describe, expect, it } from 'vitest';
import { applyApprovedPlan, applyProposals, approvedNotFound, approvedProgress, buildNamingRequest, namingCandidates } from './naming.js';
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

describe('renaming the reports themselves', () => {
  it('never offers a report by default', () => {
    expect(namingCandidates(freshRows(), REPORTS).map((r) => r.current)).not.toContain('NiSource Metric Dashboard.pbix');
  });

  /**
   * The reports in this library read nothing from it — checked by opening them,
   * not guessed at. What is still true is that a .pbix is referenced from
   * *outside* the library by whatever published dataset or refresh was set up
   * from it, and nothing here or inside the file can see that. So it is a
   * switch a person turns on, not a default.
   */
  it('offers one when the caller has decided to rename reports', () => {
    const names = namingCandidates(freshRows(), REPORTS, { includeReports: true }).map((r) => r.current);
    expect(names).toContain('NiSource Metric Dashboard.pbix');
  });

  it('holds an approved report name to the convention like any other', () => {
    const rows = freshRows();
    applyApprovedPlan(rows, [{ from: 'NiSource Metric Dashboard.pbix', to: 'not a valid name.pbix' }], {
      reports: REPORTS,
      includeReports: true,
    });
    const row = rows.find((r) => r.current.endsWith('.pbix'));
    expect(row?.proposed).toBe('');
  });

  it('applies an approved report name that does satisfy it', () => {
    const rows = freshRows();
    applyApprovedPlan(
      rows,
      [{ from: 'NiSource Metric Dashboard.pbix', to: 'NIPSCO_Metric-Dashboard_2026-01-01_Final_v1.0.pbix' }],
      { reports: REPORTS, includeReports: true },
    );
    const row = rows.find((r) => r.current.endsWith('.pbix'));
    expect(row?.proposed).toBe('NIPSCO_Metric-Dashboard_2026-01-01_Final_v1.0.pbix');
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
      ['category', 'date', 'document', 'file', 'folder', 'project', 'status', 'version', 'why'].sort(),
    );
    expect(req.text.format.schema.properties.proposals.items.properties.status.enum).toContain('Draft');
  });

  /**
   * Yancey on 3 September, when offered a template instead: *"No, I wanted to
   * give it freedom a little bit on the folder structure."* Freedom is the
   * model's half; "a little bit" is enforced in src/structure.js. What the
   * request has to do is make the job well-posed.
   */
  it('asks where each document belongs as well as what it is called', () => {
    const req = buildNamingRequest([]);
    const props = req.text.format.schema.properties.proposals.items.properties;
    expect(props.folder.description).toContain('Name only, never a path');
    expect(req.text.format.schema.required).toContain('folders');
    expect(req.instructions).toContain('FOUR OR FIVE folders across twenty documents');
    expect(req.instructions).toContain('a folder holding');
    expect(req.instructions).toContain('Never a numbered stage folder');
  });

  /** It cannot reuse a folder name it has never seen. */
  it('shows the folders that already exist, so reuse beats invention', () => {
    const text = firstText(
      buildNamingRequest([{ name: 'a.xlsx', path: '/x/a.xlsx' }], {
        folders: ['/x/Resources and References', '/x/1. Assigned Tasks'],
      }),
    );
    expect(text).toContain('/x/Resources and References');
    expect(text).toContain('Folders that already exist, paths copied exactly');
    expect(text.indexOf('/x/Resources')).toBeLessThan(text.indexOf('Files to name'));
  });

  it('says nothing about folders when the library has none', () => {
    const text = firstText(buildNamingRequest([{ name: 'a.xlsx', path: '/x/a.xlsx' }]));
    expect(text.startsWith('Files to name:')).toBe(true);
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

/**
 * The review page is what a person approves. If arming the switch re-runs the
 * model, what gets applied is not what was approved — and the model is not
 * deterministic: two runs an hour apart over the same unchanged library
 * proposed `Time-Zone-Conversion-Table` and then `Timezone-Conversion-Table`.
 * These hold the approved list to being the authority, and hold it to the same
 * guards a fresh proposal has to pass.
 */
/** The prose block of a request, which is where everything but the files goes. */
/** @param {any} req @returns {string} */
const firstText = (req) => String(req.input[0].content[0].text || '');

describe('what the model says about where a document belongs', () => {
  /** @param {any} extra */
  const named = (extra) => {
    const rows = freshRows();
    applyProposals(
      rows,
      [{ file: 'Metrics .xlsx', project: 'NIPSCO', document: 'Idling-Cost-Analysis', status: 'Review', date: '2026-03-02', version: 'v0.4', category: 'analysis', why: 'header', ...extra }],
      { files: FILES, reports: REPORTS },
    );
    return rows.find((r) => r.current === 'Metrics .xlsx');
  };

  it('records the folder it suggested', () => {
    expect(named({ folder: 'Dashboards' })?.folderSuggestion).toBe('Dashboards');
  });

  /**
   * Where a document belongs and what it is called are separate judgements, so
   * a name that fails the convention must not take a good filing decision down
   * with it.
   */
  it('keeps the folder even when the name it proposed is refused', () => {
    const row = named({ document: '', folder: 'Dashboards' });
    expect(row?.proposed, 'no document name means the model did not name it').not.toContain('NIPSCO');
    expect(row?.folderSuggestion).toBe('Dashboards');
  });

  it('records nothing when it made no suggestion', () => {
    expect(named({ folder: '' })?.folderSuggestion).toBe('');
  });
});

describe('applying a plan that was actually approved', () => {
  /** @param {any[]} rows */
  const withModelNames = (rows) =>
    applyProposals(
      rows,
      [
        { file: 'Metrics .xlsx', project: '', document: 'Whatever-The-Model-Says-Today', status: '', date: '2026-01-01', version: '', category: '', why: '' },
        { file: 'Screenshot 2025-04-06 at 6.34.18 PM.png', project: '', document: 'Also-Not-Approved', status: '', date: '2026-01-01', version: '', category: '', why: '' },
      ],
      { files: FILES, reports: REPORTS },
    );

  it('applies the approved name, not the one this run proposed', () => {
    const rows = withModelNames(freshRows());
    applyApprovedPlan(rows, [{ from: 'Metrics .xlsx', to: 'NIPSCO_Idling-Cost-Analysis_2026-03-02_Review_v0.4.xlsx' }], { reports: REPORTS });
    const row = rows.find((r) => r.current === 'Metrics .xlsx');
    expect(row?.proposed).toBe('NIPSCO_Idling-Cost-Analysis_2026-03-02_Review_v0.4.xlsx');
    expect(row?.namedBy).toBe('approved');
  });

  /**
   * The 3 September failure, in one test. Yancey dropped eight documents into
   * the library, ran the organiser, and got eight refusals reading "not on the
   * approved list, so this run leaves it alone" — the model had read and named
   * every one, and the gate discarded the names before he could look at them.
   * A plan governs what is APPLIED. It must not govern what may be PROPOSED,
   * or a file that arrives after the plan is signed can never be reviewed.
   */
  it('keeps a name nobody approved yet as a proposal rather than throwing it away', () => {
    const rows = withModelNames(freshRows());
    applyApprovedPlan(rows, [{ from: 'Metrics .xlsx', to: 'NIPSCO_Idling-Cost_2026-03-02_Review_v0.4.xlsx' }], { reports: REPORTS });
    const other = rows.find((r) => r.current.indexOf('6.34.18') !== -1);
    expect(other?.proposed).toBe('Also-Not-Approved_2026-01-01_Draft_v0.1.png');
    expect(other?.approved).toBe(false);
    expect(other?.blocked).toBe('');
    expect(other?.awaiting).toContain('approve it');
  });

  it('clears the approved one to be applied and leaves the rest waiting', () => {
    const rows = withModelNames(freshRows());
    applyApprovedPlan(rows, [{ from: 'Metrics .xlsx', to: 'NIPSCO_Idling-Cost_2026-03-02_Review_v0.4.xlsx' }], { reports: REPORTS });
    expect(rows.find((r) => r.current === 'Metrics .xlsx')?.approved).toBe(true);
    expect(rows.filter((r) => r.proposed && r.approved === true).length).toBe(1);
    expect(rows.filter((r) => r.proposed && r.approved === false).length).toBeGreaterThan(0);
  });

  /**
   * A plan that says nothing about where a file goes has not approved a
   * destination, so arming the move switch must move nothing. Saying so is what
   * `into` is for.
   */
  it('approves a destination only where the plan names one', () => {
    const rows = withModelNames(freshRows());
    applyApprovedPlan(rows, [{ from: 'Metrics .xlsx', to: 'NIPSCO_Idling-Cost_2026-03-02_Review_v0.4.xlsx' }], { reports: REPORTS });
    expect(rows.find((r) => r.current === 'Metrics .xlsx')?.approvedInto).toBeFalsy();

    const withInto = withModelNames(freshRows());
    applyApprovedPlan(
      withInto,
      [{ from: 'Metrics .xlsx', to: 'NIPSCO_Idling-Cost_2026-03-02_Review_v0.4.xlsx', into: ROOT + '/Task 2' }],
      { reports: REPORTS },
    );
    expect(withInto.find((r) => r.current === 'Metrics .xlsx')?.approvedInto).toBe(ROOT + '/Task 2');
  });

  /**
   * A folder a model suggests is no more deterministic than a name it writes.
   * Two runs an hour apart grouped the same two workbooks into `Dashboards` and
   * then did not, so the folder was created once and refused the next time by
   * the rule that a new folder must hold two documents. A line carrying only a
   * destination pins the filing without touching the name.
   */
  it('takes a line that approves where a document goes and nothing else', () => {
    const rows = withModelNames(freshRows());
    applyApprovedPlan(rows, [{ from: 'Metrics .xlsx', into: ROOT + '/Task 2' }], { reports: REPORTS });
    const row = rows.find((r) => r.current === 'Metrics .xlsx');
    expect(row?.approvedInto).toBe(ROOT + '/Task 2');
    // The name it proposed is untouched, and still unapproved.
    expect(row?.proposed).toBeTruthy();
    expect(row?.approved).toBe(false);
    expect(row?.namedBy).not.toBe('approved');
  });

  /**
   * The state after a plan has been applied. The renamed file comes back round
   * as a fresh proposal, that proposal is dropped because it is not on the
   * list, and the file must then read as already correct rather than blocked.
   */
  it('reports an already-applied file as correct, not as a problem', () => {
    const done = [{ name: 'SCE_Plan_2026-04-15_Draft_v0.1.docx', path: ROOT + '/Task 2/SCE_Plan_2026-04-15_Draft_v0.1.docx', modified: '2026-04-15T10:00:00Z' }];
    const rows = planRenames(done, {});
    applyProposals(
      rows,
      [{ file: done[0]?.name, project: 'SCE', document: 'Something-Else-Today', status: 'Draft', date: '2026-04-15', version: 'v0.1', category: '', why: '' }],
      { files: done, reports: [] },
    );
    expect(rows[0]?.proposed).toBeTruthy();
    applyApprovedPlan(rows, [{ from: 'Some Other File.xlsx', to: 'SCE_Other_2026-04-15_Draft_v0.1.xlsx' }], { reports: [] });
    expect(rows[0]?.proposed).toBe('');
    expect(rows[0]?.compliant).toBe(true);
    expect(rows[0]?.blocked).toBe('');
  });

  it('still refuses a guarded file, even when it is on the list', () => {
    const rows = freshRows();
    applyApprovedPlan(rows, [{ from: 'NiSource Metric Dashboard.pbix', to: 'NIPSCO_Metric-Dashboard_2026-01-01_Final_v1.0.pbix' }], { reports: REPORTS });
    const row = rows.find((r) => r.current.endsWith('.pbix'));
    expect(row?.proposed).toBe('');
  });

  it('refuses an approved name that would change the file extension', () => {
    const rows = freshRows();
    applyApprovedPlan(rows, [{ from: 'Metrics .xlsx', to: 'NIPSCO_Idling-Cost_2026-03-02_Review_v0.4.pdf' }], { reports: REPORTS });
    const row = rows.find((r) => r.current === 'Metrics .xlsx');
    expect(row?.proposed).toBe('');
    expect(row?.blocked).toContain('file extension');
  });

  it('refuses an approved name that does not satisfy the convention', () => {
    const rows = freshRows();
    applyApprovedPlan(rows, [{ from: 'Metrics .xlsx', to: 'Idling Cost Final_Final.xlsx' }], { reports: REPORTS });
    const row = rows.find((r) => r.current === 'Metrics .xlsx');
    expect(row?.proposed).toBe('');
    expect(row?.blocked).toContain('does not satisfy the convention');
  });

  it('refuses both when the approved list lands two files on one name', () => {
    const rows = freshRows();
    applyApprovedPlan(
      rows,
      [
        { from: 'Screenshot 2025-04-06 at 6.34.18 PM.png', to: 'NIPSCO_Chart_2026-01-01_Draft_v0.1.png' },
        { from: 'Screenshot 2025-04-06 at 6.34.50 PM.png', to: 'NIPSCO_Chart_2026-01-01_Draft_v0.1.png' },
      ],
      { reports: REPORTS },
    );
    for (const row of rows.filter((r) => r.current.indexOf('Screenshot') !== -1)) {
      expect(row.proposed).toBe('');
      expect(row.blocked).toContain('same name');
    }
  });

  it('marks a file compliant when the approved name is what it is already called', () => {
    const already = [{ name: 'SCE_Plan_2026-04-15_Draft_v0.1.docx', path: ROOT + '/Task 2/SCE_Plan_2026-04-15_Draft_v0.1.docx', modified: '2026-04-15T10:00:00Z' }];
    const rows = planRenames(already, {});
    const name = String(already[0]?.name);
    applyApprovedPlan(rows, [{ from: name, to: name }], { reports: [] });
    expect(rows[0]?.compliant).toBe(true);
    expect(rows[0]?.proposed).toBe('');
  });
});

/**
 * The check that would have caught a half-applied plan the first time.
 */
/**
 * Undo. The record of what changed is only an undo list if it can actually be
 * run, and it cannot be run through the ordinary path: every name it restores
 * is one the convention rejects, which is why it was changed in the first
 * place.
 */
describe('putting a name back', () => {
  const CURRENT = [{ name: 'SCE_Plan_2026-04-15_Draft_v0.1.docx', path: ROOT + '/Task 2/SCE_Plan_2026-04-15_Draft_v0.1.docx', modified: '2026-04-15T10:00:00Z' }];

  it('refuses to restore a messy name through the ordinary path', () => {
    const rows = planRenames(CURRENT, {});
    applyApprovedPlan(rows, [{ from: 'SCE_Plan_2026-04-15_Draft_v0.1.docx', to: 'Plan FINAL v2 copy.docx' }], { reports: [] });
    expect(rows[0]?.proposed).toBe('');
    expect(rows[0]?.blocked).toContain('does not satisfy the convention');
  });

  it('restores it when the caller says this is a revert', () => {
    const rows = planRenames(CURRENT, {});
    applyApprovedPlan(rows, [{ from: 'SCE_Plan_2026-04-15_Draft_v0.1.docx', to: 'Plan FINAL v2 copy.docx' }], { reports: [], revert: true });
    expect(rows[0]?.proposed).toBe('Plan FINAL v2 copy.docx');
  });

  it('still refuses to change the extension, revert or not', () => {
    const rows = planRenames(CURRENT, {});
    applyApprovedPlan(rows, [{ from: 'SCE_Plan_2026-04-15_Draft_v0.1.docx', to: 'Plan FINAL v2 copy.pdf' }], { reports: [], revert: true });
    expect(rows[0]?.proposed).toBe('');
    expect(rows[0]?.blocked).toContain('file extension');
  });

  it('still refuses two files landing on one name, revert or not', () => {
    const two = [
      { name: 'A_2026-04-15_Draft_v0.1.png', path: ROOT + '/T/A_2026-04-15_Draft_v0.1.png', modified: '2026-04-15T10:00:00Z' },
      { name: 'B_2026-04-15_Draft_v0.1.png', path: ROOT + '/T/B_2026-04-15_Draft_v0.1.png', modified: '2026-04-15T10:00:00Z' },
    ];
    const rows = planRenames(two, {});
    applyApprovedPlan(
      rows,
      [
        { from: 'A_2026-04-15_Draft_v0.1.png', to: 'Screenshot.png' },
        { from: 'B_2026-04-15_Draft_v0.1.png', to: 'Screenshot.png' },
      ],
      { reports: [], revert: true },
    );
    for (const row of rows) expect(row.proposed).toBe('');
  });
});

describe('an approved line whose file is not there', () => {
  it('is reported rather than silently skipped', () => {
    const rows = freshRows();
    const missing = approvedNotFound(rows, [
      { from: 'Metrics .xlsx', to: 'NIPSCO_Idling-Cost_2026-03-02_Review_v0.4.xlsx' },
      { from: 'Gone.xlsx', to: 'NIPSCO_Gone_2026-03-02_Review_v0.4.xlsx' },
    ]);
    expect(missing).toEqual(['Gone.xlsx']);
  });

  /**
   * The exact shape of the failure: macOS writes U+202F, a narrow no-break
   * space, before AM/PM. Retyping the name with an ordinary space produces a
   * line that looks identical and matches nothing.
   */
  it('catches a name retyped with an ordinary space where the file has U+202F', () => {
    const real = 'Screenshot 2025-04-06 at 6.34.18\u202fPM.png';
    const rows = planRenames([{ name: real, path: ROOT + '/Task 2/' + real, modified: '2026-04-15T10:00:00Z' }], {});
    const retyped = 'Screenshot 2025-04-06 at 6.34.18 PM.png';
    expect(retyped).not.toBe(real);
    expect(approvedNotFound(rows, [{ from: retyped, to: 'NIPSCO_Chart_2026-01-01_Draft_v0.1.png' }])).toEqual([retyped]);
    expect(approvedNotFound(rows, [{ from: real, to: 'NIPSCO_Chart_2026-01-01_Draft_v0.1.png' }])).toEqual([]);
  });

  it('says nothing when every approved line is present', () => {
    expect(approvedNotFound(freshRows(), [{ from: 'Metrics .xlsx', to: 'x' }])).toEqual([]);
  });
});

/**
 * Reporting "the file is not there" for a line that was carried out made every
 * run after a successful one look like a failure. Run 10302 said nine approved
 * files were "no longer in the library under that name" when all nine were
 * sitting there under exactly the name that had been approved.
 */
describe('where an approved line has got to', () => {
  const inLibrary = [
    { current: 'SCE_Plan_2026-04-15_Draft_v0.1.docx' },
    { current: 'Still Messy.xlsx' },
  ];

  it('calls a line whose result is in the library done, not missing', () => {
    const { done, missing } = approvedProgress(inLibrary, [
      { from: 'Plan FINAL v2 copy.docx', to: 'SCE_Plan_2026-04-15_Draft_v0.1.docx' },
    ]);
    expect(done).toEqual(['SCE_Plan_2026-04-15_Draft_v0.1.docx']);
    expect(missing).toEqual([]);
  });

  it('calls a line with neither its file nor its result adrift', () => {
    const { done, missing } = approvedProgress(inLibrary, [
      { from: 'Vanished.xlsx', to: 'SCE_Gone_2026-04-15_Draft_v0.1.xlsx' },
    ]);
    expect(done).toEqual([]);
    expect(missing).toEqual(['Vanished.xlsx']);
  });

  it('says nothing about a line this run is about to carry out', () => {
    expect(approvedProgress(inLibrary, [{ from: 'Still Messy.xlsx', to: 'x.xlsx' }])).toEqual({
      done: [],
      missing: [],
    });
  });

  /** A filing line has no result name to look for, so absence is absence. */
  it('treats a filing line with no name as adrift when its file is gone', () => {
    expect(approvedProgress(inLibrary, [{ from: 'Vanished.xlsx', into: '/a/b' }]).missing).toEqual([
      'Vanished.xlsx',
    ]);
  });

  it('is what approvedNotFound reports', () => {
    expect(
      approvedNotFound(inLibrary, [
        { from: 'Plan FINAL v2 copy.docx', to: 'SCE_Plan_2026-04-15_Draft_v0.1.docx' },
        { from: 'Vanished.xlsx', to: 'SCE_Gone_2026-04-15_Draft_v0.1.xlsx' },
      ]),
    ).toEqual(['Vanished.xlsx']);
  });
});
