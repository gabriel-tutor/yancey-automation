import { describe, expect, it } from 'vitest';
import { approvedFrom, mergeOverrides, offeredChanges } from './selfserve.js';

/**
 * The front end is the first part of this that a person who cannot read the
 * code will operate, which makes two questions worth more than the rest:
 * whether a form submission can talk the deployment into something it was not
 * built to do, and whether what somebody ticked is what actually happens.
 */
const LIB = '/sites/TaskManagement/Shared Documents/Power BI Task Management';

/** Index and find both return `undefined` to a type checker; a missing one is a failing test either way.
 * @template T @param {T | undefined} value @returns {T} */
const must = (value) => {
  if (value === undefined) throw new Error('expected the list to have it');
  return value;
};

const BASE = {
  host: 'https://upperformpartners.sharepoint.com',
  sitePath: '/sites/TaskManagement',
  libraryTitle: 'Documents',
  maxItems: 5000,
  maxChangesPerRun: 12,
  maxNewFolders: 6,
  minPerFolder: 2,
  assumeStatus: 'Draft',
  applyRenames: false,
  applyMoves: false,
  applyStructure: false,
  applyFolderRenames: false,
  applyProposed: false,
  revert: false,
  renameReports: true,
  approvedPlan: [{ from: 'baked.xlsx', to: 'Baked.xlsx' }],
  readContents: true,
  canArm: true,
};

describe('what a form submission may change about a run', () => {
  it('leaves the config alone when nothing is asked for', () => {
    const { config, refused } = mergeOverrides(BASE, {});
    expect(config).toEqual(BASE);
    expect(refused).toEqual([]);
  });

  it('arms a switch when the deployment allows it', () => {
    const { config } = mergeOverrides(BASE, { applyRenames: true });
    expect(config.applyRenames).toBe(true);
    expect(config.applyMoves).toBe(false);
  });

  it('refuses to arm anything on a copy that may not write, and says so', () => {
    const { config, refused } = mergeOverrides(
      Object.assign({}, BASE, { canArm: false }),
      { applyRenames: true, applyMoves: true }
    );
    expect(config.applyRenames).toBe(false);
    expect(config.applyMoves).toBe(false);
    expect(refused).toHaveLength(2);
    expect(refused[0]).toContain('not allowed to change the library');
  });

  it('takes canArm from the options when the caller states it', () => {
    const { config } = mergeOverrides(BASE, { applyRenames: true }, { canArm: false });
    expect(config.applyRenames).toBe(false);
  });

  it('will not let a form reach settings that belong to the deployment', () => {
    const { config, refused } = mergeOverrides(BASE, {
      host: 'https://evil.example.com',
      sitePath: '/sites/Somewhere',
      libraryTitle: 'Other',
      maxItems: 10,
      modelNaming: 'gpt-3',
    });
    expect(config.host).toBe(BASE.host);
    expect(config.sitePath).toBe(BASE.sitePath);
    expect(config.libraryTitle).toBe('Documents');
    expect(config.maxItems).toBe(5000);
    expect(refused).toHaveLength(5);
  });

  it('lets a caller switch a setting off, which is choosing to do less', () => {
    const { config, refused } = mergeOverrides(BASE, { renameReports: false, readContents: false });
    expect(config.renameReports).toBe(false);
    expect(config.readContents).toBe(false);
    expect(refused).toEqual([]);
  });

  it('refuses to switch one on that the deployment has off', () => {
    // Turning renameReports on puts the .pbix files in scope; turning
    // readContents on sends document text out of the tenant. A caller may
    // narrow what a run does and never widen it.
    const careful = Object.assign({}, BASE, { renameReports: false, readContents: false });
    const { config, refused } = mergeOverrides(careful, { renameReports: true, readContents: true });
    expect(config.renameReports).toBe(false);
    expect(config.readContents).toBe(false);
    expect(refused).toHaveLength(2);
    expect(refused[0]).toContain('but not on');
  });

  it('leaves one on when the deployment already has it on', () => {
    const { config, refused } = mergeOverrides(BASE, { renameReports: true, readContents: true });
    expect(config.renameReports).toBe(true);
    expect(config.readContents).toBe(true);
    expect(refused).toEqual([]);
  });

  it('lowers the cap but never raises it', () => {
    expect(mergeOverrides(BASE, { maxChangesPerRun: 4 }).config.maxChangesPerRun).toBe(4);
    const raised = mergeOverrides(BASE, { maxChangesPerRun: 500 });
    expect(raised.config.maxChangesPerRun).toBe(12);
    expect(raised.refused[0]).toContain('lowered but not raised');
  });

  it('keeps the built cap when the number is nonsense', () => {
    const { config, refused } = mergeOverrides(BASE, { maxChangesPerRun: 'lots' });
    expect(config.maxChangesPerRun).toBe(12);
    expect(refused[0]).toContain('not a number');
  });

  it('lowers a ceiling but never raises one', () => {
    expect(mergeOverrides(BASE, { maxNewFolders: 2 }).config.maxNewFolders).toBe(2);
    const raised = mergeOverrides(BASE, { maxNewFolders: 40 });
    expect(raised.config.maxNewFolders).toBe(6);
    expect(raised.refused[0]).toContain('lowered but not raised past 6');
  });

  it('raises a floor but never lowers one', () => {
    // minPerFolder runs the other way: a higher floor means fewer new folders,
    // so up is the conservative direction and down is the one to refuse.
    expect(mergeOverrides(BASE, { minPerFolder: 5 }).config.minPerFolder).toBe(5);
    const dropped = mergeOverrides(BASE, { minPerFolder: 1 });
    expect(dropped.config.minPerFolder).toBe(2);
    expect(dropped.refused[0]).toContain('raised but not lowered below 2');
  });

  it('takes a status the convention recognises', () => {
    expect(mergeOverrides(BASE, { assumeStatus: 'Review' }).config.assumeStatus).toBe('Review');
    expect(mergeOverrides(BASE, { assumeStatus: 'Final' }).config.assumeStatus).toBe('Final');
  });

  it('takes an empty status, which refuses to guess at all', () => {
    const { config, refused } = mergeOverrides(BASE, { assumeStatus: '' });
    expect(config.assumeStatus).toBe('');
    expect(refused).toEqual([]);
  });

  it('refuses a status the convention would then throw away', () => {
    const { config, refused } = mergeOverrides(BASE, { assumeStatus: 'Werking' });
    expect(config.assumeStatus).toBe('Draft');
    expect(refused[0]).toContain('must be one of');
  });

  it('replaces the baked plan rather than adding to it', () => {
    const { config } = mergeOverrides(BASE, { approvedPlan: [{ from: 'a.xlsx', to: 'B.xlsx' }] });
    expect(config.approvedPlan).toEqual([{ from: 'a.xlsx', to: 'B.xlsx' }]);
  });

  it('drops a plan line that asks for nothing', () => {
    const { config } = mergeOverrides(BASE, {
      approvedPlan: [{ from: 'a.xlsx' }, { from: '', to: 'B.xlsx' }, { from: 'c.xlsx', into: '/x' }],
    });
    expect(config.approvedPlan).toEqual([{ from: 'c.xlsx', into: '/x' }]);
  });

  it('reads overrides that arrive as a JSON string', () => {
    const { config } = mergeOverrides(BASE, JSON.stringify({ applyRenames: true }));
    expect(config.applyRenames).toBe(true);
  });

  it('applies nothing when the overrides will not parse', () => {
    const { config, refused } = mergeOverrides(BASE, '{not json');
    expect(config).toEqual(BASE);
    expect(refused[0]).toContain('not readable');
  });

  it('treats a missing payload as no overrides', () => {
    expect(mergeOverrides(BASE, null).config).toEqual(BASE);
    expect(mergeOverrides(BASE, []).config).toEqual(BASE);
  });
});

const RUN = {
  library: LIB,
  scanned: 27,
  compliant: 21,
  readContents: true,
  truncated: false,
  pages: 1,
  beyondModel: 0,
  newFolders: [LIB + '/Reliability Reviews'],
  folderRenames: [{ from: LIB + '/misc', to: 'Working Files' }],
  toRename: [
    { file: LIB + '/DIMP_Dashboard.xlsx', becomes: '2026-08-14_DIMP-Dashboard_Upperform_v0.1_Draft.xlsx' },
    { file: LIB + '/reports/Notes.docx', becomes: '2026-08-01_Notes_Upperform_v0.1_Draft.docx' },
  ],
  toMove: [
    { file: LIB + '/DIMP_Dashboard.xlsx', into: LIB + '/Reliability Reviews' },
    { file: LIB + '/Loose.pdf', into: LIB + '/Reliability Reviews' },
  ],
  refused: [{ file: 'weird~name.xlsx', because: 'the name has a character SharePoint will not take' }],
  folderRefusals: [],
  reportsUnreadable: [],
};

describe('a run turned into things to tick', () => {
  it('offers folders before the documents that go in them', () => {
    const offered = offeredChanges(RUN);
    expect(must(offered[0]).kind).toBe('folder');
    expect(must(offered[1]).kind).toBe('folderRename');
  });

  it('makes one tick out of a document that is renamed and filed', () => {
    const offered = offeredChanges(RUN);
    const both = offered.filter((o) => o.kind === 'rename+move');
    expect(both).toHaveLength(1);
    expect(must(both[0]).from).toBe(LIB + '/DIMP_Dashboard.xlsx');
    expect(must(both[0]).to).toContain('DIMP-Dashboard');
    expect(must(both[0]).into).toBe(LIB + '/Reliability Reviews');
    expect(must(both[0]).createsFolder).toBe(true);
  });

  it('offers a filing with no rename on its own', () => {
    const offered = offeredChanges(RUN);
    const move = must(offered.find((o) => o.kind === 'move'));
    expect(move.from).toBe(LIB + '/Loose.pdf');
  });

  it('numbers every label, so two files of one name are never the same tick', () => {
    const twins = {
      library: LIB,
      scanned: 2,
      compliant: 0,
      toRename: [
        { file: LIB + '/a/Report.xlsx', becomes: 'X.xlsx' },
        { file: LIB + '/b/Report.xlsx', becomes: 'X.xlsx' },
      ],
    };
    const offered = offeredChanges(twins);
    expect(must(offered[0]).label).not.toBe(must(offered[1]).label);
    expect(new Set(offered.map((o) => o.label)).size).toBe(2);
  });

  it('shows paths without the part every path shares', () => {
    const offered = offeredChanges(RUN);
    const move = must(offered.find((o) => o.kind === 'move'));
    expect(move.label).toContain('Loose.pdf');
    expect(move.label).not.toContain('/sites/TaskManagement');
  });

});

describe('ticks turned back into a plan', () => {
  const offered = offeredChanges(RUN);
  /** @param {string} kind */
  const label = (kind) => must(offered.find((o) => o.kind === kind)).label;

  it('binds the names that were on the page, not what a later run invents', () => {
    const { approvedPlan } = approvedFrom([label('rename')], offered);
    expect(approvedPlan).toEqual([
      { from: LIB + '/reports/Notes.docx', to: '2026-08-01_Notes_Upperform_v0.1_Draft.docx' },
    ]);
  });

  it('arms only the switch the tick needs', () => {
    const { overrides } = approvedFrom([label('rename')], offered);
    expect(overrides.applyRenames).toBe(true);
    expect(overrides.applyMoves).toBe(false);
    expect(overrides.applyStructure).toBe(false);
    expect(overrides.applyFolderRenames).toBe(false);
  });

  it('never arms applyProposed, because the second run may name things differently', () => {
    const { overrides } = approvedFrom(offered.map((o) => o.label), offered);
    expect(overrides.applyProposed).toBe(false);
  });

  it('agrees to create the folder when the filing needs one', () => {
    const { overrides } = approvedFrom([label('move')], offered);
    expect(overrides.applyMoves).toBe(true);
    expect(overrides.applyStructure).toBe(true);
  });

  it('does not create a folder for a filing into one that is already there', () => {
    const settled = { offered: offeredChanges(Object.assign({}, RUN, { newFolders: [] })) };
    const move = must(settled.offered.find((o) => o.kind === 'move'));
    const { overrides } = approvedFrom([move.label], settled.offered);
    expect(overrides.applyMoves).toBe(true);
    expect(overrides.applyStructure).toBe(false);
  });

  it('carries a rename and a filing as one line', () => {
    const { approvedPlan, overrides } = approvedFrom([label('rename+move')], offered);
    expect(approvedPlan).toEqual([
      {
        from: LIB + '/DIMP_Dashboard.xlsx',
        to: '2026-08-14_DIMP-Dashboard_Upperform_v0.1_Draft.xlsx',
        into: LIB + '/Reliability Reviews',
      },
    ]);
    expect(overrides.applyRenames).toBe(true);
    expect(overrides.applyMoves).toBe(true);
  });

  it('arms folder renaming only when a folder rename is ticked', () => {
    expect(approvedFrom([label('folderRename')], offered).overrides.applyFolderRenames).toBe(true);
    expect(approvedFrom([label('rename')], offered).overrides.applyFolderRenames).toBe(false);
  });

  it('reads ticks that come back as one comma-separated string', () => {
    const { chosen } = approvedFrom(label('rename') + ',' + label('move'), offered);
    expect(chosen).toHaveLength(2);
  });

  it('matches on the number, so a label that got reworded still lands', () => {
    const { chosen } = approvedFrom(['2. something else entirely'], offered);
    expect(must(chosen[0]).index).toBe(2);
  });

  it('reports a tick it cannot place rather than silently dropping it', () => {
    const { chosen, unknown } = approvedFrom(['not from this page'], offered);
    expect(chosen).toEqual([]);
    expect(unknown).toEqual(['not from this page']);
  });

  it('counts a change once however many times it was ticked', () => {
    const { chosen } = approvedFrom([label('rename'), label('rename')], offered);
    expect(chosen).toHaveLength(1);
  });

  it('arms nothing when nothing was ticked', () => {
    const { approvedPlan, overrides } = approvedFrom([], offered);
    expect(approvedPlan).toEqual([]);
    expect(overrides.applyRenames).toBe(false);
    expect(overrides.applyMoves).toBe(false);
    expect(overrides.applyStructure).toBe(false);
    expect(overrides.applyFolderRenames).toBe(false);
  });

  it('produces a plan the merge will accept unchanged', () => {
    const { overrides } = approvedFrom(offered.map((o) => o.label), offered);
    const { config, refused } = mergeOverrides(BASE, overrides);
    expect(refused).toEqual([]);
    expect(config.approvedPlan).toHaveLength(3);
    expect(config.applyRenames).toBe(true);
    expect(config.applyMoves).toBe(true);
  });
});

/**
 * Choosing which library to point at.
 *
 * Yancey asked for this on 4 September, watching a run finish: *"what happens
 * if I want to point it at another folder?"* — and it is the one setting the
 * whole file was written to refuse, because pointing a workflow that renames
 * documents at somewhere new is the definition of widening what a run does.
 *
 * So it opens under a policy the DEPLOYMENT states, never on a caller's say-so:
 *
 *   nothing stated          locked, exactly as before
 *   allowedSites: [...]     the caller may pick one of those
 *   allowAnySiteOnHost      any well-formed site, on the deployment's own host
 *
 * The host is never any of a caller's business under any of them. It is the
 * tenant, and no choice of tenant is a narrowing.
 */
describe('choosing which library a run points at', () => {
  const OPEN = Object.assign({}, BASE, { allowAnySiteOnHost: true });
  const LISTED = Object.assign({}, BASE, {
    allowedSites: ['/sites/TaskManagement', '/sites/WebExcel'],
  });

  it('stays locked when the deployment has said nothing, as it always was', () => {
    const { config, refused } = mergeOverrides(BASE, { sitePath: '/sites/Somewhere' });
    expect(config.sitePath).toBe(BASE.sitePath);
    expect(refused.join(' ')).toContain('this deployment is fixed to');
  });

  it('lets a caller pick a site the deployment listed', () => {
    const { config, refused } = mergeOverrides(LISTED, { sitePath: '/sites/WebExcel' });
    expect(config.sitePath).toBe('/sites/WebExcel');
    expect(refused).toEqual([]);
  });

  it('refuses a site the deployment did not list, and says which are allowed', () => {
    const { config, refused } = mergeOverrides(LISTED, { sitePath: '/sites/Payroll' });
    expect(config.sitePath).toBe(BASE.sitePath);
    expect(refused.join(' ')).toContain('/sites/TaskManagement');
    expect(refused.join(' ')).toContain('/sites/WebExcel');
  });

  it('lets any well-formed site through when the deployment opens the host', () => {
    const { config, refused } = mergeOverrides(OPEN, { sitePath: '/sites/AnythingElse' });
    expect(config.sitePath).toBe('/sites/AnythingElse');
    expect(refused).toEqual([]);
  });

  /**
   * The host is the tenant. The SharePoint credential cannot leave it, and a
   * caller naming a different one is either confused or trying something.
   */
  it('never takes the host, however open the deployment is', () => {
    const { config, refused } = mergeOverrides(OPEN, {
      host: 'https://evil.example.com',
      sitePath: '/sites/AnythingElse',
    });
    expect(config.host).toBe(BASE.host);
    expect(refused.join(' ')).toContain('host');
  });

  it.each([
    ['a full URL', 'https://upperformpartners.sharepoint.com/sites/X'],
    ['a climb out of the site collection', '/sites/TaskManagement/../../Payroll'],
    ['somewhere that is not a site', '/personal/yancey_partnersperform_com'],
    ['a bare word', 'TaskManagement'],
    ['nothing at all', '   '],
    ['a double slash', '/sites//TaskManagement'],
  ])('refuses %s as a site path', (_why, path) => {
    const { config, refused } = mergeOverrides(OPEN, { sitePath: path });
    expect(config.sitePath).toBe(BASE.sitePath);
    expect(refused.join(' ')).toContain('sitePath');
  });

  /**
   * The library title is interpolated into `getbytitle('...')` in the request
   * URL. While it was a constant that was nobody's problem; the moment a caller
   * can set it, an apostrophe closes the quote. Anything but plain characters
   * is refused rather than escaped, because refusing is checkable and escaping
   * is a thing you get wrong once.
   */
  it.each([
    ["an apostrophe, which closes the quote", "Docs') or 1 eq 1--"],
    ['a slash', 'Shared/Documents'],
    ['a percent escape', 'Docs%27'],
    ['empty', ''],
  ])('refuses a library title containing %s', (_why, title) => {
    const { config, refused } = mergeOverrides(OPEN, { libraryTitle: title });
    expect(config.libraryTitle).toBe('Documents');
    expect(refused.join(' ')).toContain('libraryTitle');
  });

  it('takes a plain library title', () => {
    const { config, refused } = mergeOverrides(OPEN, { libraryTitle: 'Shared Documents' });
    expect(config.libraryTitle).toBe('Shared Documents');
    expect(refused).toEqual([]);
  });

  /**
   * If the caller's spelling won, two spellings of one site would be two
   * different strings travelling into the request URL, and which library got
   * written to would depend on how somebody typed it.
   */
  it('answers with the deployment spelling of a listed site, not the caller\'s', () => {
    const { config, refused } = mergeOverrides(LISTED, { sitePath: '/SITES/webexcel' });
    expect(config.sitePath).toBe('/sites/WebExcel');
    expect(refused).toEqual([]);
  });

  it('keeps the deployment library when only the site is changed', () => {
    const { config } = mergeOverrides(OPEN, { sitePath: '/sites/WebExcel' });
    expect(config.sitePath).toBe('/sites/WebExcel');
    expect(config.libraryTitle).toBe('Documents');
  });

  /**
   * Changing where a run points must never be the thing that also arms it.
   * The ticks arm a run and nothing else does.
   */
  it('arms nothing by being pointed somewhere new', () => {
    const { config } = mergeOverrides(OPEN, { sitePath: '/sites/WebExcel', libraryTitle: 'Docs' });
    for (const key of ['applyRenames', 'applyMoves', 'applyStructure', 'applyFolderRenames', 'applyProposed']) {
      expect(/** @type {any} */ (config)[key]).toBe(false);
    }
  });
});
