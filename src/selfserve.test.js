import { describe, expect, it } from 'vitest';
import { approvedFrom, mergeOverrides, outcomeHtml, reviewFields } from './selfserve.js';

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
    const { offered } = reviewFields(RUN);
    expect(must(offered[0]).kind).toBe('folder');
    expect(must(offered[1]).kind).toBe('folderRename');
  });

  it('makes one tick out of a document that is renamed and filed', () => {
    const { offered } = reviewFields(RUN);
    const both = offered.filter((o) => o.kind === 'rename+move');
    expect(both).toHaveLength(1);
    expect(must(both[0]).from).toBe(LIB + '/DIMP_Dashboard.xlsx');
    expect(must(both[0]).to).toContain('DIMP-Dashboard');
    expect(must(both[0]).into).toBe(LIB + '/Reliability Reviews');
    expect(must(both[0]).createsFolder).toBe(true);
  });

  it('offers a filing with no rename on its own', () => {
    const { offered } = reviewFields(RUN);
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
    const { offered } = reviewFields(twins);
    expect(must(offered[0]).label).not.toBe(must(offered[1]).label);
    expect(new Set(offered.map((o) => o.label)).size).toBe(2);
  });

  it('shows paths without the part every path shares', () => {
    const { offered } = reviewFields(RUN);
    const move = must(offered.find((o) => o.kind === 'move'));
    expect(move.label).toContain('Loose.pdf');
    expect(move.label).not.toContain('/sites/TaskManagement');
  });

  it('puts a checkbox on the page with one option per change', () => {
    const { fields, offered } = reviewFields(RUN);
    const box = must(fields.find((f) => f.fieldType === 'checkbox'));
    expect(box.fieldOptions.values).toHaveLength(offered.length);
    expect(box.requiredField).toBe(false);
  });

  it('offers nothing to tick when there is nothing to do', () => {
    const { fields, nothing } = reviewFields({ library: LIB, scanned: 27, compliant: 27 });
    expect(nothing).toBe(true);
    expect(fields.some((f) => f.fieldType === 'checkbox')).toBe(false);
  });

  it('says out loud when the run did not see the whole library', () => {
    const { fields } = reviewFields(Object.assign({}, RUN, { truncated: true, pages: 12 }));
    expect(must(fields[0]).html).toContain('did not see the whole library');
  });

  it('says when more documents could be renamed than were looked at', () => {
    const { fields } = reviewFields(Object.assign({}, RUN, { beyondModel: 8 }));
    expect(must(fields[0]).html).toContain('8 more document');
  });

  it('says when nothing inside a document was read', () => {
    const { fields } = reviewFields(Object.assign({}, RUN, { readContents: false }));
    expect(must(fields[0]).html).toContain('nothing inside a document was read');
  });

  it('shows what was left alone and why', () => {
    const { fields } = reviewFields(RUN);
    expect(must(fields[0]).html).toContain('weird~name.xlsx');
    expect(must(fields[0]).html).toContain('SharePoint will not take');
  });

  it('says one reason once, however many things it applies to', () => {
    // The first live run put the same sentence on the page eight times.
    const stage = 'it sits inside a numbered stage folder';
    const { fields } = reviewFields(Object.assign({}, RUN, {
      folderRefusals: [
        { name: 'Dashboards', because: stage },
        { name: 'TCO Program Materials', because: stage },
        { name: 'Task 1', because: stage },
        { name: 'Lonely', because: 'only one document was put in it' },
      ],
    }));
    const html = must(fields[0]).html;
    expect(html.split(stage)).toHaveLength(2);
    expect(html).toContain('Dashboards, TCO Program Materials, Task 1');
    expect(html).toContain('only one document was put in it');
  });

  it('names the same folder once even when it is refused twice', () => {
    const stage = 'inside a numbered stage folder';
    const { fields } = reviewFields(Object.assign({}, RUN, {
      folderRefusals: [
        { name: 'Dashboards', because: stage },
        { name: 'Dashboards', because: stage },
      ],
    }));
    expect(must(fields[0]).html.split('Dashboards')).toHaveLength(2);
  });

  it('leads with nothing needing changing when the library is already tidy', () => {
    const { fields, nothing } = reviewFields({ library: 'Documents', scanned: 27, compliant: 27 });
    expect(nothing).toBe(true);
    const html = must(fields[0]).html;
    expect(html).toContain('Nothing needs changing');
    expect(html).toContain('27 documents in the Documents library');
    expect(html).not.toContain('proposed');
  });

  it('escapes a file name that would otherwise be markup', () => {
    const { fields } = reviewFields(Object.assign({}, RUN, {
      refused: [{ file: '<img src=x onerror=alert(1)>', because: 'no' }],
    }));
    expect(must(fields[0]).html).not.toContain('<img');
    expect(must(fields[0]).html).toContain('&lt;img');
  });
});

describe('ticks turned back into a plan', () => {
  const { offered } = reviewFields(RUN);
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
    const settled = reviewFields(Object.assign({}, RUN, { newFolders: [] }));
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

describe('the page at the end', () => {
  it('puts failures before successes', () => {
    const html = outcomeHtml({
      applied: [{ what: 'renamed a.xlsx' }],
      failed: [{ what: 'b.xlsx', because: 'locked by another user' }],
    });
    expect(html.indexOf('could not be carried out')).toBeLessThan(html.indexOf('carried out.'));
    expect(html).toContain('locked by another user');
  });

  it('says plainly when nothing was ticked out of changes that were offered', () => {
    const html = outcomeHtml({ applied: [], failed: [] }, { chosen: [], proposed: 3 });
    expect(html).toContain('Nothing was ticked');
  });

  it('says nothing needed changing when nothing was ever offered', () => {
    // Not the same as ticking nothing, and the difference is the whole message:
    // one is a shrug, the other is the library being in good order.
    const html = outcomeHtml({ scanned: 27, compliant: 27, library: 'Documents', applied: [], failed: [] }, { chosen: [], proposed: 0 });
    expect(html).toContain('Nothing needs changing');
    expect(html).toContain('27 documents');
    expect(html).not.toContain('Nothing was ticked');
  });

  it('still reports what a tidy run refused to do', () => {
    const html = outcomeHtml(
      { scanned: 27, compliant: 27, applied: [], failed: [], folderRefusals: [{ name: 'Dashboards', because: 'it is a numbered stage folder' }] },
      { chosen: [], proposed: 0 }
    );
    expect(html).toContain('Dashboards');
    expect(html).toContain('numbered stage folder');
  });

  it('says that running again finishes what did not fit', () => {
    const html = outcomeHtml({ applied: [{ what: 'x' }], failed: [], heldBack: [{ what: 'y' }] });
    expect(html).toContain('Run it again');
  });

  it('repeats what the deployment refused to do', () => {
    const html = outcomeHtml({ applied: [], failed: [] }, {
      proposed: 1,
      refused: ['applyRenames was asked for, but this copy is not allowed to change the library'],
    });
    expect(html).toContain('not allowed to change the library');
  });

  it('escapes what it reports', () => {
    const html = outcomeHtml({ applied: [], failed: [{ what: '<b>x</b>', because: 'no' }] });
    expect(html).not.toContain('<b>x</b>');
  });
});
