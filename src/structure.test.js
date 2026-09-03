import { describe, expect, it } from 'vitest';
import { planFolderRenames, planStructure, safeFolderName } from './structure.js';

/**
 * Yancey asked for a structure the model invents rather than one he specifies,
 * and then bounded it himself: *"I wanted to give it freedom a little bit on
 * the folder structure."* These are the rules that make "a little bit" mean
 * something, tested against the library he was actually looking at.
 */
const LIB = '/sites/TaskManagement/Shared Documents/Power BI Task Management';
const REFS = LIB + '/Resources and References';

/** @param {string} name @param {string} folder @param {string} [suggestion] */
const at = (name, folder, suggestion) => ({
  current: name,
  path: folder + '/' + name,
  folderSuggestion: suggestion || '',
  placement: 'settled',
  placementWhy: 'filed',
  destination: '',
});

describe('what can be a folder name', () => {
  it('takes an ordinary name', () => {
    expect(safeFolderName('Dashboards').ok).toBe(true);
    expect(safeFolderName('Rate Case & Filings').ok).toBe(true);
  });

  it('refuses characters SharePoint will not accept', () => {
    for (const bad of ['A/B', 'A:B', 'A?B', 'A*B', 'A|B', 'A<B', 'A#B', 'A%B']) {
      expect(safeFolderName(bad).ok, bad).toBe(false);
    }
  });

  it('refuses a name that would read as a stage folder', () => {
    expect(safeFolderName('5. Archived').ok).toBe(false);
    expect(safeFolderName('5. Archived').reason).toContain('stage folder');
  });

  it('refuses leading or trailing spaces, and a trailing full stop', () => {
    expect(safeFolderName(' Dashboards').ok).toBe(false);
    expect(safeFolderName('Dashboards ').ok).toBe(false);
    expect(safeFolderName('Dashboards.').ok).toBe(false);
  });

  it('refuses something too short to mean anything, or too long to read', () => {
    expect(safeFolderName('AB').ok).toBe(false);
    expect(safeFolderName('D'.repeat(60)).ok).toBe(false);
  });
});

describe('inventing a folder structure', () => {
  /**
   * The failure mode that makes a model look clever and be useless: one folder
   * per document. The pile with extra clicks in front of it.
   */
  it('refuses a new folder that would hold a single document', () => {
    const rows = [at('a.xlsx', REFS, 'Dashboards'), at('b.xlsx', REFS, 'Rosters')];
    const plan = planStructure(rows, { folders: [REFS] });
    expect(plan.create).toEqual([]);
    expect(plan.refused.length).toBe(2);
    expect(plan.refused[0]?.because).toContain('holding one file is the pile it came from');
    expect(rows[0]?.placement).toBe('settled');
  });

  it('creates one when enough documents share it, and files them all', () => {
    const rows = [
      at('a.xlsx', REFS, 'Dashboards'),
      at('b.xlsx', REFS, 'Dashboards'),
      at('c.xlsx', REFS, 'Dashboards'),
    ];
    const plan = planStructure(rows, { folders: [REFS] });
    expect(plan.create).toEqual([REFS + '/Dashboards']);
    for (const row of rows) {
      expect(row.placement).toBe('move');
      expect(row.destination).toBe(REFS + '/Dashboards/' + row.current);
    }
    expect(rows[0]?.placementWhy).toContain('grouped with 2 other documents');
  });

  /** Filing one document into a folder somebody already made is not restructuring. */
  it('will put a single document into a folder that already exists', () => {
    const rows = [at('a.xlsx', LIB, 'Resources and References')];
    const plan = planStructure(rows, { folders: [LIB, REFS] });
    expect(plan.create).toEqual([]);
    expect(rows[0]?.placement).toBe('move');
    expect(rows[0]?.destination).toBe(REFS + '/a.xlsx');
    expect(rows[0]?.placementWhy).toContain('belongs with what is already in');
  });

  it('spends the cap on the folders doing the most work', () => {
    const rows = [
      at('a.xlsx', REFS, 'Big'),
      at('b.xlsx', REFS, 'Big'),
      at('c.xlsx', REFS, 'Big'),
      at('d.xlsx', REFS, 'Small'),
      at('e.xlsx', REFS, 'Small'),
    ];
    const plan = planStructure(rows, { folders: [REFS], maxNewFolders: 1 });
    expect(plan.create).toEqual([REFS + '/Big']);
    expect(plan.refused.map((r) => r.folder)).toEqual(['Small', 'Small']);
    expect(plan.refused[0]?.because).toContain('as many new folders as it is allowed to');
  });

  /** A stage folder records progress. Nothing in here reorganises one. */
  it('never restructures inside a numbered stage folder', () => {
    const stage = LIB + '/1. Assigned Tasks';
    const rows = [at('a.xlsx', stage, 'Dashboards'), at('b.xlsx', stage, 'Dashboards')];
    const plan = planStructure(rows, { folders: [stage] });
    expect(plan.create).toEqual([]);
    expect(plan.refused[0]?.because).toContain('how far the work has got');
  });

  it('leaves a Power BI report alone unless reports are in scope on this run', () => {
    const rows = [at('r.pbix', REFS, 'Dashboards'), at('s.pbix', REFS, 'Dashboards')];
    expect(planStructure(rows, { folders: [REFS] }).create).toEqual([]);
    const fresh = [at('r.pbix', REFS, 'Dashboards'), at('s.pbix', REFS, 'Dashboards')];
    expect(planStructure(fresh, { folders: [REFS], includeReports: true }).create).toEqual([
      REFS + '/Dashboards',
    ]);
  });

  it('refuses a folder whose name a file beside it already has', () => {
    const rows = [at('a.xlsx', REFS, 'Dashboards'), at('b.xlsx', REFS, 'Dashboards')];
    const plan = planStructure(rows, {
      folders: [REFS],
      files: [{ name: 'Dashboards', path: REFS + '/Dashboards' }],
    });
    expect(plan.create).toEqual([]);
    expect(plan.refused[0]?.because).toContain('a file of that name already sits beside it');
  });

  it('creates the folder beside the documents, never somewhere else in the library', () => {
    const rows = [at('a.xlsx', REFS, 'Trackers'), at('b.xlsx', REFS, 'Trackers')];
    planStructure(rows, { folders: [REFS] });
    expect(rows[0]?.destination.startsWith(REFS + '/')).toBe(true);
  });

  it('does nothing for a document already in a folder of that name', () => {
    const rows = [at('a.xlsx', REFS, 'Resources and References')];
    const plan = planStructure(rows, { folders: [REFS] });
    expect(plan.create).toEqual([]);
    expect(plan.refused).toEqual([]);
    expect(rows[0]?.placement).toBe('settled');
  });

  /**
   * The model answers with a name, not a path, so a name has to resolve to a
   * place. A folder of that name beside the document wins; failing that, the
   * one folder of that name anywhere in the library.
   */
  it('files into the one folder of that name anywhere in the library', () => {
    const rows = [at('a.xlsx', LIB + '/Task 2', 'Resources and References')];
    planStructure(rows, { folders: [LIB, LIB + '/Task 2', REFS] });
    expect(rows[0]?.placement).toBe('move');
    expect(rows[0]?.destination).toBe(REFS + '/a.xlsx');
  });

  it('prefers the folder beside the document over one further away', () => {
    const near = LIB + '/Task 2/Dashboards';
    const rows = [at('a.xlsx', LIB + '/Task 2', 'Dashboards')];
    planStructure(rows, { folders: [LIB, LIB + '/Task 2', near, LIB + '/Dashboards'] });
    expect(rows[0]?.destination).toBe(near + '/a.xlsx');
  });

  /** Two folders of one name is a choice, and choosing belongs to a person. */
  it('refuses when two folders share the suggested name and neither is beside it', () => {
    const rows = [at('a.xlsx', LIB + '/Task 2', 'Dashboards')];
    const plan = planStructure(rows, {
      folders: [LIB, LIB + '/Task 2', LIB + '/Dashboards', REFS + '/Dashboards'],
    });
    expect(plan.create).toEqual([]);
    expect(plan.refused[0]?.because).toContain('choosing between them belongs to a person');
  });

  /**
   * A stage-shaped name is refused by `safeFolderName` before it gets this far.
   * This is the other route in: an innocently named folder that happens to sit
   * inside a stage folder.
   */
  it('never files a document into a folder inside a numbered stage folder', () => {
    const inside = LIB + '/1. Assigned Tasks/Dashboards';
    const rows = [at('a.xlsx', LIB, 'Dashboards')];
    const plan = planStructure(rows, { folders: [LIB, LIB + '/1. Assigned Tasks', inside] });
    expect(plan.create).toEqual([]);
    expect(rows[0]?.placement).toBe('settled');
    expect(plan.refused[0]?.because).toContain('asserts how far the work has got');
  });

  /**
   * A pinned destination is a person's decision, and the model does not get to
   * overrule it. The order matters: the approved plan is read first and the
   * structure second, so without this the model's suggestion would land on top
   * of the plan's — and `row.destination` would then disagree with where the
   * call actually goes, which is a review page describing a move that is not
   * the move being made.
   */
  it('leaves alone a document whose destination a person already approved', () => {
    const rows = [
      { ...at('a.xlsx', REFS, 'Dashboards'), approvedInto: LIB + '/Task 2', placement: 'move', destination: LIB + '/Task 2/a.xlsx' },
      at('b.xlsx', REFS, 'Dashboards'),
      at('c.xlsx', REFS, 'Dashboards'),
    ];
    const plan = planStructure(rows, { folders: [LIB, REFS, LIB + '/Task 2'] });
    expect(rows[0]?.destination, 'still the approved one').toBe(LIB + '/Task 2/a.xlsx');
    expect(plan.create).toEqual([REFS + '/Dashboards']);
    expect(rows[1]?.destination).toBe(REFS + '/Dashboards/b.xlsx');
  });

  it('does not count an approved document towards a folder it is not going in', () => {
    const rows = [
      { ...at('a.xlsx', REFS, 'Dashboards'), approvedInto: LIB + '/Task 2', placement: 'move', destination: LIB + '/Task 2/a.xlsx' },
      at('b.xlsx', REFS, 'Dashboards'),
    ];
    const plan = planStructure(rows, { folders: [LIB, REFS, LIB + '/Task 2'] });
    expect(plan.create, 'only b wanted Dashboards, and one document is not a folder').toEqual([]);
  });

  it('ignores a document the model made no suggestion for', () => {
    const rows = [at('a.xlsx', REFS)];
    const plan = planStructure(rows, { folders: [REFS] });
    expect(plan.create).toEqual([]);
    expect(plan.refused).toEqual([]);
  });
});

describe('renaming a folder', () => {
  const KAREN = LIB + '/PD alignment Excel and Power BI - Karen';

  it('renames a messy one', () => {
    const plan = planFolderRenames([{ from: KAREN, to: 'PD Alignment' }], { folders: [LIB, KAREN] });
    expect(plan.renames).toEqual([{ from: KAREN, to: LIB + '/PD Alignment', why: '' }]);
  });

  it('never renames a numbered stage folder', () => {
    const stage = LIB + '/1. Assigned Tasks';
    const plan = planFolderRenames([{ from: stage, to: 'Assigned' }], { folders: [stage] });
    expect(plan.renames).toEqual([]);
    expect(plan.refused[0]?.because).toContain('numbering is what makes the sequence readable');
  });

  /**
   * Everything inside a folder moves with it, so a report inside is a report
   * being repathed — the same decision as renaming the report itself.
   */
  it('holds a folder containing a report to the same decision as the report', () => {
    const files = [{ name: 'Performance Management System.pbix', path: KAREN + '/Performance Management System.pbix' }];
    const guarded = planFolderRenames([{ from: KAREN, to: 'PD Alignment' }], { folders: [LIB, KAREN], files });
    expect(guarded.renames).toEqual([]);
    expect(guarded.refused[0]?.because).toContain('holds a Power BI report');

    const allowed = planFolderRenames([{ from: KAREN, to: 'PD Alignment' }], {
      folders: [LIB, KAREN],
      files,
      includeReports: true,
    });
    expect(allowed.renames.length).toBe(1);
  });

  /**
   * Found on the first live run: the model proposed renaming
   * "1. Assigned Tasks/Task 1" to "Dashboard Drafts". The folder itself is not
   * stage-shaped, but renaming it repaths every file inside it — files the
   * file-level rule refuses to touch. A guard that holds one file at a time and
   * is walked past a folder at a time is not a guard.
   */
  it('refuses a folder that sits inside a numbered stage folder', () => {
    const inside = LIB + '/1. Assigned Tasks/Task 1';
    const plan = planFolderRenames([{ from: inside, to: 'Dashboard Drafts' }], {
      folders: [LIB, LIB + '/1. Assigned Tasks', inside],
    });
    expect(plan.renames).toEqual([]);
    expect(plan.refused[0]?.because).toContain('repaths everything in it');
  });

  it('refuses when a folder of that name is already beside it', () => {
    const plan = planFolderRenames([{ from: KAREN, to: 'Resources and References' }], {
      folders: [LIB, KAREN, REFS],
    });
    expect(plan.renames).toEqual([]);
    expect(plan.refused[0]?.because).toContain('already sits beside it');
  });

  it('refuses a folder that is not there any more', () => {
    const plan = planFolderRenames([{ from: LIB + '/Gone', to: 'Somewhere' }], { folders: [LIB] });
    expect(plan.refused[0]?.because).toContain('no folder at that path any more');
  });

  it('ignores a rename to the name it already has', () => {
    const plan = planFolderRenames([{ from: REFS, to: 'Resources and References' }], { folders: [REFS] });
    expect(plan.renames).toEqual([]);
    expect(plan.refused).toEqual([]);
  });
});

/**
 * SharePoint compares paths without regard to case, so `Dashboards` and
 * `dashboards` are one folder. Treating them as two would have this proposing a
 * folder that already exists, and the create failing on a name collision it
 * could have seen coming.
 */
describe('folders whose names differ only in case', () => {
  it('files into the folder that is there rather than making another', () => {
    const rows = [at('a.xlsx', LIB, 'resources and references')];
    const plan = planStructure(rows, { folders: [LIB, REFS] });
    expect(plan.create).toEqual([]);
    expect(rows[0]?.destination, 'the path SharePoint actually holds').toBe(REFS + '/a.xlsx');
  });

  it('refuses a folder rename onto a name that differs only in case', () => {
    const KAREN = LIB + '/PD alignment Excel and Power BI - Karen';
    const plan = planFolderRenames([{ from: KAREN, to: 'RESOURCES AND REFERENCES' }], {
      folders: [LIB, KAREN, REFS],
    });
    expect(plan.renames).toEqual([]);
    expect(plan.refused[0]?.because).toContain('already sits beside it');
  });

  it('renames a folder the listing spelled differently', () => {
    const odd = LIB + '/pd alignment';
    const plan = planFolderRenames([{ from: LIB + '/PD Alignment', to: 'PD Alignment Notes' }], {
      folders: [LIB, odd],
    });
    expect(plan.renames.length, 'the folder is there, spelled another way').toBe(1);
    expect(plan.renames[0]?.from, 'moved from the path that exists').toBe(odd);
  });

  it('will not put a document beside a file of that name in any case', () => {
    const rows = [at('a.xlsx', REFS, 'Dashboards'), at('b.xlsx', REFS, 'Dashboards')];
    const plan = planStructure(rows, {
      folders: [REFS],
      files: [{ name: 'DASHBOARDS', path: REFS + '/DASHBOARDS' }],
    });
    expect(plan.create).toEqual([]);
    expect(plan.refused[0]?.because).toContain('a file of that name already sits beside it');
  });
});
