import { describe, expect, it } from 'vitest';
import {
  fileChangeUrl,
  fittingActions,
  folderChangeUrl,
  folderCreateUrl,
  foldersNeeded,
  odataQueryValue,
  plannedChanges,
  refuseImpossible,
} from './changes.js';

/**
 * Both of these tests exist because of a specific failure on 3 September, in
 * front of the client. Neither is hypothetical.
 */
const HOST = 'https://upperformpartners.sharepoint.com';
const BASE = HOST + '/sites/TaskManagement/_api/web';
const LIB = '/sites/TaskManagement/Shared Documents/Power BI Task Management';

/** @param {any} r @returns {any} */
const row = (r) => ({
  current: 'x.xlsx',
  path: LIB + '/x.xlsx',
  proposed: '',
  destination: '',
  placement: 'settled',
  assumptions: [],
  placementWhy: '',
  ...r,
});

describe('the URL that broke in the meeting', () => {
  /**
   * The exact failure: Yancey created "PD alignment Excel and Power BI - Karen",
   * moved two files into it, and the rename came back with "The length of the
   * URL for this request exceeds the configured maxUrlLength". ASP.NET caps the
   * URL *path* at 260 characters, and moveto carries the path twice.
   */
  it('keeps the path under the 260-character ceiling that failed', () => {
    const folder = LIB + '/PD alignment Excel and Power BI - Karen';
    const url = fileChangeUrl(
      BASE,
      folder + '/Performance Management System.xlsx',
      folder + '/Performance-Management-System-Workbook_2026-07-09_Draft_v0.1.xlsx',
    );
    const path = url.slice(0, url.indexOf('?'));
    expect(path.length, 'the part ASP.NET measures').toBeLessThan(260);
  });

  it('puts both paths in the query string, where the ceiling is 2048', () => {
    const url = fileChangeUrl(BASE, LIB + '/a b.xlsx', LIB + '/c d.xlsx');
    const query = url.slice(url.indexOf('?') + 1);
    expect(query).toContain('%20');
    expect(query.length).toBeLessThan(2048);
    expect(url).toContain('decodedurl=@old');
    expect(url).toContain('newurl=@new');
  });

  /** flags=0 is the promise never to overwrite, and it survives the rewrite. */
  it('still refuses to overwrite', () => {
    expect(fileChangeUrl(BASE, '/a/b.xlsx', '/a/c.xlsx')).toContain('flags=0');
  });
});

describe('encoding a path into a query string', () => {
  it('doubles a quote, the way OData wants', () => {
    expect(odataQueryValue("/a/Yancey's file.xlsx")).toBe("/a/Yancey''s%20file.xlsx");
  });

  /** Encode the percent first, or every later escape gets escaped again. */
  it('does not escape its own escapes', () => {
    expect(odataQueryValue('/a/50% done.xlsx')).toBe('/a/50%25%20done.xlsx');
  });

  it('encodes the characters that would end the value or start a new parameter', () => {
    expect(odataQueryValue('/a/A&B.xlsx')).toBe('/a/A%26B.xlsx');
    expect(odataQueryValue('/a/what?.xlsx')).toBe('/a/what%3F.xlsx');
    expect(odataQueryValue('/a/C#.xlsx')).toBe('/a/C%23.xlsx');
    expect(odataQueryValue('/a/a+b.xlsx')).toBe('/a/a%2Bb.xlsx');
  });

  it('leaves an ordinary path alone apart from its spaces', () => {
    expect(odataQueryValue('/sites/x/Shared Documents/a.xlsx')).toBe('/sites/x/Shared%20Documents/a.xlsx');
  });
});

describe('one change per file', () => {
  /**
   * The bug: a file wanting a new name and a new folder produced two calls,
   * both starting from the path it had before. Whichever ran second could not
   * find the file.
   */
  it('renames and moves in a single call, to the new folder under the new name', () => {
    const changes = plannedChanges(
      [
        row({
          current: 'CAIDI deliverables .docx',
          path: LIB + '/CAIDI deliverables .docx',
          proposed: 'NIPSCO_CAIDI-Deliverables_2026-01-01_Draft_v0.1.docx',
          placement: 'move',
          destination: LIB + '/Resources and References/CAIDI deliverables .docx',
        }),
      ],
      { applyRenames: true, applyMoves: true },
    );
    expect(changes.length).toBe(1);
    expect(changes[0]?.kind).toBe('rename and move');
    expect(changes[0]?.newPath).toBe(
      LIB + '/Resources and References/NIPSCO_CAIDI-Deliverables_2026-01-01_Draft_v0.1.docx',
    );
  });

  it('renames in place when only renaming is armed', () => {
    const changes = plannedChanges(
      [
        row({
          proposed: 'SCE_Plan_2026-01-01_Draft_v0.1.xlsx',
          placement: 'move',
          destination: LIB + '/Elsewhere/x.xlsx',
        }),
      ],
      { applyRenames: true, applyMoves: false },
    );
    expect(changes[0]?.kind).toBe('rename');
    expect(changes[0]?.newPath).toBe(LIB + '/SCE_Plan_2026-01-01_Draft_v0.1.xlsx');
  });

  it('moves under the old name when only moving is armed', () => {
    const changes = plannedChanges(
      [
        row({
          proposed: 'SCE_Plan_2026-01-01_Draft_v0.1.xlsx',
          placement: 'move',
          destination: LIB + '/Elsewhere/x.xlsx',
        }),
      ],
      { applyRenames: false, applyMoves: true },
    );
    expect(changes[0]?.kind).toBe('move');
    expect(changes[0]?.newPath).toBe(LIB + '/Elsewhere/x.xlsx');
  });

  it('does nothing at all when neither switch is armed', () => {
    expect(
      plannedChanges([row({ proposed: 'a_2026-01-01_Draft_v0.1.xlsx' })], {}),
    ).toEqual([]);
  });

  /** Composing from parts can land where the file already is. That is not a call. */
  it('emits nothing when the composed path is where the file already sits', () => {
    const changes = plannedChanges(
      [row({ placement: 'move', destination: LIB + '/x.xlsx' })],
      { applyRenames: true, applyMoves: true },
    );
    expect(changes).toEqual([]);
  });
});

describe('what a pinned plan gates', () => {
  /**
   * The failure Yancey saw: he dropped eight files in, the model named every
   * one, and all eight came back "not on the approved list, so this run leaves
   * it alone". A plan has to gate what is *applied* without discarding what is
   * *proposed*, or the review loop has nothing new to review.
   */
  const unapproved = row({
    current: 'Latest_ActionTracker_YD.xlsx',
    path: LIB + '/Latest_ActionTracker_YD.xlsx',
    proposed: 'NIPSCO_Action-Tracker_2026-02-01_Draft_v0.1.xlsx',
    approved: false,
  });

  it('does not apply a name nobody approved', () => {
    expect(plannedChanges([unapproved], { applyRenames: true, requireApproval: true })).toEqual([]);
  });

  it('applies one that was approved', () => {
    const changes = plannedChanges([{ ...unapproved, approved: true }], {
      applyRenames: true,
      requireApproval: true,
    });
    expect(changes[0]?.newName).toBe('NIPSCO_Action-Tracker_2026-02-01_Draft_v0.1.xlsx');
  });

  it('applies everything proposed when no plan is pinned', () => {
    const changes = plannedChanges([unapproved], { applyRenames: true, requireApproval: false });
    expect(changes.length).toBe(1);
  });

  /**
   * A plan is a list of names, so it binds names. Reading its silence about
   * folders as a refusal would mean pinning a plan forbade every folder the
   * agent proposed — the same mistake as gating proposals, one level up. The
   * folder is governed by its own switch, which is what a person arms after
   * reading the page.
   */
  it('still moves a file the plan says nothing about, when moving is armed', () => {
    const wantsMove = row({
      placement: 'move',
      destination: LIB + '/Resources and References/x.xlsx',
      approved: false,
    });
    const changes = plannedChanges([wantsMove], { applyMoves: true, requireApproval: true });
    expect(changes.length).toBe(1);
    expect(changes[0]?.newPath).toBe(LIB + '/Resources and References/x.xlsx');
    // ...but its name is untouched, because that is what the plan governs.
    expect(changes[0]?.kind).toBe('move');
  });

  it('uses the destination a plan line names, over the one this run worked out', () => {
    const changes = plannedChanges(
      [
        row({
          placement: 'move',
          destination: LIB + '/Somewhere The Model Chose/x.xlsx',
          approvedInto: LIB + '/Resources and References',
        }),
      ],
      { applyMoves: true, requireApproval: true },
    );
    expect(changes[0]?.newPath).toBe(LIB + '/Resources and References/x.xlsx');
  });

  it('moves nothing when moving is not armed, whatever the plan says', () => {
    const wantsMove = row({
      placement: 'move',
      destination: LIB + '/Resources and References/x.xlsx',
      approvedInto: LIB + '/Resources and References',
    });
    expect(plannedChanges([wantsMove], { applyMoves: false, requireApproval: true })).toEqual([]);
  });
});

describe('the folder calls', () => {
  it('creates a folder with its path in the query string', () => {
    const url = folderCreateUrl(BASE, LIB + '/Dashboards');
    expect(url.slice(0, url.indexOf('?')).length).toBeLessThan(260);
    expect(url).toContain("addUsingPath(DecodedUrl=@p)?@p='");
    expect(url).toContain('Power%20BI%20Task%20Management/Dashboards');
  });

  it('renames a folder the same way it moves a file, and never overwrites', () => {
    const url = folderChangeUrl(BASE, LIB + '/PD alignment Excel and Power BI - Karen', LIB + '/PD Alignment');
    expect(url).toContain('GetFolderByServerRelativePath');
    expect(url).toContain('flags=0');
    expect(url.slice(0, url.indexOf('?')).length).toBeLessThan(260);
  });
});

/**
 * Where a change would land, and whether it can.
 *
 * Everything above answers "what call do we send". These answer the question
 * that has to come first: would the call succeed? `flags=0` means SharePoint
 * refuses rather than overwrites, so a collision is a failed call rather than a
 * lost document — but a run that plans a call it knows will fail is not
 * reviewing anything, it is guessing and finding out in front of the client.
 */
describe('a change that cannot land', () => {
  const REFS = LIB + '/Resources and References';
  const DASH = REFS + '/Dashboards';

  /** @param {any[]} rows @param {any[]} [files] @param {any} [opts] */
  const refuse = (rows, files, opts) => {
    refuseImpossible(rows, files || [], opts);
    return rows;
  };

  it('refuses two documents that would land on one path', () => {
    const rows = [
      row({ current: 'a.xlsx', path: LIB + '/One/a.xlsx', placement: 'move', destination: DASH + '/a.xlsx' }),
      row({ current: 'a.xlsx', path: LIB + '/Two/a.xlsx', placement: 'move', destination: DASH + '/a.xlsx' }),
    ];
    refuse(rows);
    for (const r of rows) {
      expect(r.placement).toBe('settled');
      expect(r.destination).toBe('');
      expect(r.blocked).toContain('would land on the same path');
    }
  });

  /** Two files can share a name while they sit in different folders. */
  it('refuses a move onto a document that is already there', () => {
    const rows = [
      row({ current: 'a.xlsx', path: LIB + '/One/a.xlsx', placement: 'move', destination: DASH + '/a.xlsx' }),
    ];
    refuse(rows, [{ name: 'a.xlsx', path: DASH + '/a.xlsx' }]);
    expect(rows[0]?.placement).toBe('settled');
    expect(rows[0]?.blocked).toContain('already sits');
  });

  it('allows the move when the document there is itself moving away', () => {
    const rows = [
      row({ current: 'a.xlsx', path: LIB + '/One/a.xlsx', placement: 'move', destination: DASH + '/a.xlsx' }),
      row({ current: 'a.xlsx', path: DASH + '/a.xlsx', placement: 'move', destination: LIB + '/Two/a.xlsx' }),
    ];
    refuse(rows, [
      { name: 'a.xlsx', path: LIB + '/One/a.xlsx' },
      { name: 'a.xlsx', path: DASH + '/a.xlsx' },
    ]);
    expect(rows[0]?.placement).toBe('move');
    expect(rows[1]?.placement).toBe('move');
  });

  /** SharePoint compares paths without case, so these are the same path. */
  it('sees a collision that differs only in case', () => {
    const rows = [
      row({ current: 'a.xlsx', path: LIB + '/One/a.xlsx', placement: 'move', destination: DASH + '/A.XLSX' }),
    ];
    refuse(rows, [{ name: 'a.xlsx', path: DASH + '/a.xlsx' }]);
    expect(rows[0]?.blocked).toContain('already sits');
  });

  it('refuses a rename onto a name already in the same folder', () => {
    const rows = [
      row({ current: 'Messy.xlsx', path: REFS + '/Messy.xlsx', proposed: 'SCE_Plan_2026-01-01_Draft_v0.1.xlsx' }),
    ];
    refuse(rows, [{ name: 'SCE_Plan_2026-01-01_Draft_v0.1.xlsx', path: REFS + '/SCE_Plan_2026-01-01_Draft_v0.1.xlsx' }]);
    expect(rows[0]?.proposed).toBe('');
    expect(rows[0]?.blocked).toContain('already sits');
  });

  /**
   * SharePoint Online refuses a server-relative path over 400 characters, and
   * the failure it gives back says nothing about which document caused it.
   */
  it('refuses a path longer than SharePoint accepts', () => {
    const deep = LIB + '/' + 'A folder with quite a long name'.repeat(15);
    const rows = [row({ current: 'a.xlsx', path: LIB + '/a.xlsx', placement: 'move', destination: deep + '/a.xlsx' })];
    refuse(rows);
    expect(rows[0]?.blocked).toContain('longer than the 400 characters');
    expect(rows[0]?.placement).toBe('settled');
  });

  it('leaves a change that lands somewhere free entirely alone', () => {
    const rows = [
      row({
        current: 'a.xlsx',
        path: LIB + '/One/a.xlsx',
        proposed: 'SCE_Plan_2026-01-01_Draft_v0.1.xlsx',
        placement: 'move',
        destination: DASH + '/a.xlsx',
      }),
    ];
    refuse(rows, [{ name: 'a.xlsx', path: LIB + '/One/a.xlsx' }]);
    expect(rows[0]?.proposed).toBe('SCE_Plan_2026-01-01_Draft_v0.1.xlsx');
    expect(rows[0]?.placement).toBe('move');
    expect(rows[0]?.blocked).toBeFalsy();
  });

  /** A file that is not changing cannot collide with itself. */
  it('says nothing about a document that is staying put', () => {
    const rows = [row({ current: 'a.xlsx', path: LIB + '/a.xlsx' })];
    refuse(rows, [{ name: 'a.xlsx', path: LIB + '/a.xlsx' }]);
    expect(rows[0]?.blocked).toBeFalsy();
  });

  /** A plan line's destination is the one that gets used, so it is the one checked. */
  it('checks the destination a plan pinned, not the one the run worked out', () => {
    const rows = [
      row({
        current: 'a.xlsx',
        path: LIB + '/One/a.xlsx',
        placement: 'move',
        destination: LIB + '/Somewhere/a.xlsx',
        approvedInto: DASH,
      }),
    ];
    refuse(rows, [{ name: 'a.xlsx', path: DASH + '/a.xlsx' }]);
    expect(rows[0]?.blocked).toContain('already sits');
  });
});

describe('what fits in one run', () => {
  /** @param {string} group @param {string} name */
  const action = (group, name) => ({ group, newName: name });

  /**
   * A grouping is one folder plus every document that goes into it. Slicing the
   * list at the cap could cut between them, which creates a folder and leaves
   * the documents outside it — a state nobody chose and nobody asked for.
   */
  it('never sends half a folder', () => {
    const actions = [
      action('f:Dashboards', 'Dashboards'),
      action('f:Dashboards', 'a.xlsx'),
      action('f:Dashboards', 'b.xlsx'),
      action('one:c', 'c.xlsx'),
    ];
    const { send, held } = fittingActions(actions, 2);
    expect(send.map((a) => a.newName)).toEqual(['c.xlsx']);
    expect(held).toEqual([{ group: 'f:Dashboards', actions: 3 }]);
  });

  it('sends whole groups in order while they fit', () => {
    const actions = [
      action('f:Dashboards', 'Dashboards'),
      action('f:Dashboards', 'a.xlsx'),
      action('one:b', 'b.xlsx'),
    ];
    const { send, held } = fittingActions(actions, 3);
    expect(send.map((a) => a.newName)).toEqual(['Dashboards', 'a.xlsx', 'b.xlsx']);
    expect(held).toEqual([]);
  });

  /** Skipping rather than stopping means a run still makes progress. */
  it('carries on past a group that does not fit', () => {
    const actions = [
      action('f:Big', '1'),
      action('f:Big', '2'),
      action('f:Big', '3'),
      action('one:small', '4'),
    ];
    const { send } = fittingActions(actions, 2);
    expect(send.map((a) => a.newName)).toEqual(['4']);
  });

  it('sends nothing when there is no room at all', () => {
    expect(fittingActions([action('a', 'x')], 0).send).toEqual([]);
  });
});

describe('the folders the documents need', () => {
  const REFS = LIB + '/Resources and References';
  /** @param {string} into @param {Partial<any>} [over] */
  const heading = (into, over) =>
    row({ placement: 'move', destination: into + '/x.xlsx', ...over });

  it('creates every missing level, shallowest first', () => {
    const rows = [heading(REFS + '/Dashboards/Weekly')];
    expect(foldersNeeded(rows, [LIB, REFS]).create).toEqual([
      REFS + '/Dashboards',
      REFS + '/Dashboards/Weekly',
    ]);
  });

  /**
   * Written the narrow way first — only folders a pinned plan named — and a
   * move to a folder nothing had thought to create went straight through to a
   * call that could only fail.
   */
  it('names a folder the run invented, not only one a plan pinned', () => {
    const rows = [heading(REFS + '/Dashboards')];
    expect(rows[0]?.approvedInto).toBeUndefined();
    expect(foldersNeeded(rows, [LIB, REFS]).create).toEqual([REFS + '/Dashboards']);
  });

  it('matches an existing folder without case', () => {
    const rows = [heading(LIB + '/RESOURCES AND REFERENCES')];
    expect(foldersNeeded(rows, [LIB, REFS]).create).toEqual([]);
  });

  /**
   * A destination with no known ancestor cannot be created safely: the walk up
   * would end at /sites, and this has no business making site collections.
   */
  it('refuses a destination with no ancestor it knows exists', () => {
    const rows = [heading('/sites/Elsewhere/Shared Documents/Somewhere')];
    const { create, unreachable } = foldersNeeded(rows, [LIB, REFS]);
    expect(create).toEqual([]);
    expect(unreachable).toEqual(['/sites/Elsewhere/Shared Documents/Somewhere']);
  });

  it('names each folder once however many documents want it', () => {
    const rows = [heading(REFS + '/Dashboards'), heading(REFS + '/Dashboards')];
    expect(foldersNeeded(rows, [LIB, REFS]).create).toEqual([REFS + '/Dashboards']);
  });

  it('uses the folder a plan pinned over the one the run worked out', () => {
    const rows = [heading(LIB + '/Somewhere', { approvedInto: REFS + '/Dashboards' })];
    expect(foldersNeeded(rows, [LIB, REFS]).create).toEqual([REFS + '/Dashboards']);
  });

  it('says nothing about a document that is staying where it is', () => {
    expect(foldersNeeded([row({})], [LIB]).create).toEqual([]);
  });
});

/**
 * Names that are not plain ASCII.
 *
 * The escaping here covers what would end an OData string literal or start a new
 * query parameter. Everything else — accented letters, CJK, an em-dash in a
 * title — is left exactly as it is, because the HTTP layer percent-encodes a URL
 * on the way out and encoding it twice would send the escape rather than the
 * character. This is a decision, not an omission.
 */
describe('a file name that is not plain ASCII', () => {
  it('passes accented and non-Latin characters through untouched', () => {
    expect(odataQueryValue('/a/Résumé — 日本語.docx')).toBe('/a/Résumé%20—%20日本語.docx');
  });

  it('still escapes what would break the request around them', () => {
    expect(odataQueryValue("/a/Café & Bar's 50%.docx")).toBe("/a/Café%20%26%20Bar''s%2050%25.docx");
  });
});
