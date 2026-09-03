import { describe, expect, it } from 'vitest';
import { fileChangeUrl, folderChangeUrl, folderCreateUrl, odataQueryValue, plannedChanges } from './changes.js';

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
