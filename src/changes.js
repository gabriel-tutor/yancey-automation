/**
 * Turning a reviewed plan into calls SharePoint will actually accept.
 *
 * This module exists because of two failures on 3 September, both of them in
 * the same place: the step that decides what to send.
 *
 * **A file needed two conflicting calls.** Renaming and moving were built as
 * separate action lists, so a document that wanted a new name *and* a new
 * folder produced two calls, both starting from the path it had before. The
 * first succeeded and the second could not: the file was no longer there. To
 * SharePoint a rename and a move are one operation — both change a path — so a
 * file gets exactly one call, to its new folder under its new name.
 *
 * **The URL ran out of room.** `moveto` carries the path twice, and ASP.NET
 * caps the URL *path* at 260 characters. Yancey created a folder called
 * "PD alignment Excel and Power BI - Karen", moved two files into it, and the
 * rename came back with *"The length of the URL for this request exceeds the
 * configured maxUrlLength"* at about 385 characters. The paths now travel in
 * the query string behind OData parameter aliases, where the ceiling is 2048.
 */
import { folderOf } from './library.js';

/** Characters that would end the value early, or mean something else, inside a query string. */
/** @type {Array<[string, string]>} */
const NEEDS_ESCAPING = [
  ["'", "''"], // OData escapes a quote by doubling it; must run before percent-encoding
  ['%', '%25'], // must run before anything that introduces a percent
  ['#', '%23'],
  ['&', '%26'],
  ['+', '%2B'],
  ['?', '%3F'],
  [' ', '%20'],
];

/**
 * A server-relative path, encoded to sit inside an OData string literal that is
 * itself inside a query string.
 *
 * Order matters twice over: the quote is doubled first because that is an OData
 * rule rather than an encoding one, and the percent is encoded before anything
 * that introduces one, or the escapes escape each other.
 *
 * @param {string} path
 * @returns {string}
 */
export function odataQueryValue(path) {
  let out = String(path);
  for (const [from, to] of NEEDS_ESCAPING) out = out.split(from).join(to);
  return out;
}

/**
 * The call that moves a file from one path to another.
 *
 * `flags=0` is the promise never to overwrite: if anything already sits on the
 * new path, SharePoint fails the call rather than replacing what is there.
 *
 * @param {string} base e.g. https://host/sites/Site/_api/web
 * @param {string} oldPath server-relative
 * @param {string} newPath server-relative
 * @returns {string}
 */
export function fileChangeUrl(base, oldPath, newPath) {
  return (
    String(base) +
    "/GetFileByServerRelativePath(decodedurl=@old)/moveto(newurl=@new,flags=0)?@old='" +
    odataQueryValue(oldPath) +
    "'&@new='" +
    odataQueryValue(newPath) +
    "'"
  );
}

/**
 * The call that creates a folder, if it is not already there.
 *
 * @param {string} base
 * @param {string} path server-relative path of the folder to create
 * @returns {string}
 */
export function folderCreateUrl(base, path) {
  return String(base) + "/folders/addUsingPath(DecodedUrl=@p)?@p='" + odataQueryValue(path) + "'";
}

/**
 * The call that renames a folder. A folder is moved exactly as a file is.
 *
 * @param {string} base
 * @param {string} oldPath
 * @param {string} newPath
 * @returns {string}
 */
export function folderChangeUrl(base, oldPath, newPath) {
  return (
    String(base) +
    "/GetFolderByServerRelativePath(decodedurl=@old)/moveto(newurl=@new,flags=0)?@old='" +
    odataQueryValue(oldPath) +
    "'&@new='" +
    odataQueryValue(newPath) +
    "'"
  );
}

/** @param {string} path @returns {string} */
function nameOf(path) {
  const cut = String(path).lastIndexOf('/');
  return cut === -1 ? String(path) : String(path).slice(cut + 1);
}

/**
 * One change per file, or none.
 *
 * A row can want a new name, a new folder, or both. Both is still one call, and
 * that is the whole point of this function: the destination is composed from
 * whichever parts are actually approved, and a row whose composed destination
 * equals where it already is produces nothing.
 *
 * `requireApproval` is what a pinned plan turns on. Without it, everything this
 * run proposes is fair game — right for a sandbox. With it, only what a person
 * signed off is applied, and the rest stays a proposal on the page rather than
 * disappearing.
 *
 * It binds **names**, because a list of names is what a plan is. It does not
 * bind folders: a plan that never mentions a destination has not refused one,
 * and reading it as a refusal would mean a pinned plan silently forbade every
 * folder the agent proposed — the same mistake as gating proposals, one level
 * up. Where a plan line *does* name a destination in `into`, that is the folder
 * used. Where it does not, the move is governed by `applyMoves`, which is the
 * switch a person arms after reading the page.
 *
 * @param {any[]} rows plan rows
 * @param {{ applyRenames?: boolean, applyMoves?: boolean, requireApproval?: boolean }} [options]
 * @returns {Array<{ kind: string, oldName: string, newName: string, oldPath: string, newPath: string, why: string[] }>}
 */
export function plannedChanges(rows, options) {
  const opts = options || {};
  const changes = [];

  for (const row of Array.isArray(rows) ? rows : []) {
    const path = String(row.path || '');
    if (!path) continue;

    const nameApproved = opts.requireApproval ? row.approved === true : true;

    const wantName = opts.applyRenames === true && Boolean(row.proposed) && nameApproved;
    const wantMove = opts.applyMoves === true && row.placement === 'move' && Boolean(row.destination);
    if (!wantName && !wantMove) continue;

    // A plan line that names a destination overrides what this run worked out.
    const into = row.approvedInto ? String(row.approvedInto) : folderOf(String(row.destination));
    const folder = wantMove ? into : folderOf(path);
    const name = wantName ? String(row.proposed) : nameOf(path);
    const newPath = folder + '/' + name;

    // Composing from parts can land exactly where the file already is — a move
    // to its own folder, say. That is not a change and must not be a call.
    if (newPath === path) continue;

    changes.push({
      kind: wantName && wantMove ? 'rename and move' : wantName ? 'rename' : 'move',
      oldName: nameOf(path),
      newName: name,
      oldPath: path,
      newPath,
      why: wantName ? row.assumptions || [] : [String(row.placementWhy || '')],
    });
  }

  return changes;
}
