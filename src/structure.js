/**
 * Letting the model invent a folder structure, and holding it to a small one.
 *
 * Yancey asked for this directly on the 3 September call:
 *
 *   "Can it come up with its own folder structure and rename the files and put
 *    it in its own folder? Can it think and come up with its own folder
 *    structure?"
 *
 * and then answered the design question himself when offered a template
 * instead:
 *
 *   "No, I wanted to give it freedom a little bit on the folder structure."
 *
 * *A little bit* is the whole specification, and it is the part a machine has
 * to supply. A model given twenty documents and told to group them will happily
 * return twenty folders of one, which is not a structure — it is the same pile
 * with more clicks in it. So the model names the folders and the rules decide
 * which of them are allowed to exist:
 *
 *   a NEW folder must hold at least two documents   a folder of one is a pile
 *   there is a cap on how many appear in one run    freedom, a little bit
 *   it is created beside the documents it holds     never a global reshuffle
 *   the name has to be a name SharePoint accepts    and not collide
 *
 * An existing folder is exempt from the first rule: moving one document into a
 * folder somebody already made is filing, not restructuring.
 *
 * The model answers with a folder *name*, not a path, and a name is resolved to
 * a place in three steps: a folder of that name beside the document, then the
 * one folder of that name anywhere in the library, then a new folder beside the
 * document. Two folders sharing a name and neither beside the document is a
 * choice, and choosing belongs to a person — the same standard the placing
 * rules already hold themselves to.
 */
import { folderOf, isReport } from './library.js';

/** A numbered stage folder records how far work has got. Nothing here touches one. */
const STAGE_FOLDER = /^\d+[.)]?\s+\S/;

/** What SharePoint will not accept in a folder name, plus what reads as a mistake. */
const FORBIDDEN = /["*:<>?/\\|~#%{}]/;
const SHAPE = /^[A-Za-z0-9][A-Za-z0-9 &'_-]{1,48}[A-Za-z0-9)]$/;
const RESERVED = new Set(['.', '..', 'con', 'prn', 'aux', 'nul', 'forms', '_vti_']);

/**
 * Whether a name can be a folder, and what is wrong with it if not.
 *
 * @param {string} name
 * @returns {{ ok: boolean, reason: string }}
 */
export function safeFolderName(name) {
  const value = String(name || '').trim();
  if (!value) return { ok: false, reason: 'no folder name was given' };
  if (value !== String(name || '')) {
    return { ok: false, reason: 'a folder name cannot start or end with a space' };
  }
  if (value.length < 3) return { ok: false, reason: 'a folder name that short says nothing' };
  if (value.length > 50) return { ok: false, reason: 'too long to read in a breadcrumb' };
  if (FORBIDDEN.test(value)) return { ok: false, reason: 'contains a character SharePoint will not accept' };
  if (RESERVED.has(value.toLowerCase())) return { ok: false, reason: 'that name is reserved' };
  if (value.endsWith('.')) return { ok: false, reason: 'a folder name cannot end with a full stop' };
  if (STAGE_FOLDER.test(value)) {
    return { ok: false, reason: 'that looks like a numbered stage folder, which records progress rather than kind' };
  }
  if (!SHAPE.test(value)) return { ok: false, reason: 'letters, numbers, spaces, & and - only' };
  return { ok: true, reason: '' };
}

/**
 * The folders that are there, looked up without case and answered with the
 * spelling SharePoint actually holds.
 *
 * SharePoint compares paths without regard to case, so `Dashboards` and
 * `dashboards` are one folder. Treating them as two proposes a folder that
 * already exists and fails the create on a collision it could have seen. And
 * once matched, the answer has to be the path that exists rather than the one
 * suggested, or everything downstream carries a path the library does not have.
 *
 * @param {string[] | undefined} folders
 * @returns {(path: string) => string}
 */
function folderLookup(folders) {
  /** @type {Map<string, string>} */
  const byLower = new Map();
  for (const folder of folders || []) byLower.set(String(folder).toLowerCase(), String(folder));
  return (path) => byLower.get(String(path).toLowerCase()) || '';
}

/** @param {string} path @returns {string} */
function lastSegment(path) {
  const cut = String(path || '').lastIndexOf('/');
  return cut === -1 ? String(path || '') : String(path).slice(cut + 1);
}

/** @param {string} path @returns {boolean} */
function insideStageFolder(path) {
  return folderOf(path)
    .split('/')
    .some((segment) => STAGE_FOLDER.test(segment));
}

/** @param {string} folderPath @returns {boolean} */
function isOrInStageFolder(folderPath) {
  return String(folderPath)
    .split('/')
    .some((segment) => STAGE_FOLDER.test(segment));
}

/**
 * Decide which of the model's folder suggestions become real.
 *
 * Rows are mutated in place: an accepted suggestion sets `placement`,
 * `destination` and `placementWhy`, exactly as the deterministic placing rules
 * do, so everything downstream treats an invented folder and an existing one
 * the same way. Refusals are recorded rather than dropped, because a suggestion
 * that was thrown away for a reason is worth reading.
 *
 * @param {any[]} rows plan rows, carrying `folderSuggestion` from the model
 * @param {{
 *   folders?: string[],
 *   files?: Array<{ name: string, path: string }>,
 *   maxNewFolders?: number,
 *   minPerFolder?: number,
 *   includeReports?: boolean
 * }} [options]
 * @returns {{ create: string[], refused: Array<{ file: string, folder: string, because: string }>, used: string[] }}
 */
export function planStructure(rows, options) {
  const opts = options || {};
  const asItIs = folderLookup(opts.folders);
  const maxNew = opts.maxNewFolders === undefined ? 6 : Number(opts.maxNewFolders);
  const minPer = opts.minPerFolder === undefined ? 2 : Number(opts.minPerFolder);

  // A folder cannot take the name of a file already sitting beside it, and
  // SharePoint does not care which case that file's name is in.
  const takenByAFile = new Set(
    (opts.files || []).map((f) => (folderOf(String(f.path)) + '/' + String(f.name)).toLowerCase()),
  );

  /** @type {Array<{ file: string, folder: string, because: string }>} */
  const refused = [];
  /** @type {Map<string, any[]>} candidate folder path -> rows wanting it */
  const wanted = new Map();

  for (const row of Array.isArray(rows) ? rows : []) {
    const suggestion = String(row.folderSuggestion || '').trim();
    if (!suggestion) continue;

    const path = String(row.path || '');
    const parent = folderOf(path);

    if (insideStageFolder(path)) {
      refused.push({
        file: row.current,
        folder: suggestion,
        because:
          'it sits inside a numbered stage folder, and those record how far the work has got — nothing about a file says which stage it belongs in',
      });
      continue;
    }
    if (isReport(String(row.current)) && !opts.includeReports) {
      refused.push({
        file: row.current,
        folder: suggestion,
        because: 'moving a Power BI report changes its path exactly as renaming it would, and reports are not being changed on this run',
      });
      continue;
    }
    const safe = safeFolderName(suggestion);
    if (!safe.ok) {
      refused.push({ file: row.current, folder: suggestion, because: safe.reason });
      continue;
    }
    if (lastSegment(parent) === suggestion) {
      // Already in a folder of that name. Nothing to do, and not a refusal.
      continue;
    }

    // A name beside the document wins; then the one folder of that name
    // anywhere; otherwise a new folder beside the document.
    const beside = parent + '/' + suggestion;
    const here = asItIs(beside);
    const elsewhere = (opts.folders || [])
      .map(String)
      .filter((f) => f.toLowerCase() !== beside.toLowerCase() && lastSegment(f).toLowerCase() === suggestion.toLowerCase());
    let target = here || beside;
    if (!here) {
      if (elsewhere.length === 1) {
        target = String(elsewhere[0]);
      } else if (elsewhere.length > 1) {
        refused.push({
          file: row.current,
          folder: suggestion,
          because:
            elsewhere.length + ' folders in the library are called ' + suggestion + ', and choosing between them belongs to a person',
        });
        continue;
      }
    }

    if (isOrInStageFolder(target)) {
      refused.push({
        file: row.current,
        folder: suggestion,
        because: 'that would file it inside a numbered stage folder, which asserts how far the work has got',
      });
      continue;
    }
    if (takenByAFile.has(target.toLowerCase())) {
      refused.push({
        file: row.current,
        folder: suggestion,
        because: 'a file of that name already sits beside it, so the folder cannot be created',
      });
      continue;
    }
    if (target + '/' + String(row.current) === path) {
      // Resolved to the folder it is already in. Nothing to do.
      continue;
    }

    const group = wanted.get(target) || [];
    group.push(row);
    wanted.set(target, group);
  }

  /** @type {string[]} */
  const create = [];
  /** @type {string[]} */
  const used = [];

  // Biggest groups first, so the cap spends itself on the folders that are
  // doing the most work rather than on whichever happened to be listed first.
  const groups = [...wanted.entries()].sort((a, b) => b[1].length - a[1].length);

  for (const [target, group] of groups) {
    const alreadyThere = Boolean(asItIs(target));
    if (!alreadyThere && group.length < minPer) {
      for (const row of group) {
        refused.push({
          file: row.current,
          folder: lastSegment(target),
          because:
            'only this one document was put in it, and a folder holding one file is the pile it came from with an extra click in front',
        });
      }
      continue;
    }
    if (!alreadyThere && create.length >= maxNew) {
      for (const row of group) {
        refused.push({
          file: row.current,
          folder: lastSegment(target),
          because: 'this run had already proposed as many new folders as it is allowed to',
        });
      }
      continue;
    }
    if (!alreadyThere) create.push(target);
    used.push(target);

    for (const row of group) {
      row.placement = 'move';
      row.destination = target + '/' + String(row.current);
      row.placementWhy = alreadyThere
        ? 'it belongs with what is already in ' + lastSegment(target)
        : 'grouped with ' + (group.length - 1) + ' other document' + (group.length === 2 ? '' : 's') + ' into a new ' + lastSegment(target) + ' folder';
    }
  }

  return { create, refused, used };
}

/**
 * Folder renames the model has proposed, held to the same rules.
 *
 * A folder rename moves everything inside it, so the blast radius is the folder
 * multiplied by its contents. The extra guard beyond a file's: a folder holding
 * a Power BI report changes that report's own path, so it is only renameable on
 * a run that is already willing to rename reports.
 *
 * @param {Array<{ from: string, to: string, why?: string }>} proposals
 * @param {{ folders?: string[], files?: Array<{ name: string, path: string }>, includeReports?: boolean }} [options]
 * @returns {{ renames: Array<{ from: string, to: string, why: string }>, refused: Array<{ folder: string, because: string }> }}
 */
export function planFolderRenames(proposals, options) {
  const opts = options || {};
  const asItIs = folderLookup(opts.folders);
  const files = opts.files || [];
  /** @type {Array<{ from: string, to: string, why: string }>} */
  const renames = [];
  /** @type {Array<{ folder: string, because: string }>} */
  const refused = [];

  for (const proposal of Array.isArray(proposals) ? proposals : []) {
    const asked = String((proposal && proposal.from) || '');
    const to = String((proposal && proposal.to) || '').trim();
    if (!asked || !to) continue;

    // The model copies a path back; SharePoint holds one spelling of it. Move
    // the one that exists, not the one that was typed.
    const from = asItIs(asked);
    if (!from) {
      refused.push({ folder: asked, because: 'there is no folder at that path any more' });
      continue;
    }

    const current = lastSegment(from);
    if (current.toLowerCase() === to.toLowerCase()) continue;
    if (STAGE_FOLDER.test(current)) {
      refused.push({
        folder: current,
        because: 'it is a numbered stage folder, and the numbering is what makes the sequence readable',
      });
      continue;
    }
    // And nothing *inside* one either. Renaming a folder repaths every file in
    // it, so allowing it here would move documents the file-level rule refuses
    // to touch — the guard would hold one file at a time and be walked straight
    // past a folder at a time. It cost a good suggestion on the first live run:
    // the model wanted "1. Assigned Tasks/DIMP Folder" to become "DIMP
    // Materials", which is better, and also wanted "Task 1" to become
    // "Dashboard Drafts", which is a person's own task structure.
    if (isOrInStageFolder(folderOf(from))) {
      refused.push({
        folder: current,
        because:
          'it sits inside a numbered stage folder, and renaming a folder repaths everything in it — those documents are held back one at a time for the same reason',
      });
      continue;
    }
    const safe = safeFolderName(to);
    if (!safe.ok) {
      refused.push({ folder: current, because: safe.reason });
      continue;
    }

    const parent = folderOf(from);
    const target = parent + '/' + to;
    if (asItIs(target)) {
      refused.push({ folder: current, because: 'a folder called ' + to + ' already sits beside it' });
      continue;
    }

    // Everything inside moves with it, so a report inside is a report being
    // repathed.
    const holdsAReport = files.some(
      (f) => String(f.path).indexOf(from + '/') === 0 && isReport(String(f.name)),
    );
    if (holdsAReport && !opts.includeReports) {
      refused.push({
        folder: current,
        because: 'it holds a Power BI report, whose own path would change with it, and reports are not being changed on this run',
      });
      continue;
    }

    renames.push({ from, to: target, why: String((proposal && proposal.why) || '') });
  }

  return { renames, refused };
}
