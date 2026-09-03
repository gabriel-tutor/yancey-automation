/**
 * Fill the organiser plan artifact from a real run of WF-F.
 *
 *   node scripts/build-organiser-plan.mjs
 *
 * Same rule as the Gate 2 artifact: nothing is transcribed. Hand-copying the
 * Gate 2 drafts once dropped a sign-off and would have had Yancey mark the
 * system down for a fault it never committed. Here the risk is worse — a
 * mistyped file name in a rename plan is a plan for the wrong rename.
 *
 * This used to re-derive the plan locally by calling planRenames, which was
 * honest while the names came from rules. It is not any more: the names now
 * come from a model reading the documents, and no local rerun can reproduce
 * them. Deriving anything here would produce a page that agrees with itself and
 * with nothing else. So the page is built from the run WF-F actually did, saved
 * whole:
 *
 *   fixtures/sharepoint/organiser-run.json   the run — rows, folders, summary
 *   fixtures/sharepoint/renames-applied.json what has already been changed
 *
 * To refresh it: run WF-F, then save its `Review The Library`, `Apply Proposals`
 * and `Summary` node output into the run fixture.
 */
import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { folderOf, libraryRootOf } from '../src/library.js';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const runPath = join(root, 'fixtures/sharepoint/organiser-run.json');
const appliedPath = join(root, 'fixtures/sharepoint/renames-applied.json');
const htmlPath = join(root, 'docs/organiser-plan.html');

if (!existsSync(runPath)) {
  console.error('fixtures/sharepoint/organiser-run.json is missing. Run WF-F and save its output.');
  process.exit(2);
}

/**
 * @typedef {{
 *   current: string, path: string, compliant: boolean, proposed: string, blocked: string,
 *   assumptions: string[], namedBy?: string, why?: string, approved?: boolean, awaiting?: string,
 *   category?: string, categoryWhy?: string, placement?: string, placementWhy?: string,
 *   destination?: string, folderSuggestion?: string
 * }} Row
 */
/** @type {{ executionId: string, ranAt: string, summary: any, folders: string[], rows: Row[] }} */
const run = JSON.parse(readFileSync(runPath, 'utf8'));
const rows = run.rows || [];

/**
 * Why a file was refused, as a group a person can act on. The reason text is
 * the authority; this only decides which heading it sits under.
 *
 * @param {string} blocked
 * @returns {string}
 */
function refusalGroup(blocked) {
  if (blocked.indexOf('a Power BI report is referenced by path') !== -1) return 'is-a-report';
  if (blocked.indexOf('opens with the same words as the report') !== -1) return 'feeds-a-report';
  if (blocked.indexOf('shares a folder with the report') !== -1) return 'feeds-a-report';
  if (blocked.indexOf('is a date under one convention') !== -1) return 'ambiguous-date';
  if (blocked.indexOf('already carries a version') !== -1) return 'unreadable-version';
  if (blocked.indexOf('same name') !== -1) return 'collision';
  if (blocked.indexOf('no file extension') !== -1) return 'not-a-document';
  return 'other';
}

// The library root is the same for every row and only makes the page harder to
// read — but it is not a constant, and writing it as one made this a page about
// one library rather than a page about a library. It is whatever folder every
// path in the run happens to sit inside.
const libraryRoot = libraryRootOf([
  ...(run.folders || []).map(String),
  ...rows.map((r) => folderOf(String(r.path))),
]);

/** @param {string} path @returns {string} */
const withoutRoot = (path) => {
  const value = String(path);
  if (!libraryRoot) return value;
  if (value === libraryRoot) return '';
  return value.indexOf(libraryRoot + '/') === 0 ? value.slice(libraryRoot.length + 1) : value;
};

/**
 * Where a document lives, without the part of the path that is the same for
 * everything on the page.
 *
 * @param {string} path
 * @returns {string}
 */
function folderLabel(path) {
  return withoutRoot(folderOf(path)) || '(library root)';
}

/** @param {string} path */
const asFolderLabel = (path) => withoutRoot(path) || '(library root)';

// The folder tree, counted from the rows so the page and the run cannot
// disagree about how many files sit where. Folders the run listed but that hold
// no documents are kept, because an empty folder is a fact about the library.
/** @type {Record<string, { path: string, depth: number, files: number, kinds: Record<string, number> }>} */
const folders = {};
/** @param {string} label */
function folderEntry(label) {
  return (
    folders[label] ||
    (folders[label] = {
      path: label,
      depth: label === '(library root)' ? 0 : label.split('/').length - 1,
      files: 0,
      kinds: {},
    })
  );
}
for (const path of run.folders || []) folderEntry(asFolderLabel(path));
for (const row of rows) {
  const seen = folderEntry(folderLabel(row.path));
  seen.files += 1;
  const kind = row.category || 'other';
  seen.kinds[kind] = (seen.kinds[kind] || 0) + 1;
}

// What has actually been renamed, straight out of the runs that did it. This is
// also the undo list: every line is a rename that can be reversed by hand.
/** @type {Array<{ kind?: string, from: string, to: string, fromFolder?: string, assumptions: string[], applied: string, run: string }>} */
const history = existsSync(appliedPath) ? JSON.parse(readFileSync(appliedPath, 'utf8')) : [];
const byName = new Map(rows.map((r) => [r.current, r]));

/**
 * A file can be renamed more than once, and this library holds several that
 * were: the first pass gave them well-formed but mechanical names, and the
 * agent then gave them names that say what they hold. Showing both hops as two
 * separate changes is noise, and it breaks the check below, because the middle
 * name is not in the library any more.
 *
 * So chains are collapsed to what a person actually needs: what a file in the
 * library today was originally called, and when it last changed.
 */
// Only a rename can be superseded. A move keeps the document's name, so its
// `from` and `to` are the same string, and treating that as a chain would have
// each move superseding itself and vanishing off the page. A created folder is
// not a name change at all.
const renames = history.filter((r) => (r.kind || 'rename') === 'rename');
const supersededBy = new Map();
for (const r of renames) supersededBy.set(r.from, r);

/** @param {{ from: string, to: string }} record @returns {string} */
function originOf(record) {
  let from = record.from;
  const seen = new Set([from]);
  for (;;) {
    const earlier = renames.find((h) => h.to === from && h.from !== from);
    if (!earlier || seen.has(earlier.from)) return from;
    from = earlier.from;
    seen.add(from);
  }
}

/** @param {{ kind?: string }} r */
const isFolder = (r) => String(r.kind || '').indexOf('folder') !== -1;

// Only the last hop of each rename chain is shown, carrying the original name.
// A move is not part of that chain even though it sits in one: it changed a
// folder, not a name, and reporting it as "X became Y, in a new folder" would
// claim a rename that this entry did not do. It keeps its own name on both
// sides, which is what makes the page able to say what actually changed.
const applied = history
  .filter((r) => isFolder(r) || !supersededBy.has(r.to))
  .map((r) => ({ ...r, from: (r.kind || 'rename') === 'rename' ? originOf(r) : r.from }));

// A record claiming something the library does not hold would be worse than no
// record: this is the undo list. Folders are checked against the folder tree and
// documents against the documents, because they are different things.
const folderPaths = new Set((run.folders || []).map(String));
for (const r of applied) {
  const there = isFolder(r) ? folderPaths.has(String(r.to)) : byName.has(r.to);
  if (!there) {
    console.error(
      'the record claims a ' + (r.kind || 'rename') + ' to ' + r.to + ', which is not in the library.',
    );
    process.exit(2);
  }
}

const data = {
  library: 'Documents',
  site: '/sites/TaskManagement',
  generated: String(run.ranAt || '').slice(0, 10),
  execution: String(run.executionId || ''),
  naming: String(run.summary?.modelNote || ''),
  namedByModel: Number(run.summary?.namedByModel || 0),
  readContents: run.summary?.readContents === true,
  scanned: rows.length,
  // A pinned plan governs what is APPLIED, never what may be proposed. Without
  // this the page cannot tell "approved and waiting to run" from "the agent has
  // just read this and would like you to look" — and the run that could not
  // tell them apart reported eight freshly-named documents as refusals.
  pinned: run.summary?.pinned === true,
  // What the run could not see or could not fit. Left off the page, these are
  // the difference between "nothing else needed doing" and "nothing else was
  // looked at", which read identically and mean opposite things.
  truncated: run.summary?.truncated === true,
  listed: Number(run.summary?.listed || 0),
  pages: Number(run.summary?.pages || 1),
  beyondModel: Number(run.summary?.beyondModel || 0),
  // A group is a folder and everything that goes into it, a folder rename, or a
  // single document — the key says which, and the page should say it in the
  // words a person uses rather than the ones the run used.
  heldBack: (run.summary?.heldBack || []).map((/** @type {any} */ h) => {
    const key = String(h.group || '');
    const at = key.indexOf(':');
    const kind = at === -1 ? '' : key.slice(0, at);
    const path = at === -1 ? key : key.slice(at + 1);
    return {
      folder: kind === 'file' ? path.slice(path.lastIndexOf('/') + 1) : asFolderLabel(path),
      actions: Number(h.actions || 0),
    };
  }),
  unreachableFolders: (run.summary?.unreachableFolders || []).map(asFolderLabel),
  impossibleSwaps: (run.summary?.impossibleSwaps || []).map((/** @type {any} */ x) => ({
    file: String(x.file || ''),
    because: String(x.because || ''),
  })),
  // What it came up with when asked to organise rather than to rename.
  newFolders: (run.summary?.newFolders || []).map(asFolderLabel),
  folderRenames: (run.summary?.folderRenames || []).map((/** @type {any} */ r) => ({
    from: asFolderLabel(r.from),
    to: asFolderLabel(r.to),
    why: r.why || '',
  })),
  folderRefusals: [
    ...(run.summary?.folderRefusals || []),
    ...(run.summary?.folderRenameRefusals || []),
  ].map((/** @type {any} */ r) => ({ file: r.file || '', folder: r.folder || '', because: r.because || '' })),
  folders: Object.values(folders).sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0)),
  applied: applied.map((r) => ({
    kind: r.kind || 'rename',
    from: r.from,
    to: isFolder(r) ? asFolderLabel(String(r.to)) : r.to,
    fromFolder: r.fromFolder ? asFolderLabel(String(r.fromFolder)) : '',
    applied: r.applied,
    assumptions: r.assumptions,
    folder: isFolder(r)
      ? asFolderLabel(String(r.to))
      : folderLabel(String((byName.get(r.to) || { path: '' }).path)),
  })),
  rows: rows.map((r) => ({
    current: r.current,
    folder: folderLabel(r.path),
    compliant: r.compliant,
    proposed: r.proposed,
    blocked: r.blocked,
    group: r.blocked ? refusalGroup(r.blocked) : '',
    assumptions: r.assumptions || [],
    // Which of the two named it, and on what evidence. This is the whole point
    // of the rewrite, so it goes on the page rather than staying in the run.
    // Carried through even once the name has been applied — a file whose name
    // was written by reading it stays a file whose name was written by reading
    // it, and blanking it here made the page disagree with its own count.
    namedBy: r.namedBy || '',
    why: r.why || '',
    // Signed off, or still only proposed. Both are real states and the page has
    // to show them apart.
    approved: r.approved === true,
    awaiting: r.awaiting || '',
    kind: r.category || '',
    kindWhy: r.categoryWhy || '',
    placement: r.placement || '',
    placementWhy: r.placementWhy || '',
    into: r.destination ? asFolderLabel(folderOf(r.destination)) : '',
  })),
};

const html = readFileSync(htmlPath, 'utf8');
const OPEN = '<script type="application/json" id="data">';
const CLOSE = '</script>';
const start = html.indexOf(OPEN);
if (start === -1) {
  console.error('docs/organiser-plan.html has no data block to fill.');
  process.exit(2);
}
const end = html.indexOf(CLOSE, start);
const rebuilt = html.slice(0, start + OPEN.length) + '\n' + JSON.stringify(data, null, 2) + '\n' + html.slice(end);
writeFileSync(htmlPath, rebuilt);

const ready = data.rows.filter((r) => r.proposed && (!data.pinned || r.approved)).length;
const waiting = data.rows.filter((r) => r.proposed && data.pinned && !r.approved).length;
const refused = data.rows.filter((r) => r.blocked).length;
const correct = data.rows.filter((r) => r.compliant).length;
const moves = data.rows.filter((r) => r.placement === 'move' && r.into).length;
console.log(
  `scanned=${data.scanned} renamed=${data.applied.length} correct=${correct} approved=${ready} waiting=${waiting} ` +
    `(${data.namedByModel} named by the model) refused=${refused} moves=${moves} newFolders=${data.newFolders.length} ` +
    `folderRenames=${data.folderRenames.length} folders=${data.folders.length} -> docs/organiser-plan.html`,
);
