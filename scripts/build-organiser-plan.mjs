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
import { folderOf } from '../src/library.js';

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
 *   assumptions: string[], namedBy?: string, why?: string,
 *   category?: string, categoryWhy?: string, placement?: string, placementWhy?: string
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

/**
 * The library root is the same for every row and only makes the page harder to
 * read. Show where a file lives, not the whole server path.
 *
 * @param {string} path
 * @returns {string}
 */
function folderLabel(path) {
  const trimmed = folderOf(path).replace(/^\/sites\/TaskManagement\/Shared Documents\/?/, '');
  return trimmed || '(library root)';
}

/** @param {string} path */
const asFolderLabel = (path) =>
  String(path).replace(/^\/sites\/TaskManagement\/Shared Documents\/?/, '') || '(library root)';

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
const supersededBy = new Map();
for (const r of history) supersededBy.set(r.from, r);

/** @param {{ from: string, to: string }} record @returns {string} */
function originOf(record) {
  let from = record.from;
  const seen = new Set([from]);
  for (;;) {
    const earlier = history.find((h) => h.to === from);
    if (!earlier || seen.has(earlier.from)) return from;
    from = earlier.from;
    seen.add(from);
  }
}

// Only the last hop of each chain is shown, carrying the original name.
const applied = history
  .filter((r) => !supersededBy.has(r.to))
  .map((r) => ({ ...r, from: originOf(r) }));

for (const r of applied) {
  if (!byName.has(r.to)) {
    console.error('the record claims a rename to ' + r.to + ', which is not in the library.');
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
  folders: Object.values(folders).sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : 0)),
  applied: applied.map((r) => ({
    kind: r.kind || 'rename',
    from: r.from,
    to: r.to,
    fromFolder: r.fromFolder || '',
    applied: r.applied,
    assumptions: r.assumptions,
    folder: folderLabel(String((byName.get(r.to) || { path: '' }).path)),
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
    kind: r.category || '',
    kindWhy: r.categoryWhy || '',
    placement: r.placement || '',
    placementWhy: r.placementWhy || '',
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

const ready = data.rows.filter((r) => r.proposed).length;
const refused = data.rows.filter((r) => r.blocked).length;
const correct = data.rows.filter((r) => r.compliant).length;
const unfiled = data.rows.filter((r) => r.placement === 'unfiled').length;
console.log(
  `scanned=${data.scanned} renamed=${data.applied.length} correct=${correct} outstanding=${ready} (${data.namedByModel} named by the model) refused=${refused} unfiled=${unfiled} folders=${data.folders.length} -> docs/organiser-plan.html`,
);
