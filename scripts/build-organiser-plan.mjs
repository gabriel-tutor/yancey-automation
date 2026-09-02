/**
 * Fill the organiser plan artifact from a real library listing.
 *
 *   node scripts/build-organiser-plan.mjs
 *
 * Same rule as the Gate 2 artifact: nothing is transcribed. Hand-copying the
 * Gate 2 drafts once dropped a sign-off and would have had Yancey mark the
 * system down for a fault it never committed. Here the risk is worse — a
 * mistyped file name in a rename plan is a plan for the wrong rename. So the
 * page is generated from the listing WF-F actually read, through the same
 * planRenames the workflow runs.
 *
 * The listing comes from WF-F's `List The Library` node, saved as
 * fixtures/sharepoint/task-management-listing.json.
 */
import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { planRenames } from '../src/filenames.js';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const listingPath = join(root, 'fixtures/sharepoint/task-management-listing.json');
const htmlPath = join(root, 'docs/organiser-plan.html');

if (!existsSync(listingPath)) {
  console.error('fixtures/sharepoint/task-management-listing.json is missing.');
  process.exit(2);
}

/** @type {Array<{ name: string, path: string, modified: string }>} */
const listed = JSON.parse(readFileSync(listingPath, 'utf8'));

// WF-F's Plan node drops folders by FSObjType. The listing keeps no such field
// once saved, and planRename refuses anything without an extension anyway, but
// filtering here keeps the counts on the page equal to the counts in the run.
const files = listed.filter((f) => /\.[A-Za-z0-9]+$/.test(String(f.name)));

const rows = planRenames(files, { assumeStatus: 'Draft', assumeVersion: { major: 0, minor: 1 } });

/**
 * Why a file was refused, as a group a person can act on. The reason text is
 * the authority; this only decides which heading it sits under.
 *
 * @param {string} blocked
 * @returns {string}
 */
function categorise(blocked) {
  if (blocked.indexOf('is a date under one convention') !== -1) return 'ambiguous-date';
  if (blocked.indexOf('already carries a version') !== -1) return 'unreadable-version';
  if (blocked.indexOf('a Power BI report is referenced by path') !== -1) return 'is-a-report';
  if (blocked.indexOf('opens with the same words as the report') !== -1) return 'feeds-a-report';
  if (blocked.indexOf('shares a folder with the report') !== -1) return 'feeds-a-report';
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
  const withoutFile = String(path).slice(0, String(path).lastIndexOf('/'));
  const trimmed = withoutFile.replace(/^\/sites\/TaskManagement\/Shared Documents\/?/, '');
  return trimmed || '(library root)';
}

const data = {
  library: 'Documents',
  site: '/sites/TaskManagement',
  generated: new Date().toISOString().slice(0, 10),
  scanned: rows.length,
  rows: rows.map((r) => ({
    current: r.current,
    folder: folderLabel(r.path),
    compliant: r.compliant,
    proposed: r.proposed,
    blocked: r.blocked,
    category: r.blocked ? categorise(r.blocked) : '',
    assumptions: r.assumptions,
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
console.log(`scanned=${data.scanned} correct=${correct} ready=${ready} refused=${refused} -> docs/organiser-plan.html`);
