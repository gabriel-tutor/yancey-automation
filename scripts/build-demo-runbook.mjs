/**
 * Fill the demo runbook from what the library actually holds.
 *
 *   node scripts/build-demo-runbook.mjs
 *
 * Same rule as everywhere else here: no file name on this page is typed by
 * hand. The page is read aloud in front of the client, so a runbook that quotes
 * a rename the workflow is not about to perform is worse than no runbook.
 *
 * It used to be built from config/approved-plan.json alone, which was right
 * while that plan was pending and became a lie the moment it was applied: on
 * 2026-09-04 the page still promised fifteen renames that had already happened,
 * against a library that was by then 100% compliant. A plan is what somebody
 * approved; only the library says what is still true. So this reads the
 * measured listing as well, and REFUSES to build a page claiming a rename whose
 * new name is already sitting in the library.
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { nameChanges, reconcile, renameChains } from '../src/undo.js';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const read = (/** @type {string} */ p) => JSON.parse(readFileSync(join(root, p), 'utf8'));

const plan = read('config/approved-plan.json');
const undo = read('fixtures/sharepoint/renames-applied.json');
const listing = read('fixtures/sharepoint/library-2026-09-04.json');
const htmlPath = join(root, 'docs/demo-runbook.html');

/** @type {string[]} */
const live = listing.documents.map((/** @type {any} */ d) => String(d.name));
const liveSet = new Set(live);

/**
 * Which of four states a plan line is in, decided by the library rather than by
 * the plan. Three of these were obvious; the fourth was not, and the guard
 * below found it.
 *
 *   pending    the old name is there and the new one is free - it can be applied
 *   applied    the old name is gone and the new one is there - it has happened
 *   colliding  BOTH are there - it cannot be applied, and the workflow's own
 *              flags=0 would refuse it rather than overwrite
 *   gone       neither is there - the document has left the library
 *
 * `colliding` turned up on 2026-09-04, when a document that had already been
 * renamed was copied back into the library under its original name during a
 * demo. Calling that pending would put a rename on the page that the workflow
 * is about to refuse, which is the same lie in a subtler form.
 *
 * @param {any} line
 */
function stateOf(line) {
  const from = liveSet.has(String(line.from));
  const to = liveSet.has(String(line.to));
  if (from && !to) return 'pending';
  if (!from && to) return 'applied';
  if (from && to) return 'colliding';
  return 'gone';
}

const pending = plan.plan.filter((/** @type {any} */ l) => stateOf(l) === 'pending');
const applied = plan.plan.filter((/** @type {any} */ l) => stateOf(l) === 'applied');
const colliding = plan.plan.filter((/** @type {any} */ l) => stateOf(l) === 'colliding');

// The guard, kept even though the split above makes it unreachable: a page
// promising a rename that has already happened is the exact failure this file
// was written to prevent, and an assertion that cannot fire is cheaper than
// finding out it could.
const alreadyDone = pending.filter((/** @type {any} */ l) => liveSet.has(String(l.to)));
if (alreadyDone.length > 0) {
  console.error('refusing to build: the plan claims these are pending, but the new name is already in the library:');
  for (const line of alreadyDone) console.error('  ' + line.from + ' -> ' + line.to);
  process.exit(1);
}

const record = reconcile(undo, live, plan.plan);
if (record.deadEnds.length > 0 || record.unrecorded.length > 0) {
  console.error('refusing to build: the undo list does not reconcile with the library.');
  console.error('  dead ends: ' + record.deadEnds.length + ', unrecorded renames: ' + record.unrecorded.length);
  console.error('  run the undo tests (src/undo.test.js) — the page must not promise an undo that is incomplete.');
  process.exit(1);
}

// Every chain, with the disclosure the run made at the time.
const changes = nameChanges(undo);
const byFrom = new Map(changes.map((/** @type {any} */ e) => [String(e.from), e]));
const history = renameChains(undo)
  .filter((c) => liveSet.has(c.current))
  .map((c) => {
    const first = byFrom.get(c.original) || {};
    return {
      original: c.original,
      current: c.current,
      hops: c.hops,
      assumptions: first.assumptions || [],
      reconstructed: first.reconstructed === true,
    };
  })
  .sort((a, b) => a.current.localeCompare(b.current));

const moves = undo
  .filter((/** @type {any} */ e) => e.toFolder && e.fromFolder && e.toFolder !== e.fromFolder)
  .map((/** @type {any} */ e) => ({ file: e.to, from: e.fromFolder, into: e.toFolder }));
const foldersMade = undo
  .filter((/** @type {any} */ e) => e.kind === 'create folder')
  .map((/** @type {any} */ e) => String(e.to).replace(/^\/sites\/TaskManagement\/Shared Documents\//, ''));

const data = {
  capturedOn: listing.capturedOn,
  scanned: live.length,
  compliant: listing.documents.filter((/** @type {any} */ d) => d.compliant).length,
  folders: listing.folders.length,
  traced: history.length,
  reconstructed: history.filter((h) => h.reconstructed).length,
  history,
  moves,
  foldersMade,
  pending: pending.map((/** @type {any} */ l) => ({ from: l.from, to: String(l.to), why: l.why })),
  colliding: colliding.map((/** @type {any} */ l) => ({ from: l.from, to: String(l.to) })),
  applied: applied.length,
  reverse: history.map((h) => ({ from: h.current, to: h.original })),
  approvedFrom: plan.approvedFrom,
  untouched: record.untraced,
};

const html = readFileSync(htmlPath, 'utf8');
const OPEN = '<script type="application/json" id="data">';
const start = html.indexOf(OPEN);
if (start === -1) {
  console.error('docs/demo-runbook.html has no data block to fill.');
  process.exit(2);
}
const end = html.indexOf('</script>', start);
writeFileSync(
  htmlPath,
  html.slice(0, start + OPEN.length) + '\n' + JSON.stringify(data, null, 2) + '\n' + html.slice(end),
);
console.log(
  'library ' + data.scanned + ' documents, ' + data.compliant + ' compliant  ·  ' +
  data.traced + ' traced back (' + data.reconstructed + ' reconstructed)  ·  ' +
  data.pending.length + ' pending, ' + data.colliding.length + ' colliding  ->  docs/demo-runbook.html',
);
