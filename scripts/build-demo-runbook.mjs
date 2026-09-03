/**
 * Fill the demo runbook from the plan that is actually pinned.
 *
 *   node scripts/build-demo-runbook.mjs
 *
 * Same rule as everywhere else here: no file name on a page is typed by hand.
 * A runbook that quotes a rename the workflow is not about to perform is worse
 * than no runbook, because it is read aloud in front of the client.
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
/**
 * A plan line approves a name, a destination, or both. The rename table and the
 * undo list are about names, so they are built from the lines that carry one; a
 * line with only `into` is a filing decision and gets its own short list.
 * @type {{ plan: Array<{ from: string, to?: string, into?: string, why: string }>, approvedFrom: string }}
 */
const plan = JSON.parse(readFileSync(join(root, 'config/approved-plan.json'), 'utf8'));
const renames = plan.plan.filter((/** @type {any} */ l) => l.to);
const filings = plan.plan.filter((/** @type {any} */ l) => !l.to && l.into);
const run = JSON.parse(readFileSync(join(root, 'fixtures/sharepoint/organiser-run.json'), 'utf8'));
const htmlPath = join(root, 'docs/demo-runbook.html');

/** @type {Array<{ current: string, destination: string, placementWhy: string, compliant: boolean, placement: string }>} */
const rows = run.rows || [];
const moves = rows.filter((r) => r.placement === 'move' && r.destination);

const data = {
  plan: renames.map((line) => ({ from: line.from, to: String(line.to), why: line.why })),
  reverse: renames.map((line) => ({ from: String(line.to), to: line.from })),
  filing: filings.map((line) => ({ from: line.from, into: String(line.into), why: line.why })),
  moves: moves.map((r) => ({
    file: r.current,
    into: String(r.destination).replace(/^\/sites\/TaskManagement\/Shared Documents\//, '').replace(/\/[^/]+$/, ''),
    because: r.placementWhy,
  })),
  scanned: rows.length,
  compliant: rows.filter((r) => r.compliant).length,
  reportsRead: (run.summary && run.summary.reportsRead) || 0,
  approvedFrom: plan.approvedFrom,
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
  `plan=${data.plan.length} filing=${data.filing.length} moves=${data.moves.length} scanned=${data.scanned} -> docs/demo-runbook.html`,
);
