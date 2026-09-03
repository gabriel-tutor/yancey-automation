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
/** @type {{ plan: Array<{ from: string, to: string, why: string }>, approvedFrom: string }} */
const plan = JSON.parse(readFileSync(join(root, 'config/approved-plan.json'), 'utf8'));
const run = JSON.parse(readFileSync(join(root, 'fixtures/sharepoint/organiser-run.json'), 'utf8'));
const htmlPath = join(root, 'docs/demo-runbook.html');

/** @type {Array<{ current: string, destination: string, placementWhy: string, compliant: boolean, placement: string }>} */
const rows = run.rows || [];
const moves = rows.filter((r) => r.placement === 'move' && r.destination);

const data = {
  plan: plan.plan.map((line) => ({ from: line.from, to: line.to, why: line.why })),
  reverse: plan.plan.map((line) => ({ from: line.to, to: line.from })),
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
  `plan=${data.plan.length} moves=${data.moves.length} scanned=${data.scanned} -> docs/demo-runbook.html`,
);
