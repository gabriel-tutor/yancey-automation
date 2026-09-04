/**
 * The library as a run actually measured it, saved as a fixture.
 *
 *   node scripts/pull-library.mjs <transcript.jsonl> <out.json> [runId]
 *
 * docs/demo-runbook.html is built from this rather than from the approved plan,
 * because a plan records what somebody signed off and can say nothing about
 * whether it has happened yet. On 2026-09-04 the page still promised fifteen
 * renames that had been applied the day before.
 *
 * Taken from the transcript for the same reason scripts/pull-deployed.mjs is:
 * typing a listing in by hand would make every check downstream a check of the
 * typing. Read the Review The Library node's output with
 * get_workflow_execution, then run this.
 */
import { createReadStream, writeFileSync } from 'node:fs';
import { createInterface } from 'node:readline';

const [transcript, out, runId] = process.argv.slice(2);
if (!transcript || !out) {
  console.error('usage: node scripts/pull-library.mjs <transcript.jsonl> <out.json> [runId]');
  process.exit(2);
}

const MARKER = '"library":"Documents"';
/** @type {any} */
let found = null;

const lines = createInterface({ input: createReadStream(transcript), crlfDelay: Infinity });
for await (const line of lines) {
  if (!line.includes(MARKER)) continue;
  let at = 0;
  for (;;) {
    const hit = line.indexOf(MARKER, at);
    if (hit === -1) break;
    at = hit + 1;
    const start = line.lastIndexOf('{', hit);
    if (start === -1) continue;
    let depth = 0;
    let inString = false;
    let escaped = false;
    let end = -1;
    for (let i = start; i < line.length; i += 1) {
      const c = line[i];
      if (escaped) { escaped = false; continue; }
      if (c === '\\') { escaped = true; continue; }
      if (c === '"') { inString = !inString; continue; }
      if (inString) continue;
      if (c === '{') depth += 1;
      else if (c === '}') { depth -= 1; if (depth === 0) { end = i; break; } }
    }
    if (end === -1) continue;
    try {
      const parsed = JSON.parse(line.slice(start, end + 1));
      // The LAST one in the transcript, so a refresh picks up the newest read.
      if (parsed && Array.isArray(parsed.rows) && Array.isArray(parsed.files)) found = parsed;
    } catch { /* not this one */ }
  }
}

if (!found) {
  console.error('no library listing in the transcript — read Review The Library first');
  process.exit(1);
}

const today = new Date().toISOString().slice(0, 10);
writeFileSync(out, JSON.stringify({
  _note: 'Measured, not typed. Taken verbatim from a Review The Library output via scripts/pull-library.mjs, ' +
    'so the runbook is built from what the library actually holds rather than from a plan somebody believes is still pending.',
  capturedOn: today,
  run: runId || '',
  library: found.library,
  site: found.site,
  folders: found.folders,
  documents: found.rows.map((/** @type {any} */ r) => ({
    name: String(r.current),
    path: String(r.path),
    compliant: r.compliant === true,
  })),
}, null, 2) + '\n');

const docs = found.rows.length;
const ok = found.rows.filter((/** @type {any} */ r) => r.compliant === true).length;
console.log(out + '  ' + docs + ' documents, ' + ok + ' compliant, ' + (docs - ok) + ' not, ' + found.folders.length + ' folders');
