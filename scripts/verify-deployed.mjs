/**
 * Does the workflow that is running match the workflow this repo builds?
 *
 * The only honest answer to that is a byte comparison, and it has earned its
 * place repeatedly: it caught 4.5KB of dead prompt riding into the apply node,
 * a stray newline from splicing a node out of two pieces, and a line in
 * production's Read Reports that the repo has never produced — a hand-paste
 * that was never committed. None of those would have failed a behaviour test.
 *
 *   node scripts/emit-nodes.mjs wf-f dist/expected.json
 *   node scripts/verify-deployed.mjs <deployed.json> dist/expected.json
 *
 * `deployed.json` is what the n8n MCP's get_workflow_details writes when the
 * answer is too large to return inline, so no container access is needed.
 *
 * Exits non-zero when anything differs, so it can gate a deploy.
 *
 * Normalise before comparing, or every run is a false alarm: esbuild emits
 * \uXXXX escapes where a paste round-trips to the character itself. Same string
 * at runtime, different bytes. Decode on BOTH sides, then compare exactly.
 */
import { readFileSync } from 'node:fs';

/** @param {string} value */
const decodeEscapes = (value) =>
  String(value).replace(/\\u([0-9a-fA-F]{4})/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)));

const [deployedPath, expectedPath] = process.argv.slice(2);
if (!deployedPath || !expectedPath) {
  console.error('usage: node scripts/verify-deployed.mjs <deployed.json> <expected.json>');
  process.exit(2);
}

const saved = JSON.parse(readFileSync(deployedPath, 'utf8'));
const workflow = saved.workflow || saved;
/** @type {Record<string, string>} */
const expected = JSON.parse(readFileSync(expectedPath, 'utf8'));

/** Every node that carries code or raw JSON, keyed by the name on the canvas.
 * @type {Record<string, string>} */
const deployed = {};
for (const node of workflow.nodes || []) {
  const parameters = node.parameters || {};
  if (typeof parameters.jsCode === 'string') deployed[node.name] = parameters.jsCode;
  else if (typeof parameters.jsonOutput === 'string') deployed[node.name] = parameters.jsonOutput;
}

let wrong = 0;
for (const [name, want] of Object.entries(expected)) {
  if (!(name in deployed)) {
    console.log('MISSING  ' + name + '  — the deployed workflow has no such node');
    wrong += 1;
    continue;
  }
  const there = decodeEscapes(String(deployed[name]));
  const here = decodeEscapes(String(want));
  if (there === here) {
    console.log('MATCH    ' + name.padEnd(22) + String(here.length).padStart(7));
    continue;
  }

  wrong += 1;
  console.log('DIFFER   ' + name.padEnd(22) + 'deployed=' + there.length + ' built=' + here.length);
  // Say WHERE, because "these 45KB differ" is not a finding anybody can act on.
  const shortest = Math.min(there.length, here.length);
  let at = shortest;
  for (let i = 0; i < shortest; i += 1) {
    if (there[i] !== here[i]) {
      at = i;
      break;
    }
  }
  const from = Math.max(0, at - 60);
  if (at < shortest) {
    console.log('         first difference at character ' + at);
    console.log('         deployed: ' + JSON.stringify(there.slice(from, at + 60)));
    console.log('         built   : ' + JSON.stringify(here.slice(from, at + 60)));
  } else {
    const longer = there.length > here.length ? there : here;
    const side = there.length > here.length ? 'deployed' : 'built';
    console.log('         ' + side + ' carries ' + Math.abs(there.length - here.length) +
      ' extra characters from ' + at + ': ' + JSON.stringify(longer.slice(at, at + 80)));
  }
}

console.log('');
console.log(workflow.name + '  ·  ' + (workflow.nodeCount || (workflow.nodes || []).length) + ' nodes  ·  ' +
  (workflow.active ? 'published' : 'NOT published'));
console.log(wrong === 0
  ? 'every node matches the build'
  : wrong + ' node(s) differ from the build');
process.exit(wrong === 0 ? 0 : 1);
