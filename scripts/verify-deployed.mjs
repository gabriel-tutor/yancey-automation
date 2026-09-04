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
 *
 * Sticky notes are checked too, keyed by their heading — n8n names a sticky node
 * with a generated hash that differs between deployments, so the node name is no
 * key. They are documentation rather than behaviour, which is why nothing looked
 * at them, and why the sandbox's note went on saying the webhook was
 * "unguessable but not authenticated" for a day after it was authenticated.
 *
 * Node bodies are not the whole deployment. Where a failure is routed and how
 * long a run may take are workflow SETTINGS, and the SDK builder cannot express
 * them — `workflow(id, name)` takes an id and a name. So no generator produces
 * them, the n8n editor is the only place they exist, and nothing would notice
 * them being lost. config/deploy.json records what they must be per target and
 * this checks the deployment against that, under the same TARGET convention as
 * the build: `TARGET=dev npm run verify -- ...`.
 */
import { readFileSync } from 'node:fs';
import { TARGET } from '../workflows/shared.mjs';

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
const emitted = JSON.parse(readFileSync(expectedPath, 'utf8'));
/** The emitter grew a shape; a flat map is still the older one. @type {Record<string, string>} */
const expected = emitted.nodes || emitted;
/** @type {Record<string, string>} */
const expectedNotes = emitted.stickies || {};

/** Every node that carries code or raw JSON, keyed by the name on the canvas.
 * @type {Record<string, string>} */
const deployed = {};
for (const node of workflow.nodes || []) {
  const parameters = node.parameters || {};
  if (typeof parameters.jsCode === 'string') deployed[node.name] = parameters.jsCode;
  else if (typeof parameters.jsonOutput === 'string') deployed[node.name] = parameters.jsonOutput;
}

/** Deployed sticky notes, keyed by heading the way the build keys them. */
/** @type {Record<string, string>} */
const deployedNotes = {};
for (const node of workflow.nodes || []) {
  if (node.type !== 'n8n-nodes-base.stickyNote') continue;
  const content = String((node.parameters || {}).content || '');
  const heading = (content.split('\n')[0] || '').trim();
  if (heading) deployedNotes[heading] = content;
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

for (const [heading, want] of Object.entries(expectedNotes)) {
  const there = deployedNotes[heading];
  if (there === undefined) {
    // A heading the deployment does not carry means the note was edited on the
    // canvas, which is where a note starts drifting from what it describes.
    console.log('MISSING  note ' + JSON.stringify(heading));
    wrong += 1;
    continue;
  }
  if (decodeEscapes(there) === decodeEscapes(String(want))) {
    console.log('MATCH    note ' + heading);
    continue;
  }
  wrong += 1;
  console.log('DIFFER   note ' + heading + '  deployed=' + there.length + ' built=' + String(want).length);
}

// A note the build does not produce is the same drift from the other side: the
// convention here is that every workflow is generated, with no exceptions.
for (const heading of Object.keys(deployedNotes)) {
  if (Object.keys(expectedNotes).length === 0) break;
  if (expectedNotes[heading] !== undefined) continue;
  console.log('EXTRA    note ' + JSON.stringify(heading) + ' — on the canvas, not in the build');
  wrong += 1;
}

// Verifying a dev deployment with TARGET unset would hold it to production's
// error handler and report a difference that is really a mistake in the
// command. The name carries the suffix, so say so rather than failing oddly.
const suffix = TARGET.nameSuffix || '';
const named = String(workflow.name || '');
if (suffix ? !named.endsWith(suffix) : named.includes(' \u00b7 dev')) {
  console.log('');
  console.log('WRONG TARGET  ' + JSON.stringify(named) + ' is not a ' + TARGET.name + ' workflow');
  // The fix is always the OTHER target, so read it off the file, not off TARGET.
  console.log('              re-run with ' + (suffix ? 'TARGET unset' : 'TARGET=dev'));
  process.exit(2);
}

// Settings the generators cannot build, so the repo asserts them instead.
const rules = /** @type {any} */ (TARGET.workflowSettings || {});
const haveSettings = workflow.settings || {};

// Every workflow is named WF-X, and that letter is the unit — which is what the
// timeout is keyed on, because the right ceiling is not one number for a poller
// that runs every minute and an organiser that reads a whole library.
const letter = /^WF-([A-Z])\b/.exec(String(workflow.name || ''));
const unit = letter ? 'wf' + letter[1] : '';
if (!unit) {
  console.log('');
  console.log('UNNAMED  ' + JSON.stringify(String(workflow.name || '')) +
    ' does not start with WF-<letter>, so no unit rule can apply to it');
  process.exit(2);
}

/** @type {Record<string, unknown>} */
const wantSettings = {};
// WF-E is the handler. Pointing it at itself is a loop, so it is the one
// workflow with no error workflow — one exception, so there is one to remember.
if (rules.errorWorkflow && unit !== rules.errorWorkflowExcept) {
  wantSettings.errorWorkflow = rules.errorWorkflow;
}
if (rules.saveDataErrorExecution) wantSettings.saveDataErrorExecution = rules.saveDataErrorExecution;
const ceiling = (rules.executionTimeout || {})[unit];
if (ceiling) wantSettings.executionTimeout = ceiling;

for (const [key, want] of Object.entries(wantSettings)) {
  const have = haveSettings[key];
  if (have === want) {
    console.log('MATCH    setting ' + key.padEnd(30) + JSON.stringify(want));
    continue;
  }
  wrong += 1;
  console.log('DIFFER   setting ' + key.padEnd(30) +
    'deployed=' + JSON.stringify(have === undefined ? null : have) + ' required=' + JSON.stringify(want));
}
if (unit === rules.errorWorkflowExcept && haveSettings.errorWorkflow) {
  wrong += 1;
  console.log('DIFFER   setting errorWorkflow                 ' + unit +
    ' is the error handler and must route nowhere, but points at ' + JSON.stringify(haveSettings.errorWorkflow));
}

console.log('');
console.log(workflow.name + '  ·  ' + (workflow.nodeCount || (workflow.nodes || []).length) + ' nodes  ·  ' +
  (workflow.active ? 'published' : 'NOT published') + '  ·  checked against ' + TARGET.name);
console.log(wrong === 0
  ? 'every node and setting matches what the repo requires'
  : wrong + ' difference(s) from what the repo requires');
process.exit(wrong === 0 ? 0 : 1);
