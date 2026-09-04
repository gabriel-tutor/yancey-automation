/**
 * A deployed workflow, taken out of the session transcript rather than retyped.
 *
 * scripts/verify-deployed.mjs needs the deployed workflow as a file. The n8n
 * MCP saves a large answer to one and returns a small one inline, so WF-F
 * arrives as a file and WF-G does not. Copying an inline answer into a file by
 * hand would make the comparison circular — it would prove the build matches
 * somebody's typing, which is exactly the failure the byte check exists to
 * catch. The transcript holds every tool result verbatim, so take it from there.
 *
 *   # whole workflow, from one get_workflow_details
 *   node scripts/pull-deployed.mjs <transcript.jsonl> <workflowId> <out.json>
 *
 *   # nodes from the PUBLISHED version, name and settings from the live read
 *   node scripts/pull-deployed.mjs <transcript.jsonl> <workflowId> <out.json> --version <versionId>
 *
 * The second form exists because get_workflow_version returns the version that
 * is actually running but carries no settings, and a workflow's draft can be
 * ahead of what is published. Both halves still come from the transcript; this
 * only joins them, and refuses if the live read is not at that same version.
 *
 * An n8n API key would make this unnecessary — the repo could fetch a workflow
 * directly. Until there is one, this is how a small workflow gets verified.
 */
import { createReadStream, writeFileSync } from 'node:fs';
import { createInterface } from 'node:readline';

/**
 * Every self-contained JSON object on `line` whose opening brace precedes an
 * occurrence of `marker`. Walks braces while honouring strings and escapes,
 * because a regex cannot find the end of a JSON object that contains code.
 *
 * @param {string} line
 * @param {string} marker
 * @returns {any[]}
 */
function objectsAt(line, marker) {
  const found = [];
  let at = 0;
  for (;;) {
    const hit = line.indexOf(marker, at);
    if (hit === -1) return found;
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
      else if (c === '}') {
        depth -= 1;
        if (depth === 0) { end = i; break; }
      }
    }
    if (end === -1) continue;
    try { found.push(JSON.parse(line.slice(start, end + 1))); } catch { /* not JSON on its own */ }
  }
}

const args = process.argv.slice(2);
const flag = args.indexOf('--version');
const wantVersion = flag === -1 ? '' : String(args[flag + 1] || '');
const [transcript, workflowId, out] = flag === -1 ? args : args.slice(0, flag);

if (!transcript || !workflowId || !out) {
  console.error('usage: node scripts/pull-deployed.mjs <transcript.jsonl> <workflowId> <out.json> [--version <versionId>]');
  process.exit(2);
}

/** @type {any} */ let live = null;
/** @type {any} */ let version = null;
const liveMarker = '"workflow":{"id":"' + workflowId + '"';

const lines = createInterface({ input: createReadStream(transcript), crlfDelay: Infinity });
for await (const line of lines) {
  if (line.includes(liveMarker)) {
    for (const o of objectsAt(line, liveMarker)) {
      if (o.workflow && o.workflow.id === workflowId) live = o;
    }
  }
  if (wantVersion && line.includes('"versionId":"' + wantVersion + '"')) {
    for (const o of objectsAt(line, '"success":true')) {
      if (o.workflowId === workflowId && o.versionId === wantVersion && Array.isArray(o.nodes)) version = o;
    }
  }
}

if (!live) {
  console.error('no get_workflow_details for ' + workflowId + ' in the transcript');
  process.exit(1);
}

let nodes = live.workflow.nodes;
if (wantVersion) {
  if (!version) {
    console.error('no get_workflow_version for ' + wantVersion + ' in the transcript');
    process.exit(1);
  }
  // Settings from a live read of a DIFFERENT version would describe a workflow
  // nobody is running. Refuse rather than quietly mix the two.
  if (live.workflow.activeVersionId !== wantVersion) {
    console.error('the live read is at ' + live.workflow.activeVersionId + ', not ' + wantVersion +
      '; re-read the workflow after publishing');
    process.exit(1);
  }
  nodes = version.nodes;
}

if (!Array.isArray(nodes)) {
  console.error('that read carries no nodes — use detailLevel "full", or pass --version');
  process.exit(1);
}

writeFileSync(out, JSON.stringify({ workflow: { ...live.workflow, nodes } }, null, 2));
console.log(out);
console.log('  ' + live.workflow.name);
console.log('  nodes    from ' + (wantVersion ? 'version ' + wantVersion : 'the workflow read') + ' (' + nodes.length + ')');
console.log('  settings from the live read, activeVersionId ' + live.workflow.activeVersionId);
