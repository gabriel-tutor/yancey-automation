/**
 * The exact Code-node bodies a workflow should be running, written to a file.
 *
 * Deploying through the n8n MCP means reproducing whole node bodies by hand,
 * and retyping is not a hypothetical risk here: a plan typed into the Config
 * node once lost three U+202F characters and half-applied. This is the other
 * half of the answer — emit what the build says, then compare it against what
 * is deployed with scripts/verify-deployed.mjs.
 *
 *   node scripts/emit-nodes.mjs wf-f dist/wf-f-nodes.json
 *   TARGET=dev node scripts/emit-nodes.mjs wf-g dist/wf-g-dev-nodes.json
 *
 * Every Code node, not only the ones a given change happens to touch. It
 * started as the four being edited, which was fine while every target already
 * had the rest; it is not fine for a fresh copy, where a partial emitter leaves
 * half a workflow that cannot run and no sign of which half.
 *
 * Sticky notes are emitted too. They are documentation on the canvas rather
 * than behaviour, which is exactly why nothing looked at them and why the
 * sandbox's note went on saying the webhook was "unguessable but not
 * authenticated" for a day after it was authenticated. A note that lies is
 * worse than no note: it sends the next person to add protection that is
 * already there, or to treat a route as a secret.
 *
 * They are keyed by their heading — the first line — because n8n names a sticky
 * node with a generated hash that differs between the two deployments, so the
 * node name is no key at all. The heading is written in this repo and is stable.
 */
import { writeFileSync } from 'node:fs';
import { codeNode } from './bundle.mjs';
import { HELPERS } from '../workflows/helpers-manifest.mjs';
import {
  APPLY,
  COMPOSE,
  CONFIG,
  CONFIG_MERGE,
  CONFIRM,
  PICK,
  PICK_REPORTS,
  READ_CONTENTS,
  READ_REPORTS,
  REVIEW,
  SUMMARY,
} from '../workflows/wf-f-organizer.mjs';
import { ANSWER, DECIDE } from '../workflows/wf-g-selfserve.mjs';
import { build as buildOrganiser } from '../workflows/wf-f-organizer.mjs';
import { build as buildEndpoint } from '../workflows/wf-g-selfserve.mjs';

/**
 * The sticky contents a generator produces, read back out of the SDK source it
 * builds. Taking them from the built code rather than from exported constants
 * keeps every generator uniform — two of WF-F's four notes are written inline
 * at the call — and means what is emitted is literally what would be deployed.
 *
 * @param {string} code @returns {Record<string, string>}
 */
function stickiesFrom(code) {
  /** @type {Record<string, string>} */
  const found = {};
  let at = 0;
  for (;;) {
    const call = code.indexOf('sticky(', at);
    if (call === -1) return found;
    at = call + 1;
    const quote = code.indexOf('"', call);
    if (quote === -1) continue;
    // Walk the JSON string literal to its close, honouring escapes.
    let end = -1;
    for (let i = quote + 1; i < code.length; i += 1) {
      if (code[i] === '\\') { i += 1; continue; }
      if (code[i] === '"') { end = i; break; }
    }
    if (end === -1) continue;
    let content;
    try { content = JSON.parse(code.slice(quote, end + 1)); } catch { continue; }
    const heading = (String(content).split('\n')[0] || '').trim();
    if (!heading) continue;
    if (found[heading] !== undefined) {
      throw new Error('two sticky notes share the heading ' + JSON.stringify(heading) +
        ', so neither can be matched to a deployment; make the headings distinct');
    }
    found[heading] = content;
  }
}

/** The organiser. Ten Code nodes plus the Set node that holds the built config. */
async function organiser() {
  const nodes = {
    // Defaults is a Set node carrying JSON, not code. Config is what everything
    // downstream reads: Defaults folded together with a caller's overrides.
    Defaults: JSON.stringify(CONFIG, null, 2),
    Config: await codeNode(HELPERS['wf-f/Config'], CONFIG_MERGE),
    'Pick The Reports': await codeNode(HELPERS['wf-f/Pick The Reports'], PICK_REPORTS),
    'Read Reports': await codeNode(HELPERS['wf-f/Read Reports'], READ_REPORTS),
    'Review The Library': await codeNode(HELPERS['wf-f/Review The Library'], REVIEW),
    'Pick What To Name': await codeNode(HELPERS['wf-f/Pick What To Name'], PICK),
    'Read Contents': await codeNode(HELPERS['wf-f/Read Contents'], READ_CONTENTS),
    'Compose Request': await codeNode(HELPERS['wf-f/Compose Request'], COMPOSE),
    'Apply Proposals': await codeNode(HELPERS['wf-f/Apply Proposals'], APPLY),
    // These two carry no helpers, so they are the node body verbatim.
    'Confirm What Changed': CONFIRM,
    Summary: SUMMARY,
  };
  return { nodes, stickies: stickiesFrom(await buildOrganiser()) };
}

/** The web endpoint. Two Code nodes; the rest of it is a webhook and a reply. */
async function endpoint() {
  const nodes = {
    Decide: await codeNode(HELPERS['wf-g/Decide'], DECIDE),
    Answer: await codeNode(HELPERS['wf-g/Answer'], ANSWER),
  };
  return { nodes, stickies: stickiesFrom(await buildEndpoint()) };
}

/** @type {Record<string, () => Promise<{ nodes: Record<string, string>, stickies: Record<string, string> }>>} */
const EMITTERS = { 'wf-f': organiser, 'wf-g': endpoint };

const which = String(process.argv[2] || '');
const path = process.argv[3];
if (!EMITTERS[which] || !path) {
  console.error('usage: node scripts/emit-nodes.mjs <' + Object.keys(EMITTERS).join('|') + '> <out.json>');
  process.exit(2);
}

const emit = EMITTERS[which];
const out = await emit();
writeFileSync(path, JSON.stringify(out, null, 2));
console.log(path);
for (const [name, body] of Object.entries(out.nodes)) {
  console.log('  ' + String(body.length).padStart(7) + '  ' + name);
}
for (const [heading, body] of Object.entries(out.stickies)) {
  console.log('  ' + String(body.length).padStart(7) + '  note: ' + heading);
}
