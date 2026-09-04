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

/** The organiser. Ten Code nodes plus the Set node that holds the built config. */
async function organiser() {
  return {
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
}

/** The web endpoint. Two Code nodes; the rest of it is a webhook and a reply. */
async function endpoint() {
  return {
    Decide: await codeNode(HELPERS['wf-g/Decide'], DECIDE),
    Answer: await codeNode(HELPERS['wf-g/Answer'], ANSWER),
  };
}

/** @type {Record<string, () => Promise<Record<string, string>>>} */
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
for (const [name, body] of Object.entries(out)) {
  console.log('  ' + String(body.length).padStart(7) + '  ' + name);
}
