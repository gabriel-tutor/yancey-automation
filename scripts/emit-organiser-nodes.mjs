/**
 * The exact node bodies for WF-F, written out so a deployment updates the live
 * workflow with what the build produced rather than with anything retyped.
 *
 * Retyping is not a hypothetical risk here: a plan typed by hand into the
 * Config node silently lost three U+202F characters and half-applied.
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
  SAFETY_NOTE,
  STRUCTURE_NOTE,
  SUMMARY,
} from '../workflows/wf-f-organizer.mjs';

/**
 * Every Code node, not only the ones a given change happens to touch.
 *
 * It started as the four I was editing, which was fine while the only target
 * was a workflow that already had the rest. It is not fine for a fresh copy:
 * the sandbox was created with placeholder bodies, and a partial emitter would
 * have left half of it unable to run with no obvious sign of which half.
 *
 * `Confirm What Changed` and `Summary` carry no helpers, so they are the node
 * body verbatim; the rest are tree-shaken against workflows/helpers-manifest.mjs.
 */
const out = {
  // Two nodes now, not one. Defaults holds the settings the build produced;
  // Config is what everything downstream reads, and it is Defaults folded
  // together with whatever a caller was allowed to change.
  Defaults: JSON.stringify(CONFIG, null, 2),
  Config: await codeNode(HELPERS['wf-f/Config'], CONFIG_MERGE),
  'Pick The Reports': await codeNode(HELPERS['wf-f/Pick The Reports'], PICK_REPORTS),
  'Read Reports': await codeNode(HELPERS['wf-f/Read Reports'], READ_REPORTS),
  'Review The Library': await codeNode(HELPERS['wf-f/Review The Library'], REVIEW),
  'Pick What To Name': await codeNode(HELPERS['wf-f/Pick What To Name'], PICK),
  'Read Contents': await codeNode(HELPERS['wf-f/Read Contents'], READ_CONTENTS),
  'Compose Request': await codeNode(HELPERS['wf-f/Compose Request'], COMPOSE),
  'Apply Proposals': await codeNode(HELPERS['wf-f/Apply Proposals'], APPLY),
  'Confirm What Changed': CONFIRM,
  Summary: SUMMARY,
  notes: { structure: STRUCTURE_NOTE, safety: SAFETY_NOTE },
};

const path = process.argv[2] || 'dist/wf-f-nodes.json';
writeFileSync(path, JSON.stringify(out, null, 2));
console.log(path);
for (const [k, v] of Object.entries(out)) {
  if (typeof v === 'string') console.log('  ' + String(v.length).padStart(7) + '  ' + k);
}
