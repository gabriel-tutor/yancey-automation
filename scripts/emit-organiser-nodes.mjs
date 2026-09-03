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
import { APPLY, COMPOSE, CONFIG, SAFETY_NOTE, STRUCTURE_NOTE, SUMMARY } from '../workflows/wf-f-organizer.mjs';

const out = {
  Config: JSON.stringify(CONFIG, null, 2),
  'Compose Request': await codeNode(HELPERS['wf-f/Compose Request'], COMPOSE),
  'Apply Proposals': await codeNode(HELPERS['wf-f/Apply Proposals'], APPLY),
  Summary: SUMMARY,
  notes: { structure: STRUCTURE_NOTE, safety: SAFETY_NOTE },
};

const path = process.argv[2] || 'dist/wf-f-nodes.json';
writeFileSync(path, JSON.stringify(out, null, 2));
console.log(path);
for (const [k, v] of Object.entries(out)) {
  if (typeof v === 'string') console.log('  ' + String(v.length).padStart(7) + '  ' + k);
}
