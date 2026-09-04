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
 * Every node body a generator builds, read back out of the SDK source it emits.
 *
 * This started as a hand-written list per workflow, which covered WF-F and WF-G
 * and nothing else - so eight of the ten deployed workflows had no way to be
 * checked against the repo at all, and one of them (WF-B's sandbox) turned out
 * to be running August's logic while production had a September fix.
 *
 * Reading the built source instead means every generator is covered by
 * construction, including the ones whose node bodies are written inline at the
 * call rather than exported. The pairing is safe because a node's `name` always
 * precedes its `parameters` in the SDK config, so the nearest `name:` looking
 * backwards from a body is that body's own.
 *
 * @param {string} code @returns {Record<string, string>}
 */
export function nodeBodiesFrom(code) {
  /** @type {Record<string, string>} */
  const found = {};
  for (const key of ['jsCode', 'jsonOutput']) {
    let at = 0;
    for (;;) {
      const hit = code.indexOf(key + ': "', at);
      if (hit === -1) break;
      at = hit + 1;

      // Walk the JSON string literal to its close, honouring escapes.
      const opens = hit + key.length + 2;
      let end = -1;
      for (let i = opens + 1; i < code.length; i += 1) {
        if (code[i] === '\\') { i += 1; continue; }
        if (code[i] === '"') { end = i; break; }
      }
      if (end === -1) continue;
      let body;
      try { body = JSON.parse(code.slice(opens, end + 1)); } catch { continue; }

      // The node's own name, which sits a few lines above it in the config.
      const window = code.slice(Math.max(0, hit - 1200), hit);
      const names = [...window.matchAll(/name:\s*(['"])((?:\\.|(?!\1).)*)\1/g)];
      const last = names[names.length - 1];
      if (!last || !last[2]) continue;
      found[last[2]] = body;
    }
  }
  return found;
}

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

/**
 * Every generator, by the name its target uses in config/deploy.json.
 *
 * A generator that refuses to build for the current target - the seeder in
 * production, the SharePoint ones with no credential - is reported rather than
 * emitted, because a build that skips is working as intended and an empty file
 * is not an answer.
 */
/** @type {Record<string, () => Promise<{ build: () => Promise<string> }>>} */
const GENERATORS = {
  'wf-a': () => import('../workflows/wf-a-classify.mjs'),
  'wf-b': () => import('../workflows/wf-b-draft.mjs'),
  'wf-c': () => import('../workflows/wf-c-digest.mjs'),
  'wf-d': () => import('../workflows/wf-d-attachments.mjs'),
  'wf-e': () => import('../workflows/wf-e-errors.mjs'),
  'wf-f': () => import('../workflows/wf-f-organizer.mjs'),
  'wf-g': () => import('../workflows/wf-g-selfserve.mjs'),
  'wf-p': () => import('../workflows/wf-p-permissions.mjs'),
  'wf-s': () => import('../workflows/wf-s-seed.mjs'),
  'wf-t': () => import('../workflows/wf-t-preflight.mjs'),
};

/**
 * The organiser and the endpoint keep their hand-written lists as well, and the
 * two are compared. If reading the built source ever disagrees with naming the
 * constants directly, that is worth knowing before it is trusted for the eight
 * workflows that have no second opinion.
 */
async function organiserByHand() {
  return {
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
  };
}

async function endpointByHand() {
  return {
    Decide: await codeNode(HELPERS['wf-g/Decide'], DECIDE),
    Answer: await codeNode(HELPERS['wf-g/Answer'], ANSWER),
  };
}

/** @param {string} which */
async function emit(which) {
  const load = GENERATORS[which];
  if (!load) throw new Error('no generator called ' + which);
  const module = await load();
  const source = await module.build();
  const nodes = nodeBodiesFrom(source);

  // The two that can be checked both ways are checked both ways.
  const byHand = which === 'wf-f' ? await organiserByHand() : which === 'wf-g' ? await endpointByHand() : null;
  if (byHand) {
    for (const [name, body] of Object.entries(byHand)) {
      if (nodes[name] === body) continue;
      throw new Error(
        'reading the built source disagrees with the hand-written list for ' + name +
        ' (source=' + String(nodes[name] || '').length + ', constants=' + body.length + ')'
      );
    }
  }

  return {
    nodes,
    stickies: stickiesFrom(source),
  };
}

const which = String(process.argv[2] || '');
const path = process.argv[3];
if (!GENERATORS[which] || !path) {
  console.error('usage: node scripts/emit-nodes.mjs <' + Object.keys(GENERATORS).join('|') + '> <out.json>');
  process.exit(2);
}

const out = await emit(which);
writeFileSync(path, JSON.stringify(out, null, 2));
console.log(path);
for (const [name, body] of Object.entries(out.nodes)) {
  console.log('  ' + String(body.length).padStart(7) + '  ' + name);
}
for (const [heading, body] of Object.entries(out.stickies)) {
  console.log('  ' + String(body.length).padStart(7) + '  note: ' + heading);
}
