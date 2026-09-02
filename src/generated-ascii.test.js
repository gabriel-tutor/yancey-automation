/**
 * Every Code node body this repo generates, checked for characters that do not
 * survive being moved as text.
 *
 * `src/bundle.test.js` already checks the tree-shaken helpers, and that check
 * immediately found a literal non-breaking space inside a normaliser and the
 * euro and sterling signs inside the guard that decides whether an invented fee
 * reaches a draft. But it only covers the shared helpers. The rest of a node —
 * the part written directly in the generator — went unchecked, and that is
 * where the next one was: four em dashes in the seeder's fixture subjects.
 *
 * They were already written as — in the source, which looks safe and is
 * not: inside a template literal JavaScript decodes the escape at evaluation
 * time, so a literal em dash landed in the generated node anyway. Writing
 * \\u2014 is what actually emits an escape. Nothing but a check on the output
 * would have caught that, because the input looked correct.
 *
 * Sticky notes are deliberately not covered: they are prose for a person
 * reading the canvas, and a mangled dash there is cosmetic rather than a
 * silently broken regex.
 */
import { describe, expect, it, beforeAll, vi } from 'vitest';

/**
 * Two generators refuse to build for the wrong environment, and they refuse in
 * opposite directions: the seeder will not build for production because it
 * plants fabricated mail, and the organiser will not build for dev because dev
 * has no SharePoint credential. So each is built under the target that owns it
 * rather than under one target with the guards worked around — a check that has
 * to disarm a safety rule to run is checking the wrong thing.
 *
 * shared.mjs reads the target once at import, so the registry is reset between
 * the two passes and the module is re-evaluated.
 */
const UNDER_PROD = [
  'wf-a-classify.mjs',
  'wf-b-draft.mjs',
  'wf-c-digest.mjs',
  'wf-d-attachments.mjs',
  'wf-e-errors.mjs',
  'wf-f-organizer.mjs',
  'wf-t-preflight.mjs',
];
const UNDER_DEV = ['wf-s-seed.mjs'];
const GENERATORS = [...UNDER_PROD, ...UNDER_DEV].sort();

/** @type {Record<string, string>} */
const built = {};

beforeAll(async () => {
  /** @type {Array<[string, string[]]>} */
  const passes = [
    ['prod', UNDER_PROD],
    ['dev', UNDER_DEV],
  ];
  for (const [target, files] of passes) {
    vi.resetModules();
    process.env.TARGET = target;
    for (const file of files) {
      const mod = await import('../workflows/' + file);
      built[file] = await mod.build();
    }
  }
}, 60000);

/**
 * Every `jsCode: "…"` in generated SDK source, decoded back to the JavaScript
 * that will actually run inside the node.
 *
 * @param {string} source
 * @returns {string[]}
 */
function nodeBodies(source) {
  const bodies = [];
  let at = source.indexOf('jsCode: "');
  while (at !== -1) {
    let k = at + 9;
    let escaped = '';
    while (k < source.length) {
      const c = source[k];
      if (c === '\\') {
        escaped += (source[k] ?? '') + (source[k + 1] ?? '');
        k += 2;
        continue;
      }
      if (c === '"') break;
      escaped += c;
      k += 1;
    }
    bodies.push(JSON.parse('"' + escaped + '"'));
    at = source.indexOf('jsCode: "', k);
  }
  return bodies;
}

describe('generated Code node bodies travel as text safely', () => {
  it.each(GENERATORS)('%s emits only ASCII', (file) => {
    const bodies = nodeBodies(built[file] ?? '');
    expect(bodies.length, file + ' produced no Code nodes to check').toBeGreaterThan(0);

    const offenders = [];
    for (const body of bodies) {
      for (const m of body.matchAll(/[^\x00-\x7F]/g)) {
        offenders.push({
          codePoint: 'U+' + (m[0]?.codePointAt(0) ?? 0).toString(16).toUpperCase().padStart(4, '0'),
          near: body.slice(Math.max(0, (m.index ?? 0) - 40), (m.index ?? 0) + 40),
        });
      }
    }
    expect(offenders).toEqual([]);
  });

  // A generator that silently produced nothing would pass every check above.
  it('builds a body for every generator', () => {
    for (const file of GENERATORS) {
      expect(String(built[file] || '').length, file).toBeGreaterThan(500);
    }
  });
});
