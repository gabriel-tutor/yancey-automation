import { describe, expect, it } from 'vitest';
import { existsSync, readFileSync } from 'node:fs';

/**
 * The deployment map, held to the rules that keep the n8n instance legible.
 *
 * Every rule here exists because breaking it already cost something. A dev
 * build once called the production WF-B and would have drafted into Yancey's
 * mailbox from the sandbox. The seeder — which writes fabricated mail — was
 * recorded under the production target and bound to a credential whose display
 * name lied about which mailbox it was. Neither was caught by anything running;
 * both were caught by a person reading carefully, which is not a control.
 */
const deploy = JSON.parse(readFileSync(new URL('../config/deploy.json', import.meta.url), 'utf8'));

/**
 * Every workflow unit, and the generator that builds it.
 * @type {Record<string, string | undefined>}
 */
const GENERATORS = {
  wfA: 'wf-a-classify.mjs',
  wfB: 'wf-b-draft.mjs',
  wfC: 'wf-c-digest.mjs',
  wfD: 'wf-d-attachments.mjs',
  wfE: 'wf-e-errors.mjs',
  wfF: 'wf-f-organizer.mjs',
  wfT: 'wf-t-preflight.mjs',
  wfS: 'wf-s-seed.mjs',
};

const targets = Object.keys(deploy.targets);

describe('the deployment map', () => {
  it('has exactly the two targets the build knows about', () => {
    expect(targets.sort()).toEqual(['dev', 'prod']);
  });

  // The generator composes every name as NAME + nameSuffix. Production keeps
  // the bare name; dev adds the marker. Sorting a list by name then puts a unit
  // beside its dev copy, which is how you notice one has drifted.
  it('names dev copies by suffix, so they sort beside the original', () => {
    expect(deploy.targets.prod.nameSuffix).toBe('');
    expect(deploy.targets.dev.nameSuffix).toBe(' · dev');
  });

  it.each(targets)('%s names only workflows that have a generator', (target) => {
    for (const unit of Object.keys(deploy.targets[target].workflows)) {
      expect(GENERATORS[unit], unit + ' has no generator').toBeDefined();
      const path = new URL('../workflows/' + GENERATORS[unit], import.meta.url);
      expect(existsSync(path), GENERATORS[unit] + ' is missing').toBe(true);
    }
  });

  it.each(targets)('%s gives every workflow a real id', (target) => {
    for (const [unit, id] of Object.entries(deploy.targets[target].workflows)) {
      expect(typeof id, target + '.' + unit).toBe('string');
      expect(String(id).length, target + '.' + unit + ' is empty').toBeGreaterThan(8);
    }
  });

  /**
   * The seeder plants six invented emails into whatever mailbox its credential
   * owns. Its id sat under the production target for five days. The generator
   * refuses to build it for production, so this is the second lock rather than
   * the first, and it is the one that reads at a glance.
   */
  it('keeps the seeder out of production entirely', () => {
    expect(deploy.targets.prod.workflows.wfS).toBeUndefined();
    expect(deploy.targets.dev.workflows.wfS).toBeTruthy();
  });

  /**
   * A dev workflow holding a production id is the failure that matters most
   * here: it does not look wrong, it does not error, and it writes to a client
   * mailbox. No id may appear under both targets.
   */
  it('never lets one id belong to two environments', () => {
    const seen = new Map();
    const clashes = [];
    for (const target of targets) {
      for (const [unit, id] of Object.entries(deploy.targets[target].workflows)) {
        const already = seen.get(id);
        if (already) clashes.push(id + ' is both ' + already + ' and ' + target + '.' + unit);
        else seen.set(id, target + '.' + unit);
      }
    }
    expect(clashes).toEqual([]);
  });

  it('never lets one mailbox credential serve two environments', () => {
    expect(deploy.targets.prod.outlookCredentialId).toBeTruthy();
    expect(deploy.targets.dev.outlookCredentialId).toBeTruthy();
    expect(deploy.targets.prod.outlookCredentialId).not.toBe(deploy.targets.dev.outlookCredentialId);
  });

  /**
   * There used to be a top-level workflows block holding production ids, and a
   * target that failed to name a workflow fell back to it silently. The block
   * is gone and the build now throws instead; this stops it coming back.
   */
  it('has no top-level workflow ids to fall back to', () => {
    expect(deploy.workflows).toBeUndefined();
  });

  it('writes the conventions down rather than leaving them to memory', () => {
    expect(deploy._conventions).toBeDefined();
    for (const key of ['naming', 'tags', 'archiving', 'everything_is_generated']) {
      expect(String(deploy._conventions[key] || '').length, key).toBeGreaterThan(40);
    }
  });
});

/**
 * The hand-built tools. These exist because a generated workflow cannot create
 * a test condition — WF-S seeds no attachments, so without a seeder that does,
 * WF-D cannot be exercised end to end in the sandbox at all.
 *
 * They were nearly archived on the argument that everything should be
 * generated. That argument was right about production and wrong about the
 * sandbox: reading execution data replaced the helpers that inspected a run,
 * but nothing replaces a seeder. What the repo must not do is forget they
 * exist, which is what this block is for.
 */
describe('the hand-built tools', () => {
  const entries = Object.entries(deploy.diagnostics || {}).filter(([key]) => !key.startsWith('_'));

  it('are listed at all, so no workflow in the instance is unaccounted for', () => {
    expect(entries.length).toBeGreaterThan(0);
  });

  it.each(entries.map(([key]) => key))('%s says what it is and why it is kept', (key) => {
    const tool = deploy.diagnostics[key];
    expect(String(tool.id || '').length, key + '.id').toBeGreaterThan(8);
    expect(String(tool.why || '').length, key + '.why').toBeGreaterThan(40);
  });

  // A WF- name is a promise that the repo generates it. These are not
  // generated, so they must not wear one.
  it.each(entries.map(([key]) => key))('%s is named Tool, not WF-', (key) => {
    expect(deploy.diagnostics[key].name).toMatch(/^Tool · /);
  });

  it('never reuses the id of a generated workflow', () => {
    const generated = new Set();
    for (const target of targets) {
      for (const id of Object.values(deploy.targets[target].workflows)) generated.add(id);
    }
    for (const [key, tool] of entries) {
      expect(generated.has(tool.id), key + ' shares an id with a generated workflow').toBe(false);
    }
  });
});
