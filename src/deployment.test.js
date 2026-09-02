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
  wfP: 'wf-p-permissions.mjs',
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

  /**
   * Three workflows were once hand-built, and the honest way to keep them was
   * to write their generators rather than to soften the rule they broke. This
   * fails if a diagnostics escape hatch is ever reintroduced.
   */
  it('keeps no hand-built exceptions', () => {
    expect(deploy.diagnostics).toBeUndefined();
  });

  it('names the SharePoint probe under the target that owns the library', () => {
    expect(deploy.targets.prod.workflows.wfP).toBeTruthy();
    expect(deploy.targets.dev.workflows.wfP).toBeUndefined();
  });

  it('writes the conventions down rather than leaving them to memory', () => {
    expect(deploy._conventions).toBeDefined();
    for (const key of ['naming', 'tags', 'archiving', 'everything_is_generated', 'guards']) {
      expect(String(deploy._conventions[key] || '').length, key).toBeGreaterThan(40);
    }
  });
});
