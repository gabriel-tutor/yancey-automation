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
  wfG: 'wf-g-selfserve.mjs',
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

  /**
   * Settings the generators cannot produce.
   *
   * `workflow(id, name)` is the whole SDK surface for a workflow, so where a
   * failure is routed and how long a run may take exist only in the n8n editor.
   * The repo cannot build them; what it can do is say what they must be, and
   * fail here when they say something else. Without this the rule lives in one
   * person's memory, which is how production kept a line in Read Reports that
   * the repo had never produced.
   */
  it.each(targets)('%s says what its deployed workflows must be set to', (target) => {
    const settings = deploy.targets[target].workflowSettings;
    expect(settings, target + ' records no workflow settings').toBeDefined();
    expect(settings.executionTimeout).toBeGreaterThan(0);
    expect(settings.saveDataErrorExecution).toBe('all');
  });

  /**
   * A sandbox that routes its failures to the production error handler writes
   * dev noise into the client's error log and alerts a client about it. Every
   * target must handle its own failures, and the only id that can be right is
   * that target's own WF-E.
   */
  it.each(targets)('%s routes its failures to its own error handler', (target) => {
    const settings = deploy.targets[target].workflowSettings;
    expect(settings.errorWorkflow, target + ' names no error workflow').toBeTruthy();
    expect(settings.errorWorkflow).toBe(deploy.targets[target].workflows.wfE);
  });

  it('writes the conventions down rather than leaving them to memory', () => {
    expect(deploy._conventions).toBeDefined();
    for (const key of ['naming', 'tags', 'archiving', 'everything_is_generated', 'guards']) {
      expect(String(deploy._conventions[key] || '').length, key).toBeGreaterThan(40);
    }
  });
});

/**
 * The organiser sandbox.
 *
 * It exists so that deploying half-finished node code cannot disturb the
 * workflow Yancey is running — n8n has no published version of it
 * (`activeVersionId` is null), so a manual Execute runs whatever was last
 * saved, and there is no other separation.
 *
 * It reads the SAME library through the SAME credential, because there is only
 * one and the real edge cases are in the real library. That makes it isolated
 * for deployment and not for writes, so the one thing it must never ship is an
 * armed switch: a dev build that can write to production is a second production
 * workflow nobody is watching.
 */
describe('the organiser sandbox', () => {
  const dev = deploy.targets.dev;

  it('can reach the library at all', () => {
    expect(dev.sharepointCredential, 'without this the organiser refuses to build').toBeTruthy();
    expect(dev.sharepointHost).toBeTruthy();
    expect(dev.sharepointSite).toBeTruthy();
  });

  it('is a different workflow from the one Yancey runs', () => {
    expect(dev.workflows.wfF, 'the sandbox has no workflow to deploy into').toBeTruthy();
    expect(dev.workflows.wfF).not.toBe(deploy.targets.prod.workflows.wfF);
  });

  it('says in the file what it is and is not isolated from', () => {
    expect(dev._wfFNote).toContain('SAME library');
    expect(dev._wfFNote).toContain('read-only');
  });

  /**
   * A build failure rather than a silent correction: quietly changing what
   * somebody asked for is how a setting stops meaning anything.
   */
  it('refuses to build with any switch that could write armed', async () => {
    const { CONFIG } = await import('../workflows/wf-f-organizer.mjs');
    for (const key of ['applyRenames', 'applyMoves', 'applyStructure', 'applyFolderRenames', 'applyProposed']) {
      const armed = /** @type {any} */ (CONFIG)[key];
      expect(armed, key + ' ships armed, which the sandbox build refuses').toBe(false);
    }
  });
});
