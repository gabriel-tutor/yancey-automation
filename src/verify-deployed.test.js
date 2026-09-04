import { describe, expect, it } from 'vitest';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

/**
 * The check that gates a deploy.
 *
 * It answers one question — is the workflow that is running the workflow this
 * repo builds — and it has earned its place three times: 4.5KB of dead prompt
 * riding into the apply node, a stray newline from splicing a node out of two
 * pieces, and a line in production's Read Reports the repo had never produced.
 * None of those would have failed a behaviour test.
 *
 * It now also checks workflow SETTINGS, which no generator can build: the SDK's
 * `workflow(id, name)` takes an id and a name, so where a failure is routed
 * exists only in the n8n editor. That is the same shape of drift, and until
 * this it was the shape nothing looked at.
 */
const ROOT = new URL('..', import.meta.url).pathname;
const deploy = JSON.parse(readFileSync(join(ROOT, 'config/deploy.json'), 'utf8'));
const work = mkdtempSync(join(tmpdir(), 'verify-'));

/** @param {{ name: string, jsCode: string, settings?: Record<string, unknown>, notes?: string[] }} how */
function deployed(how) {
  const path = join(work, 'deployed-' + Math.random().toString(36).slice(2) + '.json');
  const nodes = /** @type {any[]} */ ([{ name: 'Decide', parameters: { jsCode: how.jsCode } }]);
  for (const content of how.notes || []) {
    // n8n names a sticky with a generated hash, which is exactly why the check
    // keys on the heading instead. Use a name nothing could match on.
    nodes.push({
      name: 'Sticky Note ' + Math.random().toString(16).slice(2, 10),
      type: 'n8n-nodes-base.stickyNote',
      parameters: { content },
    });
  }
  writeFileSync(path, JSON.stringify({
    workflow: {
      name: how.name,
      active: true,
      settings: how.settings === undefined ? { ...expectedSettings('prod') } : how.settings,
      nodes,
    },
  }));
  return path;
}

/**
 * What a workflow of this unit must carry, per the recorded rules — not a flat
 * copy of them: the time limit is keyed by unit, and WF-E carries no error
 * workflow because pointing the handler at itself is a loop.
 *
 * @param {string} target @param {string} [unit]
 */
function expectedSettings(target, unit) {
  const of = unit || 'wfF';
  const rules = deploy.targets[target].workflowSettings;
  /** @type {Record<string, unknown>} */
  const want = { saveDataErrorExecution: rules.saveDataErrorExecution };
  if (of !== rules.errorWorkflowExcept) want.errorWorkflow = rules.errorWorkflow;
  const ceiling = rules.executionTimeout[of];
  if (ceiling) want.executionTimeout = ceiling;
  return want;
}

/** The older, flat emitter shape — still read, so old files do not become lies.
 * @param {string} jsCode */
function built(jsCode) {
  const path = join(work, 'built-' + Math.random().toString(36).slice(2) + '.json');
  writeFileSync(path, JSON.stringify({ Decide: jsCode }));
  return path;
}

/** What the emitter writes now: node bodies and sticky notes keyed by heading.
 * @param {string} jsCode @param {string[]} notes */
function builtWith(jsCode, notes) {
  const path = join(work, 'built-' + Math.random().toString(36).slice(2) + '.json');
  /** @type {Record<string, string>} */
  const stickies = {};
  for (const content of notes) stickies[(content.split('\n')[0] || '').trim()] = content;
  writeFileSync(path, JSON.stringify({ nodes: { Decide: jsCode }, stickies }));
  return path;
}

/**
 * Run it the way a person would, because the exit code is half of what it is
 * for — a check that reports a difference and exits 0 gates nothing.
 *
 * @param {string} deployedPath @param {string} builtPath @param {string} [target]
 */
function run(deployedPath, builtPath, target) {
  const env = { ...process.env };
  if (target) env.TARGET = target;
  else delete env.TARGET;
  try {
    const out = execFileSync('node', ['scripts/verify-deployed.mjs', deployedPath, builtPath], {
      cwd: ROOT, env, encoding: 'utf8',
    });
    return { code: 0, out };
  } catch (error) {
    const e = /** @type {any} */ (error);
    return { code: e.status, out: String(e.stdout || '') + String(e.stderr || '') };
  }
}

const PROD = 'WF-F · Organise Documents';
const DEV = 'WF-F · Organise Documents · dev';

describe('verifying a deployment against the build', () => {
  it('passes when every node body and every required setting agrees', () => {
    const result = run(deployed({ name: PROD, jsCode: 'return [];' }), built('return [];'));
    expect(result.code).toBe(0);
    expect(result.out).toContain('every node and setting matches');
  });

  it('fails on a node body that differs, and says where', () => {
    const result = run(
      deployed({ name: PROD, jsCode: 'return [{ json: { ok: true } }];' }),
      built('return [{ json: { ok: false } }];')
    );
    expect(result.code).toBe(1);
    expect(result.out).toContain('DIFFER   Decide');
    expect(result.out, 'a character offset is what makes it actionable').toContain('first difference at character');
  });

  /**
   * The line in production's Read Reports was a hand-paste that no behaviour
   * test could see. A node the deployment does not have at all is the same
   * failure from the other direction.
   */
  it('fails when the deployment is missing a node the build produces', () => {
    const path = join(work, 'no-nodes.json');
    writeFileSync(path, JSON.stringify({ workflow: { name: PROD, settings: expectedSettings('prod'), nodes: [] } }));
    const result = run(path, built('return [];'));
    expect(result.code).toBe(1);
    expect(result.out).toContain('MISSING  Decide');
  });

  /**
   * The whole reason settings are checked here: a redeploy through the editor
   * can drop them and every node still matches.
   */
  it('fails when a required setting is absent from the deployment', () => {
    const result = run(
      deployed({ name: PROD, jsCode: 'return [];', settings: { executionOrder: 'v1' } }),
      built('return [];')
    );
    expect(result.code).toBe(1);
    expect(result.out).toContain('DIFFER   setting errorWorkflow');
    expect(result.out).toContain('deployed=null');
  });

  it('fails when a setting is present but points somewhere else', () => {
    const result = run(
      deployed({
        name: PROD,
        jsCode: 'return [];',
        settings: { ...expectedSettings('prod'), errorWorkflow: 'someOtherWorkflow' },
      }),
      built('return [];')
    );
    expect(result.code).toBe(1);
    expect(result.out).toContain('DIFFER   setting errorWorkflow');
  });

  it('holds the sandbox to the sandbox error handler, not production\'s', () => {
    const result = run(
      deployed({ name: DEV, jsCode: 'return [];', settings: expectedSettings('dev') }),
      built('return [];'),
      'dev'
    );
    expect(result.code).toBe(0);
    expect(result.out).toContain(deploy.targets.dev.workflows.wfE);
    expect(result.out, 'the sandbox must not route failures to the client log')
      .not.toContain(deploy.targets.prod.workflows.wfE);
  });

  /**
   * Checking a dev deployment against production's settings would report a
   * difference that is really a mistake in the command. Say so instead, and
   * name the fix — the advice was inverted the first time, which is worse than
   * no advice.
   */
  it('refuses a dev deployment checked as production, and names the fix', () => {
    const result = run(deployed({ name: DEV, jsCode: 'return [];' }), built('return [];'));
    expect(result.code).toBe(2);
    expect(result.out).toContain('WRONG TARGET');
    expect(result.out).toContain('re-run with TARGET=dev');
  });

  it('refuses a production deployment checked as dev, and names the fix', () => {
    const result = run(deployed({ name: PROD, jsCode: 'return [];' }), built('return [];'), 'dev');
    expect(result.code).toBe(2);
    expect(result.out).toContain('WRONG TARGET');
    expect(result.out).toContain('re-run with TARGET unset');
  });

  /**
   * esbuild emits \uXXXX escapes where a paste round-trips to the character
   * itself: the same string at runtime, different bytes. Every run would be a
   * false alarm without this, and a check that cries wolf gets switched off.
   */
  it('treats an escaped character and the character itself as the same', () => {
    const result = run(
      deployed({ name: PROD, jsCode: 'const dot = "·";' }),
      built('const dot = "\\u00b7";')
    );
    expect(result.code).toBe(0);
  });
});

/**
 * Sticky notes.
 *
 * They are documentation rather than behaviour, which is why nothing looked at
 * them for as long as nothing did — and why the sandbox's note went on saying
 * the webhook was "unguessable but not authenticated" for a day after it was
 * authenticated. A note that lies is worse than no note: it sends the next
 * person to add protection that is already there.
 *
 * They are keyed by heading because n8n names a sticky node with a generated
 * hash that differs between the two deployments, so the node name is no key.
 */
describe('the notes on the canvas', () => {
  const A = '## A note on the canvas\n\nWhat this workflow is for.';
  const B = '### Another note\n\nWhy it is safe.';

  it('passes when every note the build produces is on the canvas, word for word', () => {
    const result = run(
      deployed({ name: PROD, jsCode: 'return [];', notes: [A, B] }),
      builtWith('return [];', [A, B])
    );
    expect(result.code).toBe(0);
    expect(result.out).toContain('MATCH    note ## A note on the canvas');
  });

  it('fails when a note has been edited on the canvas', () => {
    const edited = A.replace('What this workflow is for.', 'Something somebody typed here instead.');
    const result = run(
      deployed({ name: PROD, jsCode: 'return [];', notes: [edited, B] }),
      builtWith('return [];', [A, B])
    );
    expect(result.code).toBe(1);
    expect(result.out).toContain('DIFFER   note ## A note on the canvas');
  });

  it('fails when the build produces a note the canvas never received', () => {
    const result = run(
      deployed({ name: PROD, jsCode: 'return [];', notes: [A] }),
      builtWith('return [];', [A, B])
    );
    expect(result.code).toBe(1);
    expect(result.out).toContain('MISSING  note "### Another note"');
  });

  /**
   * The convention here is that every workflow is generated, with no
   * exceptions. A note somebody added on the canvas is the deployment carrying
   * something the repo does not produce, which is the same drift as a
   * hand-pasted line in a Code node.
   */
  it('fails on a note that is on the canvas but not in the build', () => {
    const stray = '## Someone typed this straight into n8n\n\nAnd nobody rebuilt it.';
    const result = run(
      deployed({ name: PROD, jsCode: 'return [];', notes: [A, B, stray] }),
      builtWith('return [];', [A, B])
    );
    expect(result.code).toBe(1);
    expect(result.out).toContain('EXTRA    note "## Someone typed this straight into n8n"');
  });

  /**
   * A workflow whose notes are checked and one whose notes are not must not be
   * the same call. An emitted file from before notes were emitted carries none,
   * and every note on that canvas would otherwise report as EXTRA.
   */
  it('says nothing about notes when the build file carries none', () => {
    const result = run(
      deployed({ name: PROD, jsCode: 'return [];', notes: [A, B] }),
      built('return [];')
    );
    expect(result.code).toBe(0);
    expect(result.out).not.toContain('note');
  });
});

/**
 * The rules are per unit, and the one that is not universal is the error
 * handler: WF-E must route nowhere, because routing it at itself is a loop.
 */
describe('the rules that are not the same for every workflow', () => {
  it('holds the organiser to a longer run than the poller, as recorded', () => {
    const limits = deploy.targets.prod.workflowSettings.executionTimeout;
    expect(limits.wfF).toBeGreaterThan(limits.wfA);

    const result = run(
      deployed({ name: PROD, jsCode: 'return [];', settings: { ...expectedSettings('prod', 'wfF'), executionTimeout: limits.wfA } }),
      built('return [];')
    );
    expect(result.code, 'the poller\'s ceiling on the organiser is not the recorded one').toBe(1);
    expect(result.out).toContain('DIFFER   setting executionTimeout');
  });

  it('lets the error handler route nowhere', () => {
    const result = run(
      deployed({ name: 'WF-E · Error Handler', jsCode: 'return [];', settings: expectedSettings('prod', 'wfE') }),
      built('return [];')
    );
    expect(result.code).toBe(0);
    expect(result.out, 'WF-E must not be required to have one').not.toContain('DIFFER   setting errorWorkflow');
  });

  it('fails the error handler if somebody points it at itself', () => {
    const result = run(
      deployed({
        name: 'WF-E · Error Handler',
        jsCode: 'return [];',
        settings: { ...expectedSettings('prod', 'wfE'), errorWorkflow: deploy.targets.prod.workflows.wfE },
      }),
      built('return [];')
    );
    expect(result.code).toBe(1);
    expect(result.out).toContain('must route nowhere');
  });

  /**
   * Every workflow here is named WF-X and that letter is the unit, so a name
   * that does not follow the convention has no rule that applies to it. Saying
   * so beats silently checking nothing.
   */
  it('refuses a workflow whose name carries no unit', () => {
    const result = run(
      deployed({ name: 'Some workflow somebody made in the editor', jsCode: 'return [];' }),
      built('return [];')
    );
    expect(result.code).toBe(2);
    expect(result.out).toContain('UNNAMED');
  });
});
