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

/** @param {{ name: string, jsCode: string, settings?: Record<string, unknown> }} how */
function deployed(how) {
  const path = join(work, 'deployed-' + Math.random().toString(36).slice(2) + '.json');
  writeFileSync(path, JSON.stringify({
    workflow: {
      name: how.name,
      active: true,
      settings: how.settings === undefined ? { ...expectedSettings('prod') } : how.settings,
      nodes: [{ name: 'Decide', parameters: { jsCode: how.jsCode } }],
    },
  }));
  return path;
}

/** @param {string} target */
function expectedSettings(target) {
  const { _note, ...rest } = deploy.targets[target].workflowSettings;
  return rest;
}

/** @param {string} jsCode */
function built(jsCode) {
  const path = join(work, 'built-' + Math.random().toString(36).slice(2) + '.json');
  writeFileSync(path, JSON.stringify({ Decide: jsCode }));
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
