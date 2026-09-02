import { describe, expect, it } from 'vitest';
import { subset } from '../scripts/bundle.mjs';
import { HELPERS } from '../workflows/helpers-manifest.mjs';
import * as lib from './agent-lib.js';
import { DEEP_FORWARD_CHAIN, OUTLOOK_HTML_BODY, graphMessage } from './fixtures.js';

/**
 * The tree-shaken subsets are what actually run inside n8n. These assertions
 * prove each one still behaves like the module the unit tests cover —
 * otherwise the suite would be grading code that never ships.
 */
/**
 * @param {string} source
 * @param {string} expression
 * @returns {any}
 */
const evalIn = (source, expression) =>
  new Function(source + '\nreturn (' + expression + ');')();

/** @type {Record<string, string>} */
const subsets = {};
/** @type {Record<string, { lib?: string[], prompts?: string[], unzip?: string[], attachments?: string[], filenames?: string[] }>} */
const manifest = HELPERS;
for (const [nodeKey, needs] of Object.entries(manifest)) {
  subsets[nodeKey] = await subset(needs);
}

/**
 * @param {string} nodeKey
 * @returns {string}
 */
const helpersFor = (nodeKey) => {
  const source = subsets[nodeKey];
  if (source === undefined) throw new Error('no subset built for ' + nodeKey);
  return source;
};

describe('tree-shaken Code node helpers', () => {
  it.each(Object.keys(HELPERS))('%s emits runnable, module-free code', (nodeKey) => {
    const source = helpersFor(nodeKey);
    const needs = manifest[nodeKey] ?? {};
    expect(source.length).toBeGreaterThan(0);
    expect(/^\s*(import|export)\s/m.test(source)).toBe(false);
    for (const name of [...(needs.lib ?? []), ...(needs.prompts ?? [])]) {
      const kind = evalIn(source, 'typeof ' + name);
      expect(kind, nodeKey + ' -> ' + name).not.toBe('undefined');
    }
  });

  it('drops helpers a node does not use', () => {
    expect(helpersFor('wf-b/Compose Draft Prompt')).not.toContain('FALLBACK_PROFILE');
    expect(helpersFor('wf-a/Prepare Email')).not.toContain('function voiceGuard');
    expect(helpersFor('wf-c/Build Digest')).not.toContain('function normalizeGraphMessage');
  });

  it('trims a forward chain identically to the module', () => {
    const source = helpersFor('wf-a/Prepare Email');
    const msg = graphMessage();
    expect(evalIn(source, 'normalizeGraphMessage(' + JSON.stringify(msg) + ')')).toEqual(
      lib.normalizeGraphMessage(msg),
    );
  });

  it('converts an Outlook html body identically to the module', () => {
    const source = helpersFor('wf-a/Prepare Email');
    const msg = graphMessage({ body: { contentType: 'html', content: OUTLOOK_HTML_BODY } });
    expect(evalIn(source, 'normalizeGraphMessage(' + JSON.stringify(msg) + ')')).toEqual(
      lib.normalizeGraphMessage(msg),
    );
    expect(DEEP_FORWARD_CHAIN.length).toBeGreaterThan(0);
  });

  it('still forces a pricing request to flag_for_yancey', () => {
    const source = helpersFor('wf-a/Apply Classification');
    const result = evalIn(
      source,
      'normalizeClassification({ label: "opportunity", confidence: 0.95, requires_draft: true }, ' +
        '{ cleanText: "Can you provide an estimate of timeline and cost?", subject: "scope", senderEmail: "a@b.com", confidenceMin: 0.6 })',
    );
    expect(result.label).toBe('flag_for_yancey');
    expect(result.requiresDraft).toBe(false);
  });

  it('still blocks an invented fee', () => {
    const source = helpersFor('wf-b/Voice Guard');
    const blocked = evalIn(source, 'voiceGuard("The engagement runs about $180,000.", "send the workbook")');
    expect(blocked.ok).toBe(false);
    const allowed = evalIn(source, 'voiceGuard("I will send the workbook.", "send the workbook")');
    expect(allowed.ok).toBe(true);
  });

  it('still renders a digest with the no-send promise', () => {
    const source = helpersFor('wf-c/Build Digest');
    const html = evalIn(
      source,
      'renderDigestHtml(bucketDigestRows([], { now: "2026-08-26T12:30:00Z", liveDraftIds: [], followupDays: 3 }), { dateLabel: "Wednesday, August 26" })',
    );
    expect(html).toContain('never sends anything');
    expect(html).toContain('Needs you');
  });

  it('carries the full classifier prompt, profile placeholder included', () => {
    const source = helpersFor('wf-a/Resolve Context');
    const prompt = evalIn(source, 'buildClassifySystemPrompt("THE-PROFILE")');
    expect(prompt).toContain('THE-PROFILE');
    expect(prompt).toContain('flag_for_yancey');
    expect(prompt).toContain('=== END BUSINESS PROFILE ===');
  });
});
