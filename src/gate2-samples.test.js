import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { SAMPLES } from '../scripts/gate2-samples.mjs';
import { addressedToMailbox, normalizeClassification, normalizeGraphMessage } from './agent-lib.js';

/**
 * Gate 2 asks Yancey to score the voice of a reply. That is only worth his time
 * if the assistant would actually write that reply — and for a while it was
 * not: every one of the five samples was mail his own rules hold for review,
 * four from Chuck Coleman and one from outside the domain. He was asked to rate
 * replies that would never be sent.
 *
 * These assertions hold the sample list to the policy that is deployed, so the
 * two cannot drift apart again in silence.
 */

const ROOT = join(import.meta.dirname, '..');
const DEPLOY = JSON.parse(readFileSync(join(ROOT, 'config/deploy.json'), 'utf8'));
const POLICY = DEPLOY.policy;
const TARGET = DEPLOY.targets.prod;

/** The most permissive verdict the classifier could return for this message. */
const MODEL_WANTS_A_DRAFT = {
  label: 'opportunity',
  confidence: 0.95,
  requires_draft: true,
  urgency: 'this_week',
  reasoning: 'live engagement work',
  asks_of_yancey: ['something'],
};

/** @param {string} id */
function verdictFor(id) {
  const fixture = JSON.parse(readFileSync(join(ROOT, 'fixtures/golden', id + '.json'), 'utf8'));
  const n = normalizeGraphMessage(fixture.message, { maxChars: 6000 });
  const authorised = addressedToMailbox(
    n.toRecipients,
    n.ccRecipients,
    [TARGET.draftingMailbox].concat(TARGET.draftingAliases || []),
  );
  return {
    sender: n.senderEmail,
    ...normalizeClassification(MODEL_WANTS_A_DRAFT, {
      cleanText: n.cleanText,
      threadText: n.threadText,
      subject: n.subject,
      senderEmail: n.senderEmail,
      confidenceMin: 0.8,
      alwaysFlagSenders: POLICY.alwaysFlagSenders,
      internalDomains: POLICY.internalDomains,
      authorised,
    }),
  };
}

describe('Gate 2 rates replies that would actually be written', () => {
  it('has samples at all', () => {
    expect(SAMPLES.length).toBeGreaterThan(0);
  });

  it('every sample is one the assistant would draft a reply to', () => {
    const held = SAMPLES.map((s) => ({ id: s.id, ...verdictFor(s.id) })).filter((v) => !v.requiresDraft);
    expect(
      held.map((h) => h.id + ' held by ' + (h.guardFlags || 'policy')),
      'Gate 2 would be asking Yancey to score replies that are never written',
    ).toEqual([]);
  });

  it('no sample is from a sender on the always-flag list', () => {
    for (const sample of SAMPLES) {
      const v = verdictFor(sample.id);
      expect(v.guardFlags, sample.id + ' is from ' + v.sender).not.toContain('always_flag_sender');
    }
  });

  it('no sample is from outside the internal domains', () => {
    for (const sample of SAMPLES) {
      const v = verdictFor(sample.id);
      expect(v.guardFlags, sample.id + ' is from ' + v.sender).not.toContain('external_sender');
    }
  });

  it('describes why each one is worth rating', () => {
    for (const sample of SAMPLES) {
      expect(sample.why, sample.id).toBeTruthy();
      expect(sample.register, sample.id).toBeTruthy();
    }
  });
});
