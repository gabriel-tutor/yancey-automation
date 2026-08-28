import { describe, expect, it } from 'vitest';
import { codeNode } from '../scripts/bundle.mjs';
import { HELPERS } from '../workflows/helpers-manifest.mjs';
import { VOICE_GUARD } from '../workflows/wf-b-draft.mjs';
import { APPLY_CLASSIFICATION } from '../workflows/wf-a-classify.mjs';

/**
 * The Code node *bodies* — not just the helper library they import.
 *
 * Two production bugs got in through this gap: a paste that silently dropped a
 * non-breaking space, and an orphaned draft left in the mailbox when the model
 * failed. Both were invisible to the unit tests because the node bodies were
 * only ever strings. These run them.
 */

const voiceGuardCode = await codeNode(HELPERS['wf-b/Voice Guard'], VOICE_GUARD);
const applyClassificationCode = await codeNode(
  HELPERS['wf-a/Apply Classification'],
  APPLY_CLASSIFICATION,
);

/**
 * @param {Record<string, any>} map
 * @returns {(name: string) => any}
 */
const nodeRef = (map) => (name) => {
  if (!(name in map)) throw new Error("node '" + name + "' did not run");
  const value = map[name];
  return { first: () => ({ json: value }), all: () => [{ json: value }], item: { json: value } };
};

/**
 * @param {string} body
 * @param {Record<string, any>} globals
 * @returns {any}
 */
const run = (body, globals) => {
  const names = Object.keys(globals);
  return new Function(...names, body)(...names.map((n) => globals[n]));
};

describe('WF-B Voice Guard node body', () => {
  const cfg = { maxDraftWords: 150, signOff: 'Thanks,', signature: 'Yancey' };
  const req = {
    messageId: 'AAMkMSG',
    messageIdEnc: 'AAMkMSG%3D',
    subject: 'RE: Vector CI',
    cleanText: 'Please send the crossarm workbook.',
  };
  const composed = {
    draftId: 'AAMkDRAFT',
    draftIdEnc: 'AAMkDRAFT%3D',
    draftWebLink: 'https://outlook/x',
    draftBodyHtml: '<html><body><div>quoted</div></body></html>',
  };
  /** @param {any} inputJson */
  const globals = (inputJson) => ({
    $: nodeRef({ Config: cfg, 'Draft Request': req, 'Compose Draft Prompt': composed }),
    $input: { first: () => ({ json: inputJson }) },
  });

  it('writes a reply when the model returns one', () => {
    const out = run(voiceGuardCode, globals({
      content: 'Hey Chuck,\n\nI will pull the crossarm workbook together before the Momentum Room.\n\nThanks,\n\nYancey',
    }));
    expect(out[0].json.ok).toBe(true);
    expect(out[0].json.status).toBe('drafted');
    expect(out[0].json.finalHtml).toContain('Hey Chuck');
  });

  // The orphaned-draft bug. createReply has already put an empty reply in the
  // mailbox; if this throws, only the blocked branch would have deleted it.
  it('absorbs a model failure and routes it to the cleanup branch', () => {
    const out = run(voiceGuardCode, globals({ error: { message: 'OpenAI 500: upstream timeout' } }));
    expect(out[0].json.ok).toBe(false);
    expect(out[0].json.status).toBe('draft_failed');
    expect(out[0].json.guardFlags).toBe('draft_generation_failed');
    expect(out[0].json.finalHtml).toBe('');
    expect(out[0].json.extraCategories).toEqual(['AI-Review']);
  });

  it('keeps every identifier the cleanup branch needs', () => {
    const out = run(voiceGuardCode, globals({ error: 'boom' }));
    // Deleting the empty draft needs this one.
    expect(out[0].json.draftIdEnc).toBe('AAMkDRAFT%3D');
    // Tagging and logging the original message need these.
    expect(out[0].json.messageId).toBe('AAMkMSG');
    expect(out[0].json.messageIdEnc).toBe('AAMkMSG%3D');
  });
});

describe('WF-A Apply Classification node body', () => {
  const cfg = {
    confidenceMin: 0.6,
    followupDays: 3,
    alwaysFlagSenders: [],
    draftingMailbox: 'gtutor@partnersperform.com',
  };
  const prepared = {
    messageId: 'M1', messageIdEnc: 'M1', conversationId: 'C1', internetMessageId: '<x>',
    receivedAt: '2026-08-27T09:00:00Z',
    senderEmail: 'ccoleman@partnersperform.com', senderName: 'Charles Coleman',
    subject: 'RE: Vector CI', webLink: 'https://o/x', hasAttachments: false, categories: [],
    cleanText: 'Can you pull the workbook together before Thursday?',
    threadText: 'Can you pull the workbook together before Thursday?',
    toRecipients: ['gtutor@partnersperform.com'], ccRecipients: [],
  };
  const model = {
    label: 'opportunity', confidence: 0.95, reasoning: 'live work',
    requires_draft: true, urgency: 'this_week', asks_of_yancey: ['send workbook'],
  };
  /** @param {any} prep */
  const globals = (prep) => ({
    $: nodeRef({ Config: cfg, 'Prepare Email': prep, 'Resolve Context': { profileSource: 'file' } }),
    $json: model,
    $itemIndex: 0,
  });

  it('drafts for the mailbox drafting was authorised for', () => {
    const out = run(applyClassificationCode, globals(prepared));
    expect(out.json.requiresDraft).toBe(true);
    expect(out.json.guardFlags).toBe('');
  });

  it('accepts the mailbox on cc, in any case', () => {
    const out = run(applyClassificationCode, globals({
      ...prepared,
      toRecipients: ['kdunlap@partnersperform.com'],
      ccRecipients: ['GTutor@PartnersPerform.com'],
    }));
    expect(out.json.requiresDraft).toBe(true);
  });

  // The credential-swap case: the agent is suddenly reading a different
  // mailbox. It must keep classifying, and stop writing.
  it('stops drafting when the mailbox is not the authorised one', () => {
    const out = run(applyClassificationCode, globals({
      ...prepared,
      toRecipients: ['ydawson@partnersperform.com'],
      ccRecipients: [],
    }));
    expect(out.json.requiresDraft).toBe(false);
    expect(out.json.guardFlags).toBe('mailbox_not_addressed');
    expect(out.json.label).toBe('opportunity');
  });

  it('still applies the pricing guard from the whole thread', () => {
    const out = run(applyClassificationCode, globals({
      ...prepared,
      cleanText: "Let's put together what it would take.",
      threadText: "Let's put together what it would take.\n\nCan you provide an estimate of timeline and cost?",
    }));
    expect(out.json.label).toBe('flag_for_yancey');
    expect(out.json.requiresDraft).toBe(false);
    expect(out.json.guardFlags).toContain('pricing_guard');
  });
});

/**
 * The same node body, configured exactly as production is: Yancey's rules of
 * 2026-08-28. This is the layer that actually ships — the library is tested
 * separately, but a config key spelled wrong in the node would pass there and
 * fail here.
 */
describe("WF-A Apply Classification under Yancey's rules", () => {
  const cfg = {
    confidenceMin: 0.6,
    followupDays: 3,
    alwaysFlagSenders: ['ccoleman@partnersperform.com'],
    internalDomains: ['partnersperform.com'],
    draftingMailbox: 'YDawson@partnersperform.com',
  };
  const base = {
    messageId: 'M1', messageIdEnc: 'M1', conversationId: 'C1', internetMessageId: '<x>',
    receivedAt: '2026-08-27T09:00:00Z',
    subject: 'RE: Vector CI', webLink: 'https://o/x', hasAttachments: false, categories: [],
    cleanText: 'Can you pull the workbook together before Thursday?',
    threadText: 'Can you pull the workbook together before Thursday?',
    toRecipients: ['YDawson@partnersperform.com'], ccRecipients: [],
  };
  const model = {
    label: 'opportunity', confidence: 0.95, reasoning: 'live work',
    requires_draft: true, urgency: 'this_week', asks_of_yancey: ['send workbook'],
  };

  /** @param {any} prep @param {any} [modelOut] */
  const runWith = (prep, modelOut) =>
    run(applyClassificationCode, {
      $: nodeRef({ Config: cfg, 'Prepare Email': prep, 'Resolve Context': { profileSource: 'file' } }),
      $json: modelOut || model,
      $itemIndex: 0,
    });

  it('drafts for a colleague', () => {
    const out = runWith({ ...base, senderEmail: 'scorkran@partnersperform.com', senderName: 'Shawn Corkran' });
    expect(out.json.label).toBe('opportunity');
    expect(out.json.requiresDraft).toBe(true);
    expect(out.json.categories).toContain('AI-Action');
  });

  it('never drafts for Chuck Coleman', () => {
    const out = runWith({ ...base, senderEmail: 'Ccoleman@partnersperform.com', senderName: 'Charles Coleman' });
    expect(out.json.label).toBe('flag_for_yancey');
    expect(out.json.requiresDraft).toBe(false);
    expect(out.json.guardFlags).toContain('always_flag_sender');
    expect(out.json.categories).toContain('AI-Review');
  });

  it('flags an outsider rather than replying to them', () => {
    const out = runWith({ ...base, senderEmail: 'Jessica.Doucette@sce.com', senderName: 'Jessica Doucette' });
    expect(out.json.label).toBe('flag_for_yancey');
    expect(out.json.requiresDraft).toBe(false);
    expect(out.json.guardFlags).toContain('external_sender');
  });

  // The rule he gave is "flag without drafting", not "escalate everything
  // external". Five of his junk fixtures are outside senders; promoting those
  // would put PayPal receipts in the bucket he reads at 07:30.
  it('leaves an outside newsletter as noise', () => {
    const out = runWith(
      { ...base, senderEmail: 'service@paypal.com', senderName: 'PayPal' },
      { label: 'noise', confidence: 0.97, reasoning: 'receipt', requires_draft: false, urgency: 'fyi', asks_of_yancey: [] },
    );
    expect(out.json.label).toBe('noise');
    expect(out.json.status).toBe('resolved');
    expect(out.json.guardFlags).toBe('');
  });

  it('does not resurrect a junk invoice into the review bucket', () => {
    const out = runWith(
      {
        ...base,
        senderEmail: 'service@paypal.com',
        subject: 'Invoice from JEREMIAH CHRISTIAN PLACENTE ROSALES (0028)',
        cleanText: 'You sent a payment request for a contractor invoice.',
        threadText: 'You sent a payment request for a contractor invoice.',
      },
      { label: 'noise', confidence: 0.99, reasoning: 'contractor invoice', requires_draft: false, urgency: 'fyi', asks_of_yancey: [] },
    );
    expect(out.json.label).toBe('noise');
    expect(out.json.guardFlags).toBe('');
    expect(out.json.categories).not.toContain('AI-Review');
  });

  it('names the category that held a colleague message back', () => {
    const out = runWith({
      ...base,
      senderEmail: 'dyoes@partnersperform.com',
      cleanText: 'Can you approve the change order before Thursday?',
      threadText: 'Can you approve the change order before Thursday?',
    });
    expect(out.json.label).toBe('flag_for_yancey');
    expect(out.json.guardFlags).toContain('sensitivity:approval');
  });

  it('still drafts consulting mail that only sounds like a flagged category', () => {
    const out = runWith({
      ...base,
      senderEmail: 'dyoes@partnersperform.com',
      cleanText: 'Sustain benefits through performance monitoring. Reducing operating expense improves cash flow.',
      threadText: 'Sustain benefits through performance monitoring. Reducing operating expense improves cash flow.',
    });
    expect(out.json.label).toBe('opportunity');
    expect(out.json.requiresDraft).toBe(true);
  });
});
