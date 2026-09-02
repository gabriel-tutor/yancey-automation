import { describe, expect, it } from 'vitest';
import { buildDraftSystemPrompt, buildDraftUserPrompt } from './prompts.js';

/**
 * The drafting prompt used to tell the model, unconditionally, that it had not
 * read the attachments. Now that WF-D does read most of them, that instruction
 * would make the assistant lie about its own work — so the wording has to
 * follow what was actually extracted.
 */

/** @param {Partial<Record<string, any>>} over */
const ctx = (over) =>
  Object.assign(
    {
      subject: 'RE: Vector CI working session',
      senderName: 'Charles Coleman',
      senderEmail: 'ccoleman@partnersperform.com',
      cleanText: 'Numbers are in the workbook.',
      reasoning: 'live engagement work',
      asks: ['confirm the SLT date'],
      hasAttachments: true,
      label: 'opportunity',
      urgency: 'this_week',
    },
    over,
  );

describe('draft prompt — attachments', () => {
  it('quotes the contents when they were read', () => {
    const out = buildDraftUserPrompt(ctx({ attachmentText: '- tracker.xlsx:\nF12-North | 88.4' }));
    expect(out).toContain('Attachments: yes — contents below');
    expect(out).toContain('--- ATTACHMENTS ---');
    expect(out).toContain('F12-North | 88.4');
  });

  it('still says not to summarise when nothing could be read', () => {
    const out = buildDraftUserPrompt(ctx({ attachmentText: '' }));
    expect(out).toContain('acknowledge without summarising');
    expect(out).not.toContain('--- ATTACHMENTS ---');
  });

  it('says none when there are none', () => {
    const out = buildDraftUserPrompt(ctx({ hasAttachments: false }));
    expect(out).toContain('Attachments: none');
  });

  it('puts the attachments before the triage notes, not after the sign-off', () => {
    const out = buildDraftUserPrompt(ctx({ attachmentText: '- tracker.xlsx:\n88.4' }));
    expect(out.indexOf('--- ATTACHMENTS ---')).toBeLessThan(out.indexOf('--- TRIAGE NOTES ---'));
  });

  // The rule has to cover both cases in one sentence, because the system prompt
  // is built once and does not know what this particular message carried.
  it('ties the hard rule to whether contents are present', () => {
    const rules = buildDraftSystemPrompt('profile text', { maxWords: 180 });
    expect(rules).toContain('Acknowledge attachments without claiming to have read them');
    expect(rules).toContain('unless their contents are quoted');
  });
});
