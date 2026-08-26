import { describe, expect, it } from 'vitest';
import {
  addBusinessDays,
  bucketDigestRows,
  buildClassifierInput,
  businessDaysBetween,
  decideCategories,
  extractModelText,
  htmlToText,
  injectReplyIntoDraftHtml,
  normalizeClassification,
  normalizeGraphMessage,
  isPassThroughHead,
  parseModelJson,
  pricingGuard,
  renderDigestHtml,
  sensitivityGuard,
  stripLegalFooter,
  stripQuotedHistory,
  textToHtmlParagraphs,
  truncate,
  voiceGuard,
} from './agent-lib.js';
import {
  BARE_FORWARD,
  DEEP_FORWARD_CHAIN,
  NO_QUOTE_BODY,
  OUTLOOK_HTML_BODY,
  graphMessage,
} from './fixtures.js';

describe('htmlToText', () => {
  it('strips Outlook markup and decodes entities', () => {
    const text = htmlToText(OUTLOOK_HTML_BODY);
    expect(text).toContain('Hi Yancey,');
    expect(text).toContain('estimate of timeline and cost?');
    expect(text).toContain('Kris Murray');
    expect(text).not.toContain('<p');
    expect(text).not.toContain('MsoNormal');
    expect(text).not.toContain('&nbsp;');
  });

  it('drops style and script blocks entirely', () => {
    expect(htmlToText('<style>.x{color:red}</style><p>Body</p>')).toBe('Body');
    expect(htmlToText('<script>alert(1)</script><p>Body</p>')).toBe('Body');
  });

  it('turns list items and breaks into readable lines', () => {
    expect(htmlToText('<ul><li>One</li><li>Two</li></ul>')).toBe('- One\n- Two');
    expect(htmlToText('a<br>b')).toBe('a\nb');
  });

  it('decodes numeric and hex entities', () => {
    expect(htmlToText('<p>Yancey&#39;s deck &#x2014; ready</p>')).toBe("Yancey's deck — ready");
  });

  it('returns empty string for empty input', () => {
    expect(htmlToText('')).toBe('');
    expect(htmlToText(/** @type {any} */ (undefined))).toBe('');
  });
});

describe('stripQuotedHistory', () => {
  it('keeps only the newest message in a deep forward chain', () => {
    const result = stripQuotedHistory(DEEP_FORWARD_CHAIN);
    expect(result.trimmed).toBe(true);
    expect(result.text).toContain('crossarm numbers');
    expect(result.text).toContain('Chuck');
    expect(result.text).not.toContain('jessica.doucette@sce.com');
    expect(result.text).not.toContain('Original Message');
    expect(result.text).not.toContain('68% reduction');
  });

  it('counts every boundary in the chain, not just the first', () => {
    expect(stripQuotedHistory(DEEP_FORWARD_CHAIN).boundaries).toBeGreaterThanOrEqual(3);
  });

  it('extends past the boundary when the new text is only "FYI"', () => {
    const result = stripQuotedHistory(BARE_FORWARD);
    expect(result.text).toContain('FYI');
    expect(result.text).toContain('leader discipline session');
    expect(result.extended).toBeGreaterThan(0);
  });

  it('leaves a message with no quoted history untouched', () => {
    const result = stripQuotedHistory(NO_QUOTE_BODY);
    expect(result.trimmed).toBe(false);
    expect(result.boundaries).toBe(0);
    expect(result.text).toBe(NO_QUOTE_BODY);
  });

  it('does not cut on the word "from:" used in ordinary prose', () => {
    const prose = 'The data from: the Pulse dashboard is attached.\n\nThanks,\nYancey';
    expect(stripQuotedHistory(prose).trimmed).toBe(false);
  });

  it('cuts on a Gmail-style attribution line', () => {
    const body = 'Sounds good.\n\nOn Mon, Aug 3, 2026 at 9:00 AM Phil Winter wrote:\nOriginal content here';
    const result = stripQuotedHistory(body);
    expect(result.text).toBe('Sounds good.');
  });

  it('cuts on a run of ">" quoted lines', () => {
    const body = 'Agreed, let us proceed with the sprint plan.\n\n> earlier line one\n> earlier line two';
    expect(stripQuotedHistory(body).text).toBe('Agreed, let us proceed with the sprint plan.');
  });

  it('never returns an empty body', () => {
    const onlyQuote = '________________________________\nFrom: a@b.com\nSent: today\nTo: c@d.com\nSubject: x\n\nreal content';
    expect(stripQuotedHistory(onlyQuote).text.length).toBeGreaterThan(0);
  });
});

describe('stripLegalFooter', () => {
  it('removes a confidentiality notice appended after real content', () => {
    const body = 'Here is the deck for review, let me know your thoughts before Friday.\n\nCONFIDENTIALITY NOTICE: this message is intended only for the addressee.';
    expect(stripLegalFooter(body)).toBe('Here is the deck for review, let me know your thoughts before Friday.');
  });

  it('does not swallow a message that opens with the marker', () => {
    const body = 'Confidentiality notice questions came up in the SLT session.';
    expect(stripLegalFooter(body)).toBe(body);
  });
});

describe('truncate', () => {
  it('flags truncation past the limit', () => {
    const result = truncate('x'.repeat(50), 10);
    expect(result.truncated).toBe(true);
    expect(result.text).toContain('[...truncated...]');
  });

  it('leaves short text alone', () => {
    expect(truncate('short', 100)).toEqual({ text: 'short', truncated: false });
  });
});

describe('normalizeGraphMessage', () => {
  it('normalises a text-bodied Graph message', () => {
    const n = normalizeGraphMessage(graphMessage());
    expect(n.messageId).toBe('AAMkAGI2_msg_001');
    expect(n.senderEmail).toBe('ccoleman@partnersperform.com');
    expect(n.senderName).toBe('Charles Coleman');
    expect(n.hasAttachments).toBe(true);
    expect(n.categories).toEqual(['Blue category']);
    expect(n.cleanText).toContain('crossarm numbers');
    expect(n.cleanText).not.toContain('Original Message');
  });

  it('converts an html body before trimming quotes', () => {
    const n = normalizeGraphMessage(
      graphMessage({ body: { contentType: 'html', content: OUTLOOK_HTML_BODY } }),
    );
    expect(n.bodyContentType).toBe('html');
    expect(n.cleanText).toContain('estimate of timeline and cost?');
    expect(n.cleanText).not.toContain('<p');
  });

  it('normalises CRLF in a text body', () => {
    const n = normalizeGraphMessage(
      graphMessage({ body: { contentType: 'text', content: 'line one\r\nline two\r\n\r\nline three' } }),
    );
    expect(n.cleanText).not.toContain('\r');
    expect(n.cleanText).toBe('line one\nline two\n\nline three');
  });

  it('survives a message with no body at all', () => {
    const n = normalizeGraphMessage({ id: 'x' });
    expect(n.messageId).toBe('x');
    expect(n.subject).toBe('(no subject)');
    expect(n.cleanText).toBe('');
  });

  it('builds a classifier payload that carries the routing metadata', () => {
    const input = buildClassifierInput(normalizeGraphMessage(graphMessage()));
    expect(input).toContain('SUBJECT: RE: Vector CI working session');
    expect(input).toContain('HAS_ATTACHMENTS: yes');
    expect(input).toContain('FORWARD_CHAIN_DEPTH:');
    expect(input).toContain('--- EMAIL BODY (newest message only) ---');
  });
});

describe('parseModelJson', () => {
  it('accepts a plain object', () => {
    expect(parseModelJson({ label: 'noise' }).label).toBe('noise');
  });

  it('unwraps common OpenAI node envelopes', () => {
    expect(parseModelJson({ output: { label: 'admin' } }).label).toBe('admin');
    expect(parseModelJson({ message: { content: '{"label":"opportunity"}' } }).label).toBe('opportunity');
    expect(parseModelJson({ content: [{ text: '{"label":"noise"}' }] }).label).toBe('noise');
  });

  it('strips markdown fences', () => {
    expect(parseModelJson('```json\n{"label":"admin"}\n```').label).toBe('admin');
  });

  it('returns null when there is no JSON to find', () => {
    expect(parseModelJson('sorry, I cannot help')).toBeNull();
    expect(parseModelJson(null)).toBeNull();
  });
});

describe('pricingGuard — spec Gate 3', () => {
  it('fires on a request for an estimate with no figure in-thread', () => {
    const guard = pricingGuard(
      'Is this possible, and can you provide an estimate of timeline and cost?',
      'ILI program scoping',
    );
    expect(guard.triggered).toBe(true);
  });

  it('stands down when Yancey already stated a figure in-thread', () => {
    const guard = pricingGuard(
      'Following your $175-200K fixed-fee range, can you provide an estimate of the schedule?',
      'Scope',
    );
    expect(guard.triggered).toBe(false);
  });

  it('does not fire on ordinary delivery talk', () => {
    expect(pricingGuard('The O&M cost baseline slide is attached.', 'Deck').triggered).toBe(false);
  });

  // Regression: Gate 1 on Yancey's own mail (opp-05, "Re: Pulse Power BI").
  // NIPSCO asked for an estimate of timeline and cost; two replies later the
  // newest message was only "let's write a proposal of what it would take".
  // Scanning the quote-stripped newest message missed the fee ask entirely and
  // the agent drafted a reply about a proposal.
  it('fires when the fee ask is only in the quoted chain', () => {
    const newest = "Let's write a proposal of what it would take, including NIPSCO responsibility and access requirements.";
    const thread =
      newest +
      '\n\nFrom: DeFelice, Louis\nSubject: Pulse Power BI\n\n' +
      'Is this possible and if are you able to provide an estimate of timeline and cost? Thanks';

    expect(pricingGuard(newest, 'Re: Pulse Power BI').triggered).toBe(false);
    expect(pricingGuard(thread, 'Re: Pulse Power BI').triggered).toBe(true);
  });

  it('recognises commercial scope terms', () => {
    expect(
      pricingGuard('Can you share a statement of work for the T-Week engagement?', 'SOW').triggered,
    ).toBe(true);
  });

  it('does not fire on a deck merely called a proposal', () => {
    expect(
      pricingGuard(
        'Team, here are my thoughts for a deck to guide the August 3 SLT kickoff. Weigh in as you see fit.',
        'DRAFT Proposal for Kickoff Meeting Deck',
      ).triggered,
    ).toBe(false);
  });
});

describe('thread-wide guard wiring', () => {
  it('normalizeGraphMessage exposes the quoted chain as threadText', () => {
    const n = normalizeGraphMessage(graphMessage({ body: { contentType: 'text', content: DEEP_FORWARD_CHAIN } }));
    expect(n.threadText.length).toBeGreaterThan(n.cleanText.length);
    expect(n.cleanText).not.toContain('Original Message');
  });

  it('normalizeClassification flags pricing found only in threadText', () => {
    const raw = {
      label: 'opportunity',
      confidence: 0.98,
      reasoning: 'Chuck is directing a proposal.',
      requires_draft: true,
      urgency: 'this_week',
      asks_of_yancey: ['write a proposal'],
    };
    const ctx = {
      cleanText: "Let's write a proposal of what it would take.",
      threadText:
        "Let's write a proposal of what it would take.\n\nAre you able to provide an estimate of timeline and cost?",
      subject: 'Re: Pulse Power BI',
      senderEmail: 'ccoleman@partnersperform.com',
      confidenceMin: 0.6,
    };
    const verdict = normalizeClassification(raw, ctx);
    expect(verdict.label).toBe('flag_for_yancey');
    expect(verdict.requiresDraft).toBe(false);
    expect(verdict.guardFlags).toContain('pricing_guard');
  });
});

describe('sensitivityGuard', () => {
  it('fires on legal subject matter', () => {
    expect(sensitivityGuard('a@b.com', 'Our attorney has raised a breach of contract concern.', '', []).triggered).toBe(true);
  });

  it('fires on an always-flag sender', () => {
    expect(sensitivityGuard('legal@sce.com', 'hi', '', ['legal@sce.com']).triggered).toBe(true);
  });

  it('fires on an always-flag domain', () => {
    expect(sensitivityGuard('anyone@lawfirm.com', 'hi', '', ['@lawfirm.com']).triggered).toBe(true);
  });

  it('stays quiet on normal engagement traffic', () => {
    expect(sensitivityGuard('lou@nisource.com', 'CAIDI sprint update', '', []).triggered).toBe(false);
  });
});

describe('normalizeClassification', () => {
  const base = {
    cleanText: 'Attaching the SLT deck for your review before the Momentum Room.',
    subject: 'Vector deck',
    senderEmail: 'ccoleman@partnersperform.com',
    confidenceMin: 0.6,
  };

  it('passes a confident opportunity through as draftable', () => {
    const c = normalizeClassification(
      { label: 'opportunity', confidence: 0.9, requires_draft: true, urgency: 'this_week', reasoning: 'live engagement work' },
      base,
    );
    expect(c.label).toBe('opportunity');
    expect(c.requiresDraft).toBe(true);
    expect(c.guards).toEqual([]);
  });

  it('falls back to flag_for_yancey when the model output is unusable', () => {
    const c = normalizeClassification('total nonsense', base);
    expect(c.label).toBe('flag_for_yancey');
    expect(c.requiresDraft).toBe(false);
    expect(c.guards).toContain('unparseable_label');
  });

  it('never drafts for noise', () => {
    const c = normalizeClassification({ label: 'noise', confidence: 0.99, requires_draft: true }, base);
    expect(c.requiresDraft).toBe(false);
  });

  it('never drafts for flag_for_yancey', () => {
    const c = normalizeClassification({ label: 'flag_for_yancey', confidence: 0.99, requires_draft: true }, base);
    expect(c.requiresDraft).toBe(false);
  });

  it('withholds the draft below the confidence floor', () => {
    const c = normalizeClassification({ label: 'opportunity', confidence: 0.4, requires_draft: true }, base);
    expect(c.requiresDraft).toBe(false);
    expect(c.guards).toContain('below_confidence_min');
  });

  it('overrides a confident opportunity when the pricing guard fires', () => {
    const c = normalizeClassification(
      { label: 'opportunity', confidence: 0.95, requires_draft: true },
      { ...base, cleanText: 'Can you provide an estimate of timeline and cost for the ILI work?' },
    );
    expect(c.label).toBe('flag_for_yancey');
    expect(c.requiresDraft).toBe(false);
    expect(c.guards).toContain('pricing_guard');
  });

  it('rescales a percentage-style confidence', () => {
    expect(normalizeClassification({ label: 'noise', confidence: 85 }, base).confidence).toBe(0.85);
  });
});

describe('decideCategories', () => {
  it('preserves the user categories and adds the AI markers', () => {
    expect(decideCategories(['Blue category'], 'opportunity')).toEqual([
      'Blue category',
      'AI-Processed',
      'AI-Action',
    ]);
  });

  it('marks noise as processed only', () => {
    expect(decideCategories([], 'noise')).toEqual(['AI-Processed']);
  });

  it('marks a flagged message for review', () => {
    expect(decideCategories([], 'flag_for_yancey')).toEqual(['AI-Processed', 'AI-Review']);
  });

  it('is idempotent when re-run over its own output', () => {
    const once = decideCategories(['Blue category'], 'opportunity');
    expect(decideCategories(once, 'opportunity')).toEqual(once);
  });
});

describe('voiceGuard — spec Gate 2', () => {
  const source = 'Can you send the updated crossarm workbook before the Momentum Room?';

  it('blocks a draft that invents a dollar figure', () => {
    const guard = voiceGuard('Happy to help. The engagement runs about $180,000.\n\nThanks,\n\nYancey', source);
    expect(guard.ok).toBe(false);
    expect(guard.hard.join()).toContain('invented_currency');
  });

  it('blocks a draft that invents a percentage', () => {
    const guard = voiceGuard('We saw a 68% reduction on the last program.\n\nThanks,\n\nYancey', source);
    expect(guard.ok).toBe(false);
    expect(guard.hard.join()).toContain('invented_percent');
  });

  it('allows a figure that was already in the inbound message', () => {
    const guard = voiceGuard(
      'Confirming the $175,000 figure you mentioned.\n\nThanks,\n\nYancey',
      'We understood the fixed fee to be $175,000 for the program.',
    );
    expect(guard.ok).toBe(true);
  });

  it('flags an invented date as medium rather than blocking', () => {
    const guard = voiceGuard('I will send it on March 12.\n\nThanks,\n\nYancey', source);
    expect(guard.ok).toBe(true);
    expect(guard.medium.join()).toContain('invented_date');
  });

  it('removes the forbidden opener', () => {
    const guard = voiceGuard('I hope this email finds you well. Sending the workbook today.\n\nThanks,\n\nYancey', source);
    expect(guard.text).not.toMatch(/finds you well/i);
    expect(guard.soft).toContain('removed_forbidden_opener');
  });

  it('appends the sign-off when the model omits it', () => {
    const guard = voiceGuard('Sending the workbook today.', source);
    expect(guard.text).toMatch(/Thanks,\n\nYancey$/);
    expect(guard.soft).toContain('appended_signoff');
  });

  it('rewrites a wrong sign-off into his own', () => {
    const guard = voiceGuard('Sending the workbook today.\n\nBest regards,\nYancey Dawson', source);
    expect(guard.text).toMatch(/Thanks,\n\nYancey$/);
    expect(guard.text).not.toMatch(/Best regards/i);
  });

  it('leaves a correct sign-off untouched', () => {
    const guard = voiceGuard('Sending the workbook today.\n\nThanks,\n\nYancey', source);
    expect(guard.soft).not.toContain('appended_signoff');
  });

  it('strips a stray code fence and subject line', () => {
    const guard = voiceGuard('```\nSubject: RE: workbook\nSending it over.\n\nThanks,\n\nYancey\n```', source);
    expect(guard.text.startsWith('Sending it over.')).toBe(true);
  });

  it('rejects an empty draft', () => {
    const guard = voiceGuard('   ', source);
    expect(guard.ok).toBe(false);
    expect(guard.hard).toContain('empty_draft');
  });

  it('reports over-length as medium', () => {
    const long = 'word '.repeat(300) + '\n\nThanks,\n\nYancey';
    expect(voiceGuard(long, source).medium.join()).toContain('over_length');
  });
});

describe('textToHtmlParagraphs / injectReplyIntoDraftHtml', () => {
  it('wraps paragraphs and escapes html', () => {
    const html = textToHtmlParagraphs('Hi <Dean>,\nline two\n\nThanks,\n\nYancey');
    expect(html).toContain('&lt;Dean&gt;');
    expect(html).toContain('<br>');
    expect(html.match(/<p>/g)).toHaveLength(3);
  });

  it('inserts the reply just inside the body tag, keeping the quoted history', () => {
    const draft = '<html><body><div>quoted original</div></body></html>';
    const out = injectReplyIntoDraftHtml(draft, '<p>my reply</p>');
    expect(out.indexOf('my reply')).toBeLessThan(out.indexOf('quoted original'));
    expect(out).toContain('quoted original');
  });

  it('falls back to prepending when there is no body tag', () => {
    const out = injectReplyIntoDraftHtml('<div>quoted</div>', '<p>reply</p>');
    expect(out.startsWith('<p>reply</p>')).toBe(true);
  });
});

describe('business day helpers', () => {
  it('skips the weekend when adding days', () => {
    // 2026-08-20 is a Thursday.
    expect(addBusinessDays(new Date('2026-08-20T12:00:00Z'), 3).toISOString().slice(0, 10)).toBe('2026-08-25');
  });

  it('counts only weekdays between two dates', () => {
    expect(businessDaysBetween(new Date('2026-08-20T12:00:00Z'), new Date('2026-08-25T12:00:00Z'))).toBe(3);
  });

  it('returns zero when the range is reversed', () => {
    expect(businessDaysBetween(new Date('2026-08-25T12:00:00Z'), new Date('2026-08-20T12:00:00Z'))).toBe(0);
  });
});

describe('bucketDigestRows', () => {
  const now = '2026-08-26T12:30:00Z';

  const rows = [
    {
      id: 1,
      message_id: 'm1',
      subject: 'SLT deck review',
      sender_name: 'Jessica Doucette',
      label: 'opportunity',
      urgency: 'today',
      reasoning: 'client reply on active deliverable',
      status: 'classified',
      received_at: '2026-08-26T08:00:00Z',
      web_link: 'https://outlook/m1',
    },
    {
      id: 2,
      message_id: 'm2',
      subject: 'Crossarm workbook',
      sender_name: 'Chuck Coleman',
      label: 'opportunity',
      urgency: 'this_week',
      status: 'drafted',
      draft_id: 'd2',
      draft_created_at: '2026-08-24T09:00:00Z',
      received_at: '2026-08-24T08:00:00Z',
    },
    {
      id: 3,
      message_id: 'm3',
      subject: 'ILI scoping',
      sender_name: 'Kris Murray',
      label: 'opportunity',
      status: 'drafted',
      draft_id: 'd3-gone',
      draft_created_at: '2026-08-21T09:00:00Z',
      received_at: '2026-08-21T08:00:00Z',
    },
    {
      id: 4,
      message_id: 'm4',
      subject: 'Damage prevention proposal',
      sender_name: 'Phil Winter',
      label: 'opportunity',
      status: 'classified',
      received_at: '2026-08-17T08:00:00Z',
    },
    {
      id: 5,
      message_id: 'm5',
      subject: 'Paychex retirement plans',
      label: 'noise',
      status: 'classified',
      received_at: '2026-08-26T07:00:00Z',
    },
    {
      id: 6,
      message_id: 'm6',
      subject: 'already handled',
      label: 'opportunity',
      status: 'classified',
      received_at: '2026-08-26T07:00:00Z',
      resolved_at: '2026-08-26T09:00:00Z',
    },
  ];

  const buckets = bucketDigestRows(rows, { now, liveDraftIds: ['d2'], followupDays: 3 });

  it('puts fresh unactioned work in Needs You', () => {
    expect(buckets.needsYou.map((r) => r.subject)).toEqual(['SLT deck review']);
  });

  it('lists drafts that still exist in the mailbox', () => {
    expect(buckets.draftsWaiting.map((r) => r.subject)).toEqual(['Crossarm workbook']);
    expect(buckets.draftsWaiting[0].waitingDays).toBe(2);
  });

  it('treats a vanished draft as sent and marks the row resolved', () => {
    expect(buckets.resolvedIds).toEqual(['3']);
    expect(buckets.resolvedMessageIds).toEqual(['m3']);
  });

  it('surfaces a thread that has gone quiet past the threshold', () => {
    expect(buckets.followupsDue.map((r) => r.subject)).toEqual(['Damage prevention proposal']);
  });

  it('never shows noise or already-resolved rows', () => {
    const all = JSON.stringify(buckets);
    expect(all).not.toContain('Paychex');
    expect(all).not.toContain('already handled');
  });

  it('sorts Needs You by urgency', () => {
    const sorted = bucketDigestRows(
      [
        { id: 1, message_id: 'a', subject: 'fyi one', label: 'opportunity', urgency: 'fyi', received_at: now },
        { id: 2, message_id: 'b', subject: 'today one', label: 'opportunity', urgency: 'today', received_at: now },
      ],
      { now, liveDraftIds: [], followupDays: 3 },
    );
    expect(sorted.needsYou.map((r) => r.subject)).toEqual(['today one', 'fyi one']);
  });

  it('handles an empty log without throwing', () => {
    const empty = bucketDigestRows([], { now, liveDraftIds: [], followupDays: 3 });
    expect(empty.stats).toEqual({
      needsYou: 0,
      draftsWaiting: 0,
      followupsDue: 0,
      resolved: 0,
      autoClosed: 0,
      scanned: 0,
    });
  });

  it('auto-closes an undrafted row that has gone stale', () => {
    const stale = bucketDigestRows(
      [{ id: 9, message_id: 'old', subject: 'ancient', label: 'opportunity', received_at: '2026-06-01T08:00:00Z' }],
      { now, liveDraftIds: [], followupDays: 3, autoCloseDays: 21 },
    );
    expect(stale.stats.autoClosed).toBe(1);
    expect(stale.resolvedMessageIds).toEqual(['old']);
    expect(stale.followupsDue).toHaveLength(0);
  });

  it('ignores synthetic empty items from alwaysOutputData', () => {
    expect(bucketDigestRows([{}], { now, liveDraftIds: [], followupDays: 3 }).stats.needsYou).toBe(0);
  });
});

describe('renderDigestHtml', () => {
  const buckets = bucketDigestRows(
    [
      {
        id: 1,
        message_id: 'm1',
        subject: 'Deck <review> & sign-off',
        sender_name: 'Jessica Doucette',
        label: 'flag_for_yancey',
        urgency: 'today',
        reasoning: 'asks for a fee',
        guard_flags: 'pricing_guard',
        received_at: '2026-08-26T08:00:00Z',
        web_link: 'https://outlook/m1',
      },
    ],
    { now: '2026-08-26T12:30:00Z', liveDraftIds: [], followupDays: 3 },
  );

  it('escapes user content into the html', () => {
    const html = renderDigestHtml(buckets, { dateLabel: 'Wednesday, August 26' });
    expect(html).toContain('Deck &lt;review&gt; &amp; sign-off');
    expect(html).not.toContain('Deck <review>');
  });

  it('renders all three sections plus the no-send promise', () => {
    const html = renderDigestHtml(buckets, { dateLabel: 'Wednesday, August 26' });
    expect(html).toContain('Needs you');
    expect(html).toContain('Drafts waiting');
    expect(html).toContain('Follow-ups due');
    expect(html).toContain('cannot send email');
  });

  it('shows empty-state copy rather than a blank section', () => {
    const html = renderDigestHtml(buckets, { dateLabel: 'x' });
    expect(html).toContain('No drafts waiting.');
  });

  it('adds the shadow-mode banner when drafting is off', () => {
    const html = renderDigestHtml(buckets, { dateLabel: 'x', mode: 'shadow' });
    expect(html).toContain('Shadow mode');
  });

  it('marks a flagged item as never drafted', () => {
    expect(renderDigestHtml(buckets, { dateLabel: 'x' })).toContain('FLAGGED — no draft written');
  });
});

describe('isPassThroughHead', () => {
  it('treats pointer phrases as pass-through', () => {
    for (const head of ['FYI', 'fyi.', 'See below', 'Please review', 'Thoughts?', 'For your awareness', '']) {
      expect(isPassThroughHead(head)).toBe(true);
    }
  });

  it('keeps a short but self-contained reply', () => {
    for (const head of ['Sounds good.', 'Received! Thank you!', 'Approved.', 'Agreed, let us proceed.']) {
      expect(isPassThroughHead(head)).toBe(false);
    }
  });

  it('ignores a trailing signature line when judging the head', () => {
    expect(isPassThroughHead('FYI\n\nChuck')).toBe(true);
    expect(isPassThroughHead('Please see below and confirm the crossarm numbers.\n\nChuck')).toBe(false);
  });
});

describe('extractModelText', () => {
  it('reads a plain string', () => {
    expect(extractModelText('hello')).toBe('hello');
  });

  it('unwraps the OpenAI node envelopes', () => {
    expect(extractModelText({ output_text: 'a' })).toBe('a');
    expect(extractModelText({ content: [{ text: 'b' }] })).toBe('b');
    expect(extractModelText({ message: { content: 'c' } })).toBe('c');
    expect(extractModelText({ output: [{ content: [{ text: 'd' }] }] })).toBe('d');
  });

  it('returns an empty string when there is no text', () => {
    expect(extractModelText({ usage: { tokens: 5 } })).toBe('');
    expect(extractModelText(null)).toBe('');
  });
});
