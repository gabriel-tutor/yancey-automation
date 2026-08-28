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
  describeDraftFailure,
  sameAddress,
  addressedToMailbox,
  isPassThroughHead,
  parseModelJson,
  pricingGuard,
  renderDigestHtml,
  normalizeErrorPayload,
  senderDomain,
  isInternalSender,
  sensitivityGuard,
  stripLegalFooter,
  stripQuotedHistory,
  textToHtmlParagraphs,
  truncate,
  voiceGuard,
  voicePreChecks,
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
    // The promise Yancey actually gets to read. He spotted that "cannot send
    // email" sat oddly beside an emailed digest, so both halves are asserted:
    // nothing sends from his mailbox, and the digest comes from elsewhere.
    expect(html).toContain('never sends anything');
    expect(html).toContain('written straight into your inbox');
    expect(html).not.toContain('cannot send email');
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

describe('describeDraftFailure — nothing is left stranded in the mailbox', () => {
  const request = { messageId: 'AAMkMSG', messageIdEnc: 'AAMkMSG%3D' };
  const draft = { draftIdEnc: 'AAMkDRAFT%3D' };

  it('reports a model failure as a fault, not a decision', () => {
    const out = describeDraftFailure({
      item: { error: { message: 'OpenAI 500: upstream timeout' } },
      request,
      draft,
    });
    expect(out.status).toBe('draft_failed');
    expect(out.guardFlags).toBe('draft_generation_failed');
    // the empty draft still exists and must be reachable for deletion
    expect(out.draftIdEnc).toBe('AAMkDRAFT%3D');
  });

  it('falls back to the trigger for identifiers when Voice Guard never ran', () => {
    const out = describeDraftFailure({ item: { error: 'boom' }, request, draft });
    expect(out.messageId).toBe('AAMkMSG');
    expect(out.messageIdEnc).toBe('AAMkMSG%3D');
  });

  it('prefers the id on the item over the one from the prompt node', () => {
    const out = describeDraftFailure({
      item: { error: 'boom', draftIdEnc: 'FROM_ITEM' },
      request,
      draft,
    });
    expect(out.draftIdEnc).toBe('FROM_ITEM');
  });

  it('says plainly when there is no draft to delete', () => {
    const out = describeDraftFailure({ item: { error: 'boom' }, request, draft: {} });
    expect(out.draftIdEnc).toBe('');
  });

  it('carries a short error note without letting it run away', () => {
    const out = describeDraftFailure({
      item: { error: { message: 'x'.repeat(900) } },
      request,
      draft,
    });
    expect(out.note.length).toBeLessThanOrEqual(200);
  });

  it('never throws on a shape it has not seen', () => {
    const out = describeDraftFailure({ item: null, request: null, draft: null });
    expect(out.status).toBe('draft_failed');
    expect(out.draftIdEnc).toBe('');
  });
});

describe('sameAddress — the drafting interlock primitive', () => {
  it('matches regardless of case and padding', () => {
    expect(sameAddress('  YDawson@PartnersPerform.com ', 'ydawson@partnersperform.com')).toBe(true);
  });

  it('refuses a different mailbox', () => {
    expect(sameAddress('ydawson@partnersperform.com', 'gtutor@partnersperform.com')).toBe(false);
  });

  it('fails closed when the mailbox could not be read', () => {
    expect(sameAddress('', 'gtutor@partnersperform.com')).toBe(false);
  });

  it('fails closed when no mailbox is configured', () => {
    // An unset interlock must not silently mean "draft into anything".
    expect(sameAddress('gtutor@partnersperform.com', '')).toBe(false);
    expect(sameAddress('gtutor@partnersperform.com', undefined)).toBe(false);
  });
});

describe('addressedToMailbox — stops drafting if the credential moves', () => {
  const BOX = 'gtutor@partnersperform.com';

  it('accepts a message addressed directly to the mailbox', () => {
    expect(addressedToMailbox(['gtutor@partnersperform.com'], [], BOX)).toBe(true);
  });

  it('accepts a message where the mailbox is only on cc', () => {
    // Yancey is cc rather than to on several of his own OPPORTUNITY examples.
    expect(addressedToMailbox(['kdunlap@partnersperform.com'], [BOX], BOX)).toBe(true);
  });

  it('ignores case and padding', () => {
    expect(addressedToMailbox([' GTutor@PartnersPerform.com '], [], BOX)).toBe(true);
  });

  it('refuses when the mailbox is nowhere on the message', () => {
    // This is the credential-swap case: mail for a different mailbox entirely.
    expect(addressedToMailbox(['ydawson@partnersperform.com'], [], BOX)).toBe(false);
  });

  it('fails closed when no mailbox is configured', () => {
    expect(addressedToMailbox([BOX], [], '')).toBe(false);
  });

  it('fails closed on missing or malformed recipient lists', () => {
    expect(addressedToMailbox(null, null, BOX)).toBe(false);
    expect(addressedToMailbox('not-an-array', undefined, BOX)).toBe(false);
    expect(addressedToMailbox([null, 42, {}], [], BOX)).toBe(false);
  });
});

describe('pricingGuard suppression — only a quoted FEE stands the guard down', () => {
  // Regression. Reading the whole thread for the ask was right; reading it for
  // the suppression was not. Any currency anywhere in 12k chars of quoted
  // history was releasing a live fee request.
  const ASK = 'Can you provide an estimate of cost for extending this into the northern district?';

  it('still fires on a bare fee ask', () => {
    expect(pricingGuard(ASK, 'ILI extension').triggered).toBe(true);
  });

  it('is NOT stood down by unrelated figures quoted below', () => {
    // $40.6B is SCE capital plan wording, verbatim from Yancey's own opp-04.
    const thread =
      ASK +
      '\n\nFrom: Charles Coleman\nSubject: Vector SLT Kickoff\n\n' +
      'SCE plans to deploy approximately $40.6B from 2026 to 2030, with annual ' +
      'capital spend approaching $9B and rate base growing to $67.9B.';
    expect(pricingGuard(thread, 'ILI extension').triggered).toBe(true);
  });

  it('is NOT stood down by a number in a signature block', () => {
    const thread = ASK + '\n\nLou DeFelice | Manager\nOffice 219-851-1800 | 500k customers served';
    expect(pricingGuard(thread, 'ILI extension').triggered).toBe(true);
  });

  it('IS stood down when Yancey already quoted a fee in the thread', () => {
    // demo email #2: the figure sits next to fee-stating language.
    const thread =
      'Can you provide a breakdown of what the fee covers across the five deliverables?' +
      '\n\nFrom: Yancey Dawson\nSubject: Preliminary estimate\n\n' +
      'The indicative range is $175K to $200K on a fixed-fee basis. The estimate is ' +
      'preliminary and non-binding; the final fee would be confirmed at scoping.';
    const guard = pricingGuard(thread, 'RE: Integrating Gas Control into T-Week Planning');
    expect(guard.triggered).toBe(false);
    expect(guard.reason).toContain('already');
  });

  it('is stood down by an hourly rate stated in-thread', () => {
    const thread = ASK + '\n\nAs discussed our blended hourly rate is $265 per hour.';
    expect(pricingGuard(thread, 'Scope').triggered).toBe(false);
  });
});

describe('normalizeClassification owns the interlock flag', () => {
  const raw = {
    label: 'opportunity', confidence: 0.95, reasoning: 'live work',
    requires_draft: true, urgency: 'this_week', asks_of_yancey: ['send workbook'],
  };
  const base = {
    cleanText: 'Can you pull the workbook together?',
    subject: 'RE: Vector CI',
    senderEmail: 'ccoleman@partnersperform.com',
    confidenceMin: 0.6,
  };

  it('drafts when the mailbox is authorised', () => {
    const v = normalizeClassification(raw, { ...base, authorised: true });
    expect(v.requiresDraft).toBe(true);
    expect(v.guardFlags).toBe('');
  });

  it('withholds the draft and flags it, joined like every other guard', () => {
    const v = normalizeClassification(raw, { ...base, authorised: false });
    expect(v.requiresDraft).toBe(false);
    expect(v.guardFlags).toBe('mailbox_not_addressed');
    expect(v.label).toBe('opportunity');
  });

  it('joins the interlock flag with others rather than hand-rolling a comma', () => {
    const v = normalizeClassification(
      { ...raw, confidence: 0.2 },
      { ...base, authorised: false },
    );
    // Low confidence already blocks the draft, so only that flag applies.
    expect(v.guardFlags).toBe('below_confidence_min');
  });

  it('treats an absent mailbox context as no restriction', () => {
    // The golden-set harness scores prompts, not deployments.
    expect(normalizeClassification(raw, base).requiresDraft).toBe(true);
  });
});

describe('voicePreChecks — spec §11 Gate 2 automated pre-checks', () => {
  /**
   * @param {{ checks: { name: string, pass: boolean, detail: string }[] }} result
   * @param {string} name
   */
  const check = (result, name) => {
    const hit = result.checks.find((c) => c.name === name);
    if (!hit) throw new Error('no check named ' + name);
    return hit;
  };

  const SOURCE = 'Can you pull the crossarm workbook together before Thursday? We saved $6.4M last year.';
  const GOOD = [
    'Hey Chuck,',
    '',
    'I will pull the crossarm workbook together and send it before the Momentum Room.',
    '',
    'Thanks,',
    '',
    'Yancey',
  ].join('\n');

  it('passes a draft in voice', () => {
    const r = voicePreChecks(GOOD, SOURCE, {});
    expect(r.pass).toBe(true);
    expect(r.checks.every((c) => c.pass)).toBe(true);
  });

  it('names all four checks the spec lists', () => {
    const names = voicePreChecks(GOOD, SOURCE, {}).checks.map((c) => c.name);
    expect(names).toEqual(['sign_off', 'no_forbidden_opener', 'no_invented_digits', 'under_word_cap']);
  });

  it('fails a missing sign-off', () => {
    const r = voicePreChecks('Hey Chuck,\n\nWill do.', SOURCE, {});
    expect(r.pass).toBe(false);
    expect(check(r, 'sign_off').pass).toBe(false);
  });

  it('fails the forbidden opener', () => {
    const r = voicePreChecks('Hi Chuck,\n\nI hope this email finds you well.\n\nThanks,\n\nYancey', SOURCE, {});
    expect(check(r, 'no_forbidden_opener').pass).toBe(false);
  });

  it('fails a currency figure the source never mentioned', () => {
    const r = voicePreChecks(
      'Hey Chuck,\n\nThat should land around $250K.\n\nThanks,\n\nYancey',
      SOURCE,
      {},
    );
    expect(check(r, 'no_invented_digits').pass).toBe(false);
    expect(check(r, 'no_invented_digits').detail).toContain('$250K');
  });

  it('allows a figure that IS in the source', () => {
    const r = voicePreChecks(
      'Hey Chuck,\n\nThe $6.4M number is the one to use.\n\nThanks,\n\nYancey',
      SOURCE,
      {},
    );
    expect(check(r, 'no_invented_digits').pass).toBe(true);
  });

  it('fails a percentage the source never mentioned', () => {
    const r = voicePreChecks('Hey Chuck,\n\nAbout 68% of them.\n\nThanks,\n\nYancey', SOURCE, {});
    expect(check(r, 'no_invented_digits').pass).toBe(false);
  });

  it('enforces the spec cap of 150 words, not the guard\'s looser 220', () => {
    const long = 'Hey Chuck,\n\n' + 'word '.repeat(160) + '\n\nThanks,\n\nYancey';
    const r = voicePreChecks(long, SOURCE, {});
    expect(check(r, 'under_word_cap').pass).toBe(false);
    expect(check(r, 'under_word_cap').detail).toContain('150');
  });
});

describe('pricingGuard — "how much" must not fire on its own', () => {
  // Found writing a demo: "how much" sat in BOTH the ask list and the question
  // list, so one innocuous phrase satisfied a rule meant to need two different
  // signals. A commercial-property pitch was escalated to flag_for_yancey.
  it('ignores "how much" used about anything but money', () => {
    const body =
      'Given how much Upperform has grown since you took Suite 1170, it might ' +
      'be worth a look before we go to market.';
    expect(pricingGuard(body, 'Private offices two blocks from Main Street').triggered).toBe(false);
  });

  it('still fires when "how much" is asked about money', () => {
    expect(pricingGuard('Roughly how much would the assessment cost?', 'ILI').triggered).toBe(true);
  });

  it('still fires on the plain fee ask', () => {
    expect(
      pricingGuard('Can you provide an estimate of timeline and cost?', 'Scoping').triggered,
    ).toBe(true);
  });
});

describe('sensitivityGuard — word stems that never matched', () => {
  // The trailing \b in the alternation meant "indemnif" could not match
  // "indemnification", which is the only form anyone actually writes.
  it('catches every form of indemnify', () => {
    for (const word of ['indemnification', 'indemnify', 'indemnities', 'indemnified']) {
      expect(sensitivityGuard('a@b.com', 'They added an ' + word + ' obligation.', '', []).triggered)
        .toBe(true);
    }
  });

  it('catches every form of whistleblow', () => {
    for (const word of ['whistleblower', 'whistleblowing', 'whistleblow']) {
      expect(sensitivityGuard('a@b.com', 'A ' + word + ' complaint was raised.', '', []).triggered)
        .toBe(true);
    }
  });

  it('catches the MSA clause language on its own', () => {
    const body =
      'The limitation of liability cap has been reduced, and they have added an ' +
      'indemnification obligation covering third-party claims.';
    expect(sensitivityGuard('a@b.com', body, 'MSA renewal', []).triggered).toBe(true);
  });

  it('does not fire on ordinary delivery talk', () => {
    expect(
      sensitivityGuard('a@b.com', 'The crossarm workbook is attached for review.', 'Deck', []).triggered,
    ).toBe(false);
  });
});

describe('digest never claims shadow mode while drafts are sitting there', () => {
  const shell = {
    resolvedIds: [],
    resolvedMessageIds: [],
    stats: { needsYou: 0, draftsWaiting: 0, followupsDue: 0, resolved: 0, autoClosed: 0, scanned: 0 },
  };
  const buckets = {
    ...shell,
    needsYou: [],
    draftsWaiting: [
      { subject: 'Vector CI working session', sender: 'Charles Coleman', urgency: 'today', webLink: 'https://o/1' },
    ],
    followupsDue: [],
  };
  const empty = { ...shell, needsYou: [], draftsWaiting: [], followupsDue: [] };

  it('shows the shadow banner when nothing has been drafted', () => {
    const html = renderDigestHtml(empty, { mode: 'shadow', mailboxLabel: 'dev', dateLabel: 'Thursday' });
    expect(html).toContain('No replies are being drafted');
  });

  // The live digest printed "No replies are being drafted" directly above a
  // list of six drafts, because the banner came from a hand-set config string
  // rather than from what actually happened.
  it('drops the banner once a draft exists, whatever the config says', () => {
    const html = renderDigestHtml(buckets, { mode: 'shadow', mailboxLabel: 'dev', dateLabel: 'Thursday' });
    expect(html).not.toContain('No replies are being drafted');
    expect(html).toContain('Vector CI working session');
  });

  it('shows no banner in live mode', () => {
    const html = renderDigestHtml(empty, { mode: 'live', mailboxLabel: 'dev', dateLabel: 'Thursday' });
    expect(html).not.toContain('No replies are being drafted');
  });
});

describe('normalizeErrorPayload — the alert has to say what broke', () => {
  // Every error row ever written said "unknown workflow" / "Unknown error",
  // because the Code node read $input (the Config node) instead of the Error
  // Trigger. Two real failures were captured and both were useless.
  const REAL = {
    execution: {
      id: '18',
      url: 'https://n8n/execution/18',
      lastNodeExecuted: 'Generate Draft',
      mode: 'integrated',
      error: { message: 'OpenAI 500: upstream timeout', stack: 'Error: OpenAI 500\n  at draft' },
    },
    workflow: { id: 'Rjpk5KuPopQCpCQL', name: 'WF-B · Draft Reply' },
  };

  it('pulls the workflow, node and message out of a real payload', () => {
    const e = normalizeErrorPayload(REAL, '2026-08-27T10:00:00Z');
    expect(e.workflowName).toBe('WF-B · Draft Reply');
    expect(e.workflowId).toBe('Rjpk5KuPopQCpCQL');
    expect(e.executionId).toBe('18');
    expect(e.nodeName).toBe('Generate Draft');
    expect(e.errorMessage).toBe('OpenAI 500: upstream timeout');
    expect(e.stack).toContain('at draft');
  });

  it('finds a message id anywhere in the payload', () => {
    const withMsg = { ...REAL, execution: { ...REAL.execution, data: { messageId: 'AAMkMSG123456' } } };
    const e = normalizeErrorPayload(withMsg, '2026-08-27T10:00:00Z');
    expect(e.messageId).toBe('AAMkMSG123456');
    expect(e.hasMessage).toBe(true);
    expect(e.messageIdEnc).toBe(encodeURIComponent('AAMkMSG123456'));
  });

  it('says the payload was empty rather than inventing "Unknown error"', () => {
    const e = normalizeErrorPayload({}, '2026-08-27T10:00:00Z');
    // "Unknown error" reads like a real message. This has to read like a fault
    // in the reporting, because that is what it is.
    expect(e.errorMessage).toContain('empty');
    expect(e.workflowName).toContain('unreported');
    expect(e.hasMessage).toBe(false);
  });

  it('never throws on a shape it has not seen', () => {
    expect(() => normalizeErrorPayload(null, '2026-08-27T10:00:00Z')).not.toThrow();
    expect(() => normalizeErrorPayload('nonsense', '2026-08-27T10:00:00Z')).not.toThrow();
  });

  it('truncates a runaway message and stack', () => {
    const huge = { execution: { error: { message: 'x'.repeat(3000), stack: 'y'.repeat(9000) } } };
    const e = normalizeErrorPayload(huge, '2026-08-27T10:00:00Z');
    expect(e.errorMessage.length).toBeLessThanOrEqual(1000);
    expect(e.stack.length).toBeLessThanOrEqual(2000);
  });
});

describe('pricingGuard — the ask and the question have to be in the same breath', () => {
  // Found by Gate 1 disagreeing with Yancey on opp-02. PRICING_QUESTION matched
  // "send me your" in "please send me your updates for the items you are
  // working on", while PRICING_ASK matched "cost" and "estimate" in a different
  // message about idling analysis. Two unrelated phrases, hundreds of
  // characters apart, read together as a fee request.
  //
  // Over a 12,000-character thread a generic question form and a money noun
  // co-occur constantly. They now have to be near each other.

  it('does not fire on a status request that happens to share a thread with cost talk', () => {
    const thread = [
      'Please send me your updates for the items you are working on.',
      '',
      'Separately: total 2025 idling represents an annual cost of $4M, and we may',
      'need a low to medium confidence estimate for the hypothesis target.',
    ].join('\n');
    expect(pricingGuard(thread, 'RE: ACTION NEEDED').triggered).toBe(false);
  });

  it('still fires when the ask and the question are one sentence', () => {
    expect(
      pricingGuard('Is this possible, and can you provide an estimate of timeline and cost?', 'ILI')
        .triggered,
    ).toBe(true);
  });

  it('still fires across a line break within the same request', () => {
    expect(
      pricingGuard('Before we take it to the steering group,\ncan you provide an estimate of cost?', 'Scoping')
        .triggered,
    ).toBe(true);
  });

  it('does not fire on "can you share the deck" beside unrelated budget talk', () => {
    const thread = [
      'Can you share the revised deck before Thursday?',
      '',
      'For context, the O&M gap grows from $11M in 2026 to $48M in 2027 and the',
      'budget position is tighter than last cycle.',
    ].join('\n');
    expect(pricingGuard(thread, 'RE: Vector').triggered).toBe(false);
  });

  it('fires when the fee ask is buried in a quoted chain, as before', () => {
    const thread = [
      "Let's put together what it would take and walk Kris through it next week.",
      '',
      '-----Original Message-----',
      'From: Kris Murray',
      '',
      'Is this possible, and can you provide an estimate of timeline and cost?',
    ].join('\n');
    expect(pricingGuard(thread, 'RE: Northern district ILI extension').triggered).toBe(true);
  });
});

/**
 * Yancey's rules of 2026-08-28: draft only for colleagues, never for Chuck
 * Coleman, and flag his named categories.
 */
describe("Yancey's drafting rules", () => {
  const INTERNAL = ['partnersperform.com'];

  /** @param {Partial<Record<string, unknown>>} over */
  function raw(over) {
    return Object.assign(
      {
        label: 'opportunity',
        confidence: 0.95,
        reasoning: 'Colleague asking for a plan review.',
        requires_draft: true,
        urgency: 'this_week',
        asks_of_yancey: ['review the plan'],
      },
      over || {},
    );
  }

  /** @param {string} senderEmail @param {Partial<Record<string, unknown>>} over */
  function ctx(senderEmail, over) {
    return Object.assign(
      {
        cleanText: 'Can you take a look at the project plan before Thursday?',
        subject: 'Plan review',
        senderEmail,
        confidenceMin: 0.6,
        internalDomains: INTERNAL,
      },
      over || {},
    );
  }

  it('drafts for a colleague', () => {
    const v = normalizeClassification(raw({}), ctx('scorkran@partnersperform.com', {}));
    expect(v.label).toBe('opportunity');
    expect(v.requiresDraft).toBe(true);
    expect(v.guards).toEqual([]);
  });

  it('flags an outsider instead of drafting', () => {
    const v = normalizeClassification(raw({}), ctx('jessica.doucette@sce.com', {}));
    expect(v.label).toBe('flag_for_yancey');
    expect(v.requiresDraft).toBe(false);
    expect(v.guards).toContain('external_sender');
  });

  // The rule is "flag without drafting", not "promote everything external".
  // Ten of the twenty-one fixtures are external noise; routing those into the
  // review bucket would put a newsletter in front of him every morning.
  it('leaves external noise as noise', () => {
    const v = normalizeClassification(
      raw({ label: 'noise', requires_draft: false }),
      ctx('billing@paypal.com', {}),
    );
    expect(v.label).toBe('noise');
    expect(v.requiresDraft).toBe(false);
    expect(v.guards).not.toContain('external_sender');
  });

  it('never drafts for Chuck Coleman even though he is internal', () => {
    const v = normalizeClassification(
      raw({}),
      ctx('ccoleman@partnersperform.com', {
        alwaysFlagSenders: ['ccoleman@partnersperform.com'],
      }),
    );
    expect(v.label).toBe('flag_for_yancey');
    expect(v.requiresDraft).toBe(false);
    expect(v.guardFlags).toContain('always_flag_sender');
  });

  it('keeps the classifier verdict alongside the policy outcome', () => {
    const v = normalizeClassification(raw({}), ctx('jessica.doucette@sce.com', {}));
    expect(v.modelLabel).toBe('opportunity');
    expect(v.label).toBe('flag_for_yancey');
  });

  it('does nothing when no internal domain is configured', () => {
    const v = normalizeClassification(
      raw({}),
      ctx('jessica.doucette@sce.com', { internalDomains: [] }),
    );
    expect(v.label).toBe('opportunity');
    expect(v.requiresDraft).toBe(true);
  });

  // Gate 1 caught this: three of the five emails Yancey labelled JUNK are
  // PayPal invoices titled "Invoice from ...". With the topic vocabulary applied
  // unconditionally, every one of them was promoted into his 07:30 review
  // bucket — a 60% false-positive rate on his own junk mail.
  it('does not resurrect junk that happens to say "invoice"', () => {
    const v = normalizeClassification(
      raw({ label: 'noise', requires_draft: false }),
      ctx('service@paypal.com', {
        subject: 'Invoice from JEREMIAH CHRISTIAN PLACENTE ROSALES (0028)',
        cleanText: 'You sent a payment request for a contractor invoice.',
      }),
    );
    expect(v.label).toBe('noise');
    expect(v.guards).toEqual([]);
  });

  // The counterpart: a genuine legal matter still overrules the classifier,
  // because that guard exists precisely for the case where the model is wrong.
  it('still overrules the classifier on a legal matter', () => {
    const v = normalizeClassification(
      raw({ label: 'noise', requires_draft: false }),
      ctx('someone@example.com', {
        subject: 'Notice of litigation',
        cleanText: 'Our attorney has filed a lawsuit regarding the inspection records.',
      }),
    );
    expect(v.label).toBe('flag_for_yancey');
    expect(v.guardFlags).toContain('sensitivity:legal');
  });

  it('names the category that held a message back', () => {
    const v = normalizeClassification(
      raw({}),
      ctx('scorkran@partnersperform.com', {
        cleanText: 'Can you approve the change order before Thursday?',
      }),
    );
    expect(v.label).toBe('flag_for_yancey');
    expect(v.guardFlags).toContain('sensitivity:approval');
  });

  it('still drafts consulting mail that only sounds like a flagged category', () => {
    const v = normalizeClassification(
      raw({}),
      ctx('dyoes@partnersperform.com', {
        cleanText:
          'Sustain benefits through performance monitoring. Reducing operating expense improves cash flow.',
      }),
    );
    expect(v.label).toBe('opportunity');
    expect(v.requiresDraft).toBe(true);
  });
});

describe('sender domain matching fails closed', () => {
  it('reads the domain half', () => {
    expect(senderDomain('Ccoleman@PartnersPerform.com')).toBe('partnersperform.com');
  });

  it.each([
    ['', 'no address'],
    ['not-an-address', 'no @'],
    ['someone@', 'no domain'],
  ])('treats %s as external (%s)', (address) => {
    expect(isInternalSender(address, ['partnersperform.com'])).toBe(false);
  });

  it('does not accept a subdomain as internal', () => {
    expect(isInternalSender('a@mail.partnersperform.com', ['partnersperform.com'])).toBe(false);
  });

  it('does not accept a lookalike suffix as internal', () => {
    expect(isInternalSender('a@notpartnersperform.com', ['partnersperform.com'])).toBe(false);
  });

  it('matches the configured domain regardless of case or a leading @', () => {
    expect(isInternalSender('YDawson@partnersperform.com', ['@PartnersPerform.com'])).toBe(true);
  });
});
