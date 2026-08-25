/**
 * Test fixtures. These are synthetic stand-ins shaped like the real traffic
 * described in the business profile — they are NOT Yancey's training emails.
 * Drop the real 20 into fixtures/golden/ for the Gate 1 harness.
 */

/** A deep Outlook forward chain: two lines of new text on top of three quoted messages. */
export const DEEP_FORWARD_CHAIN = [
  'Yancey,',
  '',
  'Please see below and let me know if you have questions on the crossarm numbers.',
  '',
  'Chuck',
  '',
  '________________________________',
  'From: Jessica Doucette <jessica.doucette@sce.com>',
  'Sent: Monday, August 18, 2026 9:14 AM',
  'To: Charles Coleman <ccoleman@partnersperform.com>',
  'Subject: RE: Vector CI working session — crossarm capitalization',
  '',
  'Chuck, the SLT deck looks right to me. Dean had one comment on slide 12.',
  '',
  '-----Original Message-----',
  'From: Charles Coleman <ccoleman@partnersperform.com>',
  'Sent: Friday, August 15, 2026 4:02 PM',
  'To: Jessica Doucette <jessica.doucette@sce.com>',
  'Subject: Vector CI working session — crossarm capitalization',
  '',
  'Jessica, attaching the workbook ahead of the Momentum Room.',
  '',
  'On Thu, Aug 14, 2026 at 11:20 AM, Dean Heiss wrote:',
  '> We reviewed 307 reports and retired 188 of them.',
  '> That is a 68% reduction against the baseline.',
].join('\n');

/** A forward with almost no new text on top — the meaning lives in the quote. */
export const BARE_FORWARD = [
  'FYI',
  '',
  '________________________________',
  'From: Lou DeFelice <lou.defelice@nisource.com>',
  'Sent: Tuesday, August 19, 2026 7:45 AM',
  'To: Yancey Dawson <YDawson@partnersperform.com>',
  'Subject: NIPSCO power delivery — ETR and CAIDI sprint',
  '',
  'Yancey, can you join the leader discipline session next week? We want to walk the Pulse dashboard.',
].join('\n');

export const OUTLOOK_HTML_BODY = [
  '<html><head><style>p{margin:0}</style></head><body>',
  '<div class="WordSection1"><p class="MsoNormal">Hi Yancey,<o:p></o:p></p>',
  '<p class="MsoNormal">&nbsp;</p>',
  '<p class="MsoNormal">Is this possible, and can you provide an estimate of timeline and cost?<o:p></o:p></p>',
  '<p class="MsoNormal">&nbsp;</p>',
  '<p class="MsoNormal">Thanks,<br>Kris Murray</p></div>',
  '</body></html>',
].join('');

export const NO_QUOTE_BODY = [
  'Hey Yancey,',
  '',
  'Received! Thank you!',
  '',
  'Ashley',
].join('\n');

/**
 * @param {Partial<Record<string, any>>} [overrides]
 */
export function graphMessage(overrides) {
  return Object.assign(
    {
      id: 'AAMkAGI2_msg_001',
      conversationId: 'AAQkAGI2_conv_001',
      internetMessageId: '<abc@partnersperform.com>',
      receivedDateTime: '2026-08-20T14:03:00Z',
      subject: 'RE: Vector CI working session',
      from: { emailAddress: { name: 'Charles Coleman', address: 'CCOLEMAN@partnersperform.com' } },
      toRecipients: [{ emailAddress: { name: 'Yancey Dawson', address: 'YDawson@partnersperform.com' } }],
      ccRecipients: [],
      hasAttachments: true,
      webLink: 'https://outlook.office.com/mail/id/AAMkAGI2_msg_001',
      categories: ['Blue category'],
      body: { contentType: 'text', content: DEEP_FORWARD_CHAIN },
    },
    overrides || {},
  );
}
