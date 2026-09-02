import { E, GRAPH, S, SDK_IMPORT, OUTLOOK_CREDENTIAL, SUFFIX, TABLES, TARGET } from './shared.mjs';
import {
  docxFixture,
  pdfFixture,
  pngFixture,
  pptxFixture,
  toBase64,
  xlsxFixture,
} from '../src/office-fixtures.js';

/**
 * Attachments for the fixtures, built at build time from text you can read
 * rather than pasted in as base64.
 *
 * There used to be two extra hand-built seeders for this — one for attachments,
 * one for PDFs — because the generated seeder carried none. Folding them in
 * leaves one seeder that covers every path WF-D takes: Word, PowerPoint and
 * Excel read in process, a PDF that goes to the model for vision, and an image
 * small enough that the triage skips it unread.
 *
 * @param {string} name
 * @param {string} contentType
 * @param {Uint8Array} bytes
 */
const attach = (name, contentType, bytes) => ({
  '@odata.type': '#microsoft.graph.fileAttachment',
  name,
  contentType,
  contentBytes: toBase64(bytes),
});

const OOXML = 'application/vnd.openxmlformats-officedocument';

const FIXTURES = {
  docx: attach(
    'Crossarm-Capitalization_2026-09-01_Draft_v0.1.docx',
    OOXML + '.wordprocessingml.document',
    docxFixture([
      'Crossarm capitalization — field notes',
      'Four hundred and twelve poles surveyed in the northern district.',
      'Recommend treating the replacement programme as capital, not O&M.',
    ]),
  ),
  xlsx: attach(
    'Crossarm-Survey_2026-09-01_Draft_v0.1.xlsx',
    OOXML + '.spreadsheetml.sheet',
    xlsxFixture('Survey', [
      ['District', 'Poles', 'Status'],
      ['Northern', '412', 'In progress'],
      ['Southern', '188', 'Complete'],
    ]),
  ),
  pptx: attach(
    'Momentum-Room_2026-09-01_Draft_v0.1.pptx',
    OOXML + '.presentationml.presentation',
    pptxFixture([
      ['Momentum Room — crossarm programme'],
      ['CAIDI trend, twelve months'],
      ['Next steps and owners'],
    ]),
  ),
  pdf: attach(
    'DIMP-Field-Report_2026-09-01_Draft_v0.1.pdf',
    'application/pdf',
    pdfFixture([
      'DIMP field report - northern district',
      'Pigging complete on eleven of fourteen segments.',
      'Two cross bores identified and cleared.',
      'No further action requested at this time.',
    ]),
  ),
  // Below the reading threshold, so the triage should skip it and say so
  // rather than spend a model call being told what a logo is.
  logo: attach('signature.png', 'image/png', pngFixture()),
};

const BUILD_MESSAGES = `
const FIXTURES = ${JSON.stringify(FIXTURES)};
// Representative fixtures planted straight into the inbox via Graph. We cannot
// *send* mail (no Mail.Send, by design), but Mail.ReadWrite lets us create a
// message directly in a folder, which is enough to exercise the pipeline.
//
// Rewritten 2026-08-28 for Yancey's rules: a reply is written only for a
// colleague at partnersperform.com who is not Chuck Coleman and trips no topic
// guard. Two of these fixtures exist to prove the rules do NOT fire where they
// should not, which is the half that a careless implementation gets wrong.
const messages = [
  {
    label: 'expects: opportunity -> draft written (internal colleague, the path that still works)',
    subject: 'Vector CI working session \\u2014 crossarm workbook',
    fromName: 'Shawn Corkran',
    fromAddress: 'scorkran@partnersperform.com',
    body: [
      'Yancey,',
      '',
      'Can you pull the crossarm workbook together before Thursday\\'s Momentum Room?',
      'I want to walk the SLT through the capitalization approach and it would help',
      'to have your numbers in front of them.',
      '',
      'Shawn'
    ].join('\\n')
  },
  {
    label: 'expects: flag_for_yancey -> NO draft (always_flag_sender \\u2014 Chuck)',
    subject: 'Pulse rollout \\u2014 next steps',
    fromName: 'Charles Coleman',
    fromAddress: 'ccoleman@partnersperform.com',
    body: [
      'Yancey,',
      '',
      'Jessica asked whether we can walk the SLT through the Pulse rollout at',
      'Thursday\\'s session. Can you put the sequencing together and let me know if',
      'that timing works on your end?',
      '',
      'Chuck'
    ].join('\\n')
  },
  {
    label: 'expects: flag_for_yancey -> NO draft (sensitivity:approval, internal sender)',
    subject: 'Change order on the T-Week pilot',
    fromName: 'Darryl Yoes',
    fromAddress: 'dyoes@partnersperform.com',
    body: [
      'Yancey,',
      '',
      'The pilot needs an extra week of field support. Can you approve the change',
      'order so I can get the crews scheduled?',
      '',
      'Darryl'
    ].join('\\n')
  },
  {
    label: 'expects: opportunity -> draft written (regression: consulting words that look like guard terms)',
    subject: 'O&M framing for the portfolio review',
    fromName: 'Darryl Yoes',
    fromAddress: 'dyoes@partnersperform.com',
    body: [
      'Yancey,',
      '',
      'For the portfolio deck: reducing operating expense improves cash flow and',
      'reinforces credit metrics. Directors stay accountable for realizing benefits',
      'through performance monitoring.',
      '',
      'Can you sanity check that framing before I circulate it?',
      '',
      'Darryl'
    ].join('\\n')
  },
  {
    label: 'expects: flag_for_yancey -> NO draft (pricing guard AND external_sender)',
    subject: 'ILI program \\u2014 northern district',
    fromName: 'Kris Murray',
    fromAddress: 'kris.murray@nisource.com',
    body: [
      'Yancey,',
      '',
      'We are looking at extending the pigging and ILI work into the northern',
      'district next year. Is this possible, and can you provide an estimate of',
      'timeline and cost?',
      '',
      'Thanks,',
      'Kris'
    ].join('\\n')
  },
  {
    label: 'expects: noise -> NOT flagged (outside sender, but junk stays junk)',
    subject: 'Your September utility operations briefing is here',
    fromName: 'GridSignal Media',
    fromAddress: 'updates@gridsignal.example.com',
    body: [
      'Hi there,',
      '',
      'This month: predictive dispatch models that actually ship, and why most AMI',
      'data never reaches the field.',
      '',
      'Read the briefing: https://gridsignal.example.com/briefing',
      '',
      'You are receiving this because you subscribed to GridSignal updates.',
      'Unsubscribe | Manage preferences'
    ].join('\\n')
  },
  {
    label: 'expects: opportunity -> draft written; docx, xlsx and pptx all read in process, none leaving the server',
    subject: 'Crossarm pack for Thursday',
    fromName: 'Shawn Corkran',
    fromAddress: 'scorkran@partnersperform.com',
    attachments: [FIXTURES.docx, FIXTURES.xlsx, FIXTURES.pptx],
    body: [
      'Yancey,',
      '',
      'Notes, survey and the deck attached. Anything you want changed before',
      'Thursday?',
      '',
      'Shawn'
    ].join('\\n')
  },
  {
    label: 'expects: opportunity -> draft written; the PDF is read by the model, signature.png is skipped unread',
    subject: 'DIMP field report',
    fromName: 'Darryl Yoes',
    fromAddress: 'dyoes@partnersperform.com',
    attachments: [FIXTURES.pdf, FIXTURES.logo],
    body: [
      'Yancey,',
      '',
      'Field report attached. Nothing needed from you unless something looks',
      'wrong in it.',
      '',
      'Darryl'
    ].join('\\n')
  }
];

return messages.map(m => ({
  json: {
    expectation: m.label,
    payload: {
      subject: m.subject,
      body: { contentType: 'Text', content: m.body },
      attachments: m.attachments || [],
      from: { emailAddress: { name: m.fromName, address: m.fromAddress } },
      sender: { emailAddress: { name: m.fromName, address: m.fromAddress } },
      // Must be the mailbox this target drafts into, or the authorisation
      // interlock withholds every draft and the seeder tests nothing. It was
      // hardcoded to Yancey's address, which made it silently useless in dev.
      toRecipients: [{ emailAddress: { name: 'Inbox owner', address: ${S(TARGET.draftingMailbox)} } }],
      isRead: false,
      // A message POSTed into a folder is created with isDraft: true, and Graph
      // refuses createReply on a draft (ErrorInvalidReferenceItem). Writing
      // PR_MESSAGE_FLAGS clears MSGFLAG_UNSENT so the fixture behaves like real
      // received mail. Without this the fixture is untestable, not the workflow.
      singleValueExtendedProperties: [{ id: 'Integer 0x0E07', value: '1' }]
    }
  }
}));
`;

/**
 * Never buildable against production.
 *
 * The seeder plants fake mail directly into whichever mailbox the target names.
 * That is exactly right in dev and indefensible in Yancey's inbox, and the only
 * thing that used to prevent it was a sentence on a sticky note. `npm run build`
 * defaults to prod, so without this a routine rebuild produces a workflow whose
 * whole purpose is writing invented emails into a live client mailbox.
 */
if (TARGET.name === 'prod') {
  throw new Error(
    'refusing to build the seeder for production. It plants fabricated mail in ' +
      'the target mailbox; use TARGET=dev.',
  );
}

export const NAME = 'WF-S · Seed Test Inbox' + SUFFIX;

export async function build() {
  return `${SDK_IMPORT}

const runSeeder = trigger({
  type: 'n8n-nodes-base.manualTrigger',
  version: 1,
  config: { name: 'Seed Test Emails', parameters: {}, position: [0, 0] },
  output: [{}]
});

const config = node({
  type: 'n8n-nodes-base.set',
  version: 3.5,
  config: {
    name: 'Config',
    parameters: { mode: 'raw', jsonOutput: ${S(JSON.stringify({ graphBase: GRAPH }, null, 2))} },
    position: [200, 0]
  },
  output: [{ graphBase: ${S(GRAPH)} }]
});

const buildMessages = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Build Test Messages',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(BUILD_MESSAGES)} },
    position: [400, 0]
  },
  output: [{ expectation: 'expects: opportunity -> draft written (internal colleague)', payload: {} }]
});

const plantInInbox = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Plant in Inbox',
    // Graph throttles concurrent writes to a single mailbox: posting six
    // fixtures back to back returns 429 ApplicationThrottled
    // ("Application is over its MailboxConcurrency limit") after the first.
    // One at a time with a pause between, and a retry for the residual case.
    retryOnFail: true,
    maxTries: 3,
    waitBetweenTries: 5000,
    parameters: {
      method: 'POST',
      url: ${E("{{ $('Config').first().json.graphBase }}/me/mailFolders/inbox/messages")},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftOutlookOAuth2Api',
      sendBody: true,
      contentType: 'json',
      specifyBody: 'json',
      jsonBody: ${E('{{ JSON.stringify($json.payload) }}')},
      options: {
        timeout: 20000,
        batching: { batch: { batchSize: 1, batchInterval: 2000 } },
        response: { response: { responseFormat: 'json' } }
      }
    },
    credentials: { microsoftOutlookOAuth2Api: newCredential(${S(OUTLOOK_CREDENTIAL)}) },
    position: [600, 0]
  },
  output: [{ id: 'AAMkSEED', subject: 'Vector CI working session — crossarm workbook' }]
});

const summarise = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Summary',
    parameters: {
      mode: 'runOnceForAllItems',
      language: 'javaScript',
      jsCode: "const planted = $input.all().map((i, n) => ({\\n  subject: i.json.subject,\\n  from: ((i.json.from || {}).emailAddress || {}).address || '(mailbox owner)',\\n  expectation: ($('Build Test Messages').all()[n] || { json: {} }).json.expectation,\\n  categories: i.json.categories || []\\n}));\\nreturn [{ json: { planted, count: planted.length } }];"
    },
    position: [800, 0]
  },
  output: [{ planted: [], count: 2 }]
});

const note = sticky(
  ${S('## WF-S · Seed Test Inbox\n\nPlants representative emails straight into the dev inbox so the full pipeline can be exercised on demand.\n\nThe agent has no `Mail.Send`, so it cannot email itself a fixture — but `Mail.ReadWrite` allows creating a message directly inside a folder, which is enough. That is why this works at all without weakening the invariant.\n\nEach fixture carries an `expectation` string. Run this, then run WF-A, then check the result against those expectations:\n\n- **crossarm workbook** -> `opportunity`, draft written, `AI-Drafted`\n- **ILI program** -> `flag_for_yancey`, **no draft**, `AI-Review`\n\nThe second one is the one that matters: it proves the pricing guard still holds when drafting is switched **on**.\n\nDev mailbox only. Never run this against Yancey\'s mailbox.')},
  [runSeeder, config, buildMessages, plantInInbox, summarise],
  { color: 6, width: 1000, height: 300 },
);

export default workflow('upperform-wf-s-seed', ${S(NAME)})
  .add(runSeeder)
  .to(config)
  .to(buildMessages)
  .to(plantInInbox)
  .to(summarise)
  .add(note);
`;
}
