import { E, GRAPH, S, SDK_IMPORT } from './shared.mjs';

const BUILD_MESSAGES = `
// Representative fixtures planted straight into the inbox via Graph. We cannot
// *send* mail (no Mail.Send, by design), but Mail.ReadWrite lets us create a
// message directly in a folder, which is enough to exercise the pipeline.
const messages = [
  {
    label: 'expects: opportunity -> draft written',
    subject: 'Vector CI working session — crossarm workbook',
    fromName: 'Charles Coleman',
    fromAddress: 'ccoleman@partnersperform.com',
    body: [
      'Yancey,',
      '',
      'Jessica asked whether we can walk the SLT through the crossarm capitalization',
      'approach at Thursday\\'s Momentum Room. Can you pull the workbook together and',
      'let me know if that timing works on your end?',
      '',
      'Chuck'
    ].join('\\n')
  },
  {
    label: 'expects: flag_for_yancey -> NO draft (pricing guard)',
    subject: 'ILI program — northern district',
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
  }
];

return messages.map(m => ({
  json: {
    expectation: m.label,
    payload: {
      subject: m.subject,
      body: { contentType: 'Text', content: m.body },
      from: { emailAddress: { name: m.fromName, address: m.fromAddress } },
      sender: { emailAddress: { name: m.fromName, address: m.fromAddress } },
      toRecipients: [{ emailAddress: { name: 'Yancey Dawson', address: 'YDawson@partnersperform.com' } }],
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

export const NAME = 'WF-S · Seed Test Inbox';

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
  output: [{ expectation: 'expects: opportunity -> draft written', payload: {} }]
});

const plantInInbox = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Plant in Inbox',
    parameters: {
      method: 'POST',
      url: ${E("{{ $('Config').first().json.graphBase }}/me/mailFolders/inbox/messages")},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftOutlookOAuth2Api',
      sendBody: true,
      contentType: 'json',
      specifyBody: 'json',
      jsonBody: ${E('{{ JSON.stringify($json.payload) }}')},
      options: { timeout: 20000, response: { response: { responseFormat: 'json' } } }
    },
    credentials: { microsoftOutlookOAuth2Api: newCredential('Outlook – Dev (gtutor)') },
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
