import { DEPLOY, E, GRAPH, S, SDK_IMPORT } from './shared.mjs';

const CONFIG = {
  alertFrom: 'agent@partnersperform.com',
  alertTo: 'gabdotcode01@gmail.com',
  graphBase: GRAPH,
};

const NORMALIZE_ERROR = `
const items = $input.all();
const out = [];
for (const item of items) {
  const j = item.json || {};
  const ex = j.execution || {};
  const wf = j.workflow || {};
  const err = ex.error || j.error || {};

  let messageId = '';
  const probe = JSON.stringify(j).match(/"message(?:Id|_id)"\\s*:\\s*"([^"]{6,})"/);
  if (probe && probe[1]) messageId = probe[1];

  out.push({
    json: {
      occurredAt: new Date().toISOString(),
      workflowName: String(wf.name || 'unknown workflow'),
      workflowId: String(wf.id || ''),
      executionId: String(ex.id || ''),
      executionUrl: String(ex.url || ''),
      nodeName: String(ex.lastNodeExecuted || err.node || ''),
      mode: String(ex.mode || ''),
      messageId,
      messageIdEnc: messageId ? encodeURIComponent(messageId) : '',
      hasMessage: messageId.length > 0,
      errorMessage: String(err.message || 'Unknown error').slice(0, 1000),
      stack: String(err.stack || '').slice(0, 2000)
    }
  });
}
return out;
`;

export const NAME = 'WF-E · Error Handler';

export async function build() {
  return `${SDK_IMPORT}

const onWorkflowError = trigger({
  type: 'n8n-nodes-base.errorTrigger',
  version: 1,
  config: { name: 'On Workflow Error', parameters: {}, position: [0, 0] },
  output: [{ execution: { id: '1042', lastNodeExecuted: 'Fetch Inbox', error: { message: 'Graph 429' } }, workflow: { id: 'abc', name: 'WF-A · Poll & Classify' } }]
});

const config = node({
  type: 'n8n-nodes-base.set',
  version: 3.5,
  config: {
    name: 'Config',
    parameters: { mode: 'raw', jsonOutput: ${S(JSON.stringify(CONFIG, null, 2))} },
    position: [200, 0]
  },
  output: ${JSON.stringify([CONFIG])}
});

const normalizeError = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Normalize Error',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(NORMALIZE_ERROR)} },
    position: [400, 0]
  },
  output: [{ occurredAt: '2026-08-26T12:00:00.000Z', workflowName: 'WF-A · Poll & Classify', workflowId: 'abc', executionId: '1042', executionUrl: 'https://n8n.example/execution/1042', nodeName: 'Fetch Inbox', mode: 'trigger', messageId: '', messageIdEnc: '', hasMessage: false, errorMessage: 'Graph 429', stack: '' }]
});

const logError = node({
  type: 'n8n-nodes-base.dataTable',
  version: 1.1,
  config: {
    name: 'Log Error',
    onError: 'continueRegularOutput',
    parameters: {
      resource: 'row',
      operation: 'insert',
      dataTableId: { __rl: true, mode: 'id', value: ${S(DEPLOY.dataTables.errorLog)}, cachedResultName: 'agent_error_log' },
      columns: {
        mappingMode: 'defineBelow',
        value: {
          occurred_at: ${E('{{ $json.occurredAt }}')},
          workflow_name: ${E('{{ $json.workflowName }}')},
          workflow_id: ${E('{{ $json.workflowId }}')},
          execution_id: ${E('{{ $json.executionId }}')},
          node_name: ${E('{{ $json.nodeName }}')},
          message_id: ${E('{{ $json.messageId }}')},
          error_message: ${E('{{ $json.errorMessage }}')},
          stack: ${E('{{ $json.stack }}')}
        },
        schema: [
          { id: 'occurred_at', displayName: 'occurred_at', required: false, defaultMatch: false, display: true, type: 'dateTime', canBeUsedToMatch: true },
          { id: 'workflow_name', displayName: 'workflow_name', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'workflow_id', displayName: 'workflow_id', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'execution_id', displayName: 'execution_id', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'node_name', displayName: 'node_name', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'message_id', displayName: 'message_id', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'error_message', displayName: 'error_message', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'stack', displayName: 'stack', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true }
        ]
      }
    },
    position: [640, -160]
  },
  output: [{ id: 1, createdAt: '2026-08-26T12:00:00.000Z', updatedAt: '2026-08-26T12:00:00.000Z' }]
});

const alertBuilder = node({
  type: 'n8n-nodes-base.emailSend',
  version: 2.1,
  config: {
    name: 'Alert the Builder',
    onError: 'continueRegularOutput',
    parameters: {
      operation: 'send',
      fromEmail: ${E('{{ $(\'Config\').first().json.alertFrom }}')},
      toEmail: ${E('{{ $(\'Config\').first().json.alertTo }}')},
      subject: ${E('[Upperform agent] {{ $json.workflowName }} failed at {{ $json.nodeName }}')},
      emailFormat: 'text',
      text: ${E('Workflow: {{ $json.workflowName }}\nNode: {{ $json.nodeName }}\nExecution: {{ $json.executionId }}\nWhen: {{ $json.occurredAt }}\nMessage id: {{ $json.messageId }}\n\n{{ $json.errorMessage }}\n\n{{ $json.executionUrl }}')},
      options: { appendAttribution: false }
    },
    credentials: { smtp: newCredential('SMTP – Digest Sender') },
    position: [640, 0]
  },
  output: [{ accepted: ['gabdotcode01@gmail.com'] }]
});

const hasMessageId = ifElse({
  version: 2.3,
  config: {
    name: 'Error Tied to a Message?',
    parameters: {
      conditions: {
        options: { caseSensitive: true, leftValue: '', typeValidation: 'loose', version: 2 },
        conditions: [{ id: 'has-message', leftValue: ${E('{{ $json.hasMessage }}')}, operator: { type: 'boolean', operation: 'true', singleValue: true } }],
        combinator: 'and'
      },
      looseTypeValidation: true
    },
    position: [640, 180]
  },
  output: [{ hasMessage: true, messageIdEnc: 'AAMk' }]
});

const readCategories = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Read Message Categories',
    onError: 'continueRegularOutput',
    parameters: {
      method: 'GET',
      url: ${E('{{ $(\'Config\').first().json.graphBase }}/me/messages/{{ $json.messageIdEnc }}?$select=categories')},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftOutlookOAuth2Api',
      options: { timeout: 20000, response: { response: { responseFormat: 'json', neverError: true } } }
    },
    credentials: { microsoftOutlookOAuth2Api: newCredential('Outlook – Dev (gtutor)') },
    position: [880, 180]
  },
  output: [{ categories: ['AI-Processed', 'AI-Action'] }]
});

const tagAiError = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Tag Message AI-Error',
    onError: 'continueRegularOutput',
    parameters: {
      method: 'PATCH',
      url: ${E('{{ $(\'Config\').first().json.graphBase }}/me/messages/{{ $(\'Normalize Error\').first().json.messageIdEnc }}')},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftOutlookOAuth2Api',
      sendBody: true,
      contentType: 'json',
      specifyBody: 'json',
      jsonBody: ${E('{{ JSON.stringify({ categories: ($json.categories || []).filter(c => c !== "AI-Error").concat(["AI-Error"]) }) }}')},
      options: { timeout: 20000 }
    },
    credentials: { microsoftOutlookOAuth2Api: newCredential('Outlook – Dev (gtutor)') },
    position: [1120, 180]
  },
  output: [{ id: 'AAMk', categories: ['AI-Processed', 'AI-Action', 'AI-Error'] }]
});

const noteHeader = sticky(
  ${S('## WF-E · Error Handler\n\nSet as the Error Workflow on WF-A, WF-B and WF-C.\n\nEvery failure lands here: it is written to `agent_error_log`, emailed to the builder, and — when the failure can be traced to a specific message — that message is tagged `AI-Error` in Outlook so nothing silently disappears from the inbox.\n\nEdit the alert addresses in **Config**.')},
  [onWorkflowError, config, normalizeError],
  { color: 3, width: 620, height: 260 }
);

const noteBranches = sticky(
  ${S('### Three independent branches\n\nEach side effect has `onError: continueRegularOutput` so a failure in the error handler itself can never cascade. Losing the Outlook tag must not cost you the log row or the alert email.')},
  [logError, alertBuilder, hasMessageId, readCategories, tagAiError],
  { color: 5, width: 720, height: 420 }
);

export default workflow('upperform-wf-e-errors', ${S(NAME)})
  .add(onWorkflowError)
  .to(config)
  .to(normalizeError)
  .to(logError)
  .add(normalizeError)
  .to(alertBuilder)
  .add(normalizeError)
  .to(hasMessageId.onTrue(readCategories.to(tagAiError)))
  .add(noteHeader)
  .add(noteBranches);
`;
}
