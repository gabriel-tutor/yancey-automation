import { codeNode } from '../scripts/bundle.mjs';
import { CLASSIFY_SCHEMA } from '../src/prompts.js';
import { HELPERS } from './helpers-manifest.mjs';
import { DEPLOY, E, GRAPH, S, SDK_IMPORT, contextLoaderNodes } from './shared.mjs';

const CONFIG = {
  modelClassify: 'gpt-5.6-luna',
  temperatureClassify: 0,
  confidenceMin: 0.6,
  followupDays: 3,
  pollBatch: 25,
  lookbackHours: 6,
  maxBodyChars: 6000,
  draftingEnabled: false,
  alwaysFlagSenders: [],
  graphBase: GRAPH,
};

const GRAPH_SELECT =
  'id,conversationId,internetMessageId,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,body,hasAttachments,categories,webLink';

const RESOLVE_CONTEXT = `
const cfg = $('Config').first().json;
const raw = $input.first().json || {};
const fromFile = typeof raw.profileText === 'string' ? raw.profileText.trim() : '';
const usable = fromFile.length > 200;
const profile = usable ? fromFile : FALLBACK_PROFILE;

const lookbackHours = Number(cfg.lookbackHours) || 6;
const since = new Date(Date.now() - lookbackHours * 3600 * 1000)
  .toISOString()
  .replace(/\\.\\d{3}Z$/, 'Z');

return [{
  json: {
    profile,
    profileSource: usable ? 'file' : 'fallback',
    classifySystem: buildClassifySystemPrompt(profile),
    graphFilter: 'receivedDateTime ge ' + since,
    since
  }
}];
`;

const PREPARE_EMAIL = `
const cfg = $('Config').first().json;
const maxChars = Number(cfg.maxBodyChars) || 6000;
const out = [];

for (const item of $input.all()) {
  const n = normalizeGraphMessage(item.json, { maxChars });
  if (!n.messageId) continue;
  n.messageIdEnc = encodeURIComponent(n.messageId);
  n.classifierInput = buildClassifierInput(n);
  out.push({ json: n });
}

return out;
`;

const APPLY_CLASSIFICATION = `
const cfg = $('Config').first().json;

// Pair this classification back to the message it came from. $().item is the
// supported route; the index fallback keeps the run alive if any upstream node
// ever drops pairedItem, which would otherwise break every message at once.
let src;
try {
  src = $('Prepare Email').item.json;
} catch (e) {
  const prepared = $('Prepare Email').all();
  const fallback = prepared[$itemIndex] || prepared[0];
  src = fallback ? fallback.json : {};
}
if (!src || !src.messageId) throw new Error('Apply Classification could not pair with Prepare Email');

const nowIso = new Date().toISOString();

const c = normalizeClassification($json, {
  cleanText: String(src.cleanText || ''),
  subject: String(src.subject || ''),
  senderEmail: String(src.senderEmail || ''),
  confidenceMin: Number(cfg.confidenceMin) || 0.6,
  alwaysFlagSenders: Array.isArray(cfg.alwaysFlagSenders) ? cfg.alwaysFlagSenders : []
});

const receivedAt = src.receivedAt || nowIso;

return {
  json: {
    messageId: String(src.messageId || ''),
    messageIdEnc: String(src.messageIdEnc || ''),
    conversationId: String(src.conversationId || ''),
    internetMessageId: String(src.internetMessageId || ''),
    receivedAt,
    senderEmail: String(src.senderEmail || ''),
    senderName: String(src.senderName || ''),
    subject: String(src.subject || ''),
    webLink: String(src.webLink || ''),
    hasAttachments: Boolean(src.hasAttachments),
    cleanText: String(src.cleanText || ''),
    label: c.label,
    confidence: c.confidence,
    urgency: c.urgency,
    reasoning: c.reasoning,
    asks: JSON.stringify(c.asks),
    requiresDraft: c.requiresDraft,
    guardFlags: c.guardFlags,
    status: c.label === 'noise' ? 'resolved' : 'classified',
    followupDueAt: addBusinessDays(new Date(receivedAt), Number(cfg.followupDays) || 3).toISOString(),
    categories: decideCategories(src.categories, c.label),
    profileSource: $('Resolve Context').first().json.profileSource
  }
};
`;

export const NAME = 'WF-A · Poll & Classify';

export async function build() {
  const wfB = DEPLOY.workflows.wfB;
  if (!wfB) throw new Error('config/deploy.json: workflows.wfB must be set before building WF-A');

  const resolveContextCode = await codeNode(
    HELPERS['wf-a/Resolve Context'],
    RESOLVE_CONTEXT,
  );
  const prepareEmailCode = await codeNode(
    HELPERS['wf-a/Prepare Email'],
    PREPARE_EMAIL,
  );
  const applyClassificationCode = await codeNode(
    HELPERS['wf-a/Apply Classification'],
    APPLY_CLASSIFICATION,
  );

  return `${SDK_IMPORT}

const everyTwoMinutes = trigger({
  type: 'n8n-nodes-base.scheduleTrigger',
  version: 1.3,
  config: {
    name: 'Every 2 Minutes',
    parameters: { rule: { interval: [{ field: 'minutes', minutesInterval: 2 }] } },
    position: [0, 0]
  },
  output: [{}]
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
${contextLoaderNodes({ x: 400, y: 0 })}
const resolveContext = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Resolve Context',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(resolveContextCode)} },
    position: [800, 0]
  },
  output: [{ profile: '# Upperform Partners — Agent Context File', profileSource: 'file', classifySystem: 'You triage the inbox of a utilities management consultant...', graphFilter: 'receivedDateTime ge 2026-08-26T06:00:00Z', since: '2026-08-26T06:00:00Z' }]
});

const fetchInbox = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Fetch Inbox',
    retryOnFail: true,
    maxTries: 3,
    waitBetweenTries: 5000,
    parameters: {
      method: 'GET',
      url: ${E("{{ $('Config').first().json.graphBase }}/me/mailFolders/inbox/messages")},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftOutlookOAuth2Api',
      sendQuery: true,
      specifyQuery: 'keypair',
      queryParameters: {
        parameters: [
          { name: '$select', value: ${S(GRAPH_SELECT)} },
          { name: '$filter', value: ${E("{{ $('Resolve Context').first().json.graphFilter }}")} },
          { name: '$orderby', value: 'receivedDateTime desc' },
          { name: '$top', value: ${E("{{ $('Config').first().json.pollBatch }}")} }
        ]
      },
      sendHeaders: true,
      specifyHeaders: 'keypair',
      headerParameters: {
        parameters: [{ name: 'Prefer', value: 'outlook.body-content-type="text"' }]
      },
      options: { timeout: 30000, response: { response: { responseFormat: 'json' } } }
    },
    credentials: { microsoftOutlookOAuth2Api: newCredential('Outlook – Dev (gtutor)') },
    position: [1000, 0]
  },
  output: [{ value: [{ id: 'AAMkAGI2', subject: 'RE: Vector CI working session', categories: [], receivedDateTime: '2026-08-26T11:58:00Z' }] }]
});

const splitMessages = node({
  type: 'n8n-nodes-base.splitOut',
  version: 1,
  config: {
    name: 'Split Out Messages',
    parameters: { fieldToSplitOut: 'value', include: 'noOtherFields' },
    position: [1200, 0]
  },
  output: [{ id: 'AAMkAGI2', subject: 'RE: Vector CI working session', categories: [] }]
});

const skipProcessed = node({
  type: 'n8n-nodes-base.filter',
  version: 2.3,
  config: {
    name: 'Skip Already Labelled',
    parameters: {
      conditions: {
        options: { caseSensitive: true, leftValue: '', typeValidation: 'loose', version: 2 },
        conditions: [{ id: 'not-processed', leftValue: ${E("{{ ($json.categories || []).includes('AI-Processed') }}")}, operator: { type: 'boolean', operation: 'false', singleValue: true } }],
        combinator: 'and'
      },
      looseTypeValidation: true
    },
    position: [1400, 0]
  },
  output: [{ id: 'AAMkAGI2', categories: [] }]
});

const skipLogged = node({
  type: 'n8n-nodes-base.dataTable',
  version: 1.1,
  config: {
    name: 'Skip Already Logged',
    parameters: {
      resource: 'row',
      operation: 'rowNotExists',
      dataTableId: { __rl: true, mode: 'id', value: ${S(DEPLOY.dataTables.emailLog)}, cachedResultName: 'agent_email_log' },
      matchType: 'allConditions',
      filters: { conditions: [{ keyName: 'message_id', condition: 'eq', keyValue: ${E('{{ $json.id }}')} }] }
    },
    position: [1600, 0]
  },
  output: [{ id: 'AAMkAGI2', categories: [] }]
});

const prepareEmail = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Prepare Email',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(prepareEmailCode)} },
    position: [1800, 0]
  },
  output: [{ messageId: 'AAMkAGI2', messageIdEnc: 'AAMkAGI2', conversationId: 'AAQk', subject: 'RE: Vector CI working session', senderEmail: 'ccoleman@partnersperform.com', senderName: 'Charles Coleman', receivedAt: '2026-08-26T11:58:00Z', hasAttachments: true, categories: [], cleanText: 'Please send the crossarm workbook before the Momentum Room.', classifierInput: 'SUBJECT: RE: Vector CI working session...' }]
});

const classifyEmail = node({
  type: '@n8n/n8n-nodes-langchain.openAi',
  version: 2.3,
  config: {
    name: 'Classify Email',
    retryOnFail: true,
    maxTries: 3,
    waitBetweenTries: 2000,
    parameters: {
      resource: 'text',
      operation: 'response',
      modelId: { __rl: true, mode: 'id', value: ${E("{{ $('Config').first().json.modelClassify }}")} },
      responses: { values: [{ type: 'text', role: 'user', content: ${E('{{ $json.classifierInput }}')} }] },
      simplify: true,
      options: {
        instructions: ${E("{{ $('Resolve Context').first().json.classifySystem }}")},
        temperature: ${E("{{ $('Config').first().json.temperatureClassify }}")},
        store: false,
        promptCacheKey: 'upperform-classify-v1',
        textFormat: {
          textOptions: {
            type: 'json_schema',
            name: 'upperform_email_triage',
            strict: true,
            description: 'Triage verdict for one inbound email.',
            schema: ${S(JSON.stringify(CLASSIFY_SCHEMA))}
          }
        }
      }
    },
    credentials: { openAiApi: newCredential('OpenAI – Upperform') },
    position: [2000, 0]
  },
  output: [{ label: 'opportunity', confidence: 0.88, reasoning: 'live engagement work on an active deliverable', requires_draft: true, urgency: 'this_week', asks_of_yancey: ['send the crossarm workbook'] }]
});

const applyClassification = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Apply Classification',
    parameters: { mode: 'runOnceForEachItem', language: 'javaScript', jsCode: ${S(applyClassificationCode)} },
    position: [2200, 0]
  },
  output: [{ messageId: 'AAMkAGI2', messageIdEnc: 'AAMkAGI2', conversationId: 'AAQk', internetMessageId: '<abc@partnersperform.com>', receivedAt: '2026-08-26T11:58:00Z', senderEmail: 'ccoleman@partnersperform.com', senderName: 'Charles Coleman', subject: 'RE: Vector CI working session', webLink: 'https://outlook.office.com/mail/id/AAMkAGI2', hasAttachments: true, cleanText: 'Please send the crossarm workbook.', label: 'opportunity', confidence: 0.88, urgency: 'this_week', reasoning: 'live engagement work', asks: '["send the crossarm workbook"]', requiresDraft: true, guardFlags: '', status: 'classified', followupDueAt: '2026-08-31T11:58:00.000Z', categories: ['AI-Processed', 'AI-Action'], profileSource: 'file' }]
});

const tagInOutlook = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Tag in Outlook',
    retryOnFail: true,
    maxTries: 3,
    waitBetweenTries: 5000,
    parameters: {
      method: 'PATCH',
      url: ${E("{{ $('Config').first().json.graphBase }}/me/messages/{{ $json.messageIdEnc }}")},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftOutlookOAuth2Api',
      sendBody: true,
      contentType: 'json',
      specifyBody: 'json',
      jsonBody: ${E('{{ JSON.stringify({ categories: $json.categories }) }}')},
      options: { timeout: 20000 }
    },
    credentials: { microsoftOutlookOAuth2Api: newCredential('Outlook – Dev (gtutor)') },
    position: [2400, 0]
  },
  output: [{ id: 'AAMkAGI2', categories: ['AI-Processed', 'AI-Action'] }]
});

const logClassification = node({
  type: 'n8n-nodes-base.dataTable',
  version: 1.1,
  config: {
    name: 'Log Classification',
    parameters: {
      resource: 'row',
      operation: 'insert',
      dataTableId: { __rl: true, mode: 'id', value: ${S(DEPLOY.dataTables.emailLog)}, cachedResultName: 'agent_email_log' },
      columns: {
        mappingMode: 'defineBelow',
        value: {
          message_id: ${E("{{ $('Apply Classification').item.json.messageId }}")},
          conversation_id: ${E("{{ $('Apply Classification').item.json.conversationId }}")},
          internet_message_id: ${E("{{ $('Apply Classification').item.json.internetMessageId }}")},
          received_at: ${E("{{ $('Apply Classification').item.json.receivedAt }}")},
          sender: ${E("{{ $('Apply Classification').item.json.senderEmail }}")},
          sender_name: ${E("{{ $('Apply Classification').item.json.senderName }}")},
          subject: ${E("{{ $('Apply Classification').item.json.subject }}")},
          label: ${E("{{ $('Apply Classification').item.json.label }}")},
          confidence: ${E("{{ $('Apply Classification').item.json.confidence }}")},
          urgency: ${E("{{ $('Apply Classification').item.json.urgency }}")},
          reasoning: ${E("{{ $('Apply Classification').item.json.reasoning }}")},
          asks: ${E("{{ $('Apply Classification').item.json.asks }}")},
          requires_draft: ${E("{{ $('Apply Classification').item.json.requiresDraft }}")},
          has_attachments: ${E("{{ $('Apply Classification').item.json.hasAttachments }}")},
          web_link: ${E("{{ $('Apply Classification').item.json.webLink }}")},
          status: ${E("{{ $('Apply Classification').item.json.status }}")},
          guard_flags: ${E("{{ $('Apply Classification').item.json.guardFlags }}")},
          followup_due_at: ${E("{{ $('Apply Classification').item.json.followupDueAt }}")}
        },
        schema: [
          { id: 'message_id', displayName: 'message_id', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'conversation_id', displayName: 'conversation_id', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'internet_message_id', displayName: 'internet_message_id', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'received_at', displayName: 'received_at', required: false, defaultMatch: false, display: true, type: 'dateTime', canBeUsedToMatch: true },
          { id: 'sender', displayName: 'sender', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'sender_name', displayName: 'sender_name', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'subject', displayName: 'subject', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'label', displayName: 'label', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'confidence', displayName: 'confidence', required: false, defaultMatch: false, display: true, type: 'number', canBeUsedToMatch: true },
          { id: 'urgency', displayName: 'urgency', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'reasoning', displayName: 'reasoning', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'asks', displayName: 'asks', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'requires_draft', displayName: 'requires_draft', required: false, defaultMatch: false, display: true, type: 'boolean', canBeUsedToMatch: true },
          { id: 'has_attachments', displayName: 'has_attachments', required: false, defaultMatch: false, display: true, type: 'boolean', canBeUsedToMatch: true },
          { id: 'web_link', displayName: 'web_link', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'status', displayName: 'status', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'guard_flags', displayName: 'guard_flags', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'followup_due_at', displayName: 'followup_due_at', required: false, defaultMatch: false, display: true, type: 'dateTime', canBeUsedToMatch: true }
        ]
      }
    },
    position: [2600, 0]
  },
  output: [{ id: 12, createdAt: '2026-08-26T12:00:00.000Z' }]
});

const needsDraft = ifElse({
  version: 2.3,
  config: {
    name: 'Draft This One?',
    parameters: {
      conditions: {
        options: { caseSensitive: true, leftValue: '', typeValidation: 'loose', version: 2 },
        conditions: [
          { id: 'requires-draft', leftValue: ${E("{{ $('Apply Classification').item.json.requiresDraft }}")}, operator: { type: 'boolean', operation: 'true', singleValue: true } },
          { id: 'drafting-on', leftValue: ${E("{{ $('Config').first().json.draftingEnabled }}")}, operator: { type: 'boolean', operation: 'true', singleValue: true } }
        ],
        combinator: 'and'
      },
      looseTypeValidation: true
    },
    position: [2800, 0]
  },
  output: [{ id: 12 }]
});

const callDraftWorkflow = node({
  type: 'n8n-nodes-base.executeWorkflow',
  version: 1.3,
  config: {
    name: 'Draft Reply (WF-B)',
    onError: 'continueRegularOutput',
    parameters: {
      mode: 'each',
      source: 'database',
      workflowId: { __rl: true, mode: 'id', value: ${S(wfB)}, cachedResultName: 'WF-B · Draft Reply' },
      workflowInputs: {
        mappingMode: 'defineBelow',
        value: {
          messageId: ${E("{{ $('Apply Classification').item.json.messageId }}")},
          messageIdEnc: ${E("{{ $('Apply Classification').item.json.messageIdEnc }}")},
          conversationId: ${E("{{ $('Apply Classification').item.json.conversationId }}")},
          subject: ${E("{{ $('Apply Classification').item.json.subject }}")},
          senderName: ${E("{{ $('Apply Classification').item.json.senderName }}")},
          senderEmail: ${E("{{ $('Apply Classification').item.json.senderEmail }}")},
          cleanText: ${E("{{ $('Apply Classification').item.json.cleanText }}")},
          label: ${E("{{ $('Apply Classification').item.json.label }}")},
          urgency: ${E("{{ $('Apply Classification').item.json.urgency }}")},
          reasoning: ${E("{{ $('Apply Classification').item.json.reasoning }}")},
          asks: ${E("{{ $('Apply Classification').item.json.asks }}")},
          hasAttachments: ${E("{{ $('Apply Classification').item.json.hasAttachments }}")}
        },
        matchingColumns: [],
        schema: [
          { id: 'messageId', displayName: 'messageId', required: false, defaultMatch: false, display: true, canBeUsedToMatch: true, type: 'string' },
          { id: 'messageIdEnc', displayName: 'messageIdEnc', required: false, defaultMatch: false, display: true, canBeUsedToMatch: true, type: 'string' },
          { id: 'conversationId', displayName: 'conversationId', required: false, defaultMatch: false, display: true, canBeUsedToMatch: true, type: 'string' },
          { id: 'subject', displayName: 'subject', required: false, defaultMatch: false, display: true, canBeUsedToMatch: true, type: 'string' },
          { id: 'senderName', displayName: 'senderName', required: false, defaultMatch: false, display: true, canBeUsedToMatch: true, type: 'string' },
          { id: 'senderEmail', displayName: 'senderEmail', required: false, defaultMatch: false, display: true, canBeUsedToMatch: true, type: 'string' },
          { id: 'cleanText', displayName: 'cleanText', required: false, defaultMatch: false, display: true, canBeUsedToMatch: true, type: 'string' },
          { id: 'label', displayName: 'label', required: false, defaultMatch: false, display: true, canBeUsedToMatch: true, type: 'string' },
          { id: 'urgency', displayName: 'urgency', required: false, defaultMatch: false, display: true, canBeUsedToMatch: true, type: 'string' },
          { id: 'reasoning', displayName: 'reasoning', required: false, defaultMatch: false, display: true, canBeUsedToMatch: true, type: 'string' },
          { id: 'asks', displayName: 'asks', required: false, defaultMatch: false, display: true, canBeUsedToMatch: true, type: 'string' },
          { id: 'hasAttachments', displayName: 'hasAttachments', required: false, defaultMatch: false, display: true, canBeUsedToMatch: true, type: 'boolean' }
        ],
        attemptToConvertTypes: false,
        convertFieldsToString: false
      },
      options: { waitForSubWorkflow: true }
    },
    position: [3020, -140]
  },
  output: [{ ok: true, draftId: 'AAMkDRAFT' }]
});

const noteHeader = sticky(
  ${S('## WF-A · Poll & Classify\n\nRuns every 2 minutes. Reads the inbox, classifies each new message, paints the Outlook category, writes a log row, and (once drafting is switched on) hands the message to WF-B.\n\n**Everything tunable lives in Config** — model, confidence floor, poll size, always-flag senders, and `draftingEnabled`. Nothing else in this workflow should need editing.')},
  [everyTwoMinutes, config, loadProfileFile, extractProfileText, resolveContext],
  { color: 4, width: 960, height: 260 },
);

const noteIdempotency = sticky(
  ${S('### Two independent idempotency guards\n\n1. **Outlook categories** — anything already carrying `AI-Processed` is dropped. This is also the UI: Yancey sees the coloured flags in the inbox he already uses.\n2. **The log table** — `rowNotExists` on `message_id` catches the window between tagging and a crash.\n\nThe `$filter` only asks Graph for mail newer than `lookbackHours`, so a normal run reads a handful of messages, not the whole inbox.')},
  [fetchInbox, splitMessages, skipProcessed, skipLogged],
  { color: 5, width: 800, height: 300 },
);

const noteShadow = sticky(
  ${S('### Shadow mode lives in this IF\n\n`draftingEnabled` is **false** at deploy time. WF-A labels only — exactly the 2–3 day shadow period in the spec.\n\nWhen Yancey signs off on the labels, flip `draftingEnabled` to `true` in **Config**. Nothing else changes.')},
  [needsDraft, callDraftWorkflow],
  { color: 3, width: 460, height: 260 },
);

export default workflow('upperform-wf-a-classify', ${S(NAME)})
  .add(everyTwoMinutes)
  .to(config)
  .to(loadProfileFile)
  .to(extractProfileText)
  .to(resolveContext)
  .to(fetchInbox)
  .to(splitMessages)
  .to(skipProcessed)
  .to(skipLogged)
  .to(prepareEmail)
  .to(classifyEmail)
  .to(applyClassification)
  .to(tagInOutlook)
  .to(logClassification)
  .to(needsDraft.onTrue(callDraftWorkflow))
  .add(noteHeader)
  .add(noteIdempotency)
  .add(noteShadow);
`;
}
