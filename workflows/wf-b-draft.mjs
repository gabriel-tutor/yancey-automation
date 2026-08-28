import { codeNode } from '../scripts/bundle.mjs';
import { HELPERS } from './helpers-manifest.mjs';
import { DEPLOY, E, GRAPH, S, SDK_IMPORT, contextLoaderNodes, OUTLOOK_CREDENTIAL, SUFFIX, TABLES } from './shared.mjs';

const CONFIG = {
  modelDraft: 'gpt-5.6-terra',
  maxDraftWords: 150,
  signOff: 'Thanks,',
  // Yancey's answer of 2026-08-28 when asked for his signature block: the name
  // alone, no title or contact block.
  signature: 'Yancey Dawson',
  // "I do not currently use a booking link." Left empty deliberately — the
  // draft prompt turns the empty value into an explicit prohibition rather
  // than silently dropping the rule.
  bookingLink: '',
  graphBase: GRAPH,
};

const RESOLVE_CONTEXT = `
const cfg = $('Config').first().json;
const raw = $input.first().json || {};
const fromFile = typeof raw.profileText === 'string' ? raw.profileText.trim() : '';
const usable = fromFile.length > 200;
const profile = usable ? fromFile : FALLBACK_PROFILE;

return [{
  json: {
    profile,
    profileSource: usable ? 'file' : 'fallback',
    draftSystem: buildDraftSystemPrompt(profile, {
      maxWords: Number(cfg.maxDraftWords) || 150,
      bookingLink: String(cfg.bookingLink || ''),
      signOff: String(cfg.signOff || 'Thanks,'),
      signature: String(cfg.signature || 'Yancey')
    })
  }
}];
`;

const COMPOSE_PROMPT = `
const req = $('Draft Request').first().json;
const draft = $input.first().json || {};

let asks = [];
try { asks = JSON.parse(req.asks || '[]'); } catch (e) { asks = []; }

const draftId = String(draft.id || '');
if (!draftId) throw new Error('Graph createReply returned no draft id');

return [{
  json: {
    draftId,
    draftIdEnc: encodeURIComponent(draftId),
    draftWebLink: String(draft.webLink || ''),
    draftBodyHtml: draft.body && draft.body.content ? String(draft.body.content) : '',
    userPrompt: buildDraftUserPrompt({
      subject: String(req.subject || ''),
      senderName: String(req.senderName || ''),
      senderEmail: String(req.senderEmail || ''),
      cleanText: String(req.cleanText || ''),
      reasoning: String(req.reasoning || ''),
      asks: Array.isArray(asks) ? asks : [],
      hasAttachments: Boolean(req.hasAttachments),
      attachmentText: String(req.attachmentText || ''),
      label: String(req.label || ''),
      urgency: String(req.urgency || '')
    })
  }
}];
`;

export const VOICE_GUARD = `
const cfg = $('Config').first().json;
const req = $('Draft Request').first().json;
const composed = $('Compose Draft Prompt').first().json;
const incoming = $input.first().json;

// 'Generate Draft' continues on error after its retries, so a model failure
// arrives here as an item carrying \`error\` instead of content. Treat it as a
// block rather than letting it throw: createReply has already put an empty
// reply in the mailbox, and only the blocked branch below deletes it. Every
// identifier the branch needs comes from the trigger and the prompt node, both
// of which have run by this point.
if (incoming && incoming.error) {
  const failed = describeDraftFailure({ item: incoming, request: req, draft: composed });
  return [{
    json: {
      ok: false,
      status: failed.status,
      needsReview: false,
      guardFlags: failed.guardFlags,
      note: failed.note,
      hard: [],
      medium: [],
      soft: [],
      words: 0,
      draftText: '',
      finalHtml: '',
      extraCategories: ['AI-Review'],
      draftId: composed.draftId,
      draftIdEnc: failed.draftIdEnc,
      draftWebLink: composed.draftWebLink,
      draftCreatedAt: new Date().toISOString(),
      messageId: failed.messageId,
      messageIdEnc: failed.messageIdEnc,
      subject: String(req.subject || '')
    }
  }];
}

const generated = extractModelText(incoming);
const maxWords = Number(cfg.maxDraftWords) || 150;

const guard = voiceGuard(generated, String(req.subject || '') + '\\n' + String(req.cleanText || ''), {
  signOff: String(cfg.signOff || 'Thanks,'),
  signature: String(cfg.signature || 'Yancey'),
  maxWords: maxWords + 70
});

const flags = guard.hard.concat(guard.medium).concat(guard.soft);
const needsReview = guard.medium.length > 0;

return [{
  json: {
    ok: guard.ok,
    status: guard.ok ? 'drafted' : 'draft_blocked',
    note: '',
    needsReview,
    guardFlags: flags.join(',').slice(0, 480),
    hard: guard.hard,
    medium: guard.medium,
    soft: guard.soft,
    words: guard.words,
    draftText: guard.text,
    finalHtml: injectReplyIntoDraftHtml(composed.draftBodyHtml, textToHtmlParagraphs(guard.text)),
    extraCategories: needsReview ? ['AI-Drafted', 'AI-Review'] : ['AI-Drafted'],
    draftId: composed.draftId,
    draftIdEnc: composed.draftIdEnc,
    draftWebLink: composed.draftWebLink,
    draftCreatedAt: new Date().toISOString(),
    messageId: String(req.messageId || ''),
    messageIdEnc: String(req.messageIdEnc || ''),
    subject: String(req.subject || '')
  }
}];
`;

const OUTLOOK_CRED =
  'credentials: { microsoftOutlookOAuth2Api: newCredential(' + S(OUTLOOK_CREDENTIAL) + ') }';

export const NAME = 'WF-B · Draft Reply' + SUFFIX;

export async function build() {
  const resolveContextCode = await codeNode(
    HELPERS['wf-b/Resolve Context'],
    RESOLVE_CONTEXT,
  );
  const composePromptCode = await codeNode(HELPERS['wf-b/Compose Draft Prompt'], COMPOSE_PROMPT);
  const voiceGuardCode = await codeNode(
    HELPERS['wf-b/Voice Guard'],
    VOICE_GUARD,
  );

  return `${SDK_IMPORT}

const draftRequest = trigger({
  type: 'n8n-nodes-base.executeWorkflowTrigger',
  version: 1.2,
  config: {
    name: 'Draft Request',
    parameters: {
      inputSource: 'workflowInputs',
      workflowInputs: {
        values: [
          { name: 'messageId', type: 'string' },
          { name: 'messageIdEnc', type: 'string' },
          { name: 'conversationId', type: 'string' },
          { name: 'subject', type: 'string' },
          { name: 'senderName', type: 'string' },
          { name: 'senderEmail', type: 'string' },
          { name: 'cleanText', type: 'string' },
          { name: 'label', type: 'string' },
          { name: 'urgency', type: 'string' },
          { name: 'reasoning', type: 'string' },
          { name: 'asks', type: 'string' },
          { name: 'hasAttachments', type: 'boolean' },
          { name: 'attachmentText', type: 'string' }
        ]
      }
    },
    position: [0, 0]
  },
  output: [{ messageId: 'AAMkAGI2', messageIdEnc: 'AAMkAGI2', conversationId: 'AAQk', subject: 'RE: Vector CI working session', senderName: 'Charles Coleman', senderEmail: 'ccoleman@partnersperform.com', cleanText: 'Please send the crossarm workbook before the Momentum Room.', label: 'opportunity', urgency: 'this_week', reasoning: 'live engagement work', asks: '["send the crossarm workbook"]', hasAttachments: true }]
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
  output: [{ profile: '# Upperform Partners — Agent Context File', profileSource: 'file', draftSystem: "Write Yancey Dawson's reply..." }]
});

const createReplyDraft = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Create Reply Draft',
    retryOnFail: true,
    maxTries: 3,
    waitBetweenTries: 3000,
    parameters: {
      method: 'POST',
      url: ${E("{{ $('Config').first().json.graphBase }}/me/messages/{{ $('Draft Request').first().json.messageIdEnc }}/createReply")},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftOutlookOAuth2Api',
      sendBody: true,
      contentType: 'json',
      specifyBody: 'json',
      jsonBody: '{}',
      options: { timeout: 30000, response: { response: { responseFormat: 'json' } } }
    },
    ${OUTLOOK_CRED},
    position: [1000, 0]
  },
  output: [{ id: 'AAMkDRAFT', webLink: 'https://outlook.office.com/mail/id/AAMkDRAFT', body: { contentType: 'html', content: '<html><body><div>quoted original</div></body></html>' } }]
});

const composeDraftPrompt = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Compose Draft Prompt',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(composePromptCode)} },
    position: [1200, 0]
  },
  output: [{ draftId: 'AAMkDRAFT', draftIdEnc: 'AAMkDRAFT', draftWebLink: 'https://outlook.office.com/mail/id/AAMkDRAFT', draftBodyHtml: '<html><body><div>quoted original</div></body></html>', userPrompt: 'INBOUND MESSAGE...' }]
});

const generateDraft = node({
  type: '@n8n/n8n-nodes-langchain.openAi',
  version: 2.3,
  config: {
    name: 'Generate Draft',
    retryOnFail: true,
    maxTries: 3,
    waitBetweenTries: 2000,
    onError: 'continueRegularOutput',
    parameters: {
      resource: 'text',
      operation: 'response',
      modelId: { __rl: true, mode: 'id', value: ${E("{{ $('Config').first().json.modelDraft }}")} },
      responses: { values: [{ type: 'text', role: 'user', content: ${E('{{ $json.userPrompt }}')} }] },
      simplify: true,
      options: {
        instructions: ${E("{{ $('Resolve Context').first().json.draftSystem }}")},
        store: false,
        promptCacheKey: 'upperform-draft-v1'
      }
    },
    credentials: { openAiApi: newCredential('OpenAI – Upperform') },
    position: [1400, 0]
  },
  output: [{ content: 'Hey Chuck,\\n\\nI will pull the crossarm workbook together and send it before the Momentum Room.\\n\\nThanks,\\n\\nYancey' }]
});

const voiceGuardNode = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Voice Guard',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(voiceGuardCode)} },
    position: [1600, 0]
  },
  output: [{ ok: true, needsReview: false, guardFlags: '', words: 28, draftText: 'Hey Chuck, ...', finalHtml: '<html><body><p>Hey Chuck,</p></body></html>', extraCategories: ['AI-Drafted'], draftId: 'AAMkDRAFT', draftIdEnc: 'AAMkDRAFT', draftWebLink: 'https://outlook.office.com/mail/id/AAMkDRAFT', draftCreatedAt: '2026-08-26T12:00:00.000Z', messageId: 'AAMkAGI2', messageIdEnc: 'AAMkAGI2', subject: 'RE: Vector CI working session' }]
});

const draftPassedGuard = ifElse({
  version: 2.3,
  config: {
    name: 'Draft Passed the Guard?',
    parameters: {
      conditions: {
        options: { caseSensitive: true, leftValue: '', typeValidation: 'loose', version: 2 },
        conditions: [{ id: 'guard-ok', leftValue: ${E('{{ $json.ok }}')}, operator: { type: 'boolean', operation: 'true', singleValue: true } }],
        combinator: 'and'
      },
      looseTypeValidation: true
    },
    position: [1800, 0]
  },
  output: [{ ok: true }]
});

const writeDraftBody = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Write Draft Body',
    retryOnFail: true,
    maxTries: 3,
    waitBetweenTries: 3000,
    parameters: {
      method: 'PATCH',
      url: ${E("{{ $('Config').first().json.graphBase }}/me/messages/{{ $json.draftIdEnc }}")},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftOutlookOAuth2Api',
      sendBody: true,
      contentType: 'json',
      specifyBody: 'json',
      jsonBody: ${E('{{ JSON.stringify({ body: { contentType: "HTML", content: $json.finalHtml } }) }}')},
      options: { timeout: 30000 }
    },
    ${OUTLOOK_CRED},
    position: [2020, -140]
  },
  output: [{ id: 'AAMkDRAFT', isDraft: true }]
});

const readCategories = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Read Original Categories',
    onError: 'continueRegularOutput',
    parameters: {
      method: 'GET',
      url: ${E("{{ $('Config').first().json.graphBase }}/me/messages/{{ $('Voice Guard').first().json.messageIdEnc }}?$select=categories")},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftOutlookOAuth2Api',
      options: { timeout: 20000, response: { response: { responseFormat: 'json', neverError: true } } }
    },
    ${OUTLOOK_CRED},
    position: [2240, -140]
  },
  output: [{ categories: ['AI-Processed', 'AI-Action'] }]
});

const tagDrafted = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Tag AI-Drafted',
    onError: 'continueRegularOutput',
    parameters: {
      method: 'PATCH',
      url: ${E("{{ $('Config').first().json.graphBase }}/me/messages/{{ $('Voice Guard').first().json.messageIdEnc }}")},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftOutlookOAuth2Api',
      sendBody: true,
      contentType: 'json',
      specifyBody: 'json',
      jsonBody: ${E('{{ JSON.stringify({ categories: ($json.categories || []).filter(c => c !== "AI-Drafted" && c !== "AI-Review").concat($("Voice Guard").first().json.extraCategories) }) }}')},
      options: { timeout: 20000 }
    },
    ${OUTLOOK_CRED},
    position: [2460, -140]
  },
  output: [{ id: 'AAMkAGI2', categories: ['AI-Processed', 'AI-Action', 'AI-Drafted'] }]
});

const updateLogDrafted = node({
  type: 'n8n-nodes-base.dataTable',
  version: 1.1,
  config: {
    name: 'Log Draft Written',
    onError: 'continueRegularOutput',
    parameters: {
      resource: 'row',
      operation: 'update',
      dataTableId: { __rl: true, mode: 'id', value: ${S(TABLES.emailLog)}, cachedResultName: ${S(TABLES.emailLogName)} },
      matchType: 'allConditions',
      filters: { conditions: [{ keyName: 'message_id', condition: 'eq', keyValue: ${E("{{ $('Voice Guard').first().json.messageId }}")} }] },
      columns: {
        mappingMode: 'defineBelow',
        value: {
          draft_id: ${E("{{ $('Voice Guard').first().json.draftId }}")},
          draft_created_at: ${E("{{ $('Voice Guard').first().json.draftCreatedAt }}")},
          draft_web_link: ${E("{{ $('Voice Guard').first().json.draftWebLink }}")},
          status: 'drafted',
          guard_flags: ${E("{{ $('Voice Guard').first().json.guardFlags }}")}
        },
        schema: [
          { id: 'draft_id', displayName: 'draft_id', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'draft_created_at', displayName: 'draft_created_at', required: false, defaultMatch: false, display: true, type: 'dateTime', canBeUsedToMatch: true },
          { id: 'draft_web_link', displayName: 'draft_web_link', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'status', displayName: 'status', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'guard_flags', displayName: 'guard_flags', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true }
        ]
      }
    },
    position: [2680, -140]
  },
  output: [{ id: 12, updatedAt: '2026-08-26T12:00:00.000Z' }]
});

const deleteEmptyDraft = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Delete the Empty Draft',
    onError: 'continueRegularOutput',
    parameters: {
      method: 'DELETE',
      url: ${E("{{ $('Config').first().json.graphBase }}/me/messages/{{ $json.draftIdEnc }}")},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftOutlookOAuth2Api',
      options: { timeout: 20000, response: { response: { neverError: true } } }
    },
    ${OUTLOOK_CRED},
    position: [2020, 180]
  },
  output: [{}]
});

const readCategoriesBlocked = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Read Categories (Blocked)',
    onError: 'continueRegularOutput',
    parameters: {
      method: 'GET',
      url: ${E("{{ $('Config').first().json.graphBase }}/me/messages/{{ $('Voice Guard').first().json.messageIdEnc }}?$select=categories")},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftOutlookOAuth2Api',
      options: { timeout: 20000, response: { response: { responseFormat: 'json', neverError: true } } }
    },
    ${OUTLOOK_CRED},
    position: [2240, 180]
  },
  output: [{ categories: ['AI-Processed', 'AI-Action'] }]
});

const tagBlockedForReview = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Tag AI-Review (No Draft)',
    onError: 'continueRegularOutput',
    parameters: {
      method: 'PATCH',
      url: ${E("{{ $('Config').first().json.graphBase }}/me/messages/{{ $('Voice Guard').first().json.messageIdEnc }}")},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftOutlookOAuth2Api',
      sendBody: true,
      contentType: 'json',
      specifyBody: 'json',
      // AI-Action is stripped as well as the others. Yancey asked what the
      // relationship between AI-Action and AI-Drafted is, and the honest answer
      // only works if the categories are mutually exclusive: AI-Action means the
      // assistant is handling it, AI-Review means it is back with him. A failed
      // draft used to end up carrying both, which is precisely the confusion he
      // was asking about.
      jsonBody: ${E('{{ JSON.stringify({ categories: ($json.categories || []).filter(c => c !== "AI-Drafted" && c !== "AI-Review" && c !== "AI-Action").concat(["AI-Review"]) }) }}')},
      options: { timeout: 20000 }
    },
    ${OUTLOOK_CRED},
    position: [2460, 180]
  },
  output: [{ id: 'AAMkAGI2', categories: ['AI-Processed', 'AI-Review'] }]
});

const updateLogBlocked = node({
  type: 'n8n-nodes-base.dataTable',
  version: 1.1,
  config: {
    name: 'Log Draft Blocked',
    onError: 'continueRegularOutput',
    parameters: {
      resource: 'row',
      operation: 'update',
      dataTableId: { __rl: true, mode: 'id', value: ${S(TABLES.emailLog)}, cachedResultName: ${S(TABLES.emailLogName)} },
      matchType: 'allConditions',
      filters: { conditions: [{ keyName: 'message_id', condition: 'eq', keyValue: ${E("{{ $('Voice Guard').first().json.messageId }}")} }] },
      columns: {
        mappingMode: 'defineBelow',
        value: {
          status: ${E("{{ $('Voice Guard').first().json.status || 'draft_blocked' }}")},
          guard_flags: ${E("{{ $('Voice Guard').first().json.guardFlags }}")}
        },
        schema: [
          { id: 'status', displayName: 'status', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true },
          { id: 'guard_flags', displayName: 'guard_flags', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true }
        ]
      }
    },
    position: [2680, 180]
  },
  output: [{ id: 12, updatedAt: '2026-08-26T12:00:00.000Z' }]
});

const noteHeader = sticky(
  ${S('## WF-B · Draft Reply\n\nCalled by WF-A, once per message that earned a draft. Never runs on its own schedule.\n\n**The no-send invariant lives here.** The reply is created with Graph `createReply`, which returns a *draft* — it is not the n8n Outlook "reply" operation, which sends. The Entra app has no `Mail.Send` scope, so even a bug cannot put mail on the wire.')},
  [draftRequest, config, resolveContext],
  { color: 4, width: 760, height: 260 },
);

const noteStandalone = sticky(
  ${S('### Do not press "Execute workflow" on this one\n\nWF-B is a **sub-workflow**. Its trigger only receives data when WF-A calls it, so running it standalone leaves every input `null` and `Create Reply Draft` posts to a URL with an empty message id.\n\nGraph answers **405 — The OData request is not supported**, which reads like an HTTP method bug and is really a missing id.\n\n**To test drafting, run WF-A instead.** It fetches real mail, classifies it, and calls this workflow with all twelve inputs filled in.\n\nTo iterate on WF-B alone: run WF-A once, then pin the output of **Draft Request** from that execution.')},
  [draftRequest],
  { color: 3, width: 460, height: 250 },
);

const noteGuard = sticky(
  ${S('### Voice Guard = spec Gate 2, in code\n\n**hard** (invented currency or percentage) -> no draft is written and the empty Graph draft is deleted, so Yancey never sees a stray blank reply.\n**medium** (invented date or large number, over length) -> the draft is written *and* the message also gets `AI-Review`.\n**soft** (missing sign-off, "I hope this email finds you well") -> repaired silently.\n\nThe rules are unit-tested in `src/agent-lib.test.js`.')},
  [generateDraft, voiceGuardNode, draftPassedGuard],
  { color: 3, width: 620, height: 300 },
);

const noteBlocked = sticky(
  ${S('### Blocked path\n\nDelete the empty draft, tag the message `AI-Review`, record why in `guard_flags`. The item then shows up under "Needs you" in the next digest instead of quietly vanishing.')},
  [deleteEmptyDraft, readCategoriesBlocked, tagBlockedForReview, updateLogBlocked],
  { color: 6, width: 900, height: 220 },
);

export default workflow('upperform-wf-b-draft', ${S(NAME)})
  .add(draftRequest)
  .to(config)
  .to(loadProfileFile)
  .to(extractProfileText)
  .to(resolveContext)
  .to(createReplyDraft)
  .to(composeDraftPrompt)
  .to(generateDraft)
  .to(voiceGuardNode)
  .to(draftPassedGuard
    .onTrue(writeDraftBody.to(readCategories.to(tagDrafted.to(updateLogDrafted))))
    .onFalse(deleteEmptyDraft.to(readCategoriesBlocked.to(tagBlockedForReview.to(updateLogBlocked)))))
  .add(noteHeader)
  .add(noteStandalone)
  .add(noteGuard)
  .add(noteBlocked);
`;
}
