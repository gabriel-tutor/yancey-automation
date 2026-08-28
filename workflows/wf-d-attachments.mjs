import { codeNode } from '../scripts/bundle.mjs';
import { HELPERS } from './helpers-manifest.mjs';
import { E, GRAPH, S, SDK_IMPORT, OUTLOOK_CREDENTIAL, SUFFIX } from './shared.mjs';

const CONFIG = {
  graphBase: GRAPH,
  // Cheap model. Describing a slide or a screenshot is not a hard task, and
  // this runs on every attachment that is not plain text.
  modelAttachments: 'gpt-5.6-luna',
  // Anything larger is named in the summary but not opened.
  maxBytes: 8 * 1024 * 1024,
  // Inline images below this are signatures, logos and tracking pixels.
  inlineImageFloor: 40 * 1024,
  // Characters of attachment text allowed into the classifier prompt in total.
  totalCharBudget: 6000,
  perItemCharBudget: 2500,
  // Attachments per message. Beyond this the rest are named but not opened —
  // a forty-file dump is a filing exercise, not something to read at 07:30.
  maxAttachments: 8,
};

const TRIAGE = `
const cfg = $('Config').first().json;
const req = $('Attachment Request').first().json;
const listed = $input.first().json;
const all = Array.isArray(listed.value) ? listed.value : [];

const limits = {
  maxBytes: Number(cfg.maxBytes) || 8388608,
  inlineImageFloor: Number(cfg.inlineImageFloor) || 40960
};
const maxCount = Number(cfg.maxAttachments) || 8;

// Nothing attached: still emit one item so the summary node always runs and
// the caller always gets a reply, rather than the branch going quiet.
if (all.length === 0) {
  return [{ json: {
    messageId: String(req.messageId || ''),
    messageIdEnc: String(req.messageIdEnc || ''),
    attachmentId: '', label: '', route: 'none', reason: '', needsBytes: false, index: 0
  }}];
}

const out = [];
for (let i = 0; i < all.length; i++) {
  const a = all[i] || {};
  const verdict = i < maxCount
    ? triageAttachment(a, limits)
    : { route: 'skip', reason: 'beyond the first ' + maxCount + ' attachments on this message', ext: '', label: String(a.name || 'attachment') };

  const needsBytes = verdict.route === 'text' || verdict.route === 'office'
    || verdict.route === 'pdf' || verdict.route === 'image';

  out.push({ json: {
    messageId: String(req.messageId || ''),
    messageIdEnc: String(req.messageIdEnc || ''),
    attachmentId: String(a.id || ''),
    name: String(a.name || 'attachment'),
    contentType: String(a.contentType || ''),
    size: Number(a.size) || 0,
    label: verdict.label,
    route: verdict.route,
    reason: verdict.reason,
    needsBytes: needsBytes,
    index: i
  }});
}
return out;
`;

const EXTRACT = `
const cfg = $('Config').first().json;
const meta = $('Triage Attachments').item.json;
const fetched = $input.item.json || {};
const b64 = String(fetched.contentBytes || '');

const base = {
  attachmentId: meta.attachmentId,
  label: meta.label,
  route: meta.route,
  reason: meta.reason,
  index: meta.index,
  needsModel: false,
  text: '',
  error: ''
};

if (!b64) {
  base.error = 'Graph returned no content for this attachment';
  return { json: base };
}

// Everything that can be read without leaving the process is read here.
// PDFs and images are the only routes that cost a model call.
try {
  const bytes = Uint8Array.from(Buffer.from(b64, 'base64'));

  if (meta.route === 'text') {
    base.text = Buffer.from(b64, 'base64').toString('utf8');
    return { json: base };
  }

  if (meta.route === 'office') {
    base.text = officeText(bytes, meta.name);
    if (!base.text) base.error = 'no readable text found in the document body';
    return { json: base };
  }
} catch (e) {
  base.error = String(e.message || e).slice(0, 160);
  return { json: base };
}

// PDFs go to the model as files rather than through a text extractor, because
// a scanned deck has no text layer at all and would come back empty. Images go
// as images. Both are the same call shape.
const isPdf = meta.route === 'pdf';
const dataUri = 'data:' + (isPdf ? 'application/pdf' : (meta.contentType || 'image/png')) + ';base64,' + b64;

const instruction = isPdf
  ? 'Summarise this document for an inbox triage system. State what it is, who it appears to be for, and the specific figures, dates, decisions or requests it contains. Do not speculate beyond what is on the page. Under 200 words.'
  : 'Describe this image for an inbox triage system. If it contains text, a chart or a table, read the values out. If it is a screenshot of an error, quote the error. Say plainly if it is only a logo or decoration. Under 150 words.';

base.needsModel = true;
base.modelRequest = {
  model: String(cfg.modelAttachments || 'gpt-5.6-luna'),
  store: false,
  input: [{
    role: 'user',
    content: [
      { type: 'input_text', text: instruction },
      isPdf
        ? { type: 'input_file', filename: String(meta.name || 'document.pdf'), file_data: dataUri }
        : { type: 'input_image', image_url: dataUri }
    ]
  }]
};
return { json: base };
`;

const READ_MODEL = `
const meta = $('Extract').item.json;
const raw = $input.item.json || {};

let text = '';
try {
  // The Responses API nests the text a few levels down and the shape has moved
  // before. Walk it rather than trusting one path.
  const out = raw.output || [];
  for (const part of out) {
    const content = (part && part.content) || [];
    for (const c of content) {
      if (c && typeof c.text === 'string') text += c.text;
    }
  }
  if (!text && typeof raw.output_text === 'string') text = raw.output_text;
} catch (e) {
  text = '';
}

return { json: {
  attachmentId: meta.attachmentId,
  label: meta.label,
  route: meta.route,
  reason: meta.reason,
  index: meta.index,
  text: text,
  error: text ? '' : 'the model returned nothing for this file'
}};
`;

const SUMMARISE = `
const cfg = $('Config').first().json;
const req = $('Attachment Request').first().json;

// The authoritative list is what triage decided, so an attachment that was
// deliberately skipped is still reported. Extraction results are matched onto
// it by id; anything missing simply has no text.
const triaged = $('Triage Attachments').all().map(function (i) { return i.json; });

const extracted = {};
const incoming = $input.all();
for (const item of incoming) {
  const j = item.json || {};
  if (j.attachmentId) extracted[j.attachmentId] = j;
}

const items = [];
for (const t of triaged) {
  if (t.route === 'none') continue;
  const got = extracted[t.attachmentId] || {};
  items.push({
    label: t.label,
    route: t.route,
    reason: t.reason,
    text: got.text || '',
    error: got.error || ''
  });
}

const summary = summariseAttachments(items, {
  totalBudget: Number(cfg.totalCharBudget) || 6000,
  perItem: Number(cfg.perItemCharBudget) || 2500
});

return [{ json: {
  messageId: String(req.messageId || ''),
  attachmentText: summary.text,
  attachmentsRead: summary.read,
  attachmentsSkipped: summary.skipped,
  attachmentsFailed: summary.failed,
  attachmentCount: items.length
}}];
`;

export const NAME = 'WF-D · Read Attachments' + SUFFIX;

export async function build() {
  const triageCode = await codeNode(HELPERS['wf-d/Triage Attachments'], TRIAGE);
  const extractCode = await codeNode(HELPERS['wf-d/Extract'], EXTRACT);
  const summariseCode = await codeNode(HELPERS['wf-d/Summarise'], SUMMARISE);

  return `${SDK_IMPORT}

const attachmentRequest = trigger({
  type: 'n8n-nodes-base.executeWorkflowTrigger',
  version: 1.2,
  config: {
    name: 'Attachment Request',
    parameters: {
      inputSource: 'workflowInputs',
      workflowInputs: {
        values: [
          { name: 'messageId', type: 'string' },
          { name: 'messageIdEnc', type: 'string' }
        ]
      }
    },
    position: [0, 0]
  },
  output: [{ messageId: 'AAMkAGI2', messageIdEnc: 'AAMkAGI2' }]
});

const config = node({
  type: 'n8n-nodes-base.set',
  version: 3.5,
  config: {
    name: 'Config',
    parameters: { mode: 'raw', jsonOutput: ${S(JSON.stringify(CONFIG, null, 2))} },
    position: [200, 0]
  },
  output: [${S(CONFIG)}]
});

const listAttachments = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'List Attachments',
    retryOnFail: true,
    maxTries: 3,
    waitBetweenTries: 3000,
    parameters: {
      method: 'GET',
      url: ${E("{{ $('Config').first().json.graphBase }}/me/messages/{{ $('Attachment Request').first().json.messageIdEnc }}/attachments")},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftOutlookOAuth2Api',
      sendQuery: true,
      queryParameters: {
        parameters: [
          { name: '$select', value: 'id,name,contentType,size,isInline' }
        ]
      },
      options: { timeout: 30000, response: { response: { responseFormat: 'json' } } }
    },
    credentials: { microsoftOutlookOAuth2Api: newCredential(${S(OUTLOOK_CREDENTIAL)}) },
    position: [400, 0]
  },
  output: [{ value: [{ id: 'AAMkATT', name: 'Q3 deck.pptx', contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', size: 240000, isInline: false }] }]
});

const triage = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Triage Attachments',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(triageCode)} },
    position: [620, 0]
  },
  output: [{ attachmentId: 'AAMkATT', label: 'Q3 deck.pptx (234 KB)', route: 'office', needsBytes: true, index: 0 }]
});

const needsBytes = ifElse({
  version: 2.2,
  config: {
    name: 'Worth Opening?',
    parameters: {
      conditions: {
        options: { caseSensitive: true, leftValue: '', typeValidation: 'loose' },
        conditions: [
          { id: 'needs-bytes', leftValue: ${E('{{ $json.needsBytes }}')}, operator: { type: 'boolean', operation: 'true', singleValue: true } }
        ],
        combinator: 'and'
      }
    },
    position: [840, 0]
  }
});

const fetchContent = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Fetch Content',
    onError: 'continueRegularOutput',
    retryOnFail: true,
    maxTries: 2,
    waitBetweenTries: 3000,
    parameters: {
      method: 'GET',
      url: ${E("{{ $('Config').first().json.graphBase }}/me/messages/{{ $json.messageIdEnc }}/attachments/{{ $json.attachmentId }}")},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftOutlookOAuth2Api',
      options: {
        timeout: 60000,
        // One at a time. Graph throttles concurrent reads of a single mailbox,
        // and an eight-attachment message would trip it exactly as the tagging
        // step did.
        batching: { batch: { batchSize: 1, batchInterval: 300 } },
        response: { response: { responseFormat: 'json' } }
      }
    },
    credentials: { microsoftOutlookOAuth2Api: newCredential(${S(OUTLOOK_CREDENTIAL)}) },
    position: [1060, -120]
  },
  output: [{ id: 'AAMkATT', name: 'Q3 deck.pptx', contentBytes: 'UEsDBBQ...' }]
});

const extract = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Extract',
    parameters: { mode: 'runOnceForEachItem', language: 'javaScript', jsCode: ${S(extractCode)} },
    position: [1280, -120]
  },
  output: [{ attachmentId: 'AAMkATT', label: 'Q3 deck.pptx (234 KB)', text: '[Slide 1] Vector CI', needsModel: false }]
});

const needsModel = ifElse({
  version: 2.2,
  config: {
    name: 'Needs the Model?',
    parameters: {
      conditions: {
        options: { caseSensitive: true, leftValue: '', typeValidation: 'loose' },
        conditions: [
          { id: 'needs-model', leftValue: ${E('{{ $json.needsModel }}')}, operator: { type: 'boolean', operation: 'true', singleValue: true } }
        ],
        combinator: 'and'
      }
    },
    position: [1500, -120]
  }
});

const askModel = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Read It With Vision',
    onError: 'continueRegularOutput',
    retryOnFail: true,
    maxTries: 2,
    waitBetweenTries: 5000,
    parameters: {
      method: 'POST',
      url: 'https://api.openai.com/v1/responses',
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'openAiApi',
      sendBody: true,
      contentType: 'json',
      specifyBody: 'json',
      jsonBody: ${E('{{ JSON.stringify($json.modelRequest) }}')},
      options: { timeout: 120000 }
    },
    credentials: { openAiApi: newCredential('OpenAI account') },
    position: [1720, -220]
  },
  output: [{ output: [{ content: [{ text: 'A slide deck covering Q3 reliability performance.' }] }] }]
});

const readModel = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Read Model Reply',
    parameters: { mode: 'runOnceForEachItem', language: 'javaScript', jsCode: ${S(READ_MODEL)} },
    position: [1940, -220]
  },
  output: [{ attachmentId: 'AAMkATT', label: 'Q3 deck.pptx (234 KB)', text: 'A slide deck covering Q3 reliability performance.' }]
});

const allResults = merge({
  version: 3.2,
  config: {
    name: 'All Results',
    parameters: { mode: 'append', numberInputs: 3 },
    position: [1980, 0]
  }
});

const summarise = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Summarise',
    alwaysOutputData: true,
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(summariseCode)} },
    position: [2180, 0]
  },
  output: [{ messageId: 'AAMkAGI2', attachmentText: '- Q3 deck.pptx (234 KB):\\n[Slide 1] Vector CI', attachmentsRead: 1, attachmentsSkipped: 0, attachmentsFailed: 0, attachmentCount: 1 }]
});

const noteHeader = sticky(
  ${S('## WF-D · Read Attachments\n\nCalled by WF-A when a message has attachments. Returns one line of text per attachment, capped, for the classifier to read.\n\n**Everything is named, including what was not opened.** An attachment the assistant silently ignored is worse than one it says it could not read — Yancey can act on the second and cannot see the first.\n\n### Routes\n\n| Kind | How |\n|---|---|\n| `.docx` `.pptx` `.xlsx` | unzipped and read in process |\n| `.pdf` | sent to the model as a file, so scans work too |\n| images | sent to the model as an image |\n| text, csv, json, xml, html | decoded directly |\n| `.pbix` `.xls` `.doc` archives, media | named, not opened |\n\nThe Office path carries its own DEFLATE decompressor because n8n\'s Code node refuses `require(\'zlib\')` — probed 2026-08-29, "Module \'zlib\' is disallowed". See `src/unzip.js`.')},
  [attachmentRequest, config, listAttachments, triage],
  { color: 6, width: 820, height: 300 },
);

const noteCost = sticky(
  ${S('### Only two routes cost anything\n\nOffice files and text are read in process and cost nothing but milliseconds. PDFs and images go to the model, so the cheap classifier model is used rather than the drafting one.\n\n`maxAttachments` caps how many are opened per message; the rest are named. `maxBytes` refuses anything large. Inline images below `inlineImageFloor` are signatures and logos and are skipped without a call.')},
  [needsModel, askModel, readModel],
  { color: 5, width: 700, height: 260 },
);

export default workflow('upperform-wf-d-attachments', ${S(NAME)})
  .add(attachmentRequest)
  .to(config)
  .to(listAttachments)
  .to(triage)
  .to(needsBytes
    .onTrue(fetchContent.to(extract).to(needsModel
      .onTrue(askModel.to(readModel).to(allResults.input(2)))
      .onFalse(allResults.input(1))))
    .onFalse(allResults.input(0)))
  .add(allResults)
  .to(summarise)
  .add(noteHeader)
  .add(noteCost);
`;
}
