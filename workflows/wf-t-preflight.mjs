import { E, GRAPH, S, SDK_IMPORT, OUTLOOK_CREDENTIAL, SUFFIX, TABLES } from './shared.mjs';

const CONFIG = {
  graphBase: GRAPH,
  // The mailbox this credential is expected to be bound to. Also the target of
  // the deliberate send attempt, so the blast radius is one self-addressed mail.
  expectedMailbox: 'gtutor@partnersperform.com',
};

const VERDICT = `
const cfg = $('Config').first().json;

function res(node) {
  const r = $(node).first().json || {};
  const body = r.body || {};
  const err = body.error || {};
  return {
    code: Number(r.statusCode || 0),
    errCode: String(err.code || ''),
    errMsg: String(err.message || '').slice(0, 200),
    body
  };
}

const checks = [];
function expectOk(name, node, why) {
  const r = res(node);
  const passed = r.code >= 200 && r.code < 300;
  checks.push({
    check: name,
    proves: why,
    expected: '2xx',
    statusCode: r.code,
    result: passed ? 'PASS' : 'FAIL',
    detail: passed ? '' : r.errCode + (r.errMsg ? ': ' + r.errMsg : '')
  });
  return passed;
}

const readOk    = expectOk('Read inbox',    'Read Inbox',    'Mail.Read \\u2014 WF-A can poll');
const draftsOk  = expectOk('List drafts',   'List Drafts',   'Mail.Read \\u2014 WF-C can reconcile');
const createOk  = expectOk('Create draft',  'Create Draft',  'Mail.ReadWrite \\u2014 WF-B can write a reply');
const deleteOk  = expectOk('Delete draft',  'Delete Draft',  'Mail.ReadWrite \\u2014 WF-B can bin a blocked draft');

// The send check is the whole point, so it is graded strictly. Only an explicit
// authorization refusal counts. Anything else means Graph rejected the request
// before it ever evaluated permissions, which proves nothing.
const send = res('Attempt Send');
let sendResult, sendDetail;

if (send.code >= 200 && send.code < 300) {
  sendResult = 'CRITICAL';
  sendDetail = 'Microsoft ACCEPTED the send. This credential holds Mail.Send.';
} else if (send.code === 401 || send.code === 403) {
  sendResult = 'PASS';
  sendDetail = 'Refused: ' + send.errCode + (send.errMsg ? ' \\u2014 ' + send.errMsg : '');
} else {
  sendResult = 'INCONCLUSIVE';
  sendDetail = 'Graph rejected the request with ' + send.code + ' ' + send.errCode +
    ' before evaluating permissions. This does NOT prove sending is blocked. Fix the request and re-run.';
}

checks.push({
  check: 'Attempt to SEND',
  proves: 'Mail.Send is NOT granted \\u2014 the core invariant',
  expected: '401 or 403 from Microsoft',
  statusCode: send.code,
  result: sendResult,
  detail: sendDetail
});

const created = $('Create Draft').first().json.body || {};
const from = (created.from || {}).emailAddress || (created.sender || {}).emailAddress || {};
const actualMailbox = String(from.address || '').toLowerCase();
const expectedMailbox = String(cfg.expectedMailbox || '').toLowerCase();
const mailboxMatches = !actualMailbox || !expectedMailbox || actualMailbox === expectedMailbox;

if (!mailboxMatches) {
  checks.push({
    check: 'Mailbox identity',
    proves: 'the credential is bound to the mailbox you think it is',
    expected: expectedMailbox,
    statusCode: 0,
    result: 'FAIL',
    detail: 'This credential is connected to ' + actualMailbox + ', not ' + expectedMailbox
  });
}

const allReadsOk = readOk && draftsOk && createOk && deleteOk;
const verdict = allReadsOk && sendResult === 'PASS' && mailboxMatches ? 'PASS' : 'FAIL';

return [{
  json: {
    verdict,
    canDoTheJob: allReadsOk,
    noSendInvariantProven: sendResult === 'PASS',
    sendCheck: sendResult,
    alarm: sendResult === 'CRITICAL'
      ? 'STOP. Mail.Send is granted. Revoke consent in Entra (Enterprise applications -> the app -> Permissions), fix Enabled Scopes, reconnect. Do NOT point this at a client mailbox.'
      : sendResult === 'INCONCLUSIVE'
        ? 'The send check did not run properly. Treat the invariant as UNVERIFIED until this reads PASS.'
        : '',
    mailbox: actualMailbox || '(not reported)',
    inboxSample: ((res('Read Inbox').body || {}).value || [])
      .map(m => ({ subject: m.subject, categories: m.categories || [] })),
    openDrafts: ((res('List Drafts').body || {}).value || []).length,
    checks,
    checkedAt: new Date().toISOString()
  }
}];
`;

export const NAME = 'WF-T · Preflight (scope proof)' + SUFFIX;

/**
 * One Graph call on the preflight canvas.
 *
 * @param {string} varName @param {string} label @param {string} method
 * @param {string} urlSuffix @param {string} position the canvas coordinates, as they appear
 * in the generated source @param {string} [jsonBody]
 */
function graphNode(varName, label, method, urlSuffix, position, jsonBody) {
  return `
const ${varName} = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: ${S(label)},
    onError: 'continueRegularOutput',
    parameters: {
      method: ${S(method)},
      url: ${E("{{ $('Config').first().json.graphBase }}" + urlSuffix)},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftOutlookOAuth2Api',${jsonBody ? `
      sendBody: true,
      contentType: 'json',
      specifyBody: 'json',
      jsonBody: ${jsonBody},` : ''}
      options: {
        timeout: 20000,
        response: { response: { responseFormat: 'json', fullResponse: true, neverError: true } }
      }
    },
    credentials: { microsoftOutlookOAuth2Api: newCredential(${S(OUTLOOK_CREDENTIAL)}) },
    position: [${position}]
  },
  output: [{ statusCode: 200, body: {} }]
});
`;
}

export async function build() {
  const createBody = S(JSON.stringify({
    subject: 'PREFLIGHT probe \\u2014 safe to delete',
    body: { contentType: 'Text', content: 'Created and deleted by WF-T to prove draft write access.' },
    toRecipients: [],
  }));

  const sendBody = E(
    '{{ JSON.stringify({ message: { subject: "PREFLIGHT \\u2014 this must never arrive", ' +
      'body: { contentType: "Text", content: "If this message exists, the Entra app holds Mail.Send and the no-send invariant is broken." }, ' +
      'toRecipients: [{ emailAddress: { address: $("Config").first().json.expectedMailbox } }] }, ' +
      'saveToSentItems: false }) }}',
  );

  return `${SDK_IMPORT}

const runPreflight = trigger({
  type: 'n8n-nodes-base.manualTrigger',
  version: 1,
  config: { name: 'Run Preflight', parameters: {}, position: [0, 0] },
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
${graphNode('readInbox', 'Read Inbox', 'GET', '/me/mailFolders/inbox/messages?$top=3&$select=id,subject,categories', '400, 0')}${graphNode('listDrafts', 'List Drafts', 'GET', '/me/mailFolders/drafts/messages?$top=5&$select=id', '600, 0')}${graphNode('createDraft', 'Create Draft', 'POST', '/me/messages', '800, 0', createBody)}${graphNode('deleteDraft', 'Delete Draft', 'DELETE', '/me/messages/{{ encodeURIComponent($(\'Create Draft\').first().json.body.id || \'missing\') }}', '1000, 0')}${graphNode('attemptSend', 'Attempt Send', 'POST', '/me/sendMail', '1200, 0', sendBody)}
const verdict = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Verdict',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(VERDICT)} },
    position: [1420, 0]
  },
  output: [{ verdict: 'PASS', canDoTheJob: true, noSendInvariantProven: true, sendCheck: 'PASS', mailbox: 'gtutor@partnersperform.com', openDrafts: 0, checks: [] }]
});

const note = sticky(
  ${S('## WF-T · Preflight — proof, not assumption\n\nRun against any mailbox before pointing the agent at it. Set `expectedMailbox` in **Config** first.\n\nFour checks prove the credential can do exactly what the agent needs: read the inbox, list drafts, create a draft, delete one. The fifth **deliberately attempts to send** and requires Microsoft to refuse it.\n\n### The send check is graded strictly\n\n- **401 / 403** -> PASS. Sending is genuinely denied.\n- **2xx** -> CRITICAL. The credential holds `Mail.Send`. Stop.\n- **anything else** -> INCONCLUSIVE, *not* a pass. A 400 means Graph rejected the request shape before it ever looked at permissions, which proves nothing.\n\nThat last rule exists because the first version of this test scored a malformed-request 400 as a pass and reported the invariant as holding on evidence that did not support it. A safety check that goes green without testing anything is worse than no check.\n\nScreenshot the Verdict output for the handover doc. A consent screenshot shows what was *asked for*; this shows what was *granted*.')},
  [runPreflight, config, readInbox, listDrafts, createDraft, deleteDraft, attemptSend, verdict],
  { color: 3, width: 1600, height: 320 },
);

export default workflow('upperform-wf-t-preflight', ${S(NAME)})
  .add(runPreflight)
  .to(config)
  .to(readInbox)
  .to(listDrafts)
  .to(createDraft)
  .to(deleteDraft)
  .to(attemptSend)
  .to(verdict)
  .add(note);
`;
}
