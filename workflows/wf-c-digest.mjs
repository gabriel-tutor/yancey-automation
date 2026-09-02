import { codeNode } from '../scripts/bundle.mjs';
import { HELPERS } from './helpers-manifest.mjs';
import { DEPLOY, E, GRAPH, S, SDK_IMPORT, OUTLOOK_CREDENTIAL, SUFFIX, TABLES, TARGET } from './shared.mjs';

const CONFIG = {
  // Display identity only — nothing is sent, so this address never has to exist
  // or be verified anywhere. It is listed in WF-A's ignoreSenders so the agent
  // does not pick up and classify its own digest on the next poll.
  digestFrom: 'assistant@agent.partnersperform.com',
  digestTo: TARGET.digestTo,
  digestCc: '',
  // Yancey, 2026-08-28: "Please send the daily summary at 7:30 AM Central,
  // Monday through Friday." The trigger below is already 07:30 Mon-Fri, and the
  // America/Chicago timezone is a workflow *setting* the SDK cannot express —
  // it is applied via MCP and survives node updates. Verify it after any
  // recreate-from-scratch, or the digest silently moves to 07:30 UTC, which is
  // 01:30 his time.
  followupDays: 3,
  autoCloseDays: 21,
  // Yancey, 2026-08-28: "For the initial rollout, please include emails
  // classified as noise in a 'Processed as Noise' section of the morning
  // summary. I want visibility into possible misclassifications while the
  // system is learning." Set false to retire the section once he is satisfied.
  showNoise: true,
  // Drafting is on in production, so the shadow-mode banner must not appear.
  mode: 'live',
  mailboxLabel: TARGET.mailboxLabel,
  graphBase: GRAPH,
};

const BUILD_DIGEST = `
const cfg = $('Config').first().json;
const draftsResponse = $('Fetch Open Drafts').first().json || {};
const rows = $input.all().map(i => i.json).filter(r => r && r.message_id);

const errors = $('Fetch Recent Errors').all()
  .map(i => i.json)
  .filter(r => r && r.error_message);

let nowIso, dateLabel, weekday;
try {
  nowIso = $now.toISO();
  dateLabel = $now.toFormat('cccc, LLLL d');
  weekday = $now.weekday;
} catch (e) {
  const d = new Date();
  nowIso = d.toISOString();
  dateLabel = d.toDateString();
  weekday = d.getUTCDay() === 0 ? 7 : d.getUTCDay();
}

// A truncated draft listing cannot prove a draft is gone. When Graph paginates,
// treat every logged draft as still present so nothing is wrongly auto-resolved.
const draftsComplete = !draftsResponse['@odata.nextLink'];
const liveDraftIds = draftsComplete
  ? (draftsResponse.value || []).map(d => String(d.id))
  : rows.map(r => String(r.draft_id || '')).filter(Boolean);

const buckets = bucketDigestRows(rows, {
  now: nowIso,
  liveDraftIds,
  followupDays: Number(cfg.followupDays) || 3,
  lookbackHours: weekday === 1 ? 72 : 24,
  autoCloseDays: Number(cfg.autoCloseDays) || 21
});

const showNoise = cfg.showNoise !== false;

const html = renderDigestHtml(buckets, {
  dateLabel,
  mailboxLabel: String(cfg.mailboxLabel || ''),
  mode: String(cfg.mode || 'live'),
  errors: errors.slice(0, 5),
  showNoise
});

// Delivered by writing the message straight into the mailbox rather than
// emailing it. Mail.ReadWrite is the permission that already writes Yancey's
// drafts, so this needs nothing new - and because the POST targets /me, the
// digest can only ever land in the mailbox the credential owns. It cannot be
// misaddressed to a third party the way an SMTP recipient can.
//
// The 0x0E07 extended property clears MSGFLAG_UNSENT, without which Graph files
// a POSTed message as a draft instead of as received mail.
const graphPayload = {
  subject: 'Inbox digest \\u2014 ' + buckets.stats.needsYou + ' need you, ' +
    buckets.stats.draftsWaiting + ' drafts waiting',
  body: { contentType: 'HTML', content: html },
  from: { emailAddress: { name: 'Upperform inbox assistant', address: String(cfg.digestFrom || '') } },
  sender: { emailAddress: { name: 'Upperform inbox assistant', address: String(cfg.digestFrom || '') } },
  toRecipients: [{ emailAddress: { name: 'Yancey Dawson', address: String(cfg.digestTo || '') } }],
  isRead: false,
  singleValueExtendedProperties: [{ id: 'Integer 0x0E07', value: '1' }]
};

return [{
  json: {
    graphPayload,
    subject: 'Inbox digest \\u2014 ' + buckets.stats.needsYou + ' need you, ' +
      buckets.stats.draftsWaiting + ' drafts waiting',
    html,
    text: renderDigestText(buckets, { dateLabel, showNoise }),
    stats: buckets.stats,
    draftsComplete,
    errorCount: errors.length,
    resolvedAt: nowIso,
    resolved: buckets.resolvedMessageIds.map(m => ({ message_id: m }))
  }
}];
`;

export const NAME = 'WF-C · Daily Digest & Follow-ups' + SUFFIX;

export async function build() {
  const buildDigestCode = await codeNode(
    HELPERS['wf-c/Build Digest'],
    BUILD_DIGEST,
  );

  return `${SDK_IMPORT}

const weekdayMorning = trigger({
  type: 'n8n-nodes-base.scheduleTrigger',
  version: 1.3,
  config: {
    name: 'Weekdays 07:30 Central',
    parameters: {
      rule: {
        interval: [{
          field: 'weeks',
          weeksInterval: 1,
          triggerAtDay: [1, 2, 3, 4, 5],
          triggerAtHour: 7,
          triggerAtMinute: 30
        }]
      }
    },
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

const fetchOpenDrafts = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Fetch Open Drafts',
    retryOnFail: true,
    maxTries: 3,
    waitBetweenTries: 5000,
    parameters: {
      method: 'GET',
      url: ${E("{{ $('Config').first().json.graphBase }}/me/mailFolders/drafts/messages")},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftOutlookOAuth2Api',
      sendQuery: true,
      specifyQuery: 'keypair',
      queryParameters: {
        parameters: [
          { name: '$select', value: 'id' },
          { name: '$top', value: '500' }
        ]
      },
      options: { timeout: 30000, response: { response: { responseFormat: 'json' } } }
    },
    credentials: { microsoftOutlookOAuth2Api: newCredential(${S(OUTLOOK_CREDENTIAL)}) },
    position: [400, 0]
  },
  output: [{ value: [{ id: 'AAMkDRAFT' }] }]
});

const fetchRecentErrors = node({
  type: 'n8n-nodes-base.dataTable',
  version: 1.1,
  config: {
    name: 'Fetch Recent Errors',
    alwaysOutputData: true,
    onError: 'continueRegularOutput',
    parameters: {
      resource: 'row',
      operation: 'get',
      dataTableId: { __rl: true, mode: 'id', value: ${S(TABLES.errorLog)}, cachedResultName: ${S(TABLES.errorLogName)} },
      matchType: 'allConditions',
      filters: { conditions: [{ keyName: 'occurred_at', condition: 'gt', keyValue: ${E('{{ $now.minus({ hours: 24 }).toISO() }}')} }] },
      returnAll: false,
      limit: 20,
      orderBy: true,
      orderByColumn: 'occurred_at',
      orderByDirection: 'DESC'
    },
    position: [620, 0]
  },
  output: [{ id: 1, occurred_at: '2026-08-26T09:00:00.000Z', workflow_name: 'WF-A · Poll & Classify', node_name: 'Fetch Inbox', error_message: 'Graph 429' }]
});

const fetchLogRows = node({
  type: 'n8n-nodes-base.dataTable',
  version: 1.1,
  config: {
    name: 'Fetch Open Log Rows',
    executeOnce: true,
    alwaysOutputData: true,
    parameters: {
      resource: 'row',
      operation: 'get',
      dataTableId: { __rl: true, mode: 'id', value: ${S(TABLES.emailLog)}, cachedResultName: ${S(TABLES.emailLogName)} },
      matchType: 'allConditions',
      filters: {
        // The label-neq-noise condition that used to sit here is gone: noise
        // rows have to be fetched for the "Processed as noise" section. They are
        // closed in the same run, so each appears exactly once.
        // (No backticks in this comment - it lives inside a template literal.)
        conditions: [
          { keyName: 'resolved_at', condition: 'isEmpty' }
        ]
      },
      returnAll: true
    },
    position: [840, 0]
  },
  output: [{ id: 12, message_id: 'AAMkAGI2', subject: 'RE: Vector CI working session', sender_name: 'Charles Coleman', label: 'opportunity', urgency: 'this_week', status: 'classified', received_at: '2026-08-26T11:58:00Z', web_link: 'https://outlook.office.com/mail/id/AAMkAGI2' }]
});

const buildDigest = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Build Digest',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(buildDigestCode)} },
    position: [1060, 0]
  },
  output: [{ subject: 'Inbox digest — 3 need you, 2 drafts waiting', html: '<div>...</div>', text: 'Upperform inbox assistant...', stats: { needsYou: 3, draftsWaiting: 2, followupsDue: 1, resolved: 1, autoClosed: 0, scanned: 7 }, draftsComplete: true, errorCount: 0, resolvedAt: '2026-08-26T12:30:00.000Z', resolved: [{ message_id: 'AAMkOLD' }] }]
});

const sendDigest = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Deliver Digest (Graph)',
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
      jsonBody: ${E('{{ JSON.stringify($json.graphPayload) }}')},
      options: { timeout: 30000 }
    },
    credentials: { microsoftOutlookOAuth2Api: newCredential(${S(OUTLOOK_CREDENTIAL)}) },
    position: [1280, -140]
  },
  output: [{ id: 'AAMkDIGEST', subject: 'Inbox digest', isDraft: false }]
});

const splitResolved = node({
  type: 'n8n-nodes-base.splitOut',
  version: 1,
  config: {
    name: 'Split Resolved Rows',
    parameters: { fieldToSplitOut: 'resolved', include: 'noOtherFields' },
    position: [1280, 160]
  },
  output: [{ message_id: 'AAMkOLD' }]
});

const markResolved = node({
  type: 'n8n-nodes-base.dataTable',
  version: 1.1,
  config: {
    name: 'Close Resolved Rows',
    onError: 'continueRegularOutput',
    parameters: {
      resource: 'row',
      operation: 'update',
      dataTableId: { __rl: true, mode: 'id', value: ${S(TABLES.emailLog)}, cachedResultName: ${S(TABLES.emailLogName)} },
      matchType: 'allConditions',
      filters: { conditions: [{ keyName: 'message_id', condition: 'eq', keyValue: ${E('{{ $json.message_id }}')} }] },
      columns: {
        mappingMode: 'defineBelow',
        value: {
          resolved_at: ${E("{{ $('Build Digest').first().json.resolvedAt }}")},
          status: 'closed'
        },
        schema: [
          { id: 'resolved_at', displayName: 'resolved_at', required: false, defaultMatch: false, display: true, type: 'dateTime', canBeUsedToMatch: true },
          { id: 'status', displayName: 'status', required: false, defaultMatch: false, display: true, type: 'string', canBeUsedToMatch: true }
        ]
      }
    },
    position: [1500, 160]
  },
  output: [{ id: 9, updatedAt: '2026-08-26T12:30:00.000Z' }]
});

const noteHeader = sticky(
  ${S('## WF-C · Daily Digest & Follow-ups\n\n07:30 America/Chicago, Monday to Friday. Monday looks back 72 hours so nothing from Friday afternoon is missed.\n\n**The digest is written into the mailbox, not emailed.** A POST to `/me/mailFolders/inbox/messages` under `Mail.ReadWrite` — the same permission that writes the drafts. `sendMail` is never called, so the no-send promise made on the go-live call still holds, and it holds more simply than before: nothing is sent by anyone.\n\nThis replaced SMTP because Resend refused to deliver to Yancey without a verified domain. Writing to the mailbox needs no domain, no third-party service, and cannot be misaddressed — the POST targets `/me`, so it can only ever land in the mailbox the credential owns.\n\n**Do not "simplify" this to `/me/sendMail`.** That would grant the automation the ability to email anyone as Yancey, contradict the footer he reads every morning, and undo the request to strip `Mail.Send` from the Entra app.')},
  [weekdayMorning, config, fetchOpenDrafts],
  { color: 4, width: 620, height: 260 },
);

const noteReconcile = sticky(
  ${S('### Draft reconciliation in one Graph call\n\nInstead of polling each draft, the workflow lists the Drafts folder once and compares. A logged draft that is no longer there was sent or discarded, so the row is closed.\n\n`alwaysOutputData` on both reads is what makes the digest go out on a quiet day — the heartbeat is the point.')},
  [fetchRecentErrors, fetchLogRows, buildDigest],
  { color: 5, width: 680, height: 280 },
);

export default workflow('upperform-wf-c-digest', ${S(NAME)})
  .add(weekdayMorning)
  .to(config)
  .to(fetchOpenDrafts)
  .to(fetchRecentErrors)
  .to(fetchLogRows)
  .to(buildDigest)
  .to(sendDigest)
  .add(buildDigest)
  .to(splitResolved)
  .to(markResolved)
  .add(noteHeader)
  .add(noteReconcile);
`;
}
