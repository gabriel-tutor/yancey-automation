/**
 * WF-G — the endpoint behind the page Yancey opens.
 *
 *   "can we do a frontend / abstract tool for yancey so he can run it on his
 *    own like chrome extension?"
 *   "let's do webhook in a custom html"
 *
 * One webhook, called twice by a page that holds the list in between:
 *
 *   POST {}                     look, change nothing, answer with what it found
 *   POST { chosen: [...] }      carry out exactly those, answer with what it did
 *
 * That is the whole protocol, and it is why this is five nodes rather than the
 * eleven a multi-page form needed. The browser remembers what was offered, so
 * nothing here has to.
 *
 * The browser proposes and WF-F disposes. What comes back through the webhook
 * is a *request*, never an instruction: `approvedFrom` keeps only entries that
 * are shaped like a change, `mergeOverrides` keeps only settings a caller is
 * allowed to touch and refuses all arming unless canArm says this deployment
 * may write, and WF-F then re-reads the library and re-checks every proposed
 * name against the convention, the guards and the live collisions. A hostile
 * payload can at worst ask for something that gets refused.
 *
 * The answer is DATA, not markup. Presentation lives in web/index.html, where
 * it can be designed and changed without a workflow deployment.
 */
import { codeNode } from '../scripts/bundle.mjs';
import { HELPERS } from './helpers-manifest.mjs';
import { E, S, SDK_IMPORT, SUFFIX, TARGET } from './shared.mjs';

/**
 * What the caller asked for, turned into what WF-F is told.
 *
 * An empty body is a look. A body carrying `chosen` is a request to carry those
 * out, and `approvedFrom` turns them into an approved plan — which BINDS, so
 * the applying run does what was on the page rather than what its own model
 * call comes back with the second time.
 */
export const DECIDE = `
const incoming = $input.first().json || {};
const body = incoming.body || incoming;
const chosen = Array.isArray(body.chosen) ? body.chosen : [];
const settings = body.settings && typeof body.settings === 'object' && !Array.isArray(body.settings)
  ? body.settings
  : {};

// approvedFrom matches ticks against what was offered. Here the caller sends
// the changes themselves, so they are both - which also means every entry has
// to survive the same shape check as a tick, and one that does not is counted
// rather than quietly dropped.
const decided = approvedFrom(chosen.map(function (c) { return (c && c.label) || ''; }), chosen);

// Settings first, ticks second, and the order is the point: approvedFrom always
// writes all five arming switches, so a caller that hides applyRenames in its
// settings has it overwritten rather than honoured. What arms a run is what was
// ticked, and nothing else.
const overrides = Object.assign({}, settings, decided.overrides);

return [{ json: {
  overrides: JSON.stringify(overrides),
  applying: decided.chosen.length > 0,
  chosenCount: decided.chosen.length,
  unusable: decided.unknown.length
}}];
`;

/**
 * The answer.
 *
 * Everything the page needs to draw itself and nothing it does not: what was
 * scanned, what could change, what was refused and why, what the run could not
 * see, and — on an applying call — what actually happened.
 */
export const ANSWER = `
const summary = $input.first().json || {};
const asked = $('Decide').first().json || {};

const offered = offeredChanges(summary);

// A refusal is worth saying once, however many things it applies to. The first
// live run of the earlier form put the same sentence on one page eight times.
function group(entries) {
  const out = [];
  const seen = {};
  for (const entry of entries) {
    if (!entry || !entry.because) continue;
    const at = seen[entry.because];
    if (at === undefined) {
      seen[entry.because] = out.length;
      out.push({ names: entry.name ? [entry.name] : [], because: entry.because });
    } else if (entry.name && out[at].names.indexOf(entry.name) === -1) {
      out[at].names.push(entry.name);
    }
  }
  return out;
}

return [{ json: {
  ok: true,
  applying: asked.applying === true,
  library: String(summary.library || ''),
  scanned: Number(summary.scanned) || 0,
  compliant: Number(summary.compliant) || 0,
  offered: offered,

  // What the run could NOT see or do. A page that hides these claims a
  // completeness the run did not have.
  truncated: summary.truncated === true,
  pages: Number(summary.pages) || 1,
  beyondModel: Number(summary.beyondModel) || 0,
  readContents: summary.readContents === true,
  refused: group((summary.refused || []).map(function (r) {
    return { name: String(r.file || ''), because: String(r.because || '') };
  })),
  folderRefusals: group((summary.folderRefusals || []).map(function (r) {
    return typeof r === 'string'
      ? { name: '', because: r }
      : { name: String(r.name || r.folder || ''), because: String(r.because || r.why || '') };
  })),
  unreadable: summary.reportsUnreadable || [],

  // What this deployment would not be talked into, and whether it may write at
  // all. A page that cannot say why it did less than it was asked to is a page
  // that looks broken.
  canArm: summary.canArm === true,
  overrideRefusals: summary.overrideRefusals || [],
  unusable: Number(asked.unusable) || 0,

  // What the run actually used, so the page can put the toggles back where the
  // deployment left them rather than where the person moved them - and the
  // whole configuration alongside, so the page can show what the tool is
  // pointed at without anyone opening n8n.
  settings: {
    renameReports: summary.renameReports === true,
    readContents: summary.readContents === true,
    maxChangesPerRun: Number(summary.maxChangesPerRun) || 0
  },
  config: summary.configuration || {},

  // Only meaningful on an applying call.
  applied: summary.applied || [],
  failed: summary.failed || [],
  heldBack: summary.heldBack || []
}}];
`;

export const NAME = 'WF-G · Organise Documents (web)' + SUFFIX;

/**
 * The webhook path.
 *
 * Long and unguessable, and no longer load-bearing: the webhook requires an
 * X-Organise-Key header, so the path is a route rather than a secret. It never
 * reaches a browser either — the page talks to a Vercel function and the
 * function talks to this — so nothing is lost by it being ugly.
 */
export const PATH = 'organise-documents-a7f3c19e4b2d';

const HEADER_NOTE = [
  "## WF-G · the endpoint behind Yancey's page",
  '',
  'One webhook, called twice by `web/index.html` (hosted on Vercel):',
  '',
  '- `POST {}` — look at the library, change nothing, answer with what it found',
  '- `POST { chosen: [...] }` — carry out exactly those, answer with what it did',
  '',
  'The browser holds the list between the two calls, which is why this is five',
  'nodes and not eleven.',
  '',
  'The answer is **data, not markup** — the page is designed in the repo, not here.',
].join('\n');

const SAFETY_NOTE = [
  '### The browser proposes, WF-F disposes',
  '',
  'What arrives here is a **request**, never an instruction:',
  '',
  '1. `approvedFrom` keeps only entries shaped like a change.',
  "2. `mergeOverrides`, in WF-F's Config, keeps only settings a caller may touch,",
  '   lowers the cap but never raises it, and refuses **all** arming unless',
  '   `canArm` says this deployment may write.',
  '3. WF-F re-reads the library and re-checks every name against the convention,',
  '   the guards and the live collisions.',
  '',
  'A hostile payload can at worst ask for something that gets refused.',
  '',
  '**The webhook requires Header Auth.** `X-Organise-Key`, held by the Vercel',
  'function and never by a browser. Without it this answers 403 before running',
  'anything. The path is a route, not a secret.',
].join('\n');

/**
 * The sandbox says so on its own canvas.
 *
 * The dev copies were created with hand-written notes because the generator had
 * no way to say "this is the sandbox", and a hand-written note is one nobody
 * rebuilds — this one went on claiming the webhook was unauthenticated for a
 * day after it was not. Generating it means the next rebuild corrects it.
 */
const SANDBOX_NOTE = [
  '## Sandbox — not the one Yancey runs',
  '',
  'This calls the **dev** organiser, whose `canArm` is false: every arming switch',
  'is forced off, so a run reads the library and refuses to change it. The page',
  'says so rather than pretending.',
  '',
  'It reads the **same library through the same credential** — there is only one,',
  'and the real edge cases are in the real library. Isolated for deployment, not',
  'for writes.',
  '',
  'Prove it here. Then production.',
].join('\n');

export async function build() {
  const wfF = (TARGET.workflows && TARGET.workflows.wfF) || '';
  if (!wfF) throw new Error('no WF-F workflow id for target ' + TARGET.name + '; deploy the organiser first');

  // No credential, no build. An endpoint that renames a client's documents is
  // not something to ship open by omission, and a missing setting is the most
  // ordinary way that happens.
  const webhookAuth = TARGET.webhookAuthCredential;
  if (!webhookAuth) {
    throw new Error(
      'refusing to build the web endpoint for ' + TARGET.name +
      ' with no webhookAuthCredential: the webhook would accept anything that found the path'
    );
  }

  const decideCode = await codeNode(HELPERS['wf-g/Decide'], DECIDE);
  const answerCode = await codeNode(HELPERS['wf-g/Answer'], ANSWER);
  const path = PATH + (TARGET.name === 'prod' ? '' : '-' + TARGET.name);

  return `${SDK_IMPORT}

const hook = trigger({
  type: 'n8n-nodes-base.webhook',
  version: 2.1,
  config: {
    name: 'Asked',
    parameters: {
      httpMethod: 'POST',
      path: ${S(path)},
      // The header is the guard; the path is only a route. Declared here so a
      // rebuild cannot quietly hand back an open endpoint - the deployment had
      // this before the repo did, which is the shape of drift that ends with
      // somebody redeploying the protection away and nobody noticing.
      authentication: 'headerAuth',
      responseMode: 'responseNode',
      options: {}
    },
    credentials: { httpHeaderAuth: newCredential(${S(webhookAuth)}) },
    position: [0, 0]
  },
  output: [{ body: { chosen: [] } }]
});

const decide = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Decide',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(decideCode)} },
    position: [220, 0]
  },
  output: [{ overrides: '{}', applying: false, chosenCount: 0, unusable: 0 }]
});

const organise = node({
  type: 'n8n-nodes-base.executeWorkflow',
  version: 1.3,
  config: {
    name: 'Run The Organiser',
    onError: 'continueRegularOutput',
    parameters: {
      mode: 'once',
      source: 'database',
      workflowId: { __rl: true, mode: 'id', value: ${S(wfF)}, cachedResultName: ${S('WF-F · Organise Documents' + SUFFIX)} },
      workflowInputs: {
        mappingMode: 'defineBelow',
        value: { overrides: ${E('{{ $json.overrides }}')} },
        matchingColumns: [],
        schema: [
          { id: 'overrides', displayName: 'overrides', required: false, defaultMatch: false, display: true, canBeUsedToMatch: true, type: 'string' }
        ],
        attemptToConvertTypes: false,
        convertFieldsToString: true
      },
      options: { waitForSubWorkflow: true }
    },
    position: [440, 0]
  },
  output: [{ scanned: 27, compliant: 27, toRename: [], toMove: [], newFolders: [], applied: [], failed: [] }]
});

const answer = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Answer',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(answerCode)} },
    position: [660, 0]
  },
  output: [{ ok: true, scanned: 27, compliant: 27, offered: [], applied: [], failed: [] }]
});

const respond = node({
  type: 'n8n-nodes-base.respondToWebhook',
  version: 1.1,
  config: {
    name: 'Answer The Page',
    parameters: {
      respondWith: 'allIncomingItems',
      options: {
        responseHeaders: {
          entries: [
            { name: 'Cache-Control', value: 'no-store' }
          ]
        }
      }
    },
    position: [880, 0]
  },
  output: [{}]
});

const noteHeader = sticky(
  ${S(HEADER_NOTE)},
  [hook, decide, organise],
  { color: 4, width: 760, height: 320 }
);

const noteSafety = sticky(
  ${S(SAFETY_NOTE)},
  [answer, respond],
  { color: 3, width: 700, height: 420 }
);
${TARGET.name === 'prod' ? '' : `
const noteSandbox = sticky(
  ${S(SANDBOX_NOTE)},
  [organise],
  { color: 7, width: 700, height: 320 }
);
`}
export default workflow(${S('upperform-wf-g-selfserve' + (TARGET.name === 'prod' ? '' : '-' + TARGET.name))}, ${S(NAME)})
  .add(hook)
  .to(decide)
  .to(organise)
  .to(answer)
  .to(respond)
  .add(noteHeader)
  .add(noteSafety)${TARGET.name === 'prod' ? '' : '\n  .add(noteSandbox)'};
`;
}
