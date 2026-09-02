import { E, S, SDK_IMPORT, SUFFIX, TARGET } from './shared.mjs';

/**
 * WF-P — why SharePoint is refusing.
 *
 * This exists because working it out the first time took most of a day and the
 * answer was not where anyone was looking. Writes were refused; the assumption
 * was that gtutor lacked permission on the site. He did not. The token did.
 *
 * Two endpoints tell those apart, and the difference between them is the whole
 * diagnostic:
 *
 *   _api/web/effectiveBasePermissions        what THIS TOKEN can do
 *   _api/web/getusereffectivepermissions(@u) what THE PERSON is entitled to
 *
 * The first is the intersection of the app's OAuth scope with the person's site
 * role. The second is the permission engine's answer for the person alone. When
 * the second is wider than the first, the scope is the limiter and no amount of
 * site permission will help — which is exactly what happened, because the
 * credential asks for `.default` and an unconsented permission yields a quietly
 * weaker token rather than an error.
 *
 * Read-only. It sends three GETs and writes nothing.
 */
const CONFIG = {
  host: TARGET.sharepointHost || '',
  sitePath: TARGET.sharepointSite || '',
  // SharePoint identifies a person by claim, not by address. Stored raw and
  // encoded at request time: pre-encoding it here does not survive, because
  // n8n re-encodes an expression substituted into a URL and %3A becomes %253A,
  // which SharePoint answers with an empty mask rather than an error.
  checkUser: 'i:0#.f|membership|gtutor@partnersperform.com',
};

/**
 * Both endpoints answer in OData XML sometimes and JSON other times, depending
 * on the Accept header SharePoint feels like honouring. Read either.
 *
 * No colons anywhere in the thrown string: the n8n MCP returns an execution's
 * error text and keeps only what follows the last one, so a colon truncates the
 * answer this workflow exists to give.
 */
const VERDICT = `
const mask = function (json) {
  const body = typeof json.data === 'string' ? json.data : JSON.stringify(json || {});
  const num = function (k) {
    const xml = body.match(new RegExp('<d:' + k + '[^>]*>\\\\s*(\\\\d+)\\\\s*</d:' + k + '>', 'i'));
    if (xml) { return Number(xml[1]); }
    const js = body.match(new RegExp('"' + k + '"\\\\s*:\\\\s*"?(\\\\d+)'));
    return js ? Number(js[1]) : 0;
  };
  return { low: num('Low'), high: num('High') };
};

// SPBasePermissions, low word. Edit is the one a rename or a move needs.
const verbs = function (m) {
  if (!m.low && !m.high) { return 'UNREADABLE'; }
  const has = function (bit) { return (m.low & bit) === bit ? 'yes' : 'NO'; };
  return 'view=' + has(1) + ' add=' + has(2) + ' edit=' + has(4) + ' delete=' + has(8);
};

const token = mask($('Token Can Do').first().json || {});

// Reading another user's entitlement needs EnumeratePermissions, which an
// ordinary member does not hold. That is a limit, not a fault, and saying
// "unreadable" without saying why makes a working probe look broken.
const personRaw = $('Person May Do').first().json || {};
const personCode = personRaw.error
  ? String((personRaw.details && personRaw.details.httpCode) || 'error')
  : '';
const person = mask(personRaw);

// Reads succeed under any scope that can see the site, so this reflects the
// person's real standing rather than the token's ceiling. Direct membership
// only, so a site reached through a Microsoft 365 group shows none of them.
const groupsRaw = $('Which Groups').first().json || {};
const groups = Array.isArray(groupsRaw.value) && groupsRaw.value.length
  ? groupsRaw.value.map(function (g) { return g.Title; }).join(' + ')
      .replace(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi, '#')
  : 'none-direct';

const canWrite = (token.low & 4) === 4;
const flat = function (v) { return String(v === undefined || v === null ? '' : v).replace(/\\s+/g, ' ').replace(/[:]/g, '-'); };

const personText = personCode
  ? 'unreadable (' + personCode + ' - reading another user needs EnumeratePermissions, which this account does not hold)'
  : verbs(person) + ' (low=' + person.low + ')';

// The comparison only earns its keep when the token is blocked. When writes
// work there is nothing to explain, and when the person cannot be read the
// scope is still the first thing to check, because it is the usual cause.
const conclusion = canWrite
  ? 'the token carries write, so nothing is blocked and the comparison is not needed'
  : personCode
    ? 'the token cannot write and the person entitlement could not be read - check the app scope first, naming AllSites.Write explicitly in the credential rather than relying on .default, then reconnect'
    : (person.low & 4) === 4
      ? 'the person may edit but the token cannot, so the OAuth scope is the limiter - name AllSites.Write explicitly in the credential and reconnect'
      : 'neither the token nor the person can edit, so this is a site permission question rather than a scope one';

// Returns rather than throws. A diagnostic that reports by failing shows red
// on the canvas, and a red run is exactly what someone reaches for this
// workflow to explain.
return [{
  json: {
    verdict: canWrite ? 'WRITE-OK' : 'WRITE-BLOCKED',
    token: { permissions: verbs(token), low: token.low, high: token.high },
    person: personCode ? { unreadable: personCode, why: 'reading another user needs EnumeratePermissions, which this account does not hold' } : { permissions: verbs(person), low: person.low },
    directGroups: flat(groups),
    conclusion: conclusion
  }
}];
`;

/**
 * Only buildable where a SharePoint credential exists.
 *
 * The mirror of the seeder's guard. The dev target has no SharePoint site to
 * point at, and a probe built against a target that cannot reach SharePoint
 * would report a refusal that means nothing.
 */
if (!TARGET.sharepointCredential) {
  throw new Error(
    'refusing to build the SharePoint probe for a target with no sharepointCredential. ' +
      'Only the target that owns the library can answer this.',
  );
}

export const NAME = 'WF-P · SharePoint Permissions (read-only)' + SUFFIX;

/**
 * @param {string} varName
 * @param {string} label
 * @param {string} urlSuffix appended to host + sitePath + /_api/web
 * @param {string} position
 * @param {string} [query]
 */
function readNode(varName, label, urlSuffix, position, query) {
  return `
const ${varName} = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: ${S(label)},
    onError: 'continueRegularOutput',
    parameters: {
      method: 'GET',
      url: ${E("{{ $('Config').first().json.host }}{{ $('Config').first().json.sitePath }}/_api/web" + urlSuffix)},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftSharePointOAuth2Api',${query ? `
      sendQuery: true,
      queryParameters: { parameters: ${query} },` : ''}
      sendHeaders: true,
      headers: { parameters: [{ name: 'Accept', value: 'application/json;odata=nometadata' }] },
      options: { timeout: 30000 }
    },
    credentials: { microsoftSharePointOAuth2Api: newCredential(${S(TARGET.sharepointCredential)}) },
    position: [${position}]
  },
  output: [{ Low: 0, High: 0 }]
});
`;
}

export async function build() {
  return `${SDK_IMPORT}

const runProbe = trigger({
  type: 'n8n-nodes-base.manualTrigger',
  version: 1,
  config: { name: 'Run Probe', parameters: {}, position: [0, 0] },
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
${readNode('tokenCanDo', 'Token Can Do', '/effectiveBasePermissions', '420, 0')}${readNode('whichGroups', 'Which Groups', '/currentuser/groups', '640, 0', "[{ name: '$select', value: 'Title' }]")}${readNode('personMayDo', 'Person May Do', "/getusereffectivepermissions(@u)?@u='{{ encodeURIComponent($('Config').first().json.checkUser) }}'", '860, 0')}
const verdict = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Verdict',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(VERDICT)} },
    position: [1080, 0]
  },
  output: [{}]
});

const note = sticky(
  ${S('## WF-P · SharePoint Permissions\n\n**Read-only. Three GETs, nothing written.**\n\nRun this when SharePoint refuses something, before changing any permission. It answers the only question worth asking first: is it the person or the token?\n\n| Endpoint | Question |\n|---|---|\n| `effectiveBasePermissions` | what **this token** can do — the app scope intersected with the person\'s role |\n| `getusereffectivepermissions` | what **the person** is entitled to, from the permission engine alone |\n\n**When the second is wider than the first, the scope is the limiter** and no amount of site permission will help.\n\n### The trap this was built for\n\nThe credential asks for `.default`, which means *whatever this app is already consented for*. A permission that has been added but not consented yields a quietly weaker token rather than an error — and reconnecting picks up nothing new, because there is nothing new to pick up. Name the scope explicitly instead:\n\n```\nopenid offline_access https://upperformpartners.sharepoint.com/AllSites.Read https://upperformpartners.sharepoint.com/AllSites.Write\n```\n\nThen disconnect and reconnect. Entra either prompts for consent or says plainly what is missing.\n\n**Do not press "Grant admin consent"** on the API permissions page. It grants everything configured, including `AllSites.FullControl` — tenant-wide full control, for a workflow that renames files.\n\n### Reading the groups line\n\n`Limited Access System Group For ...` and a `SharingLinks...` entry look like sharing-link access rather than membership. They are not conclusive: site rights arriving through a Microsoft 365 group never appear as a *direct* SharePoint group membership, so this line can look alarming while the person has full Contribute.')},
  [runProbe, config, tokenCanDo, whichGroups, personMayDo, verdict],
  { color: 4, width: 1240, height: 300 },
);

export default workflow('upperform-wf-p-permissions', ${S(NAME)})
  .add(runProbe)
  .to(config)
  .to(tokenCanDo)
  .to(whichGroups)
  .to(personMayDo)
  .to(verdict)
  .add(note);
`;
}
