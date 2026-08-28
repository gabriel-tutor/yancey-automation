import { E, GRAPH, S, SDK_IMPORT, OUTLOOK_CREDENTIAL, SUFFIX, TARGET } from './shared.mjs';

const NAME = 'WF-S0 · SharePoint access probe' + SUFFIX;

/**
 * Which credential the probe runs as.
 *
 * Today this is the Phase 1 Outlook credential, because that is the thing whose
 * SharePoint reach we needed to measure — and the answer was 403 on all four
 * checks. After the Entra grant lands there will be a *different* credential of
 * a *different* type, so both are parameters rather than literals. Point
 * `sharepointCredential` / `sharepointCredentialType` at the new one in
 * config/deploy.json, rebuild, and the same four checks re-run unchanged.
 *
 * Keeping the checks identical across the change is the whole point: a green
 * run afterwards is only meaningful if it is the same test that went red.
 */
const CRED_NAME = TARGET.sharepointCredential || OUTLOOK_CREDENTIAL;
const CRED_TYPE = TARGET.sharepointCredentialType || 'microsoftOutlookOAuth2Api';

/**
 * Refuse to build a prod probe that would run as Yancey's mail credential.
 *
 * `npm run build` defaults to prod, so without this the diagnostic silently
 * acquires `Outlook – Yancey` and any deploy would put probe traffic on the
 * client's live credential. The calls are read-only, so the blast radius is
 * small — but "harmless in effect" is not a reason to leave a footgun cocked,
 * and WF-S's dev-only rule is a sticky note nobody reads at build time.
 *
 * Once a real SharePoint credential is configured for prod this stops applying,
 * because at that point pointing the probe at prod is the entire intent.
 */
if (TARGET.name === 'prod' && !TARGET.sharepointCredential) {
  throw new Error(
    'refusing to build the probe against prod with the mail credential. ' +
      'Use TARGET=dev, or set targets.prod.sharepointCredential once the SharePoint ' +
      'credential exists (see docs/PHASE2-SHAREPOINT.md).',
  );
}

const CONFIG = { graphBase: GRAPH };

/** The four checks, in the order they run. */
const CHECKS = [
  {
    varName: 'myDrive',
    label: 'My OneDrive',
    path: '/me/drive',
    proves: 'Files.* — the token can see a drive at all',
  },
  {
    varName: 'rootSite',
    label: 'Root Site',
    path: '/sites/root',
    proves: 'Sites.Read — the token can resolve a site',
  },
  {
    varName: 'searchSites',
    label: 'Search Sites',
    path: '/sites?search=*',
    proves: 'Sites.Read.All — site discovery, needed to find the target library',
  },
  {
    varName: 'followedSites',
    label: 'Followed Sites',
    path: '/me/followedSites',
    proves: 'delegated Sites access on behalf of the signed-in user',
  },
];

const VERDICT = `
function res(name) {
  const r = $(name).first().json || {};
  const body = r.body || {};
  const err = body.error || {};
  const code = Number(r.statusCode || 0);
  return {
    check: name,
    statusCode: code,
    ok: code >= 200 && code < 300,
    errCode: String(err.code || ''),
    errMsg: String(err.message || '').slice(0, 200),
    body: body
  };
}

const names = ${S(CHECKS.map((c) => c.label))};
const proves = ${S(CHECKS.map((c) => c.proves))};

const checks = [];
for (const name of names) {
  checks.push(res(name));
}

// Collect whatever sites the token could actually see. On a passing run this is
// the payload that makes the next conversation concrete: real site names and
// IDs, rather than a taxonomy proposed against imagined folders.
const sites = [];
for (const c of checks) {
  if (c.check === 'My OneDrive') continue;
  const v = c.body && c.body.value;
  if (Array.isArray(v)) {
    for (const s of v) {
      sites.push({ name: s.displayName || s.name, id: s.id, url: s.webUrl });
    }
  } else if (c.body && c.body.webUrl) {
    sites.push({ name: c.body.displayName || c.body.name, id: c.body.id, url: c.body.webUrl });
  }
}

let canReadSites = false;
for (const c of checks) {
  if (c.check !== 'My OneDrive' && c.ok) canReadSites = true;
}

const summary = [];
for (let i = 0; i < checks.length; i++) {
  const c = checks[i];
  summary.push({
    check: c.check,
    proves: proves[i],
    expected: '2xx',
    statusCode: c.statusCode,
    result: c.ok ? 'PASS' : 'FAIL',
    detail: c.ok ? '' : c.errCode + (c.errMsg ? ': ' + c.errMsg : '')
  });
}

return [{
  json: {
    verdict: canReadSites ? 'SharePoint is reachable' : 'SharePoint scopes are NOT granted',
    credential: ${S(CRED_NAME)},
    credentialType: ${S(CRED_TYPE)},
    canReadOwnDrive: checks[0].ok,
    canReadSites: canReadSites,
    whatIsNeeded: canReadSites
      ? 'none — this token already carries Sites/Files access'
      : 'a separate credential against an app registration holding Sites access. See docs/PHASE2-SHAREPOINT.md',
    sitesFound: sites,
    checks: summary,
    checkedAt: new Date().toISOString()
  }
}];
`;

/**
 * One read-only Graph GET. `neverError` + `fullResponse` so a 403 lands in the
 * data as a status code to be reported, rather than failing the run — the
 * failures *are* the measurement here.
 * @param {{ varName: string, label: string, path: string }} check
 * @param {number} x
 * @returns {string}
 */
function probeNode(check, x) {
  return `
const ${check.varName} = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: ${S(check.label)},
    onError: 'continueRegularOutput',
    parameters: {
      method: 'GET',
      url: ${E("{{ $('Config').first().json.graphBase }}" + check.path)},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: ${S(CRED_TYPE)},
      options: {
        timeout: 20000,
        response: { response: { responseFormat: 'json', fullResponse: true, neverError: true } }
      }
    },
    credentials: { ${CRED_TYPE}: newCredential(${S(CRED_NAME)}) },
    position: [${x}, 0]
  },
  output: [{ statusCode: 200, body: {} }]
});
`;
}

const STICKY = `## WF-S0 · SharePoint access probe

Answers one question before any Phase 2 work is scoped: **can this credential see SharePoint at all?**

Every call is a GET. Nothing here can write.

### Result on 2026-08-28, against the Phase 1 Outlook credential

All four checks returned **403**. Graph's wording on \`/me/followedSites\` is the informative one — *"not supported with the provided scopes"* — which is a statement about the token, not about permissions on any site.

n8n's Outlook credential type requests mail scopes only at authorization time. An access token cannot carry scopes that were never requested, so no amount of tenant-side consent widens this one. A separate credential is required regardless of what the Entra app holds.

### Re-run it after the grant

Point \`sharepointCredential\` and \`sharepointCredentialType\` at the new credential in \`config/deploy.json\` and rebuild. The four checks stay identical on purpose: a green run only means something if it is the same test that went red.

On a passing run, \`sitesFound\` carries the real site names and IDs — which is what makes a taxonomy proposal concrete rather than hypothetical.

Reasoning and the admin request: \`docs/PHASE2-SHAREPOINT.md\`, \`docs/access-request-sharepoint.md\`.`;

export async function build() {
  const probes = CHECKS.map((check, i) => probeNode(check, 400 + i * 200)).join('');
  const varNames = CHECKS.map((c) => c.varName);
  const lastX = 400 + CHECKS.length * 200;

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
  output: [${S(CONFIG)}]
});
${probes}
const verdict = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Verdict',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(VERDICT)} },
    position: [${lastX}, 0]
  },
  output: [{ verdict: 'SharePoint is reachable', canReadOwnDrive: true, canReadSites: true, sitesFound: [], checks: [] }]
});

const note = sticky(
  ${S(STICKY)},
  [runProbe, config, ${varNames.join(', ')}, verdict],
  { color: 4, width: 1600, height: 340 },
);

export default workflow('upperform-wf-s0-probe', ${S(NAME)})
  .add(runProbe)
  .to(config)
  .to(${varNames.join(')\n  .to(')})
  .to(verdict)
  .add(note);
`;
}
