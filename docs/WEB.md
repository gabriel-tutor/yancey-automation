# The page Yancey opens

A one-page tool at a Vercel URL that runs the document organiser. No install, no
n8n login, works on a phone. He opens it, presses a button, reads what it wants
to do, ticks what he agrees with, and sees what happened.

```
  browser                 Vercel                     n8n
  ────────                ──────                     ───
  web/index.html  ──────► web/api/organise.js ─────► WF-G webhook
                          (holds the URL              │
                           and the secret)            ▼
                                                     WF-F organiser
```

## Why a proxy and not a direct call

The page could call the n8n webhook itself. It does not, for two reasons that
are the same reason:

- **The webhook URL is the only thing guarding a workflow that renames a
  client's documents.** Called from the browser it would be in view-source on
  every laptop that opens the page.
- **CORS.** A cross-origin call needs n8n to allow the Vercel origin. Same
  origin needs nothing.

The function is deliberately thin. Everything that decides anything is in the
workflow, where it is tested.

## The protocol

One endpoint, called twice, and the browser holds the list in between — which is
why the workflow is five nodes rather than the eleven a multi-page form needed.

| Call | Body | What happens |
| --- | --- | --- |
| Look | `{ "chosen": [] }` | Reads the library, changes nothing, answers with what it found |
| Carry out | `{ "chosen": [ …changes… ] }` | Applies exactly those, answers with what it did |

The `chosen` entries are the objects the look returned, sent back unmodified.

### The browser proposes, WF-F disposes

What arrives at the webhook is a **request**, never an instruction. Three things
stand between it and the library:

1. `approvedFrom` keeps only entries shaped like a change, and counts the rest.
2. `mergeOverrides`, in WF-F's `Config` node, keeps only settings a caller may
   touch — never the host, the site or the library — lowers the change cap but
   never raises it, and refuses **all** arming unless `canArm` says this
   deployment may write.
3. WF-F re-reads the library and re-checks every proposed name against the
   convention, the guards and the live collisions.

A hostile payload can at worst ask for something that gets refused.

The ticks become an **approved plan**, and an approved plan binds. That matters
because the model is not deterministic: two runs an hour apart over the same
unchanged library proposed `Time-Zone-Conversion-Table` and then
`Timezone-Conversion-Table`. Both are reasonable; neither is the one somebody
read. Binding the plan is what makes the middle step an approval rather than a
demonstration.

## Who can use it

One shared passcode, because there is one person. Accounts, an identity
provider and per-user roles would be machinery for a population that does not
exist, and machinery nobody exercises is machinery nobody notices breaking.

```
  browser                     Vercel
  ────────                    ──────
  POST /api/session ────────► compares the passcode in constant time
                              ◄──── httpOnly signed cookie (12 hours)

  POST /api/organise ───────► no valid cookie → 401, and the request
                              never reaches n8n or SharePoint
```

What the browser holds is a token saying only when it expires, signed with
`SESSION_SECRET`. It is not the passcode, it is not readable from page
JavaScript (`HttpOnly`), it never crosses plain HTTP (`Secure`), and another
site cannot make the browser spend it (`SameSite=Lax`). Editing the expiry in
devtools invalidates the signature — `src/session.test.js` proves that rather
than assuming it.

**The gate is on the endpoint, not only on the page.** A page that hides a
button is a page, not a guard; `/api/organise` is what reaches SharePoint, so
`/api/organise` is what asks.

**It fails closed.** A deployment with no `ORGANISE_PASSCODE` set lets nobody
in and says so. A missing setting is a broken deployment, not an open one.

### What this is not

Sign-in sleeps 600ms on every attempt, right or wrong. That is not a rate
limiter and should not be mistaken for one — a serverless function keeps no
state between requests, so there is nowhere to count failures. It buys time;
the length of the passcode does the rest. If this ever guards something more
than one person's document library, put a real limiter in front of it.

### Changing the passcode

```sh
cd web
vercel env rm ORGANISE_PASSCODE production --yes
printf 'the-new-passcode' | vercel env add ORGANISE_PASSCODE production
vercel --prod
```

Everyone signed in stays signed in until their token expires. To end every
session at once, rotate `SESSION_SECRET` the same way — every token that was
issued stops verifying immediately.

## Environment variables

Set these on the Vercel project (Settings → Environment Variables), not in the
repo:

| Name | Value |
| --- | --- |
| `N8N_ORGANISE_URL` | `https://n8n-wevu.srv1921545.hstgr.cloud/webhook/organise-documents-a7f3c19e4b2d` — production. The dev copy is the same path with `-dev`. |
| `N8N_ORGANISE_SECRET` | Optional today. See hardening below. |
| `ORGANISE_PASSCODE` | The passcode Yancey types. Required — without it nobody can sign in. |
| `SESSION_SECRET` | Signs the session cookie. Rotate it to end every session at once. |

The dev sandbox has its own path (`…-dev`). Its `canArm` is false, so it will
read the library and refuse to change it — and the page says so rather than
pretending. To point the page at the sandbox instead, swap `N8N_ORGANISE_URL`
for the `-dev` path and redeploy; nothing else changes.

## What production means here

| | |
| --- | --- |
| Page | https://yancey-document-organiser.vercel.app |
| Endpoint | WF-G `IoEjcbqG0OBu6yc6`, published |
| Organiser | WF-F `CC8jik1RGti4JXZk`, published, `canArm: true` |
| Failures | routed to WF-E, logged to `agent_error_log` and alerted |
| Sandbox | WF-G `spVTAZQWeL9PONgW` → WF-F `6r1FMIaY7nZBnGWl`, `canArm: false`, both published |

Every node body in both production workflows was byte-compared against the build
before publishing, and every change was proved on the sandbox against this same
library first.

### What is tested, and where

| | |
| --- | --- |
| The naming, placement and apply logic | `src/*.test.js` — the bulk of the suite |
| Every Code node body, executed | `src/nodes.test.js` |
| The page's rendering, settings and storage | `src/web-page.test.js`, read out of the shipped `web/index.html` |
| The page's wiring — that every control has a listener, that every element the script reaches for exists | `src/web-page.test.js` |
| The proxy: the session gate, what it forwards, what it refuses to leak | `src/web-api.test.js` |
| The session token and cookie | `src/session.test.js` |
| Which library a run may be pointed at, and the refusals | `src/selfserve.test.js`, `src/web-page.test.js` |
| The deploy gate itself: node diffs, missing settings, canvas notes, the wrong target | `src/verify-deployed.test.js` |
| That the undo record reconciles with the library | `src/undo.test.js` |
| That the demo runbook never promises a rename already applied | `src/demo-runbook.test.js`, read out of the built page |
| That each target routes failures to its own error handler | `src/deployment.test.js` |

The wiring tests exist because two of the three bugs the page produced were
invisible to unit tests of the functions involved: the functions were right and
nothing connected them. Each of those tests has been checked by putting the bug
back and confirming the test fails.

## The webhook requires Header Auth

Both deployments answer `403` to a request that does not carry the right
`X-Organise-Key` header, before running anything. The key is held by the Vercel
function and never by a browser.

| | |
| --- | --- |
| Header | `X-Organise-Key` |
| Credential | `Header Auth account`, `oHVC3A1NTFlDgDb8`, shared by both targets |
| Set on | WF-G production and WF-G sandbox, both published |
| Sent by | `web/api/organise.js`, whenever `N8N_ORGANISE_SECRET` is set |

Measured rather than assumed: no header → 403 in 0.67s; a wrong key → 403; the
right key → 200; the page end to end → 200, 29 scanned, 29 already correct, 0
written.

**The path is a route, not a secret.** It is committed to this repo, and that is
now fine — before the header it was the only thing guarding a workflow that
renames a client's documents, and production's `canArm` is true.

The generator declares the authentication too, and **refuses to build** for a
target with no `webhookAuthCredential`:

```
SKIP wf-g-selfserve -> refusing to build the web endpoint for prod with no
webhookAuthCredential: the webhook would accept anything that found the path
```

That refusal exists because the deployment had this protection before the repo
did, which is how somebody redeploys it away without noticing.

## What the repo cannot build, and checks instead

The SDK builder is `workflow(id, name)` — an id and a name. It cannot express
workflow **settings**, so where a failure is routed, whether error runs are kept
and how long a run may take exist only in the n8n editor. No generator produces
them and, until this, nothing looked at them. That is exactly the shape of drift
that left a line in production's `Read Reports` the repo had never built.

`config/deploy.json` records what they must be, per target:

| Setting | Production | Sandbox |
| --- | --- | --- |
| `errorWorkflow` | `XSKk6eHTFGqftrUM` (WF-E) | `E302rNL5D7NuBy5c` (WF-E · dev) |
| `saveDataErrorExecution` | `all` | `all` |
| `executionTimeout` | per unit — 900s for the organiser, 300s for the poller, 120s for a diagnostic | the same, and never longer |

The time limit is per unit because the right ceiling is not one number: a poller
that runs every minute must not still be running when the next one starts, and
the organiser reads, names and files a whole library. What is always wrong is no
ceiling at all.

**The error handler is the one exception.** Pointing WF-E at itself is a loop, so
it is the only workflow that routes nowhere — named in the file as
`errorWorkflowExcept` rather than remembered.

Three things hold all of this: `src/deployment.test.js` fails if a target's
`errorWorkflow` is not that target's own WF-E or if a unit has no time limit,
`src/verify-deployed.test.js` covers the check itself, and `npm run verify`
compares the recorded settings against the deployment beside the node bodies.

### What the audit of 2026-09-04 found

Every workflow in both targets was read. Three were wrong:

- **WF-D · Read Attachments, in production and active.** No error workflow, no
  error-execution data, no time limit. It reads Yancey's mail attachments and
  sends PDFs and images to OpenAI; a failure logged nothing, alerted nobody and
  left nothing to read afterwards. WF-A, WF-B and WF-C had all three since
  August — WF-D was added on 2 September and was missed.
- **Both organiser sandbox copies**, and the dev copy of WF-D: the same three
  missing. A sandbox failure failed silently.
- **Both WF-G workflows carried no tags**, so they appeared in no environment
  filter and no system filter, against the convention that every workflow has
  exactly one of each.

All fixed. The diagnostics (WF-T, WF-P, WF-S) were given the same settings even
though n8n only fires an error workflow for production executions and they are
hand-run — one uniform rule with one exception is easier to keep than several.

## Choosing which library it looks at

Asked for on 4 September, watching a run finish:

> *"Perfect, perfect. That's what I'm looking for. And what happens if I want to
> point it at another folder?"*

This is the one setting `mergeOverrides` was written to refuse. Pointing a
workflow that renames documents at somewhere new is plainly **widening** what a
run does, and the rule the whole file is built on is that a caller may narrow a
run and never widen it.

So it does not open on a caller's say-so. It opens under a policy the
**deployment** states, which makes it the same shape as `assumeStatus` picking
from the statuses the convention recognises — a choice between options that were
declared, not a free hand.

| `allowedSites` | `allowAnySiteOnHost` | What a caller may do |
| --- | --- | --- |
| empty | `false` | Nothing. Locked to the built site, exactly as before. |
| a list | either | Pick one of the listed sites, and nothing else. |
| empty | `true` | Name any well-formed site path — on this host. |

**Locked is the default.** A deployment that says nothing behaves exactly as it
did before this existed, and the test that proved the site was unreachable still
passes unchanged.

**The host is on none of those rows and never will be.** It is the tenant. No
choice of tenant is a narrowing, and the SharePoint credential cannot leave it
anyway.

### Why production is open

Both targets ship `allowAnySiteOnHost: true`. The reasoning, so it can be argued
with rather than inherited:

- The credential already holds `AllSites.Write` on this host, so choosing where
  to look grants no access it did not already have.
- Nothing is written without a review and explicit ticks, and `canArm` still
  gates the deployment.
- Whoever is choosing has already passed the passcode gate and holds the
  webhook key.

What it does change is **blast radius**, and one thing worth being deliberate
about: `readContents` is on, so pointing at a new library sends *that* library's
documents to OpenAI. To tighten, replace the flag with an `allowedSites` list —
`mergeOverrides` prefers the list whenever one is present.

### The library title is validated, not escaped

`libraryTitle` is interpolated into `getbytitle('<title>')` in the request URL.
While it was a constant that was nobody's problem; the moment a caller can set
it, an apostrophe closes the quote and the rest is read as OData.

It is **refused** rather than escaped — plain letters, digits, spaces, dots,
dashes and underscores only — because a refusal is something a test can check
and an escape is a thing you get wrong once. `src/selfserve.test.js` sends
`Docs') or 1 eq 1--` and asserts the deployment's own library survives.

### Proved on the sandbox before production

| Sent | Ran against | Refusal |
| --- | --- | --- |
| `/SITES/TaskManagement` | `/SITES/TaskManagement` | none — honoured end to end |
| `/personal/yancey_…` | `/sites/TaskManagement` | "must look like /sites/Something" |
| `Docs') or 1 eq 1--` | `Documents` | "must be plain letters…" |

The first is the one that matters: a case variant SharePoint resolves anyway, so
the run succeeded and the value came back in `configuration.sitePath` — proving
it travelled into the request rather than being quietly dropped.

The answer reports where the copy **may** be pointed separately from where the
run **was** pointed (`homeSite`, `homeLibrary`, `allowedSites`,
`allowAnySiteOnHost`), because a run that was pointed somewhere must not report
that somewhere as the only place it can go.

## Sticky notes are checked too

They are documentation on the canvas rather than behaviour, which is exactly why
nothing looked at them. The sandbox's note went on saying the webhook was
"unguessable but not authenticated" for a day after it was authenticated, and a
note that lies is worse than no note: it sends the next person to add protection
that is already there.

They are keyed by their **heading**, because n8n names a sticky node with a
generated hash that differs between the two deployments, so the node name is no
key at all. `npm run emit` reads them back out of the SDK source each generator
builds, so a note written inline at the call is covered like any other.

Checking them found production's canvas describing an older workflow: the header
note did not mention that reports are opened for their Power Query, the note on
what stops this going wrong predated `approvedPlan`, the URL-length limit and the
folder ordering, and the note on the folder structure — the one thing a client
asked for by name — had never been deployed at all. The sandbox had one
hand-written note where the build produces five.

A note **on the canvas but not in the build** fails too. The convention here is
that every workflow is generated, with no exceptions, so a note somebody typed
into n8n is the same drift as a hand-pasted line in a Code node.

## Proving the deployment is the repo

Deploying through the n8n MCP means reproducing node bodies by hand, so the only
honest answer to "is the workflow that is running the workflow this repo builds"
is a byte comparison. It has earned its place: it caught 4.5KB of dead prompt
riding into the apply node, a stray newline from splicing a node out of two
pieces, and a line in production's Read Reports the repo had never produced.
None of those would have failed a behaviour test.

```sh
# what the build says the nodes should be
npm run emit wf-f dist/expected.json          # or: TARGET=dev npm run emit ...
npm run emit wf-g dist/expected-web.json

# what is actually deployed
npm run verify -- <deployed.json> dist/expected.json      # TARGET=dev for the sandbox
```

It exits non-zero on any difference and says which node and which character, so
it can gate a deploy. Normalisation matters and it does it for you: esbuild
emits `\uXXXX` escapes where a paste round-trips to the character itself — the
same string at runtime, different bytes — so both sides are decoded before
comparing. It checks the recorded settings in the same pass, and refuses
outright if the file is a sandbox workflow being checked against production's
settings, or the reverse — that difference is a mistake in the command, not in
the deployment, and reporting it as drift would train people to ignore it.

### Getting the deployed file

The n8n MCP writes `get_workflow_details` to a file when the answer is large and
returns it inline when it is small. WF-F arrives as a file; WF-G does not.

**Do not copy an inline answer into a file by hand.** That makes the comparison
circular — it proves the build matches somebody's typing, which is the exact
failure the byte check exists to catch. The session transcript holds every tool
result verbatim, so take it from there:

```sh
T=~/.claude/projects/<project>/<session>.jsonl

# whole workflow, from one get_workflow_details
npm run pull -- $T IoEjcbqG0OBu6yc6 dist/deployed-wf-g.json

# nodes from the PUBLISHED version, name and settings from the live read
npm run pull -- $T spVTAZQWeL9PONgW dist/deployed-wf-g-dev.json --version <versionId>
```

The second form exists because a workflow's draft can be ahead of what is
published, and `get_workflow_version` returns the version that is actually
running but carries no settings. It refuses to join the two halves if the live
read is not at that same version.

**An n8n API key would make all of this unnecessary** — the repo could fetch a
workflow directly, and `npm run deploy` could replace pasting node bodies into
the editor. That is the single largest remaining piece of manual handling here.

### Verified on 2026-09-04

| | Nodes | Canvas notes | Settings |
| --- | --- | --- | --- |
| WF-F production | 11 of 11 | 4 of 4 | match |
| WF-F sandbox | 11 of 11 | 5 of 5 | match |
| WF-G production | 2 of 2 | 2 of 2 | match |
| WF-G sandbox | 2 of 2 | 3 of 3 | match |

The sandboxes carry one note more than production: a banner saying they are the
sandbox, which the generator emits for dev only. It used to be hand-written,
which is why its sibling was the note that went stale.

### What is not checked

Node **positions** on the canvas, and the connections between nodes. A node
moved or rewired by hand would pass. Bodies, notes and settings are what a
rebuild would overwrite, and they are what this covers.

## Running it locally

```sh
cd web
npx vercel dev
```

`vercel dev` reads the same environment variables once they are pulled with
`npx vercel env pull`.

## Deploying

```sh
cd web
npx vercel --prod
```

The project root is `web/`, not the repo root — the repo is the automation
suite, and the page is one part of it.

## What the page shows that a status readout would not

Every one of these is a fact about what the run could **not** do, and each is on
the page because a version without it read as either broken or falsely complete:

- the library was read only in part (nothing is applied while that is true)
- more documents could be renamed than this run looked at
- this deployment is not allowed to change anything
- a setting the caller asked for was refused, and why
- changes that did not fit in one run
- every document and folder left alone, grouped by reason rather than repeated
  once per file

And when there is nothing to do, it says so as good news — "all 27 documents
already follow the convention" — rather than showing an empty list.
