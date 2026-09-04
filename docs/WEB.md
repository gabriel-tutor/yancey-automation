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
| Sandbox | WF-G `spVTAZQWeL9PONgW` → WF-F `6r1FMIaY7nZBnGWl`, `canArm: false` |

Every node body in both production workflows was byte-compared against the build
before publishing, and every change was proved on the sandbox against this same
library first.

## The one step left: Header Auth on the webhook

**This is the last thing standing between production and being properly closed,
and it needs two minutes in the n8n UI that no API can do for you.**

Today the webhook is guarded by an unguessable path and by never being shown to
a browser. That was proportionate for a sandbox. It is thinner in production for
two reasons: the path is committed to this repo, and production's `canArm` is
true, so anyone who reached the webhook with a well-formed payload could rename
a client's documents.

Vercel already sends the header. Only the n8n half is missing:

1. **Credentials → New → Header Auth.** Name it `Organise webhook key`.
   - Header name: `X-Organise-Key`
   - Value: the key printed when `N8N_ORGANISE_SECRET` was set. Read it back with
     `cd web && vercel env pull` if you no longer have it.
2. **WF-G → the `Asked` node → Authentication → Header Auth**, pick that
   credential, save, publish.
3. Nothing to do on the Vercel side — `web/api/organise.js` sends the header
   whenever `N8N_ORGANISE_SECRET` is set, which it is.

Until step 2 is done, treat the webhook path as a secret: it is in this
repository, so anyone with read access to the repo can reach production.

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
