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

## Environment variables

Set these on the Vercel project (Settings → Environment Variables), not in the
repo:

| Name | Value |
| --- | --- |
| `N8N_ORGANISE_URL` | `https://n8n-wevu.srv1921545.hstgr.cloud/webhook/<path>` — the WF-G webhook's production URL |
| `N8N_ORGANISE_SECRET` | Optional today. See hardening below. |

The dev sandbox has its own path (`…-dev`) and its own Vercel deployment if you
want one. Its `canArm` is false, so it will read the library and refuse to change
it — and the page says so rather than pretending.

## Hardening: add Header Auth (two minutes, worth doing)

Today the webhook is protected by an unguessable path and by never being shown
to a browser. That is a first line, not the last one. To close it properly:

1. In n8n: **Credentials → New → Header Auth**. Name it `Organise webhook key`.
   Header name `X-Organise-Key`, value a long random string.
2. Open **WF-G → the `Asked` node** → Authentication → **Header Auth** → pick
   that credential. Save and publish.
3. In Vercel: set `N8N_ORGANISE_SECRET` to the same random string. Redeploy.

The proxy already sends the header when the variable is set, so step 3 is the
only change on this side.

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
