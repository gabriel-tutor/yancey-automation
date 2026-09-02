# Upperform Email Assistant

Source of truth for the eight n8n workflows running the Upperform Partners
inbox assistant and, since Phase 2, the SharePoint document organiser. The workflows in n8n are **generated from this repo** — edit
here, rebuild, redeploy. Do not hand-edit the Code nodes in the n8n UI; the
next rebuild overwrites them.

Client: Upperform Partners (Yancey Dawson) · Builder: GabDotCode / Lumtri
Spec: `~/Downloads/upperform-ai-agent/02-specs/2026-08-26-email-agent-spec-FINAL.md`

## The invariant

**The agent cannot send email.** The Entra app is never granted `Mail.Send`.
Replies are created with Graph `createReply`, which returns a draft, and the
daily digest goes out over SMTP from a separate sender precisely so this stays
true. Three things in this repo exist only to protect that promise — the
`pricingGuard`, the `voiceGuard`, and the SMTP digest path. Don't simplify them
away.

## Layout

| Path | What it is |
|---|---|
| `src/agent-lib.js` | Every pure function the Code nodes use. 80 unit tests. |
| `src/prompts.js` | Classification + drafting prompts and the JSON schema. |
| `src/*.test.js` | Vitest suite, including bundle-parity tests. |
| `workflows/*.mjs` | Generators that emit n8n Workflow SDK source. |
| `workflows/helpers-manifest.mjs` | Which helpers each Code node needs. |
| `scripts/bundle.mjs` | Tree-shakes the library down to each node's closure. |
| `scripts/build-workflows.mjs` | Writes `dist/*.sdk.ts`. |
| `scripts/golden-set.mjs` | Gate 1 / Gate 3 regression harness. |
| `context/business-profile.md` | Deploy to `/data/context/` on the VPS. |
| `config/deploy.json` | Live IDs (workflows, data tables, project). |
| `fixtures/golden/` | The regression set. See `fixtures/README.md`. |

## Commands

```bash
npm install
npm test            # 94 tests
npm run typecheck   # strict tsc over JSDoc-typed JS
npm run build       # regenerate dist/*.sdk.ts
npm run verify      # all three
OPENAI_API_KEY=sk-... npm run gate1
```

## Live workflows

Fifteen, and every one is generated from `workflows/*.mjs`. There are no
exceptions: a workflow in n8n with no generator here is scaffolding, and
scaffolding is archived once the answer it was built for is written down.

**Production** — tagged `prod`:

| Workflow | ID | Trigger | Tags |
|---|---|---|---|
| WF-A · Poll & Classify | `dSGzw1yk46JTHD3Z` | every minute | `prod` `email-agent` |
| WF-B · Draft Reply | `Rjpk5KuPopQCpCQL` | called by WF-A | `prod` `email-agent` |
| WF-C · Daily Digest & Follow-ups | `vCbgJoNqfuVOqWGH` | weekdays 07:30 CT | `prod` `email-agent` |
| WF-D · Read Attachments | `ZYV9V281C1Kojzpz` | called by WF-A | `prod` `email-agent` |
| WF-E · Error Handler | `XSKk6eHTFGqftrUM` | on failure | `prod` `email-agent` |
| WF-F · Organise Documents | `CC8jik1RGti4JXZk` | manual | `prod` `documents` |
| WF-T · Preflight (scope proof) | `1fCjIsw2JjVGkNJb` | manual | `prod` `email-agent` `diagnostic` |
| WF-P · SharePoint Permissions (read-only) | `c6gA3jbwVxR4cWoY` | manual | `prod` `documents` `diagnostic` |

**Dev sandbox** — tagged `dev`, pointed at the `gtutor` mailbox:

| Workflow | ID | Tags |
|---|---|---|
| WF-A · Poll & Classify · dev | `O6f4ck4gjG2KePBL` | `dev` `email-agent` |
| WF-B · Draft Reply · dev | `Ps0RYrfT3auUNwKw` | `dev` `email-agent` |
| WF-C · Daily Digest & Follow-ups · dev | `bbKJUs1WTX8kjixP` | `dev` `email-agent` |
| WF-D · Read Attachments · dev | `P2BfDd7Y782I13ZD` | `dev` `email-agent` |
| WF-E · Error Handler · dev | `E302rNL5D7NuBy5c` | `dev` `email-agent` |
| WF-T · Preflight (scope proof) · dev | `8Y3MEkP7A4kEiuGK` | `dev` `email-agent` `diagnostic` |
| WF-S · Seed Test Inbox · dev | `CNttEItjnRC4RkZZ` | `dev` `email-agent` `diagnostic` |

WF-F has no dev copy: there is no second SharePoint library to point one at, and
it changes nothing unless deliberately armed.

`WF-S · Seed Test Inbox · dev` plants eight fixtures, two of them carrying
attachments — a .docx, .xlsx and .pptx read in process, a PDF that goes to the
model for vision, and a signature image small enough that the triage skips it.
Those files are built by `src/office-fixtures.js` from readable text, not pasted
in as base64, and every one is round-tripped through the real extractors in the
test suite.

Two generators refuse to build for the wrong target, and they mirror each other:
`wf-s-seed.mjs` refuses production because it writes fabricated mail into
whatever mailbox the target names, and `wf-p-permissions.mjs` refuses any target
with no SharePoint credential, because a probe that cannot reach SharePoint
reports a refusal that means nothing. A build that skips one of these is working.

### The conventions, and why each one

**Names.** `WF-X · Title` in production, the same plus ` · dev` in the sandbox —
composed by the generator as `NAME + targets.<env>.nameSuffix`, never typed into
the n8n editor, because the next build renames it back. Sorting the list by name
puts a unit beside its own dev copy, which is how you notice one has drifted.

**Tags.** Exactly one environment tag and one system tag on everything, plus
`diagnostic` on anything that must never be scheduled or left published. `prod`
exists because production used to be identified by the *absence* of `dev`, and
no filter can express an absence. `phase-1` and `upperform` were removed: both
sat on every workflow, so neither selected anything.

**What gets archived.** Scaffolding goes once the answer it was built for is
written down, and anything that would now give a *wrong* answer goes regardless
— a probe that misleads is worse than no probe. But what creates a test
condition stays. Reading execution data replaced the four helpers that inspected
a run; nothing replaces a seeder, and the dev sandbox is not something to tidy
away.

**Nothing falls back.** Each target names its own workflow IDs and the build
throws if one is missing. There is no shared block to inherit from — a dev build
of WF-A once inherited the production WF-B and would have drafted into Yancey's
mailbox from the sandbox. `src/deployment.test.js` holds these rules, including
that no ID and no mailbox credential may serve two environments, and that the
seeder never appears under production.

Base URL: `https://n8n-wevu.srv1921545.hstgr.cloud`
Data tables: `agent_email_log_yancey` (`cKOjAskKEn5wYTuf`), `agent_error_log` (`Q6j7VEadubVstttP`)
SharePoint: `https://upperformpartners.sharepoint.com/sites/TaskManagement`, credential `We80s2s9pbxJWWp2`

## Changing behaviour

| You want to | Change |
|---|---|
| Tune the classifier's judgement | `context/business-profile.md` on the VPS. Loaded fresh every run — no redeploy. |
| Change the taxonomy or few-shots | `src/prompts.js`, then `npm run verify` and redeploy WF-A. |
| Turn drafting on | `draftingEnabled: true` in WF-A's **Config** node. Nothing else. |
| Change model, confidence floor, poll size, always-flag senders | WF-A **Config** node. |
| Reconsider model choice | See "Model choice" below before switching — the naming is misleading. |
| Change digest time, recipient, follow-up threshold | WF-C **Config** node + the Schedule Trigger. |
| Stop reading attachments | Clear `wfD` from the target in `config/deploy.json`. That is the whole off switch — WF-A only reads attachments when its target names one. |
| Let the organiser actually rename | `applyRenames: true` in WF-F's **Config** node, with `maxChangesPerRun` as the cap. It rests `false`, so a run reviews and changes nothing. |
| Stop the organiser reading documents | `readContents: false` in WF-F's **Config** node. The model then sees only names, folders and kinds, and nothing from inside a document leaves the tenant. Naming gets worse, not broken. |
| Change a guard rule | `src/agent-lib.js` + its test, then rebuild and redeploy. |

## Two deviations from the spec, deliberate

1. **n8n Data Tables instead of Postgres (spec §8).** The deployed n8n is the
   Hostinger Docker Manager catalog build, which is SQLite-backed and has no
   Postgres alongside it. Adding one means an extra container, an extra
   credential, and an extra thing to back up, for two tables of metadata. Data
   Tables are n8n's native persistence, survive restarts, and are visible in the
   UI. The column layout mirrors the spec's DDL, so moving to Postgres later is
   a straight port — `03-deployment/docker-compose.FALLBACK.yml` still describes
   that stack if row-level SQL analytics ever becomes the requirement.

2. **`Prefer: outlook.body-content-type="text"` on the inbox read.** Graph does
   the HTML-to-text conversion server-side, which is more reliable than regex
   stripping and costs fewer tokens. `htmlToText` is still in the library and
   still runs whenever Graph returns HTML anyway, so nothing is lost.

## Model choice

`gpt-5.6-luna` for classification, `gpt-5.6-terra` for drafting. Both are set in
the **Config** node of WF-A and WF-B respectively, and both LLM nodes carry a
stable `promptCacheKey` so the identical prompt prefix bills at the cached rate.

Do not swap these for a `-mini` or `-nano` tier on the assumption it is cheaper.
As of the July 30 repricing:

| Model | Input /1M | Output /1M |
|---|---|---|
| gpt-5.6-luna | $0.20 | $1.20 |
| gpt-5.4-mini | $0.75 | $4.50 |
| gpt-5.6-terra | $2.00 | $12.00 |

Luna undercuts the mini tier by nearly 4x on input. At this workload —
roughly 1,300 classifications and 220 drafts a month — the split above costs
about **$2.65/month**: ~$0.80 for classification, ~$1.85 for drafting.

Classification is where the volume is, so it gets the cheap model. Drafting is
low-volume and is the only client-visible output, so it gets the better one; the
~$1.70/month premium over running luna there is not worth the voice quality.

Neither node sets `temperature`. The 5.6 family rejects it outright —
`Unsupported parameter: 'temperature' is not supported with this model` — so the
spec's `temperatureClassify: 0` had to go. Determinism for classification comes
from the strict `json_schema` response format instead, which constrains the
output shape regardless of sampling. If you ever move to an older model that does
accept `temperature`, add it back deliberately.

The real arbiter is `npm run gate1`. If a cheaper model still clears 90%
agreement on Yancey's labelled set with zero flagged-and-drafted items, take it.
Decide with the harness, not with the price list.
