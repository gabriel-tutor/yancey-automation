# Upperform Email Assistant — Phase 1

Source of truth for the four n8n workflows running the Upperform Partners
inbox assistant. The workflows in n8n are **generated from this repo** — edit
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

| Workflow | ID | Trigger | Publish when |
|---|---|---|---|
| WF-A · Poll & Classify | `dSGzw1yk46JTHD3Z` | every 2 min | credentials attached + Gate 1 passed |
| WF-B · Draft Reply | `Rjpk5KuPopQCpCQL` | called by WF-A | credentials attached |
| WF-C · Daily Digest & Follow-ups | `vCbgJoNqfuVOqWGH` | weekdays 07:30 CT | after shadow mode |
| WF-E · Error Handler | `XSKk6eHTFGqftrUM` | on failure | before the others |
| WF-T · Preflight (scope proof) | `1fCjIsw2JjVGkNJb` | manual | never — diagnostic only |

Base URL: `https://n8n-wevu.srv1921545.hstgr.cloud`
Data tables: `agent_email_log` (`zGcPKOA90SMVEenJ`), `agent_error_log` (`Q6j7VEadubVstttP`)

Next steps are in [`docs/DEPLOY.md`](docs/DEPLOY.md).

## Changing behaviour

| You want to | Change |
|---|---|
| Tune the classifier's judgement | `context/business-profile.md` on the VPS. Loaded fresh every run — no redeploy. |
| Change the taxonomy or few-shots | `src/prompts.js`, then `npm run verify` and redeploy WF-A. |
| Turn drafting on | `draftingEnabled: true` in WF-A's **Config** node. Nothing else. |
| Change model, confidence floor, poll size, always-flag senders | WF-A **Config** node. |
| Reconsider model choice | See "Model choice" below before switching — the naming is misleading. |
| Change digest time, recipient, follow-up threshold | WF-C **Config** node + the Schedule Trigger. |
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
