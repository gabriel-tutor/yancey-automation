# Deploy — what is left to do

Everything that can be built without touching a secret is built. Four workflows
and two data tables are live in n8n and configured. What remains needs
credentials, a file on the VPS, and Yancey.

Work top to bottom. Steps 1–6 are yours alone; step 8 is the call with Yancey.

---

## 1. Mount the business profile on the VPS

The classifier loads Upperform's context from a file at runtime, so it can be
tuned without redeploying anything.

```bash
ssh root@31.220.53.70
mkdir -p /root/context
nano /root/context/business-profile.md    # paste context/business-profile.md
```

Then mount it. Find the project directory with `docker compose ls`, and under
the n8n service's `volumes:` add:

```yaml
      - /root/context:/data/context:ro
```

```bash
docker compose down && docker compose up -d
```

**Verify:** open WF-A → **Resolve Context** → Execute step. `profileSource`
must read `file`. If it reads `fallback`, the mount is wrong — the agent will
still run, on a reduced built-in profile, but Yancey's taxonomy and voice notes
will be missing and classification quality drops.

## 2. Create the three credentials

In n8n → **Credentials** → New. Names matter: the nodes already reference them.

| Name | Type | Notes |
|---|---|---|
| `Outlook – Dev (gtutor)` | Microsoft Outlook OAuth2 API | Client ID/Secret from the Entra app. Check the displayed **OAuth Redirect URL** reads `https://n8n-wevu.srv1921545.hstgr.cloud/rest/oauth2-credential/callback` before connecting. |
| `OpenAI – Upperform` | OpenAI | Key from Yancey's platform.openai.com account. |
| `SMTP – Digest Sender` | SMTP | The dedicated `agent@` sender. **Not** Yancey's mailbox — a separate sender is what keeps the no-send invariant true. |

Before connecting Outlook, confirm the Entra app's **Supported account types**
is "Accounts in any organizational directory and personal Microsoft accounts".
The single-tenant default fails with AADSTS50194 and is the single most common
way this step goes wrong.

## 3. Attach the credentials to the nodes

The HTTP Request nodes call Graph through `predefinedCredentialType`, so n8n
could not auto-assign them at creation. Open each and pick the credential:

- **WF-A:** Fetch Inbox · Tag in Outlook · Classify Email
- **WF-B:** Create Reply Draft · Write Draft Body · Read Original Categories ·
  Tag AI-Drafted · Delete the Empty Draft · Read Categories (Blocked) ·
  Tag AI-Review (No Draft) · Generate Draft
- **WF-C:** Fetch Open Drafts · Send Digest (SMTP)
- **WF-E:** Read Message Categories · Tag Message AI-Error · Alert the Builder

## 4. Check the model IDs

`modelClassify` and `modelDraft` in each **Config** node carry the values from
the spec (`gpt-5.6-luna`, `gpt-5.6-terra`). Once the OpenAI credential exists,
open **Classify Email**, switch the model field to list mode, and confirm those
IDs are actually available on Yancey's account. If they are not, pick the real
IDs and put them back in **Config** — every model reference reads from there.

## 5. Publish WF-E and WF-B, then wire the error workflow

```
publish WF-E (XSKk6eHTFGqftrUM)
publish WF-B (Rjpk5KuPopQCpCQL)
```

Both are safe to publish: WF-E only fires on a failure, WF-B only when called.
Neither has a schedule.

Then, on **WF-A, WF-B and WF-C**: Settings → **Error Workflow** →
`WF-E · Error Handler`. n8n refuses this until WF-E is published, which is why
it comes after. Without it, failures are silent.

## 6. Create the Outlook categories

In the dev mailbox (and later Yancey's), create these five categories. They are
the whole UI — the coloured flags are what he sees on day one.

`AI-Processed` grey · `AI-Action` orange · `AI-Review` red ·
`AI-Drafted` green · `AI-Error` purple

Graph will accept a category name that does not exist yet, but it will not show
a colour until the category is defined in the mailbox.

## 7. Run Gate 1 before pointing anything at a real inbox

Convert Yancey's 20 labelled emails into fixtures (format in
`fixtures/README.md` — pulling each one from Graph is exact and takes five
minutes), then:

```bash
OPENAI_API_KEY=sk-... npm run gate1
```

The harness runs the production prompt, schema and guard chain. It passes only
at ≥90% agreement, zero flagged-and-drafted items, every pricing case held, and
at least 20 fixtures from Yancey's own set. Six synthetic fixtures ship with the
repo; they exercise the harness but deliberately do not count toward Gate 1.

Mislabels are usually fixed in `context/business-profile.md` first — that needs
no redeploy. Only reach for `src/prompts.js` when the taxonomy itself is wrong.

## 8. Shadow mode, then the go-live call

1. Publish **WF-A** against the dev mailbox. `draftingEnabled` is already
   `false`, so it labels and logs but drafts nothing.
2. Watch three or four runs. `agent_email_log` should fill; the inbox should
   show colours.
3. Run the go-live call from `04-runbooks/setup-call-runbook.md`. Swap WF-A,
   WF-B and WF-C to the `Outlook – Production (Yancey)` credential first.
4. Two to three business days of labels only. Yancey reports mislabels; tune the
   profile file.
5. Flip `draftingEnabled` to `true` in WF-A's **Config**. Publish **WF-C**.
   Set `mode` to `live` in WF-C's Config so the shadow-mode banner disappears
   from the digest.

## Before handover

- [ ] Back up `N8N_ENCRYPTION_KEY`. Lose it and every credential above must be
      reconnected from scratch, including a second call with Yancey.
- [ ] Change the gtutor password — it arrived as a screenshot inside an email
      thread and is sitting permanently in two mailboxes.
- [ ] Confirm `EXECUTIONS_DATA_PRUNE=true` and `EXECUTIONS_DATA_MAX_AGE=168` so
      email bodies do not linger in execution history.
- [ ] Take a Hostinger snapshot once WF-A has run clean for a day.
- [ ] Screenshot the Entra permission list (read/draft, no send) for the
      handover doc — it is the evidence behind the promise.
- [ ] Put this repo under version control and hand it over with the credentials.

---

## What has already been verified

Run in n8n against the live instance, with credentialed nodes pinned so nothing
touched a real mailbox or wrote a row:

- **WF-C end to end.** Four log rows in, one draft present in the pinned Drafts
  listing. The row whose draft had vanished (`MSG-C`) was correctly detected as
  sent and pushed through `Split Resolved Rows` to `Close Resolved Rows`. The
  live draft stayed in "Drafts Waiting" at 2 business days. Buckets came out
  1 needs-you / 1 draft-waiting / 1 follow-up / 1 resolved.

- **WF-A end to end, including Gate 3.** Two messages in. A deep Outlook forward
  chain from Chuck trimmed to "FYI" plus the one forwarded message that carries
  the meaning, dropping the older `-----Original Message-----` section. A pricing
  request from a client contact was pinned so the model returned
  `opportunity / confidence 0.95 / requires_draft: true` — and `pricingGuard`
  overrode it to `flag_for_yancey`, `requiresDraft: false`,
  `guard_flags: pricing_guard`, categories `AI-Processed + AI-Review`.
  **That is the spec's Gate 3 holding in the deployed workflow, against a model
  that actively wanted to draft a reply.** `profileSource` read `file`,
  confirming the mounted-profile path.

- **94 unit tests + strict typecheck** over the quote trimmer, HTML converter,
  guards, voice guard, and digest renderer, including tests that the tree-shaken
  code actually shipped into each n8n Code node behaves identically to the module
  under test.

Still unproven until credentials exist: the Graph calls themselves, the OpenAI
calls, SMTP delivery, and OAuth. Those are steps 2–3 above.

## Open items still owed by Yancey

From the spec §13 and the client doc: exact signature block for first contacts,
booking link, digest time confirmation, follow-up threshold, and any senders
that should always be flagged. The first two are `signature` / `bookingLink` in
WF-B's Config; the last is `alwaysFlagSenders` in WF-A's Config, which accepts
either full addresses or `@domain.com`.
