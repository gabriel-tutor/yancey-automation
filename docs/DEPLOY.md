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

Done on 2026-08-26. The live compose is `/docker/n8n-wevu/docker-compose.yml`
(find it with `docker compose ls`), and the applied diff was:

```yaml
    environment:
      - TZ=${TZ}
      - N8N_RESTRICT_FILE_ACCESS_TO=/data/context   # added
      - EXECUTIONS_DATA_PRUNE=true                  # added
      - EXECUTIONS_DATA_MAX_AGE=168                 # added
    volumes:
      - n8n_data:/home/node/.n8n
      - /root/context:/data/context:ro  # added
```

The two pruning variables were missing entirely, so n8n was storing every
execution — full email bodies included — indefinitely. Spec §10 requires a
7-day window precisely so client mail does not linger on disk. A backup of the
original sits at `docker-compose.yml.bak-preprofile`.

**`N8N_RESTRICT_FILE_ACCESS_TO` is the one that catches people out.** The mount
alone is not enough. This variable used to be unset, meaning "allow everything",
but since **n8n v1.123.5 it defaults to `/home/node/.n8n-files`** — so a correctly
mounted `/data/context` is refused with:

```
Access to the file is not allowed.
```

The filesystem is fine at that point; n8n is the one saying no. `ls` from inside
the container will happily show the file, which makes this confusing to diagnose.
This VPS runs n8n 2.35.7, well past the change. Scoping the variable to
`/data/context` alone is also tighter than the old default, since the profile is
the only file these workflows ever read.

If you ever need more than one path, note the separator is unreliable: the docs
say colon, but Docker deployments generally need a semicolon.

Mount the *directory*, not the file. Editors write a new file and rename it on
save, which changes the inode; a single-file mount would leave the container
reading the old one forever.

**Verify:** open WF-A → **Resolve Context** → Execute step. `profileSource`
must read `file`. If it reads `fallback`, the mount is wrong — the agent will
still run, on a reduced built-in profile, but Yancey's taxonomy and voice notes
will be missing and classification quality drops.

### Access to the VPS

Root login is by SSH key. Gabriel's key is registered through hPanel → VPS →
SSH keys, labelled `gabriel-macbook-gabdotcode01@gmail.com` so Yancey can
identify and revoke it at handover. Host key fingerprint:
`SHA256:uShZrPIML5j8Me9d+rnHj0oDpIoilHYjmT8xng8k8SM` — if SSH ever warns that it
changed and nobody rebuilt the server, investigate before reconnecting.

Editing the profile afterwards needs no restart and no redeploy:

```bash
ssh root@31.220.53.70
nano /root/context/business-profile.md
```

WF-A reads the file fresh on every run, so the next 1-minute poll picks it up.
That is the tuning loop for shadow mode. Keep `context/business-profile.md` in
this repo in sync, or the handover copy drifts from what is actually running.

## 2. Create the three credentials

In n8n -> **Credentials** -> New. Names matter: the nodes already reference them
by name, and the dash in the Outlook names is an en-dash (–), not a hyphen.

| Name | Type |
|---|---|
| `Outlook – Dev (gtutor)` | Microsoft Outlook OAuth2 API |
| `OpenAI – Upperform` | OpenAI |
| `SMTP – Digest Sender` | SMTP |

The SMTP sender must be the dedicated `agent@` mailbox, **not** Yancey's. A
separate sender is what lets the Graph app stay read-and-draft only.

### 2b. SMTP via Resend — verified working

Do not point this at a Microsoft 365 mailbox. n8n's SMTP credential only speaks
basic auth, and Exchange Online disables SMTP AUTH basic auth by default from the
end of December 2026 — it would work at go-live and die around New Year, months
after handover, with no migration path available in the node.

Resend, confirmed sending on 2026-08-26:

| Field | Value |
|---|---|
| Host | `smtp.resend.com` |
| Port | `465` |
| SSL/TLS | on |
| User | `resend` (literally that word) |
| Password | the API key, `re_...` |

Ports 465/2465 are implicit TLS; 25/587/2587 are STARTTLS. 465 has the fewest
failure modes.

**Until a domain is verified**, Resend sends only from `onboarding@resend.dev`
and only *to* the account owner's own address. Anything else is refused with a
550 that helpfully names the permitted address:

```
550 You can only send testing emails to your own email address
    (gtutor@partnersperform.com).
```

That constraint is useful rather than annoying: it proves rendering, SMTP auth
and delivery end to end without touching DNS, which is how the first digest was
tested.

**Before go-live**, verify a **subdomain** — `agent.partnersperform.com`, sending
as `assistant@agent.partnersperform.com`. Not the root domain.
`partnersperform.com` already carries an SPF record for Microsoft 365, and SPF
permits exactly one record per domain; adding a second produces a permerror that
can degrade the firm's entire mail flow. A subdomain has its own SPF and DKIM and
cannot affect M365. Those DNS records are a client dependency — raise them early,
not on the day.

Then restore the real values: WF-C Config `digestFrom` / `digestTo`, and WF-E
Config `alertFrom` / `alertTo`. Both currently hold the sandbox address, so error
alerts would not reach you as things stand.

### 2a. The Outlook credential — the scope trap

**Read this before clicking Connect.** n8n's Microsoft Outlook credential ships
with a hidden default scope string:

```
openid offline_access Contacts.Read Contacts.ReadWrite Calendars.Read
Calendars.Read.Shared Calendars.ReadWrite Mail.ReadWrite Mail.ReadWrite.Shared
Mail.Send Mail.Send.Shared MailboxSettings.Read
```

`Mail.Send` and `Mail.Send.Shared` are in there. The Entra app uses dynamic
consent — the scopes requested at sign-in are the scopes granted, regardless of
what the app registration lists. Accept the defaults and Yancey's permission
screen will say **"Send mail as you"**, which is precisely the thing you will be
narrating as *not* on the list. It breaks the project's core promise, live, in
front of the client.

n8n has the fix, but it is off by default:

Creating the credential from inside a node gives it the default name
*"Microsoft Outlook account"*. Rename it first — the Graph nodes in all four
workflows were generated referencing `Outlook – Dev (gtutor)` by name, so
matching it saves hand-picking the credential on roughly fourteen nodes.

1. Rename the credential to `Outlook – Dev (gtutor)` (en-dash, not a hyphen).
2. Confirm the **OAuth Redirect URL** n8n displays reads
   `https://n8n-wevu.srv1921545.hstgr.cloud/rest/oauth2-credential/callback`.
   If it shows `localhost:5678`, stop — `N8N_HOST` / `WEBHOOK_URL` are wrong and
   OAuth cannot work until they are fixed. The Authorization and Access Token
   URLs should both use `/common/`, the multitenant endpoint.
3. Paste the **Client ID** and **Client Secret** from the Entra app. The secret
   is the **Value** column, not the Secret ID next to it. Leave
   **Authentication** on `Client Secret`.
4. Change **Allowed HTTP Request Domains** from `All` to `Specific Domains` and
   enter `graph.microsoft.com`. This restriction applies only when a credential
   is used from an HTTP Request node — which is every Graph call these workflows
   make. Left on `All`, a mistyped URL in any HTTP node would send Yancey's
   mailbox bearer token to whatever host was typed. It is a useful backstop, not
   a primary control: n8n has published advisories about bypasses of it, so the
   scope list in the next step remains the real guarantee.
5. Turn on **Custom Scopes**.
6. Set **Enabled Scopes** to exactly:

```
openid offline_access User.Read Mail.Read Mail.ReadWrite
```

7. Save, then **Connect my account**.

Steps 5 and 6 must happen **before** Connect. Consent is granted at sign-in and
persists as an enterprise-app grant in the tenant until revoked, so connecting
once with the defaults means reconnecting *and* clearing the stale grant under
Entra -> Enterprise applications -> the app -> Permissions.

That set covers every call these workflows make — reading the inbox,
`createReply`, PATCHing a draft body, setting categories, deleting a blocked
draft, listing the Drafts folder — and nothing else. `Mail.ReadWrite` grants
draft creation but not sending. This credential type defines no test request, so
trimming the scopes will not cause a spurious connection failure.

**What the consent screen should show:** read and write access to mail, sign-in,
and "maintain access to data you have given it access to" (`offline_access`).
No send. Screenshot it — the handover doc wants it as evidence.

Also verify on the Entra side before connecting:

- **Supported account types** = "Accounts in any organizational directory and
  personal Microsoft accounts". The single-tenant default fails with
  `AADSTS50194`. Fix under Authentication, or `signInAudience` in the Manifest.
- **Authentication -> Web -> Redirect URIs** contains the callback URL above
  exactly: https, no trailing slash, no stray whitespace.

If your n8n build has no **Custom Scopes** toggle, do not proceed with the
defaults. The fallback is a Generic OAuth2 credential with the scope string
above, which also means switching the Graph HTTP nodes from
`predefinedCredentialType` to `genericCredentialType`.

At go-live this is repeated as `Outlook – Production (Yancey)` — same client ID
and secret, same custom scopes, connected from Yancey's browser on his own
machine. Then swap the credential on the nodes in WF-A, WF-B and WF-C.

## 3. Attach the credentials to the nodes

The HTTP Request nodes call Graph through `predefinedCredentialType`, so n8n
could not auto-assign them at creation. Open each and pick the credential:

- **WF-A:** Fetch Inbox · Tag in Outlook · Classify Email
- **WF-B:** Create Reply Draft · Write Draft Body · Read Original Categories ·
  Tag AI-Drafted · Delete the Empty Draft · Read Categories (Blocked) ·
  Tag AI-Review (No Draft) · Generate Draft
- **WF-C:** Fetch Open Drafts · Send Digest (SMTP)
- **WF-E:** Read Message Categories · Tag Message AI-Error · Alert the Builder

## 3b. Prove the no-send invariant — mandatory gate

Run **WF-T · Preflight (scope proof)** (`1fCjIsw2JjVGkNJb`) before pointing the
agent at any mailbox, and again after every credential change.

Set `expectedMailbox` in its **Config**, then Execute. Four checks confirm the
credential can read the inbox, list drafts, create a draft and delete one. The
fifth deliberately calls `POST /me/sendMail` at the mailbox itself and requires
Microsoft to refuse it.

**Do not proceed unless `sendCheck` reads `PASS`.**

| `sendCheck` | Meaning |
|---|---|
| `PASS` (401/403) | Sending is genuinely denied. The invariant holds. |
| `CRITICAL` (2xx) | The credential holds `Mail.Send`. Stop and remediate below. |
| `INCONCLUSIVE` | Graph rejected the request shape before checking permissions. Proves nothing. |

Screenshot the Verdict output for the handover doc. A consent screenshot shows
what was *asked for*; this shows what was actually *granted*.

### If it comes back CRITICAL

This happened on the first real run against the dev mailbox on 2026-08-26, and
it is the expected failure mode, so treat it as likely rather than exotic.

**Trimming the scope string does not retract a grant that already exists.** If
the app was ever connected once with n8n's defaults — a single exploratory
connect is enough — Entra records `Mail.Send` on the service principal, and the
token endpoint can keep handing it back no matter what the new authorize request
asks for.

1. Entra -> **Enterprise applications** -> the app -> **Permissions** -> revoke
   the consented permissions. Deleting the enterprise application (the service
   principal) also works and is cleaner; the app *registration* is unaffected.
2. In n8n, confirm **Custom Scopes** is on and **Enabled Scopes** is exactly
   `openid offline_access User.Read Mail.Read Mail.ReadWrite`. **Save.**
3. Reconnect. With the grant revoked you get a fresh consent screen — read it.
4. Re-run WF-T and require `PASS`.
5. Delete the `PREFLIGHT — this must never arrive` message from the inbox.

Note that the agent's own code never calls `sendMail`, so a `CRITICAL` result
does not mean anything was sent by the assistant. It means the *permission*
exists, which makes "it could not send even if it tried" untrue — and that
sentence is the product.

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

Do this by hand. Creating categories through Graph would need
`MailboxSettings.ReadWrite`, which step 2a deliberately leaves out.

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

### Live run against the dev mailbox, 2026-08-26

WF-A executed end to end against gtutor's real inbox with real Graph calls and
real `gpt-5.6-luna` classification. Five messages in, five classified, and the
categories are visible in Outlook:

| Message | Label | Categories applied |
|---|---|---|
| PREFLIGHT — this must never arrive | flag_for_yancey | AI-Processed, AI-Review |
| AI Agent Status Meeting | admin | AI-Processed, AI-Action |
| There's more to explore in Microsoft 365 | noise | AI-Processed |
| Power up your productivity with Microsoft 365 | noise | AI-Processed |
| Email and Teams | admin | AI-Processed, AI-Action |

Both Microsoft promotional mails were correctly called `noise` and marked
`resolved` at insert, which is what keeps them out of the digest. Neither `admin`
item was marked draftable — the model judged no reply was warranted, which is the
right call on a meeting notice.

Two defects that only a live call could have surfaced, both now fixed:

- **`temperature` is not supported by gpt-5.6.** The run failed outright with
  `Unsupported parameter: 'temperature' is not supported with this model`. The
  spec's `temperatureClassify: 0` is gone; determinism comes from the strict
  `json_schema` response format instead.
- **CRLF line endings survived in text bodies.** Only the HTML path normalised
  them, so `\r` characters were being shipped to the model. Fixed in
  `normalizeGraphMessage` with a test.

### Digest delivered, 2026-08-26

WF-C ran end to end and a real digest arrived: Drafts folder read from Graph,
log rows bucketed, HTML rendered, delivered over Resend SMTP. Subject read
`Inbox digest — 2 need you, 0 drafts waiting`, with the shadow-mode banner, the
two flagged items, empty-state copy under Drafts Waiting and Follow-ups Due, and
the no-send promise in the footer. Timestamps rendered in Central.

The digest content was correct on the first attempt; only the recipient was
wrong, and Resend's 550 named the right one.

### End-to-end validation on genuinely delivered mail, 2026-08-26

Four emails sent from an external Gmail account to the dev mailbox, with WF-A
published and running on its 1-minute schedule. No manual runs, nothing pinned.

| Email | Label | Conf | Draft | guard_flags |
|---|---|---|---|---|
| Following up on the T&D reliability conversation | opportunity | 0.97 | written | — |
| Scoping the damage prevention review | flag_for_yancey | 0.98 | none | `pricing_guard` |
| Reduce your software spend in 2027 | noise | 0.99 | none | — |
| MSA amendment — indemnification language | flag_for_yancey | 0.99 | none | `sensitivity_guard` |

The first two are the same email; the second merely asks what he charges. One
drafted, one refused. Both safety rails fired independently — `pricing_guard` on
the fee request, `sensitivity_guard` on the attorney and breach-of-contract
language — on mail that genuinely arrived over SMTP rather than seeded fixtures.

Latency from arrival to classification was 43s, 112s and 39s.

The written draft cleared the Voice Guard with no hard, medium or soft flags, and
opened `Hi Gabriel,` rather than `Hey` — the profile reserves `Hey` for known
colleagues and the sender was external, which is the greeting rule working
without being prompted for it.

Two design choices were validated incidentally. Gmail delivered HTML, but
`Fetch Inbox` received clean text, so the `Prefer: outlook.body-content-type`
header does the conversion server-side on third-party mail. And runs where
nothing was new finished in ~0.3s with no model call, because both idempotency
gates dropped everything before it reached the classifier. And note `profileSource` came back
`fallback` on every message — step 1 has not been done, so all of the above was
classified on the reduced built-in profile, not Yancey's real one.

## Gate 1 + Gate 3 — PASSED on Yancey's own labelled mail (2026-08-27)

Yancey's 20 training emails were retrieved from Gmail and converted into
fixtures. His subject-line labels split 10 OPPORTUNITY / 5 JUNK / 5 MY WRITING
STYLE. Only the first 15 are classification cases; the writing-style five are
mail he **wrote**, so they live in `fixtures/voice/` and calibrate the drafting
voice instead. `REQUIRED_REAL_FIXTURES` was corrected from 20 to 15 to match —
padding the count with outbound mail would have measured nothing.

Final run, `gpt-5.6-luna`, 21 fixtures (15 real + 6 synthetic):

```
Did the right thing 12/15   (+3 surfaced-not-drafted)
Missed real work   0
Drafted junk       0
Label agreement    100.0%
Gate 3 pricing     2/2 held
GATE 1 + GATE 3: PASS
```

All five JUNK items classified `noise` at >= 0.99, including the hardest one in
the set: a named Paychex representative, warm-introduced by Yancey's own
accountant, proposing a specific meeting time, in a thread Yancey had already
replied to twice. Nothing on its surface says junk. It is junk because
retirement-plan sales is not Upperform work, and that distinction is now the
first question the classifier is told to ask.

### The defect this found

The first run failed. `opp-05` ("Re: Pulse Power BI") was drafted when it should
have been held. NIPSCO had asked *"are you able to provide an estimate of
timeline and cost?"*, but by the time Chuck replied *"let's write a proposal of
what it would take"* the fee ask was one level down the quoted chain.
`pricingGuard` was reading `cleanText` — the quote-stripped newest message — so
it never saw the ask, and the agent wrote a reply about a proposal. This is the
exact failure mode Gate 3 exists to prevent, and only real mail surfaced it: the
synthetic pricing fixture puts the ask in the newest message, so it passed
throughout.

Fixed by giving the guards the whole thread:

- `normalizeGraphMessage` now returns `threadText` (post-footer-strip,
  pre-quote-strip, capped at 12000 chars). It is used by the deterministic
  guards only and is never sent to the model.
- `normalizeClassification` passes `ctx.threadText || ctx.cleanText` to
  `pricingGuard`. WF-A's Apply Classification node forwards it.
- `PRICING_ASK` gained the unambiguous commercial terms: scope of work,
  statement of work, SOW, level of effort, LOE, time and materials, fixed-fee,
  day rate, bid. Bare "proposal" was deliberately **not** added — one of
  Yancey's own emails is titled "DRAFT Proposal for Kickoff Meeting Deck" and
  is about a slide deck, not a fee.

Three regression tests cover it, including the assertion that the newest message
alone does not trigger while the full thread does.

Also fixed: the harness was still sending `temperature: 0`, which `gpt-5.6`
rejects. It had been removed from WF-A and WF-B but not from `golden-set.mjs`,
so the harness was not reproducing production. Removed. The consequence is that
sampling cannot be pinned — labels have been stable at 100% across runs, but the
draft / no-draft split moves by an item or two. Judge a run by the gate outcome.

### Note on `safe-hold`

Three items Yancey labelled OPPORTUNITY were surfaced without a draft. That is
not a miss. `opp-02` is a status update addressed to Chuck with Yancey cc,
`opp-07` is a colleague answering Chuck rather than asking Yancey anything, and
`opp-09` is a review addressed to Kathleen with Yancey cc. Each appears in the
digest; the agent simply declined to put words in his mouth where no question
was aimed at him. Both the harness and `fixtures/README.md` score these as
passing.

## Deploying the guard fix to the live instance (2026-08-27)

Profile v3 went to `/root/context/business-profile.md` on the VPS (v2 backed up
alongside as `.bak-v2`) and is visible inside the container at
`/data/context/business-profile.md`. No restart needed — WF-A reads it fresh on
every run.

The code fix needed WF-A's `Prepare Email` and `Apply Classification` nodes
replaced. There is no n8n API key on this instance, so the MCP `update_workflow`
tool was the only route, and it takes the node body inline.

**Verify before publishing.** `update_workflow` edits the draft; the published
version keeps serving until `publish_workflow` is called. That gap is the safety
margin — use it.

The first paste of `Prepare Email` was wrong. Hashing the deployed node against
the local build caught it:

```
Prepare Email          411a8e74...  <- deployed
Prepare Email          2f186ecd...  <- local build
```

The cause was a **literal non-breaking space** inside `htmlToText`:

```
local:     out.replace(/\xc2\xa0/g, " ")   <- U+00A0, normalises NBSP
deployed:  out.replace(/\x20/g, " ")       <- plain space, a silent no-op
```

The invisible character did not survive being retyped, and nothing would have
flagged it: the code is valid JS, n8n reported no validation warnings, and a
test execution passed because the idempotency gate skipped every already-labelled
message so the node never ran. HTML mail would simply have kept its
non-breaking spaces, quietly degrading quote-boundary detection.

Re-sent with `\u00A0` written as an explicit escape so the character cannot be
dropped again.

**Hashes are the wrong check here** — esbuild emits `"\u2019"` while a paste
round-trips to `"’"`, which is the same string. Verify behaviour instead. The
deployed helpers were pulled back out of the SQLite database, `eval`'d, and run
against `src/agent-lib.js` on the same inputs:

```
PASS  htmlToText normalizes &nbsp; (the regression)
PASS  entity table parity
PASS  normalizeGraphMessage full parity
PASS  threadText keeps the fee ask / cleanText does not
PASS  pricingGuard: newest message alone -> false, whole thread -> true
PASS  deployed guard chain holds opp-05
PASS  deployed == local classification
DEPLOYED CODE MATCHES LOCAL BEHAVIOUR
```

Only then published. `workflow_entity` now shows `active: 1` with
`versionId 4c415e4e-2ced-40d1-bed7-385306d7da6a`, matching the `activeVersionId`
returned by `publish_workflow`. Note `workflow_published_version` is empty on
this instance — `workflow_entity` is the live row.

Reading the node code back out for comparison:

```
docker exec n8n-wevu-n8n-1 node -e "
const {DatabaseSync}=require('node:sqlite');
const db=new DatabaseSync('/home/node/.n8n/database.sqlite',{readOnly:true});
const row=db.prepare('select nodes from workflow_entity where id=?').get('<id>');
for (const n of JSON.parse(row.nodes)) console.log(n.name, n.parameters.jsCode?.length);
"
```

## Live run of the demo set (2026-08-27)

All five demo emails sent from Gmail into the dev mailbox and picked up
automatically on the 1-minute schedule. Every outcome matched what
`docs/demo-emails.md` predicted.

| Subject | Label | Conf | Draft | guard_flags |
|---|---|---|---|---|
| RE: Northern district ILI extension | flag_for_yancey | 0.99 | no | **pricing_guard** |
| RE: Integrating Gas Control into T-Week Planning | flag_for_yancey | 0.98 | no | *(none)* |
| Introduction — business insurance review | noise | 0.99 | no | — |
| Damage Prevention workbook — RACI follow-up | opportunity | 0.99 | **yes** | — |
| Human reliability programme | opportunity | 0.97 | **yes** | — |

Arrival to classification: 62–71 seconds. Drafts written 3–6 seconds after that.

The first two rows are the ones worth showing. Both are held with no draft and
look identical in Outlook, but they were caught by different mechanisms. The
ILI thread tripped `pricing_guard` — its own reasoning field talks only about
developing an approach and tone, because the model saw nothing about money in
the newest message and the deterministic guard overruled it after reading the
chain. The Gas Control thread has an empty guard column and reasoning that
names "pricing and scope commitments" outright: the guard stood down because a
figure already existed in-thread, and the model held it on its own judgement.

Both drafts passed the Voice Guard with no hard, medium or soft flags.

The RACI reply is the one to read closely. Jennifer asked whether cross bore
remediation should sit with Construction or Integrity. The draft does not
answer:

> I'll take a look at the cross bore remediation RACI assignment with Adam's
> input in mind, along with where the QA/QC process document should land
> against the gap analysis.
>
> I'll come back to you before the working session on the 9th so we can get
> those items squared away.

It picked the date out of the inbound message, opened `Hi Jennifer,` rather than
`Hey` because she is a client contact, and took no position on a question only
Yancey can settle. 82 words.

The human reliability reply describes the offering accurately — practical human
performance tools, reducing error-likely situations, strengthening defenses —
which is real upperform.com content, and correctly did **not** volunteer
case-study numbers, per profile rule 10. 112 words, no fee mentioned.

### Two false alarms while checking, both mine

`guard_flags` first came back empty on every row because the query used
camelCase; the columns are snake_case (`guard_flags`, `requires_draft`,
`sender`, `received_at`). And latency first computed as ~18000s because
`createdAt` and `followup_due_at` are stored without a timezone suffix, so
`new Date()` parsed them as container-local (America/Chicago) rather than UTC.
Neither is a defect. Worth knowing before reading these tables in anger:
follow-up on a Wednesday arrival correctly lands the following Monday.

## Closing the remaining gaps (2026-08-27)

Three things were still open. All three are now fixed in source, covered by
tests, and deployed.

### 1. An orphaned draft was left in the mailbox when the model failed

`Create Reply Draft` puts an empty reply in the mailbox *before* the model is
called. If `Generate Draft` failed after its three retries the workflow threw,
and the only node that deletes that empty reply sits on the blocked branch —
so the empty draft stayed in the mailbox. This had already happened once.

The obvious fix — routing the error output to the cleanup branch — does not
work, because every node on that branch reads `$('Voice Guard')`, and on the
failure path the Voice Guard never ran. The expressions would throw and the
draft would still be stranded.

What works instead needs no rewiring at all. `Generate Draft` now continues on
error, so the failure arrives at the Voice Guard as an item carrying `error`.
The Voice Guard recognises it and returns `ok: false` along with every
identifier the cleanup branch needs — and it already sourced those from the
trigger and `Compose Draft Prompt`, both of which have run by that point. The
existing `Draft Passed the Guard?` branch then does exactly the right thing.

The log distinguishes the two cases: `draft_blocked` is a decision,
`draft_failed` is a fault. `guard_flags` reads `draft_generation_failed`.

### 2. No interlock on which mailbox gets drafted into

`draftingEnabled` was the only thing standing between the agent and writing
replies into a live client mailbox. Re-pointing the Outlook credential at
Yancey — the whole plan at handover — would have started drafting immediately
if that flag was still true.

`Config.draftingMailbox` now names the mailbox drafting is authorised for, and
`Apply Classification` checks the message's own to/cc against it. Mail for a
different mailbox is still classified, labelled and logged; it just does not
draft, and records `mailbox_not_addressed` so the reason is visible.

No extra Graph call — the recipients are already on the message. The trade is
that mail reaching Yancey only by bcc or a distribution list will not draft.
That is deliberate: it fails towards labelling. Worth knowing that all twenty
of the emails Yancey labelled himself have him on to or cc, so his real
traffic is unaffected.

**At handover, `draftingMailbox` must be changed at the same moment as the
credential.** Leaving it blank blocks all drafting rather than allowing it.

### 3. Source and live had drifted on the alert address

`wf-e-errors.mjs` said `alertTo: gabdotcode01@gmail.com`. The Resend sandbox
only delivers to the account owner, so that address is accepted and then
dropped — a broken agent would have failed silently. The live instance already
had the working address; only source was wrong, which meant a rebuild would
have quietly broken alerting. Source now matches live.

### The node bodies are tested now

Both bugs deployed this week got in through the same gap: the Code node bodies
were only ever strings, so nothing executed them. `src/nodes.test.js` runs them
with stubbed n8n globals — `$`, `$input`, `$json`, `$itemIndex` — and asserts
the failure paths directly, including the model-failure case that produced the
orphan. `VOICE_GUARD` and `APPLY_CLASSIFICATION` are exported from the workflow
modules for this.

That change pulled `workflows/shared.mjs` into the typecheck graph for the
first time and surfaced four implicit-`any` errors. Annotated rather than
excluded.

### Deployment

Config first, then the interlock that reads it — deploying `Apply
Classification` against a Config with no `draftingMailbox` would have stopped
all drafting. Both node bodies came back byte-identical to the local build this
time (no invisible characters in them), and the deployed code was pulled back
out of SQLite and run through the same simulations before publishing. WF-B
published before WF-A, since WF-A calls it.

## Code review findings, and the regression it caught (2026-08-27)

A two-axis review (standards + spec) was run against `0db8db8...HEAD`. The spec
axis found a real regression in the change that had just been deployed.

### The pricing guard got weaker in the commit that strengthened it

Reading the whole thread for the fee **ask** was right. Reading it the same way
for the **suppression** was not. `pricingGuard` stands down when a figure
already exists in-thread, and that check was matching any currency anywhere in
12,000 characters of quoted history:

```
bare fee ask                     -> held        (correct)
same ask + quoted "$40.6B"       -> RELEASED    (wrong)
same ask + "500k" in a signature -> RELEASED    (wrong)
```

`$40.6B` is SCE capital-plan wording, verbatim in Yancey's own `opp-04`
training email. So a genuine fee request arriving in any Vector thread would
have been let through — Gate 3 briefly weaker than before it read threads at
all.

Fixed by making the stand-down narrower than the ask. `feeAlreadyQuoted`
requires a money figure (digits with a currency marker — `500k` alone no longer
counts) with fee-*stating* language within 200 characters: `fixed-fee`,
`indicative range`, `retainer`, `per hour`, `the fee covers`, and so on.
Yancey's own quote reads "$175K to $200K on a fixed-fee basis" — figure and
phrase inches apart. A capital plan has no such words near its numbers.

Five regression tests cover it, including both false-suppression cases.

### The reported agreement figure was flattering itself

Gate 1 printed one agreement number averaged across all 21 fixtures, six of
which are synthetic smoke tests written in-house. Spec §11 asks for agreement
with *Yancey's* labels. The gate now turns on his 15 alone and reports the
overall figure separately:

```
Agreement (Yancey) 93.3%  (14/15, target >= 90%)
Agreement (all)    95.2%  (incl. synthetic smoke fixtures)
```

The headline number previously quoted as 100% was the diluted one. The honest
figure for his set moves between roughly 93% and 100% run to run, because
`gpt-5.6` rejects `temperature` and sampling cannot be pinned.

### Documentation contradicted the code

`fixtures/README.md` described `expected` as "a diagnostic, not the gate" while
`golden-set.mjs` failed the run below 90% agreement against it. The README now
says what the code does, and states plainly that `opp-05` disagrees with
Yancey's own label on purpose.

### Accepted, not fixed

The drafting interlock is **scope creep** against spec §5, which conditions
drafting only on label and confidence. It is kept because it fails closed and
guards a real handover hazard, but it is recorded here as an addition to the
spec rather than an implementation of it.

## Standards-axis findings acted on (2026-08-27)

Two hard violations and four judgement calls. The hard ones are fixed; two of
the judgement calls are fixed; two are recorded as accepted.

**Guard composition — fixed.** The interlock hand-rolled comma-joining of
`guardFlags` in the WF-A node body, when every other guard pushes onto the
`guards` array inside `normalizeClassification`. That split ownership of
`requiresDraft` across two files, in a template literal `checkJs` never reads —
so the next guard would have needed edits in both. It now flows through
`ctx.authorised` like pricing and sensitivity do. Absent means "no mailbox
context", which is what the golden-set harness has; only an explicit `false`
withholds a draft.

**Second coercion idiom — fixed.** `String(x || '').toLowerCase().trim()` is
the file's idiom (`sensitivityGuard`, `recipientAddresses`). The new function
had invented its own.

**Speculative Generality — fixed.** `describeDraftFailure` carried a
`blockedByGuard` branch and a `hasDraftToDelete` field that never ran: the
guard path computes its own status inline, and the helper is only reached when
the model errored. Both removed, and the call site now uses `failed.draftIdEnc`
rather than reaching past the helper to `composed.draftIdEnc` — one source of
truth for the id that deletes the empty draft.

**Mysterious Name — fixed.** `mailboxMatches(entry, box)` compared a *recipient*
to a mailbox. Renamed `sameAddress`.

**Accepted, not fixed.** The status vocabulary (`drafted` / `draft_blocked` /
`draft_failed`) is spelled in four places, one of them inside an n8n expression
string that cannot import a constant, so a shared enum would only partly help.
And `VOICE_GUARD` / `APPLY_CLASSIFICATION` are the only two node-body constants
exported, purely so `nodes.test.js` can reach them; registering bodies in
`HELPERS` alongside their helper lists would be tidier and is worth doing when
the next node body needs a test.

**Two fidelity notes on the new harness**, worth knowing before trusting it:
`nodeRef` returns the same value from `first()`, `all()` and `item`, so the
`$itemIndex` pairing that `APPLY_CLASSIFICATION`'s own comment calls subtle is
never exercised. And the cc test uses `GTutor@PartnersPerform.com`, which
cannot occur in production because `recipientAddresses` lowercases first — the
case-insensitivity it proves is defensive, not load-bearing.

Re-verified after the refactor: 132 tests, typecheck clean, both deployed node
bodies simulated against the local build, and a live email through the whole
chain came back `opportunity` 0.98, drafted, no guard flags.

## Gate 2 — voice (2026-08-27)

The spec's third gate, and the one that was missing. §11:

> **Gate 2 — voice:** draft replies to 5 opportunity samples; Yancey rates
> voice fit 1–5; average ≥ 4. Automated pre-checks: contains "Thanks," +
> "Yancey"; no "I hope this email finds you well"; no invented digits (regex
> for currency/percent not present in the source); under 150 words.

Half of that is a human judgement and cannot be automated away, so
`npm run gate2` runs in two phases.

**Phase 1** drafts replies to five of Yancey's own OPPORTUNITY emails through
the exact prompt, model (`gpt-5.6-terra`) and voice guard WF-B uses, runs the
four named pre-checks, and writes `fixtures/voice/gate2-drafts.md` — a rating
sheet showing what each person wrote and what the assistant replied.

**Phase 2** reads the scores from `fixtures/voice/gate2-ratings.json` and gates
at an average of 4.

The five samples are chosen to spread the registers his voice has to cover,
not to be easy: Chuck asking his opinion (colleague, "Hey"), an SCE contact on
a live deck (client, "Hi"), PG&E business development (where salesmanship is
the failure mode), a thread where he is only cc, and a plan review with an
open question he cannot answer from the thread.

`voicePreChecks` is deliberately separate from `voiceGuard`. The guard both
judges *and repairs*, and its length flag sits at a looser 220 words; the gate
only reports, at the spec's own 150. Eight unit tests cover it.

First run, all five drafted, **0 pre-check failures and no guard repairs
needed** — the model produced the sign-off and avoided the forbidden opener
unaided.

### One thing to watch when rating

Draft 4 (`opp-09`) replies as though Yancey owns the task — "I'll download the
Teams version, remove all comments, and send it over to Ashley". In the real
thread that is Kathleen's job and Yancey is only cc. The reply is in his voice
and reads well, which is exactly why it needs a human score: an automated check
cannot tell that the right reply here is a shorter one that does not take over
someone else's action. That is the sort of thing a 2 or 3 is for, and the fix
would be a line in the profile about replying from cc.

### Running it

```
OPENAI_API_KEY=sk-... npm run gate2      # phase 1: draft + pre-check
# Yancey fills in fixtures/voice/gate2-ratings.json
npm run gate2                            # phase 2: score the gate
```

The rating sheet is also published as an artifact he can use on a phone:
**https://claude.ai/code/artifact/efd87e6e-1fd6-4bb6-8947-65c2ae9a4c46**

It shows each inbound email beside the reply, takes a 1–5 tap per draft, runs
a live average against the bar of 4, and emits the exact JSON body of
`fixtures/voice/gate2-ratings.json` to paste back. Regenerate the artifact
whenever the drafts are regenerated — a score has to refer to text the person
actually read.

The `opp-09` card carries a deliberate warning rather than presenting all five
as equally sound. Nudging the rater toward a defect the automated checks
cannot see is the point of having a human in this gate at all.

Drafts regenerate on each run unless `--keep` is passed, so a score always
refers to text somebody actually read.

## Two guard defects found by writing a demo (2026-08-27)

Building a second demo set and checking the predictions before handing them
over surfaced two more defects. Neither would have been found by the golden
set, because neither phrase appears in Yancey's twenty emails.

**"how much" satisfied the pricing rule by itself.** `pricingGuard` requires
`PRICING_ASK && PRICING_QUESTION` — a money noun *and* a question form, two
independent signals. "how much" was in both lists, so one phrase satisfied
both. A commercial-property pitch reading *"given how much Upperform has
grown"* was escalated to `flag_for_yancey`. It fails safe, but it is the
failure that erodes trust fastest: junk landing in the "needs you" bucket. It
is a question form, not a money noun, and now lives only in `PRICING_QUESTION`.

**Two sensitivity stems had never matched anything.** The pattern ends
`...|indemnif|...)\b/i`, and `indemnif\b` cannot match `indemnification` —
the only form anyone writes. Same for `whistleblow` against `whistleblower`.
Both were intended as stems and neither had ever fired. Now `indemni\w*` (short
enough to also catch `indemnities`, which `indemnif` would have missed) and
`whistleblow\w*`.

The earlier MSA test email did trigger `sensitivity_guard`, which masked this —
but it fired on `attorney` and `breach of contract`, never on the
indemnification clause the email was named for.

Gate 1 re-run after both fixes: PASS, no regression. Deployed and verified
against seven behavioural cases on the live node.

## Exercising the paths nothing had ever run (2026-08-27)

Two more defects, both in code that had never been run against real data.

### The digest contradicted itself

Running WF-C end to end for the first time produced a digest that said, in a
banner at the top:

> Shadow mode: the assistant is labelling only. No replies are being drafted.

and then, two sections below:

> **DRAFTS WAITING (6)**

The banner came from a hand-set `mode: shadow` string in Config that no longer
matched reality once `draftingEnabled` was turned on. Yancey would have read
that every morning. It now derives from the run: no banner is printed if any
draft exists, whatever the config claims.

### Every error alert ever written was blank

`agent_email_log`'s sibling table held two real failures. Both read:

```
workflow_name: "unknown workflow"
error_message: "Unknown error"
execution_id:  ""   node_name: ""   stack: ""
```

`Normalize Error` called `$input.all()` — but `Config` sits between the Error
Trigger and the normaliser, so `$input` was the config object and every field
fell back. The same class of mistake as the Voice Guard one: reading the wrong
node. It now reads `$('On Workflow Error')` through a tested
`normalizeErrorPayload`, and an empty payload says so rather than "Unknown
error", which reads like something the workflow actually produced.

### Proving it, rather than assuming

WF-E cannot be run manually — it only accepts an Error Trigger — and n8n fires
error workflows on production executions only. So a throwaway workflow was
created that throws on purpose, pointed at WF-E, published, run in production
mode, then unpublished and archived.

The result, where before every field was blank:

```
workflow_name: WF-X · Fire Drill (safe to delete)
execution_id:  463
node_name:     Fail On Purpose
error_message: deliberate failure to prove WF-E captures what broke [line 1]
stack:         Error: FIRE DRILL: ... at VmCodeWrapper ...
```

And the alert email was accepted by SMTP:

```
accepted: ["gtutor@partnersperform.com"]   rejected: []
response: 250 9ad843db-d08c-46e1-bc7c-60e745823fcd
```

That is the first time the safety net has been shown to work. If the agent
breaks in Yancey's mailbox, somebody now finds out, and the message says what
broke and where.

### WF-C is one commit behind live, deliberately

The digest fix is in source and tested, but WF-C's node in n8n still has the
old banner. WF-C is inactive, so nothing is affected — and go-live already
requires rebuilding and redeploying it, because `digestTo` has to change from
the dev mailbox to Yancey at the same moment. **Rebuild and redeploy WF-C as
part of that step; do not publish the node currently sitting there.**

## Gate 1 reaches 100% by fixing the last false positive (2026-08-27)

The remaining disagreement with Yancey was `opp-02`, Shawn's status update on
crossarm capitalization and vehicle idling. The agent held it with
`pricing_guard`; Yancey had called it OPPORTUNITY.

It was not a judgement call. `PRICING_QUESTION` matched **"send me your"** in
*"please send me your updates for the items you are working on"*, while
`PRICING_ASK` matched **"cost"** and **"estimate"** in a different message about
idling analysis. Two unrelated phrases, 107 characters apart, read together as a
fee request.

That is a design flaw rather than one bad term. The rule was *any money noun
anywhere* AND *any question form anywhere*, and over a 12,000-character thread
those co-occur constantly in consulting mail — the thread-wide scan amplified
it enormously.

Both halves now require **proximity**, the same technique already proven in
`feeAlreadyQuoted`. The window was measured, not guessed:

```
MUST FIRE  one sentence     closest = 19    can you provide <-> estimate
MUST FIRE  across a break   closest = 19    can you provide <-> estimate
MUST FIRE  buried chain     closest = 19    can you provide <-> estimate
MUST NOT   status request   closest = 107   send me your    <-> cost
MUST NOT   share the deck   closest = 122   Can you share   <-> budget
```

60 characters sits between them with margin on both sides.

Result, stable across three consecutive runs:

```
Agreement (Yancey) 100.0%  (15/15, target >= 90%)
Missed real work   0
Drafted junk       0
Gate 3 pricing     2/2 held
```

This 100% is earned rather than diluted: it comes from removing a false
positive, not from averaging Yancey's fixtures in with synthetic ones. The
figure to quote is **100% across three runs, having been 93.3% before this
fix** — and it is worth saying that sampling is not pinned, so a future run
could differ.

## Artifacts, and one more hand-transcription bug (2026-08-27)

All three artifacts republished to their existing URLs.

| Artifact | Audience | What it covers |
|---|---|---|
| [Inbox assistant](https://claude.ai/code/artifact/91a37a65-cdcf-4262-8dc8-06cfd778ddf9) | Yancey | What happens to one email, why it cannot send, how it scored |
| [How it decides](https://claude.ai/code/artifact/d5e8ef1e-39d9-4bb8-9983-54292ec8013b) | Either | The four verdicts, which rules overrule the model, six real emails traced |
| [Agent internals](https://claude.ai/code/artifact/54c05d0c-79d6-4815-b93a-1386b9172548) | Builder | Node-by-node sequence, config reference, failure modes |
| [Voice check](https://claude.ai/code/artifact/efd87e6e-1fd6-4bb6-8947-65c2ae9a4c46) | Yancey | Gate 2 rating sheet, generated from the harness |

`how-it-decides.html` is the one to reach for when the question is "why did it
do that?" — it walks the taxonomy, the rule chain in execution order, both
proximity windows with the measured numbers behind them, and six real emails
from Yancey's own inbox traced from arrival to Outlook category. Three of the
six are cases the system got wrong at some point today, shown with what changed.

### The client explainer was claiming something untrue

It said **"15 of 15 — it reached the same verdict you did on every one."** That
was written when Gate 1 reported 100%, but that figure averaged Yancey's fifteen
in with six synthetic fixtures. His own number at the time was 14/15. The claim
is true again now, having been earned properly, but it was not true when it was
published, and the page now says plainly that the score is measured rather than
promised and that it read 93% before the proximity fix.

### The voice-check artifact had a transcription bug of its own

Building it by copying the five drafts across by hand dropped opp-09's
`Thanks, / Yancey`. The page was showing Yancey a draft with no sign-off — and
inviting him to score it. He would have marked it down for a fault the system
never committed. The same failure mode as the non-breaking space in the node
paste: what gets retyped gets damaged.

So nothing is transcribed any more. `npm run gate2` writes
`fixtures/voice/gate2-drafts.json` beside the markdown sheet, and
`npm run gate2:artifact` injects it into the page, re-running the four Gate 2
pre-checks on the way and refusing to publish if any fail. Updating the rating
sheet is now two commands rather than a careful copy:

```
OPENAI_API_KEY=sk-... npm run gate2      # draft + pre-check, writes the JSON
npm run gate2:artifact                   # inject, re-check, ready to publish
```

## Swapping the mailbox to Yancey

### The shape of it

All **17 Outlook nodes across 5 workflows** point at one credential,
`gBDmBd3jHGBqqQjl`. Nothing repoints per node — re-authenticate that single
credential and every workflow follows the sign-in at once.

| Workflow | Outlook nodes |
|---|---|
| WF-A · Poll & Classify | 2 |
| WF-B · Draft Reply | 7 |
| WF-C · Daily Digest | 1 |
| WF-E · Error Handler | 2 |
| WF-T · Preflight | 5 |

Do **not** create a second credential and migrate node by node. A half-finished
migration leaves WF-A reading Yancey and WF-B writing into the dev mailbox,
which is far worse than losing the dev connection. Re-auth is atomic; node
migration is not.

### Phase 0 — fix the scopes, and prove it on your own mailbox first

This is the part that must not be done live in front of Yancey. `Mail.Send` is
currently granted, and WF-T reports `202 Accepted` on its send attempt.

1. n8n → Credentials → **Outlook – Dev (gtutor)**.
2. Turn on **Custom Scopes** and set exactly:
   `openid offline_access User.Read Mail.Read Mail.ReadWrite`
3. **Connect my account**, sign in as gtutor@, approve.
4. Run **WF-T** and read `sendCheck`. It must now be `PASS`, with a **401 or
   403** on the send attempt.

If it still returns 202, the tenant has a lingering consent grant: Entra →
Enterprise applications → the app → Permissions → remove the consent, then
repeat step 3.

**Do not go past this step until WF-T passes on the dev mailbox.** The whole
point is to find out the scope change works before Yancey's mail is involved.

### Phase 1 — the swap itself, on the call

5. WF-A Config: `draftingEnabled: false`, `draftingMailbox:
   YDawson@partnersperform.com`. Publish. Do this **before** the credential
   changes — the interlock fails closed, so if the credential moves first,
   drafting stops anyway, but this way nothing is ambiguous.
6. Same credential → **Connect my account** → Yancey signs in **on his own
   machine**. He reads the permission list before approving; sending is not on
   it. That is the moment the explainer tells him to check.
7. Rename the credential to *Outlook – Yancey (prod)* so nobody later assumes
   it points at the dev mailbox.
8. Run **WF-T** again — now against his mailbox. Must `PASS`.
9. Run **WF-A** once manually. Coloured labels appear in his inbox live. No
   drafts, because drafting is off.

### Phase 2 — shadow, 2–3 business days

He reports mislabels; the profile is tuned on the VPS and picked up on the next
poll with no restart. Spec §12 step 4.

### Phase 3 — turn drafting on

10. Confirm `draftingMailbox` is his address, then set `draftingEnabled: true`.
    Publish. Watch the first few drafts land before leaving it alone.

### Phase 4 — the digest

11. **Blocked on a verified sending domain.** The Resend sandbox only delivers
    to the account owner, so the digest physically cannot reach Yancey today.
    Verify `agent.partnersperform.com`, update `digestFrom`.
12. WF-C: `digestTo` → Yancey, and **rebuild and redeploy WF-C from source** —
    its node in n8n is deliberately a commit behind and still carries the false
    shadow-mode banner. Then publish.

### The two things that will bite

**`draftingMailbox` and the credential must move together.** The interlock
checks the message's own To/Cc against that address. Swap the credential and
forget the config and the agent labels Yancey's mail correctly but silently
never drafts — recorded as `mailbox_not_addressed` in `guard_flags`, so it is
diagnosable, but only if you know to look.

**`alertTo` stays `gtutor@partnersperform.com`.** Error alerts should reach the
builder, not the client. That address is also the only one the Resend sandbox
delivers to, which is why alerting works today and the digest does not.

## The dev sandbox (2026-08-28)

Yancey's mailbox is now live on the published workflows, so testing needs its
own everything. The build is target-aware: `npm run build` is prod,
`TARGET=dev npm run build` is the sandbox.

| Resource | prod | dev |
|---|---|---|
| Outlook credential | `gBDmBd3jHGBqqQjl` Outlook – Yancey | `vNJuWdSXhjzB0Jo1` Outlook – Dev (gtutor) |
| Email log | `zGcPKOA90SMVEenJ` | `JZP4Wzctro2gA3ya` agent_email_log_dev |
| Error log | `Q6j7VEadubVstttP` | `y7IZKnJoNHt6sNiq` agent_error_log_dev |
| WF-A | `dSGzw1yk46JTHD3Z` | `O6f4ck4gjG2KePBL` |
| WF-B | `Rjpk5KuPopQCpCQL` | `Ps0RYrfT3auUNwKw` |
| WF-C | `vCbgJoNqfuVOqWGH` | `bbKJUs1WTX8kjixP` |
| WF-E | `XSKk6eHTFGqftrUM` | `E302rNL5D7NuBy5c` |
| WF-T | archived | `8Y3MEkP7A4kEiuGK` |

### Duplicating in the UI points everything at production

n8n copies node parameters verbatim, so a duplicated workflow inherits the
live credential, the live data tables, the live sub-workflow id and the live
error workflow. `dev - WF-A` arrived calling **prod WF-B**, with
`draftingEnabled: true` and `draftingMailbox: YDawson@partnersperform.com` —
running it would have driven drafts into a client mailbox from a sandbox.

Two other traps in the same operation. Duplicates come back with
`availableInMCP: false`, so they cannot be edited over MCP until that is turned
on. And WF-A's two Outlook nodes arrived **unbound** rather than pointing at
prod, which is the safer failure but easy to miss.

An error workflow must be **published** before it can be set as one, so the
order is: wire dev WF-E, publish it, then wire the rest.

Verified afterwards by scanning every dev node for a prod credential, table,
sub-workflow id or error workflow: **no dev workflow touches production.**

### The build hardcoded a credential name, which had become dangerous

Generators emitted `newCredential('Outlook – Dev (gtutor)')`. Yancey's
credential was renamed to *Outlook – Yancey* and a **new** credential took the
old name — so rebuilding prod from source would have repointed his workflows at
the dev mailbox. The name now comes from the target config.

Worth knowing: `newCredential()` by name creates a *placeholder*, not a
binding. Credentials have to be attached by id with `setNodeCredential`.

### State

`dev WF-B` and `dev WF-E` are published because a sub-workflow and an error
workflow have to be. Neither self-triggers. `dev WF-A` and `dev WF-C` are
deliberately **unpublished** — publishing WF-A starts a 1-minute poll, and a
scheduled dev digest is not wanted. Run them manually.

`dev WF-A` has `draftingEnabled: false` and `draftingMailbox` set to gtutor, so
the interlock fails closed even if the flag is flipped without thinking.

## Production status, 2026-08-28

### Live on Yancey's mailbox

| | |
|---|---|
| WF-A · Poll & Classify | **active**, every minute |
| WF-B · Draft Reply | **active**, called by WF-A |
| WF-C · Daily Digest | **active**, weekdays 07:30 Central |
| WF-E · Error Handler | **active**, proven to deliver |

`draftingEnabled: true`, `draftingMailbox: YDawson@partnersperform.com`. It is
classifying his real mail and writing drafts he is sending — his Drafts folder
came back empty on the last check, which means the replies are being used.

### What is proven, and how

```
Gate 1  agreement with Yancey's own labels   100%  (15/15, three runs)
        real work missed                     0
        junk drafted                         0
Gate 3  pricing cases held                   2/2
        269 unit tests, typecheck clean
```

The error path was proven end to end by deliberately failing a throwaway
workflow: the row carried the workflow, execution id, failing node, message and
stack, and SMTP accepted the alert. Before that fix every error row read
"unknown workflow / Unknown error".

### Closed today

`mode` was still `shadow` on WF-C while drafting was on, so the daily digest
would have printed *"No replies are being drafted"* above a list of drafts, and
headed it *"(dev mailbox)"*. Both were wrong and both were client-facing. Now
`live`, labelled with his address.

### Three things are NOT closed

**1. `Mail.Send` is granted on Yancey's mailbox.** WF-T:
`noSendInvariantProven: false`, send attempt returns `202 Accepted`. A brand
new credential authorised from scratch inherited it too, which rules out the
n8n scope settings and points at **tenant-level admin consent on the Entra app
registration**. It needs an Entra admin on partnersperform.com to remove the
admin-consented `Mail.Send`, then both preflights re-run.

Nothing in the agent calls `sendMail` — WF-B only ever calls `createReply`, and
that is covered by tests. So the guarantee currently rests on the code rather
than on the permission being absent. That is a weaker promise than the client
explainer describes, and the explainer says so.

**2. The digest cannot reach Yancey.** `digestTo` is deliberately still
`gtutor@`. The Resend sandbox only delivers to the account owner, so pointing it
at him would 550 rather than arrive. Verify a sending domain
(`agent.partnersperform.com`), update `digestFrom`, then switch `digestTo`.

**3. The 2–3 day shadow period was skipped.** Spec §12 step 4 wants
classify-only on the live mailbox first, so he can correct mislabels before
anything is written in his name. Drafting was on from the swap. Not a defect —
but it was a side effect rather than a decision, and it is worth telling him
that is what happened.

### One known deviation

WF-C's `Build Digest` node in n8n is one commit behind source: the fix that
derives the shadow banner from the run rather than from config is not deployed.
`mode: live` makes the banner unreachable, so the defect cannot surface, and
deploying a 12,000-character node body to an active production workflow was not
worth the risk on the morning of a client meeting. Deploy it with the `digestTo`
change, which requires rebuilding WF-C anyway.

## Closing out Phase 1 (2026-08-28)

### The digest cannot reach Yancey — tested, not assumed

Pointed `digestTo` at him and ran it. Resend answered:

```
550 You can only send testing emails to your own email address
(gtutor@partnersperform.com). To send emails to other recipients,
please verify...
```

Reverted to `gtutor@` so the digest keeps working rather than failing every
weekday at 07:30. **This is a Resend account action, not a workflow change:**
verify `agent.partnersperform.com`, point `digestFrom` at it, then switch
`digestTo`. One config edit once the domain is live.

Until then Gabriel receives the digest and can forward it. It is correct in
every other respect — right mailbox label, no false shadow banner, only his
mail in it.

### Yancey's mailbox got a clean log

The log table still held 32 open rows from the dev-mailbox era. Two of them
were already appearing in his digest under "Needs you", with *Open in Outlook*
links pointing into a different mailbox, and the rest would have surfaced as
phantom follow-ups from Monday.

Production now writes to **`agent_email_log_yancey`** (`cKOjAskKEn5wYTuf`).
The old table is left intact as an archive of the build.

Drafting was switched off for the changeover so re-classification could not
produce duplicate drafts for mail he had already dealt with. It turned out not
to matter — the `Skip Already Labelled` gate reads Outlook categories, so his
already-tagged mail is not reprocessed at all. The new table therefore fills
with **new** mail from here on, which is the right behaviour. Drafting is back
on.

### A stale label worth knowing about

Prod WF-E's Outlook nodes still displayed `Outlook – Dev (gtutor)` while being
bound to the correct credential id. n8n resolves by id, so it was harmless —
but that name now belongs to a *different* credential, so reading the workflow
would tell you the error handler was pointed at the wrong mailbox. Relabelled.

The general lesson, twice over now: **n8n caches credential display names in
the node.** After renaming a credential, every workflow that used it shows the
old name until touched. Trust the id.

### Where Phase 1 stands

Live on Yancey's mailbox and working: classify, draft, error alerting. Gate 1
100% on his own labels, Gate 3 2/2, 269 tests, typecheck clean. An isolated dev
sandbox exists with its own credential, tables and workflows.

Three things remain, none of them code:

1. **`Mail.Send` is granted at tenant level.** Needs an Entra admin on
   partnersperform.com. Until then the no-send guarantee rests on the code
   rather than on the permission being absent.
2. **Digest delivery** needs a verified sending domain.
3. **Gate 2** needs Yancey's five voice ratings.

## Polling moved to one minute (2026-08-28)

Yancey asked in the meeting whether it could be faster. Measured before
answering, because "just lower the number" is not an answer:

| Interval | Average wait | Worst case | Execution data | Retained (7d) |
|---|---|---|---|---|
| 2 min | ~60s | 120s | 31 MB/day | 218 MB |
| **1 min** | **~30s** | **60s** | **62 MB/day** | **436 MB** |
| 30 s | ~15s | 30s | 124 MB/day | 871 MB |

Every poll stores 44 KB whether or not anything arrived, because `Fetch Inbox`
pulls the full `lookbackHours` window each time and the gates discard what has
already been seen. Disk is not the constraint — 92 GB free.

**The model spend does not change.** Classification only runs on genuinely new
mail, so a faster poll adds Graph calls (free, far below throttling limits) and
nothing else. Idle polls finish in 213 ms.

Confirmed live rather than assumed:

```
12:29:10 -> 60s
12:28:10 -> 60s
12:27:10 -> 60s
12:26:10 -> 120s   <- before the change
```

### If he wants genuinely instant

Polling is the wrong tool past this point. **Microsoft Graph change
notifications** push a webhook the moment mail arrives — seconds, event-driven,
no idle polling at all. n8n is already on a public HTTPS URL so it is feasible.
The real cost is lifecycle: mail subscriptions expire and must be renewed every
~3 days, there is a validation handshake on creation, and a slow fallback poll
is still wanted to catch anything missed. That is a Phase 2 piece of work, not
a config change.

Worth saying plainly to him: operationally, a draft appearing in 30 seconds
rather than 2 minutes does not change his day — he is not watching the inbox.
It changes how responsive the thing *feels*, which is a real reason, just not a
technical one.

### 30 seconds, if asked

Fine on disk, but pair it with `saveDataSuccessExecution: none` on WF-A. At
2,880 runs a day the execution list becomes unusable for debugging, drowned in
idle polls — which is the actual cost, not storage.

## Yancey's answers, 2026-08-28 — all five received and implemented

He replied to every open question and added one instruction nobody had asked
for. Everything below is deployed to the prod drafts.

| He asked for | Where it lives | State |
|---|---|---|
| Signature: `Yancey Dawson` | WF-B Config `signature` | done |
| No booking link | WF-B Config `bookingLink` empty | done — the empty value now produces an explicit prohibition in the draft prompt rather than silently dropping the rule, so the model cannot offer to "send a calendar link" |
| Digest 07:30 Central, Mon–Fri | WF-C trigger + workflow timezone | already correct; **verified live** — the trigger reports `America/Chicago (UTC-05:00)` and execution 1247 fired at 12:30:35 UTC |
| Follow-up after 3 business days | `followupDays` in WF-A and WF-C | already correct; `addBusinessDays` skips weekends |
| Always flag Chuck Coleman and anyone outside partnersperform.com | `config/deploy.json` → `policy` | done |

### The one that changes the product

> "Please always flag emails from Chuck Coleman and anyone outside of
> partnersperform.com without drafting a response."

A reply is now written only for a colleague at `partnersperform.com` who is not
Chuck and trips no topic guard. Measured against his own labelled set, that
takes drafting from 10 of 10 opportunities to **3 of 10** — Chuck sent six of
them, and one was external. He keeps all of the triage and most of the drafting
goes away. That is a deliberate, conservative posture, not a defect, and
widening it later is one line in `policy`.

**The external rule deliberately does not touch `noise`.** Ten of the
twenty-one fixtures are outside senders classified as junk; promoting those to
`flag_for_yancey` would put PayPal receipts and newsletters into the bucket he
reads at 07:30. Only mail that would otherwise have been *drafted* is promoted.

### Topic guards he named

`legal, contracts, HR, pricing, financial matters, client complaints, media
requests, confidential information, approval requests` — added as
`CLIENT_POLICY_TOPICS` in `src/agent-lib.js`, each reporting its own category
so the digest can say *why* something was held.

Pricing keeps its existing two-signal guard rather than firing on any money
word: it is the gate that caught opp-05, and external pricing asks are now
covered by the sender rule anyway. The financial vocabulary is deliberately
narrow — Upperform's day job *is* budgets, costs and operating expense, and
opp-04 says "sustain benefits" and "reducing operating expense" while being
exactly the kind of mail he wants drafted. Bare `benefits` and `expense` would
have broken it; there are regression probes for both.

### The contradiction he spotted

> "The assistant supposedly cannot send email, but the document says it will
> send you a summary email."

He was right to ask, and the old wording was sloppy in two ways. The digest goes
out over **SMTP through Resend**, a service with no access to his mailbox — the
Graph app is not involved in sending at all. And "cannot send email" claimed
more than has been proven: the Entra app still holds `Mail.Send` (see the open
item below), so the honest claim is that *no workflow sends from his mailbox*.
The digest footer now says exactly that.

## Verified live in the dev sandbox, 2026-08-28

Every rule was exercised against real mail delivered to
`gtutor@partnersperform.com` before anything was published to Yancey's
workflows. Seeded fixtures for the internal cases, and two genuinely sent
emails from an outside address for the external rule.

| Message | Sender | Result | Guard |
|---|---|---|---|
| crossarm workbook | scorkran@ (colleague) | opportunity, **draft written** | — |
| O&M framing | dyoes@ (colleague) | opportunity, **draft written** | — |
| Pulse rollout | **ccoleman@** | flag, no draft | `always_flag_sender` |
| change order | dyoes@ | flag, no draft | `sensitivity:approval` |
| ILI program | kris.murray@nisource.com | flag, no draft | `pricing_guard` |
| AGA follow-up | outside address | flag, no draft | **`external_sender`** |
| GridSignal newsletter | outside address | **noise**, not flagged | — |

Three of these are the ones that would have caught a careless implementation.

**The AGA follow-up is the behaviour change.** The classifier understood it
perfectly — *"a utility reliability leader is following up on Yancey's CAIDI
work and requesting a call to discuss a potential engagement approach"*, 0.98
confidence. Before 2026-08-28 that would have been drafted. The sender rule held
it instead, and it carries `AI-Review` rather than `AI-Action`.

**The newsletter proves the rule does not over-reach.** Also an outside sender,
also correctly understood, and it stayed `noise` with no guard and no
`AI-Review` category — so it never reaches the 07:30 digest.

**The O&M email proves the topic guards do not over-reach.** It says "reducing
operating expense" and "realizing benefits", one word away from the financial
and HR vocabularies, and it drafted normally.

The draft that was written, in full:

> Hey Shawn,
>
> Yes, I'll pull the crossarm workbook together for Thursday's Momentum Room and
> make sure the capitalization approach is laid out clearly for the SLT.
>
> I'll get you a draft ahead of the session so we can make any tweaks needed.
>
> Thanks,
>
> Yancey Dawson

45 words, his signature as he specified it, no booking link offered, voice guard
clean. `followup_due_at` came back **2026-09-02** — Friday plus three business
days with the weekend skipped.

### Two bugs the dev run found that had nothing to do with the new rules

1. **WF-S was bound to Yancey's live mailbox.** Its `Plant in Inbox` node used
   credential `gBDmBd3jHGBqqQjl`, whose cached label read "Outlook – Dev
   (gtutor)". That ID is *Outlook – Yancey*. Running the seeder would have
   written six fabricated emails into a live client inbox while its sticky note
   promised dev-only. `deploy.json` carried the same wrong mapping. Both fixed;
   the seeder now refuses to build for prod.
2. **`Tag in Outlook` throttles under burst.** Thirteen messages returned 429
   `ApplicationThrottled`; `retryOnFail` does not help because every request
   hits the ceiling together. WF-A dies at that node, so the log rows below it
   never get written and those messages never reach the digest — a burst after a
   quiet weekend would have read as the agent ignoring the inbox. Now serialised.

## Gate 1 after the policy change — and the defect it caught

**PASS at 100% agreement on Yancey's fifteen labelled emails, twice.**

```
Agreement (Yancey)  100.0%  (15/15)
Classifier vs YD    100.0%  (15/15 junk vs not-junk)
Missed real work    0        Drafted junk    0
Gate 3 pricing      2/2      Policy gates    9/9
```

The first run did **not** pass — 80% — and the failure was worth more than the
pass. Three of the five emails Yancey labelled JUNK are PayPal contractor
invoices titled *"Invoice from JEREMIAH CHRISTIAN PLACENTE ROSALES"*. The
classifier read all three correctly as `noise`; the new `financial` vocabulary
saw "invoice" and promoted every one of them into his 07:30 review bucket.

A 60% false-positive rate on his own junk mail, and precisely the outcome the
external-sender rule had been written to avoid — the same reasoning simply had
not been applied to the topic guards.

### The fix, and the distinction it rests on

The two guard families answer different questions, so they now behave
differently:

- **`sensitivityGuard`** — always-flag senders, plus legal, HR and press. A
  liability backstop. It deliberately overrules the classifier, because it
  exists for the case where the classifier is wrong. Unconditional.
- **`policyTopicGuard`** — Yancey's named categories. These describe mail he
  wants to *answer himself*. If the classifier already called something junk,
  there is nothing to answer and nothing to withhold. Applied only to mail that
  would otherwise have been replied to.

Both directions are now asserted: an invoice-titled junk email stays `noise`,
and a `noise`-classified message mentioning litigation is still flagged.

The second failure was bookkeeping — `synthetic-03` is from Chuck Coleman and
its expectation had not been updated with the real `opp-*` fixtures.

Stability: two consecutive runs, 100% both times. The only variance is how many
messages the model chooses to surface rather than draft, which moves in the
conservative direction and fails nothing.

## Live state, 2026-08-28 (second round)

Published and active: **WF-A**, **WF-B**, **WF-C**, **WF-E**. WF-A polls every
minute; WF-C writes the digest into Yancey's inbox at 07:30 America/Chicago,
Mon–Fri. Gate 1 + Gate 3 PASS at 100% on his fifteen labelled emails, twice, at
the new 0.80 floor. 436 tests, typecheck clean.

Verified in production (execution 1476): the digest carries a real SCE email
from James Vallecorsa held as *"outside partnersperform.com"*, and a
"Processed as noise" entry for a Quran Coleman message with the classifier's
reasoning attached so a misclassification would be visible.

### His six items

| # | Asked for | State |
|---|---|---|
| 1 | Chuck Coleman always flagged, never drafted | already live; unchanged |
| 2 | Hold anything commercial, **even if a price was quoted** | built — suppression removed, categorical term list added |
| 3 | Upperform scheduling vs personal admin vs vendor scheduling | built — explicit rule in the classifier prompt |
| 4 | "Processed as Noise" section for the rollout | built — each item shown once, then closed |
| 5 | Confidence floor 0.80 | built |
| 6 | AI-Action vs AI-Drafted, and draft failure | clarified **and fixed** — they were not mutually exclusive |

### On item 2, and where the implementation is narrower than his sentence

He named six things: pricing, fees, estimates, proposals, scope changes,
commercial terms. Four are matched literally. **`estimate`, `proposal`, `price`
and `cost` are deliberately not**, because in his own mail they carry innocent
meanings — "DRAFT Proposal for Kickoff Meeting Deck" is a deck, and opp-02's
"low to medium confidence estimate for the hypothesis target" is a statistical
estimate in an email he wants drafted. Matching them literally held two emails
from his own labelled set for no reason.

Those four stay behind the two-signal test that also needs a question form, and
his sentence is carried in full by the classifier prompt, so the semantic
judgement is made by the model and the literal list is the deterministic
backstop underneath it. Worth saying to him plainly rather than claiming the
rule is purely literal.

### On item 6, which turned out to be a defect

The categories were not mutually exclusive. A message classified `opportunity`
got `AI-Action`; if its draft then failed, WF-B added `AI-Review` without
removing `AI-Action`, so it carried both. That is exactly the confusion he was
asking about. Fixed: `AI-Action` is cleared whenever no draft is written.

The four categories now mean:

- **AI-Processed** — the assistant read it. Always applied.
- **AI-Action** — it needs Yancey, and the assistant is handling it.
- **AI-Review** — the assistant deliberately did not draft. Over to him.
- **AI-Drafted** — a reply is waiting in Drafts.

`AI-Action` and `AI-Review` never appear together. If a draft fails, the empty
draft is deleted, the message is tagged `AI-Review`, the row is logged
`draft_failed`, and WF-E raises an alert.

## Verified 2026-08-29

Two things checked rather than asserted.

**Send permission — still granted.** WF-T re-run today, execution 1514, against
the dev credential so nothing lands in Yancey's mailbox. Read, list, create and
delete all PASS; the deliberate send attempt returned **202 Accepted**, so the
verdict is `noSendInvariantProven: false`. The answer given to him is correct and
now rests on a run from today rather than on memory. The Entra app is the same
for both credentials, which is why the dev test proves the point.

**Production is running what the repo says.** WF-A execution 1515 shows
`confidenceMin 0.8`, Chuck on `alwaysFlagSenders`, `internalDomains` set, and
`profileSource: file` — so the business profile mounted on the VPS is the one the
model reads, and it contains the corrected pricing rule. Source, build, deploy
and runtime all agree.

## Why so few drafts — measured, not estimated

Ran the classifier over all ten of Yancey's OPPORTUNITY emails and captured
`asks_of_yancey` per message:

- **7 of 10 contain a genuine request directed at him.** The mail is actionable
  and the classifier reads it correctly.
- Of those, the three where a reply was judged genuinely warranted — opp-06,
  opp-08, opp-10 — are **all from Chuck Coleman**.
- The remaining four are status updates and discussion contributions where a
  reply adds nothing beyond an acknowledgement.

So the always-flag rule removes the entire drafting surface in that sample. Not
because the rule is wrong, but because Chuck is the person who assigns Yancey
work, and assigned work is what produces a reply worth writing. On his labelled
set those are the same fact.

**No code should change for this.** The lever already exists and is one line:
remove Chuck from `policy.alwaysFlagSenders` in `config/deploy.json`. His mail
would then draft only when it is not commercial, not legal, not an approval
request and the classifier is confident — every other gate still applies. That
is Yancey's decision, and it is easier to make after a week of watching.

## Client documents — all brought current 2026-08-29

They had drifted into actively misleading territory, which matters because
Yancey holds the links.

| Document | Was wrong about | Now |
|---|---|---|
| `how-it-decides.html` | 0.6 floor, the pricing carve-out he removed, no external or always-flag rule, no noise section | rewritten: nine gates in the order they run, plus the measured drafting finding above |
| `workflow-reference.html` | 0.6 floor, WF-C "not published", digest over SMTP, `digestTo` pointing at dev, missing `internalDomains` / `draftingAliases` / `showNoise` | current |
| `gate2-voice-check.html` | asked him to score five drafts the system would now never write | marked superseded, replaced with a real draft the assistant produced |
| `controls-and-definitions.html` | — | new, answers his 28 August email |

## Phase 2 — the document organiser, 2026-09-02

**SharePoint was never blocked.** The earlier conclusion — that the automation
could not reach it and needed `Sites.Selected` plus a tenant admin — came from
probing SharePoint endpoints with the *Outlook* credential. A mail credential
returning 403 on SharePoint proves that a mail credential cannot read
SharePoint. It was read as if it proved more. `docs/PHASE2-SHAREPOINT.md` and
`docs/access-request-sharepoint.md` now open by saying so; the access request is
withdrawn, unsent.

What is actually true:

| | |
|---|---|
| Credential type | `microsoftSharePointOAuth2Api` — its own type, not Graph |
| Permissions | `AllSites.Read` + `AllSites.Write` under **Office 365 SharePoint Online** (`00000003-0000-0ff1-ce00-000000000000`), on the Entra app already used for Outlook |
| Admin needed | **None.** Neither permission requires admin consent |
| Host | `upperformpartners.sharepoint.com` — not `partnersperform` |

**The `.default` trap is the one worth remembering.** The credential requests
`…/.default`, which means *whatever this app is already consented for*. A
permission added but not consented therefore yields a quietly weaker token
rather than an error, and reconnecting the credential picks up nothing new.
Naming the scope explicitly is what forces Entra to prompt:

```
openid offline_access https://upperformpartners.sharepoint.com/AllSites.Read https://upperformpartners.sharepoint.com/AllSites.Write
```

Two endpoints tell the two possible causes apart and are worth keeping:
`_api/web/effectiveBasePermissions` reports what *the token* can do;
`_api/web/getusereffectivepermissions(@u)` reports what *the person* is entitled
to. When the second is wider than the first, the scope is the limiter and no
amount of site permission will help.

**Do not press "Grant admin consent"** on the API permissions page. It grants
everything configured, including `AllSites.FullControl` — tenant-wide full
control, for a workflow that renames files.

### What WF-F does

Reviews the library, categorises every document, checks each name against
Yancey's convention, and reports where each file sits. It rests with
`applyRenames: false`, so a run produces a review and changes nothing.

It **proposes no moves**, deliberately. A move changes a path exactly as a
rename changes a name, so it breaks a Power BI refresh the same way — and a
`.pbix` keeps its source paths inside itself. Moving anything between the
numbered stage folders would assert how far work has got, which no file records.

Applied on 2026-09-02 after the plan was read: six renames, no failures. The
library now reads 21 scanned, 6 compliant, 0 outstanding, 15 refused with a
reason each, 1 unfiled. The record and the undo list are in
`fixtures/sharepoint/renames-applied.json`.

## Organising the n8n instance, 2026-09-02

Twenty-four workflows, two naming conventions, tags that selected nothing, a
duplicate, and eight pieces of scaffolding whose questions were already
answered. Fourteen remain.

**The duplicate.** Two workflows were called `WF-F · Organise Documents`. The
first, `ZEIPYYQTBsmOuRek`, was created at 02:01 and abandoned at 02:03; the one
that was built out is `qh29SRwChF2g9bYB`, created at 02:14. Identical names are
how the wrong one gets opened and edited, so the dead one is renamed
`zz · WF-F duplicate` until it is archived.

**Two naming conventions.** WF-A, B, C and E carried a `dev - ` prefix; WF-D, T
and the probes carried a ` · dev` suffix. The generators compose every name as
`NAME + targets.<env>.nameSuffix`, and `targets.dev.nameSuffix` is `" · dev"` —
so the prefixed four were drift from before that existed. All four renamed. The
suffix form also sorts a unit beside its own dev copy, which is how you notice
one has drifted from the other.

**`WF-S · Seed Test Inbox` carried no environment marker at all.** It plants six
fabricated emails into whatever mailbox its credential owns, and it read as
production. It is bound to the dev credential and its generator refuses to build
for production, so only the name was wrong — but until 28 August it really was
bound to Yancey's live mailbox, through a credential whose cached display name
said "Outlook – Dev (gtutor)" while the ID belonged to Yancey. Renamed
`· dev`, tagged `diagnostic`, and its description now names the credential ID
rather than its label.

**Tags selected nothing.** `phase-1` and `upperform` sat on every tagged
workflow, and production was identified by the *absence* of `dev`, which no
filter can express. Replaced with one environment tag (`prod` | `dev`), one
system tag (`email-agent` | `documents`), and `diagnostic` on anything that must
never be scheduled or left published. WF-D and WF-F had no tags at all and were
invisible to every filter.

**Descriptions that had gone wrong.** WF-A said it polls every 2 minutes; it has
polled every minute since 28 August. WF-C said the digest goes out over SMTP;
it has been written into the mailbox over Graph for days, which is a stronger
guarantee than the sentence claimed. Every dev copy carried production's
description verbatim, so four workflows pointed at the `gtutor` mailbox while
claiming to act on Yancey's.

### Two config faults found while doing it

**The seeder's ID was recorded under production.** `targets.prod.workflows.wfS`
held the dev seeder's ID and `targets.dev.workflows` had no `wfS` at all —
exactly backwards for the one workflow that must never touch a client mailbox.

**A missing ID fell back to production.** `wf-a-classify.mjs` read
`(TARGET.workflows && TARGET.workflows.wfB) || DEPLOY.workflows.wfB`, and the
comment directly above it described the bug that fallback causes: a dev build of
WF-A calling the production WF-B, drafting into Yancey's mailbox from the
sandbox. The fallback is gone, the build throws instead, and the top-level
`workflows` block — which duplicated the production IDs and existed only to be
fallen back to — is deleted.

`src/deployment.test.js` holds all of this: no ID and no mailbox credential may
serve two environments, the seeder never appears under production, there is no
top-level block to inherit from, every named workflow has a generator, and the
conventions are written into `config/deploy.json` rather than left to memory.

### Ten archived, and the rule kept rather than softened

Twenty-four workflows became fifteen, and every one is now generated from
`workflows/*.mjs`. There are no hand-built exceptions left.

**Archived (10):**

| Workflow | Why |
|---|---|
| `zz · WF-F duplicate` | abandoned after two minutes; nothing referenced it |
| `WF-S0 · SharePoint access probe · dev` | **would give the wrong answer today** — probed SharePoint with the Outlook credential |
| `WF-S9 · Find the Task Management site` | answered; the host is in `deploy.json` |
| `WF-S3`, `WF-S4`, `WF-S5`, `WF-S6` | inspected a run, a log table or a draft, from before execution data could be fetched directly |
| `Tool · Seed Attachment Fixtures · dev` | folded into WF-S |
| `Tool · Seed PDF Fixtures · dev` | folded into WF-S |
| `Tool · SharePoint Permissions (read-only)` | replaced by the generated WF-P |

The line between archiving and keeping is not how a workflow was built. It is
whether something else now does the job:

> Reading execution data replaced everything that **inspected** a run.
> Nothing replaces something that **creates the condition** to inspect.

Three were kept on that basis and then, rather than living as documented
exceptions to "everything is generated", were folded into the repo properly.

### The seeder now covers every path WF-D takes

`WF-S · Seed Test Inbox · dev` plants eight fixtures. Two carry attachments: a
.docx, .xlsx and .pptx that are read in process and never leave the server, a
PDF that goes to the model for vision, and a signature image below the reading
threshold that the triage skips unread. That is every branch of the attachment
code, exercisable from one workflow.

The files are **built, not pasted.** `src/office-fixtures.js` writes stored-method
ZIPs, a structurally complete single-page PDF with a resolvable cross-reference
table, and a 1×1 PNG — a few hundred bytes each, from text you can read in the
source. A pasted .docx once went in with the right length and a valid ZIP
signature and inflated to nothing, and it read exactly like a broken decoder
rather than a broken fixture. Every fixture is round-tripped through the real
extractors in the test suite, and the CRC-32 is checked against Node's own
implementation rather than a constant copied in by hand.

Verified live: eight messages planted, all `201 Created`, `hasAttachments` true
on both, `isDraft` false so they behave like received mail.

### WF-P, and the diagnostic that had to explain itself

`WF-P · SharePoint Permissions (read-only)` replaces the hand-built probe. Three
GETs, nothing written. It reports the gap between `effectiveBasePermissions`
(what the token can do) and `getusereffectivepermissions` (what the person may
do); when the second is wider, the OAuth scope is the limiter.

Running it caught something the hand-built version had hidden. The second
endpoint now returns **403**, because reading another user's entitlement needs
`EnumeratePermissions` and an ordinary site member does not hold it. The verdict
said `PERSON UNREADABLE`, which reads like a broken probe rather than a known
limit. It now names the status code and the reason, and the conclusion adapts:
when the token can write there is nothing to explain, and when the person cannot
be read the app scope is still the first thing to check, because it is the usual
cause.

### One more hazard class closed

`src/generated-ascii.test.js` builds every generator and asserts that every Code
node body is pure ASCII. `src/bundle.test.js` already did this for the shared
helpers and found a literal non-breaking space in a normaliser and the euro and
sterling signs in the guard that decides whether an invented fee reaches a
draft — but it covered only the helpers, not the part written directly in a
generator. That is where the next ones were: four em dashes in the seeder's
fixture subjects, three in WF-C including **the digest subject line Yancey reads
every morning**, and six in WF-T.

They were already written as escapes in the source, which looks safe and is not.
Inside a template literal JavaScript decodes `\u2014` at evaluation time, so a
literal em dash landed in the generated node anyway; `\\u2014` is what emits an
escape. Nothing but a check on the output would have caught it, because the
input looked correct. Verified afterwards that the digest subject still renders
as a real em dash, so nothing Yancey sees changed.

## Still open

Three things, none of them mine to close.

1. **`Mail.Send` is still granted** on the Entra app. Unused by anything — the
   digest is written, not sent — but a promise resting on the code being correct
   is weaker than one resting on the permission being absent. Needs a Microsoft
   365 administrator on `partnersperform.com`. It was going to be bundled with
   the SharePoint access request; that request is withdrawn, so **this is now
   the only administrator item**. **Yancey asked directly whether the
   send-permission test has passed: it has not.** WF-T still reports
   `noSendInvariantProven: false`.
2. **Gate 2 needs Yancey's ratings.** The samples were regenerated — the old
   five were all mail his rules now hold for review, so scoring them would have
   measured nothing. Four genuinely draftable fixtures replaced them and are
   pinned by `src/gate2-samples.test.js`. The sheet is
   `fixtures/voice/gate2-drafts.md`, scores go into
   `fixtures/voice/gate2-ratings.json`, then `npm run gate2`. Needs a 4.0
   average.
3. **The OpenAI key needs rotating.** It was pasted into a session to run the
   gates and was never written to this repo — verified against the working tree
   and the full git history. It lives only in n8n, encrypted under
   `N8N_ENCRYPTION_KEY`. Rotating it is the account owner's action.

### Also worth saying out loud

**Drafting volume is low, and Yancey should hear it from Gabriel rather than
notice it.** With his rules, on his own twenty-email training set, no draft is
written for any of the ten opportunities: six are Chuck, one is external, and on
the remaining three the classifier itself judged that no reply was warranted.
The triage is doing the work; the drafting is deliberately conservative.

**The client-facing artifacts show viewers a pinned version.** Updating them
updates the live page, but anyone holding an old link keeps seeing the version
pinned when it was shared. After the 2026-09-02 pass, re-share:
Controls and Definitions, How It Decides, Upperform Agent Internals, and the
Task Management Library Review.
