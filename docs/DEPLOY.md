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

WF-A reads the file fresh on every run, so the next 2-minute poll picks it up.
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
published and running on its 2-minute schedule. No manual runs, nothing pinned.

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
automatically on the 2-minute schedule. Every outcome matched what
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

## Open items still owed by Yancey

From the spec §13 and the client doc: exact signature block for first contacts,
booking link, digest time confirmation, follow-up threshold, and any senders
that should always be flagged. The first two are `signature` / `bookingLink` in
WF-B's Config; the last is `alwaysFlagSenders` in WF-A's Config, which accepts
either full addresses or `@domain.com`.
