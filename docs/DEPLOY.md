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
      - EXECUTIONS_DATA_PRUNE=true      # added
      - EXECUTIONS_DATA_MAX_AGE=168     # added
    volumes:
      - n8n_data:/home/node/.n8n
      - /root/context:/data/context:ro  # added
```

The two pruning variables were missing entirely, so n8n was storing every
execution — full email bodies included — indefinitely. Spec §10 requires a
7-day window precisely so client mail does not linger on disk. A backup of the
original sits at `docker-compose.yml.bak-preprofile`.

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

Still unproven: SMTP delivery (WF-C has never sent a real digest) and the
drafting path (WF-B has not run live). And note `profileSource` came back
`fallback` on every message — step 1 has not been done, so all of the above was
classified on the reduced built-in profile, not Yancey's real one.

## Open items still owed by Yancey

From the spec §13 and the client doc: exact signature block for first contacts,
booking link, digest time confirmation, follow-up threshold, and any senders
that should always be flagged. The first two are `signature` / `bookingLink` in
WF-B's Config; the last is `alwaysFlagSenders` in WF-A's Config, which accepts
either full addresses or `@domain.com`.
