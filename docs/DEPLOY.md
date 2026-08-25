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

1. Confirm the **OAuth Redirect URL** n8n displays reads
   `https://n8n-wevu.srv1921545.hstgr.cloud/rest/oauth2-credential/callback`.
   If it shows `localhost:5678`, stop — `N8N_HOST` / `WEBHOOK_URL` are wrong and
   OAuth cannot work until they are fixed.
2. Paste the **Client ID** and **Client Secret** from the Entra app. The secret
   is the **Value** column, not the Secret ID next to it.
3. Turn on **Custom Scopes**.
4. Set **Enabled Scopes** to exactly:

```
openid offline_access User.Read Mail.Read Mail.ReadWrite
```

5. Save, then **Connect my account**.

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

Still unproven until credentials exist: the Graph calls themselves, the OpenAI
calls, SMTP delivery, and OAuth. Those are steps 2–3 above.

## Open items still owed by Yancey

From the spec §13 and the client doc: exact signature block for first contacts,
booking link, digest time confirmation, follow-up threshold, and any senders
that should always be flagged. The first two are `signature` / `bookingLink` in
WF-B's Config; the last is `alwaysFlagSenders` in WF-A's Config, which accepts
either full addresses or `@domain.com`.
