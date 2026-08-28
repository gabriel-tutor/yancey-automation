# Phase 2 — SharePoint organizer: access comes first

Yancey's ask, verbatim from the 2026-08-28 call:

> I want to organize it and I want to see if it can rename things and I want to
> see if it can organize it in a way, but I want the AI to be able to do it.

And his assumption, a minute later:

> I gave you access to this and I actually gave you admin rights on up
> performance stuff. So I mean, you should be able to do stuff. You got admin
> rights, so.

That assumption is reasonable and it is wrong — not about SharePoint itself,
where the admin rights are real, but about what the *automation's* token can
reach. Those are two different things, and the gap is the first thing Phase 2
has to close.

---

## What was measured

`WF-S0 · SharePoint access probe` (`mwvTuEG1Ta42XlNP`) — a manual-trigger,
read-only workflow: four GETs against Graph with `neverError`, then a verdict
node. Nothing in it can write. Run 1266, 2026-08-28, on the dev credential
`Outlook – Dev (gtutor)`.

All four failed:

| Check | Endpoint | Result |
|---|---|---|
| My OneDrive | `/me/drive` | **403** `accessDenied: Access denied` |
| Root Site | `/sites/root` | **403** `accessDenied: Request Doesn't have the required Permission scopes to access a site.` |
| Search Sites | `/sites?search=*` | **403** `accessDenied: Access denied` |
| Followed Sites | `/me/followedSites` | **403** `accessDenied: This operation is not supported with the provided scopes` |

Graph is not saying *you may not touch this site*. It is saying *this token was
never issued the scopes that would let it try*. Note the third message
especially: **"not supported with the provided scopes"** — that is a statement
about the token, not about permissions on any particular site.

Keep this workflow. Its value is on the far side of the access change: run it
again after the grant and it answers PASS/FAIL in two seconds, with the site
list populated, without anyone guessing.

---

## Why it fails, and which cause actually matters

**Cause A — structural, certain.** n8n's `microsoftOutlookOAuth2Api` credential
type requests a fixed scope list at authorization time: mail, calendars,
contacts. It never asks for `Sites.*` or `Files.*`. An OAuth access token can
only carry scopes that were requested when it was minted, so this credential
*cannot* reach SharePoint no matter what the tenant consents to. This is a fact
about n8n, independent of Upperform's tenant.

**Cause B — probable, and it is a tenant question.** Even with the right
credential type, the Entra app has to hold the Sites/Files permissions and have
them consented. Phase 1 already showed this tenant grants permissions at the app
level rather than per-user: a brand-new credential inherited `Mail.Send` without
anyone asking for it. So consent here is almost certainly an admin action, not
something a user click resolves.

Both are fixed by the same move, so the ambiguity costs nothing: a **separate
credential of a different type**, against an app registration that holds the
right permissions. n8n ships `microsoftSharePointOAuth2Api` as a distinct
credential type for exactly this, and self-hosted instances must supply their own
Entra client ID and secret.

---

## The recommendation: bound the blast radius at the scope, not in our code

The obvious ask is delegated `Sites.ReadWrite.All` + `Files.ReadWrite.All`. It
works, and it means the automation can rewrite **every SharePoint site the
signing user can reach**. Signed in as Gabriel — who now has admin rights — that
is all of Upperform.

Phase 1 taught precisely why that is the wrong shape. The no-send invariant held
because no workflow called `sendMail`, but the *token* could send the whole time.
Careful design is not a security boundary. Scope is.

**Ask instead for `Sites.Selected`** (application permission). An admin grants
the app write access to one named site collection and nothing else. Every other
library in the tenant is unreachable by construction, even if a workflow has a
bug. It also turns Yancey's decision from a hard question into an easy one:
*"the AI can rewrite anything in our SharePoint"* versus *"the AI can only touch
the one library you point it at."*

Trade-offs, stated rather than buried:

- `Sites.Selected` is app-only, so it uses a client-credentials flow rather than
  the authorization-code flow n8n's SharePoint credential expects. The HTTP
  Request nodes get a generic OAuth2 credential configured for client
  credentials. The organizer design was already raw-HTTP-against-Graph — the
  native SharePoint node has no rename, no move, no folder ops and no `delta`,
  so it was never going to carry this — so nothing is lost.
- File changes appear in version history as the app, not as a person. For an
  audit trail that is arguably better, but Yancey should hear it before he sees
  it.
- The per-site grant is a one-time admin command rather than a consent click.

Fallback if app-only proves awkward: delegated `Sites.ReadWrite.All` +
`Files.ReadWrite.All` + `offline_access`, with the wider blast radius accepted
and mitigated only by the dry-run and approval gates in the design.

---

## Ask for the Mail.Send removal in the same breath

Phase 1's one unresolved risk is that Yancey's mailbox credential holds
`Mail.Send` — a send attempt returns `202 Accepted`. It is not reachable by any
workflow, but it is there, and it needs an Entra admin to remove.

That is the same admin, doing the same kind of task, in the same conversation.
Getting both into one request turns two multi-day round-trips into one. Do not
split them.

---

## What Thursday looks like

**If the access lands first:** the meeting is concrete. Run WF-S0, enumerate his
actual document library, and put a taxonomy proposal in front of him built from
his real folders and real filenames — which is the one thing that makes the
"can it rename things?" conversation productive rather than hypothetical. The
design's prerequisite is a taxonomy sign-off, and that is far easier to get when
he is looking at his own files.

**If it does not:** the meeting is a demo against a seeded library in the dev
tenant, plus the access request explained in person. Workable, and noticeably
weaker — he cannot sign off a taxonomy for files he is not looking at.

The lead time is the whole argument for sending the request now rather than
after Thursday.

## Open

- [ ] Send Yancey the access request (see `docs/access-request-sharepoint.md`)
- [ ] Decide app-only vs delegated once the admin replies
- [ ] Register/extend the Entra app; add the credential in n8n
- [ ] Re-run WF-S0 to confirm PASS
- [ ] Taxonomy sign-off from Yancey — prerequisite, not a formality
