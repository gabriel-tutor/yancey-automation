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
  credentials, and n8n's native SharePoint node becomes unusable. See
  "What the native node can and cannot do" below — the honest accounting is
  that this costs us two operations out of roughly eight, which is not enough
  to trade a bounded blast radius for.
- File changes appear in version history as the app, not as a person. For an
  audit trail that is arguably better, but Yancey should hear it before he sees
  it.
- The per-site grant is a one-time admin command rather than a consent click.

Fallback if app-only proves awkward: delegated `Sites.ReadWrite.All` +
`Files.ReadWrite.All` + `offline_access`, with the wider blast radius accepted
and mitigated only by the dry-run and approval gates in the design.

---

## What the native node can and cannot do

Checked against the node's type definitions rather than assumed, because an
earlier draft of this document got it wrong in the direction that would have
embarrassed us in front of the client.

**It can rename.** `file:update` takes an optional `fileName` ("if not
specified, the original file name will be used") and `changeFileContent`
defaults to `false`, so a name change without touching content is a supported
operation. Half of Yancey's ask — *"I want to see if it can rename things"* —
is a first-class feature of a node that ships with n8n.

**It cannot move, and it cannot enumerate.** The eleven actions are file
download/update/upload, list-item create/upsert/delete/get/getAll/update, and
list get/getAll. There is no folder resource, no destination parameter on
`file:update`, no create-folder, and no `delta`. The `getFolders`/`getFiles`
search-list methods exist only to populate UI pickers, not as data operations.

So the other half of the ask — *"organize it in a way"* — needs raw HTTP
against Graph regardless. Inventory (`/drive/root/delta`), folder creation, and
`PATCH parentReference` for moves all have to be written by hand. The design was
already raw-HTTP for those stages; what changes is only that we should stop
describing the native node as useless. It is not. It is half of the job, and
worth pointing at when explaining why the other half needs custom work.

One genuinely useful thing it does give us: `item:create` / `item:update` /
`item:getAll` against a SharePoint List is exactly the approval-queue mechanism
in Stage 3. If we end up on the delegated fallback rather than `Sites.Selected`,
that stage gets simpler.

---

## What the libraries actually look like

Observed 2026-08-28 from Teams in the browser, where Gabriel's admin rights do
work. Two channels of two teams — a partial sample, not an inventory. Treat
everything here as a strong signal and none of it as a count.

**These are Teams-backed sites, not one document library.** "Task Management"
and "Upperform Partners" are separate teams, so separate SharePoint sites, with
channels as folders inside each team's Documents library. The spec's "one site
collection in v1" needs restating: it is one *team* in v1, and a `Sites.Selected`
grant is per-site, so the request should name each team we intend to touch.

### Yancey has already built a taxonomy, and it is a good one

`Power BI Task Management` contains, all created by Yancey Dawson on 3 April
2025:

```
1. Assigned Tasks
2. In Progress
3. Under Review
4. Completed Tasks
```

Numbered so they sort in process order. Whatever we propose should extend this,
not replace it — he will recognise his own scheme, and a proposal that quietly
overwrites it reads as the tool not having looked.

**It has already drifted**, which is the argument for the work: sitting
unnumbered alongside those four are `Resources and References` and `PD alignment
Excel and Power BI - Karen`. A person's name in a folder name is the classic
signal of a structure that stopped being maintained.

### The demonstrable problem is version sprawl

Three files in `Web Excel`, all modified 21 July, same base name:

```
OM_Savings_Idea_Capture_FinalVersion.xlsx    Jeremiah Rosales
OM_Savings_Idea_Capture_v4.xlsx              Charles Coleman
OM_Savings_Idea_Capture_WithCostBucke….xlsx  Jeremiah Rosales
```

A "FinalVersion" and a "v4" co-existing, by different authors, on the same day.
That is the single most persuasive thing to open Thursday with, because it needs
no explanation — and it involves the President, so it is visibly not a junior
filing problem.

Across the sample there are **at least six competing idioms for the same
concept**: `v1`, `v4`, `FinalVersion`, `Latest_`, `Draft_`, `WithCostBuckets`.
Separators are equally mixed — `OM_Savings_Idea_Capture` (underscores),
`Draft_Assessment Tool_6 Walkthru_26-Jan-…` (underscores, spaces and hyphens in
one name), `… - Karen` (spaced hyphen). Dates appear inline in at least one
format. Owner initials appear in `Latest_ActionTracker_YD.xlsx`.

This is exactly the shape of problem a deterministic pass handles well: the
inconsistency is mechanical, so most of the value lands without the model
needing to be clever about content.

### The spec's DocType vocabulary is wrong for this content

§3 assumes a consulting document library — Proposal, SOW, Contract, Report,
Deck, Invoice, Playbook, Template, Notes, Data, Misc. What is actually here is
an operational workspace: `.xlsx` trackers, dashboards and one `.pbix`.
`DIMP_Dashboard`, `Electric_Operations_Daily_Review_Dashboard`,
`Daily_Employee_Project_Time_Tracking`, `Mooyah_Cleaning_Gantt`,
`TCO Pilot Program Oversight Dashboard`, `Latest_ActionTracker`.

Dashboard, Tracker and Gantt are not in the vocabulary and are most of what is
there. Revisit the closed list against the real inventory before proposing it —
a taxonomy whose top category is `Misc` will not survive the meeting.

### One thing that must not be automated

The four numbered folders are **workflow states, not document categories**.
Moving a file from `2. In Progress` to `3. Under Review` is a status change with
business meaning — it says a person finished something. An organizer that
"tidies" across those folders would silently be changing task status.

Renaming inside them is fine. Moving between them is not a filing decision and
should be out of scope for automated moves, or gated far harder than a plain
library. Worth saying out loud on Thursday: it demonstrates that the tool
distinguishes what folders *mean* from what they are called, which is the
distinction that makes "let the AI do it" a comfortable thing to agree to.

### Note on confidentiality

This section names real Upperform files and staff. Same class of material as
`fixtures/golden/`, and the same mitigation applies: this repository has no git
remote and should not acquire one.

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

Browser access changes this materially. The taxonomy proposal no longer waits
on the API grant, because the content can be read by eye — and a proposal built
from `OM_Savings_Idea_Capture_FinalVersion.xlsx` is worth more in that meeting
than one built from a seeded dev library, grant or no grant.

So Thursday is concrete either way. What the grant changes is only whether we
can *show it running*.

**With access:** run WF-S0 live, enumerate the library, and demonstrate a
dry-run rename plan against real files. He sees the thing work.

**Without:** present the proposal against the real filenames, screenshots and
all, and demo the mechanism against a seeded dev library. He sees what it would
do, and signs off the taxonomy — which is the actual prerequisite and the thing
that gates every later stage.

Either way the access request should go now: it is on the critical path for
building, just no longer for the meeting. Sending it also stops Thursday
becoming the place we ask, which would put the whole of Phase 2 a week behind
for a reason Yancey would rightly find avoidable.

## Open

- [ ] Send Yancey the access request (see `docs/access-request-sharepoint.md`),
      naming both teams rather than "one site"
- [ ] Build the naming-convention proposal from the observed filenames —
      does not depend on the grant
- [ ] Revisit the §3 DocType vocabulary against what is actually in the
      libraries (Dashboard / Tracker / Gantt are missing and are most of it)
- [ ] Decide app-only vs delegated once the admin replies
- [ ] Register/extend the Entra app; add the credential in n8n
- [ ] Re-run WF-S0 to confirm PASS
- [ ] Taxonomy sign-off from Yancey — prerequisite, not a formality
- [ ] Confirm with Yancey that the four numbered status folders are out of
      scope for automated moves
