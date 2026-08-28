# Access request — SharePoint (draft for Gabriel to send)

Two parts. The first is for Yancey and stays in plain language. The second is a
block he can forward untouched to whoever administers Upperform's Microsoft 365
tenant — Yancey may or may not be that person, and asking him to translate would
be how detail gets lost.

Send both together. The Entra change has multi-day lead time and Thursday's
meeting is much better if it has already landed.

---

## Part 1 — to Yancey

> Hi Yancey,
>
> Following up on the SharePoint piece before Thursday — this is the "some
> needed access" I mentioned on the call.
>
> The admin rights you gave me do work: I can open the SharePoint sites myself
> in the browser. But the automation signs in separately from me, and that
> separate sign-in is currently limited to email only — it was set up for the
> inbox assistant, so email is all it was ever given. When it tries to look at
> SharePoint, Microsoft refuses it. I tested this rather than assumed it.
>
> So before the AI can rename or reorganise anything, someone with Microsoft 365
> admin on partnersperform.com needs to widen that permission. I have written
> the exact change below — it should be about ten minutes for whoever handles
> your Microsoft admin.
>
> One thing worth your decision rather than mine. There are two ways to do it:
>
> **Option A — limited to one library.** You tell me which document library you
> want organised, and the automation is given access to that one and nothing
> else. Everything else in your SharePoint stays out of reach, permanently, no
> matter what. This is what I would recommend.
>
> **Option B — access to everything.** Simpler to set up, but the automation
> could technically reach any SharePoint site in the company. The safety would
> then rest on my code being careful rather than on the permission itself.
>
> I would rather build it the narrow way. If Option A, let me know which teams
> to include — from what I can see there is at least "Task Management" and
> "Upperform Partners", and each one is granted separately, so it is worth
> naming them all now rather than coming back to you twice.
>
> I have also asked them to remove one leftover permission from the email
> assistant while they are in there. It is not being used and nothing is wrong,
> but it is broader than that assistant needs and this is the natural moment to
> tidy it.
>
> Once it is done I have a two-second check that confirms it worked, and I will
> come to Thursday with a proposed folder structure built from your actual
> files rather than a generic example.
>
> Thanks,
> Gabriel

---

## Part 2 — forward this to the Microsoft 365 / Entra admin

> **Subject: Entra app permission change — n8n automation (2 items)**
>
> There is an existing Entra app registration used by an n8n automation for
> Upperform's inbox assistant. Two changes are needed.
>
> **1. Remove `Mail.Send`.**
> The app currently holds `Mail.Send` and it is not used — the assistant creates
> reply drafts and never sends. Please remove the permission and revoke the
> existing consent grant so the change takes effect on refresh. This is a
> reduction in access, nothing depends on it.
>
> **2. Add SharePoint access. Preferred approach first.**
>
> *Preferred — `Sites.Selected`:*
> - Add the **application** permission `Sites.Selected` (Microsoft Graph) and
>   grant admin consent.
> - Then grant the app `write` on the single named site only, e.g.
>   ```
>   POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions
>   {
>     "roles": ["write"],
>     "grantedToIdentities": [
>       { "application": { "id": "{client-id}", "displayName": "n8n automation" } }
>     ]
>   }
>   ```
> - The site will be confirmed before you run this. No other site becomes
>   reachable, which is the point of choosing this option.
> - This uses the client-credentials flow, so a **client secret** is needed
>   (expiry noted below).
>
> *Fallback — delegated, only if app-only is impractical:*
> - Delegated permissions `Sites.ReadWrite.All`, `Files.ReadWrite.All`,
>   `offline_access`, with admin consent.
> - Note this grants the automation write access to every SharePoint site the
>   signing-in user can reach, which is why it is the fallback.
>
> **What we need back:**
> - Tenant ID, Client ID
> - Client secret (and its expiry date, so renewal can be diaried before it
>   breaks in production)
> - Confirmation of which option was applied
>
> **If a new app registration is created instead of extending the existing one,**
> it needs a Web platform redirect URI pointing at the n8n callback:
> `https://n8n-wevu.srv1921545.hstgr.cloud/rest/oauth2-credential/callback`
> — copy the exact value from the n8n credential screen before entering it, as
> n8n displays the authoritative one.

---

## Notes for Gabriel, not for sending

- Do not soften "I tested this rather than assumed it." Yancey stated on the
  call that the admin rights should be sufficient. Correcting that gently but
  unambiguously is the whole job of the message — a vaguer version invites
  "but I already gave you access" and costs a round-trip.
- The Mail.Send removal is deliberately framed as tidying, not as a problem. It
  is genuinely not a problem — no workflow can reach it — and presenting it as
  one would be alarming out of proportion to the risk. But it does need an
  admin, and this is the only admin conversation on the calendar.
- If the reply is "just do Option B, it's faster": that is Yancey's call to
  make and it is a legitimate one. Take it, and say plainly that the dry-run
  default and the approval gate are then carrying the safety on their own.
- Client secret expiry is the classic thing nobody diaries. Get the date in
  writing and put it in `docs/DEPLOY.md` the day it arrives.
