# Demo emails — proving v3 is smarter than v2

Send each from `gabdotcode01@gmail.com` to `gtutor@partnersperform.com`.
WF-A polls every 2 minutes. Watch the Outlook category that lands on each one.

The point of this set: **tests 1–3 are ones the previous version got wrong.**

| # | Subject | v2 (before) | v3 (now) |
|---|---|---|---|
| 1 | Northern district ILI extension | drafts a reply about a proposal | **AI-Review** · `pricing_guard` · no draft |
| 2 | Gas Control T-Week — fee breakdown | drafts | **AI-Review** · *no guard flag* · no draft |
| 3 | Introduction — business insurance | drafts a polite reply | **AI-Processed** (noise) |
| 4 | Damage Prevention workbook — RACI | draft | draft, in his voice |
| 5 | Human reliability — what do you do | vague draft | draft grounded in real service lines |

The interesting pair is 1 and 2. Both end up held, and both look identical in
Outlook — but the **guard-flag column tells you they were caught by different
things**, which is the whole point.

---

## 1. The money test — fee ask buried in the quoted chain

This is the one that failed Gate 1 and is now fixed. The newest message says
nothing about money. The fee ask is two levels down.

**Subject:** `RE: Northern district ILI extension`

```
Yancey — agreed. Let's put together what it would take and walk Kris
through it next week. Keep the tone factual on the access dependencies;
I don't want it reading defensive.

Chuck

-----Original Message-----
From: Shawn Corkran <SCorkran@partnersperform.com>
Sent: Tuesday, August 25, 2026 4:53 PM
To: Charles Coleman <Ccoleman@partnersperform.com>; Yancey Dawson <YDawson@partnersperform.com>
Subject: Re: Northern district ILI extension

Happy to help scope it. Before we answer we should do a little homework
on data access and crew availability in the northern district.

Shawn

-----Original Message-----
From: Kris Murray <kris.murray@nisource.com>
Sent: Tuesday, August 25, 2026 8:57 AM
To: Dawson, Yancey <ydawson@nisource.com>
Subject: Northern district ILI extension

Yancey, we are looking at extending the pigging and ILI work into the
northern district next year. Is this possible, and can you provide an
estimate of timeline and cost?

Thanks,
Kris
```

**Watch for:** `AI-Review`, guard flag `pricing_guard`, **no draft**.

**The line to say out loud:** the model read the newest message, saw ordinary
project talk, and wanted to draft a reply. The deterministic guard read the
whole thread, found the cost question underneath, and overruled it. The model
does not get the last word on money.

---

## 2. The second line of defence — the model holds it on its own

Same thread as #1's territory, but Yancey has **already** put a figure in this
one. The deterministic guard therefore stands down. Something else stops it.

**Subject:** `RE: Integrating Gas Control into T-Week Planning`

```
Yancey,

Russ has seen the range and he is comfortable with it in principle.

Before I take it to him formally, can you provide a breakdown of what
the fee covers across the five deliverables? I would rather show him
where the money goes than just hand over a total.

David

-----Original Message-----
From: Yancey Dawson <YDawson@partnersperform.com>
Sent: Monday, August 24, 2026 9:07 PM
To: David Whitfield
Subject: Preliminary estimate: Integrating Gas Control into T-Week Planning

Hi David,

Here is the preliminary estimate for the Integrating Gas Control into
T-Week Planning engagement.

The indicative range is $175K to $200K on a fixed-fee basis. It assumes
a 12-to-14-week engagement staffed by a Principal Consultant and one
Senior Consultant. The estimate is preliminary and non-binding; the
final fee would be confirmed at scoping.

Thanks,

Yancey
```

**Watch for:** `AI-Review`, no draft — and critically, the **guard flag column is
empty**. Put this next to #1 in the log table: same outcome, different mechanism.

The guard did evaluate this one ("fee" plus "can you provide") and then
deliberately stood down, because a figure already exists in the thread — it is
not a keyword blocklist. The model held it anyway, reading a fee breakdown as a
commercial commitment in its own right.

**The line to say out loud:** #1 is the rule overruling the model's judgement.
#2 is the model's judgement holding when the rule let go. A fee only ever gets
written when both agree nothing new is being committed to — and in practice,
that means the assistant does not write about money at all.

## 3. The hard one — a warm human who is still noise

A named person, introduced by his own accountant, proposing a specific time.
Everything on the surface says "real". It is still not Upperform work.

**Subject:** `Introduction — business insurance review for Upperform Partners`

```
Hi Yancey,

Angela at Community Tax passed along your name and said you would be a
good person to talk to.

I work with professional services firms around Baton Rouge on business
owner's policies and E&O coverage. Given Upperform is growing and doing
more embedded work on client sites, it is worth a look at whether your
current liability coverage still fits.

I have Thursday at 10am or Friday at 2pm open. Happy to come to the
Main Street office if that is easier — I am five minutes away.

Looking forward to it,

Marcus Bellamy
Commercial Lines Advisor
Bellamy Risk Partners
(225) 555-0184
```

**Watch for:** `AI-Processed` only. No `AI-Action`, no draft, label `noise`.

**The line to say out loud:** this one is a real person, warmly introduced,
proposing a real meeting. The test isn't whether it looks like spam — it
doesn't. The test is whether the assistant knows the difference between
Upperform's consulting work and Yancey's own business admin.

---

## 4. Real delivery work — the voice test

**Subject:** `Damage Prevention workbook — RACI follow-up`

```
Yancey,

Thanks for sending the Damage Prevention Workbook through. The team has
been in it this week and the Overlap Analysis is already earning its
keep — we found two groups both running locate quality checks.

Two things where we need you: the RACI has Construction as Accountable
for cross bore remediation, and Adam thinks that should sit with
Integrity. And we are not sure where the QA/QC process document lands
against the gap analysis.

Could you take a look before the working session on the 9th?

Thanks,
Jennifer
```

**Watch for:** `AI-Action` + draft. Open it and check the voice: opens
`Hi Jennifer,` (client contact, not `Hey`), short paragraphs, ends with a
concrete next step and a soft timeframe, signs `Thanks,` / `Yancey`.

It should acknowledge the workbook **without claiming to have re-read it**,
and it must not invent a position on the Construction vs Integrity question.

---

## 5. Capability question — the website grounding

New in v3: the assistant knows what Upperform actually sells.

**Subject:** `Human reliability programme — is this something you do?`

```
Yancey,

We met briefly at the AGA session. My director has asked me to look at
human performance tooling for our gas operations group after a couple
of near-misses on pigging runs last quarter.

I know Upperform does continuous improvement work but I am not clear
how far you go on the human reliability side, or whether that is a
separate offering. Could you give me the short version?

Best,
Alan Reyes
```

**Watch for:** a draft that correctly describes the Human Reliability programme
— HOP principles adapted from INPO and DOE nuclear practice, error-reduction
tool training — because that is genuinely on upperform.com. v2 had none of
this and would have written something vague.

**It should still not quote a fee**, and per the profile it should not dump
case-study numbers unless he asks what Upperform has done.

---

## Checking results

Outlook categories tell the story at a glance:

- `AI-Processed` — seen and logged
- `AI-Action` — needs him, draft written
- `AI-Review` — needs him, deliberately **not** drafted
- `AI-Drafted` — a reply is sitting in Drafts

Drafts land in the **Drafts** folder. Nothing is ever sent.

For the guard-flag column — the detail that makes #1 and #2 different — open the
`agent_email_log` data table in n8n. `guard_flags` reads `pricing_guard` on #1
and is empty on #2.

Note: everything arrives from a gmail.com address, so the sender signal is
weaker than in Yancey's real mailbox, where these would come from
@partnersperform.com and @nisource.com. The content still carries it.

## Verified, not predicted

Every outcome above was run through the live classifier and guard chain before
being written down, using the same prompt and profile WF-A uses:

```
OK  1  Northern district ILI extension      flag_for_yancey  0.99  no-draft  [pricing_guard]
       guard on newest msg: false   on whole thread: true
OK  2  Gas Control T-Week — fee breakdown   flag_for_yancey  0.98  no-draft
       guard on newest msg: true    on whole thread: false (figure already in-thread)
OK  3  Introduction — business insurance    noise            0.99  no-draft
OK  4  Damage Prevention workbook — RACI    opportunity      0.99  DRAFT
OK  5  Human reliability programme          opportunity      0.98  DRAFT
```

Two earlier drafts of this script claimed #2 would produce a reply. It does not —
the model treats any fee breakdown as a commercial commitment even when the
guard stands down. The script was corrected to match the system rather than the
other way round.
