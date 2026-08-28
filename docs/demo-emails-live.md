# Demo set 2 — the one to send in front of Yancey

Round 1 (`demo-emails.md`) has all been sent and logged; re-sending hits the
idempotency gate. This is the fresh set, and it demonstrates the work done on
27 Aug rather than repeating round 1.

Four Gmail drafts are already sitting in `gabdotcode01@gmail.com`, addressed to
`gtutor@partnersperform.com`. Send them together — they arrive in one
poll and light up three different Outlook categories at once.

| # | Subject | Lands as | Why it is in the set |
|---|---|---|---|
| 1 | RE: Northern district ILI extension | `AI-Review` · `pricing_guard` | The fee ask is buried two levels down **and** the thread quotes $40.6B of unrelated capital figures |
| 2 | Public Awareness deck — pocket card | `AI-Action` + draft | A real ask, in his own project vocabulary |
| 3 | Private offices two blocks from Main Street | `AI-Processed` (noise) | A warm, specific, human pitch that is still not Upperform work |
| 4 | NiSource MSA renewal — liability cap | `AI-Review` · `sensitivity_guard` | Contract language the agent must never touch |

Every outcome below was run through the live classifier and guard chain before
being written down.

---

## 1. The one that shows tonight's fix

Send this one first and let it sit. The newest message is Chuck talking about
tone and timing — no money words at all. Two levels down, NIPSCO asks for *"an
estimate of timeline and cost"*. And in between, Chuck has pasted SCE capital
figures: **$40.6B**, **$9B**, **$48.2B**, **$5.0B**.

That combination is the whole point:

- Reading only the newest message, the guard sees nothing and the model drafts
  a reply about a proposal.
- Reading the whole thread but treating *any* currency as "a price was already
  quoted", the guard stands down — because $40.6B is sitting right there — and
  again lets it through.

**Both of those were real defects, and both were fixed today.** The guard now
reads the whole thread for the ask, and stands down only when a figure has
fee-stating language beside it. A capital plan has no such words near its
numbers.

Expect `AI-Review`, guard flag `pricing_guard`, **no draft**.

> **The line to say:** "It found the fee request two replies down, and it was
> not fooled by the forty billion dollars sitting between them. Those are
> different numbers doing different jobs, and it knows which one is a price."

## 2. A real reply, in his voice

Jennifer asks two questions he can actually answer and one he cannot settle
alone. Expect `AI-Action` and a draft in his Drafts folder.

Open it and check three things: it opens `Hi Jennifer,` not `Hey` (she is a
client contact), it does not pick the pocket-card answer for him, and it ends
with a concrete next step and a soft timeframe.

## 3. Warm, human, and still junk

A named local property agent who knows his floor number, references a
neighbour, and offers to buy coffee. Nothing on the surface says junk.

Expect `AI-Processed` only — no `AI-Action`, no draft.

> **The line to say:** "The test is not whether it looks like spam. It does
> not. The test is whether it knows the difference between Upperform's work
> and Yancey's own office admin."

*This email also caught a defect while it was being written: the phrase "given
how much Upperform has grown" was tripping the pricing guard, because "how
much" sat in both halves of a rule meant to need two independent signals. Fixed
the same day.*

## 4. The contract one

Liability caps and indemnification. Expect `AI-Review` with
`sensitivity_guard`, and no draft — the agent will not put words in his mouth
on contract terms.

*Also found while writing this one: the stem `indemnif` could never match
`indemnification`, because of a trailing word boundary in the pattern. It had
never once fired on the word it was written for. Now `indemni\w*`, which also
catches `indemnities`.*

---

## Watching it happen

Categories in Outlook tell the story live. For the detail that separates #1
from #4 — both land as `AI-Review` — open the `agent_email_log` data table:
`guard_flags` reads `pricing_guard` on one and `sensitivity_guard` on the other.

Everything arrives from a gmail.com address, so the sender signal is weaker
than it will be in Yancey's real mailbox. The content carries it.

## Verified, not predicted

```
OK  1  Northern district ILI extension   flag_for_yancey  0.99  no-draft  [pricing_guard]
       newest message alone: false   whole thread: true
OK  2  Public Awareness deck             opportunity      0.99  DRAFT
OK  3  Private offices, Main Street      noise            0.99  no-draft
OK  4  NiSource MSA renewal              flag_for_yancey  0.99  no-draft  [sensitivity_guard]
```

The first run of this set had **two** of the four behaving differently from
what this document claimed. Rather than soften the claims, the two defects
behind them were fixed and the set re-run.
