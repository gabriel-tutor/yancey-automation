# Golden set

`npm run gate1` runs every fixture in `golden/` through the exact prompt, schema
and guard chain WF-A uses in production, and scores it against the labels Yancey
supplied.

## What Yancey sent

He forwarded 20 labelled emails on 2026-08-24, each with the label in the
subject line:

| Label | Count | Where they live |
|---|---|---|
| `TRAINING EXAMPLE – OPPORTUNITY – 01…10` | 10 | `golden/opp-*.json` |
| `TRAINING EXAMPLE – JUNK – 01…05` | 5 | `golden/junk-*.json` |
| `TRAINING EXAMPLE – MY WRITING STYLE – 01…05` | 5 | `voice/voice-*.json` |

Only the first 15 are classification cases. The writing-style five are emails
Yancey **wrote** — they calibrate the drafting voice and are quoted in
`context/business-profile.md`. Running them through the classifier would measure
nothing, so they are not in `golden/`.

## Fixture format

```json
{
  "id": "opp-05",
  "source": "yancey-training-set",
  "yanceyLabel": "OPPORTUNITY",
  "behavior": "flag",
  "expected": "flag_for_yancey",
  "gate": "pricing",
  "note": "why Yancey labelled it this way, and why we diverge if we do",
  "message": {
    "id": "AAMk...",
    "subject": "Re: Pulse Power BI",
    "receivedDateTime": "2026-07-04T14:42:00Z",
    "from": { "emailAddress": { "name": "Charles Coleman", "address": "Ccoleman@partnersperform.com" } },
    "toRecipients": [{ "emailAddress": { "address": "YDawson@partnersperform.com" } }],
    "hasAttachments": true,
    "categories": [],
    "body": { "contentType": "text", "content": "full body including the quoted chain" }
  }
}
```

- `source` must be `yancey-training-set` to count toward Gate 1. The
  `synthetic-*` files are smoke tests for the harness itself and are excluded
  from that count (they still score for label agreement).
- `yanceyLabel` is his verbatim subject-line label. Never edit it to make a
  test pass.
- `behavior` is the operational meaning of his label — `draft`, `ignore`, or
  `flag`. **This is what the gate turns on.**
- `expected` is the four-way label the deployed prompt should produce. It is a
  diagnostic, not the gate.
- `gate: "pricing"` marks a Gate 3 case: must be `flag_for_yancey` and must not
  be draftable.
- Paste the body **including** the forwarded chain. Trimming it defeats the
  point — the deep chains are the hard case, and the pricing guard now reads
  them.

## How scoring works

Yancey's labels are a two-way operational judgement: reply to this, or ignore
it. The gate scores that, not the finer taxonomy.

| Outcome | Meaning | Gate |
|---|---|---|
| `correct` | Did what Yancey's label implies | pass |
| `safe-hold` | Surfaced to him but no draft written | pass — he still sees it |
| `missed` | He said reply; agent said noise | **fail** — work disappears |
| `false-draft` | He said ignore; agent drafted | **fail** |

A `safe-hold` is not a failure. Yancey labelling something OPPORTUNITY means it
matters, not that every one needs a reply — several are FYI forwards or threads
where he is only cc. The agent surfaces all of them and writes drafts for the
ones that carry a genuine ask.

Gate 1 passes only when: zero `missed`, zero `false-draft`, ≥90% label
agreement, zero `flag_for_yancey` items marked draftable, and all 15
classification fixtures present.

## A note on run-to-run variance

`gpt-5.6` rejects the `temperature` parameter, so the harness cannot pin
sampling to 0. Labels have been stable at 100% across runs; the draft /
no-draft split moves by an item or two between runs. Judge a run by the gate
outcome, not by an exact `did the right thing` count.

## Regenerating

The fixtures were built from Gmail via MCP. The bodies are the original inbound
messages — Yancey's forwarding header block is stripped and captured as
`message.from` / `subject` / `receivedDateTime`, everything below it is kept
verbatim.
