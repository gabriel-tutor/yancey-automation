# Draft reply to Yancey — the document organiser agent

**Not sent.** Written for your approval. Nothing goes to Yancey until you say so.

---

**Re:** AI agent to review, categorize, rename and reorganize the SharePoint library

Yancey,

This is built and it ran against the Task Management library this morning. Short
version: it reviewed all 21 documents, has a new name ready for twelve of them,
and deliberately left nine alone. Nothing has been changed by it — the review is
the output, and applying it is a separate switch that a person turns on.

**It reads the documents.** That is the part worth flagging, because it is a
change from what I told you before. The first version worked from file names
only, and a file name is not enough: your format opens with a project code and a
short descriptive name, and neither of those is in `Metrics .xlsx`. They are
inside the document. That version could only shuffle the characters already
there, so it produced `Metrics_2025-04-15_Draft_v0.1.xlsx` — your format worn as
a costume over the old junk. You said you wanted the AI to do it, and doing it
properly means opening the file.

So it does, with limits I'd rather you hear from me than discover:

- Word, PowerPoint and Excel files are **unzipped on our own server**, and only
  the text that comes out is sent to the model. The file itself never leaves.
- A PDF or a screenshot goes as itself, because a scan has no text to pull out
  and an image of a dashboard can only be read by looking at it.
- **Power BI reports and the workbooks feeding them are never opened at all** —
  they are filtered out before anything is downloaded.
- It is one setting. Turn it off and the agent sees only names, folders and file
  types. The naming gets worse; nothing breaks.

**What that bought you.** Three screenshots taken thirty seconds apart on 6
April used to be indistinguishable. Now:

| was | becomes |
|---|---|
| `Screenshot 2025-04-06 at 6.34.18 PM.png` | `Sales-Dashboard-Overview_2025-04-06_Draft_v0.1.png` |
| `Screenshot 2025-04-06 at 6.34.50 PM.png` | `Net-Profit-Quarterly-Analysis_2025-04-06_Draft_v0.1.png` |
| `Screenshot 2025-04-06 at 6.35.10 PM.png` | `Units-Sold-Period-Analysis_2025-04-06_Draft_v0.1.png` |

And it now finds project codes, which is the part a pattern-matcher could never
do — `TCO`, `DIMP` and `Upperform` below were read out of the contents and the
folder path, not the file name:

| was | becomes |
|---|---|
| `Power BI TCO Pilot_YD_04072025.xlsx` | `TCO_Pilot-Project-Register_2025-04-08_Draft_v0.1.xlsx` |
| `DEL-asme-b31-8s-2004_2025-03-19_Draft_v0.1.pdf` | `DIMP_ASME-B31-8S-Standard_2025-03-19_Draft_v0.1.pdf` |
| `Metrics_2025-04-15_Draft_v0.1.xlsx` | `Upperform_Metric-Development-Register_2025-04-15_Draft_v0.1.xlsx` |

Where nothing in the document evidences a project, it adds none and says so. An
invented project code would be worse than none.

**Approval before implementation, which is what you asked for.** The agent
proposes; it does not act. Renaming and moving are two separate switches, both
off, and even switched on one run touches at most five files. The review page
lists every proposal with the reason it was chosen and every assumption it made
— where a date came from a file timestamp rather than the document, it says so
on that line.

**The nine it refused.** Every one of them is a Power BI report or a workbook
feeding one. A `.pbix` keeps the paths of its data sources inside itself, where nothing
outside the file can see them, so renaming a workbook can break a dashboard with
nothing on the surface to show it happened. Those get renamed deliberately, with
you, not in a sweep. The reason is written out in full on each one rather than
reduced to a code.

**What it won't do without you.** It won't move anything between
`1. Assigned Tasks` and `4. Completed Tasks`. That would be claiming how far a
piece of work has got, which no file records and only you know.

The review page is attached. If the twelve look right, say the word and I'll arm
it — five at a time, with the old names kept so any of them can be put back by
hand.

— Gabriel

---

## Notes for you before sending

- The counts (21 / 12 / 9) come from run `9006`, 2 September. If you re-run WF-F
  before sending, refresh them.
- "The review page is attached" — replace with the artifact link.
- The one thing I'd expect him to push back on: that files renamed on
  2 September get renamed *again*. That is not a loop. The first pass gave them
  well-formed but mechanical names; the agent gives them names that say what
  they hold. Worth saying out loud if he asks, rather than in the email.
