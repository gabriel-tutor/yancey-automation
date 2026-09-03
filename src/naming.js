/**
 * Naming by a model, checked by code.
 *
 * The first version of the organiser named files with regular expressions, and
 * it could only ever rearrange the characters already in the filename. Yancey's
 * convention opens with a project code and then a short descriptive name, and
 * neither of those is in `Metrics .xlsx`. So it produced
 * `Metrics_2025-04-15_Draft_v0.1.xlsx` — the old junk with a date bolted on,
 * wearing the convention as a costume. He asked for an AI agent on the call and
 * said so plainly: *"I want the AI to be able to do it."*
 *
 * The division of labour here is the point:
 *
 *   the model supplies the WORDS   — project, document name, status, date
 *   the code enforces the SHAPE    — format, guards, collisions, validation
 *
 * The model never writes a filename. It returns parts, and `proposeFileName`
 * assembles them, so a malformed name is not a thing the model can produce. And
 * a proposal that fails `diagnose` is refused rather than applied, which means
 * the worst a bad answer can do is leave a file alone.
 */
import { STATUSES, diagnose, proposeFileName } from './filenames.js';
import { folderOf, isReport, pathKey, reportReading, splitExtension } from './library.js';

/**
 * A line of the approved plan. It approves a name, a destination, or both — and
 * `from` is matched the way SharePoint matches, without case.
 *
 * @typedef {{ from: string, to?: string, into?: string }} ApprovedLine
 */

/** Yancey's guide: v0.x is a working draft, v1.0 is the first approved one. */
const FIRST_APPROVED = { major: 1, minor: 0 };
const FIRST_DRAFT = { major: 0, minor: 1 };

/**
 * Files a model may name — or, with `forPlacement`, files it may be shown at
 * all.
 *
 * The Power BI guards are not negotiable and never reach the model: a `.pbix`
 * keeps its data-source paths inside itself, so no amount of reading the
 * library tells you what renaming one would break. Everything else is fair
 * game, including files that already parse — those are the ones carrying no
 * project code, which is exactly what the regex could not supply.
 *
 * Naming and placing are different questions with different answers. A document
 * with no extension cannot be given a name in a convention that ends in one,
 * but it can certainly be put in the right folder — and a file the organiser
 * can never touch at all is a gap rather than a guardrail. `forPlacement`
 * relaxes the extension rule and nothing else: a report is still kept out of
 * both, because moving one changes its path exactly as renaming it would.
 *
 * @param {Array<{ current: string, path: string, blocked?: string, compliant?: boolean }>} rows
 * @param {Array<{ name: string, path?: string }>} reports
 * @param {{ dependencies?: { byFile: Record<string, { report: string, where: string }>, unreadable?: string[] }, includeReports?: boolean, forPlacement?: boolean }} [options]
 * @returns {Array<{ current: string, path: string }>}
 */
export function namingCandidates(rows, reports, options) {
  return (Array.isArray(rows) ? rows : []).filter((row) => {
    const name = String(row.current || '');
    if (!splitExtension(name).extension && !(options && options.forPlacement)) return false;
    // A report is excluded by default. Not because renaming one is unsafe in
    // the way a workbook is - the reports here read nothing from this library -
    // but because a .pbix is also referenced from outside it, by whatever
    // published dataset or refresh schedule was set up from it, and none of
    // that is visible from here or from inside the file.
    if (isReport(name) && !(options && options.includeReports)) return false;
    if (reportReading(name, String(row.path || ''), reports || [], options)) return false;
    return true;
  });
}

/** The shape the model must answer in. Names are assembled from these parts. */
const SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['proposals', 'folders'],
  properties: {
    folders: {
      type: 'array',
      description:
        'Existing folders whose names do not say what they hold. Leave empty if the folder names are already clear.',
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['from', 'to', 'why'],
        properties: {
          from: { type: 'string', description: 'the folder path, copied exactly from the list given' },
          to: { type: 'string', description: 'the new folder name only, not a path' },
          why: { type: 'string', description: 'one short sentence on what makes the new name better' },
        },
      },
    },
    proposals: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['file', 'project', 'document', 'status', 'date', 'version', 'category', 'folder', 'why'],
        properties: {
          file: { type: 'string', description: 'the current file name, copied exactly' },
          folder: {
            type: 'string',
            description:
              'The name of the folder this document belongs in — an existing folder name reused exactly, or a new grouping in Title Case. Empty string to leave it where it is. Name only, never a path.',
          },
          project: {
            type: 'string',
            description:
              'Client or project code evidenced by the document or its folder path, e.g. NIPSCO, SCE, Upperform, DIMP. Empty string if nothing evidences one. Never guess.',
          },
          document: {
            type: 'string',
            description:
              'Two to four words saying what this document IS, hyphenated, e.g. Ride-Along-Plan, CAIDI-Training-Guide, Crossarm-Survey. Never the old file name, never a date, never a time.',
          },
          status: {
            type: 'string',
            enum: ['Draft', 'Review', 'Final', 'Archived', ''],
            description: 'Only if the document says so. Empty otherwise.',
          },
          date: {
            type: 'string',
            description:
              'YYYY-MM-DD, only if the document itself carries a date that is clearly its own. Empty otherwise. Never resolve an ambiguous number like 04042025 unless the document states the date in words or unambiguous form.',
          },
          version: {
            type: 'string',
            description: 'v#.# only if the document states a version. Empty otherwise.',
          },
          category: {
            type: 'string',
            description: 'What kind of document this is, in two or three words.',
          },
          why: {
            type: 'string',
            description: 'One short sentence on what in the document evidenced the project and the name.',
          },
        },
      },
    },
  },
};

const INSTRUCTIONS = /* @__PURE__ */ [
  "You are naming documents in Yancey Dawson's SharePoint library at Upperform",
  'Partners, a management consultancy for regulated electric and gas utilities.',
  '',
  'His convention, sent 29 August 2026:',
  '',
  '  [Project]_[Document Name]_[YYYY-MM-DD]_[Status]_v#.#',
  '',
  '  SCE_PM_Ride-Along-Plan_2026-09-01_Draft_v0.1.docx',
  '  NIPSCO_CAIDI_Training-Guide_2026-08-29_Review_v0.4.pptx',
  '  Upperform_Momentum-Room_One-Pager_2026-08-29_Final_v1.0.pdf',
  '',
  'You return the PARTS. Something else assembles the name and checks the format,',
  'so do not write a file name and do not add the extension.',
  '',
  'What each part means:',
  '',
  '- PROJECT is the client or programme the document belongs to. Take it from the',
  '  document contents or its folder path. NIPSCO, SCE, Upperform, DIMP and TCO are',
  '  real ones here. If nothing in front of you evidences a project, return an empty',
  '  string. An invented project code is worse than none.',
  '',
  '- DOCUMENT is two to four words saying what the thing IS, hyphenated. This is the',
  '  part that has to be written rather than copied: "Screenshot" is not a document',
  '  name, and neither is "Metrics". Say what it shows or holds —',
  '  Crossarm-Survey, Idling-Cost-Analysis, Channel-Guide, Roster.',
  '',
  '- STATUS only if the document says so. A deck marked DRAFT is Draft; a signed',
  '  document is Final. Otherwise empty, and Draft will be assumed and disclosed.',
  '',
  '- DATE only if the document carries one that is clearly its own — a report date,',
  "  a version date, a header. Otherwise empty and the file's timestamp is used.",
  '  If a name contains an eight-digit run like 04042025, that is 4 April under one',
  '  convention and 4 May under another. Resolve it ONLY if the document states the',
  '  date unambiguously somewhere. Otherwise leave date empty.',
  '',
  '- VERSION only if the document states one.',
  '',
  'Rules that override anything you infer:',
  '- Never invent a figure, a date, a status or a project.',
  '- Never put a time of day in the name.',
  '- Never use Final_Final, Latest or New.',
  '- Two files must not end up with the same name. Where several are similar, the',
  '  document name is what tells them apart, so make each one specific.',
  '',
  'Answer for every file listed, using the file name exactly as given.',
  '',
  '',
  'THE FOLDER STRUCTURE',
  '',
  'You are also grouping these documents. This part is yours to decide: there is',
  'no template to fill in and no list of approved folder names. Read what the',
  'documents are and say where each one belongs.',
  '',
  'Give each document a FOLDER — the name of the folder it should sit in:',
  '',
  '- Reuse an existing folder name, copied exactly, when the document belongs in',
  '  one that is already there. Filing beats inventing.',
  '- Otherwise name a new grouping in Title Case, two or three words, saying what',
  '  the folder HOLDS: Dashboards, Rate Case Filings, Training Materials.',
  '- Empty string leaves the document where it is, and that is the right answer',
  '  whenever nothing better is true.',
  '',
  'What makes this a structure rather than a pile:',
  '',
  '- Aim for FOUR OR FIVE folders across twenty documents, not twenty. If a',
  '  folder would hold one document it will be refused, because a folder holding',
  '  one file is the pile it came from with an extra click in front.',
  '- Use the SAME name for documents that belong together. Two names for one idea',
  '  ("Dashboards" and "Dashboard Files") splits the group and loses both.',
  '- A new folder is created beside the documents that go in it, so group',
  '  documents that sit in the same folder today. To file something into a folder',
  '  that already exists somewhere else, give that folder name exactly.',
  '- Never a numbered stage folder like "1. Assigned Tasks". Those record how far',
  '  work has got, and nothing about a document says which stage it is in.',
  '- No slashes, no punctuation beyond & and -, three to fifty characters.',
  '',
  'FOLDERS is a separate list, for existing folders whose names do not say what',
  'they hold — copy the path exactly from the folder list and give a new name',
  'only, not a path. Leave it empty if the names are already clear. Renaming a',
  'folder moves everything inside it, so propose one only where the current name',
  'is genuinely unhelpful.',
  // @__PURE__ lets esbuild drop this from a node that never asks for
  // buildNamingRequest. Without it the prompt travelled as 4.5KB of dead code
  // inside the apply node, where nothing reads it.
].join('\n');

/**
 * The Responses API payload.
 *
 * Text-bearing files arrive as extracted text, which costs nothing but tokens.
 * PDFs and images arrive as themselves, because a scanned page or a screenshot
 * has no text to extract and looking at it is the only way to know what it
 * shows — which is exactly what the three same-day screenshots needed.
 *
 * The folder list travels too, because the model cannot reuse a folder name it
 * has never seen — and reusing one is almost always the better answer than
 * inventing another.
 *
 * @param {Array<{ name: string, path: string, folder?: string, category?: string, text?: string, dataUri?: string, isImage?: boolean, contentNote?: string }>} candidates
 * @param {{ model?: string, textPerFile?: number, folders?: string[] }} [options]
 */
export function buildNamingRequest(candidates, options) {
  const opts = options || {};
  const model = String(opts.model || 'gpt-5.6-luna');
  const perFile = opts.textPerFile === undefined ? 1200 : opts.textPerFile;

  const content = [];
  const lines = [];

  const folders = (opts.folders || []).map(String).filter(Boolean).sort();
  if (folders.length > 0) {
    lines.push('Folders that already exist, paths copied exactly:', '');
    for (const f of folders) lines.push('  ' + f);
    lines.push('');
  }

  lines.push('Files to name:', '');

  (Array.isArray(candidates) ? candidates : []).forEach((c, i) => {
    lines.push('--- FILE ' + (i + 1) + ' ---');
    lines.push('name: ' + c.name);
    lines.push('folder: ' + (c.folder || '(library root)'));
    if (c.category) lines.push('kind: ' + c.category);
    if (c.text) {
      lines.push('contents:');
      lines.push(String(c.text).slice(0, perFile));
    } else if (c.dataUri) {
      lines.push('contents: attached below as FILE ' + (i + 1));
    } else {
      // Not read and could-not-be-read are different facts, and the model
      // should not treat a switched-off library as an unreadable document.
      lines.push('contents: ' + (c.contentNote || 'not readable, name and folder only'));
    }
    lines.push('');
  });

  content.push({ type: 'input_text', text: lines.join('\n') });

  for (const c of Array.isArray(candidates) ? candidates : []) {
    if (!c.dataUri) continue;
    content.push({ type: 'input_text', text: 'FILE: ' + c.name });
    content.push(
      c.isImage
        ? { type: 'input_image', image_url: c.dataUri }
        : { type: 'input_file', filename: c.name, file_data: c.dataUri },
    );
  }

  return {
    model,
    instructions: INSTRUCTIONS,
    input: [{ role: 'user', content }],
    text: { format: { type: 'json_schema', name: 'document_names', strict: true, schema: SCHEMA } },
    store: false,
  };
}

/** @param {string} value @returns {{ major: number, minor: number } | null} */
function readVersion(value) {
  const m = /^v?(\d+)\.(\d+)$/.exec(String(value || '').trim());
  return m ? { major: Number(m[1]), minor: Number(m[2]) } : null;
}

/** @param {string} value @returns {string} */
function readDate(value) {
  const v = String(value || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(v) ? v : '';
}

/**
 * Merge the model's parts into the plan, refusing anything that does not hold.
 *
 * Every proposal is assembled by `proposeFileName` and then read back by
 * `diagnose`. A name that does not satisfy the convention is dropped and the
 * file is left alone, so the failure mode of a bad answer is inaction.
 *
 * @param {any[]} rows plan rows from planRenames
 * @param {any[]} proposals model output
 * @param {{ files?: any[], reports?: any[], assumeStatus?: string, includeReports?: boolean, dependencies?: { byFile: Record<string, { report: string, where: string }>, unreadable?: string[] } }} [options]
 */
export function applyProposals(rows, proposals, options) {
  const opts = options || {};
  const reports = opts.reports || [];
  const assumeStatus = opts.assumeStatus === undefined ? 'Draft' : opts.assumeStatus;
  const modified = new Map(
    (opts.files || []).map((f) => [String(f.path), String(f.modified || '').slice(0, 10)]),
  );
  const nameable = new Set(namingCandidates(rows, reports, opts).map((r) => String(r.current)));
  const placeable = new Set(
    namingCandidates(rows, reports, { ...opts, forPlacement: true }).map((r) => String(r.current)),
  );
  const byName = new Map(
    (Array.isArray(proposals) ? proposals : []).map((p) => [String(p && p.file), p]),
  );

  for (const row of rows) {
    const p = byName.get(String(row.current));
    if (!p) continue;

    // Defence in depth. A guarded file should never have been sent, and must
    // not be touched even if a proposal comes back for it anyway.
    if (!placeable.has(String(row.current))) continue;

    // Where a document belongs and what it is called are separate judgements,
    // so the folder is recorded before the name is checked. A file that is
    // already named correctly can still be in the wrong place, a name that
    // fails the convention should not take a good filing decision with it, and
    // a document with no extension can be filed even though it cannot be named.
    // `planStructure` decides which of these suggestions become real folders.
    row.folderSuggestion = String(p.folder || '').trim();

    if (!nameable.has(String(row.current))) continue;

    const document = String(p.document || '').trim();
    if (!document) continue;

    const status = STATUSES.indexOf(String(p.status || '')) !== -1 ? String(p.status) : '';
    const date = readDate(p.date);
    const version = readVersion(p.version);

    const assumptions = [];
    if (!p.project) assumptions.push('no project code was evidenced, so none was added');
    const useStatus = status || String(assumeStatus || '');
    if (!status && useStatus) assumptions.push('status assumed to be ' + useStatus + ', the document does not say');

    const stamp = modified.get(String(row.path)) || '';
    const useDate = date || stamp;
    if (!date && stamp) assumptions.push('date taken from the file timestamp, not the document');

    // An approved version starts at v1.0, so assuming v0.1 under a Final status
    // would produce a name his own guide calls wrong.
    const approved = useStatus === 'Final' || useStatus === 'Archived';
    const useVersion = version || (approved ? FIRST_APPROVED : FIRST_DRAFT);
    if (!version) assumptions.push('version assumed to be v' + useVersion.major + '.' + useVersion.minor);

    let proposed = '';
    try {
      proposed = proposeFileName({
        project: String(p.project || ''),
        document,
        date: useDate,
        status: useStatus,
        version: useVersion,
        extension: splitExtension(String(row.current)).extension,
      });
    } catch {
      continue;
    }

    // The model supplied words; this is where they are held to the convention.
    if (diagnose(proposed).length > 0) continue;
    if (proposed === row.current) {
      row.compliant = true;
      row.proposed = '';
      row.blocked = '';
      row.approved = false;
      continue;
    }

    row.proposed = proposed;
    row.blocked = '';
    row.compliant = false;
    // With no pinned plan, arming the switch IS the approval - the person who
    // armed it is saying "apply what you propose". `applyApprovedPlan` demotes
    // these again when a plan exists, so the flag always means the same thing:
    // this name is cleared to be applied.
    row.approved = true;
    row.awaiting = '';
    row.assumptions = assumptions;
    row.namedBy = 'model';
    row.why = String(p.why || '');
    if (p.category) {
      // The deterministic pass classified this from its extension. The model
      // classified it by reading it, which is a better answer from different
      // evidence — so the evidence line has to move with the category, or the
      // page shows "Process Guide ... from the .docx extension".
      row.category = String(p.category);
      row.categoryWhy = 'from reading the document, not its extension';
    }
  }

  refuseCollisions(rows);
  return rows;
}

/**
 * Names are only unique inside a folder, and a model can land two files on one
 * name exactly as a rule can. Both are refused rather than one being destroyed.
 *
 * @param {any[]} rows
 */
function refuseCollisions(rows) {
  /** @type {Record<string, number>} */
  const wanted = {};
  for (const row of rows) {
    const key = folderOf(row.path).toLowerCase() + '/' + String(row.proposed || row.current).toLowerCase();
    wanted[key] = (wanted[key] || 0) + 1;
  }
  for (const row of rows) {
    if (!row.proposed) continue;
    const key = folderOf(row.path).toLowerCase() + '/' + row.proposed.toLowerCase();
    if ((wanted[key] || 0) > 1) {
      row.blocked =
        String(wanted[key]) +
        ' files in this folder would end up with the same name — ' +
        row.proposed +
        ', so they need names that tell them apart';
      row.proposed = '';
      row.approved = false;
      row.awaiting = '';
    }
  }
  return rows;
}

/**
 * Where each approved line has got to.
 *
 * This exists because a hand-typed plan silently half-applied. Three macOS
 * screenshots carry U+202F, a narrow no-break space, before "PM" — and a plan
 * typed with an ordinary space matched nine of twelve lines and quietly skipped
 * the other three. Leaving a file alone is the right thing to do with a line
 * that does not match; doing it without saying so is not.
 *
 * But "the file is not there" has two very different causes, and reporting them
 * as one made every run after a successful one look like a failure: nine lines
 * came back as *"no longer in the library under that name"* when all nine had
 * simply been carried out, and the file was sitting there under exactly the
 * name that was approved. A line whose RESULT is in the library is done. Only a
 * line with neither its file nor its result is genuinely adrift, and that is
 * the one worth a person's attention.
 *
 * A line approving both a name and a destination can be half carried out: the
 * rename lands and the move does not. That line is neither done nor adrift, and
 * calling it done was how an approved move stopped happening on every later run.
 *
 * @param {any[]} rows
 * @param {ApprovedLine[]} approved
 * @returns {{ done: string[], missing: string[] }}
 */
export function approvedProgress(rows, approved) {
  const byName = new Map(
    (Array.isArray(rows) ? rows : []).map((r) => [pathKey(r.current), r]),
  );
  /** @type {string[]} */
  const done = [];
  /** @type {string[]} */
  const missing = [];

  for (const line of Array.isArray(approved) ? approved : []) {
    if (!line || !line.from) continue;
    // Still in the library under its old name: this run will deal with it.
    if (byName.has(pathKey(line.from))) continue;

    const renamed = line.to ? byName.get(pathKey(line.to)) : undefined;
    if (!renamed) {
      missing.push(String(line.from));
      continue;
    }
    // The name landed. The destination, if the line named one, may not have —
    // and this run is about to see to that, so it is neither done nor adrift.
    if (line.into && pathKey(folderOf(String(renamed.path || ''))) !== pathKey(line.into)) continue;
    done.push(String(line.to));
  }

  return { done, missing };
}

/**
 * Apply a plan a person has already read and approved, instead of whatever this
 * run's model call happens to say.
 *
 * This exists because the model is not deterministic and the difference matters.
 * Two runs an hour apart over the same unchanged library proposed
 * `Time-Zone-Conversion-Table` and then `Timezone-Conversion-Table`, and
 * `Net-Profit-Period-Dashboard` and then `Profit-Performance-Dashboard`. Both
 * are reasonable names. Neither is the one on the page Yancey approved — and a
 * review that does not bind to what gets applied is not an approval, it is a
 * demonstration. So arming the switch stops asking the model anything and does
 * exactly what was signed off.
 *
 * Every guard still runs. An approved line is a request, not an instruction:
 * the file must still exist, still be un-guarded, keep its extension, satisfy
 * the convention, and not collide.
 *
 * `revert` is the one exception, and it is narrow: putting a name *back* is not
 * a naming decision, so the convention does not apply to the name being
 * restored. It could not, since that name is the messy one — refusing it is the
 * reason it was being changed. Everything else still holds.
 *
 * What a plan gates is what gets **applied**, and only that. It must not gate
 * what may be **proposed**. An earlier version conflated the two, and on
 * 3 September Yancey dropped eight new documents into the library, ran the
 * organiser, and got eight refusals reading *"not on the approved list, so this
 * run leaves it alone"* — the model had read and named every one, and the gate
 * threw the names away before anybody could look at them. A file that arrives
 * after a plan is signed could then never be reviewed at all: the loop that is
 * supposed to close never closes. So an off-list proposal is kept and marked
 * unapproved, and the run reports two lists rather than one.
 *
 * A line can approve a name, a destination, or both. A line with `into` and no
 * `to` is a filing decision on its own — "this document belongs in that folder,
 * leave its name alone" — and it exists because the folder a model suggests is
 * no more deterministic than the name it writes. Two runs an hour apart grouped
 * the same two workbooks into `Dashboards` and then didn't, so the folder was
 * created once and refused the next time by the rule that a new folder must
 * hold two documents. Pinning the destination is what makes filing repeatable,
 * exactly as pinning the name did.
 *
 * @param {any[]} rows plan rows, already carrying this run's proposals
 * @param {ApprovedLine[]} approved the signed-off plan
 * @param {{ reports?: any[], includeReports?: boolean, revert?: boolean, dependencies?: { byFile: Record<string, { report: string, where: string }>, unreadable?: string[] } }} [options]
 */
export function applyApprovedPlan(rows, approved, options) {
  const opts = options || {};
  const allowed = new Set(namingCandidates(rows, opts.reports || [], opts).map((r) => String(r.current)));
  const byFrom = new Map();
  for (const line of Array.isArray(approved) ? approved : []) {
    if (!line || !line.from || !(line.to || line.into)) continue;
    byFrom.set(pathKey(line.from), line);
    // A line is also reachable under the name it asked for, so one whose rename
    // has already landed can still carry out the destination it approved. Keyed
    // second, so a line still waiting on its rename wins the name it starts from.
    if (line.to && !byFrom.has(pathKey(line.to))) byFrom.set(pathKey(line.to), line);
  }

  for (const row of rows) {
    const line = byFrom.get(pathKey(row.current));
    const to = line ? String(line.to || '') : '';

    // A destination is approved on its own terms. A line may carry one without
    // a name, and then it says only where the document belongs — which is a
    // placement decision, so it is recorded as one. Without this the folder
    // would be created and nothing would move into it.
    if (line && line.into) {
      row.approvedInto = String(line.into);
      const destination = String(line.into) + '/' + String(row.current);
      if (destination !== String(row.path)) {
        row.placement = 'move';
        row.destination = destination;
        row.placementWhy = 'approved on the plan as belonging in ' + String(line.into);
      }
    }

    // Nothing outside the approved list is APPLIED. It is still proposed.
    if (!to) {
      if (row.proposed) {
        row.approved = false;
        // A file whose approved name has already been applied comes back round
        // as a fresh proposal for a *different* name. That is churn rather than
        // a finding: the name it has already satisfies the convention, so say
        // so and drop the second opinion.
        if (diagnose(String(row.current)).length === 0) {
          row.compliant = true;
          row.proposed = '';
          row.blocked = '';
          row.awaiting = '';
          continue;
        }
        row.blocked = '';
        row.awaiting = 'proposed on this run — approve it and the next run applies it';
      }
      continue;
    }

    row.namedBy = 'approved';
    row.approved = false;
    row.awaiting = '';

    if (!allowed.has(String(row.current))) {
      row.blocked = row.blocked || 'approved, but it is guarded now and cannot be renamed in a sweep';
      row.proposed = '';
      continue;
    }

    // A rename must not change what kind of file something is. Nothing in the
    // convention needs it, and a .xlsx that becomes a .pdf is not a rename.
    if (splitExtension(to).extension.toLowerCase() !== splitExtension(String(row.current)).extension.toLowerCase()) {
      row.blocked = 'the approved name changes the file extension, which a rename must not do';
      row.proposed = '';
      continue;
    }

    // Putting a name back is not a naming decision, so the convention does not
    // apply to it. Without this an undo cannot run at all: the name being
    // restored is the messy one, and refusing it is the whole reason it was
    // being changed. Every other guard still holds — the file must be there,
    // keep its extension, and not collide.
    if (!opts.revert) {
      const problems = diagnose(to);
      if (problems.length > 0) {
        row.blocked = 'the approved name does not satisfy the convention — ' + problems[0];
        row.proposed = '';
        continue;
      }
    }

    if (to === row.current) {
      row.compliant = true;
      row.proposed = '';
      row.blocked = '';
      continue;
    }

    row.proposed = to;
    row.blocked = '';
    row.compliant = false;
    row.approved = true;
  }

  refuseCollisions(rows);
  return rows;
}
