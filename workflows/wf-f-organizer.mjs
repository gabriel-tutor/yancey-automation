import { codeNode } from '../scripts/bundle.mjs';
import { HELPERS } from './helpers-manifest.mjs';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { DEPLOY, E, ROOT, S, SDK_IMPORT, TARGET } from './shared.mjs';

/**
 * The plan a person has read and signed off, kept as data rather than typed
 * into the Config node by hand. Only the from/to pairs travel into the
 * workflow; the reasons and assumptions stay in the file, where they are what
 * the review page was built from.
 */
function approvedPlan() {
  try {
    const file = JSON.parse(readFileSync(join(ROOT, 'config/approved-plan.json'), 'utf8'));
    return (file.plan || []).map((line) => ({ from: String(line.from), to: String(line.to) }));
  } catch {
    return [];
  }
}

/**
 * WF-F — the document organiser.
 *
 * Yancey asked for this on the 28 August call, and was specific about how:
 *
 *   "I want to organize it and I want to see if it can rename things and I want
 *    to see if it can organize it in a way, but I want the AI to be able to do it."
 *
 * The first version did not. It named files with regular expressions, which can
 * only rearrange the characters already in a file name — so `Metrics .xlsx`
 * became `Metrics_2025-04-15_Draft_v0.1.xlsx`, the old junk with a date bolted
 * on, wearing the convention as a costume. His format opens with a project code
 * and then a short descriptive name, and neither of those is in the old file
 * name. They are inside the document.
 *
 * So the workflow now reads the documents and asks a model to name them, and
 * the rules that used to do the naming do the checking instead:
 *
 *   the model supplies the WORDS — project, document name, status, date
 *   the code enforces the SHAPE  — format, guards, collisions, validation
 *
 * The model never writes a file name; it returns parts, and `proposeFileName`
 * assembles them. Every proposal is then read back by `diagnose`, and one that
 * does not satisfy the convention is dropped. The worst a bad answer can do is
 * leave a file alone.
 */
const CONFIG = {
  host: TARGET.sharepointHost || 'https://upperformpartners.sharepoint.com',
  sitePath: TARGET.sharepointSite || '/sites/TaskManagement',
  libraryTitle: 'Documents',
  maxItems: 200,

  // Both off by default. A rename or a move breaks every saved link, Teams tab
  // and Power BI data source that points at the old path, so either happens
  // after a person has read the review, not before. They are separate switches
  // because they are separate promises.
  applyRenames: false,
  applyMoves: false,

  // Even switched on, one run touches at most this many files, renames and
  // moves together. A first live run should be small enough to check by eye
  // and undo by hand.
  maxChangesPerRun: 5,

  // The plan a person has read and signed off, as [{ from, to }].
  //
  // The model is not deterministic, and the difference matters. Two runs an
  // hour apart over the same unchanged library proposed
  // Time-Zone-Conversion-Table and then Timezone-Conversion-Table, and
  // Net-Profit-Period-Dashboard and then Profit-Performance-Dashboard. Both are
  // reasonable. Neither is the one on the page that was approved. So when this
  // list is filled in, arming the switch applies exactly these names and
  // ignores what this run's model call said - a review that does not bind to
  // what gets applied is a demonstration, not an approval.
  //
  // Leave it empty and the switches apply whatever the current run proposes,
  // which is fine for a sandbox and wrong for a library someone relies on.
  approvedPlan: approvedPlan(),

  // readContents is the consent switch. With it off the model sees only names,
  // folders and kinds, and nothing from inside a document leaves the tenant.
  // With it on, Office files are unzipped on the server and only their text
  // travels; PDFs and images travel as themselves, because a scan and a
  // screenshot have no text to extract.
  readContents: true,
  modelNaming: 'gpt-5.6-luna',
  maxNamed: 30,
  maxContentBytes: 4194304,
  textPerFile: 1200,

  // Used only where neither the document nor its name says. Draft is the
  // weakest claim, so it under-states rather than over-states, and every
  // assumption is reported beside the proposal. Set to "" to refuse to rename
  // anything whose status is not written down.
  assumeStatus: 'Draft',
  assumeVersion: { major: 0, minor: 1 },
};

/**
 * The reports, and where to fetch each one from.
 *
 * A .pbix is never renamed and never shown to the model. It is downloaded for
 * one reason: to find out what it reads, so the guard can refuse on evidence
 * rather than on two files having similar names.
 */
const PICK_REPORTS = `
const cfg = $('Config').first().json;
const listed = $input.first().json || {};
const items = Array.isArray(listed.value) ? listed.value : [];

const files = [];
for (const it of items) {
  if (!it) continue;
  if (String(it.FSObjType) === '1') continue;
  const name = String(it.FileLeafRef || '');
  if (name) files.push({ name: name, path: String(it.FileRef || '') });
}

const reports = reportsIn(files);
// Always emit one item so the chain runs. A library with no reports has no
// dependencies to discover, which is a finding rather than a reason to branch.
if (reports.length === 0) return [{ json: { name: '', path: '', candidates: [], downloadUrl: '' } }];

const lit = function (p) { return String(p).replace(/'/g, "''").replace(/ /g, '%20'); };
const base = String(cfg.host || '') + String(cfg.sitePath || '') + '/_api/web';

// Every non-report document in the library, so Read Reports can ask each
// report whether it names any of them. Carried on the first item only; the
// rest would be the same list repeated.
const candidates = files.filter(function (f) { return !isReport(f.name); }).map(function (f) { return f.name; });

return reports.map(function (r, i) {
  return { json: {
    name: r.name,
    path: r.path,
    candidates: i === 0 ? candidates : [],
    downloadUrl: base + "/GetFileByServerRelativePath(decodedurl='" + lit(r.path) + "')/$value"
  }};
});
`;

/** The deterministic pass: what each document is, and what must never be touched. */
/**
 * What each report actually reads, taken from the report rather than guessed at
 * from the outside.
 *
 * Matched to its metadata by position, the same way the document contents are.
 * A report that will not open is recorded as unreadable rather than as reading
 * nothing, because those are opposite conclusions.
 */
const READ_REPORTS = `
const meta = $('Pick The Reports').all().map(function (i) { return i.json || {}; });
const got = $input.all();

const opened = [];
for (let i = 0; i < meta.length; i++) {
  const report = meta[i];
  if (!report || !report.name) continue;
  const payload = (got[i] && got[i].json) || {};
  const b64 = typeof payload.contentBytes === 'string' ? payload.contentBytes : '';
  if (!b64) { opened.push({ name: report.name, readable: false, sources: [] }); continue; }
  try {
    opened.push({ name: report.name, bytes: Uint8Array.from(Buffer.from(b64, 'base64')) });
  } catch (err) {
    opened.push({ name: report.name, readable: false, sources: [] });
  }
}

const candidates = (meta[0] && meta[0].candidates) || [];
const dependencies = dependencyIndex(opened, candidates);
return [{ json: {
  candidates: candidates.length,
  dependencies: dependencies,
  read: dependencies.read,
  unreadable: dependencies.unreadable,
  reads: Object.keys(dependencies.byFile).map(function (k) {
    return { file: k, report: dependencies.byFile[k].report, where: dependencies.byFile[k].where };
  })
}}];
`;

const REVIEW = `
const cfg = $('Config').first().json;
const listed = $('List The Library').first().json || {};
const items = Array.isArray(listed.value) ? listed.value : [];

const dependencies = ($input.first().json || {}).dependencies || { byFile: {}, unreadable: [], read: [] };

// FileLeafRef is the name, FileRef the server-relative path. List items span
// every folder in the library, so nested files are seen too.
const files = [];
const folders = [];
for (const it of items) {
  if (!it) continue;
  const name = String(it.FileLeafRef || '');
  if (!name) continue;
  // FSObjType 1 is a folder. Folders are never renamed or moved, but the list
  // of them is what lets a homeless document be offered a home.
  if (String(it.FSObjType) === '1') { folders.push(String(it.FileRef || '')); continue; }
  files.push({ name: name, path: String(it.FileRef || ''), modified: String(it.Modified || '') });
}

const rows = planRenames(files, {
  assumeStatus: cfg.assumeStatus === '' ? '' : String(cfg.assumeStatus || 'Draft'),
  assumeVersion: cfg.assumeVersion || null,
  dependencies: dependencies
});

// What each document is, and whether it is where it should be. Keyed by path
// rather than by position, so the passes cannot drift out of step.
const extra = {};
for (const c of categoriseAll(files, { dependencies: dependencies })) {
  extra[c.path] = { category: c.category, categoryWhy: c.evidence };
}
for (const pl of planPlacements(files, { folders: folders, dependencies: dependencies })) {
  const entry = extra[pl.path] || (extra[pl.path] = {});
  entry.folder = pl.folder;
  entry.placement = pl.verdict;
  entry.placementWhy = pl.reason;
  entry.destination = pl.destination;
}
for (const r of rows) {
  const found = extra[r.path] || {};
  r.category = found.category || '';
  r.categoryWhy = found.categoryWhy || '';
  r.folder = found.folder || '';
  r.placement = found.placement || '';
  r.placementWhy = found.placementWhy || '';
  r.destination = found.destination || '';
  // Where a proposal comes from is part of the answer, so it is recorded
  // rather than inferred from how the name looks.
  r.namedBy = r.proposed ? 'rules' : '';
  r.why = '';
}

return [{ json: {
  library: String(cfg.libraryTitle || ''),
  site: String(cfg.sitePath || ''),
  files: files,
  folders: folders,
  rows: rows,
  dependencies: dependencies,
  reportsRead: dependencies.read,
  reportsUnreadable: dependencies.unreadable
}}];
`;

/** Which documents a model may name, and where to fetch each one from. */
const PICK = `
const cfg = $('Config').first().json;
const review = $input.first().json || {};
const rows = review.rows || [];
const files = review.files || [];

const reports = reportsIn(files);
const candidates = namingCandidates(rows, reports, { dependencies: review.dependencies }).slice(0, Number(cfg.maxNamed) || 30);

// Always emit something, so the branch below runs even when there is nothing
// to name and the summary still gets to speak.
if (candidates.length === 0) return [{ json: { nothing: true, readIt: false } }];

const lit = function (p) { return String(p).replace(/'/g, "''").replace(/ /g, '%20'); };
const base = String(cfg.host || '') + String(cfg.sitePath || '') + '/_api/web';
const readIt = cfg.readContents === true;

return candidates.map(function (r) {
  return { json: {
    nothing: false,
    readIt: readIt,
    name: r.current,
    path: r.path,
    folder: r.folder || '',
    category: r.category || '',
    contentNote: readIt ? '' : 'not read - reading contents is switched off',
    downloadUrl: base + "/GetFileByServerRelativePath(decodedurl='" + lit(r.path) + "')/$value"
  }};
});
`;

/**
 * Turn the downloaded bytes into something a model can use.
 *
 * Office files are unzipped here and only their text travels. PDFs and images
 * travel as themselves, because a scanned page and a screenshot have no text to
 * extract and looking at them is the only way to know what they hold — which is
 * exactly what the three same-day screenshots needed.
 *
 * Items are matched to their metadata by position. The fetch and the base64
 * conversion both emit one item per input item, including when they fail, so
 * position holds where paired-item tracking through a failed HTTP call might
 * not.
 */
const READ_CONTENTS = `
const cfg = $('Config').first().json;
const picks = $('Pick What To Name').all().map(function (i) { return i.json || {}; });
const got = $input.all();
const maxBytes = Number(cfg.maxContentBytes) || 4194304;

const out = [];
for (let i = 0; i < picks.length; i++) {
  const meta = picks[i];
  const row = Object.assign({}, meta, { text: '', dataUri: '', isImage: false, contentNote: '' });
  const payload = (got[i] && got[i].json) || {};
  const b64 = typeof payload.contentBytes === 'string' ? payload.contentBytes : '';

  if (!b64) {
    row.contentNote = 'could not be downloaded, named from its name, folder and kind';
    out.push({ json: row });
    continue;
  }

  // base64 is four characters for every three bytes.
  if (Math.floor(b64.length * 3 / 4) > maxBytes) {
    row.contentNote = 'too large to read, named from its name, folder and kind';
    out.push({ json: row });
    continue;
  }

  const lower = String(meta.name || '').toLowerCase();
  try {
    if (/\\.(docx|pptx|xlsx|xlsm)$/.test(lower)) {
      row.text = String(officeText(Uint8Array.from(Buffer.from(b64, 'base64')), lower) || '').slice(0, 6000);
      if (!row.text) row.contentNote = 'opened but held no readable text';
    } else if (/\\.pdf$/.test(lower)) {
      row.dataUri = 'data:application/pdf;base64,' + b64;
    } else if (/\\.(png|jpg|jpeg|gif|webp)$/.test(lower)) {
      row.isImage = true;
      const mime = /\\.png$/.test(lower) ? 'image/png'
        : /\\.gif$/.test(lower) ? 'image/gif'
        : /\\.webp$/.test(lower) ? 'image/webp'
        : 'image/jpeg';
      row.dataUri = 'data:' + mime + ';base64,' + b64;
    } else if (/\\.(txt|csv|tsv|md|json|xml|html?)$/.test(lower)) {
      row.text = Buffer.from(b64, 'base64').toString('utf8').slice(0, 6000);
    } else {
      row.contentNote = 'not a kind that can be opened, named from its name, folder and kind';
    }
  } catch (err) {
    // A file that will not open is a file named from its path, not a failed run.
    row.contentNote = 'could not be opened, named from its name, folder and kind';
  }

  out.push({ json: row });
}
return out;
`;

const COMPOSE = `
const cfg = $('Config').first().json;
const candidates = $input.all()
  .map(function (i) { return i.json || {}; })
  .filter(function (c) { return c && c.nothing !== true && c.name; });

return [{ json: {
  toName: candidates.length,
  modelRequest: buildNamingRequest(candidates, {
    model: String(cfg.modelNaming || 'gpt-5.6-luna'),
    textPerFile: Number(cfg.textPerFile) || 1200
  })
}}];
`;

/** The model's words, held to the convention, then turned into work to do. */
const APPLY = `
const cfg = $('Config').first().json;
const review = $('Review The Library').first().json || {};
const rows = review.rows || [];
const files = review.files || [];

// Either the model's answer, or nothing at all - because there was nothing to
// name, or because the call failed. Either way the run carries on with what the
// rules could work out, so the naming degrades rather than the run breaking.
const incoming = $input.first().json || {};
let proposals = [];
let modelNote = 'the model was not called, so nothing was renamed beyond what the rules could read';

if (incoming.error) {
  modelNote = 'the model call failed, so the rules named what they could';
} else if (Array.isArray(incoming.output)) {
  let text = '';
  for (const block of incoming.output) {
    for (const part of ((block && block.content) || [])) {
      if (part && typeof part.text === 'string') text += part.text;
    }
  }
  try {
    proposals = JSON.parse(text).proposals || [];
    modelNote = 'the model read and named ' + proposals.length + ' documents';
  } catch (err) {
    proposals = [];
    modelNote = 'the model answer could not be read, so the rules named what they could';
  }
}

const reports = reportsIn(files);
const dependencies = review.dependencies || { byFile: {}, unreadable: [], read: [] };
applyProposals(rows, proposals, {
  files: files,
  reports: reports,
  dependencies: dependencies,
  assumeStatus: cfg.assumeStatus === '' ? '' : String(cfg.assumeStatus || 'Draft')
});

// What a person signed off beats what this run happens to say. Every guard
// still runs over it: the file must still exist, still be un-guarded, keep its
// extension, satisfy the convention and not collide.
const approved = Array.isArray(cfg.approvedPlan) ? cfg.approvedPlan : [];
let planNote = 'this run proposed these names';
let approvedMissing = [];
if (approved.length > 0) {
  applyApprovedPlan(rows, approved, { reports: reports, dependencies: dependencies });
  // A line whose file is not there any more is skipped, which is right - but
  // silently skipping it is how a plan half-applies without anyone noticing.
  approvedMissing = approvedNotFound(rows, approved);
  planNote = 'applying the approved plan of ' + approved.length + ', not this run';
  if (approvedMissing.length > 0) {
    planNote += ' - ' + approvedMissing.length + ' approved file(s) are no longer in the library under that name';
  }
}

const toRename = rows.filter(function (r) { return r.proposed; });
const movable = rows.filter(function (r) { return r.placement === 'move' && r.destination; });

const summary = {
  library: String(review.library || ''),
  site: String(review.site || ''),
  scanned: rows.length,
  compliant: rows.filter(function (r) { return r.compliant; }).length,
  proposed: toRename.length,
  blocked: rows.filter(function (r) { return !r.compliant && !r.proposed; }).length,
  unfiled: rows.filter(function (r) { return r.placement === 'unfiled'; }).length,
  movesProposed: movable.length,
  namedByModel: rows.filter(function (r) { return r.namedBy === 'model'; }).length,
  reportsRead: (dependencies.read || []).length,
  reportsUnreadable: dependencies.unreadable || [],
  approvedCount: approved.length,
  approvedMissing: approvedMissing,
  planNote: planNote,
  modelNote: modelNote,
  readContents: cfg.readContents === true,
  applyRenames: cfg.applyRenames === true,
  applyMoves: cfg.applyMoves === true,
  cap: Number(cfg.maxChangesPerRun) || 5,
  rows: rows
};

// Renaming and moving are the same operation to SharePoint - both change a
// path - but they are different promises, so they are gated separately. A run
// with both switches off reviews and changes nothing, which is the resting
// state.
const actions = [];
if (cfg.applyRenames === true) {
  for (const r of toRename) {
    const folder = r.path.slice(0, r.path.lastIndexOf('/'));
    actions.push({
      kind: 'rename',
      oldName: r.current,
      newName: r.proposed,
      oldPath: r.path,
      newPath: folder + '/' + r.proposed,
      why: r.assumptions || []
    });
  }
}
if (cfg.applyMoves === true) {
  for (const r of movable) {
    actions.push({
      kind: 'move',
      oldName: r.current,
      newName: r.current,
      oldPath: r.path,
      newPath: r.destination,
      why: [r.placementWhy]
    });
  }
}

// Nothing to do still emits one item, so the summary always runs and the
// caller always gets an answer rather than a silent branch.
if (actions.length === 0) {
  return [{ json: Object.assign({ nothingToDo: true }, summary) }];
}

// SharePoint takes the path inside an OData string literal. Single quotes are
// escaped by doubling; spaces have to be percent-encoded or the request line
// breaks. The URL is built here rather than in an expression so the escaping is
// in one readable place.
const lit = function (p) { return String(p).replace(/'/g, "''").replace(/ /g, '%20'); };
const base = String(cfg.host || '') + String(cfg.sitePath || '') + '/_api/web';

return actions.slice(0, Number(cfg.maxChangesPerRun) || 5).map(function (a) {
  return { json: {
    nothingToDo: false,
    kind: a.kind,
    oldName: a.oldName,
    newName: a.newName,
    oldPath: a.oldPath,
    newPath: a.newPath,
    assumptions: a.why,
    // flags=0 means never overwrite. If anything already sits on the new path
    // the call fails rather than destroying it.
    changeUrl: base + "/GetFileByServerRelativePath(decodedurl='" + lit(a.oldPath) + "')/moveto(newurl='" + lit(a.newPath) + "',flags=0)",
    summary: summary
  }};
});
`;

const CONFIRM = `
const planned = $('Apply Proposals').all().map(function (i) { return i.json; });
const results = $input.all().map(function (i) { return i.json || {}; });

const done = [];
const failed = [];
for (let i = 0; i < planned.length; i++) {
  const p = planned[i];
  const r = results[i] || {};
  if (r.error) {
    const blob = JSON.stringify(r.error);
    const m = blob.match(/<m-?:?message[^>]*>([\\s\\S]*?)<\\/m-?:?message>/i);
    failed.push({ kind: p.kind, oldName: p.oldName, newName: p.newName, why: (m ? m[1] : blob).slice(0, 200) });
  } else {
    done.push({ kind: p.kind, oldName: p.oldName, newName: p.newName, newPath: p.newPath, assumptions: p.assumptions });
  }
}

const summary = (planned[0] && planned[0].summary) || {};
return [{ json: Object.assign({}, summary, { applied: done, failedChanges: failed }) }];
`;

/**
 * Returns its result rather than throwing it. An earlier version threw, because
 * that was the only way to read a run back through the MCP at the time, and it
 * was wrong for everyone else: a workflow that signals success by failing shows
 * red on the canvas and reads as broken to whoever is watching it.
 */
const SUMMARY = `
const r = $input.first().json || {};
const applied = r.applied || [];
const failed = r.failedChanges || [];

const headline = applied.length || failed.length
  ? applied.length + ' applied, ' + failed.length + ' failed'
  : r.proposed + ' to rename, ' + r.movesProposed + ' to move, ' +
    r.blocked + ' left alone, ' + r.compliant + ' already correct';

// The reasons are the part worth reading, so they come out as their own lists
// rather than buried inside every row.
const toRename = [];
const refused = [];
const toMove = [];
for (const row of (r.rows || [])) {
  if (row.proposed) {
    toRename.push({
      file: row.current,
      becomes: row.proposed,
      namedBy: row.namedBy || 'rules',
      why: row.why || '',
      assuming: row.assumptions || []
    });
  } else if (!row.compliant && row.blocked) {
    refused.push({ file: row.current, because: row.blocked });
  }
  if (row.placement === 'move' && row.destination) {
    toMove.push({ file: row.current, into: row.destination, because: row.placementWhy });
  }
}

return [{ json: {
  headline: headline,
  library: r.library,
  site: r.site,
  scanned: r.scanned,
  compliant: r.compliant,
  naming: r.modelNote,
  plan: r.planNote,
  namedByModel: r.namedByModel,
  reportsRead: r.reportsRead,
  reportsUnreadable: r.reportsUnreadable || [],
  approvedCount: r.approvedCount,
  approvedMissing: r.approvedMissing || [],
  readContents: r.readContents,
  toRename: toRename,
  toMove: toMove,
  refused: refused,
  unfiled: r.unfiled,
  switches: { applyRenames: r.applyRenames === true, applyMoves: r.applyMoves === true, cap: r.cap },
  applied: applied,
  failed: failed
}}];
`;

export const NAME = 'WF-F · Organise Documents';

const HEADER_NOTE = [
  '## WF-F · Organise Documents',
  '',
  'Reviews a SharePoint library and answers three questions about every document: **what is it**, **what should it be called**, and **is it in the right place**.',
  '',
  '```',
  '[Project]_[Document Name]_[YYYY-MM-DD]_[Status]_v#.#',
  'SCE_PM_Ride-Along-Plan_2026-09-01_Draft_v0.1.docx',
  '```',
  '',
  '### Why a model names them',
  '',
  'The first version used regular expressions, and a regular expression can only rearrange the characters already in a file name. The convention opens with a **project code** and then a **short descriptive name**, and neither of those is in `Metrics .xlsx` — they are inside the document. So it produced `Metrics_2025-04-15_Draft_v0.1.xlsx`: the old junk with a date bolted on, wearing the convention as a costume.',
  '',
  'On the 28 August call the ask was specific: *"I want to see if it can rename things and I want to see if it can organize it in a way, **but I want the AI to be able to do it**."*',
  '',
  '### The division of labour',
  '',
  '| | |',
  '|---|---|',
  '| the model supplies | the **words** — project, document name, status, date |',
  '| the code enforces | the **shape** — format, guards, collisions, validation |',
  '',
  'The model never writes a file name. It returns parts, and `proposeFileName` assembles them, so a malformed name is not something it can produce. Every proposal is then read back by `diagnose`, and one that does not satisfy the convention is dropped.',
  '',
  '### The reports are opened too',
  '',
  'A `.pbix` is never renamed and never shown to the model. It is downloaded for one reason: it is a ZIP, and its `DataMashup` part holds the Power Query text that names every workbook it reads. So the guard no longer refuses a document because it *resembles* a report — it refuses one the report actually names, and quotes the query. A look-alike that no report names is cleared to be renamed.',
  '',
  'A report that will not open is recorded as unreadable rather than as reading nothing, and the old resemblance guess is used for that one alone.',
  '',
  'The judgement lives in `src/naming.js`, `src/filenames.js`, `src/library.js` and `src/powerbi.js`, where it is tested. **Do not edit the code in these nodes** — the next rebuild overwrites it.',
].join('\n');

const READING_NOTE = [
  '### What the model is shown',
  '',
  '`readContents` is the consent switch. With it **off**, the model sees only names, folders and kinds, and nothing from inside a document leaves the tenant.',
  '',
  'With it **on**:',
  '',
  '| Kind | What travels |',
  '|---|---|',
  '| `.docx` `.pptx` `.xlsx` | unzipped on the server; only the extracted text is sent |',
  '| `.pdf` | the file itself, because a scan has no text layer to extract |',
  '| images | the image itself — the only way to tell two same-day screenshots apart |',
  '| everything else | nothing; named from its name, folder and kind |',
  '',
  'A file that will not download or will not open is named from its path rather than failing the run, and the summary says which. `maxNamed` caps how many documents are read at all, `maxContentBytes` refuses anything large, and `textPerFile` caps how much of each one is sent.',
  '',
  '**Power BI never reaches this branch at all.** Reports and the workbooks they read are filtered out before anything is downloaded, so no answer the model could give would be worth the risk.',
  '',
  '**If the model call fails, the run carries on with what the rules could work out.** The naming degrades; it does not break.',
].join('\n');

const SAFETY_NOTE = [
  '### What stops this going wrong',
  '',
  '**Off by default.** `applyRenames` and `applyMoves` are both false, and they are separate switches because they are separate promises. Either one breaks every saved link, Teams tab and Power BI data source pointing at the old path, so a person reads the review first.',
  '',
  '**Capped.** `maxChangesPerRun` limits one run to a handful, renames and moves together, small enough to check by eye and undo by hand.',
  '',
  '**Never overwrites.** The call uses `flags=0`, so if anything already sits on the new path it fails rather than destroying it.',
  '',
  '**Never collides.** Two files in one folder that would land on the same name are both refused — and that is checked again *after* the model answers, because a model can land two files on one name just as easily as a rule can.',
  '',
  '**Never touches Power BI.** A `.pbix` keeps the paths of the files it reads inside itself, where nothing outside can see them. Reports are never renamed, and neither is a workbook that shares a report’s folder or opens with a report’s words. The refusal names the report.',
  '',
  '**Moves one thing, and only on evidence.** A move changes a path exactly as a rename changes a name. It will never move a report, a workbook a report reads, or anything inside a numbered stage folder — that last would assert how far work has got, which no file records. The single case it offers is a document sitting above the filing with no folder of its own, and only where exactly one folder beside it says in its own name what it holds. Two candidates is a choice, and choosing belongs to a person.',
  '',
  '**Never renames a folder.** Anything without a file extension is refused.',
  '',
  '**Never invents.** The model is told an invented project code is worse than none, and an empty answer falls back to a disclosed assumption. Status defaults to the weakest claim, Draft, and says so. An eight-digit run like `04042025` is resolved only if the document states the date unambiguously; otherwise the date comes from the file timestamp and the proposal says so.',
].join('\n');

export async function build() {
  // The mirror of the seeder's refusal to build for production. A target with
  // no SharePoint credential cannot reach the library, and a build that quietly
  // produced a workflow authenticating as nobody would be worse than no build.
  if (!TARGET.sharepointCredential) {
    throw new Error(
      'refusing to build the organiser for a target with no sharepointCredential. ' +
        'Set targets.' +
        TARGET.name +
        '.sharepointCredential in config/deploy.json, or build for a target that owns the library.',
    );
  }

  const pickReportsCode = await codeNode(HELPERS['wf-f/Pick The Reports'], PICK_REPORTS);
  const readReportsCode = await codeNode(HELPERS['wf-f/Read Reports'], READ_REPORTS);
  const reviewCode = await codeNode(HELPERS['wf-f/Review The Library'], REVIEW);
  const pickCode = await codeNode(HELPERS['wf-f/Pick What To Name'], PICK);
  const readCode = await codeNode(HELPERS['wf-f/Read Contents'], READ_CONTENTS);
  const composeCode = await codeNode(HELPERS['wf-f/Compose Request'], COMPOSE);
  const applyCode = await codeNode(HELPERS['wf-f/Apply Proposals'], APPLY);

  const sp = TARGET.sharepointCredential;
  const openai = (DEPLOY.credentials && DEPLOY.credentials.openai && DEPLOY.credentials.openai.name) || 'OpenAI account';

  return `${SDK_IMPORT}

const run = trigger({
  type: 'n8n-nodes-base.manualTrigger',
  version: 1,
  config: { name: 'Run', parameters: {}, position: [0, 0] },
  output: [{}]
});

const config = node({
  type: 'n8n-nodes-base.set',
  version: 3.5,
  config: {
    name: 'Config',
    parameters: { mode: 'raw', jsonOutput: ${S(JSON.stringify(CONFIG, null, 2))} },
    position: [200, 0]
  },
  output: [${JSON.stringify(CONFIG)}]
});

const listFiles = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'List The Library',
    retryOnFail: true,
    maxTries: 3,
    waitBetweenTries: 3000,
    parameters: {
      method: 'GET',
      url: ${E("{{ $('Config').first().json.host }}{{ $('Config').first().json.sitePath }}/_api/web/lists/getbytitle('{{ $('Config').first().json.libraryTitle }}')/items")},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftSharePointOAuth2Api',
      sendQuery: true,
      queryParameters: {
        parameters: [
          { name: '$select', value: 'FileLeafRef,FileRef,Modified,FSObjType' },
          { name: '$top', value: ${E("{{ $('Config').first().json.maxItems }}")} }
        ]
      },
      sendHeaders: true,
      headers: { parameters: [{ name: 'Accept', value: 'application/json;odata=nometadata' }] },
      options: { timeout: 60000 }
    },
    credentials: { microsoftSharePointOAuth2Api: newCredential(${S(sp)}) },
    position: [400, 0]
  },
  output: [{ value: [{ FileLeafRef: 'Metrics .xlsx', FileRef: '/sites/TaskManagement/Shared Documents/Metrics .xlsx', Modified: '2025-04-15T10:00:00Z', FSObjType: 0 }] }]
});

const pickReports = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Pick The Reports',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(pickReportsCode)} },
    position: [620, 0]
  },
  output: [{ name: 'DIMP E-1 Performance Measures.pbix', downloadUrl: 'https://...' }]
});

const fetchReports = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Fetch Reports',
    onError: 'continueRegularOutput',
    parameters: {
      method: 'GET',
      url: ${E('{{ $json.downloadUrl }}')},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftSharePointOAuth2Api',
      options: {
        timeout: 180000,
        response: { response: { responseFormat: 'file', outputPropertyName: 'data' } },
        batching: { batch: { batchSize: 2, batchInterval: 400 } }
      }
    },
    credentials: { microsoftSharePointOAuth2Api: newCredential(${S(sp)}) },
    position: [840, 0]
  },
  output: [{}]
});

const reportsToBase64 = node({
  type: 'n8n-nodes-base.extractFromFile',
  version: 1.1,
  config: {
    name: 'Reports To Base64',
    onError: 'continueRegularOutput',
    parameters: {
      operation: 'binaryToPropery',
      binaryPropertyName: 'data',
      destinationKey: 'contentBytes',
      options: { keepSource: 'json' }
    },
    position: [1060, 0]
  },
  output: [{ contentBytes: 'UEsDBBQ' }]
});

const readReports = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Read Reports',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(readReportsCode)} },
    position: [1280, 0]
  },
  output: [{ read: ['DIMP E-1 Performance Measures.pbix'], unreadable: [], reads: [] }]
});

const review = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Review The Library',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(reviewCode)} },
    position: [1500, 0]
  },
  output: [{ library: 'Documents', files: [], folders: [], rows: [], dependencies: {} }]
});

const pick = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Pick What To Name',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(pickCode)} },
    position: [1720, 0]
  },
  output: [{ nothing: false, readIt: true, name: 'Metrics .xlsx', path: '/sites/TaskManagement/Shared Documents/Metrics .xlsx' }]
});

const anyToName = ifElse({
  version: 2.2,
  config: {
    name: 'Any To Name?',
    parameters: {
      conditions: {
        options: { caseSensitive: true, leftValue: '', typeValidation: 'loose' },
        conditions: [
          { id: 'has-candidates', leftValue: ${E('{{ $json.nothing }}')}, operator: { type: 'boolean', operation: 'false', singleValue: true } }
        ],
        combinator: 'and'
      }
    },
    position: [1940, 0]
  }
});

const readThem = ifElse({
  version: 2.2,
  config: {
    name: 'Read Them?',
    parameters: {
      conditions: {
        options: { caseSensitive: true, leftValue: '', typeValidation: 'loose' },
        conditions: [
          { id: 'read-contents', leftValue: ${E('{{ $json.readIt }}')}, operator: { type: 'boolean', operation: 'true', singleValue: true } }
        ],
        combinator: 'and'
      }
    },
    position: [2160, -180]
  }
});

const fetchContents = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Fetch Contents',
    onError: 'continueRegularOutput',
    parameters: {
      method: 'GET',
      url: ${E('{{ $json.downloadUrl }}')},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftSharePointOAuth2Api',
      options: {
        timeout: 120000,
        response: { response: { responseFormat: 'file', outputPropertyName: 'data' } },
        batching: { batch: { batchSize: 3, batchInterval: 300 } }
      }
    },
    credentials: { microsoftSharePointOAuth2Api: newCredential(${S(sp)}) },
    position: [2380, -300]
  },
  output: [{}]
});

const toBase64 = node({
  type: 'n8n-nodes-base.extractFromFile',
  version: 1.1,
  config: {
    name: 'To Base64',
    onError: 'continueRegularOutput',
    parameters: {
      operation: 'binaryToPropery',
      binaryPropertyName: 'data',
      destinationKey: 'contentBytes',
      options: { keepSource: 'json' }
    },
    position: [2600, -300]
  },
  output: [{ contentBytes: 'UEsDBBQ' }]
});

const readContents = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Read Contents',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(readCode)} },
    position: [2820, -300]
  },
  output: [{ name: 'Metrics .xlsx', text: 'District | Poles | Idling cost' }]
});

const contentsReady = merge({
  version: 3.2,
  config: { name: 'Read Or Not', parameters: { mode: 'append', numberInputs: 2 }, position: [3040, -180] }
});

const composeRequest = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Compose Request',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(composeCode)} },
    position: [3260, -180]
  },
  output: [{ toName: 1, modelRequest: {} }]
});

const nameThem = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Name Them',
    onError: 'continueRegularOutput',
    retryOnFail: true,
    maxTries: 2,
    waitBetweenTries: 5000,
    parameters: {
      method: 'POST',
      url: 'https://api.openai.com/v1/responses',
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'openAiApi',
      sendBody: true,
      contentType: 'json',
      specifyBody: 'json',
      jsonBody: ${E('{{ JSON.stringify($json.modelRequest) }}')},
      options: { timeout: 300000 }
    },
    credentials: { openAiApi: newCredential(${S(openai)}) },
    position: [3480, -180]
  },
  output: [{ output: [{ content: [{ text: '{"proposals":[]}' }] }] }]
});

const namedOrNot = merge({
  version: 3.2,
  config: { name: 'Named Or Not', parameters: { mode: 'append', numberInputs: 2 }, position: [3700, 0] }
});

const applyProposals = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Apply Proposals',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(applyCode)} },
    position: [3920, 0]
  },
  output: [{ nothingToDo: true, scanned: 21, compliant: 6, proposed: 13 }]
});

const shouldApply = ifElse({
  version: 2.2,
  config: {
    name: 'Anything To Apply?',
    parameters: {
      conditions: {
        options: { caseSensitive: true, leftValue: '', typeValidation: 'loose' },
        conditions: [
          { id: 'has-changes', leftValue: ${E('{{ $json.nothingToDo }}')}, operator: { type: 'boolean', operation: 'false', singleValue: true } }
        ],
        combinator: 'and'
      }
    },
    position: [4140, 0]
  }
});

const applyChange = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Apply The Change',
    onError: 'continueRegularOutput',
    parameters: {
      method: 'POST',
      url: ${E('{{ $json.changeUrl }}')},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftSharePointOAuth2Api',
      sendHeaders: true,
      headers: { parameters: [{ name: 'Accept', value: 'application/json;odata=nometadata' }] },
      options: { timeout: 60000, batching: { batch: { batchSize: 1, batchInterval: 400 } } }
    },
    credentials: { microsoftSharePointOAuth2Api: newCredential(${S(sp)}) },
    position: [4360, -180]
  },
  output: [{ ok: true }]
});

const confirm = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Confirm What Changed',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(CONFIRM)} },
    position: [4580, -180]
  },
  output: [{ applied: [], failedChanges: [] }]
});

const both = merge({
  version: 3.2,
  config: { name: 'Either Way', parameters: { mode: 'append', numberInputs: 2 }, position: [4800, 0] }
});

const summary = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Summary',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(SUMMARY)} },
    position: [5020, 0]
  },
  output: [{ headline: '13 to rename, 0 to move, 2 left alone, 6 already correct' }]
});

const noteHeader = sticky(
  ${S(HEADER_NOTE)},
  [run, config, listFiles, pickReports, fetchReports, reportsToBase64, readReports, review, pick],
  { color: 6, width: 1680, height: 460 },
);

const noteReading = sticky(
  ${S(READING_NOTE)},
  [readThem, fetchContents, toBase64, readContents, contentsReady, composeRequest, nameThem],
  { color: 4, width: 1560, height: 340 },
);

const noteSafety = sticky(
  ${S(SAFETY_NOTE)},
  [namedOrNot, applyProposals, shouldApply, applyChange, confirm, both, summary],
  { color: 5, width: 1540, height: 480 },
);

export default workflow('upperform-wf-f-organizer', ${S(NAME)})
  .add(run)
  .to(config)
  .to(listFiles)
  .to(pickReports)
  .to(fetchReports)
  .to(reportsToBase64)
  .to(readReports)
  .to(review)
  .to(pick)
  .to(anyToName
    .onTrue(readThem
      .onTrue(fetchContents.to(toBase64).to(readContents).to(contentsReady.input(0)))
      .onFalse(contentsReady.input(1)))
    .onFalse(namedOrNot.input(1)))
  .add(contentsReady)
  .to(composeRequest)
  .to(nameThem)
  .to(namedOrNot.input(0))
  .add(namedOrNot)
  .to(applyProposals)
  .to(shouldApply
    .onTrue(applyChange.to(confirm).to(both.input(0)))
    .onFalse(both.input(1)))
  .add(both)
  .to(summary)
  .add(noteHeader)
  .add(noteReading)
  .add(noteSafety);
`;
}
