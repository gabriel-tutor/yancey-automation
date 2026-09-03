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
    // A line approves a name, a destination, or both, so both are conditional.
    // `String(undefined)` is the string "undefined", which would travel into
    // the workflow as an approved file name and be refused as one.
    return (file.plan || []).map((/** @type {any} */ line) => ({
      from: String(line.from),
      to: line.to ? String(line.to) : '',
      into: line.into ? String(line.into) : '',
    }));
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
export const CONFIG = {
  host: TARGET.sharepointHost || 'https://upperformpartners.sharepoint.com',
  sitePath: TARGET.sharepointSite || '/sites/TaskManagement',
  libraryTitle: 'Documents',

  // One page of list items, and SharePoint's own ceiling for a single request.
  // It was 200, which is fine for this library and wrong as a default: a name
  // checked for collisions against half a library is not checked at all. If a
  // library ever exceeds this the run reads what it can, reports that it did,
  // and applies nothing.
  maxItems: 5000,

  // Both off by default. A rename or a move breaks every saved link, Teams tab
  // and Power BI data source that points at the old path, so either happens
  // after a person has read the review, not before. They are separate switches
  // because they are separate promises.
  applyRenames: false,
  applyMoves: false,

  // The folder structure the model invents. Creating a folder is the least
  // destructive thing here - nothing that exists changes - but moving documents
  // into it is not, so applyStructure gates the creating and applyMoves gates
  // the filing. Renaming an existing folder moves everything inside it, which
  // is why it is a third switch rather than part of the second.
  applyStructure: false,
  applyFolderRenames: false,

  // Freedom, a little bit. Yancey asked for a structure the agent invents
  // rather than one he specifies - "I wanted to give it freedom a little bit on
  // the folder structure" - and these are what stop that becoming twenty
  // folders of one document. A new folder must hold at least minPerFolder
  // documents, and one run proposes at most maxNewFolders of them.
  maxNewFolders: 6,
  minPerFolder: 2,

  // Even switched on, one run touches at most this many things - folders
  // created, files renamed and files moved, all counted together. It was five
  // while this only renamed; a grouping is now one folder plus every document
  // that goes in it, and a cap that cuts a folder off halfway leaves the
  // library in a state nobody chose. Still small enough to check by eye and
  // undo by hand.
  maxChangesPerRun: 12,

  // Undo. With this on, approvedPlan is read as "put these names back": the
  // convention is not applied to the name being restored, because that name is
  // the messy one and refusing it is the reason it was changed. Every other
  // guard still holds. It is how a rehearsal is undone, and how a change is
  // reversed without touching SharePoint by hand.
  revert: false,

  // Whether the Power BI reports themselves are renamed.
  //
  // The reports in this library read nothing from it - that is established by
  // opening them, not guessed at. What remains true is that a .pbix is
  // referenced from *outside* the library, by whatever published dataset or
  // refresh schedule was set up from it, and neither the library nor the file
  // can see that. So this is a decision rather than a default, and the undo
  // list is what makes it reversible.
  renameReports: true,

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

export const REVIEW = `
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
  dependencies: dependencies,
  renameReports: cfg.renameReports === true
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
export const PICK = `
const cfg = $('Config').first().json;
const review = $input.first().json || {};
const rows = review.rows || [];
const files = review.files || [];

const reports = reportsIn(files);
// forPlacement lets a document with no extension be shown: it cannot be given
// a name in a convention that ends in one, but it can be put in the right
// folder, and a file the organiser can never touch at all is a gap.
const candidates = namingCandidates(rows, reports, {
  dependencies: review.dependencies,
  includeReports: cfg.renameReports === true,
  forPlacement: true
}).slice(0, Number(cfg.maxNamed) || 30);

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

export const COMPOSE = `
const cfg = $('Config').first().json;
const review = $('Review The Library').first().json || {};
const candidates = $input.all()
  .map(function (i) { return i.json || {}; })
  .filter(function (c) { return c && c.nothing !== true && c.name; });

return [{ json: {
  toName: candidates.length,
  modelRequest: buildNamingRequest(candidates, {
    model: String(cfg.modelNaming || 'gpt-5.6-luna'),
    textPerFile: Number(cfg.textPerFile) || 1200,
    // It cannot reuse a folder name it has never seen, and reusing one is
    // almost always the better answer than inventing another.
    folders: review.folders || []
  })
}}];
`;

/**
 * The model's words, held to the convention, then turned into work to do.
 *
 * Exported because this is the node that failed in front of the client on
 * 3 September, in three separate ways, and a string is not something a test can
 * run. src/nodes.test.js runs it.
 */
export const APPLY = `
const cfg = $('Config').first().json;
const review = $('Review The Library').first().json || {};
const rows = review.rows || [];
const files = review.files || [];

// Either the model's answer, or nothing at all - because there was nothing to
// name, or because the call failed. Either way the run carries on with what the
// rules could work out, so the naming degrades rather than the run breaking.
const incoming = $input.first().json || {};
let proposals = [];
let modelFolders = [];
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
    const answer = JSON.parse(text);
    proposals = answer.proposals || [];
    modelFolders = answer.folders || [];
    modelNote = 'the model read and named ' + proposals.length + ' documents';
  } catch (err) {
    proposals = [];
    modelFolders = [];
    modelNote = 'the model answer could not be read, so the rules named what they could';
  }
}

// A library bigger than one page of list items is a library this run has only
// partly seen, and a name checked for collisions against half a library is not
// checked at all. Nothing is applied while this is true.
//
// The count comes from the listing itself rather than from the files and
// folders it was sorted into: an item with no name is dropped on the way, and
// counting after that could report a full page as a short one and apply changes
// against a library it had not finished reading.
const listing = $('List The Library').first().json || {};
const listed = Array.isArray(listing.value) ? listing.value.length : files.length + (review.folders || []).length;
const truncated = listed >= (Number(cfg.maxItems) || 200);

const reports = reportsIn(files);
const dependencies = review.dependencies || { byFile: {}, unreadable: [], read: [] };
applyProposals(rows, proposals, {
  files: files,
  reports: reports,
  dependencies: dependencies,
  includeReports: cfg.renameReports === true,
  assumeStatus: cfg.assumeStatus === '' ? '' : String(cfg.assumeStatus || 'Draft')
});

// What a person signed off beats what this run happens to say. Every guard
// still runs over it: the file must still exist, still be un-guarded, keep its
// extension, satisfy the convention and not collide.
//
// What a plan gates is what gets APPLIED. It must not gate what gets PROPOSED:
// on 3 September eight new documents came back "not on the approved list, so
// this run leaves it alone" after the model had read and named every one, and
// the review loop had nothing new to review. Off-list proposals are kept and
// marked unapproved instead.
const approved = Array.isArray(cfg.approvedPlan) ? cfg.approvedPlan : [];
const pinned = approved.length > 0;
let planNote = 'this run proposed these names, and arming the switch applies them';
let approvedMissing = [];
let approvedDone = [];
if (pinned) {
  applyApprovedPlan(rows, approved, {
    reports: reports,
    dependencies: dependencies,
    includeReports: cfg.renameReports === true,
    revert: cfg.revert === true
  });
  // A line whose file is not there any more is skipped, which is right - but
  // silently skipping it is how a plan half-applies without anyone noticing.
  //
  // "Not there" has two causes and they are opposite conclusions. A line whose
  // RESULT is in the library was carried out. Only a line with neither its file
  // nor its result is adrift, and reporting the two as one made every run after
  // a successful one read as a failure.
  const progress = approvedProgress(rows, approved);
  approvedMissing = progress.missing;
  approvedDone = progress.done;
  planNote = (cfg.revert === true ? 'REVERTING the approved plan of ' : 'applying the approved plan of ') + approved.length + ' - everything else is proposed and waiting';
  if (approvedDone.length > 0) {
    planNote += ' - ' + approvedDone.length + ' already carried out on an earlier run';
  }
  if (approvedMissing.length > 0) {
    planNote += ' - ' + approvedMissing.length + ' approved file(s) are no longer in the library under that name';
  }
}

// The structure the model invented, held to a small one. This always runs, so
// the review page shows what it came up with; the switches below decide whether
// any of it happens. Yancey asked for exactly this on 3 September: "Can it
// think and come up with its own folder structure?" - and then bounded it
// himself: "I wanted to give it freedom a little bit."
const structure = planStructure(rows, {
  folders: review.folders || [],
  files: files,
  maxNewFolders: Number(cfg.maxNewFolders) || 6,
  minPerFolder: Number(cfg.minPerFolder) || 2,
  includeReports: cfg.renameReports === true
});
// Nothing that follows may plan a change that cannot land. Two documents
// composed onto one path, a document already sitting where another wants to be,
// a path past SharePoint's 400-character ceiling: each is a call that fails, and
// finding out by sending it is not a review.
refuseImpossible(rows, files);

// Every move needs somewhere to land, and this asks once for all of them rather
// than trusting whichever step proposed the move to have proposed the folder
// too. A folder is not a separate decision - it is the other half of the move -
// so this covers what a plan pinned and what the agent invented alike, and a
// folder nothing lands in any more is not created at all.
//
// A folder holding a document certainly exists whether or not the listing named
// it, so the parents of every file count as known.
const knownFolders = (review.folders || []).concat(files.map(function (f) {
  return folderOf(String(f.path || ''));
}));
// cancelUnreachable is the other half of naming an unreachable destination:
// reporting it and then sending the move anyway would be the worst of both.
const needed = foldersNeeded(rows, knownFolders, { cancelUnreachable: true });
const wantedFolders = needed.create;

// The folders about to be created count as existing for the rename check, or a
// run could propose creating "Dashboards" and renaming another folder to
// "Dashboards" in the same breath, and whichever ran second would fail.
const folderPlan = planFolderRenames(modelFolders, {
  folders: (review.folders || []).concat(wantedFolders),
  files: files,
  includeReports: cfg.renameReports === true
});

// Documents past maxNamed were never shown to the model, so nothing was
// proposed for them - which is a fact about this run rather than about them,
// and reads as "nothing needed doing" unless it is said out loud.
const couldBeNamed = namingCandidates(rows, reports, {
  dependencies: dependencies,
  includeReports: cfg.renameReports === true,
  forPlacement: true
}).length;
const beyondModel = Math.max(0, couldBeNamed - (Number(cfg.maxNamed) || 30));

const toRename = rows.filter(function (r) { return r.proposed && (!pinned || r.approved === true); });
const awaiting = rows.filter(function (r) { return r.proposed && pinned && r.approved !== true; });
const movable = rows.filter(function (r) { return r.placement === 'move' && r.destination; });

const summary = {
  library: String(review.library || ''),
  site: String(review.site || ''),
  scanned: rows.length,
  compliant: rows.filter(function (r) { return r.compliant; }).length,
  proposed: toRename.length,
  awaiting: awaiting.length,
  blocked: rows.filter(function (r) { return !r.compliant && !r.proposed; }).length,
  unfiled: rows.filter(function (r) { return r.placement === 'unfiled'; }).length,
  movesProposed: movable.length,
  namedByModel: rows.filter(function (r) { return r.namedBy === 'model'; }).length,
  reportsRead: (dependencies.read || []).length,
  reportsUnreadable: dependencies.unreadable || [],
  approvedCount: approved.length,
  approvedMissing: approvedMissing,
  approvedDone: approvedDone,
  pinned: pinned,
  planNote: planNote,
  modelNote: modelNote,
  truncated: truncated,
  listed: listed,
  beyondModel: beyondModel,
  newFolders: wantedFolders,
  unreachableFolders: needed.unreachable,
  impossibleSwaps: [],
  folderRefusals: structure.refused,
  folderRenames: folderPlan.renames,
  folderRenameRefusals: folderPlan.refused,
  readContents: cfg.readContents === true,
  applyRenames: cfg.applyRenames === true,
  applyMoves: cfg.applyMoves === true,
  applyStructure: cfg.applyStructure === true,
  applyFolderRenames: cfg.applyFolderRenames === true,
  cap: Number(cfg.maxChangesPerRun) || 5,
  rows: rows
};

// Renaming and moving are the same operation to SharePoint - both change a
// path - but they are different promises, so they are gated separately. A run
// with every switch off reviews and changes nothing, which is the resting
// state.
//
// A file gets ONE call, to its new folder under its new name. An earlier
// version built renames and moves as separate lists, so a document wanting both
// produced two calls that each started from the path it had before: the first
// worked and the second could not find the file.
const base = String(cfg.host || '') + String(cfg.sitePath || '') + '/_api/web';
const creating = cfg.applyStructure === true;
const newFolders = creating ? wantedFolders : [];

// A move into a folder that does not exist yet cannot happen unless that folder
// is being created on this run - and that holds however deep the destination
// is, not just directly inside it. Filing into a folder that is already there is
// gated by applyMoves alone.
const notBeingCreated = creating ? [] : wantedFolders;
const wouldLandNowhere = function (folder) {
  return notBeingCreated.some(function (f) { return folder === f || folder.indexOf(f + '/') === 0; });
};

// A document may move onto a path another is vacating - legal, and useful when
// two documents swap folders - but only in that order, and flags=0 means the
// wrong order is a failed call. orderChanges sequences them and refuses a ring,
// which is two documents trading places and cannot be done without somewhere to
// put one of them first.
const wanted = plannedChanges(rows, {
  applyRenames: cfg.applyRenames === true,
  applyMoves: cfg.applyMoves === true,
  requireApproval: pinned
}).filter(function (c) {
  return !wouldLandNowhere(folderOf(c.newPath));
});
const sequenced = orderChanges(wanted);
const fileChanges = sequenced.ordered;
summary.impossibleSwaps = sequenced.impossible.map(function (c) {
  return { file: c.oldName, because: 'it would trade places with another document, and neither can move until one of them has somewhere else to be' };
});

// Order is the whole of the safety here. Folders are created before anything
// moves into them, and folders are renamed LAST - everything inside a folder
// moves with it, so a file change computed against the old path has to happen
// while that path is still real.
//
// Every action carries the group it belongs to, because a folder and the
// documents that go into it have to be applied together or not at all: a cap
// that fell between them would create a folder and leave its documents outside.
//
// A whole chain counts as one group. Creating "Dashboards" and "Dashboards/
// Weekly" and moving a document into the second is three actions and one
// decision - grouped separately, the cap could take the outer folder and the
// move and skip the inner one, which is a move into a folder that is not there.
// newFolders is shallowest-first, so the first that owns a path is its root.
const groupFor = function (folder) {
  const owning = newFolders.filter(function (f) { return folder === f || folder.indexOf(f + '/') === 0; });
  return owning.length > 0 ? 'folder:' + owning[0] : '';
};

const actions = [];
for (const path of newFolders) {
  actions.push({
    group: groupFor(path),
    kind: 'create folder',
    oldName: '',
    newName: path.slice(path.lastIndexOf('/') + 1),
    oldPath: '',
    newPath: path,
    why: ['grouping documents the model read and put together'],
    changeUrl: folderCreateUrl(base, path)
  });
}
for (const c of fileChanges) {
  // A chain of documents each taking the last one's place is one decision, so a
  // cap cannot send the second half of it and none of the first.
  actions.push({
    group: groupFor(folderOf(c.newPath)) || 'chain:' + c.chain,
    kind: c.kind,
    oldName: c.oldName,
    newName: c.newName,
    oldPath: c.oldPath,
    newPath: c.newPath,
    why: c.why,
    changeUrl: fileChangeUrl(base, c.oldPath, c.newPath)
  });
}
if (cfg.applyFolderRenames === true) {
  for (const r of folderPlan.renames) {
    actions.push({
      group: 'rename:' + r.from,
      kind: 'rename folder',
      oldName: r.from.slice(r.from.lastIndexOf('/') + 1),
      newName: r.to.slice(r.to.lastIndexOf('/') + 1),
      oldPath: r.from,
      newPath: r.to,
      why: [r.why],
      changeUrl: folderChangeUrl(base, r.from, r.to)
    });
  }
}

// A library only partly listed is a library whose collisions are only partly
// known, so nothing is sent until the whole of it has been read.
const fitting = truncated
  ? { send: [], held: [] }
  : fittingActions(actions, Number(cfg.maxChangesPerRun) || 5);
summary.heldBack = fitting.held;
if (truncated && actions.length > 0) {
  summary.planNote += ' - NOTHING WAS APPLIED: the library has at least ' + summary.listed +
    ' items, which is everything this run was allowed to list, so it has not seen all of it. Raise maxItems.';
}

// Nothing to do still emits one item, so the summary always runs and the
// caller always gets an answer rather than a silent branch.
if (fitting.send.length === 0) {
  return [{ json: Object.assign({ nothingToDo: true }, summary) }];
}

return fitting.send.map(function (a) {
  return { json: {
    nothingToDo: false,
    kind: a.kind,
    oldName: a.oldName,
    newName: a.newName,
    oldPath: a.oldPath,
    newPath: a.newPath,
    assumptions: a.why,
    changeUrl: a.changeUrl,
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
export const SUMMARY = `
const r = $input.first().json || {};
const applied = r.applied || [];
const failed = r.failedChanges || [];

const headline = applied.length || failed.length
  ? applied.length + ' applied, ' + failed.length + ' failed'
  : r.proposed + ' to rename, ' + r.movesProposed + ' to move, ' +
    (r.newFolders || []).length + ' new folder(s), ' +
    r.awaiting + ' waiting to be approved, ' +
    r.blocked + ' left alone, ' + r.compliant + ' already correct';

// The reasons are the part worth reading, so they come out as their own lists
// rather than buried inside every row.
//
// Two lists, not one. A pinned plan governs what is applied; everything else
// the model proposed is still a proposal, and putting it in front of a person
// is the only way the review loop ever closes. The run that showed one list
// reported eight documents as refused when what had actually happened was that
// eight documents were waiting to be read.
const toRename = [];
const awaiting = [];
const refused = [];
const toMove = [];
for (const row of (r.rows || [])) {
  if (row.proposed) {
    const line = {
      file: row.current,
      becomes: row.proposed,
      namedBy: row.namedBy || 'rules',
      why: row.why || '',
      assuming: row.assumptions || []
    };
    if (r.pinned && row.approved !== true) awaiting.push(line);
    else toRename.push(line);
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
  approvedDone: r.approvedDone || [],
  // Facts about what this run could NOT see or do. A page that hides these is
  // a page claiming a completeness the run did not have.
  truncated: r.truncated === true,
  listed: r.listed,
  beyondModel: r.beyondModel || 0,
  heldBack: r.heldBack || [],
  unreachableFolders: r.unreachableFolders || [],
  impossibleSwaps: r.impossibleSwaps || [],
  readContents: r.readContents,
  toRename: toRename,
  awaitingApproval: awaiting,
  toMove: toMove,
  refused: refused,
  // What the agent came up with when asked to organise rather than to rename.
  newFolders: r.newFolders || [],
  folderRenames: r.folderRenames || [],
  folderRefusals: (r.folderRefusals || []).concat(r.folderRenameRefusals || []),
  unfiled: r.unfiled,
  switches: {
    applyRenames: r.applyRenames === true,
    applyMoves: r.applyMoves === true,
    applyStructure: r.applyStructure === true,
    applyFolderRenames: r.applyFolderRenames === true,
    cap: r.cap
  },
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

export const STRUCTURE_NOTE = [
  '### It comes up with the folder structure',
  '',
  'From the 3 September call: *"Can it come up with its own folder structure and rename the files and put it in its own folder? **Can it think and come up with its own folder structure?**"* — and, when offered a template to fill in instead: *"No, **I wanted to give it freedom a little bit** on the folder structure."*',
  '',
  'So the model is given the documents and the folders that already exist, and it answers with a folder for each document and a rename for any folder whose name does not say what it holds. **Freedom is its half.** *A little bit* is the part a machine has to supply, because a model told to group twenty documents will happily return twenty folders of one — which is not a structure, it is the same pile with more clicks in it.',
  '',
  '| The rule | Why |',
  '|---|---|',
  '| a **new** folder must hold at least `minPerFolder` (2) | a folder holding one file is the pile it came from with an extra click in front |',
  '| at most `maxNewFolders` (6) in one run | freedom, a little bit |',
  '| created **beside** the documents that go in it | a bounded change, never a global reshuffle |',
  '| an existing folder is exempt from the minimum | filing one document into a folder somebody already made is filing, not restructuring |',
  '| a name that is already used twice is refused | two folders of one name is a choice, and choosing belongs to a person |',
  '| never a numbered stage folder, in or out | those record how far work has got, and nothing about a document says which stage it is in |',
  '',
  'Everything it proposes appears on the review whether the switches are armed or not. `applyStructure` creates the folders, `applyMoves` files the documents into them, and `applyFolderRenames` renames an existing one — three switches because they are three different promises, and renaming a folder moves everything inside it.',
].join('\n');

export const SAFETY_NOTE = [
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
  '**Moves on evidence.** A move changes a path exactly as a rename changes a name. It will never move a report, a workbook a report reads, or anything into or out of a numbered stage folder — that last would assert how far work has got, which no file records.',
  '',
  '**Never applies what nobody approved.** When `approvedPlan` is filled in, only those lines are sent. Everything else the model proposed stays a **proposal** on the review, under `awaitingApproval`. An earlier version discarded them, and the run reported eight freshly-named documents as refused when what had really happened was that eight documents were waiting to be read.',
  '',
  '**One call per file.** A document wanting a new name *and* a new folder gets a single call, to the new folder under the new name. Two calls both starting from the old path meant whichever ran second could not find the file.',
  '',
  '**A short URL.** Both paths travel in the query string behind OData parameter aliases. `moveto` carries the path twice, and ASP.NET caps the URL *path* at 260 characters — a folder with a long name was enough to break it.',
  '',
  '**Folders in order.** New folders are created first, files change next, and folders are renamed last — everything inside a folder moves with it, so a file change has to happen while its path is still real.',
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

const noteStructure = sticky(
  ${S(STRUCTURE_NOTE)},
  [namedOrNot, applyProposals],
  { color: 3, width: 700, height: 520 },
);

const noteSafety = sticky(
  ${S(SAFETY_NOTE)},
  [shouldApply, applyChange, confirm, both, summary],
  { color: 5, width: 1300, height: 600 },
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
  .add(noteStructure)
  .add(noteSafety);
`;
}
