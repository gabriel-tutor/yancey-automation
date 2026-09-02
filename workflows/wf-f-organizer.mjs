import { codeNode } from '../scripts/bundle.mjs';
import { HELPERS } from './helpers-manifest.mjs';
import { E, S, SDK_IMPORT } from './shared.mjs';

/**
 * WF-F reads a SharePoint document library, works out what each file should be
 * called under Yancey's convention, and — only when told to — renames them.
 *
 * Renaming is the destructive part and is treated as such. It is off by
 * default, capped per run, refuses to overwrite anything, refuses when two
 * files would land on the same name, and verifies afterwards that the new name
 * exists. Every decision about *what* a file should be called is made in
 * src/filenames.js, where it is tested against his four worked examples and
 * against the real contents of this library.
 */
const CONFIG = {
  host: 'https://upperformpartners.sharepoint.com',
  sitePath: '/sites/TaskManagement',
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

  // Status is a claim about whether a document is approved. Draft is the
  // weakest such claim, so assuming it under-states rather than over-states,
  // and every assumption is reported alongside the proposal. Set to "" to
  // refuse to rename anything whose status is not written in the name.
  assumeStatus: 'Draft',
  assumeVersion: { major: 0, minor: 1 },
};

const PLAN = `
const cfg = $('Config').first().json;
const listed = $input.first().json || {};
const items = Array.isArray(listed.value) ? listed.value : [];

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
  assumeVersion: cfg.assumeVersion || null
});

// What each document is, and whether it is where it should be. Keyed by path
// rather than by position, so the three passes cannot drift out of step.
const extra = {};
for (const c of categoriseAll(files)) {
  extra[c.path] = { category: c.category, categoryWhy: c.evidence };
}
for (const pl of planPlacements(files, { folders: folders })) {
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
}

const compliant = rows.filter(function (r) { return r.compliant; });
const proposals = rows.filter(function (r) { return r.proposed; });
const blocked = rows.filter(function (r) { return !r.compliant && !r.proposed; });
const unfiled = rows.filter(function (r) { return r.placement === 'unfiled'; });
const movable = rows.filter(function (r) { return r.placement === 'move' && r.destination; });

const summary = {
  library: String(cfg.libraryTitle || ''),
  site: String(cfg.sitePath || ''),
  scanned: rows.length,
  compliant: compliant.length,
  proposed: proposals.length,
  blocked: blocked.length,
  unfiled: unfiled.length,
  movesProposed: movable.length,
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
  for (const r of proposals) {
    const folder = r.path.slice(0, r.path.lastIndexOf('/'));
    actions.push({
      kind: 'rename',
      oldName: r.current,
      newName: r.proposed,
      oldPath: r.path,
      newPath: folder + '/' + r.proposed,
      why: r.assumptions
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

const take = actions.slice(0, Number(cfg.maxChangesPerRun) || 5);

return take.map(function (a) {
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
const planned = $('Plan Renames').all().map(function (i) { return i.json; });
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

const REPORT = `
const r = $input.first().json || {};

// This node used to throw its result. That was the only way to read a run back
// through the MCP at the time, and it was wrong for everyone else: a workflow
// that signals success by failing shows red on the canvas and reads as broken
// to whoever is watching it. Execution data is readable directly now, so it
// returns data like any other node.
const applied = r.applied || [];
const failed = r.failedChanges || [];

const headline = applied.length || failed.length
  ? applied.length + ' applied, ' + failed.length + ' failed'
  : r.proposed + ' to rename, ' + r.movesProposed + ' to move, ' +
    r.blocked + ' left alone, ' + r.compliant + ' already correct';

// The reasons are the part worth reading, so they come out as their own list
// rather than buried inside every row.
const refused = [];
for (const row of (r.rows || [])) {
  if (row.compliant || row.proposed) continue;
  if (!row.blocked) continue;
  refused.push({ file: row.current, because: row.blocked });
}

const toRename = [];
for (const row of (r.rows || [])) {
  if (!row.proposed) continue;
  toRename.push({ file: row.current, becomes: row.proposed, assuming: row.assumptions || [] });
}

const toMove = [];
for (const row of (r.rows || [])) {
  if (row.placement !== 'move' || !row.destination) continue;
  toMove.push({ file: row.current, into: row.destination, because: row.placementWhy });
}

return [{
  json: {
    headline: headline,
    library: r.library,
    site: r.site,
    scanned: r.scanned,
    compliant: r.compliant,
    toRename: toRename,
    toMove: toMove,
    refused: refused,
    unfiled: r.unfiled,
    switches: { applyRenames: r.applyRenames === true, applyMoves: r.applyMoves === true, cap: r.cap },
    applied: applied,
    failed: failed
  }
}];
`;

export const NAME = 'WF-F · Organise Documents';

export async function build() {
  const planCode = await codeNode(HELPERS['wf-f/Plan Renames'], PLAN);

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
    credentials: { microsoftSharePointOAuth2Api: newCredential('Microsoft SharePoint account') },
    position: [400, 0]
  },
  output: [{ value: [{ FileLeafRef: 'Working TimesKV.xlsx', FileRef: '/sites/TaskManagement/Shared Documents/Working TimesKV.xlsx', Modified: '2025-05-01T10:00:00Z', FSObjType: 0 }] }]
});

const plan = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Plan Renames',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(planCode)} },
    position: [620, 0]
  },
  output: [{ nothingToDo: false, kind: 'rename', oldName: 'Working TimesKV.xlsx', newName: 'Working-TimesKV_2025-05-01_Draft_v0.1.xlsx' }]
});

const shouldApply = ifElse({
  version: 2.2,
  config: {
    name: 'Apply Them?',
    parameters: {
      conditions: {
        options: { caseSensitive: true, leftValue: '', typeValidation: 'loose' },
        conditions: [
          { id: 'has-renames', leftValue: ${E('{{ $json.nothingToDo }}')}, operator: { type: 'boolean', operation: 'false', singleValue: true } }
        ],
        combinator: 'and'
      }
    },
    position: [840, 0]
  }
});

const rename = node({
  type: 'n8n-nodes-base.httpRequest',
  version: 4.5,
  config: {
    name: 'Rename In SharePoint',
    onError: 'continueRegularOutput',
    parameters: {
      method: 'POST',
      url: ${E('{{ $json.changeUrl }}')},
      authentication: 'predefinedCredentialType',
      nodeCredentialType: 'microsoftSharePointOAuth2Api',
      sendHeaders: true,
      headers: { parameters: [{ name: 'Accept', value: 'application/json;odata=nometadata' }] },
      options: {
        timeout: 60000,
        batching: { batch: { batchSize: 1, batchInterval: 400 } }
      }
    },
    credentials: { microsoftSharePointOAuth2Api: newCredential('Microsoft SharePoint account') },
    position: [1060, -120]
  },
  output: [{ ok: true }]
});

const confirm = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Confirm What Moved',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(CONFIRM)} },
    position: [1280, -120]
  },
  output: [{ applied: [], failedChanges: [] }]
});

const both = merge({
  version: 3.2,
  config: { name: 'Either Way', parameters: { mode: 'append', numberInputs: 2 }, position: [1500, 0] }
});

const report = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Report',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(REPORT)} },
    position: [1720, 0]
  },
  output: [{}]
});

const noteHeader = sticky(
  ${S('## WF-F · Organise Documents\n\nReviews a SharePoint library. For every document it answers three questions: **what is it**, **what should it be called**, and **is it in the right place**. It renames only when told to, and it never moves anything.\n\n```\n[Project]_[Document]_[YYYY-MM-DD]_[Status]_v#.#\nSCE_PM_Ride-Along-Plan_2026-09-01_Draft_v0.1.docx\n```\n\nWhat a file should be *called* lives in `src/filenames.js`. What it *is* and *where it sits* live in `src/library.js`, because both depend on the other files around it. Both are tested against his four worked examples and against the real contents of this library. These nodes only carry SharePoint\'s fields in and the answers back out — **do not edit the code here**, the next rebuild overwrites it.\n\n**Nothing is ever opened.** Every judgement comes from the file name, its extension, and what sits beside it. Reading the contents of a client library would mean sending those contents somewhere, and that is a decision to take knowingly.')},
  [run, config, listFiles, plan],
  { color: 6, width: 820, height: 330 },
);

const noteSafety = sticky(
  ${S('### What stops this going wrong\n\n**Off by default.** `applyRenames` and `applyMoves` are both false, and they are separate switches because they are separate promises. Either one breaks every saved link, Teams tab and Power BI data source pointing at the old path, so a person reads the review first.\n\n**Capped.** `maxChangesPerRun` limits one run to a handful, renames and moves together, small enough to check by eye and undo by hand.\n\n**Never overwrites.** The move uses `flags=0`, so if anything already sits on the new name the call fails rather than destroying it.\n\n**Never collides.** Two files in one folder that would land on the same name are both refused. Scoped to the folder, because that is where SharePoint requires a name to be unique.\n\n**Never touches Power BI.** A `.pbix` keeps the paths of the files it reads inside itself, where nothing outside can see them. Reports are never renamed in a sweep, and neither is a workbook that shares a report\'s folder or opens with a report\'s words. The refusal names the report.\n\n**Moves one thing, and only on evidence.** A move changes a path exactly as a rename changes a name, so it breaks a refresh the same way. It will never move a report, a workbook a report reads, or anything inside a numbered stage folder — that last would assert how far work has got, which no file records. The single case it offers is a document sitting above the filing with no folder of its own, and only where exactly one folder beside it says in its own name what it holds. Two candidates is a choice, and choosing belongs to a person.\n\n**Never renames a folder.** Anything without a file extension is refused, so a folder cannot be renamed even if the listing filter is ever changed.\n\n**Never invents a status.** Where the name does not carry one, `assumeStatus` supplies the weakest — Draft — and the proposal says it was assumed. Set it to `""` to refuse instead.\n\n**Says where the date came from.** Using the file\'s timestamp instead of a date in the name is an assumption, and it is reported rather than left to be discovered. An eight-digit run like `04042025` is never read, because day-first and month-first disagree and being right by luck is not being right.')},
  [shouldApply, rename, confirm],
  { color: 5, width: 700, height: 560 },
);

export default workflow('upperform-wf-f-organizer', ${S(NAME)})
  .add(run)
  .to(config)
  .to(listFiles)
  .to(plan)
  .to(shouldApply
    .onTrue(rename.to(confirm).to(both.input(0)))
    .onFalse(both.input(1)))
  .add(both)
  .to(report)
  .add(noteHeader)
  .add(noteSafety);
`;
}
