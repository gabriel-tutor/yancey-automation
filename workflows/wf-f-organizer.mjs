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

  // Off by default. A rename breaks every saved link, Teams tab and Power BI
  // data source that points at the old name, so it happens after a person has
  // read the plan, not before.
  applyRenames: false,

  // Even switched on, one run touches at most this many files. A first live run
  // should be small enough to check by eye and undo by hand.
  maxRenamesPerRun: 5,

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
for (const it of items) {
  if (!it) continue;
  if (String(it.FSObjType) === '1') continue;
  const name = String(it.FileLeafRef || '');
  if (!name) continue;
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
for (const pl of planPlacements(files)) {
  const entry = extra[pl.path] || (extra[pl.path] = {});
  entry.folder = pl.folder;
  entry.placement = pl.verdict;
  entry.placementWhy = pl.reason;
}
for (const r of rows) {
  const found = extra[r.path] || {};
  r.category = found.category || '';
  r.categoryWhy = found.categoryWhy || '';
  r.folder = found.folder || '';
  r.placement = found.placement || '';
  r.placementWhy = found.placementWhy || '';
}

const compliant = rows.filter(function (r) { return r.compliant; });
const proposals = rows.filter(function (r) { return r.proposed; });
const blocked = rows.filter(function (r) { return !r.compliant && !r.proposed; });
const unfiled = rows.filter(function (r) { return r.placement === 'unfiled'; });

const summary = {
  library: String(cfg.libraryTitle || ''),
  site: String(cfg.sitePath || ''),
  scanned: rows.length,
  compliant: compliant.length,
  proposed: proposals.length,
  blocked: blocked.length,
  unfiled: unfiled.length,
  applyRenames: cfg.applyRenames === true,
  cap: Number(cfg.maxRenamesPerRun) || 5,
  rows: rows
};

// Nothing to do still emits one item, so the report always runs and the caller
// always gets an answer rather than a silent branch.
if (proposals.length === 0 || cfg.applyRenames !== true) {
  return [{ json: Object.assign({ nothingToRename: true }, summary) }];
}

// SharePoint takes the path inside an OData string literal. Single quotes are
// escaped by doubling; spaces have to be percent-encoded or the request line
// breaks. The URL is built here rather than in an expression so the escaping is
// in one readable place.
const lit = function (p) { return String(p).replace(/'/g, "''").replace(/ /g, '%20'); };
const base = String(cfg.host || '') + String(cfg.sitePath || '') + '/_api/web';

const take = proposals.slice(0, Number(cfg.maxRenamesPerRun) || 5);

return take.map(function (r) {
  const folder = r.path.slice(0, r.path.lastIndexOf('/'));
  const newPath = folder + '/' + r.proposed;
  return { json: {
    nothingToRename: false,
    oldName: r.current,
    newName: r.proposed,
    oldPath: r.path,
    newPath: newPath,
    assumptions: r.assumptions,
    // flags=0 means never overwrite. If something already sits on the new name
    // the call fails rather than destroying it.
    renameUrl: base + "/GetFileByServerRelativePath(decodedurl='" + lit(r.path) + "')/moveto(newurl='" + lit(newPath) + "',flags=0)",
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
    failed.push({ oldName: p.oldName, newName: p.newName, why: (m ? m[1] : blob).slice(0, 200) });
  } else {
    done.push({ oldName: p.oldName, newName: p.newName, assumptions: p.assumptions });
  }
}

const summary = (planned[0] && planned[0].summary) || {};
return [{ json: Object.assign({}, summary, { renamed: done, failedRenames: failed }) }];
`;

const REPORT = `
const r = $input.first().json || {};

// No colons: the n8n MCP returns an execution's error text and keeps only what
// follows the last one, so this is the only channel that carries a result back.
const flat = function (v) { return String(v === undefined || v === null ? '' : v).replace(/\\s+/g, ' ').replace(/[:]/g, '-'); };

const parts = [];
parts.push('library= ' + flat(r.library));
parts.push('scanned=' + r.scanned);
parts.push('compliant=' + r.compliant);
parts.push('proposed=' + r.proposed);
parts.push('blocked=' + r.blocked);
parts.push('unfiled=' + (r.unfiled === undefined ? '?' : r.unfiled));

if (r.renamed) {
  parts.push('RENAMED=' + r.renamed.length);
  parts.push('failed=' + (r.failedRenames || []).length);
  const lines = r.renamed.map(function (d) { return flat(d.oldName) + ' -> ' + flat(d.newName); });
  for (const f of (r.failedRenames || [])) {
    lines.push(flat(f.oldName) + ' FAILED ' + flat(f.why));
  }
  parts.push('>> ' + lines.join(' | ').slice(0, 700));
} else {
  parts.push('applyRenames=' + (r.applyRenames === true));
  const lines = [];
  for (const row of (r.rows || [])) {
    if (row.compliant) continue;
    const note = (row.assumptions || []).length ? ' [assumed ' + (row.assumptions || []).length + ']' : '';
    lines.push(flat(row.current) + ' -> ' + flat(row.proposed || ('BLOCKED ' + row.blocked)) + note);
  }
  parts.push('>> ' + lines.slice(0, 10).join(' | ').slice(0, 700));
}

throw new Error(parts.join(' '));
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
  output: [{ nothingToRename: false, oldName: 'Working TimesKV.xlsx', newName: 'Working-TimesKV_2025-05-01_Draft_v0.1.xlsx' }]
});

const shouldApply = ifElse({
  version: 2.2,
  config: {
    name: 'Apply Them?',
    parameters: {
      conditions: {
        options: { caseSensitive: true, leftValue: '', typeValidation: 'loose' },
        conditions: [
          { id: 'has-renames', leftValue: ${E('{{ $json.nothingToRename }}')}, operator: { type: 'boolean', operation: 'false', singleValue: true } }
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
      url: ${E('{{ $json.renameUrl }}')},
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
  output: [{ renamed: [], failedRenames: [] }]
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
  ${S('## WF-F · Organise Documents\n\nReads a SharePoint library, works out what each file should be called under Yancey\'s convention, and renames them when told to.\n\n```\n[Project]_[Document]_[YYYY-MM-DD]_[Status]_v#.#\nSCE_PM_Ride-Along-Plan_2026-09-01_Draft_v0.1.docx\n```\n\nEvery decision about what a file should be *called* lives in `src/filenames.js`, tested against his four worked examples and against the real contents of this library. These nodes only carry SharePoint\'s fields into it and its answers back out.')},
  [run, config, listFiles, plan],
  { color: 6, width: 820, height: 260 },
);

const noteSafety = sticky(
  ${S('### What stops this going wrong\n\n**Off by default.** `applyRenames` is false. A rename breaks every saved link, Teams tab and Power BI data source pointing at the old name, so a person reads the plan first.\n\n**Capped.** `maxRenamesPerRun` limits one run to a handful, small enough to check by eye and undo by hand.\n\n**Never overwrites.** The move uses `flags=0`, so if anything already sits on the new name the call fails rather than destroying it.\n\n**Never collides.** Two files that would land on the same name are both refused, not silently merged.\n\n**Never invents a status.** A status says whether a document is approved. Where the name does not carry one, `assumeStatus` supplies the weakest — Draft — and the proposal says it was assumed. Set it to `""` to refuse instead.\n\n**Says where the date came from.** Using the file\'s timestamp instead of a date in the name is an assumption, and it is reported rather than left to be discovered. An eight-digit run like `04042025` is never read, because day-first and month-first disagree and being right by luck is not being right.')},
  [shouldApply, rename, confirm],
  { color: 5, width: 700, height: 380 },
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
