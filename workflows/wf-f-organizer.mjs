import { codeNode } from '../scripts/bundle.mjs';
import { HELPERS } from './helpers-manifest.mjs';
import { E, S, SDK_IMPORT } from './shared.mjs';

/**
 * WF-F reads a SharePoint document library and says which files break Yancey's
 * naming convention and what each should be called instead.
 *
 * It renames nothing. The credential holds AllSites.Read, so it could not write
 * even if it wanted to, and that is the right order: a rename is destructive and
 * irreversible from the recipient's side — a link someone saved stops working —
 * so the list gets reviewed before anything moves. `applyRenames` is the switch,
 * and turning it on needs AllSites.Write on the app registration as well.
 */
const CONFIG = {
  host: 'https://upperformpartners.sharepoint.com',
  sitePath: '/sites/TaskManagement',
  // The default library on a Teams-connected site. Its display title is
  // "Documents" while its URL says "Shared Documents".
  libraryTitle: 'Documents',
  // How many items to look at in one pass.
  maxItems: 200,
  // Renaming is a separate permission and a separate decision. See above.
  applyRenames: false,
};

const AUDIT = `
const cfg = $('Config').first().json;
const listed = $input.first().json || {};
const items = Array.isArray(listed.value) ? listed.value : [];

// FileLeafRef is the file name, FileRef the server-relative path. List items
// cover every folder in the library, so this sees nested files too.
const files = [];
for (const it of items) {
  const name = String((it && it.FileLeafRef) || '');
  const path = String((it && it.FileRef) || '');
  if (!name) continue;
  // Folders come back as items too and have no extension to judge.
  if (String(it.FSObjType) === '1' || it.FSObjType === 1) continue;
  files.push({ name: name, path: path, modified: String(it.Modified || '') });
}

const rows = [];
for (const f of files) {
  const problems = diagnose(f.name);
  if (problems.length === 0) {
    rows.push({ name: f.name, path: f.path, ok: true, problems: [], proposed: '' });
    continue;
  }

  // Only propose a new name where the pieces can be read off the old one.
  // Inventing a project or a status for someone else's document is worse than
  // saying plainly that a human has to decide.
  const dot = f.name.lastIndexOf('.');
  const extension = dot > 0 ? f.name.slice(dot + 1) : '';
  const stem = dot > 0 ? f.name.slice(0, dot) : f.name;

  // Only a year-first date is read from the name. An eight-digit run like
  // 04042025 is ambiguous between day-first and month-first, and guessing wrong
  // silently mislabels the document, so it is left to the file's own timestamp.
  const dateMatch = stem.match(/(20\\d{2})[-_.]?(\\d{2})[-_.]?(\\d{2})/);
  const modified = f.modified ? f.modified.slice(0, 10) : '';
  const date = dateMatch ? dateMatch[1] + '-' + dateMatch[2] + '-' + dateMatch[3] : modified;
  const dateFrom = dateMatch ? 'name' : (modified ? 'file timestamp' : '');

  const statusMatch = stem.match(/(^|[_\\-. ])(Draft|Review|Final|Archived)([_\\-. ]|$)/i);
  const statusWord = statusMatch ? statusMatch[2] : '';
  const status = statusWord
    ? statusWord.charAt(0).toUpperCase() + statusWord.slice(1).toLowerCase()
    : '';

  const versionMatch = stem.match(/(^|[_\\-. ])v(\\d+)(?:[._](\\d+))?([_\\-. ]|$)/i);
  const version = versionMatch
    ? { major: Number(versionMatch[2]), minor: Number(versionMatch[3] || 0) }
    : null;

  let proposed = '';
  let blocked = '';
  if (!status) {
    blocked = 'no status in the name — needs Draft, Review, Final or Archived';
  } else if (!date) {
    blocked = 'no date in the name and none on the file';
  } else if (!version) {
    blocked = 'no version in the name';
  } else {
    // Strip the parts that are being rebuilt, and keep what is left as the name.
    let remainder = stem;
    if (dateMatch) remainder = remainder.replace(dateMatch[0], ' ');
    if (statusMatch) remainder = remainder.replace(statusMatch[0], ' ');
    if (versionMatch) remainder = remainder.replace(versionMatch[0], ' ');

    try {
      proposed = proposeFileName({
        project: '',
        document: remainder,
        date: date,
        status: status,
        version: version,
        extension: extension
      });
    } catch (e) {
      blocked = String(e.message || e).slice(0, 140);
    }
  }

  rows.push({
    name: f.name,
    path: f.path,
    ok: false,
    problems: problems,
    proposed: proposed,
    blocked: blocked,
    dateFrom: proposed ? dateFrom : ''
  });
}

const good = rows.filter(function (r) { return r.ok; });
const bad = rows.filter(function (r) { return !r.ok; });
const renameable = bad.filter(function (r) { return r.proposed; });

return [{ json: {
  library: String(cfg.libraryTitle || ''),
  site: String(cfg.sitePath || ''),
  scanned: rows.length,
  compliant: good.length,
  breaking: bad.length,
  renameable: renameable.length,
  needsAHuman: bad.length - renameable.length,
  applyRenames: cfg.applyRenames === true,
  rows: rows
}}];
`;

const REPORT = `
const r = $input.first().json || {};

// No colons: the n8n MCP returns an execution's error text and keeps only what
// follows the last one, so this is the only channel that carries a result back.
const flat = function (v) { return String(v === undefined || v === null ? '' : v).replace(/\\s+/g, ' ').replace(/[:]/g, '-'); };

const lines = [];
for (const row of (r.rows || []).slice(0, 12)) {
  if (row.ok) continue;
  const note = row.dateFrom === 'file timestamp' ? ' (date from file timestamp)' : '';
  lines.push(flat(row.name) + ' -> ' + flat(row.proposed || ('BLOCKED ' + row.blocked)) + note);
}

throw new Error(
  'library= ' + flat(r.library)
  + ' scanned=' + r.scanned
  + ' compliant=' + r.compliant
  + ' breaking=' + r.breaking
  + ' renameable=' + r.renameable
  + ' needsAHuman=' + r.needsAHuman
  + ' >> ' + lines.join(' | ').slice(0, 800)
);
`;

export const NAME = 'WF-F · Organise Documents';

export async function build() {
  const auditCode = await codeNode(HELPERS['wf-f/Audit Names'], AUDIT);

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
  output: [{ value: [{ FileLeafRef: 'Task list FINAL (2).xlsx', FileRef: '/sites/TaskManagement/Shared Documents/Task list FINAL (2).xlsx', Modified: '2026-08-29T10:00:00Z', FSObjType: 0 }] }]
});

const audit = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Audit Names',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(auditCode)} },
    position: [620, 0]
  },
  output: [{ scanned: 12, compliant: 1, breaking: 11, renameable: 4, needsAHuman: 7 }]
});

const report = node({
  type: 'n8n-nodes-base.code',
  version: 2,
  config: {
    name: 'Report',
    parameters: { mode: 'runOnceForAllItems', language: 'javaScript', jsCode: ${S(REPORT)} },
    position: [840, 0]
  },
  output: [{}]
});

const noteHeader = sticky(
  ${S('## WF-F · Organise Documents\n\nReads a SharePoint library and says which files break Yancey\'s naming convention, and what each one should be called instead.\n\n```\n[Project]_[Document]_[YYYY-MM-DD]_[Status]_v#.#\nSCE_PM_Ride-Along-Plan_2026-09-01_Draft_v0.1.docx\n```\n\nThe judgement lives in `src/filenames.js`, where his four worked examples are the tests. This node only carries SharePoint\'s fields into it.')},
  [run, config, listFiles, audit],
  { color: 6, width: 820, height: 260 },
);

const noteRename = sticky(
  ${S('### It renames nothing yet\n\nTwo reasons, and the second is the real one.\n\nThe credential holds `AllSites.Read`, so it could not write if it tried. Turning renaming on needs `AllSites.Write` added to the app registration.\n\nMore to the point, a rename is destructive from the far side: a link someone saved, a Teams tab, a shortcut in a chat all stop resolving. So the proposal gets read by a person first. `applyRenames` in Config is the switch.\n\n**A file is only proposed a new name when every part can be read off the old one** — a date, a status word, a version. Where any of those is missing the row says so and waits. Inventing a status for somebody else\'s document is worse than admitting a human has to look.')},
  [report],
  { color: 5, width: 620, height: 300 },
);

export default workflow('upperform-wf-f-organizer', ${S(NAME)})
  .add(run)
  .to(config)
  .to(listFiles)
  .to(audit)
  .to(report)
  .add(noteHeader)
  .add(noteRename);
`;
}
