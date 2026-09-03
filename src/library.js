/**
 * Upperform document organiser — what a document is, and where it sits.
 *
 * `filenames.js` answers one question: what should this file be called? This
 * module answers the two that only make sense across a whole library — what
 * kind of document is this, and is it in the right place — because both
 * depend on the other files around it.
 *
 * Nothing here opens a document. Every judgement is drawn from the file name,
 * its extension, its folder, and its relationship to the other files listed
 * beside it. That is a deliberate limit: reading the contents of a client's
 * document library would mean sending those contents somewhere, and that is a
 * decision for Yancey to make knowingly rather than one to slip into.
 */

/**
 * Split a file name into stem and extension.
 *
 * Lives here rather than in filenames.js because both modules need it and the
 * rules are fiddly: a trailing dot is not an extension, and neither is the dot
 * inside a version like v0.1.
 *
 * @param {string} value
 * @returns {{ stem: string, extension: string }}
 */
export function splitExtension(value) {
  const name = String(value || '');
  const dot = name.lastIndexOf('.');
  if (dot <= 0 || dot === name.length - 1) return { stem: name, extension: '' };
  const extension = name.slice(dot + 1);
  // A real extension names a file type, and every file type has a letter in
  // it. Without this, "Plan_2026-09-01_Draft_v0.1" reads as extension "1" —
  // which would let a folder named that way past the guard that exists to stop
  // a folder ever being renamed.
  if (!/^(?=[A-Za-z0-9]*[A-Za-z])[A-Za-z0-9]+$/.test(extension)) return { stem: name, extension: '' };
  return { stem: name.slice(0, dot), extension };
}

/**
 * The folder a file sits in, as a server-relative path.
 *
 * @param {string} path
 * @returns {string}
 */
export function folderOf(path) {
  const cut = String(path || '').lastIndexOf('/');
  return cut === -1 ? '' : String(path).slice(0, cut);
}

/**
 * The words in a file name, lowercased, with punctuation treated as a break.
 *
 * @param {string} name
 * @returns {string[]}
 */
export function wordsOf(name) {
  return splitExtension(String(name || ''))
    .stem.toLowerCase()
    .split(/[^a-z0-9]+/)
    .filter(Boolean);
}

/**
 * How many words two names agree on before they diverge.
 *
 * @param {string[]} a
 * @param {string[]} b
 * @returns {number}
 */
export function sharedOpening(a, b) {
  let n = 0;
  while (n < a.length && n < b.length && a[n] === b[n]) n += 1;
  return n;
}

const POWER_BI_REPORT = /\.pbix$/i;
const COULD_FEED_A_REPORT = /\.(xlsx|xlsm|xlsb|xls|csv|tsv|txt|json|xml|parquet)$/i;

// "DIMP E-1.xlsx" against "DIMP E-1 Performance Measures.pbix" is three words
// of agreement, which is a pairing. "Metrics" against "NiSource Metric
// Dashboard" is none, which is a coincidence of vocabulary.
const SHARED_WORDS_MEANING_PAIRED = 3;

/** @param {string} name @returns {boolean} */
export function isReport(name) {
  return POWER_BI_REPORT.test(String(name || ''));
}

/** @param {string} name @returns {boolean} */
export function couldFeedAReport(name) {
  return COULD_FEED_A_REPORT.test(String(name || ''));
}

/**
 * The Power BI reports in a listing.
 *
 * @param {Array<{ name: string, path?: string }>} files
 * @returns {Array<{ name: string, path?: string }>}
 */
export function reportsIn(files) {
  return (Array.isArray(files) ? files : []).filter((f) => isReport(f && f.name));
}

/**
 * The report reading this file, if any.
 *
 * A .pbix keeps the paths of its data sources inside itself. That made this a
 * guess for a long time, resting on two signals that survive to the outside: a
 * report and its source usually share an opening, and they usually share a
 * folder. Neither is proof, and both refuse far more than they need to — nine
 * documents in this library were held back by them.
 *
 * They are now the fallback rather than the method. When the reports have been
 * opened and their queries read (see src/powerbi.js), `options.dependencies`
 * carries the answer and the guessing stops:
 *
 *   named in a query        refused, and the query line is the proof
 *   not named, all read     safe — the reports have said what they read
 *   some report unreadable  fall back to the guess, because we still cannot tell
 *
 * The middle case is the point. Certainty that a file is *not* a dependency is
 * worth as much as certainty that it is, and only reading the report gives it.
 *
 * @param {string} name
 * @param {string} path
 * @param {Array<{ name: string, path?: string }>} reports
 * @param {{ dependencies?: { byFile: Record<string, { report: string, where: string }>, unreadable?: string[] } }} [options]
 * @returns {{ report: string, because: 'query' | 'name' | 'folder', where?: string } | null}
 */
export function reportReading(name, path, reports, options) {
  if (isReport(name) || !couldFeedAReport(name)) return null;
  const list = Array.isArray(reports) ? reports : [];

  const known = options && options.dependencies;
  if (known && known.byFile) {
    const hit = known.byFile[String(name).toLowerCase()];
    if (hit) return { report: hit.report, because: 'query', where: hit.where };
    // Every report opened and none of them named this file. That is an answer,
    // not an absence of one, so the guesses below are not consulted.
    if ((known.unreadable || []).length === 0) return null;
  }

  const mine = wordsOf(name);
  const named = list.find((r) => sharedOpening(mine, wordsOf(r.name)) >= SHARED_WORDS_MEANING_PAIRED);
  if (named) return { report: named.name, because: 'name' };

  // Two files with no path do not share a folder, they share the absence of
  // one. Pairing on that would refuse renames on no evidence at all.
  const folder = folderOf(path);
  const neighbour = folder ? list.find((r) => folderOf(r.path || '') === folder) : undefined;
  if (neighbour) return { report: neighbour.name, because: 'folder' };

  return null;
}

/** @type {Record<string, string>} */
const KIND_BY_EXTENSION = {
  pbix: 'Power BI report',
  xlsx: 'spreadsheet',
  xlsm: 'spreadsheet',
  xlsb: 'spreadsheet',
  xls: 'spreadsheet',
  csv: 'spreadsheet',
  tsv: 'spreadsheet',
  docx: 'document',
  doc: 'document',
  pdf: 'document',
  rtf: 'document',
  txt: 'document',
  md: 'document',
  pptx: 'presentation',
  ppt: 'presentation',
  png: 'image',
  jpg: 'image',
  jpeg: 'image',
  gif: 'image',
  bmp: 'image',
  webp: 'image',
  heic: 'image',
  svg: 'image',
  zip: 'archive',
  '7z': 'archive',
  rar: 'archive',
  msg: 'email',
  eml: 'email',
};

const SCREEN_CAPTURE = /\b(?:screen[ _-]?shot|screen[ _-]?capture|capture d[eu])\b/i;

/**
 * What kind of document this is, and what says so.
 *
 * The evidence matters as much as the answer. "spreadsheet" because of an
 * extension is a fact; "report data" because a named report probably reads it
 * is an inference, and a person reading the review should be able to tell the
 * two apart without asking.
 *
 * @param {{ name: string, path?: string }} file
 * @param {Array<{ name: string, path?: string }>} reports
 * @param {{ dependencies?: { byFile: Record<string, { report: string, where: string }>, unreadable?: string[] } }} [options]
 * @returns {{ category: string, evidence: string }}
 */
export function categorise(file, reports, options) {
  const spec = file || { name: '' };
  const name = String(spec.name || '');
  const path = String(spec.path || '');
  const extension = splitExtension(name).extension.toLowerCase();
  const kind = KIND_BY_EXTENSION[extension] || 'other';

  if (kind === 'Power BI report') {
    return { category: 'Power BI report', evidence: 'a .pbix is a Power BI report' };
  }

  if (kind === 'image' && SCREEN_CAPTURE.test(name)) {
    return { category: 'screenshot', evidence: 'the name says it is a screen capture' };
  }

  const reading = reportReading(name, path, reports, options);
  if (reading) {
    return {
      category: 'report data',
      evidence:
        'the report ' +
        reading.report +
        (reading.because === 'name' ? ' opens with the same words' : ' sits in the same folder'),
    };
  }

  if (kind === 'other') {
    return {
      category: 'other',
      evidence: extension ? 'nothing here recognises .' + extension : 'no file extension',
    };
  }

  return { category: kind, evidence: 'from the .' + extension + ' extension' };
}

// A folder named "1. Assigned Tasks" is one step of a sequence. Where several
// siblings are numbered like this, the folders record how far work has got —
// which is a fact about the work, not about any file inside them.
const STAGE_FOLDER = /^\d+[.)]?\s+\S/;

/**
 * The numbered stage folder a file sits under, at any depth, or ''.
 *
 * @param {string} path
 * @returns {string}
 */
export function stageFolderOf(path) {
  const segments = folderOf(path).split('/');
  const hit = segments.find((segment) => STAGE_FOLDER.test(segment));
  return hit || '';
}

/**
 * The last segment of a folder path, for saying where something is without
 * repeating the whole server path back at the reader.
 *
 * @param {string} path
 * @returns {string}
 */
function folderName(path) {
  const folder = folderOf(path);
  const cut = folder.lastIndexOf('/');
  return cut === -1 ? folder : folder.slice(cut + 1);
}

/**
 * Whether other files in the listing live in folders beneath this one.
 *
 * A file whose folder has filed sub-folders under it is sitting at a branch
 * rather than at a leaf — it is above the filing, not in it.
 *
 * @param {string} folder
 * @param {string[]} folders
 * @returns {boolean}
 */
function hasFilingBeneath(folder, folders) {
  return folders.some((other) => other !== folder && other.indexOf(folder + '/') === 0);
}

/**
 * A folder whose name says it holds general material rather than the output of
 * one task. Used only to home a document that has no folder at all, and only
 * when the library offers exactly one such folder beside it — the destination
 * is read out of the library, never invented.
 */
const GENERAL_FOLDER = /\b(resources?|references?|general|shared|admin|misc|archive)\b/i;

/** @param {string} path @returns {string} */
function lastSegment(path) {
  const cut = String(path || '').lastIndexOf('/');
  return cut === -1 ? String(path || '') : String(path).slice(cut + 1);
}

/**
 * Whether a document is where it should be, and why.
 *
 * It proposes a move in one case only: a document sitting above the filing with
 * no folder of its own, where exactly one folder beside it says in its own name
 * that it holds general material. Everything else is left alone, because every
 * other move this library could justify would either break a Power BI refresh
 * it cannot see, or assert a workflow stage that no file records.
 *
 * @param {Array<{ name: string, path: string }>} files
 * @param {{ reports?: Array<{ name: string, path?: string }>, folders?: string[], dependencies?: { byFile: Record<string, { report: string, where: string }>, unreadable?: string[] } }} [options]
 * @returns {Array<{ name: string, path: string, folder: string, verdict: 'settled' | 'unfiled' | 'move', reason: string, destination: string }>}
 */
export function planPlacements(files, options) {
  const list = Array.isArray(files) ? files : [];
  const opts = options || {};
  const reports = opts.reports || reportsIn(list);
  const knownFolders = Array.isArray(opts.folders) ? opts.folders : [];
  const folders = list.map((f) => folderOf(f.path));

  return list.map((file) => {
    const name = String(file.name || '');
    const path = String(file.path || '');
    const folder = folderName(path);
    const base = { name, path, folder, destination: '' };

    if (isReport(name)) {
      return {
        ...base,
        verdict: /** @type {const} */ ('settled'),
        reason:
          'a Power BI report is referenced by path wherever it is published or set to refresh, so moving it breaks the same things renaming it would',
      };
    }

    const reading = reportReading(name, path, reports, opts);
    if (reading) {
      return {
        ...base,
        verdict: /** @type {const} */ ('settled'),
        reason:
          'the report ' +
          reading.report +
          ' very likely reads it, and a move changes the path exactly as a rename changes the name',
      };
    }

    const stage = stageFolderOf(path);
    if (stage) {
      return {
        ...base,
        verdict: /** @type {const} */ ('settled'),
        reason:
          'it sits under ' +
          stage +
          ', and those numbered folders record how far the work has got — nothing about a file says which stage it belongs in',
      };
    }

    if (hasFilingBeneath(folderOf(path), folders)) {
      // Only a folder sitting directly beside the file counts, and only if its
      // own name says what it holds. Two candidates is a choice, and choosing
      // is the part that belongs to a person.
      const here = folderOf(path);
      const candidates = knownFolders.filter(
        (f) => folderOf(f) === here && GENERAL_FOLDER.test(lastSegment(f)),
      );

      if (candidates.length === 1) {
        const into = String(candidates[0]);
        return {
          ...base,
          verdict: /** @type {const} */ ('move'),
          destination: into + '/' + name,
          reason:
            'it sits above the filing with no folder of its own, and ' +
            lastSegment(into) +
            ' is the only folder beside it that says in its own name what it holds',
        };
      }

      return {
        ...base,
        verdict: /** @type {const} */ ('unfiled'),
        reason:
          'it sits above the filing rather than in it, with folders of documents beneath it, and ' +
          (candidates.length === 0
            ? 'no folder beside it says what it holds'
            : 'more than one folder beside it could hold it, which is a choice for a person'),
      };
    }

    return {
      ...base,
      verdict: /** @type {const} */ ('settled'),
      reason: 'filed under ' + (folder || 'the library'),
    };
  });
}

/**
 * Categorise a whole listing in one pass, so the reports only get found once.
 *
 * @param {Array<{ name: string, path: string }>} files
 * @param {{ dependencies?: { byFile: Record<string, { report: string, where: string }>, unreadable?: string[] } }} [options]
 * @returns {Array<{ name: string, path: string, category: string, evidence: string }>}
 */
export function categoriseAll(files, options) {
  const list = Array.isArray(files) ? files : [];
  const reports = reportsIn(list);
  return list.map((file) => {
    const found = categorise(file, reports, options);
    return { name: String(file.name || ''), path: String(file.path || ''), ...found };
  });
}

/**
 * The folder every one of these paths sits inside.
 *
 * A library root is not a constant, and writing it as one made the review page
 * a page about *this* library rather than a page about a library: it stripped a
 * literal site path off everything, and pointed anywhere else it would have
 * rendered every folder as its full server path.
 *
 * The prefix is taken on path separators, so two folders called `Task
 * Management` and `Task Force` share `/a` rather than `/a/Task `.
 *
 * @param {string[]} paths
 * @returns {string}
 */
export function libraryRootOf(paths) {
  const real = (Array.isArray(paths) ? paths : []).map(String).filter(Boolean);
  if (real.length === 0) return '';

  let shared = String(real[0]).split('/');
  for (const path of real.slice(1)) {
    const parts = path.split('/');
    let n = 0;
    while (n < shared.length && n < parts.length && shared[n] === parts[n]) n += 1;
    shared = shared.slice(0, n);
  }

  return shared.join('/');
}

/**
 * A path or name reduced to what SharePoint actually compares.
 *
 * SharePoint matches paths and file names without regard to case, so
 * `Dashboards/Report.xlsx` and `dashboards/report.xlsx` are one document. Every
 * comparison of one path against another has to go through here, or the next
 * one written case-sensitively becomes a folder created twice, a collision
 * missed, or an approved line that silently matches nothing.
 *
 * @param {string} path
 * @returns {string}
 */
export function pathKey(path) {
  return String(path || '').toLowerCase();
}
