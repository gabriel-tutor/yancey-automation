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
 * The report that is likely reading this file, if any.
 *
 * A .pbix keeps the paths of its data sources inside itself, where nothing
 * outside can see them. These are the two signals that survive to the outside:
 * a report and its source usually share an opening, and they usually share a
 * folder. Neither is proof, and the caller is told which one fired so it can
 * say so rather than assert more than it knows.
 *
 * @param {string} name
 * @param {string} path
 * @param {Array<{ name: string, path?: string }>} reports
 * @returns {{ report: string, because: 'name' | 'folder' } | null}
 */
export function reportReading(name, path, reports) {
  if (isReport(name) || !couldFeedAReport(name)) return null;
  const list = Array.isArray(reports) ? reports : [];
  const mine = wordsOf(name);

  const named = list.find((r) => sharedOpening(mine, wordsOf(r.name)) >= SHARED_WORDS_MEANING_PAIRED);
  if (named) return { report: named.name, because: 'name' };

  const neighbour = list.find((r) => folderOf(r.path || '') === folderOf(path));
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
 * @returns {{ category: string, evidence: string }}
 */
export function categorise(file, reports) {
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

  const reading = reportReading(name, path, reports);
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
 * Whether a document is where it should be, and why.
 *
 * This deliberately proposes no moves. Every move this library could justify
 * would either break a Power BI refresh, which it cannot see, or assert a
 * workflow stage, which no file records. So it reports, and a person decides.
 * The one thing it will say is when a document has no folder at all.
 *
 * @param {Array<{ name: string, path: string }>} files
 * @param {Array<{ name: string, path?: string }>} [knownReports]
 * @returns {Array<{ name: string, path: string, folder: string, verdict: 'settled' | 'unfiled', reason: string }>}
 */
export function planPlacements(files, knownReports) {
  const list = Array.isArray(files) ? files : [];
  const reports = knownReports || reportsIn(list);
  const folders = list.map((f) => folderOf(f.path));

  return list.map((file) => {
    const name = String(file.name || '');
    const path = String(file.path || '');
    const folder = folderName(path);
    const base = { name, path, folder };

    if (isReport(name)) {
      return {
        ...base,
        verdict: /** @type {const} */ ('settled'),
        reason:
          'a Power BI report is referenced by path wherever it is published or set to refresh, so moving it breaks the same things renaming it would',
      };
    }

    const reading = reportReading(name, path, reports);
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
      return {
        ...base,
        verdict: /** @type {const} */ ('unfiled'),
        reason:
          'it sits above the filing rather than in it, with folders of documents beneath it, and nothing in the library says which one it belongs to',
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
 * @returns {Array<{ name: string, path: string, category: string, evidence: string }>}
 */
export function categoriseAll(files) {
  const list = Array.isArray(files) ? files : [];
  const reports = reportsIn(list);
  return list.map((file) => {
    const found = categorise(file, reports);
    return { name: String(file.name || ''), path: String(file.path || ''), ...found };
  });
}
