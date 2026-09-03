/**
 * What a Power BI report actually reads.
 *
 * The organiser refused to rename nine documents, and every refusal rested on
 * the same sentence: a `.pbix` keeps the paths of its data sources inside
 * itself, where nothing outside can see them. That was true of the *library* —
 * you cannot tell from a folder listing — but it was never true of the file.
 * A `.pbix` is a ZIP, and this repository already reads ZIPs.
 *
 * So the guard stops guessing. Instead of "this workbook opens with the same
 * words as that report, so the report probably reads it", the report is opened
 * and asked. A workbook the report does not name is safe to rename; a workbook
 * it does name is refused with the query line that proves it.
 *
 * Where the paths live:
 *
 *   DataMashup        the Power Query package — a length-prefixed header
 *                     wrapping *another* ZIP, whose Formulas/Section1.m holds
 *                     the M queries. This is where File.Contents(...) and
 *                     SharePoint.Files(...) appear, and it is the one that
 *                     matters.
 *   Connections       JSON connection strings, for reports that bind directly.
 *   DataModelSchema   the model, when present, often UTF-16LE.
 *
 * `Report/Layout` is deliberately *not* read. It is the visual layout, and the
 * only file names in it are Power BI's own theme resources —
 * `BaseThemes/CY24SU10.json`, `BuiltInThemes/AccessibleTidal.json`. Reading it
 * turned every report in this library into one that "depends on" a theme, which
 * is both false and the kind of false that looks like a real answer.
 *
 * That distinction is the whole point of `readable`. A report where none of the
 * three parts above exists has told us nothing, and must not be recorded as a
 * report that reads nothing — those are opposite conclusions, and only one of
 * them makes a rename safe.
 *
 * Nothing here writes: a report is read to find out what it depends on, and is
 * never itself modified.
 */
import { readZip } from './unzip.js';

/**
 * The names of every part in a package, without decoding any of them.
 *
 * `readZip` decides what to inflate by asking `wanted`, so a predicate that
 * records the name and then declines it walks the central directory for free.
 * This exists because "unreadable" is not a useful answer on its own — a report
 * that could not be read should be able to say what it does contain, so the
 * next question is answerable without another round trip.
 *
 * @param {Uint8Array} bytes
 * @returns {string[]}
 */
export function entryNames(bytes) {
  /** @type {string[]} */
  const names = [];
  try {
    readZip(bytes, (name) => {
      names.push(name);
      return false;
    });
  } catch {
    return [];
  }
  return names;
}

/** Extensions a report could plausibly be reading from. */
const DATA_FILE = /\.(xlsx|xlsm|xlsb|xls|csv|tsv|txt|json|xml|parquet|accdb|mdb)$/i;

/** The DataMashup header is four bytes of version then four of parts length. */
const MASHUP_HEADER = 8;

/**
 * Text out of a part, guessing the encoding the way Power BI writes it.
 *
 * Several parts are UTF-16LE with no BOM, which decodes as UTF-8 into a string
 * full of NUL bytes — readable enough to fool a regex into matching nothing.
 * A NUL in the first bytes is the tell.
 *
 * @param {Uint8Array} bytes
 * @returns {string}
 */
function decodePart(bytes) {
  if (!bytes || bytes.length === 0) return '';
  const head = bytes.subarray(0, Math.min(64, bytes.length));
  let nuls = 0;
  for (const b of head) if (b === 0) nuls += 1;
  const label = nuls > head.length / 4 ? 'utf-16le' : 'utf-8';
  try {
    return new TextDecoder(label).decode(bytes);
  } catch {
    return '';
  }
}

/**
 * The M queries inside a DataMashup part.
 *
 * @param {Uint8Array} bytes the DataMashup part
 * @returns {string} the query text, or '' if it cannot be opened
 */
export function mashupQueries(bytes) {
  if (!bytes || bytes.length < MASHUP_HEADER) return '';
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const partsLength = view.getUint32(4, true);
  if (partsLength <= 0 || MASHUP_HEADER + partsLength > bytes.length) return '';
  const inner = bytes.subarray(MASHUP_HEADER, MASHUP_HEADER + partsLength);
  try {
    const zip = readZip(inner, (name) => /\.m$/i.test(name) || /Section\d*\.m$/i.test(name));
    return Object.keys(zip)
      .map((name) => decodePart(zip[name] ?? new Uint8Array()))
      .join('\n');
  } catch {
    return '';
  }
}

/**
 * Every file a report names, with the line that names it.
 *
 * The M language spells a source out in full — `File.Contents("C:\Work\x.xlsx")`
 * or `SharePoint.Files("https://…")` — so the file name is recoverable without
 * understanding the query. Anything ending in a data extension is collected,
 * whether it sits in a function call or a bare string, because a name is
 * evidence of a dependency wherever it appears.
 *
 * @param {Uint8Array} bytes the whole .pbix
 * @returns {{ readable: boolean, parts: string[], entries: string[], sources: Array<{ name: string, where: string }> }}
 */
export function dataSourcesIn(bytes) {
  /** @type {Record<string, Uint8Array>} */
  let zip;
  try {
    zip = readZip(
      bytes,
      (name) => name === 'DataMashup' || name === 'Connections' || name === 'DataModelSchema',
    );
  } catch {
    return { readable: false, parts: [], entries: [], sources: [] };
  }

  const parts = [];
  const read = [];
  const mashup = zip.DataMashup;
  if (mashup) {
    const queries = mashupQueries(mashup);
    if (queries) {
      parts.push(queries);
      read.push('DataMashup');
    }
  }
  for (const key of ['Connections', 'DataModelSchema']) {
    const part = zip[key];
    if (!part) continue;
    const text = decodePart(part);
    if (!text) continue;
    parts.push(text);
    read.push(key);
  }

  // No queries and no connections is not "reads nothing" — it is "did not find
  // out". Saying otherwise would clear a rename on an answer nobody gave.
  // Say what it does hold, so an unreadable report is a question with an
  // answer attached rather than a dead end.
  if (read.length === 0) return { readable: false, parts: [], entries: entryNames(bytes), sources: [] };
  const text = parts.join('\n');

  /** @type {Map<string, string>} */
  const found = new Map();

  // Quoted strings are where a source appears, in both M and JSON.
  for (const m of text.matchAll(/"((?:[^"\\]|\\.){2,400}?)"/g)) {
    const value = String(m[1] || '').replace(/\\\\/g, '\\');
    if (!DATA_FILE.test(value)) continue;
    const cut = Math.max(value.lastIndexOf('\\'), value.lastIndexOf('/'));
    const name = cut === -1 ? value : value.slice(cut + 1);
    if (!name || !DATA_FILE.test(name)) continue;
    if (!found.has(name.toLowerCase())) found.set(name.toLowerCase(), value);
  }

  return {
    readable: true,
    parts: read,
    entries: [],
    sources: [...found.entries()].map(([, where]) => {
      const cut = Math.max(where.lastIndexOf('\\'), where.lastIndexOf('/'));
      return { name: cut === -1 ? where : where.slice(cut + 1), where };
    }),
  };
}

/**
 * Does this package mention any of these names, anywhere inside it?
 *
 * Parsing M queries was the wrong primitive. Two of the four reports in this
 * library have a `DataModel` and no `DataMashup` at all — the queries live
 * inside the model blob, which is an Analysis Services backup rather than
 * anything a ZIP reader can walk. Extracting sources from it is hard; the
 * question that actually matters is not.
 *
 * A report that reads a workbook must contain that workbook's name somewhere:
 * in a query, a connection string, or the model metadata. So the test is
 * whether the name appears in the bytes at all. That is:
 *
 *   sound         a name that appears nowhere cannot be being referenced
 *   conservative  an incidental mention refuses a rename, which is the safe way
 *                 to be wrong
 *   format-proof  it does not care how the report stores its queries
 *
 * Both UTF-8 and UTF-16LE are searched, because Power BI writes both and a
 * UTF-16 name is invisible to a UTF-8 search.
 *
 * @param {Uint8Array} bytes the whole .pbix
 * @param {string[]} names file names to look for
 * @returns {{ readable: boolean, mentioned: string[] }}
 */
export function mentionsIn(bytes, names) {
  const wanted = (Array.isArray(names) ? names : []).map(String).filter(Boolean);
  /** @type {Record<string, Uint8Array>} */
  let zip;
  try {
    zip = readZip(bytes, () => true);
  } catch {
    return { readable: false, mentioned: [] };
  }
  const parts = Object.values(zip);
  if (parts.length === 0) return { readable: false, mentioned: [] };

  const needles = wanted.map((name) => ({ name, utf8: utf8Bytes(name), utf16: utf16Bytes(name) }));
  const mentioned = [];
  for (const needle of needles) {
    const hit = parts.some((part) => contains(part, needle.utf8) || contains(part, needle.utf16));
    if (hit) mentioned.push(needle.name);
  }
  return { readable: true, mentioned };
}

/** @param {string} text @returns {Uint8Array} */
function utf8Bytes(text) {
  return new TextEncoder().encode(text);
}

/** @param {string} text @returns {Uint8Array} */
function utf16Bytes(text) {
  const out = new Uint8Array(text.length * 2);
  for (let i = 0; i < text.length; i += 1) {
    const code = text.charCodeAt(i);
    out[i * 2] = code & 0xff;
    out[i * 2 + 1] = code >> 8;
  }
  return out;
}

/**
 * Naive byte search. The haystacks here are a few megabytes and the needles a
 * few dozen bytes, which is well inside what a plain scan handles in the time a
 * download takes.
 *
 * @param {Uint8Array} hay
 * @param {Uint8Array} needle
 * @returns {boolean}
 */
function contains(hay, needle) {
  if (needle.length === 0 || hay.length < needle.length) return false;
  const first = needle[0];
  const last = hay.length - needle.length;
  for (let i = 0; i <= last; i += 1) {
    if (hay[i] !== first) continue;
    let k = 1;
    while (k < needle.length && hay[i + k] === needle[k]) k += 1;
    if (k === needle.length) return true;
  }
  return false;
}

/**
 * Turn what each report reads into the lookup the organiser needs.
 *
 * A report that could not be opened is recorded as unreadable rather than as
 * reading nothing, because those are opposite conclusions: one means "safe to
 * rename", the other means "we still cannot tell".
 *
 * @param {Array<{ name: string, path?: string, bytes?: Uint8Array, readable?: boolean, sources?: Array<{ name: string, where: string }> }>} reports
 * @param {string[]} [candidates] names of the library files a report might read
 * @returns {{ byFile: Record<string, { report: string, where: string }>, unreadable: string[], read: string[], inside: Record<string, string[]> }}
 */
export function dependencyIndex(reports, candidates) {
  /** @type {Record<string, { report: string, where: string }>} */
  const byFile = {};
  const unreadable = [];
  const read = [];
  /** @type {Record<string, string[]>} */
  const inside = {};
  const names = Array.isArray(candidates) ? candidates : [];

  for (const report of Array.isArray(reports) ? reports : []) {
    const name = String(report.name || '');
    if (!report.bytes) {
      if (report.readable === true) {
        read.push(name);
        for (const source of report.sources || []) {
          const key = String(source.name).toLowerCase();
          if (!byFile[key]) byFile[key] = { report: name, where: source.where };
        }
      } else {
        unreadable.push(name);
        inside[name] = [];
      }
      continue;
    }

    // Whether the report names the file is the decisive test, and it works
    // whatever layout the report uses. The M queries, when they exist, only
    // supply the path to quote in the reason.
    const mentions = mentionsIn(report.bytes, names);
    if (!mentions.readable) {
      unreadable.push(name);
      inside[name] = entryNames(report.bytes);
      continue;
    }
    read.push(name);

    const detail = dataSourcesIn(report.bytes);
    /** @type {Record<string, string>} */
    const where = {};
    for (const source of detail.sources) where[String(source.name).toLowerCase()] = source.where;

    for (const mentioned of mentions.mentioned) {
      const key = mentioned.toLowerCase();
      if (!byFile[key]) {
        byFile[key] = { report: name, where: where[key] || 'its name appears inside the report' };
      }
    }
  }

  return { byFile, unreadable, read, inside };
}
