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
 * @returns {{ readable: boolean, parts: string[], sources: Array<{ name: string, where: string }> }}
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
    return { readable: false, parts: [], sources: [] };
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
  if (read.length === 0) return { readable: false, parts: [], sources: [] };
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
    sources: [...found.entries()].map(([, where]) => {
      const cut = Math.max(where.lastIndexOf('\\'), where.lastIndexOf('/'));
      return { name: cut === -1 ? where : where.slice(cut + 1), where };
    }),
  };
}

/**
 * Turn what each report reads into the lookup the organiser needs.
 *
 * A report that could not be opened is recorded as unreadable rather than as
 * reading nothing, because those are opposite conclusions: one means "safe to
 * rename", the other means "we still cannot tell".
 *
 * @param {Array<{ name: string, path?: string, bytes?: Uint8Array, readable?: boolean, sources?: Array<{ name: string, where: string }> }>} reports
 * @returns {{ byFile: Record<string, { report: string, where: string }>, unreadable: string[], read: string[] }}
 */
export function dependencyIndex(reports) {
  /** @type {Record<string, { report: string, where: string }>} */
  const byFile = {};
  const unreadable = [];
  const read = [];

  for (const report of Array.isArray(reports) ? reports : []) {
    const name = String(report.name || '');
    const found = report.bytes
      ? dataSourcesIn(report.bytes)
      : { readable: report.readable === true, parts: [], sources: report.sources || [] };
    if (!found.readable) {
      unreadable.push(name);
      continue;
    }
    read.push(name);
    for (const source of found.sources) {
      const key = String(source.name).toLowerCase();
      if (!byFile[key]) byFile[key] = { report: name, where: source.where };
    }
  }

  return { byFile, unreadable, read };
}
