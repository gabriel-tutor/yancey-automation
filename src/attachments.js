/**
 * Deciding what to do with an attachment, and what to say about it.
 *
 * Extraction itself happens in the workflow — PDFs and spreadsheets go through
 * n8n's own `extractFromFile`, images go to a vision model, Office files go
 * through `unzip.js`. This module holds the decisions around that: which
 * attachments are worth opening, which route each one takes, and how the
 * results are compressed into something small enough to put in front of the
 * classifier without drowning the email itself.
 */

/**
 * Attachments larger than this are recorded but not opened.
 *
 * 8 MB was a guess and it was too low: a real 22 MB PDF was refused in dev,
 * then read end to end in 21 seconds once the ceiling was lifted. 25 MB keeps
 * us under the model's 32 MB limit for a PDF while clearing the size most mail
 * servers will accept in the first place.
 */
const MAX_BYTES = 25 * 1024 * 1024;

/**
 * Inline images below this are signatures, logos and tracking pixels. Opening
 * them costs a vision call per email and returns "a company logo" every time.
 */
const INLINE_IMAGE_FLOOR = 40 * 1024;

/**
 * No image this small holds anything a person could read — tracking pixels,
 * bullets, spacers, icons. Skipped whatever it is called and however it
 * arrived.
 */
const TINY_IMAGE_BYTES = 12 * 1024;

/**
 * A signature logo only counts as inline while the message still carries the
 * cid reference. Forwarding strips it, and plenty of clients attach the logo
 * outright, so the flag alone is not enough: measured in dev, an 18 KB SCE
 * logo cost a vision call to be told it was a logo. Names are the other half
 * of the signal.
 */
const DECORATIVE_IMAGE_WORD = /\b(logo|signature|banner|letterhead|footer|badge)\b/i;
const GENERIC_IMAGE_NAME = /^(image|img|inline|att|attachment|pixel|spacer|icon)[-_ ]*\d*\.[a-z0-9]+$/i;

/** How many attachments on one message are opened at all. */
const MAX_ATTACHMENTS = 8;

/**
 * How many of them may cost a model call. A 22 MB PDF takes about 21 seconds
 * and WF-A waits for the answer before it classifies, so this is what stops one
 * message holding up the poller. Mail servers cap a message near 25 MB anyway,
 * which does most of the work; this is the belt to that pair of braces.
 */
const MAX_MODEL_CALLS = 3;

/** Total characters of attachment text allowed into the prompt. */
const TOTAL_CHAR_BUDGET = 6000;

/** Per-attachment ceiling, so one large deck cannot crowd out three others. */
const PER_ITEM_CHAR_BUDGET = 2500;

/**
 * Extension to route. Content-Type is checked first, but Outlook frequently
 * reports `application/octet-stream` for perfectly ordinary files, so the
 * extension is what actually decides most of the time.
 * @type {Record<string, string>}
 */
const BY_EXTENSION = {
  png: 'image', jpg: 'image', jpeg: 'image', gif: 'image', webp: 'image',
  bmp: 'image', tif: 'image', tiff: 'image', heic: 'image',

  pdf: 'pdf',

  xlsx: 'office', xlsm: 'office',

  docx: 'office', pptx: 'office',

  txt: 'text', md: 'text', csv: 'text', tsv: 'text', json: 'text', xml: 'text',
  html: 'text', htm: 'text', log: 'text', yaml: 'text', yml: 'text', ics: 'text',
  eml: 'text', rtf: 'text',
};

/**
 * Types that are genuinely not readable as text, listed explicitly so the
 * summary can say what a thing *was* rather than shrugging. `.pbix` matters
 * here: Power BI files are common in this inbox and are opaque binaries.
 * @type {Record<string, string>}
 */
const KNOWN_UNREADABLE = {
  pbix: 'Power BI report',
  xls: 'Excel 97-2003 workbook',
  ods: 'OpenDocument spreadsheet',
  doc: 'Word 97-2003 document',
  ppt: 'PowerPoint 97-2003 deck',
  zip: 'archive', '7z': 'archive', rar: 'archive', tar: 'archive', gz: 'archive',
  exe: 'executable', dll: 'executable', msi: 'installer',
  mp4: 'video', mov: 'video', avi: 'video',
  mp3: 'audio', wav: 'audio', m4a: 'audio',
  psd: 'image source file', ai: 'image source file',
};

/**
 * @param {string} name
 * @returns {string}
 */
function extensionOf(name) {
  const clean = String(name || '').toLowerCase().trim();
  const dot = clean.lastIndexOf('.');
  if (dot === -1 || dot === clean.length - 1) return '';
  return clean.slice(dot + 1);
}

/**
 * Decide what happens to one attachment.
 *
 * @param {{ name?: string, contentType?: string, size?: number, isInline?: boolean,
 *           '@odata.type'?: string }} att
 * @param {{ maxBytes?: number, inlineImageFloor?: number }} [limits]
 * @returns {{ route: string, reason: string, ext: string, label: string }}
 */
function triageAttachment(att, limits) {
  const opts = limits || {};
  const maxBytes = opts.maxBytes === undefined ? MAX_BYTES : opts.maxBytes;
  const inlineFloor = opts.inlineImageFloor === undefined ? INLINE_IMAGE_FLOOR : opts.inlineImageFloor;

  const name = String(att.name || 'attachment');
  const ext = extensionOf(name);
  const size = Number(att.size) || 0;
  const type = String(att.contentType || '').toLowerCase();
  const label = name + ' (' + describeSize(size) + ')';

  // Graph returns itemAttachment and referenceAttachment alongside real files.
  // Neither has bytes to read: one is an embedded Outlook item, the other a
  // link to cloud storage.
  const odata = String(att['@odata.type'] || '');
  if (odata && odata.indexOf('fileAttachment') === -1) {
    const kind = odata.indexOf('reference') !== -1 ? 'a link to cloud storage' : 'an embedded Outlook item';
    return { route: 'skip', reason: kind + ', no file to read', ext, label };
  }

  if (size > maxBytes) {
    return { route: 'skip', reason: 'larger than the ' + describeSize(maxBytes) + ' limit', ext, label };
  }

  const looksImage = type.startsWith('image/') || BY_EXTENSION[ext] === 'image';
  if (looksImage) {
    const decorative = DECORATIVE_IMAGE_WORD.test(name) || GENERIC_IMAGE_NAME.test(name);
    if (size < inlineFloor && (att.isInline || decorative)) {
      return { route: 'skip', reason: 'signature/logo image below the reading threshold', ext, label };
    }
    if (size < TINY_IMAGE_BYTES) {
      return { route: 'skip', reason: 'image too small to hold anything readable', ext, label };
    }
  }

  if (KNOWN_UNREADABLE[ext]) {
    return { route: 'unreadable', reason: KNOWN_UNREADABLE[ext] + ', cannot be read as text', ext, label };
  }

  const byExt = BY_EXTENSION[ext];
  if (byExt) return { route: byExt, reason: '', ext, label };

  if (looksImage) return { route: 'image', reason: '', ext, label };
  if (type === 'application/pdf') return { route: 'pdf', reason: '', ext, label };
  if (type.indexOf('spreadsheetml') !== -1) return { route: 'office', reason: '', ext, label };
  if (type.indexOf('wordprocessingml') !== -1 || type.indexOf('presentationml') !== -1) return { route: 'office', reason: '', ext, label };
  if (type.startsWith('text/')) return { route: 'text', reason: '', ext, label };

  return { route: 'unreadable', reason: 'unrecognised file type' + (ext ? ' (.' + ext + ')' : ''), ext, label };
}

/**
 * @param {number} bytes
 * @returns {string}
 */
function describeSize(bytes) {
  const n = Number(bytes) || 0;
  if (n >= 1024 * 1024) return (n / (1024 * 1024)).toFixed(1).replace(/\.0$/, '') + ' MB';
  if (n >= 1024) return Math.round(n / 1024) + ' KB';
  return n + ' bytes';
}

/**
 * Squeeze extracted text down to something worth reading.
 *
 * Spreadsheets in particular arrive as thousands of near-identical rows; the
 * first screenful tells the model what the sheet is about and the rest is
 * repetition that would push the actual email out of the window.
 *
 * @param {string} value
 * @param {number} [limit]
 * @returns {string}
 */
function condense(value, limit) {
  const cap = limit === undefined ? PER_ITEM_CHAR_BUDGET : limit;
  const cleaned = String(value || '')
    .replace(/\r\n?/g, '\n')
    .replace(/[ \t]+/g, ' ')
    // Trim each line, not just the ends. PDF and spreadsheet extraction both
    // emit ragged leading indentation, and in a 6,000-character budget that is
    // real content being spent on whitespace.
    .replace(/^[ \t]+/gm, '')
    .replace(/[ \t]+$/gm, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
  if (cleaned.length <= cap) return cleaned;
  return cleaned.slice(0, cap).replace(/\s+\S*$/, '') + '\n…truncated…';
}

/**
 * Turn per-attachment results into the block that goes in front of the model.
 *
 * Everything is named, including what could not be read. An attachment the
 * assistant silently ignored is worse than one it says it could not open —
 * Yancey can act on the second and cannot see the first.
 *
 * @param {Array<{ label: string, route: string, reason?: string, text?: string, error?: string }>} items
 * @param {{ totalBudget?: number, perItem?: number }} [limits]
 * @returns {{ text: string, read: number, skipped: number, failed: number }}
 */
function summariseAttachments(items, limits) {
  const opts = limits || {};
  const total = opts.totalBudget === undefined ? TOTAL_CHAR_BUDGET : opts.totalBudget;
  const perItem = opts.perItem === undefined ? PER_ITEM_CHAR_BUDGET : opts.perItem;

  const list = Array.isArray(items) ? items : [];
  if (list.length === 0) return { text: '', read: 0, skipped: 0, failed: 0 };

  // Share the budget between whatever actually produced text.
  const withText = list.filter((i) => i && typeof i.text === 'string' && i.text.trim().length > 0);
  const share = withText.length > 0 ? Math.max(400, Math.floor(total / withText.length)) : perItem;

  const lines = [];
  let read = 0;
  let skipped = 0;
  let failed = 0;
  let spent = 0;

  for (const item of list) {
    if (!item) continue;
    const label = String(item.label || 'attachment');

    if (item.error) {
      failed += 1;
      lines.push('- ' + label + ' — could not be read (' + String(item.error).slice(0, 120) + ')');
      continue;
    }
    if (item.route === 'skip' || item.route === 'unreadable') {
      skipped += 1;
      lines.push('- ' + label + ' — not read: ' + (item.reason || 'skipped'));
      continue;
    }

    const body = condense(item.text || '', Math.min(share, perItem, Math.max(0, total - spent)));
    if (!body) {
      skipped += 1;
      lines.push('- ' + label + ' — opened, but no readable text in it');
      continue;
    }
    read += 1;
    spent += body.length;
    lines.push('- ' + label + ':\n' + body);
  }

  return { text: lines.join('\n\n'), read, skipped, failed };
}

/**
 * Triage a whole message. The per-attachment call cannot see the two limits
 * that only mean anything across the set — how many files are opened at all,
 * and how many of them may cost a model call — so they live here rather than
 * loose in a workflow node where nothing would test them.
 *
 * @param {any[]} attachments Graph attachment metadata, in the order listed
 * @param {{ maxBytes?: number, inlineImageFloor?: number,
 *           maxAttachments?: number, maxModelCalls?: number }} [limits]
 * @returns {Array<{ route: string, reason: string, ext: string, label: string,
 *                   needsBytes: boolean, index: number }>}
 */
function triageAll(attachments, limits) {
  const opts = limits || {};
  const maxAttachments = opts.maxAttachments === undefined ? MAX_ATTACHMENTS : opts.maxAttachments;
  const maxModelCalls = opts.maxModelCalls === undefined ? MAX_MODEL_CALLS : opts.maxModelCalls;
  const list = Array.isArray(attachments) ? attachments : [];

  const out = [];
  let modelCalls = 0;

  for (let i = 0; i < list.length; i += 1) {
    const verdict = triageAttachment(list[i] || {}, opts);
    let route = verdict.route;
    let reason = verdict.reason;

    if (i >= maxAttachments) {
      route = 'skip';
      reason = 'beyond the first ' + maxAttachments + ' attachments on this message';
    } else if (route === 'pdf' || route === 'image') {
      // Only a file that would actually reach the model is charged for one.
      // Office files are unzipped in process, and anything already skipped
      // costs nothing at all.
      if (modelCalls >= maxModelCalls) {
        route = 'skip';
        reason = 'only the first ' + maxModelCalls + ' files that need reading are opened per message';
      } else {
        modelCalls += 1;
      }
    }

    out.push({
      route,
      reason,
      ext: verdict.ext,
      label: verdict.label,
      needsBytes: route === 'text' || route === 'office' || route === 'pdf' || route === 'image',
      index: i,
    });
  }

  return out;
}

export {
  MAX_BYTES,
  MAX_ATTACHMENTS,
  MAX_MODEL_CALLS,
  TINY_IMAGE_BYTES,
  INLINE_IMAGE_FLOOR,
  TOTAL_CHAR_BUDGET,
  PER_ITEM_CHAR_BUDGET,
  BY_EXTENSION,
  KNOWN_UNREADABLE,
  triageAttachment,
  triageAll,
  describeSize,
  condense,
  summariseAttachments,
};
