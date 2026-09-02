/**
 * Yancey's document naming convention, sent 28 August 2026 and again on the
 * 31st:
 *
 *   [Project]_[Document Name]_[YYYY-MM-DD]_[Status]_v#.#
 *   SCE_PM_Ride-Along-Plan_2026-09-01_Draft_v0.1.docx
 *
 * His template shows three bracketed sections before the date, but every
 * worked example he gave puts two or three underscore-separated sections
 * there. So the name is whatever precedes the date, and the date, status and
 * version are the part held to an exact shape. That reading accepts all four
 * of his examples; a stricter one rejects three of them.
 *
 * Nothing here talks to SharePoint. Deciding what a file should be called is
 * the whole of the judgement, and it is worth being able to test it without a
 * credential, a tenant, or a network.
 */

/** The four he listed, in his order. Anything else is not a status. */
const STATUSES = ['Draft', 'Review', 'Final', 'Archived'];

/**
 * "v1.0 — First approved version" in his version guide, so a document marked
 * Final while still numbered v0.x contradicts itself.
 */
const FIRST_APPROVED_MAJOR = 1;

/** Names he asked for by name: "Do not use names like Final_Final, Latest, or New." */
const BANNED_NAMES = [
  { pattern: /final[_-]final/i, label: 'Final_Final' },
  { pattern: /(^|[_-])latest([_-]|$)/i, label: 'Latest' },
  { pattern: /(^|[_-])new([_-]|$)/i, label: 'New' },
];

const DATE_SHAPE = /^\d{4}-\d{2}-\d{2}$/;
const VERSION_SHAPE = /^v(\d+)\.(\d+)$/;

/**
 * Letters, digits, hyphen and underscore — everything else he asked us to
 * avoid. The dot is allowed because the version carries one of its own: v0.1
 * sits inside the stem, not in the extension.
 */
const ALLOWED_STEM = /^[A-Za-z0-9._-]+$/;

/**
 * @param {string} value
 * @returns {{ stem: string, extension: string }}
 */
function splitExtension(value) {
  const name = String(value || '');
  const dot = name.lastIndexOf('.');
  // A trailing dot, or a dot inside the version (v0.1), is not an extension.
  if (dot <= 0 || dot === name.length - 1) return { stem: name, extension: '' };
  const extension = name.slice(dot + 1);
  if (!/^[A-Za-z0-9]+$/.test(extension)) return { stem: name, extension: '' };
  return { stem: name.slice(0, dot), extension };
}

/**
 * Real calendar date, not just the right shape. 2026-02-31 passes a regex.
 * @param {string} value
 */
function isRealDate(value) {
  if (!DATE_SHAPE.test(value)) return false;
  const [y, m, d] = value.split('-').map(Number);
  if (m === undefined || d === undefined || y === undefined) return false;
  const parsed = new Date(Date.UTC(y, m - 1, d));
  return (
    parsed.getUTCFullYear() === y && parsed.getUTCMonth() === m - 1 && parsed.getUTCDate() === d
  );
}

/**
 * Pull a file name apart on his convention.
 *
 * Anchored from the right: version, then status, then date. Whatever is left
 * on the front is the name, which is why a two-section and a three-section
 * name both work.
 *
 * @param {string} value
 * @returns {{ ok: boolean, name: string, date: string, status: string,
 *             version: { major: number, minor: number } | null,
 *             extension: string }}
 */
function parseFileName(value) {
  const { stem, extension } = splitExtension(value);
  const empty = { ok: false, name: '', date: '', status: '', version: null, extension };
  if (!stem) return empty;

  const parts = stem.split('_');
  if (parts.length < 4) return { ...empty, name: stem };

  const version = String(parts[parts.length - 1] || '');
  const status = String(parts[parts.length - 2] || '');
  const date = String(parts[parts.length - 3] || '');
  const name = parts.slice(0, parts.length - 3).join('_');

  const versionMatch = VERSION_SHAPE.exec(version);
  const ok =
    name.length > 0 &&
    isRealDate(date) &&
    STATUSES.indexOf(status) !== -1 &&
    versionMatch !== null;

  return {
    ok,
    name,
    date,
    status,
    version: versionMatch
      ? { major: Number(versionMatch[1]), minor: Number(versionMatch[2]) }
      : null,
    extension,
  };
}

/**
 * Everything wrong with a file name, phrased so it can be shown to a person.
 * An empty array means it follows the convention.
 *
 * @param {string} value
 * @returns {string[]}
 */
function diagnose(value) {
  const raw = String(value || '');
  /** @type {string[]} */
  const problems = [];

  if (!raw.trim()) return ['no file name'];

  const { stem, extension } = splitExtension(raw);

  if (/\s/.test(raw)) {
    problems.push('contains a space — use hyphens within a phrase and underscores between sections');
  }

  if (stem && !ALLOWED_STEM.test(stem)) {
    problems.push('contains a special character — letters, numbers, hyphens and underscores only');
  }

  for (const banned of BANNED_NAMES) {
    if (banned.pattern.test(stem)) {
      problems.push('uses "' + banned.label + '" — the version number says which one is current');
    }
  }

  const parts = stem.split('_');
  if (parts.length < 4) {
    problems.push(
      'does not follow [Project]_[Document]_[YYYY-MM-DD]_[Status]_v#.# — for example SCE_PM_Ride-Along-Plan_2026-09-01_Draft_v0.1',
    );
    return problems;
  }

  const parsed = parseFileName(raw);
  const tailVersion = String(parts[parts.length - 1] || '');
  const tailStatus = String(parts[parts.length - 2] || '');
  const tailDate = String(parts[parts.length - 3] || '');

  if (!VERSION_SHAPE.test(tailVersion)) {
    // The status may have been the last section, meaning the version is absent.
    if (STATUSES.indexOf(tailVersion) !== -1) {
      problems.push('has no version — keep the v#.# even when the document is Final');
    } else {
      problems.push('ends with "' + tailVersion + '" rather than a version like v0.1 or v1.0');
    }
    return problems;
  }

  if (STATUSES.indexOf(tailStatus) === -1) {
    problems.push(
      '"' + tailStatus + '" is not one of the four statuses — ' + STATUSES.join(', '),
    );
  }

  if (!DATE_SHAPE.test(tailDate)) {
    problems.push('"' + tailDate + '" is not YYYY-MM-DD, which is what makes the files sort');
  } else if (!isRealDate(tailDate)) {
    problems.push('"' + tailDate + '" is not a real date');
  }

  if (
    parsed.status === 'Final' &&
    parsed.version !== null &&
    parsed.version.major < FIRST_APPROVED_MAJOR
  ) {
    problems.push('is marked Final but numbered v' + parsed.version.major + '.' + parsed.version.minor + ' — an approved version starts at v1.0');
  }

  if (!extension) {
    problems.push('has no file extension');
  }

  return problems;
}

/**
 * Tidy one section: spaces and stray punctuation become hyphens, and anything
 * he asked us to avoid is dropped rather than passed through.
 * @param {string} value
 */
function tidySection(value) {
  return String(value || '')
    .replace(/[‐-―]/g, '-')
    .replace(/[^A-Za-z0-9_\s-]/g, ' ')
    .trim()
    .replace(/\s+/g, '-')
    .replace(/-{2,}/g, '-')
    .replace(/_{2,}/g, '_')
    .replace(/^[-_]+|[-_]+$/g, '');
}

/**
 * Build a name that follows the convention. Throws rather than guessing when
 * something it cannot invent is missing — a wrong status on a client document
 * is worse than a refusal.
 *
 * @param {{ project: string, document: string, date: string, status: string,
 *           version: { major: number, minor: number }, extension: string }} parts
 * @returns {string}
 */
function proposeFileName(parts) {
  const spec = parts || {};
  const status = String(spec.status || '');
  if (STATUSES.indexOf(status) === -1) {
    throw new Error('status must be one of ' + STATUSES.join(', ') + ', got "' + status + '"');
  }

  const date = String(spec.date || '');
  if (!isRealDate(date)) {
    throw new Error('date must be a real YYYY-MM-DD, got "' + date + '"');
  }

  const version = spec.version || { major: 0, minor: 1 };
  const project = tidySection(spec.project);
  const document = tidySection(spec.document);
  if (!project && !document) throw new Error('needs a project or a document name');

  const extension = String(spec.extension || '').replace(/[^A-Za-z0-9]/g, '');
  const stem = [project, document].filter(Boolean).join('_');

  return (
    stem +
    '_' +
    date +
    '_' +
    status +
    '_v' +
    String(version.major) +
    '.' +
    String(version.minor) +
    (extension ? '.' + extension : '')
  );
}

export { STATUSES, diagnose, parseFileName, proposeFileName };
