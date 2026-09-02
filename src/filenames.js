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

/**
 * A date written year-first can be read. An eight-digit run like 04042025
 * cannot: 14042025 and 04142025 mean different things, and being right by
 * coincidence is not being right. Only the unambiguous form is read.
 */
const YEAR_FIRST_DATE = /(20\d{2})[-_.]?(\d{2})[-_.]?(\d{2})/;

/**
 * Underscore is a word character, so \b never fires between "_" and a letter
 * and "_Review_" would not match a status. These bracket on separators instead.
 */
const STATUS_IN_NAME = /(^|[_\-. ])(Draft|Review|Final|Archived)([_\-. ]|$)/i;
const VERSION_IN_NAME = /(^|[_\-. ])v(\d+)(?:[._](\d+))?([_\-. ]|$)/i;

/**
 * An eight-digit run that is not year-first, or a version glued to a word.
 * Neither can be read, and appending a canonical one anyway would leave the
 * file carrying two dates or two versions that disagree — which is worse than
 * either choice. Where one of these is present and unread, the rename is
 * refused and the digits are named so a person can decide.
 */
const UNREADABLE_DATE = /(?:^|[^\d])(\d{8})(?:[^\d]|$)/;
const UNREADABLE_VERSION = /v\d/i;

/**
 * Decide what one file should be renamed to, and be explicit about anything
 * that had to be assumed to get there.
 *
 * Status is the load-bearing assumption. It is a claim about whether a document
 * is approved, and getting it wrong upward — marking a working draft Final —
 * could have someone act on unfinished work. Draft is the weakest claim
 * available, so where the name carries no status the default under-states
 * rather than over-states, and says that it did.
 *
 * @param {{ name: string, path?: string, modified?: string }} file
 * @param {{ assumeStatus?: string | null,
 *           assumeVersion?: { major: number, minor: number } | null,
 *           project?: string }} [options]
 * @returns {{ current: string, path: string, compliant: boolean, problems: string[],
 *             proposed: string, assumptions: string[], blocked: string }}
 */
function planRename(file, options) {
  const spec = file || {};
  const opts = options || {};
  const current = String(spec.name || '');
  const path = String(spec.path || '');

  const problems = diagnose(current);
  /** @type {string[]} */
  const assumptions = [];
  const base = { current, path, compliant: false, problems, proposed: '', assumptions, blocked: '' };

  if (problems.length === 0) return { ...base, compliant: true };

  const assumeStatus = opts.assumeStatus === undefined ? 'Draft' : opts.assumeStatus;
  const assumeVersion =
    opts.assumeVersion === undefined ? { major: 0, minor: 1 } : opts.assumeVersion;

  const { stem, extension } = splitExtension(current);

  const dateMatch = YEAR_FIRST_DATE.exec(stem);
  const modified = spec.modified ? String(spec.modified).slice(0, 10) : '';
  let date = '';
  if (dateMatch) {
    date = dateMatch[1] + '-' + dateMatch[2] + '-' + dateMatch[3];
  } else if (isRealDate(modified)) {
    date = modified;
    assumptions.push('date taken from the file timestamp, not the name');
  }

  const statusMatch = STATUS_IN_NAME.exec(stem);
  let status = '';
  if (statusMatch && statusMatch[2]) {
    const word = statusMatch[2];
    status = word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
  } else if (assumeStatus) {
    status = String(assumeStatus);
    assumptions.push('status assumed to be ' + status + ', the name does not say');
  }

  const versionMatch = VERSION_IN_NAME.exec(stem);
  /** @type {{ major: number, minor: number } | null} */
  let version = null;
  if (versionMatch) {
    version = { major: Number(versionMatch[2]), minor: Number(versionMatch[3] || 0) };
  } else if (assumeVersion) {
    version = assumeVersion;
    assumptions.push('version assumed to be v' + assumeVersion.major + '.' + assumeVersion.minor);
  }

  if (!dateMatch) {
    const stray = UNREADABLE_DATE.exec(stem);
    if (stray) {
      return {
        ...base,
        blocked:
          'the name contains ' + stray[1] + ', which is a date under one convention and a different date under another — renaming would leave two dates in the name',
      };
    }
  }

  if (!versionMatch && UNREADABLE_VERSION.test(stem)) {
    return {
      ...base,
      blocked: 'the name looks like it already carries a version but not in a form that can be read — renaming would leave two',
    };
  }

  if (!status) return { ...base, blocked: 'no status in the name and none assumed' };
  if (!date) return { ...base, blocked: 'no date in the name and none on the file' };
  if (!version) return { ...base, blocked: 'no version in the name and none assumed' };

  // Whatever is left once the rebuilt parts are removed is the document name.
  // Cut by position on the original stem, right to left, rather than by
  // replacing text: removing the status first shifts the version's text so a
  // second replace silently misses and leaves "Reportv0.1-copy" behind.
  /** @type {Array<[number, number]>} */
  const cuts = [];
  if (dateMatch) cuts.push([dateMatch.index, dateMatch.index + dateMatch[0].length]);
  if (statusMatch) cuts.push([statusMatch.index, statusMatch.index + statusMatch[0].length]);
  if (versionMatch) cuts.push([versionMatch.index, versionMatch.index + versionMatch[0].length]);

  // Adjacent parts can share a separator, so overlapping cuts are merged before
  // any are applied.
  cuts.sort((a, b) => a[0] - b[0]);
  /** @type {Array<[number, number]>} */
  const merged = [];
  for (const cut of cuts) {
    const last = merged[merged.length - 1];
    if (last && cut[0] <= last[1]) last[1] = Math.max(last[1], cut[1]);
    else merged.push([cut[0], cut[1]]);
  }

  let remainder = stem;
  for (let i = merged.length - 1; i >= 0; i -= 1) {
    const cut = merged[i];
    if (!cut) continue;
    remainder = remainder.slice(0, cut[0]) + ' ' + remainder.slice(cut[1]);
  }

  let proposed = '';
  try {
    proposed = proposeFileName({
      project: String(opts.project || ''),
      document: remainder,
      date,
      status,
      version,
      extension,
    });
  } catch (err) {
    const reason = err instanceof Error ? err.message : String(err);
    return { ...base, blocked: reason.slice(0, 160) };
  }

  if (proposed === current) return { ...base, compliant: true };

  return { ...base, proposed };
}

/**
 * Plan a whole library at once, which is the only level at which collisions are
 * visible. Two files renamed onto one name would destroy one of them, so both
 * are refused rather than either being guessed at.
 *
 * @param {Array<{ name: string, path?: string, modified?: string }>} files
 * @param {{ assumeStatus?: string | null,
 *           assumeVersion?: { major: number, minor: number } | null,
 *           project?: string }} [options]
 * @returns {ReturnType<typeof planRename>[]}
 */
function planRenames(files, options) {
  const list = Array.isArray(files) ? files : [];
  const rows = list.map((f) => planRename(f, options));

  // Every name that would exist afterwards, including the ones already correct
  // which keep theirs — renaming onto an existing file is a collision too.
  /** @type {Record<string, number>} */
  const wanted = {};
  for (const row of rows) {
    const target = String(row.proposed || row.current).toLowerCase();
    wanted[target] = (wanted[target] || 0) + 1;
  }

  for (const row of rows) {
    if (!row.proposed) continue;
    const target = row.proposed.toLowerCase();
    if ((wanted[target] || 0) > 1) {
      row.blocked = 'two files would end up with the same name — ' + row.proposed;
      row.proposed = '';
    }
  }

  return rows;
}

export { STATUSES, diagnose, parseFileName, planRename, planRenames, proposeFileName };
