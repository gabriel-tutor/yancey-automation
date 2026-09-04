/**
 * Running the organiser without a person who can read the code.
 *
 * Yancey asked for this in as many words:
 *
 *   "can we do a frontend / abstract tool for yancey so he can run it on his
 *    own like chrome extension?"
 *
 * The shape he named is an example, not the requirement. The requirement is
 * that he can start a run, read what it wants to do, choose, and see what
 * happened — without n8n's editor, without me, and without anybody editing
 * config/approved-plan.json by hand, which is the step that left six correctly
 * named documents sitting untouched in front of a client.
 *
 * So the front end is three pages: start, choose, done. What makes that safe
 * rather than merely convenient is that the middle page is the approval. The
 * run that produced the list is the run that gets applied, because the ticks
 * become an approved plan and an approved plan *binds* — the model is not
 * deterministic, and a review that does not bind to what happens is a
 * demonstration rather than an approval.
 *
 * Three functions, because there are exactly three joins between a form and a
 * workflow that already works:
 *
 *   mergeOverrides   what a form submission is allowed to change about a run
 *   reviewFields     a run's proposals turned into things to tick
 *   approvedFrom     the ticks turned back into a plan and the switches for it
 *
 * The fourth, outcomeHtml, is the page at the end. It exists here rather than
 * in an expression because "what changed" is the sentence the client reads, and
 * a sentence that matters is a sentence worth testing.
 */

import { STATUSES } from './filenames.js';

/**
 * @typedef {Object} Offered
 * @property {number} index      1-based, and the thing ticks are matched on
 * @property {string} label      what the person reads
 * @property {'rename'|'move'|'rename+move'|'folder'|'folderRename'} kind
 * @property {string} [from]     the path as it is now
 * @property {string} [to]       the name it would take
 * @property {string} [into]     the folder it would land in
 * @property {boolean} [createsFolder] whether that folder does not exist yet
 */

/**
 * The only settings a form submission may change.
 *
 * Everything absent from this list — the host, the model, how much of a
 * document may be read, how many pages may be fetched — is a property of the
 * deployment and not a choice at the point of use.
 *
 * The site and the library are on it, and they are the exception that proves
 * the rule. Yancey asked for them on 4 September — *"what happens if I want to
 * point it at another folder?"* — and pointing a workflow that renames
 * documents at somewhere new is plainly widening what a run does, not
 * narrowing it. So they are not a caller's choice either: they are a choice
 * between options the DEPLOYMENT stated, which is the same shape as
 * `assumeStatus` picking from the statuses the convention recognises. A
 * deployment that states nothing stays locked exactly as it was.
 *
 * The host is on neither list and never will be. It is the tenant.
 */
const OVERRIDABLE = [
  'applyRenames',
  'applyMoves',
  'applyStructure',
  'applyFolderRenames',
  'applyProposed',
  'revert',
  'renameReports',
  'readContents',
  'approvedPlan',
  'maxChangesPerRun',
  'maxNewFolders',
  'minPerFolder',
  'assumeStatus',
  'sitePath',
  'libraryTitle',
];

/**
 * Ceilings. A caller may lower one — that is choosing to do less — and never
 * raise it past what the deployment allows.
 */
const LOWER_ONLY = ['maxChangesPerRun', 'maxNewFolders'];

/**
 * A floor, and the direction is inverted for exactly that reason: minPerFolder
 * is how many documents a new folder must hold to be worth making, so RAISING
 * it means fewer new folders. The conservative end of this one is up.
 */
const RAISE_ONLY = ['minPerFolder'];

/**
 * Settings a caller may switch OFF but never ON.
 *
 * The rule the whole of this file is built on is that a caller may *narrow*
 * what a run does and never widen it. Turning renameReports on puts the .pbix
 * files in scope, whose paths are referenced from outside the library; turning
 * readContents on sends document text out of the tenant. Both are decisions
 * that belong to the deployment. Turning either off is the caller choosing to
 * do less, which needs no permission.
 *
 * (renameReports was absent from OVERRIDABLE until the merge moved out of the
 * apply node. It is read both where naming candidates are picked and where they
 * are applied, and while the merge lived inside Apply it could only reach the
 * second — honouring it would have changed a count without changing what
 * happened. Config feeds every node, so both ends now read the same answer.)
 */
const NARROW_ONLY = ['readContents', 'renameReports'];

/** The five that turn a review into a write. */
const ARMING = [
  'applyRenames',
  'applyMoves',
  'applyStructure',
  'applyFolderRenames',
  'applyProposed',
];

/** A plan line is a rename, a move, or both — and nothing else survives.
 * @param {any} line @returns {{from: string, to?: string, into?: string} | null} */
function planLine(line) {
  if (!line || typeof line !== 'object') return null;
  const from = typeof line.from === 'string' ? line.from.trim() : '';
  if (!from) return null;
  /** @type {{from: string, to?: string, into?: string}} */
  const out = { from };
  if (typeof line.to === 'string' && line.to.trim()) out.to = line.to.trim();
  if (typeof line.into === 'string' && line.into.trim()) out.into = line.into.trim();
  // A line that says neither what to call it nor where to put it asks for
  // nothing, and a plan of those would arm the switches to do nothing.
  if (!out.to && !out.into) return null;
  return out;
}

/**
 * A SharePoint site collection path, and nothing that is not one.
 *
 * Every segment has to start with a letter or a digit, which is what refuses
 * `..` without having to reason about path climbing separately. A leading slash
 * is required, so a full URL cannot be smuggled in as a path.
 */
const SITE_PATH = /^\/(?:sites|teams)\/[A-Za-z0-9][A-Za-z0-9._-]*(?:\/[A-Za-z0-9][A-Za-z0-9._-]*)*$/i;

/**
 * A document library title, held to plain characters.
 *
 * This one is not fussiness. The title is interpolated into
 * `getbytitle('<title>')` in the request URL, and while it was a constant that
 * was nobody's problem — the moment a caller can set it, an apostrophe closes
 * the quote and the rest of the string is being read as OData. It is refused
 * rather than escaped, because a refusal is something a test can check and an
 * escape is a thing you get wrong once.
 */
const LIBRARY_TITLE = /^[A-Za-z0-9][A-Za-z0-9 ._-]{0,63}$/;

/** Longest site path worth entertaining; SharePoint's own URL limits are lower. */
const SITE_PATH_MAX = 200;

/**
 * Has this deployment opened targeting at all?
 *
 * Three states, and the first is what every deployment has until somebody says
 * otherwise: locked to the site it was built with.
 *
 * @param {any} base
 * @returns {{ open: boolean, listed: string[] }}
 */
function targeting(base) {
  const listed = Array.isArray(base && base.allowedSites)
    ? base.allowedSites.map((/** @type {any} */ s) => String(s)).filter(Boolean)
    : [];
  return { open: listed.length > 0 || (base && base.allowAnySiteOnHost === true), listed };
}

/**
 * Where a caller asked to point the run, held to what the deployment allows.
 *
 * @param {any} base the built configuration
 * @param {string} key 'sitePath' or 'libraryTitle'
 * @param {any} value what the caller asked for
 * @returns {{ value: string } | { why: string }}
 */
export function chooseTarget(base, key, value) {
  const { open, listed } = targeting(base);
  const current = String((base && base[key]) || '');

  if (!open) {
    return {
      why: key + ' cannot be set from outside: this deployment is fixed to ' +
        String((base && base.sitePath) || '') + ', and opening that is a deployment change',
    };
  }

  const wanted = typeof value === 'string' ? value.trim() : '';

  if (key === 'libraryTitle') {
    if (!LIBRARY_TITLE.test(wanted)) {
      return {
        why: 'libraryTitle must be plain letters, digits, spaces, dots, dashes or underscores — ' +
          JSON.stringify(String(value)) + ' is not, so ' + current + ' was kept',
      };
    }
    return { value: wanted };
  }

  if (wanted.length > SITE_PATH_MAX || !SITE_PATH.test(wanted)) {
    return {
      why: 'sitePath must look like /sites/Something — ' + JSON.stringify(String(value)) +
        ' does not, so ' + current + ' was kept',
    };
  }

  if (listed.length > 0) {
    // Answer with the spelling the deployment used, not the caller's, so case
    // is never the thing that decides which library gets written to.
    const found = listed.find((site) => site.toLowerCase() === wanted.toLowerCase());
    if (!found) {
      return {
        why: 'sitePath ' + wanted + ' is not one this deployment allows: ' + listed.join(', '),
      };
    }
    return { value: found };
  }

  return { value: wanted };
}

/**
 * Fold a caller's overrides onto the deployed config.
 *
 * Two rules do the work, and both are about the difference between a workflow
 * that *may* write and one that *is* writing.
 *
 * `canArm` is the deployment's own answer to whether this copy is allowed to
 * touch the library at all. The dev sandbox reads the same library through the
 * same credential as production — duplicating the workflow isolated the
 * deployment and not the writes — so a form pointed at the sandbox would
 * otherwise have been a form that changes the client's documents. With
 * canArm false every arming switch is forced off no matter what arrives, and
 * the run says so rather than quietly reviewing when it was asked to apply.
 *
 * The cap only ever goes down. Raising a blast radius is a deployment
 * decision; a person at a form may choose to do less than the build allows and
 * never more.
 *
 * @param {Record<string, any>} base the config as deployed
 * @param {any} incoming what the caller asked for
 * @param {{ canArm?: boolean }} [options]
 * @returns {{ config: Record<string, any>, refused: string[] }}
 */
export function mergeOverrides(base, incoming, options) {
  const config = Object.assign({}, base || {});
  /** @type {string[]} */
  const refused = [];
  const opts = options || {};
  const canArm = opts.canArm === undefined ? config.canArm === true : opts.canArm === true;

  let asked = incoming;
  if (typeof asked === 'string') {
    try {
      asked = JSON.parse(asked);
    } catch (err) {
      // A caller that sent something unparseable gets a review, not a guess.
      return { config, refused: ['overrides were not readable, so none were applied'] };
    }
  }
  if (!asked || typeof asked !== 'object' || Array.isArray(asked)) return { config, refused };

  for (const key of Object.keys(asked)) {
    if (OVERRIDABLE.indexOf(key) === -1) {
      refused.push(key + ' cannot be set from outside the deployment');
      continue;
    }
    const value = asked[key];

    if (ARMING.indexOf(key) !== -1) {
      if (value !== true) {
        config[key] = false;
        continue;
      }
      if (!canArm) {
        refused.push(key + ' was asked for, but this copy is not allowed to change the library');
        config[key] = false;
        continue;
      }
      config[key] = true;
      continue;
    }

    if (NARROW_ONLY.indexOf(key) !== -1) {
      if (value === false) {
        config[key] = false;
        continue;
      }
      if (base && base[key] === true) {
        config[key] = true;
        continue;
      }
      refused.push(key + ' can be switched off from outside the deployment but not on');
      continue;
    }

    if (key === 'approvedPlan') {
      const lines = Array.isArray(value) ? value.map(planLine).filter(Boolean) : [];
      if (!Array.isArray(value)) refused.push('approvedPlan was not a list, so the deployed plan was kept');
      else config.approvedPlan = lines;
      continue;
    }

    if (LOWER_ONLY.indexOf(key) !== -1 || RAISE_ONLY.indexOf(key) !== -1) {
      const wanted = Number(value);
      const built = Number(base && base[key]) || 0;
      if (!Number.isFinite(wanted) || wanted < 1) {
        refused.push(key + ' was not a number, so the built value was kept');
        continue;
      }
      const lower = LOWER_ONLY.indexOf(key) !== -1;
      if (built > 0 && lower && wanted > built) {
        refused.push(key + ' can be lowered but not raised past ' + built);
        config[key] = built;
      } else if (built > 0 && !lower && wanted < built) {
        refused.push(key + ' can be raised but not lowered below ' + built);
        config[key] = built;
      } else {
        config[key] = Math.floor(wanted);
      }
      continue;
    }

    if (key === 'assumeStatus') {
      // The status to give a document that does not say what it is. Empty is
      // the strictest answer available - it refuses to name anything whose
      // status is not written down - so it is always allowed. Anything else has
      // to be a status the convention recognises, or the run would build names
      // that its own checker then throws away.
      const wanted = typeof value === 'string' ? value.trim() : null;
      if (wanted === '') {
        config.assumeStatus = '';
      } else if (wanted !== null && STATUSES.indexOf(wanted) !== -1) {
        config.assumeStatus = wanted;
      } else {
        refused.push('assumeStatus must be one of ' + STATUSES.join(', ') + ', or empty to refuse to guess');
      }
      continue;
    }

    if (key === 'sitePath' || key === 'libraryTitle') {
      const picked = chooseTarget(base, key, value);
      if ('why' in picked) refused.push(picked.why);
      else config[key] = picked.value;
      continue;
    }

    // revert: a boolean, and it arms nothing on its own.
    config[key] = value === true;
  }

  return { config, refused };
}

/** A path shown to a person is shown without the part every path shares.
 * @param {string} path @param {string} library */
function shortPath(path, library) {
  const p = String(path || '');
  const root = String(library || '');
  if (root && p.toLowerCase().indexOf(root.toLowerCase() + '/') === 0) return p.slice(root.length + 1);
  return p;
}

/**
 * A run's proposals, turned into a list of things to tick.
 *
 * The label carries its own index because ticks come back as labels, and two
 * documents of the same name in different folders would otherwise offer the
 * same one. The index makes every label distinct whatever the library looks
 * like, and it makes the list quick to read down.
 *
 * Folders come first: a document cannot be filed into a folder nobody has
 * agreed to create, so the thing it depends on is the thing above it.
 *
 * @param {Record<string, any>} summary the Summary node's output
 * @returns {Offered[]}
 */
export function offeredChanges(summary) {
  const s = summary || {};
  const library = String(s.library || '');
  /** @type {Offered[]} */
  const offered = [];
  const created = new Set();

  for (const folder of s.newFolders || []) {
    const path = typeof folder === 'string' ? folder : String((folder && folder.path) || '');
    if (!path) continue;
    created.add(path);
    offered.push({
      index: offered.length + 1,
      label: offered.length + 1 + '. New folder · ' + shortPath(path, library),
      kind: 'folder',
      into: path,
    });
  }

  for (const rename of s.folderRenames || []) {
    const from = String((rename && rename.from) || '');
    const to = String((rename && rename.to) || '');
    if (!from || !to) continue;
    offered.push({
      index: offered.length + 1,
      label: offered.length + 1 + '. Rename folder · ' + shortPath(from, library) + ' → ' + to,
      kind: 'folderRename',
      from: from,
      to: to,
    });
  }

  // A document wanting both a name and a folder is ONE tick, because it is one
  // call to SharePoint and because agreeing to half of it is not a thing anyone
  // means. The apply side already treats it as one change; the page should not
  // pretend otherwise.
  const moves = new Map();
  for (const move of s.toMove || []) {
    if (move && move.file) moves.set(String(move.file), move);
  }

  for (const line of s.toRename || []) {
    const file = String((line && line.file) || '');
    if (!file) continue;
    const move = moves.get(file);
    moves.delete(file);
    const into = move ? String(move.into || '') : '';
    const index = offered.length + 1;
    const label = into
      ? index + '. Rename and file · ' + shortPath(file, library) + ' → ' + line.becomes + ' into ' + shortPath(into, library)
      : index + '. Rename · ' + shortPath(file, library) + ' → ' + line.becomes;
    /** @type {Offered} */
    const entry = {
      index: index,
      label: label,
      kind: into ? 'rename+move' : 'rename',
      from: file,
      to: String(line.becomes || ''),
      createsFolder: into ? created.has(into) : false,
    };
    if (into) entry.into = into;
    offered.push(entry);
  }

  for (const move of moves.values()) {
    const file = String((move && move.file) || '');
    const into = String((move && move.into) || '');
    if (!file || !into) continue;
    const index = offered.length + 1;
    offered.push({
      index: index,
      label: index + '. File · ' + shortPath(file, library) + ' → ' + shortPath(into, library),
      kind: 'move',
      from: file,
      into: into,
      createsFolder: created.has(into),
    });
  }

  return offered;
}

/** Ticks arrive as a list or as one comma-separated string depending on the field.
 * @param {any} value @returns {string[]} */
function ticked(value) {
  if (Array.isArray(value)) return value.map((v) => String(v));
  if (typeof value === 'string') return value.trim() ? value.split(',').map((v) => v.trim()) : [];
  return [];
}

/**
 * The ticks, turned back into a plan and the switches that carry it out.
 *
 * The person never sees a switch. They tick changes, and which switches that
 * needs is arithmetic: a rename needs applyRenames, a filing needs applyMoves,
 * and a filing into a folder that does not exist yet needs applyStructure as
 * well — otherwise the move is refused for landing nowhere, which would read as
 * the tool ignoring them.
 *
 * `applyProposed` stays off on purpose. It means "apply whatever this run
 * proposes", and this run is the second one — its model call may well have
 * come up with different names. The plan built here is what binds.
 *
 * @param {any} answer what the checkbox field came back with
 * @param {Offered[]} offered what reviewFields put on the page
 * @returns {{ approvedPlan: {from:string,to?:string,into?:string}[], overrides: Record<string, any>, chosen: Offered[], unknown: string[] }}
 */
export function approvedFrom(answer, offered) {
  const list = Array.isArray(offered) ? offered : [];
  const byIndex = new Map(list.map((o) => [o.index, o]));
  const byLabel = new Map(list.map((o) => [o.label, o]));

  /** @type {Offered[]} */
  const chosen = [];
  const unknown = [];
  const seen = new Set();

  for (const tick of ticked(answer)) {
    const leading = /^\s*(\d+)\s*\./.exec(tick);
    const found = (leading && byIndex.get(Number(leading[1]))) || byLabel.get(tick);
    if (!found) {
      unknown.push(tick);
      continue;
    }
    if (seen.has(found.index)) continue;
    seen.add(found.index);
    chosen.push(found);
  }

  /** @type {{from: string, to?: string, into?: string}[]} */
  const plan = [];
  /** @type {Record<string, any>} */
  const overrides = {
    applyRenames: false,
    applyMoves: false,
    applyStructure: false,
    applyFolderRenames: false,
    applyProposed: false,
  };

  /** What will actually be attempted, which is not always everything matched. */
  /** @type {Offered[]} */
  const kept = [];

  for (const change of chosen) {
    if (change.kind === 'folder') {
      overrides.applyStructure = true;
      kept.push(change);
      continue;
    }
    if (change.kind === 'folderRename') {
      overrides.applyFolderRenames = true;
      kept.push(change);
      continue;
    }
    // Matched, but asks for nothing: no path, or no name and no destination.
    // Counting it as chosen would report back a number the run cannot deliver,
    // which matters most when the caller is a browser rather than a tick list.
    if (!change.from || (!change.to && !change.into)) {
      unknown.push(change.label || String(change.index || 'an entry with nothing to do'));
      continue;
    }
    /** @type {{from: string, to?: string, into?: string}} */
    const line = { from: change.from };
    if (change.to) {
      line.to = change.to;
      overrides.applyRenames = true;
    }
    if (change.into) {
      line.into = change.into;
      overrides.applyMoves = true;
      // Filing into a folder that is not there yet only works if this same run
      // is allowed to make it. Ticking the move is taken as agreeing to that.
      if (change.createsFolder) overrides.applyStructure = true;
    }
    plan.push(line);
    kept.push(change);
  }

  overrides.approvedPlan = plan;
  return { approvedPlan: plan, overrides, chosen: kept, unknown };
}
