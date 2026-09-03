/**
 * Turning a reviewed plan into calls SharePoint will actually accept.
 *
 * This module exists because of two failures on 3 September, both of them in
 * the same place: the step that decides what to send.
 *
 * **A file needed two conflicting calls.** Renaming and moving were built as
 * separate action lists, so a document that wanted a new name *and* a new
 * folder produced two calls, both starting from the path it had before. The
 * first succeeded and the second could not: the file was no longer there. To
 * SharePoint a rename and a move are one operation — both change a path — so a
 * file gets exactly one call, to its new folder under its new name.
 *
 * **The URL ran out of room.** `moveto` carries the path twice, and ASP.NET
 * caps the URL *path* at 260 characters. Yancey created a folder called
 * "PD alignment Excel and Power BI - Karen", moved two files into it, and the
 * rename came back with *"The length of the URL for this request exceeds the
 * configured maxUrlLength"* at about 385 characters. The paths now travel in
 * the query string behind OData parameter aliases, where the ceiling is 2048.
 */
import { folderOf, pathKey } from './library.js';

/** Characters that would end the value early, or mean something else, inside a query string. */
/** @type {Array<[string, string]>} */
const NEEDS_ESCAPING = [
  ["'", "''"], // OData escapes a quote by doubling it; must run before percent-encoding
  ['%', '%25'], // must run before anything that introduces a percent
  ['#', '%23'],
  ['&', '%26'],
  ['+', '%2B'],
  ['?', '%3F'],
  [' ', '%20'],
];

/**
 * A server-relative path, encoded to sit inside an OData string literal that is
 * itself inside a query string.
 *
 * Order matters twice over: the quote is doubled first because that is an OData
 * rule rather than an encoding one, and the percent is encoded before anything
 * that introduces one, or the escapes escape each other.
 *
 * @param {string} path
 * @returns {string}
 */
export function odataQueryValue(path) {
  let out = String(path);
  for (const [from, to] of NEEDS_ESCAPING) out = out.split(from).join(to);
  return out;
}

/**
 * The call that moves a file from one path to another.
 *
 * `flags=0` is the promise never to overwrite: if anything already sits on the
 * new path, SharePoint fails the call rather than replacing what is there.
 *
 * @param {string} base e.g. https://host/sites/Site/_api/web
 * @param {string} oldPath server-relative
 * @param {string} newPath server-relative
 * @returns {string}
 */
export function fileChangeUrl(base, oldPath, newPath) {
  return (
    String(base) +
    "/GetFileByServerRelativePath(decodedurl=@old)/moveto(newurl=@new,flags=0)?@old='" +
    odataQueryValue(oldPath) +
    "'&@new='" +
    odataQueryValue(newPath) +
    "'"
  );
}

/**
 * The call that creates a folder, if it is not already there.
 *
 * @param {string} base
 * @param {string} path server-relative path of the folder to create
 * @returns {string}
 */
export function folderCreateUrl(base, path) {
  return String(base) + "/folders/addUsingPath(DecodedUrl=@p)?@p='" + odataQueryValue(path) + "'";
}

/**
 * The call that renames a folder. A folder is moved exactly as a file is.
 *
 * @param {string} base
 * @param {string} oldPath
 * @param {string} newPath
 * @returns {string}
 */
export function folderChangeUrl(base, oldPath, newPath) {
  return (
    String(base) +
    "/GetFolderByServerRelativePath(decodedurl=@old)/moveto(newurl=@new,flags=0)?@old='" +
    odataQueryValue(oldPath) +
    "'&@new='" +
    odataQueryValue(newPath) +
    "'"
  );
}

/** @param {string} path @returns {string} */
function nameOf(path) {
  const cut = String(path).lastIndexOf('/');
  return cut === -1 ? String(path) : String(path).slice(cut + 1);
}

/**
 * The folder a row would end up in.
 *
 * There is one of these because there was once more than one, and the two
 * disagreed: a document renamed and moved got a call built from its old folder
 * and a call built from its new one. Every question about where something lands
 * — the call, the collision check, the review page — asks this.
 *
 * @param {any} row
 * @returns {string}
 */
export function landingFolder(row) {
  if (row && row.placement === 'move' && row.destination) {
    // A plan line that names a destination overrides what this run worked out.
    return row.approvedInto ? String(row.approvedInto) : folderOf(String(row.destination));
  }
  return folderOf(String((row && row.path) || ''));
}

/**
 * The full path a row would end up at, name and folder together.
 *
 * @param {any} row
 * @returns {string}
 */
export function landingPath(row) {
  const path = String((row && row.path) || '');
  return landingFolder(row) + '/' + String((row && row.proposed) || nameOf(path));
}

/** SharePoint Online refuses a server-relative path longer than this. */
const MAX_PATH = 400;

/**
 * Why a row's landing cannot happen, or an empty string when it can.
 *
 * @param {any} row
 * @param {any[]} rows every row, because two of them can want one path
 * @param {Array<{ name: string, path: string }>} files what the library holds
 * @param {number} limit
 * @returns {string}
 */
function whyItCannotLand(row, rows, files, limit) {
  /** @param {any} r */
  const changing = (r) => landingPath(r) !== String(r.path || '');
  if (!changing(row)) return '';

  const landing = landingPath(row);
  const key = pathKey(landing);

  const together = rows.filter((r) => r && r.path && changing(r) && pathKey(landingPath(r)) === key);
  if (together.length > 1) {
    return (
      together.length +
      ' documents would land on the same path — ' +
      landing +
      ', so they need names or folders that tell them apart'
    );
  }

  // A document that is itself moving away frees the path it is on.
  const leaving = new Set(rows.filter((r) => r && r.path && changing(r)).map((r) => pathKey(r.path)));
  const occupied = (Array.isArray(files) ? files : []).some(
    (f) => pathKey((f && f.path) || '') === key && !leaving.has(key),
  );
  if (occupied && key !== pathKey(row.path)) {
    return 'something already sits at ' + landing + ', and nothing here overwrites what is there';
  }

  if (landing.length > limit) {
    return (
      'that path is ' + landing.length + ' characters, longer than the ' + limit + ' characters SharePoint accepts'
    );
  }
  return '';
}

/**
 * Hold back a change that cannot land, before it is ever sent.
 *
 * `flags=0` means SharePoint refuses rather than overwrites, so a collision
 * costs a failed call rather than a document — but a run that plans a call it
 * knows will fail is not reviewing anything. It is guessing, and finding out in
 * front of whoever is watching.
 *
 * Three ways a change cannot land:
 *
 *   two documents composed onto one path   whichever ran second would fail
 *   something is already sitting there     unless that something is moving away
 *   the path is longer than 400 characters SharePoint's own ceiling
 *
 * It gives up the folder before it gives up the name. A document whose
 * *destination* is taken is a document that stays where it is — it is not a
 * document that should keep being called something wrong, and cancelling both
 * also took the proposal off the waiting list, which is how a good name
 * disappears without anybody deciding to drop it. Only when the name still
 * collides where the document already sits does the name go too.
 *
 * This reads the proposal as proposed, not as gated by the switches, because
 * the review page is what a person signs off and a plan with a collision inside
 * it is a plan to fix rather than one to half-apply.
 *
 * @param {any[]} rows plan rows
 * @param {Array<{ name: string, path: string }>} files what the library holds
 * @param {{ maxPathLength?: number }} [options]
 * @returns {any[]} the same rows
 */
export function refuseImpossible(rows, files, options) {
  const opts = options || {};
  const limit = Number(opts.maxPathLength) > 0 ? Number(opts.maxPathLength) : MAX_PATH;
  const list = Array.isArray(rows) ? rows : [];

  // First give up the folder. Every reason is worked out against the state as
  // it stands before any of them is acted on, because two documents wanting one
  // path is a symmetric problem: cancelling the first would leave the second
  // looking fine, and letting it through would be this picking a winner between
  // two documents, which is exactly the choice that belongs to a person.
  const folderTrouble = list
    .filter((row) => row && row.path && row.placement === 'move' && row.destination)
    .map((row) => ({ row, why: whyItCannotLand(row, list, files, limit) }))
    .filter((found) => found.why);

  for (const { row, why } of folderTrouble) {
    row.placement = 'settled';
    row.destination = '';
    row.moveRefused = why;
    row.placementWhy = 'stays where it is — ' + why;
    // A surviving name is not a blocked document; it is a rename that will
    // happen and a move that will not, and saying otherwise takes a good
    // proposal off the waiting list.
    if (!row.proposed) row.blocked = 'stays where it is — ' + why;
  }

  // Then the name, if it still cannot land where the document already is.
  const nameTrouble = list
    .filter((row) => row && row.path)
    .map((row) => ({ row, why: whyItCannotLand(row, list, files, limit) }))
    .filter((found) => found.why);

  for (const { row, why } of nameTrouble) {
    row.proposed = '';
    row.approved = false;
    row.awaiting = '';
    row.placement = 'settled';
    row.destination = '';
    row.blocked = why;
  }

  return list;
}

/**
 * Changes in an order SharePoint can actually carry out, and the ones it cannot.
 *
 * A document may move onto a path another document is vacating — that is legal
 * and useful, and `refuseImpossible` allows it — but only in that order, and
 * `flags=0` means getting the order wrong is a failed call. Each change can be
 * waiting on at most one other, because at most one document can be leaving any
 * given path, so the dependencies form chains and rings rather than a tangle.
 *
 * A ring is two documents trading places, which cannot be done at all without
 * somewhere to put one of them first. There is no such place here, so both are
 * refused rather than half-attempted.
 *
 * Every change in a chain carries the same `chain`, because a cap that fell
 * between them would send the second half of a swap and none of the first.
 *
 * @param {any[]} changes
 * @returns {{ ordered: any[], impossible: any[] }}
 */
export function orderChanges(changes) {
  const list = Array.isArray(changes) ? changes : [];
  /** @type {Map<string, any>} the change vacating each path */
  const vacating = new Map();
  for (const change of list) vacating.set(pathKey(change.oldPath), change);

  /** @param {any} change @returns {any} the change that has to run first */
  const waitingOn = (change) => {
    const before = vacating.get(pathKey(change.newPath));
    return before && before !== change ? before : null;
  };

  /** @type {any[]} */
  const ordered = [];
  /** @type {any[]} */
  const impossible = [];
  const placed = new Set();
  let chain = 0;

  for (const change of list) {
    if (placed.has(change)) continue;

    // Walk the chain from here to whatever it ends on, or back to itself.
    /** @type {any[]} */
    const run = [];
    const seen = new Set();
    let at = change;
    let ring = false;
    while (at && !placed.has(at)) {
      if (seen.has(at)) {
        ring = true;
        break;
      }
      seen.add(at);
      run.push(at);
      at = waitingOn(at);
    }

    if (ring) {
      for (const member of run) {
        placed.add(member);
        impossible.push(member);
      }
      continue;
    }

    // `run` reads from the change that must go last to the one that goes first.
    chain += 1;
    for (let i = run.length - 1; i >= 0; i -= 1) {
      const member = run[i];
      placed.add(member);
      member.chain = chain;
      ordered.push(member);
    }
  }

  return { ordered, impossible };
}

/**
 * The actions that fit in one run, taken as whole groups.
 *
 * A grouping is one folder plus every document that goes into it. Cutting the
 * list at the cap could fall between them, which creates a folder and leaves
 * its documents outside — a state nobody chose and nobody asked for. A group
 * that does not fit is skipped rather than stopping the run, so a cap set too
 * low still makes progress on everything else, and what was held back is
 * reported rather than silently dropped.
 *
 * A group larger than the whole cap is held on every run for ever, and "no room
 * left this run" reads as though the next run will get to it. It will not, and
 * only a person raising the cap changes that — so `neverFits` says which kind
 * of held-back this is.
 *
 * @param {Array<{ group?: string }>} actions each carrying the group it belongs to
 * @param {number} cap
 * @returns {{ send: any[], held: Array<{ group: string, actions: number, neverFits: boolean }> }}
 */
export function fittingActions(actions, cap) {
  const limit = Number(cap) > 0 ? Number(cap) : 0;
  /** @type {string[]} */
  const order = [];
  /** @type {Map<string, any[]>} */
  const byGroup = new Map();

  for (const action of Array.isArray(actions) ? actions : []) {
    const key = String((action && action.group) || '');
    if (!byGroup.has(key)) {
      byGroup.set(key, []);
      order.push(key);
    }
    (byGroup.get(key) || []).push(action);
  }

  /** @type {any[]} */
  const send = [];
  /** @type {Array<{ group: string, actions: number, neverFits: boolean }>} */
  const held = [];
  let used = 0;

  for (const key of order) {
    const group = byGroup.get(key) || [];
    if (used + group.length <= limit) {
      for (const action of group) send.push(action);
      used += group.length;
    } else {
      held.push({ group: key, actions: group.length, neverFits: group.length > limit });
    }
  }

  return { send, held };
}

/**
 * One change per file, or none.
 *
 * A row can want a new name, a new folder, or both. Both is still one call, and
 * that is the whole point of this function: the destination is composed from
 * whichever parts are actually approved, and a row whose composed destination
 * equals where it already is produces nothing.
 *
 * `requireApproval` is what a pinned plan turns on. Without it, everything this
 * run proposes is fair game — right for a sandbox. With it, only what a person
 * signed off is applied, and the rest stays a proposal on the page rather than
 * disappearing.
 *
 * It binds **names**, because a list of names is what a plan is. It does not
 * bind folders: a plan that never mentions a destination has not refused one,
 * and reading it as a refusal would mean a pinned plan silently forbade every
 * folder the agent proposed — the same mistake as gating proposals, one level
 * up. Where a plan line *does* name a destination in `into`, that is the folder
 * used. Where it does not, the move is governed by `applyMoves`, which is the
 * switch a person arms after reading the page.
 *
 * @param {any[]} rows plan rows
 * @param {{ applyRenames?: boolean, applyMoves?: boolean, requireApproval?: boolean }} [options]
 * @returns {Array<{ kind: string, oldName: string, newName: string, oldPath: string, newPath: string, why: string[] }>}
 */
export function plannedChanges(rows, options) {
  const opts = options || {};
  const changes = [];

  for (const row of Array.isArray(rows) ? rows : []) {
    const path = String(row.path || '');
    if (!path) continue;

    const nameApproved = opts.requireApproval ? row.approved === true : true;

    const wantName = opts.applyRenames === true && Boolean(row.proposed) && nameApproved;
    const wantMove = opts.applyMoves === true && row.placement === 'move' && Boolean(row.destination);
    if (!wantName && !wantMove) continue;

    const folder = wantMove ? landingFolder(row) : folderOf(path);
    const name = wantName ? String(row.proposed) : nameOf(path);
    const newPath = folder + '/' + name;

    // Composing from parts can land exactly where the file already is — a move
    // to its own folder, say. That is not a change and must not be a call.
    if (newPath === path) continue;

    changes.push({
      kind: wantName && wantMove ? 'rename and move' : wantName ? 'rename' : 'move',
      oldName: nameOf(path),
      newName: name,
      oldPath: path,
      newPath,
      why: wantName ? row.assumptions || [] : [String(row.placementWhy || '')],
    });
  }

  return changes;
}

/**
 * Folders the documents would land in, that are not there yet.
 *
 * Every move needs somewhere to land, and this asks the question once for all
 * of them rather than trusting whichever step proposed the move to have also
 * proposed the folder. It was written the narrow way first — only the folders a
 * pinned plan named — and a move to a folder nothing had thought to create went
 * straight through to a call that could only fail.
 *
 * A person who approved filing a document into a folder approved the folder
 * with it, and the same is true of a grouping the agent invented: the folder is
 * not a separate decision, it is the other half of the move.
 *
 * A destination can be several levels deeper than anything that exists, and
 * `addUsingPath` will not create a folder whose parent is missing — so every
 * missing level is named, shallowest first, and creating them in that order is
 * the caller's job.
 *
 * The walk up stops at the first folder known to exist. A destination with no
 * such ancestor is reported as unreachable rather than created: walking further
 * would end at `/sites`, and nothing here has any business creating those.
 * Pass the parent folder of every file alongside the folder list, because a
 * folder holding a document certainly exists whether or not it was listed.
 *
 * A destination it cannot reach is a move that cannot happen, so
 * `cancelUnreachable` cancels it here rather than leaving the page to say it is
 * impossible while the call goes out anyway.
 *
 * @param {any[]} rows plan rows
 * @param {string[]} existing folders known to be there
 * @param {{ cancelUnreachable?: boolean }} [options]
 * @returns {{ create: string[], unreachable: string[] }}
 */
export function foldersNeeded(rows, existing, options) {
  // SharePoint compares paths without case, so a folder that differs only in
  // case is the same folder, and creating it again would fail.
  const there = new Set((Array.isArray(existing) ? existing : []).map(pathKey));
  /** @type {string[]} */
  const create = [];
  /** @type {string[]} */
  const unreachable = [];
  const planned = new Set();

  for (const row of Array.isArray(rows) ? rows : []) {
    if (!row || row.placement !== 'move' || !row.destination) continue;
    const into = landingFolder(row);
    if (!into || there.has(pathKey(into)) || planned.has(pathKey(into))) continue;

    // Walk up to the first level that is already there, collecting what is not.
    /** @type {string[]} */
    const missing = [];
    let at = into;
    while (at && at.indexOf('/') !== -1 && !there.has(pathKey(at)) && !planned.has(pathKey(at))) {
      missing.unshift(at);
      at = folderOf(at);
    }

    if (!at || (!there.has(pathKey(at)) && !planned.has(pathKey(at)))) {
      unreachable.push(into);
      if (options && options.cancelUnreachable) {
        const why =
          'nothing above ' +
          into +
          ' is a folder this run knows exists, and walking further up would leave the library altogether';
        row.placement = 'settled';
        row.destination = '';
        row.moveRefused = why;
        row.placementWhy = 'stays where it is — ' + why;
        if (!row.proposed) row.blocked = 'stays where it is — ' + why;
      }
      continue;
    }

    for (const path of missing) {
      create.push(path);
      planned.add(pathKey(path));
    }
  }

  return { create, unreachable };
}
