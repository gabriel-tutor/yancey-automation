/**
 * Is the record of what was renamed complete enough to put back?
 *
 * `fixtures/sharepoint/renames-applied.json` is described in config/deploy.json
 * as "the undo list", and a client asking *can you undo it* is asking about
 * exactly this file. On 2026-09-04 it could not have answered him: eight name
 * changes had been applied to the live library and never written down, so a
 * replay would have restored nineteen chains of twenty-seven and left the rest
 * with no way back. Nothing noticed, because nothing compared the list against
 * the library.
 *
 * That comparison is what lives here. It is deliberately arithmetic rather than
 * clever: every document now in the library either traces back to an original
 * name through the record, or is a document the tool never touched.
 */

/**
 * An applied operation. `from` and `to` are NAMES, so a move has the same name
 * on both sides and carries the folders instead; a folder creation has no
 * `from` at all.
 *
 * @typedef {{
 *   kind?: string, from?: string, to?: string,
 *   fromFolder?: string, toFolder?: string,
 *   assumptions?: string[], applied?: string, run?: string,
 *   reconstructed?: boolean, evidence?: string
 * }} Applied
 */

/** Operations that change a name — the only ones a rename chain is built from.
 * @param {Applied[]} entries */
export function nameChanges(entries) {
  return (entries || []).filter(
    (e) => e && e.kind !== 'create folder' && e.from && e.to && e.from !== e.to,
  );
}

/**
 * Original name → the name it ends up under, following renames of renames.
 *
 * A file renamed twice must collapse to one chain or the review shows the same
 * document twice under two histories. The walk stops on a repeat rather than
 * trusting the data: a cycle in a log of applied operations is a corrupt log,
 * and hanging is a worse way to report that than returning what was reachable.
 *
 * @param {Applied[]} entries
 * @returns {Array<{ original: string, current: string, hops: number }>}
 */
export function renameChains(entries) {
  const changes = nameChanges(entries);
  const next = new Map(changes.map((e) => [String(e.from), String(e.to)]));
  const isTarget = new Set(changes.map((e) => String(e.to)));

  /** @type {Array<{ original: string, current: string, hops: number }>} */
  const chains = [];
  for (const start of changes) {
    // Only begin at a name nothing else renamed INTO, or every link in a chain
    // becomes its own chain and one document appears three times.
    if (isTarget.has(String(start.from))) continue;
    const seen = new Set([String(start.from)]);
    let current = String(start.to);
    let hops = 1;
    while (next.has(current) && !seen.has(current)) {
      seen.add(current);
      current = String(next.get(current));
      hops += 1;
    }
    chains.push({ original: String(start.from), current, hops });
  }
  return chains;
}

/**
 * The record held against the library it claims to describe.
 *
 * @param {Applied[]} entries the undo list
 * @param {string[]} library every document name now in the library
 * @param {Array<{ from: string, to?: string }>} [plan] the approved plan, if one is pinned
 */
export function reconcile(entries, library, plan) {
  const live = new Set((library || []).map((n) => String(n)));
  const chains = renameChains(entries);

  // A chain that ends on a name the library does not have describes a document
  // that has since been renamed again by something that went unrecorded.
  const deadEnds = chains.filter((c) => !live.has(c.current));

  // A document with no chain was either never touched or was renamed without
  // being written down. The plan is what tells those two apart.
  const landed = new Set(chains.map((c) => c.current));
  const untraced = [...live].filter((name) => !landed.has(name));

  // A plan line whose `from` is gone from the library was applied. If no
  // recorded change starts there, it was applied and never written down.
  const changes = nameChanges(entries);
  const startsAt = new Set(changes.map((e) => String(e.from)));
  const unrecorded = (plan || [])
    .filter((line) => line && line.to && !live.has(String(line.from)) && !startsAt.has(String(line.from)))
    .map((line) => ({ from: String(line.from), pinnedTo: String(line.to), pinnedIsLive: live.has(String(line.to)) }));

  return { chains, deadEnds, untraced, unrecorded };
}
