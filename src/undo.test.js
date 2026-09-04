import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { nameChanges, reconcile, renameChains } from './undo.js';

/**
 * The undo list, held against the library it claims to describe.
 *
 * config/deploy.json calls this file "the undo list", and a client asking *can
 * you put it back* is asking about exactly it. On 2026-09-04 it could not have
 * answered him: eight name changes had been applied to Yancey's live documents
 * and never written down, so a replay would have restored nineteen chains of
 * twenty-seven. Nothing noticed for a day, because nothing compared the list to
 * the library — the list was only ever appended to.
 *
 * Every one of these fails if that happens again.
 */
const undo = JSON.parse(readFileSync(new URL('../fixtures/sharepoint/renames-applied.json', import.meta.url), 'utf8'));
const library = JSON.parse(readFileSync(new URL('../fixtures/sharepoint/library-2026-09-04.json', import.meta.url), 'utf8'));
const plan = JSON.parse(readFileSync(new URL('../config/approved-plan.json', import.meta.url), 'utf8'));
const names = library.documents.map((/** @type {any} */ d) => String(d.name));

describe('following what was renamed into what', () => {
  it('reads a plain rename as one chain', () => {
    const chains = renameChains([{ kind: 'rename', from: 'a.xlsx', to: 'b.xlsx' }]);
    expect(chains).toEqual([{ original: 'a.xlsx', current: 'b.xlsx', hops: 1 }]);
  });

  /**
   * A file renamed twice is one document with one history. Reading it as two
   * chains shows it twice on the review, under two different original names.
   */
  it('collapses a rename of a rename into one chain', () => {
    const chains = renameChains([
      { kind: 'rename', from: 'a.xlsx', to: 'b.xlsx' },
      { kind: 'rename', from: 'b.xlsx', to: 'c.xlsx' },
    ]);
    expect(chains).toEqual([{ original: 'a.xlsx', current: 'c.xlsx', hops: 2 }]);
  });

  /**
   * A move carries the same name on both sides and the folders separately.
   * Reading `from` and `to` as a rename made the first version of this walk
   * loop forever on the two Performance Management files.
   */
  it('is not fooled by a move, which has the same name on both sides', () => {
    const entries = [
      { kind: 'rename', from: 'a.xlsx', to: 'b.xlsx' },
      { kind: 'move', from: 'b.xlsx', to: 'b.xlsx', fromFolder: 'One', toFolder: 'Two' },
    ];
    expect(nameChanges(entries)).toHaveLength(1);
    expect(renameChains(entries)).toEqual([{ original: 'a.xlsx', current: 'b.xlsx', hops: 1 }]);
  });

  it('ignores a folder creation, which renames nothing', () => {
    expect(nameChanges([{ kind: 'create folder', from: '', to: '/sites/x/Dashboards' }])).toEqual([]);
  });

  /** A cycle in a log of applied operations is a corrupt log. Returning what
   * was reachable reports that; hanging does not. */
  it('stops on a cycle rather than walking forever', () => {
    const chains = renameChains([
      { kind: 'rename', from: 'a.xlsx', to: 'b.xlsx' },
      { kind: 'rename', from: 'b.xlsx', to: 'a.xlsx' },
    ]);
    expect(chains.length).toBeLessThanOrEqual(2);
  });
});

describe('the undo list against the live library', () => {
  const result = reconcile(undo, names, plan.plan);

  /**
   * A chain ending on a name the library does not have means the document was
   * renamed again by something nobody recorded — the record stops short of
   * where the file actually is, so a replay cannot find it.
   */
  it('has no chain that ends on a document the library does not have', () => {
    expect(result.deadEnds.map((/** @type {any} */ d) => d.original + ' -> ' + d.current)).toEqual([]);
  });

  /**
   * The pinned plan is the list of renames somebody signed off. One that was
   * applied and never written down is a change to a client's documents with no
   * way back, which is the whole failure this file exists to prevent.
   */
  it('records every rename the approved plan asked for and the library shows was applied', () => {
    expect(result.unrecorded.map((/** @type {any} */ u) => u.from)).toEqual([]);
  });

  /**
   * Everything else in the library must be a document the tool never touched.
   * Nothing the plan named may be sitting there untraced.
   */
  it('leaves untraced only documents no plan line ever named', () => {
    const planned = new Set(plan.plan.map((/** @type {any} */ l) => String(l.to)));
    for (const name of result.untraced) {
      expect(planned.has(name), name + ' was planned but has no recorded rename').toBe(false);
    }
  });

  it('traces most of the library back to the name it arrived with', () => {
    expect(result.chains.length).toBeGreaterThanOrEqual(27);
  });

  /**
   * Eight entries were reconstructed after the fact rather than logged when the
   * rename ran. That is worth having and worth saying: an inferred line and a
   * recorded one are not the same evidence, and a reader deserves to know which
   * they are looking at.
   */
  it('says so on any entry that was worked out afterwards rather than logged', () => {
    for (const entry of undo) {
      if (!entry.reconstructed) continue;
      expect(String(entry.evidence || '').length, entry.from + ' is reconstructed but gives no evidence')
        .toBeGreaterThan(40);
    }
  });

  it('never records an entry with no name on one side unless it made a folder', () => {
    for (const entry of undo) {
      if (entry.kind === 'create folder') continue;
      expect(String(entry.from || ''), 'an entry with no from cannot be replayed').not.toBe('');
      expect(String(entry.to || ''), 'an entry with no to cannot be replayed').not.toBe('');
    }
  });
});
