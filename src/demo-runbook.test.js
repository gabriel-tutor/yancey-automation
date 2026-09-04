import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { reconcile } from './undo.js';

/**
 * The runbook, as it is actually shipped.
 *
 * This page is read aloud in front of the client, which makes it the one
 * document here where being out of date is worse than being missing. It was:
 * on 2026-09-04 it still promised fifteen renames that had been applied the day
 * before, against a library that was by then entirely compliant. The promise
 * came from `config/approved-plan.json`, which records what somebody approved
 * and can say nothing about whether it has happened yet.
 *
 * So these read the built page rather than the builder, and check it against
 * the measured library.
 */
const html = readFileSync(new URL('../docs/demo-runbook.html', import.meta.url), 'utf8');
const library = JSON.parse(readFileSync(new URL('../fixtures/sharepoint/library-2026-09-04.json', import.meta.url), 'utf8'));
const undo = JSON.parse(readFileSync(new URL('../fixtures/sharepoint/renames-applied.json', import.meta.url), 'utf8'));
const plan = JSON.parse(readFileSync(new URL('../config/approved-plan.json', import.meta.url), 'utf8'));

const OPEN = '<script type="application/json" id="data">';
const start = html.indexOf(OPEN);
const data = JSON.parse(html.slice(start + OPEN.length, html.indexOf('</script>', start)));
const live = new Set(library.documents.map((/** @type {any} */ d) => String(d.name)));

describe('the page that gets read out loud', () => {
  it('has been built, not left with an empty data block', () => {
    expect(data.scanned, 'run: node scripts/build-demo-runbook.mjs').toBeGreaterThan(0);
  });

  it('counts the library the measurement counted', () => {
    expect(data.scanned).toBe(library.documents.length);
    expect(data.compliant).toBe(library.documents.filter((/** @type {any} */ d) => d.compliant).length);
    expect(data.folders).toBe(library.folders.length);
  });

  /**
   * The failure this file exists for. A rename is pending only while the
   * library still holds the old name; once the new one is there it has
   * happened, and a page saying otherwise is read out as a promise.
   */
  it('never promises a rename that has already happened', () => {
    for (const line of data.pending) {
      expect(live.has(String(line.to)), line.to + ' is already in the library but the page calls it pending')
        .toBe(false);
      expect(live.has(String(line.from)), line.from + ' is not in the library, so it cannot be renamed')
        .toBe(true);
    }
  });

  it('shows only history that lands on a document the library has', () => {
    for (const entry of data.history) {
      expect(live.has(String(entry.current)), entry.current + ' is on the page but not in the library').toBe(true);
    }
  });

  /** The claim the page makes about undo has to be the one the record supports. */
  it('claims a traced history no larger than the record actually reconciles', () => {
    const record = reconcile(undo, [...live], plan.plan);
    expect(data.traced).toBe(record.chains.filter((/** @type {any} */ c) => live.has(c.current)).length);
    expect(record.deadEnds).toEqual([]);
    expect(record.unrecorded).toEqual([]);
  });

  it('marks the entries that were reconstructed rather than logged', () => {
    const marked = data.history.filter((/** @type {any} */ h) => h.reconstructed).length;
    expect(marked).toBe(data.reconstructed);
    expect(data.reconstructed, 'if this is zero the marking has quietly been lost').toBeGreaterThan(0);
  });

  /**
   * Two of the three bugs the web page produced were functions that were right
   * with nothing connecting them. The same test earns its place here.
   */
  it('reaches only for elements the markup actually has', () => {
    const wanted = [...html.matchAll(/getElementById\('([^']+)'\)/g)].map((m) => m[1]);
    expect(wanted.length).toBeGreaterThan(5);
    for (const id of wanted) {
      expect(html.includes('id="' + id + '"'), 'the script reaches for #' + id + ', which is not in the page')
        .toBe(true);
    }
  });

  it('says out loud that nothing is applied without a tick', () => {
    expect(html).toContain('Nothing is applied without a tick');
  });

  /**
   * Four of the renamed documents are Power BI reports, referenced by path from
   * outside the library. The page must not let somebody rename them back and
   * forward for the sake of a livelier demo.
   */
  it('warns against reverting rather than offering it as the easy option', () => {
    expect(html).toContain('do not do this');
    expect(html).toContain('Add new documents instead');
  });
});
