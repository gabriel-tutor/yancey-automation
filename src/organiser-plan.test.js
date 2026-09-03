import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { diagnose } from './filenames.js';

/**
 * The review page is what Yancey actually reads, and two things can go wrong
 * with it silently.
 *
 * It can go stale — the data block is generated, so a run that is never rebuilt
 * into the page leaves it confidently describing a plan the workflow no longer
 * has. And it can go blank — the page renders itself from that data in the
 * browser, so one typo produces a beautifully styled empty document with no
 * error anywhere a person would look.
 *
 * The authority used to be a local rerun of planRenames, which was right while
 * the names came from rules. It is not any more: the names come from a model
 * reading the documents, and rerunning anything locally would only prove the
 * page agrees with a computation nobody performed. So the authority is the run
 * itself, saved whole, and what is checked is that the page carries that run
 * faithfully and that every name on it still satisfies the convention.
 */
const html = readFileSync(new URL('../docs/organiser-plan.html', import.meta.url), 'utf8');

/** @returns {any} */
function pageData() {
  const block = html.match(/<script type="application\/json" id="data">([\s\S]*?)<\/script>/);
  if (!block || block[1] === undefined) throw new Error('the page has no data block');
  return JSON.parse(block[1]);
}

/** @type {{ executionId: string, summary: any, rows: any[] }} */
const run = JSON.parse(
  readFileSync(new URL('../fixtures/sharepoint/organiser-run.json', import.meta.url), 'utf8'),
);

describe('the review page agrees with the run', () => {
  const data = pageData();

  it('reviews the same files the run did', () => {
    expect(data.scanned).toBe(run.rows.length);
    expect(data.execution).toBe(run.executionId);
  });

  it('carries the run’s verdicts unchanged', () => {
    const onPage = new Map(data.rows.map((/** @type {any} */ r) => [r.current, r]));
    for (const row of run.rows) {
      const shown = onPage.get(row.current);
      expect(shown, 'missing from the page: ' + row.current).toBeDefined();
      expect(shown.proposed).toBe(row.proposed);
      expect(shown.blocked).toBe(row.blocked);
      expect(shown.kind).toBe(row.category || '');
      expect(shown.placement).toBe(row.placement || '');
    }
  });

  /**
   * The workflow validates every proposal before it leaves the node, so this is
   * the second lock rather than the first — but the page is what a person reads
   * and then approves, and a name on it that does not satisfy his own rules
   * would be approved on the page's authority, not the workflow's.
   */
  it('shows no proposed name that breaks the convention', () => {
    for (const row of data.rows) {
      if (!row.proposed) continue;
      expect(diagnose(row.proposed), row.current + ' -> ' + row.proposed).toEqual([]);
    }
  });

  it('agrees with itself about how many were named by reading', () => {
    const named = data.rows.filter((/** @type {any} */ r) => r.namedBy === 'model');
    expect(named.length).toBe(data.namedByModel);
    for (const row of named) {
      expect(row.why, row.current + ' was named by the model with no reason given').toBeTruthy();
    }
  });

  /**
   * The point of the rewrite, checked against what actually happened rather
   * than what is still outstanding. A regex can rearrange the characters
   * already in a name; it cannot write "Sales-Dashboard-Overview" for a file
   * called "Screenshot 2025-04-06 at 6.34.18 PM.png", because those words are
   * in the picture.
   *
   * Every approved name has to be either applied or still proposed. A plan that
   * quietly went nowhere would otherwise leave a page that looks finished.
   */
  it('carries out every line of the approved plan', () => {
    const plan = JSON.parse(
      readFileSync(new URL('../config/approved-plan.json', import.meta.url), 'utf8'),
    );
    // Carried out means the library holds the approved name — not merely that
    // this repository's record happens to mention it. Earlier runs applied
    // names without every one reaching renames-applied.json, and a check that
    // could only see the record would call those lines adrift when the file is
    // sitting there under exactly the name that was approved.
    const landed = new Set([
      ...run.rows.map((/** @type {any} */ r) => r.current),
      ...data.applied.map((/** @type {any} */ a) => a.to),
      ...data.rows.map((/** @type {any} */ r) => r.proposed).filter(Boolean),
    ]);
    // A line approves a name, a destination, or both. Only the ones carrying a
    // name land as a rename; a filing line moves a document and leaves its name
    // exactly where it was.
    const adrift = plan.plan
      .filter((/** @type {any} */ l) => l.to)
      .map((/** @type {any} */ l) => l.to)
      .filter((/** @type {string} */ to) => !landed.has(to));
    expect(adrift, 'approved but neither applied nor proposed').toEqual([]);
  });

  // A record claiming a change that is not in the library would be worse than
  // no record: it is the undo list.
  it('only claims changes whose result is really in the library', () => {
    const present = new Set(run.rows.map((/** @type {any} */ r) => r.current));
    // A folder is not a document, and the page now records folders it made.
    const folders = new Set(data.folders.map((/** @type {any} */ f) => f.path));
    for (const applied of data.applied) {
      const where = String(applied.kind).indexOf('folder') === -1 ? present : folders;
      expect(where.has(applied.to), applied.to + ' is claimed as changed but is not there').toBe(true);
    }
  });

  /**
   * The two kinds change different things, and a record that conflated them
   * would show a move as a file renamed to itself.
   */
  it('records a rename as a changed name and a move as a changed folder', () => {
    for (const applied of data.applied) {
      if (applied.kind === 'move') {
        expect(applied.from, 'a move keeps the name').toBe(applied.to);
        expect(applied.fromFolder, 'a move must record where it came from').toBeTruthy();
        expect(applied.fromFolder).not.toBe(applied.folder);
      } else {
        expect(applied.from, 'a rename must change the name').not.toBe(applied.to);
      }
    }
  });
});

describe('the review page renders', () => {
  it('builds every section without throwing', () => {
    const script = html.match(/<script>\n([\s\S]*?)<\/script>\s*$/);
    if (!script || script[1] === undefined) throw new Error('the page has no render script');

    const ids = [...html.matchAll(/id="([a-zA-Z0-9_-]+)"/g)].map((m) => m[1]).filter(
      /** @returns {value is string} */ (value) => value !== undefined,
    );
    /** @param {string} tag */
    const node = (tag) => ({
      tagName: tag,
      className: '',
      textContent: '',
      title: '',
      hidden: false,
      style: /** @type {Record<string, string>} */ ({}),
      children: /** @type {any[]} */ ([]),
      /** @param {any} child */
      appendChild(child) {
        this.children.push(child);
        return child;
      },
    });

    /** @type {Record<string, ReturnType<typeof node>>} */
    const byId = {};
    for (const id of ids) byId[id] = node('div');
    /** @param {string} id */
    const box = (id) => {
      const found = byId[id];
      if (!found) throw new Error('the page has no #' + id);
      return found;
    };

    const dataBlock = html.match(/<script type="application\/json" id="data">([\s\S]*?)<\/script>/);
    box('data').textContent = dataBlock?.[1] ?? '{}';

    const missing = /** @type {string[]} */ ([]);
    // @ts-expect-error - a deliberately minimal stand-in for the browser
    globalThis.document = {
      /** @param {string} id */
      getElementById(id) {
        if (!(id in byId)) {
          missing.push(id);
          return node('div');
        }
        return byId[id];
      },
      /** @param {string} tag */
      createElement: (tag) => node(tag),
    };

    // Executing the page's own script is the point of this test. The source is
    // docs/organiser-plan.html from this repository, nothing is interpolated
    // into it, and it runs against the stand-in above rather than a real DOM.
    new Function(script[1])();

    expect(missing, 'the script asked for elements the page does not have').toEqual([]);

    const data = pageData();
    // Six figures, and the two that matter most are the ones the 3 September run
    // could not tell apart: what is approved and what is only proposed.
    expect(box('stats').children.length).toBe(6);
    const statLabels = box('stats').children.map((/** @type {any} */ c) =>
      c.children.map((/** @type {any} */ g) => g.textContent).join(' '),
    );
    expect(statLabels.join(' | ')).toContain('approved and still to rename');
    expect(statLabels.join(' | ')).toContain('waiting for you');
    expect(box('tree').children.length).toBe(data.folders.length);
    expect(box('applied').children.length).toBe(data.applied.length);
    expect(box('unfiled').children.length).toBeGreaterThan(0);
    expect(box('settled').children.length).toBeGreaterThan(0);
    // Refusals are a finding, not a fixture. Once the reports were opened and
    // none of them named a document in this library, there were none left — so
    // the check is that the page agrees with the run, either way.
    const refusedRows = data.rows.filter((/** @type {any} */ r) => r.blocked && !r.compliant);
    if (refusedRows.length === 0) {
      expect(box('refused').children.length).toBe(0);
    } else {
      expect(box('refused').children.length).toBeGreaterThan(0);
    }
    expect(box('stamp').textContent).toContain('renamed');
  });
});
