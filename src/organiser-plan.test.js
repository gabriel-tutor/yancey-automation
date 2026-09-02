import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { planRenames } from './filenames.js';
import { categoriseAll, planPlacements } from './library.js';

/**
 * The review page is what Yancey actually reads, and two things can go wrong
 * with it silently.
 *
 * It can go stale — the data block is generated, so an edit to the rules that
 * is never rebuilt leaves the page confidently describing a plan the workflow
 * no longer has. And it can go blank — the page renders itself from that data
 * in the browser, so one typo produces a beautifully styled empty document with
 * no error anywhere a person would look.
 *
 * These run the page's own script against a minimal DOM and hold its data to
 * the same functions the workflow runs.
 */
const html = readFileSync(new URL('../docs/organiser-plan.html', import.meta.url), 'utf8');

/** @returns {any} */
function pageData() {
  const block = html.match(/<script type="application\/json" id="data">([\s\S]*?)<\/script>/);
  if (!block || block[1] === undefined) throw new Error('the page has no data block');
  return JSON.parse(block[1]);
}

/** @type {Array<{ name: string, path: string, modified: string }>} */
const listing = JSON.parse(
  readFileSync(new URL('../fixtures/sharepoint/task-management-listing.json', import.meta.url), 'utf8'),
);
const files = listing.filter((f) => /\.[A-Za-z0-9]*[A-Za-z][A-Za-z0-9]*$/.test(String(f.name)));

describe('the review page agrees with the workflow', () => {
  const data = pageData();

  it('reviews the same files the organiser would', () => {
    expect(data.scanned).toBe(files.length);
  });

  it('carries the same rename verdicts', () => {
    const fresh = planRenames(files, { assumeStatus: 'Draft', assumeVersion: { major: 0, minor: 1 } });
    const onPage = new Map(data.rows.map((/** @type {any} */ r) => [r.current, r]));
    for (const row of fresh) {
      const shown = onPage.get(row.current);
      expect(shown, 'missing from the page: ' + row.current).toBeDefined();
      expect(shown.proposed).toBe(row.proposed);
      expect(shown.blocked).toBe(row.blocked);
    }
  });

  it('carries the same categories and placement verdicts', () => {
    const onPage = new Map(data.rows.map((/** @type {any} */ r) => [r.current, r]));
    for (const c of categoriseAll(files)) {
      expect(onPage.get(c.name).kind).toBe(c.category);
    }
    for (const p of planPlacements(files)) {
      expect(onPage.get(p.name).placement).toBe(p.verdict);
    }
  });

  // A record claiming a rename that is not in the library would be worse than
  // no record: it is the undo list.
  it('only claims renames whose result is really in the library', () => {
    const present = new Set(files.map((f) => f.name));
    for (const applied of data.applied) {
      expect(present.has(applied.to), applied.to + ' is claimed as renamed but is not there').toBe(true);
      expect(applied.from).not.toBe(applied.to);
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
    expect(box('stats').children.length).toBe(5);
    expect(box('tree').children.length).toBe(data.folders.length);
    expect(box('applied').children.length).toBe(data.applied.length);
    expect(box('unfiled').children.length).toBeGreaterThan(0);
    expect(box('settled').children.length).toBeGreaterThan(0);
    expect(box('refused').children.length).toBeGreaterThan(0);
    expect(box('stamp').textContent).toContain('renamed');
  });
});
