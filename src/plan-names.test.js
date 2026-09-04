import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';

/**
 * Every file name written into a document, held to the plan that produced it.
 *
 * This exists because a plan was retyped by hand and nine of twelve lines
 * matched. The three that did not were macOS screenshots, which carry U+202F —
 * a narrow no-break space — before AM/PM. Retyped with an ordinary space they
 * look identical and match nothing.
 *
 * The workflow now reports a line whose file it cannot find, and this stops the
 * same slip reaching a page a client reads: a name quoted in prose has to be
 * one that actually appears in the approved plan.
 */
const plan = JSON.parse(readFileSync(new URL('../config/approved-plan.json', import.meta.url), 'utf8'));

/** Both sides of every approved line, plus what the library already holds. */
const known = new Set();
for (const line of plan.plan) {
  known.add(String(line.from));
  if (line.to) known.add(String(line.to));
}
/** A line approves a name, a destination, or both. These carry a name. */
const renames = plan.plan.filter((/** @type {any} */ l) => l.to);
for (const file of JSON.parse(
  readFileSync(new URL('../fixtures/sharepoint/task-management-listing.json', import.meta.url), 'utf8'),
)) {
  known.add(String(file.name));
}
for (const applied of JSON.parse(
  readFileSync(new URL('../fixtures/sharepoint/renames-applied.json', import.meta.url), 'utf8'),
)) {
  known.add(String(applied.from));
  known.add(String(applied.to));
}
// The runbook is built from the library as it was measured, so the names it
// quotes are the ones actually sitting there — which is the point of it.
for (const doc of JSON.parse(
  readFileSync(new URL('../fixtures/sharepoint/library-2026-09-04.json', import.meta.url), 'utf8'),
).documents) {
  known.add(String(doc.name));
}

/** A name in the convention's shape, which is what gets quoted in prose. */
const QUOTED = /[A-Za-z0-9][A-Za-z0-9._-]*_\d{4}-\d{2}-\d{2}_(?:Draft|Review|Final|Archived)_v\d+\.\d+\.[A-Za-z0-9]+/g;

const DOCS = [
  '../docs/organiser-explainer.html',
  '../docs/reply-organiser-agent.md',
  '../docs/DEPLOY.md',
  '../docs/demo-runbook.html',
];

describe('file names quoted in documents are real', () => {
  it.each(DOCS)('%s quotes only names from the plan', (doc) => {
    const text = readFileSync(new URL(doc, import.meta.url), 'utf8');
    const strays = [...new Set(text.match(QUOTED) || [])].filter((name) => !known.has(name));
    expect(strays, 'not in the approved plan or the library').toEqual([]);
  });
});

/**
 * The runbook is read aloud in front of the client, so a name in it that the
 * workflow is not about to produce is worse than a name in a document.
 *
 * It used to be generated from the pinned plan and held to that, which was
 * right while the plan was pending and became wrong the moment it was applied:
 * on 2026-09-04 the page still promised fifteen renames that had happened the
 * day before. A plan records what somebody approved and can say nothing about
 * whether it has happened yet, so the page — and this — now go by the library.
 */
describe('the demo runbook matches the library', () => {
  const html = readFileSync(new URL('../docs/demo-runbook.html', import.meta.url), 'utf8');
  const block = html.match(/<script type="application\/json" id="data">([\s\S]*?)<\/script>/);
  const data = JSON.parse(block?.[1] ?? '{}');
  const live = new Set(
    JSON.parse(readFileSync(new URL('../fixtures/sharepoint/library-2026-09-04.json', import.meta.url), 'utf8'))
      .documents.map((/** @type {any} */ d) => String(d.name)),
  );

  it('calls pending exactly the approved renames the library has not had yet', () => {
    expect(data.pending.map((/** @type {any} */ p) => p.from)).toEqual(
      renames.filter((/** @type {any} */ l) => live.has(String(l.from))).map((/** @type {any} */ p) => p.from),
    );
  });

  it('reverses every traced rename for the undo, and none is a no-op', () => {
    expect(data.reverse.length).toBe(data.history.length);
    for (let i = 0; i < data.reverse.length; i += 1) {
      expect(data.reverse[i].from).toBe(data.history[i].current);
      expect(data.reverse[i].to).toBe(data.history[i].original);
      expect(data.reverse[i].from).not.toBe(data.reverse[i].to);
    }
  });

  /**
   * A folder the model proposed is on the page because it is what makes the
   * structure appear on the day: the model's grouping is not repeatable, and a
   * folder that was actually created is.
   */
  it('lists the folders it actually made, by a path inside the library', () => {
    expect(data.foldersMade.length).toBeGreaterThan(0);
    for (const folder of data.foldersMade) {
      expect(String(folder)).not.toContain('/sites/');
      expect(String(folder).length).toBeGreaterThan(3);
    }
  });

  /**
   * The old version of this asserted there was always something left to rename,
   * which stopped being true and made the failure look like a broken test
   * rather than a page that had gone out of date. What matters is not that
   * there is material — it is that the page says which of the two demos it is.
   */
  it('says which demo it is, rather than assuming there is material', () => {
    const nothingLeft = data.pending.length === 0;
    expect(nothingLeft).toBe(data.compliant === data.scanned);
    if (nothingLeft) {
      expect(html, 'with nothing to offer, the page has to say how to create some')
        .toContain('If you want it to do something live');
    }
  });
});

/**
 * The plan as the workflow receives it, not as the file holds it.
 *
 * `String(undefined)` is the string "undefined", and a line approving only a
 * destination has no name — so the naive mapping sent "undefined" into the
 * Config node as an approved file name. It would have been refused rather than
 * applied, which is the safe direction, but it would also have silently taken
 * the two documents out of the demo.
 */
describe('the plan the generator hands to the workflow', () => {
  it('never turns a missing half of a line into a name', async () => {
    const { CONFIG } = await import('../workflows/wf-f-organizer.mjs');
    for (const line of CONFIG.approvedPlan) {
      expect(line.to, line.from + ' approves a name that is not one').not.toBe('undefined');
      expect(line.into, line.from + ' is filed nowhere real').not.toBe('undefined');
      expect(Boolean(line.to || line.into), line.from + ' approves nothing at all').toBe(true);
    }
  });

  /**
   * A destination is pinned whether or not the same line also pins a name — a
   * line that started as filing-only gains a `to` the moment that document's
   * name is approved too, and it is still the thing that puts the folder beyond
   * the model's whim.
   */
  it('pins at least one destination, whether or not it also pins a name', () => {
    const filed = plan.plan.filter((/** @type {any} */ l) => l.into);
    expect(filed.length, 'the folder structure needs at least one pinned destination').toBeGreaterThan(0);
  });
});
