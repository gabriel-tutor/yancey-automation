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
 * workflow is not about to produce is worse than a name in a document. It is
 * generated from the pinned plan; this holds it to that.
 */
describe('the demo runbook matches the pinned plan', () => {
  const html = readFileSync(new URL('../docs/demo-runbook.html', import.meta.url), 'utf8');
  const block = html.match(/<script type="application\/json" id="data">([\s\S]*?)<\/script>/);
  const data = JSON.parse(block?.[1] ?? '{}');

  it('lists exactly the approved renames, in order', () => {
    expect(data.plan.map((/** @type {any} */ p) => [p.from, p.to])).toEqual(
      renames.map((/** @type {any} */ p) => [p.from, p.to]),
    );
  });

  it('reverses every rename for the undo, and none is a no-op', () => {
    expect(data.reverse.length).toBe(renames.length);
    for (let i = 0; i < data.reverse.length; i += 1) {
      expect(data.reverse[i].from).toBe(renames[i].to);
      expect(data.reverse[i].to).toBe(renames[i].from);
      expect(data.reverse[i].from).not.toBe(data.reverse[i].to);
    }
  });

  /**
   * A line with only `into` says where a document belongs and nothing about its
   * name. It is on the page because it is what makes the folder appear on the
   * day: the model's grouping is not repeatable, and this is.
   */
  it('lists every filing line, each with a destination and no rename', () => {
    expect(data.filing.map((/** @type {any} */ p) => p.from)).toEqual(
      plan.plan.filter((/** @type {any} */ l) => !l.to && l.into).map((/** @type {any} */ p) => p.from),
    );
    for (const line of data.filing) {
      expect(line.into, line.from + ' is filed nowhere').toBeTruthy();
      expect(line.to).toBeUndefined();
    }
  });

  it('has something to demonstrate at all', () => {
    expect(data.plan.length, 'nothing held back means the live run shows nothing').toBeGreaterThan(0);
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

  it('carries every filing line through with its destination', () => {
    const filings = plan.plan.filter((/** @type {any} */ l) => !l.to && l.into);
    expect(filings.length, 'the demo needs at least one pinned destination').toBeGreaterThan(0);
  });
});
