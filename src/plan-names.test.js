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
  known.add(String(line.to));
}
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

  it('lists exactly the approved plan, in order', () => {
    expect(data.plan.map((/** @type {any} */ p) => [p.from, p.to])).toEqual(
      plan.plan.map((/** @type {any} */ p) => [p.from, p.to]),
    );
  });

  it('reverses every line for the undo, and none is a no-op', () => {
    expect(data.reverse.length).toBe(plan.plan.length);
    for (let i = 0; i < data.reverse.length; i += 1) {
      expect(data.reverse[i].from).toBe(plan.plan[i].to);
      expect(data.reverse[i].to).toBe(plan.plan[i].from);
      expect(data.reverse[i].from).not.toBe(data.reverse[i].to);
    }
  });

  it('has something to demonstrate at all', () => {
    expect(data.plan.length, 'nothing held back means the live run shows nothing').toBeGreaterThan(0);
  });
});
