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

const DOCS = ['../docs/organiser-explainer.html', '../docs/reply-organiser-agent.md', '../docs/DEPLOY.md'];

describe('file names quoted in documents are real', () => {
  it.each(DOCS)('%s quotes only names from the plan', (doc) => {
    const text = readFileSync(new URL(doc, import.meta.url), 'utf8');
    const strays = [...new Set(text.match(QUOTED) || [])].filter((name) => !known.has(name));
    expect(strays, 'not in the approved plan or the library').toEqual([]);
  });
});
