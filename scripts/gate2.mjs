/**
 * Gate 2 — voice (spec §11).
 *
 *   "draft replies to 5 opportunity samples; Yancey rates voice fit 1–5;
 *    average >= 4. Automated pre-checks: contains 'Thanks,' + 'Yancey'; no
 *    'I hope this email finds you well'; no invented digits (regex for
 *    currency/percent not present in the source); under 150 words."
 *
 * Two phases, because half of this gate is a human judgement and cannot be
 * automated away:
 *
 *   1. OPENAI_API_KEY=sk-... npm run gate2
 *      Drafts replies to the five samples through the exact prompt, model and
 *      voice guard WF-B uses, runs the automated pre-checks, and writes a
 *      rating sheet for Yancey.
 *
 *   2. Yancey fills in fixtures/voice/gate2-ratings.json, then the same
 *      command scores it and passes or fails the gate.
 *
 * The drafts are generated fresh each run unless --keep is passed, so a rating
 * always refers to text someone actually read.
 */
import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import {
  buildDraftUserPrompt,
  buildDraftSystemPrompt,
} from '../src/prompts.js';
import {
  extractModelText,
  normalizeGraphMessage,
  voiceGuard,
  voicePreChecks,
} from '../src/agent-lib.js';
import { SAMPLES } from './gate2-samples.mjs';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const args = process.argv.slice(2);
const keepExisting = args.indexOf('--keep') !== -1;

const MODEL = process.env.MODEL_DRAFT || 'gpt-5.6-terra';
const MAX_WORDS = 150;
const SIGN_OFF = 'Thanks,';
const SIGNATURE = 'Yancey';
const RATING_TARGET = 4;

const ratingsPath = join(root, 'fixtures/voice/gate2-ratings.json');
const sheetPath = join(root, 'fixtures/voice/gate2-drafts.md');
const jsonPath = join(root, 'fixtures/voice/gate2-drafts.json');

/* ------------------------------------------------------------------ */
/* Phase 2 first: if a complete rating sheet exists, score it.          */
/* ------------------------------------------------------------------ */

if (existsSync(ratingsPath)) {
  const ratings = JSON.parse(readFileSync(ratingsPath, 'utf8'));
  /** @type {{ id: string, subject: string, score: number|null, comment: string }[]} */
  const rows = Array.isArray(ratings.ratings) ? ratings.ratings : [];
  const scored = rows.filter(
    (/** @type {{ score: number|null }} */ r) =>
      typeof r.score === 'number' && r.score >= 1 && r.score <= 5,
  );

  if (scored.length === rows.length && rows.length > 0) {
    const total = scored.reduce(
      (/** @type {number} */ sum, /** @type {{ score: number|null }} */ r) => sum + Number(r.score),
      0,
    );
    const average = total / scored.length;

    console.log('GATE 2 — voice\n' + '='.repeat(60));
    for (const r of scored) {
      console.log('  ' + String(r.id).padEnd(10) + r.score + '/5   ' + (r.comment || ''));
    }
    console.log('-'.repeat(60));
    console.log('  Average          ' + average.toFixed(2) + '  (target >= ' + RATING_TARGET + ')');
    console.log('  Rated by         ' + (ratings.ratedBy || '[CONFIRM who rated these]'));
    console.log('  Drafts generated ' + (ratings.generatedAt || 'unknown'));

    if (average >= RATING_TARGET) {
      console.log('\nGATE 2: PASS');
      process.exit(0);
    }
    console.log('\nGATE 2: NOT PASSED — average voice fit below ' + RATING_TARGET);
    console.log('The drafts he scored lowest are the ones to read first; the fix is');
    console.log("almost always the voice section of context/business-profile.md.");
    process.exit(1);
  }

  if (!keepExisting) {
    console.log(
      'Found ' + ratingsPath.replace(root + '/', '') + ' with ' +
        scored.length + '/' + rows.length + ' scores filled in.',
    );
    console.log('Regenerating the drafts. Pass --keep to score a partial sheet instead.\n');
  }
}

/* ------------------------------------------------------------------ */
/* Phase 1: draft the five samples through the production path.        */
/* ------------------------------------------------------------------ */

const apiKey = process.env.OPENAI_API_KEY;
if (!apiKey) {
  console.error('OPENAI_API_KEY is not set. Export the key from the Upperform OpenAI account.');
  process.exit(2);
}

const profilePath = join(root, 'context/business-profile.md');
if (!existsSync(profilePath)) {
  console.error('context/business-profile.md is missing — Gate 2 must score the real voice section.');
  process.exit(2);
}
const profile = readFileSync(profilePath, 'utf8');
const draftSystem = buildDraftSystemPrompt(profile, {
  signature: SIGNATURE,
  maxWords: MAX_WORDS,
});

/**
 * @param {string} userPrompt
 * @returns {Promise<any>}
 */
async function draft(userPrompt) {
  const response = await fetch('https://api.openai.com/v1/responses', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + apiKey },
    body: JSON.stringify({
      model: MODEL,
      instructions: draftSystem,
      input: [{ role: 'user', content: userPrompt }],
      store: false,
      prompt_cache_key: 'upperform-draft-v1',
    }),
  });
  if (!response.ok) {
    throw new Error('OpenAI ' + response.status + ': ' + (await response.text()).slice(0, 400));
  }
  return response.json();
}

const generatedAt = new Date().toISOString();
const results = [];
let preCheckFailures = 0;

for (const sample of SAMPLES) {
  const fixturePath = join(root, 'fixtures/golden', sample.id + '.json');
  if (!existsSync(fixturePath)) {
    console.error('Missing fixture ' + sample.id + ' — Gate 2 draws its samples from the golden set.');
    process.exit(2);
  }
  const fixture = JSON.parse(readFileSync(fixturePath, 'utf8'));
  const n = normalizeGraphMessage(fixture.message, { maxChars: 6000 });

  const userPrompt = buildDraftUserPrompt({
    subject: n.subject,
    senderName: n.senderName,
    senderEmail: n.senderEmail,
    cleanText: n.cleanText,
    reasoning: fixture.note || 'live engagement work',
    asks: [],
    hasAttachments: n.hasAttachments,
    label: 'opportunity',
    urgency: 'this_week',
  });

  const raw = await draft(userPrompt);
  const generated = extractModelText(raw);

  // Same guard WF-B runs, so the rated text is the text that would land.
  const guard = voiceGuard(generated, n.subject + '\n' + n.cleanText, {
    signOff: SIGN_OFF,
    signature: SIGNATURE,
    maxWords: MAX_WORDS + 70,
  });

  const pre = voicePreChecks(guard.text, n.subject + '\n' + n.cleanText, {
    signOff: SIGN_OFF,
    signature: SIGNATURE,
    maxWords: MAX_WORDS,
  });
  if (!pre.pass) preCheckFailures += 1;

  results.push({ sample, n, guard, pre, repaired: guard.soft });

  console.log(
    (pre.pass ? 'PASS  ' : 'FAIL  ') + sample.id.padEnd(10) +
      String(guard.words).padStart(3) + 'w  ' +
      pre.checks.filter((c) => !c.pass).map((c) => c.name).join(',') +
      (guard.soft.length ? '  (guard repaired: ' + guard.soft.join(',') + ')' : ''),
  );
}

/* ------------------------------------------------------------------ */
/* Write the rating sheet.                                             */
/* ------------------------------------------------------------------ */

const sheet = [
  '# Gate 2 — voice rating sheet',
  '',
  'Replies the assistant wrote to real emails from your own inbox.',
  '',
  '**Rate each 1–5 on one question only: does this sound like you?**',
  'Not whether it is a good reply — whether it is *your* reply. A 3 means you',
  'would send it after edits. A 5 means you would send it as-is.',
  '',
  'Put your scores in `fixtures/voice/gate2-ratings.json` and re-run',
  '`npm run gate2`. The gate needs an average of 4 or better.',
  '',
  'Drafts generated ' + generatedAt + ' · model ' + MODEL,
  '',
  '---',
  '',
];

let index = 0;
for (const r of results) {
  index += 1;
  sheet.push('## ' + index + '. ' + r.sample.id + ' — ' + r.n.subject);
  sheet.push('');
  sheet.push('*Why this one: ' + r.sample.why + '*');
  sheet.push('');
  sheet.push('**They wrote** (from ' + (r.n.senderName || r.n.senderEmail) + '):');
  sheet.push('');
  sheet.push('> ' + r.n.cleanText.split('\n').slice(0, 14).join('\n> '));
  sheet.push('');
  sheet.push('**The assistant replied:**');
  sheet.push('');
  sheet.push('```');
  sheet.push(r.guard.text);
  sheet.push('```');
  sheet.push('');
  sheet.push('Automated pre-checks: ' + r.pre.checks.map((c) => (c.pass ? '✓ ' : '✗ ') + c.name).join(' · '));
  if (r.repaired.length) {
    sheet.push('');
    sheet.push('*The guard repaired this one before you saw it: ' + r.repaired.join(', ') + '.*');
  }
  sheet.push('');
  sheet.push('**Your score (1–5): ____**');
  sheet.push('');
  sheet.push('---');
  sheet.push('');
}

writeFileSync(sheetPath, sheet.join('\n'));

writeFileSync(
  jsonPath,
  JSON.stringify(
    results.map((r) => ({
      id: r.sample.id,
      register: r.sample.register,
      subject: r.n.subject,
      why: r.sample.why,
      from: r.n.senderName || r.n.senderEmail,
      inbound: r.n.cleanText,
      reply: r.guard.text,
      words: r.guard.words,
      ...(r.sample.flag ? { flag: r.sample.flag } : {}),
    })),
    null,
    2,
  ) + '\n',
);

if (!existsSync(ratingsPath) || !keepExisting) {
  writeFileSync(
    ratingsPath,
    JSON.stringify(
      {
        _instructions:
          'Score each draft 1-5 on whether it sounds like Yancey. ' +
          'Read fixtures/voice/gate2-drafts.md, fill in every "score", then run npm run gate2.',
        ratedBy: '',
        generatedAt,
        model: MODEL,
        ratings: results.map((r) => ({
          id: r.sample.id,
          subject: r.n.subject,
          score: null,
          comment: '',
        })),
      },
      null,
      2,
    ) + '\n',
  );
}

console.log('\n' + '='.repeat(60));
console.log('Model              ' + MODEL);
console.log('Samples drafted    ' + results.length);
console.log('Pre-checks failed  ' + preCheckFailures + '  (must be 0)');
console.log('Rating sheet       ' + sheetPath.replace(root + '/', ''));
console.log('Artifact data      ' + jsonPath.replace(root + '/', ''));
console.log('Scores go in       ' + ratingsPath.replace(root + '/', ''));

if (preCheckFailures > 0) {
  console.log('\nGATE 2: NOT PASSED — automated pre-checks failed before Yancey sees it');
  process.exit(1);
}
console.log('\nGATE 2: awaiting Yancey\'s ratings. Pre-checks all pass.');
process.exit(0);
