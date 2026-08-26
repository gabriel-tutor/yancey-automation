/**
 * Gate 1 / Gate 3 regression harness (spec §11).
 *
 * Runs every fixture in fixtures/golden/ through the exact prompt, schema and
 * guard chain that WF-A uses in production, then scores agreement against the
 * labels Yancey supplied.
 *
 *   OPENAI_API_KEY=sk-... npm run gate1
 *   OPENAI_API_KEY=sk-... npm run gate1 -- --model gpt-5.6-luna
 *
 * Two things are scored:
 *
 *   Behaviour (the gate). Yancey's labels are a two-way operational judgement —
 *   reply to this, or ignore it. Every fixture records that as `behavior`, and
 *   the gate turns on whether the agent did the right thing. The two failures
 *   that matter are `missed` (real work classified as noise, so it silently
 *   disappears) and `false-draft` (junk that got a draft written). A
 *   `safe-hold` — surfaced to Yancey but not drafted — is not a failure: he
 *   still sees the email, the agent just declined to put words in his mouth.
 *
 *   Label agreement (diagnostic). Strict match against the four-way taxonomy
 *   the deployed prompt defines.
 *
 * Gate 1 passes only when: zero missed, zero false-draft, >= 90% label
 * agreement, zero flag_for_yancey items marked draftable, and all 15 of
 * Yancey's classification fixtures present.
 */
import { readFileSync, readdirSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import {
  buildClassifierInput,
  extractModelText,
  normalizeClassification,
  normalizeGraphMessage,
} from '../src/agent-lib.js';
import { CLASSIFY_SCHEMA, buildClassifySystemPrompt } from '../src/prompts.js';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const args = process.argv.slice(2);
/**
 * @param {string} name
 * @param {string} fallback
 * @returns {string}
 */
const flag = (name, fallback) => {
  const i = args.indexOf('--' + name);
  const value = i === -1 ? undefined : args[i + 1];
  return value === undefined ? fallback : value;
};

const MODEL = flag('model', process.env.MODEL_CLASSIFY || 'gpt-5.6-luna');
const CONFIDENCE_MIN = Number(flag('confidence-min', '0.6'));
// Yancey sent 20 labelled emails: 10 OPPORTUNITY, 5 JUNK, 5 MY WRITING STYLE.
// Only the first 15 are classification cases. The 5 writing-style samples are
// mail he *wrote* — they live in fixtures/voice/ and calibrate the draft voice;
// running them through the classifier would measure nothing.
const REQUIRED_REAL_FIXTURES = 15;
const AGREEMENT_TARGET = 0.9;

const apiKey = process.env.OPENAI_API_KEY;
if (!apiKey) {
  console.error('OPENAI_API_KEY is not set. Export the key from the Upperform OpenAI account.');
  process.exit(2);
}

const profilePath = join(root, 'context/business-profile.md');
if (!existsSync(profilePath)) {
  console.error('context/business-profile.md is missing — the harness must score the real prompt.');
  process.exit(2);
}
const profile = readFileSync(profilePath, 'utf8');
const systemPrompt = buildClassifySystemPrompt(profile);

const fixtureDir = join(root, 'fixtures/golden');
const files = existsSync(fixtureDir)
  ? readdirSync(fixtureDir).filter((f) => f.endsWith('.json')).sort()
  : [];

if (files.length === 0) {
  console.error('No fixtures in fixtures/golden/. See fixtures/README.md for the format.');
  process.exit(2);
}

/**
 * @param {string} userInput
 * @returns {Promise<any>}
 */
async function classify(userInput) {
  const response = await fetch('https://api.openai.com/v1/responses', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: 'Bearer ' + apiKey,
    },
    body: JSON.stringify({
      model: MODEL,
      instructions: systemPrompt,
      input: [{ role: 'user', content: userInput }],
      store: false,
      text: {
        format: {
          type: 'json_schema',
          name: 'upperform_email_triage',
          strict: true,
          schema: CLASSIFY_SCHEMA,
        },
      },
    }),
  });

  if (!response.ok) {
    throw new Error('OpenAI ' + response.status + ': ' + (await response.text()).slice(0, 400));
  }
  return response.json();
}

/**
 * Score the operational outcome against Yancey's own two-way judgement.
 * @param {{behavior?: string}} fixture
 * @param {{label: string, requiresDraft: boolean}} verdict
 * @returns {'correct'|'safe-hold'|'missed'|'false-draft'}
 */
function outcomeOf(fixture, verdict) {
  const drafted = verdict.requiresDraft;
  const label = verdict.label;
  const draftable = label === 'opportunity' || label === 'admin';

  if (fixture.behavior === 'draft') {
    if (draftable && drafted) return 'correct';
    if (label === 'noise') return 'missed';
    return 'safe-hold';
  }
  if (fixture.behavior === 'ignore') {
    if (drafted) return 'false-draft';
    return label === 'noise' ? 'correct' : 'safe-hold';
  }
  // behavior === 'flag': the pricing/sensitivity cases
  if (drafted) return 'false-draft';
  return label === 'flag_for_yancey' ? 'correct' : 'missed';
}

const results = [];
let realFixtures = 0;

for (const file of files) {
  const fixture = JSON.parse(readFileSync(join(fixtureDir, file), 'utf8'));
  if (fixture.source === 'yancey-training-set') realFixtures += 1;

  const normalized = normalizeGraphMessage(fixture.message, { maxChars: 6000 });
  const raw = await classify(buildClassifierInput(normalized));
  const verdict = normalizeClassification(extractModelText(raw) || raw, {
    cleanText: normalized.cleanText,
    threadText: normalized.threadText,
    subject: normalized.subject,
    senderEmail: normalized.senderEmail,
    confidenceMin: CONFIDENCE_MIN,
    alwaysFlagSenders: fixture.alwaysFlagSenders || [],
  });

  const agreed = verdict.label === fixture.expected;
  const outcome = fixture.behavior ? outcomeOf(fixture, verdict) : 'correct';
  results.push({ file, fixture, verdict, agreed, outcome });

  const mark = agreed ? 'PASS' : 'FAIL';
  console.log(
    [
      mark.padEnd(5),
      file.padEnd(30),
      ('expected=' + fixture.expected).padEnd(28),
      ('got=' + verdict.label).padEnd(24),
      'conf=' + verdict.confidence.toFixed(2),
      verdict.requiresDraft ? 'draft' : 'no-draft',
      verdict.guardFlags ? '[' + verdict.guardFlags + ']' : '',
      fixture.behavior && outcome !== 'correct' ? '<< ' + outcome.toUpperCase() : '',
    ].join(' '),
  );
  if (!agreed) console.log('        reason: ' + verdict.reasoning);
}

const scored = results.filter((r) => r.fixture.behavior);
const missed = scored.filter((r) => r.outcome === 'missed');
const falseDrafts = scored.filter((r) => r.outcome === 'false-draft');
const safeHolds = scored.filter((r) => r.outcome === 'safe-hold');
const behaviourOk = scored.filter((r) => r.outcome === 'correct').length;
const agreement = results.filter((r) => r.agreed).length / results.length;
const flaggedButDrafted = results.filter(
  (r) => r.verdict.label === 'flag_for_yancey' && r.verdict.requiresDraft,
);
const gate3 = results.filter((r) => r.fixture.gate === 'pricing');
const gate3Failures = gate3.filter(
  (r) => r.verdict.label !== 'flag_for_yancey' || r.verdict.requiresDraft,
);

console.log('\n' + '='.repeat(72));
console.log('Model              ' + MODEL);
console.log('Fixtures           ' + results.length + ' (' + realFixtures + ' from Yancey\'s labelled set)');
console.log(
  'Did the right thing ' + behaviourOk + '/' + scored.length +
    '   (+' + safeHolds.length + ' surfaced-not-drafted)',
);
console.log('Missed real work   ' + missed.length + '  (must be 0)');
console.log('Drafted junk       ' + falseDrafts.length + '  (must be 0)');
console.log('Label agreement    ' + (agreement * 100).toFixed(1) + '%  (target >= 90%)');
console.log('Flagged + drafted  ' + flaggedButDrafted.length + '  (must be 0)');
console.log('Gate 3 pricing     ' + (gate3.length - gate3Failures.length) + '/' + gate3.length + ' held');

const problems = [];
for (const r of missed) {
  problems.push('MISSED: ' + r.file + ' — Yancey said reply, agent said noise');
}
for (const r of falseDrafts) {
  problems.push('FALSE DRAFT: ' + r.file + ' — Yancey said ignore, agent drafted a reply');
}
if (agreement < AGREEMENT_TARGET) problems.push('agreement below 90%');
if (flaggedButDrafted.length > 0) problems.push('a flag_for_yancey item was marked draftable');
if (gate3Failures.length > 0) problems.push('a pricing request escaped flag_for_yancey');
if (realFixtures < REQUIRED_REAL_FIXTURES) {
  problems.push(
    'only ' + realFixtures + '/' + REQUIRED_REAL_FIXTURES +
      " fixtures come from Yancey's labelled set — Gate 1 is not decided until the real 20 are in fixtures/golden/",
  );
}

if (problems.length === 0) {
  console.log('\nGATE 1 + GATE 3: PASS');
  process.exit(0);
}
console.log('\nGATE 1 + GATE 3: NOT PASSED');
for (const p of problems) console.log('  - ' + p);
process.exit(1);
