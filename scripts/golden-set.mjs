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
 * Gate 1 passes only when: >= 90% label agreement, zero flag_for_yancey items
 * marked draftable, and at least 20 fixtures sourced from Yancey's own set.
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
const REQUIRED_REAL_FIXTURES = 20;
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
      temperature: 0,
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

const results = [];
let realFixtures = 0;

for (const file of files) {
  const fixture = JSON.parse(readFileSync(join(fixtureDir, file), 'utf8'));
  if (fixture.source === 'yancey-training-set') realFixtures += 1;

  const normalized = normalizeGraphMessage(fixture.message, { maxChars: 6000 });
  const raw = await classify(buildClassifierInput(normalized));
  const verdict = normalizeClassification(extractModelText(raw) || raw, {
    cleanText: normalized.cleanText,
    subject: normalized.subject,
    senderEmail: normalized.senderEmail,
    confidenceMin: CONFIDENCE_MIN,
    alwaysFlagSenders: fixture.alwaysFlagSenders || [],
  });

  const agreed = verdict.label === fixture.expected;
  results.push({ file, fixture, verdict, agreed });

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
    ].join(' '),
  );
  if (!agreed) console.log('        reason: ' + verdict.reasoning);
}

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
console.log('Label agreement    ' + (agreement * 100).toFixed(1) + '%  (target >= 90%)');
console.log('Flagged + drafted  ' + flaggedButDrafted.length + '  (must be 0)');
console.log('Gate 3 pricing     ' + (gate3.length - gate3Failures.length) + '/' + gate3.length + ' held');

const problems = [];
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
