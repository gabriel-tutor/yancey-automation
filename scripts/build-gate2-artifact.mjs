/**
 * Rewrite the Gate 2 artifact's data block from the harness output.
 *
 *   node scripts/build-gate2-artifact.mjs
 *
 * The artifact was first built by copying the five drafts across by hand, and
 * that dropped opp-09's "Thanks, / Yancey" sign-off — Yancey would have marked
 * it down for a fault the system never committed. Nothing is transcribed now:
 * `npm run gate2` writes fixtures/voice/gate2-drafts.json and this injects it.
 */
import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { voicePreChecks } from '../src/agent-lib.js';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const dataPath = join(root, 'fixtures/voice/gate2-drafts.json');
const htmlPath = join(root, 'docs/gate2-voice-check.html');

if (!existsSync(dataPath)) {
  console.error('fixtures/voice/gate2-drafts.json is missing. Run `npm run gate2` first.');
  process.exit(2);
}

const drafts = JSON.parse(readFileSync(dataPath, 'utf8'));
const html = readFileSync(htmlPath, 'utf8');

const OPEN = '<script type="application/json" id="data">';
const CLOSE = '</script>';
const start = html.indexOf(OPEN);
if (start === -1) {
  console.error('Could not find the data block in ' + htmlPath);
  process.exit(2);
}
const end = html.indexOf(CLOSE, start);

// A draft containing </script> would break out of the block. None ever should,
// but a rating sheet that silently corrupts its own page is not worth risking.
for (const d of drafts) {
  if (String(d.reply + d.inbound).toLowerCase().indexOf('</script') !== -1) {
    console.error('Draft ' + d.id + ' contains </script — refusing to inject.');
    process.exit(2);
  }
}

let failures = 0;
for (const d of drafts) {
  const pre = voicePreChecks(d.reply, d.subject + '\n' + d.inbound, {});
  const failed = pre.checks.filter((c) => !c.pass).map((c) => c.name);
  if (!pre.pass) failures += 1;
  console.log(
    (pre.pass ? 'PASS  ' : 'FAIL  ') + d.id.padEnd(9) +
      String(d.words).padStart(4) + 'w  ' +
      (failed.length ? failed.join(',') : 'all four checks'),
  );
}

const rebuilt =
  html.slice(0, start + OPEN.length) +
  '\n' + JSON.stringify(drafts, null, 2) + '\n' +
  html.slice(end);

writeFileSync(htmlPath, rebuilt);

console.log('\n' + drafts.length + ' drafts injected into ' + htmlPath.replace(root + '/', ''));
if (failures > 0) {
  console.log(failures + ' of them fail the Gate 2 pre-checks — fix before publishing.');
  process.exit(1);
}
console.log('All pre-checks pass. Safe to publish.');
