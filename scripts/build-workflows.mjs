import { mkdirSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const outDir = join(root, 'dist');
mkdirSync(outDir, { recursive: true });

const only = process.argv.slice(2);
/** @type {Array<[string, string]>} */
const targets = [
  ['wf-e-errors', '../workflows/wf-e-errors.mjs'],
  ['wf-b-draft', '../workflows/wf-b-draft.mjs'],
  ['wf-a-classify', '../workflows/wf-a-classify.mjs'],
  ['wf-c-digest', '../workflows/wf-c-digest.mjs'],
  ['wf-t-preflight', '../workflows/wf-t-preflight.mjs'],
  ['wf-s-seed', '../workflows/wf-s-seed.mjs'],
];

for (const [name, path] of targets) {
  if (only.length > 0 && !only.includes(name)) continue;
  try {
    const mod = await import(path);
    const source = await mod.build();
    const file = join(outDir, name + '.sdk.ts');
    writeFileSync(file, source);
    console.log(String(source.length).padStart(7), 'bytes  ', file);
  } catch (err) {
    console.error('SKIP', name, '->', err instanceof Error ? err.message : String(err));
  }
}
