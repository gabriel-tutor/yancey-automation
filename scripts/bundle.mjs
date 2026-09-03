import * as esbuild from 'esbuild';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, '..');

/**
 * @typedef {{ lib?: string[], prompts?: string[], unzip?: string[], attachments?: string[], filenames?: string[], library?: string[], naming?: string[], powerbi?: string[], structure?: string[], changes?: string[] }} HelperNeeds
 */

/** @type {Record<string, string>} */
const SOURCES = {
  lib: './src/agent-lib.js',
  prompts: './src/prompts.js',
  // Attachment ingestion. `unzip` carries a DEFLATE decompressor because n8n's
  // Code node refuses require('zlib'), and .docx/.pptx/.xlsx are ZIP archives.
  unzip: './src/unzip.js',
  attachments: './src/attachments.js',
  // Yancey's document naming convention, sent 28 August 2026.
  filenames: './src/filenames.js',
  // What a document is and where it sits — the judgements that only make
  // sense across a whole library rather than one name at a time.
  library: './src/library.js',
  // Naming by a model, checked by code. The model supplies the words; this
  // supplies the shape and refuses anything that does not hold.
  naming: './src/naming.js',

  // Reading what a Power BI report depends on, so the guard stops guessing.
  powerbi: './src/powerbi.js',

  // The folder structure the model invents, held to a small one.
  structure: './src/structure.js',
  // Turning a reviewed plan into calls SharePoint will accept.
  changes: './src/changes.js',
};

/** @type {ReadonlyArray<'lib' | 'prompts' | 'unzip' | 'attachments' | 'filenames' | 'library' | 'naming' | 'powerbi' | 'structure' | 'changes'>} */
const SOURCE_KEYS = ['lib', 'prompts', 'unzip', 'attachments', 'filenames', 'library', 'naming', 'powerbi', 'structure', 'changes'];


/**
 * Drop JSDoc/banner blocks and whole-line `//` comments. Line-based and
 * conservative: a block is only removed when the trimmed line opens with
 * `/*`, so a regex literal containing slashes is never touched.
 * @param {string} source
 * @returns {string}
 */
function stripComments(source) {
  const out = [];
  let inBlock = false;
  for (const line of source.split('\n')) {
    const trimmed = line.trim();
    if (inBlock) {
      if (trimmed.endsWith('*/')) inBlock = false;
      continue;
    }
    if (trimmed.startsWith('/*')) {
      if (!trimmed.endsWith('*/')) inBlock = true;
      continue;
    }
    if (trimmed.startsWith('//')) continue;
    out.push(line);
  }
  return out.join('\n').replace(/\n{3,}/g, '\n\n').trim();
}

/**
 * Tree-shake the shared library down to exactly the names one Code node uses.
 *
 * The n8n Code node has no module system, so the helpers have to be inlined.
 * Inlining the whole library into eight nodes would add ~300KB of duplicated
 * source to the workflows; esbuild's tree shaking keeps each node to only its
 * transitive closure. Comments are kept — these bodies are read in the n8n
 * editor during handover.
 *
 * @param {HelperNeeds} needs
 * @returns {Promise<string>}
 */
export async function subset(needs) {
  /** @type {string[]} */
  const imports = [];
  /** @type {string[]} */
  const names = [];
  for (const key of SOURCE_KEYS) {
    const wanted = needs[key];
    if (!wanted || wanted.length === 0) continue;
    imports.push(`export { ${wanted.join(', ')} } from '${SOURCES[key]}';`);
    names.push(...wanted);
  }
  if (imports.length === 0) return '';

  const result = await esbuild.build({
    stdin: { contents: imports.join('\n'), resolveDir: root, sourcefile: 'subset.js', loader: 'js' },
    bundle: true,
    format: 'esm',
    treeShaking: true,
    minify: false,
    write: false,
    target: 'es2021',
    legalComments: 'none',
    logLevel: 'silent',
  });

  const emitted = result.outputFiles[0];
  if (emitted === undefined) throw new Error('subset: esbuild produced no output');
  const out = stripComments(emitted.text);
  const stripped = out.replace(/\nexport \{[\s\S]*?\};\s*$/, '\n').trimEnd();
  if (/^\s*(import|export)\s/m.test(stripped)) {
    throw new Error('subset: emitted code still contains import/export');
  }
  for (const name of names) {
    if (!new RegExp('\\b(?:function|const|let|var|class)\\s+' + name + '\\b').test(stripped)) {
      throw new Error('subset: tree shaking dropped a requested name: ' + name);
    }
  }
  return stripped;
}

/**
 * @param {HelperNeeds} needs
 * @param {string} body node-specific code appended after the shared helpers
 * @returns {Promise<string>}
 */
export async function codeNode(needs, body) {
  const helpers = await subset(needs);
  return [
    '// ---- shared helpers (generated from src/ - edit there, not here) ----',
    helpers,
    '// ---- end shared helpers ----',
    '',
    body.trim(),
    '',
  ].join('\n');
}
