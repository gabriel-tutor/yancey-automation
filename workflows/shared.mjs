import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
export const ROOT = join(here, '..');

export const DEPLOY = JSON.parse(readFileSync(join(ROOT, 'config/deploy.json'), 'utf8'));

/**
 * Which mailbox this build targets. `npm run build` is prod (Yancey);
 * `TARGET=dev npm run build` is the gtutor sandbox.
 *
 * The credential is referenced by *name* in generated SDK code, so this is not
 * cosmetic: with the name hardcoded, renaming a credential in n8n silently
 * repoints every workflow built afterwards. Yancey's credential was renamed
 * from "Outlook – Dev (gtutor)" to "Outlook – Yancey", and a new dev
 * credential took the old name — so a prod rebuild would have pointed his
 * workflows at the dev mailbox.
 */
const TARGET_NAME = process.env.TARGET === 'dev' ? 'dev' : 'prod';
export const TARGET = { name: TARGET_NAME, ...DEPLOY.targets[TARGET_NAME] };
export const OUTLOOK_CREDENTIAL = TARGET.outlookCredential;

/**
 * Yancey's drafting rules. Deliberately not a literal in the workflow: the
 * golden-set harness reads the same block, so a gate can never certify a policy
 * that differs from the one running in his mailbox.
 */
export const POLICY = DEPLOY.policy || { internalDomains: [], alwaysFlagSenders: [] };
export const TABLES = TARGET.dataTables;

/** Suffix appended to every workflow name, so dev copies are unmistakable. */
export const SUFFIX = TARGET.nameSuffix;

/**
 * JS string literal, safe for embedding in generated SDK source.
 * @param {unknown} value
 * @returns {string}
 */
export const S = (value) => JSON.stringify(value);

/**
 * An n8n expression parameter.
 * @param {string} value
 * @returns {string}
 */
export const E = (value) => 'expr(' + JSON.stringify(value) + ')';

export const GRAPH = 'https://graph.microsoft.com/v1.0';

export const SDK_IMPORT =
  "import { workflow, node, trigger, sticky, ifElse, merge, splitOut, newCredential, expr } from '@n8n/workflow-sdk';";

/**
 * The three-node chain that loads the mounted business profile with a fallback.
 * @param {{ x: number, y: number }} position
 * @returns {string}
 */
export function contextLoaderNodes({ x, y }) {
  return `
const loadProfileFile = node({
  type: 'n8n-nodes-base.readWriteFile',
  version: 1.1,
  config: {
    name: 'Load Business Profile',
    onError: 'continueRegularOutput',
    parameters: {
      operation: 'read',
      fileSelector: ${S(DEPLOY.contextFilePath)},
      options: { dataPropertyName: 'profileFile', mimeType: 'text/markdown' }
    },
    position: [${x}, ${y}]
  },
  output: [{ profileFile: 'binary' }]
});

const extractProfileText = node({
  type: 'n8n-nodes-base.extractFromFile',
  version: 1.1,
  config: {
    name: 'Extract Profile Text',
    onError: 'continueRegularOutput',
    parameters: {
      operation: 'text',
      binaryPropertyName: 'profileFile',
      destinationKey: 'profileText'
    },
    position: [${x + 200}, ${y}]
  },
  output: [{ profileText: '# Upperform Partners — Agent Context File' }]
});
`;
}
