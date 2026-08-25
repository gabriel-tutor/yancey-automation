import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
export const ROOT = join(here, '..');

export const DEPLOY = JSON.parse(readFileSync(join(ROOT, 'config/deploy.json'), 'utf8'));

/** JS string literal, safe for embedding in generated SDK source. */
export const S = (value) => JSON.stringify(value);

/** An n8n expression parameter. */
export const E = (value) => 'expr(' + JSON.stringify(value) + ')';

export const GRAPH = 'https://graph.microsoft.com/v1.0';

export const SDK_IMPORT =
  "import { workflow, node, trigger, sticky, ifElse, splitOut, newCredential, expr } from '@n8n/workflow-sdk';";

/** The three-node chain that loads the mounted business profile with a fallback. */
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
