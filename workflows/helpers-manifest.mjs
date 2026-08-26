/**
 * Which shared helpers each n8n Code node needs.
 *
 * Single source of truth: the workflow generators build their Code nodes from
 * these entries, and src/bundle.test.js proves every entry still tree-shakes
 * to working code. Change a node's helper list here, nowhere else.
 */
export const HELPERS = {
  'wf-a/Resolve Context': { prompts: ['buildClassifySystemPrompt', 'FALLBACK_PROFILE'] },
  'wf-a/Prepare Email': { lib: ['normalizeGraphMessage', 'buildClassifierInput'] },
  'wf-a/Apply Classification': {
    lib: ['normalizeClassification', 'decideCategories', 'addBusinessDays', 'addressedToMailbox'],
  },
  'wf-b/Resolve Context': { prompts: ['buildDraftSystemPrompt', 'FALLBACK_PROFILE'] },
  'wf-b/Compose Draft Prompt': { prompts: ['buildDraftUserPrompt'] },
  'wf-b/Voice Guard': {
    lib: [
      'extractModelText',
      'voiceGuard',
      'textToHtmlParagraphs',
      'injectReplyIntoDraftHtml',
      'describeDraftFailure',
    ],
  },
  'wf-c/Build Digest': { lib: ['bucketDigestRows', 'renderDigestHtml', 'renderDigestText'] },
};
