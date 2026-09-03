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
  'wf-a/Attach Findings': { lib: ['buildClassifierInput'] },
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
  'wf-e/Normalize Error': { lib: ['normalizeErrorPayload'] },

  // Attachment ingestion. The triage node decides what to open; the extract
  // node does the in-process work (Office unzip, text decode) and builds the
  // request for anything that needs the model instead.
  'wf-d/Triage Attachments': { attachments: ['triageAll'] },
  'wf-d/Extract': { unzip: ['officeText'] },
  'wf-d/Summarise': { attachments: ['summariseAttachments'] },

  // The file organiser. Deciding what a document should be called is the whole
  // of the judgement, so it lives in src/filenames.js where it is tested.
  'wf-f/Pick The Reports': { library: ['reportsIn', 'isReport'] },
  'wf-f/Read Reports': { powerbi: ['dependencyIndex'] },
  'wf-f/Review The Library': { filenames: ['planRenames'], library: ['categoriseAll', 'planPlacements', 'reportsIn'] },
  'wf-f/Pick What To Name': { naming: ['namingCandidates'], library: ['reportsIn'] },
  'wf-f/Read Contents': { unzip: ['officeText'] },
  'wf-f/Compose Request': { naming: ['buildNamingRequest'] },
  // The whole of the apply side: what the model proposed, what a person
  // approved, the structure it invented, and the calls all of that turns into.
  'wf-f/Apply Proposals': {
    naming: ['applyProposals', 'applyApprovedPlan', 'approvedProgress', 'namingCandidates'],
    structure: ['planStructure', 'planFolderRenames'],
    changes: [
      'plannedChanges',
      'refuseImpossible',
      'fittingActions',
      'foldersNeeded',
      'fileChangeUrl',
      'folderCreateUrl',
      'folderChangeUrl',
    ],
    library: ['reportsIn', 'folderOf'],
  },
};
