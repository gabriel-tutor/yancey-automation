import { describe, expect, it } from 'vitest';
import { codeNode } from '../scripts/bundle.mjs';
import { HELPERS } from '../workflows/helpers-manifest.mjs';
import { VOICE_GUARD } from '../workflows/wf-b-draft.mjs';
import { APPLY_CLASSIFICATION } from '../workflows/wf-a-classify.mjs';
import { APPLY as ORGANISER_APPLY, CONFIG_MERGE, REVIEW as ORGANISER_REVIEW, SUMMARY as ORGANISER_SUMMARY } from '../workflows/wf-f-organizer.mjs';
import { ANSWER, DECIDE } from '../workflows/wf-g-selfserve.mjs';

/**
 * The Code node *bodies* — not just the helper library they import.
 *
 * Two production bugs got in through this gap: a paste that silently dropped a
 * non-breaking space, and an orphaned draft left in the mailbox when the model
 * failed. Both were invisible to the unit tests because the node bodies were
 * only ever strings. These run them.
 */

const voiceGuardCode = await codeNode(HELPERS['wf-b/Voice Guard'], VOICE_GUARD);
/**
 * The node reads whichever upstream node is authoritative for this build, and
 * the build substitutes the name. Production pairs with Prepare Email; a target
 * that reads attachments pairs with Attach Findings instead.
 * @param {string} sourceNode
 */
const applyClassificationFor = (sourceNode) =>
  codeNode(HELPERS['wf-a/Apply Classification'], APPLY_CLASSIFICATION.replace('SOURCE_NODE_NAME', sourceNode));

const applyClassificationCode = await applyClassificationFor('Prepare Email');
const applyWithAttachmentsCode = await applyClassificationFor('Attach Findings');

/**
 * @param {Record<string, any>} map
 * @returns {(name: string) => any}
 */
const nodeRef = (map) => (name) => {
  if (!(name in map)) throw new Error("node '" + name + "' did not run");
  const value = map[name];
  // An array stands for a node that emitted more than one item — a paginated
  // HTTP request, one item per page. Anything else is the single item it is.
  const items = (Array.isArray(value) ? value : [value]).map((json) => ({ json }));
  return { first: () => items[0], all: () => items, item: items[0] };
};

/**
 * @param {string} body
 * @param {Record<string, any>} globals
 * @returns {any}
 */
const run = (body, globals) => {
  const names = Object.keys(globals);
  return new Function(...names, body)(...names.map((n) => globals[n]));
};

describe('WF-B Voice Guard node body', () => {
  const cfg = { maxDraftWords: 150, signOff: 'Thanks,', signature: 'Yancey' };
  const req = {
    messageId: 'AAMkMSG',
    messageIdEnc: 'AAMkMSG%3D',
    subject: 'RE: Vector CI',
    cleanText: 'Please send the crossarm workbook.',
  };
  const composed = {
    draftId: 'AAMkDRAFT',
    draftIdEnc: 'AAMkDRAFT%3D',
    draftWebLink: 'https://outlook/x',
    draftBodyHtml: '<html><body><div>quoted</div></body></html>',
  };
  /** @param {any} inputJson */
  const globals = (inputJson) => ({
    $: nodeRef({ Config: cfg, 'Draft Request': req, 'Compose Draft Prompt': composed }),
    $input: { first: () => ({ json: inputJson }) },
  });

  it('writes a reply when the model returns one', () => {
    const out = run(voiceGuardCode, globals({
      content: 'Hey Chuck,\n\nI will pull the crossarm workbook together before the Momentum Room.\n\nThanks,\n\nYancey',
    }));
    expect(out[0].json.ok).toBe(true);
    expect(out[0].json.status).toBe('drafted');
    expect(out[0].json.finalHtml).toContain('Hey Chuck');
  });

  // The orphaned-draft bug. createReply has already put an empty reply in the
  // mailbox; if this throws, only the blocked branch would have deleted it.
  it('absorbs a model failure and routes it to the cleanup branch', () => {
    const out = run(voiceGuardCode, globals({ error: { message: 'OpenAI 500: upstream timeout' } }));
    expect(out[0].json.ok).toBe(false);
    expect(out[0].json.status).toBe('draft_failed');
    expect(out[0].json.guardFlags).toBe('draft_generation_failed');
    expect(out[0].json.finalHtml).toBe('');
    expect(out[0].json.extraCategories).toEqual(['AI-Review']);
  });

  it('keeps every identifier the cleanup branch needs', () => {
    const out = run(voiceGuardCode, globals({ error: 'boom' }));
    // Deleting the empty draft needs this one.
    expect(out[0].json.draftIdEnc).toBe('AAMkDRAFT%3D');
    // Tagging and logging the original message need these.
    expect(out[0].json.messageId).toBe('AAMkMSG');
    expect(out[0].json.messageIdEnc).toBe('AAMkMSG%3D');
  });
});

describe('WF-A Apply Classification node body', () => {
  const cfg = {
    confidenceMin: 0.6,
    followupDays: 3,
    alwaysFlagSenders: [],
    draftingMailbox: 'gtutor@partnersperform.com',
  };
  const prepared = {
    messageId: 'M1', messageIdEnc: 'M1', conversationId: 'C1', internetMessageId: '<x>',
    receivedAt: '2026-08-27T09:00:00Z',
    senderEmail: 'ccoleman@partnersperform.com', senderName: 'Charles Coleman',
    subject: 'RE: Vector CI', webLink: 'https://o/x', hasAttachments: false, categories: [],
    cleanText: 'Can you pull the workbook together before Thursday?',
    threadText: 'Can you pull the workbook together before Thursday?',
    toRecipients: ['gtutor@partnersperform.com'], ccRecipients: [],
  };
  const model = {
    label: 'opportunity', confidence: 0.95, reasoning: 'live work',
    requires_draft: true, urgency: 'this_week', asks_of_yancey: ['send workbook'],
  };
  /** @param {any} prep */
  const globals = (prep) => ({
    $: nodeRef({ Config: cfg, 'Prepare Email': prep, 'Resolve Context': { profileSource: 'file' } }),
    $json: model,
    $itemIndex: 0,
  });

  it('carries what the attachments said through to the log and the draft', () => {
    const out = run(applyWithAttachmentsCode, {
      $: nodeRef({
        Config: cfg,
        'Attach Findings': {
          ...prepared,
          hasAttachments: true,
          attachmentText: '- tracker.xlsx (240 KB):\nF12-North | 88.4',
          attachmentsRead: 1,
        },
        'Resolve Context': { profileSource: 'file' },
      }),
      $json: model,
      $itemIndex: 0,
    });
    expect(out.json.attachmentText).toContain('F12-North | 88.4');
    expect(out.json.attachmentsRead).toBe(1);
  });

  // The production build has no Attach Findings node at all, so the same body
  // has to work when nothing upstream ever set these fields.
  it('reports no attachment text when the build does not read attachments', () => {
    const out = run(applyClassificationCode, globals(prepared));
    expect(out.json.attachmentText).toBe('');
    expect(out.json.attachmentsRead).toBe(0);
  });

  it('drafts for the mailbox drafting was authorised for', () => {
    const out = run(applyClassificationCode, globals(prepared));
    expect(out.json.requiresDraft).toBe(true);
    expect(out.json.guardFlags).toBe('');
  });

  it('accepts the mailbox on cc, in any case', () => {
    const out = run(applyClassificationCode, globals({
      ...prepared,
      toRecipients: ['kdunlap@partnersperform.com'],
      ccRecipients: ['GTutor@PartnersPerform.com'],
    }));
    expect(out.json.requiresDraft).toBe(true);
  });

  // The credential-swap case: the agent is suddenly reading a different
  // mailbox. It must keep classifying, and stop writing.
  it('stops drafting when the mailbox is not the authorised one', () => {
    const out = run(applyClassificationCode, globals({
      ...prepared,
      toRecipients: ['ydawson@partnersperform.com'],
      ccRecipients: [],
    }));
    expect(out.json.requiresDraft).toBe(false);
    expect(out.json.guardFlags).toBe('mailbox_not_addressed');
    expect(out.json.label).toBe('opportunity');
  });

  it('still applies the pricing guard from the whole thread', () => {
    const out = run(applyClassificationCode, globals({
      ...prepared,
      cleanText: "Let's put together what it would take.",
      threadText: "Let's put together what it would take.\n\nCan you provide an estimate of timeline and cost?",
    }));
    expect(out.json.label).toBe('flag_for_yancey');
    expect(out.json.requiresDraft).toBe(false);
    expect(out.json.guardFlags).toContain('pricing_guard');
  });
});

/**
 * The same node body, configured exactly as production is: Yancey's rules of
 * 2026-08-28. This is the layer that actually ships — the library is tested
 * separately, but a config key spelled wrong in the node would pass there and
 * fail here.
 */
describe("WF-A Apply Classification under Yancey's rules", () => {
  const cfg = {
    confidenceMin: 0.6,
    followupDays: 3,
    alwaysFlagSenders: ['ccoleman@partnersperform.com'],
    internalDomains: ['partnersperform.com'],
    draftingMailbox: 'YDawson@partnersperform.com',
  };
  const base = {
    messageId: 'M1', messageIdEnc: 'M1', conversationId: 'C1', internetMessageId: '<x>',
    receivedAt: '2026-08-27T09:00:00Z',
    subject: 'RE: Vector CI', webLink: 'https://o/x', hasAttachments: false, categories: [],
    cleanText: 'Can you pull the workbook together before Thursday?',
    threadText: 'Can you pull the workbook together before Thursday?',
    toRecipients: ['YDawson@partnersperform.com'], ccRecipients: [],
  };
  const model = {
    label: 'opportunity', confidence: 0.95, reasoning: 'live work',
    requires_draft: true, urgency: 'this_week', asks_of_yancey: ['send workbook'],
  };

  /** @param {any} prep @param {any} [modelOut] */
  const runWith = (prep, modelOut) =>
    run(applyClassificationCode, {
      $: nodeRef({ Config: cfg, 'Prepare Email': prep, 'Resolve Context': { profileSource: 'file' } }),
      $json: modelOut || model,
      $itemIndex: 0,
    });

  it('drafts for a colleague', () => {
    const out = runWith({ ...base, senderEmail: 'scorkran@partnersperform.com', senderName: 'Shawn Corkran' });
    expect(out.json.label).toBe('opportunity');
    expect(out.json.requiresDraft).toBe(true);
    expect(out.json.categories).toContain('AI-Action');
  });

  it('never drafts for Chuck Coleman', () => {
    const out = runWith({ ...base, senderEmail: 'Ccoleman@partnersperform.com', senderName: 'Charles Coleman' });
    expect(out.json.label).toBe('flag_for_yancey');
    expect(out.json.requiresDraft).toBe(false);
    expect(out.json.guardFlags).toContain('always_flag_sender');
    expect(out.json.categories).toContain('AI-Review');
  });

  it('flags an outsider rather than replying to them', () => {
    const out = runWith({ ...base, senderEmail: 'Jessica.Doucette@sce.com', senderName: 'Jessica Doucette' });
    expect(out.json.label).toBe('flag_for_yancey');
    expect(out.json.requiresDraft).toBe(false);
    expect(out.json.guardFlags).toContain('external_sender');
  });

  // The rule he gave is "flag without drafting", not "escalate everything
  // external". Five of his junk fixtures are outside senders; promoting those
  // would put PayPal receipts in the bucket he reads at 07:30.
  it('leaves an outside newsletter as noise', () => {
    const out = runWith(
      { ...base, senderEmail: 'service@paypal.com', senderName: 'PayPal' },
      { label: 'noise', confidence: 0.97, reasoning: 'receipt', requires_draft: false, urgency: 'fyi', asks_of_yancey: [] },
    );
    expect(out.json.label).toBe('noise');
    expect(out.json.status).toBe('resolved');
    expect(out.json.guardFlags).toBe('');
  });

  it('does not resurrect a junk invoice into the review bucket', () => {
    const out = runWith(
      {
        ...base,
        senderEmail: 'service@paypal.com',
        subject: 'Invoice from JEREMIAH CHRISTIAN PLACENTE ROSALES (0028)',
        cleanText: 'You sent a payment request for a contractor invoice.',
        threadText: 'You sent a payment request for a contractor invoice.',
      },
      { label: 'noise', confidence: 0.99, reasoning: 'contractor invoice', requires_draft: false, urgency: 'fyi', asks_of_yancey: [] },
    );
    expect(out.json.label).toBe('noise');
    expect(out.json.guardFlags).toBe('');
    expect(out.json.categories).not.toContain('AI-Review');
  });

  it('names the category that held a colleague message back', () => {
    const out = runWith({
      ...base,
      senderEmail: 'dyoes@partnersperform.com',
      // "change order" is now a commercial term and would be held by the
      // pricing guard first, so this uses a purely internal approval ask.
      cleanText: 'Can you approve the revised T-Week plan before Thursday?',
      threadText: 'Can you approve the revised T-Week plan before Thursday?',
    });
    expect(out.json.label).toBe('flag_for_yancey');
    expect(out.json.guardFlags).toContain('sensitivity:approval');
  });

  it('still drafts consulting mail that only sounds like a flagged category', () => {
    const out = runWith({
      ...base,
      senderEmail: 'dyoes@partnersperform.com',
      cleanText: 'Sustain benefits through performance monitoring. Reducing operating expense improves cash flow.',
      threadText: 'Sustain benefits through performance monitoring. Reducing operating expense improves cash flow.',
    });
    expect(out.json.label).toBe('opportunity');
    expect(out.json.requiresDraft).toBe(true);
  });
});

/**
 * The WF-F apply node, run rather than read.
 *
 * On 3 September Yancey created a folder, moved two documents into it, dropped
 * eight more into the library and pressed Execute. Three things went wrong in
 * this one node, and every test below is one of them:
 *
 *   the URL ran past the 260-character ceiling ASP.NET enforces on a path
 *   a file wanting a new name AND a new folder produced two conflicting calls
 *   eight documents the model had just named came back "not on the approved
 *     list, so this run leaves it alone"
 */
const organiserApplyCode = await codeNode(HELPERS['wf-f/Apply Proposals'], ORGANISER_APPLY);

describe('WF-F Apply Proposals node body', () => {
  const HOST = 'https://upperformpartners.sharepoint.com';
  const SITE = '/sites/TaskManagement';
  const LIB = SITE + '/Shared Documents/Power BI Task Management';
  const KAREN = LIB + '/PD alignment Excel and Power BI - Karen';
  const REFS = LIB + '/Resources and References';

  /** @param {Partial<any>} over @returns {any} */
  const config = (over) => ({
    host: HOST,
    sitePath: SITE,
    libraryTitle: 'Documents',
    applyRenames: true,
    applyMoves: false,
    applyStructure: false,
    applyFolderRenames: false,
    maxChangesPerRun: 20,
    maxNewFolders: 6,
    minPerFolder: 2,
    renameReports: true,
    approvedPlan: [],
    assumeStatus: 'Draft',
    revert: false,
    maxItems: 5000,
    maxNamed: 30,
    ...over,
  });

  /** The two documents Yancey moved into the folder he had just made. */
  const KAREN_FILES = [
    { name: 'Performance Management System.xlsx', path: KAREN + '/Performance Management System.xlsx', modified: '2026-07-09T10:00:00Z' },
    { name: 'CAIDI deliverables .docx', path: KAREN + '/CAIDI deliverables .docx', modified: '2026-07-09T10:00:00Z' },
  ];

  /**
   * @param {{ cfg?: any, files?: any[], folders?: string[], rows?: any[], proposals?: any[], folderProposals?: any[] }} setup
   * @returns {any[]}
   */
  const apply = (setup) => {
    const files = setup.files || KAREN_FILES;
    const rows =
      setup.rows ||
      files.map((f) => ({
        current: f.name,
        path: f.path,
        proposed: '',
        blocked: '',
        compliant: false,
        assumptions: [],
        placement: 'settled',
        placementWhy: '',
        destination: '',
      }));
    const answer = JSON.stringify({
      proposals: setup.proposals || [],
      folders: setup.folderProposals || [],
    });
    const folders = setup.folders || [LIB, KAREN, REFS];
    return run(organiserApplyCode, {
      $: nodeRef({
        Config: setup.cfg || config({}),
        'List The Library': {
          value: [
            ...files.map((/** @type {any} */ f) => ({ FileLeafRef: f.name, FileRef: f.path })),
            ...folders.map((/** @type {any} */ p) => ({ FileLeafRef: p, FileRef: p, FSObjType: '1' })),
          ],
        },
        'Review The Library': {
          library: 'Documents',
          site: SITE,
          files,
          folders,
          rows,
          dependencies: { byFile: {}, unreadable: [], read: [] },
        },
      }),
      $input: { first: () => ({ json: { output: [{ content: [{ text: answer }] }] } }) },
      Buffer,
    }).map((/** @type {any} */ i) => i.json);
  };

  /** @param {any[]} out @param {string} name @returns {any} */
  const call = (out, name) => out.find((/** @type {any} */ o) => o.oldName === name);

  /**
   * The literal failure. "The length of the URL for this request exceeds the
   * configured maxUrlLength" — moveto carries the path twice, and the folder
   * Yancey named pushed it to about 385 characters.
   */
  it('keeps the request path under the ceiling that failed in the meeting', () => {
    const out = apply({
      proposals: [
        { file: 'Performance Management System.xlsx', project: '', document: 'Performance-Management-System-Workbook', status: '', date: '2026-07-09', version: '', category: 'workbook', folder: '', why: 'x' },
      ],
    });
    const url = String(call(out, 'Performance Management System.xlsx').changeUrl);
    expect(url.slice(0, url.indexOf('?')).length).toBeLessThan(260);
    expect(url).toContain("@old='");
    expect(url).toContain('flags=0');
  });

  /** Two calls from one path: the first worked, the second could not find the file. */
  it('renames and moves a document in a single call', () => {
    const out = apply({
      cfg: config({ applyMoves: true }),
      rows: [
        {
          current: 'CAIDI deliverables .docx',
          path: LIB + '/CAIDI deliverables .docx',
          proposed: '',
          blocked: '',
          compliant: false,
          assumptions: [],
          placement: 'move',
          placementWhy: 'it belongs with the references',
          destination: REFS + '/CAIDI deliverables .docx',
        },
      ],
      files: [{ name: 'CAIDI deliverables .docx', path: LIB + '/CAIDI deliverables .docx', modified: '2026-07-09T10:00:00Z' }],
      proposals: [
        { file: 'CAIDI deliverables .docx', project: 'NIPSCO', document: 'CAIDI-Deliverables', status: '', date: '2026-07-09', version: '', category: 'deliverables', folder: '', why: 'x' },
      ],
    });
    expect(out.length).toBe(1);
    expect(out[0].kind).toBe('rename and move');
    expect(out[0].newPath).toBe(REFS + '/NIPSCO_CAIDI-Deliverables_2026-07-09_Draft_v0.1.docx');
  });

  /**
   * Yancey dropped eight documents in, the model named every one, and all eight
   * came back refused. A plan gates what is applied, never what is proposed.
   */
  it('proposes a name for a new document a pinned plan has never seen', () => {
    const out = apply({
      cfg: config({
        applyRenames: true,
        approvedPlan: [{ from: 'Performance Management System.xlsx', to: 'NIPSCO_Performance-Management-System_2026-07-09_Draft_v0.1.xlsx' }],
      }),
      proposals: [
        { file: 'Performance Management System.xlsx', project: 'NIPSCO', document: 'Performance-Management-System', status: '', date: '2026-07-09', version: '', category: 'workbook', folder: '', why: 'x' },
        { file: 'CAIDI deliverables .docx', project: 'NIPSCO', document: 'CAIDI-Deliverables', status: '', date: '2026-07-09', version: '', category: 'deliverables', folder: '', why: 'x' },
      ],
    });
    const summary = out[0].summary;
    const fresh = summary.rows.find((/** @type {any} */ r) => r.current === 'CAIDI deliverables .docx');
    expect(fresh.proposed, 'the model named it, so the name must survive').toBeTruthy();
    expect(fresh.approved).toBe(false);
    expect(fresh.blocked).toBe('');
    expect(summary.awaiting).toBe(1);
    // ...and only the approved one is actually sent.
    expect(out.length).toBe(1);
    expect(out[0].oldName).toBe('Performance Management System.xlsx');
  });

  /**
   * The folder a model suggests is no more deterministic than the name it
   * writes. Runs 10264 and 10285, an hour apart, grouped the same two workbooks
   * into a new Dashboards folder and then did not — so it was created once and
   * refused the next time for holding one document. A plan line carrying only a
   * destination pins the filing, and the folder it names is created with it.
   */
  it('creates the folder a pinned destination has nowhere to land in', () => {
    const dash = REFS + '/Dashboards';
    const out = apply({
      cfg: config({
        applyRenames: false,
        applyMoves: true,
        applyStructure: true,
        approvedPlan: [
          { from: 'Performance Management System.xlsx', into: dash },
          { from: 'CAIDI deliverables .docx', into: dash },
        ],
      }),
      folders: [LIB, KAREN, REFS],
      proposals: [],
    });
    expect(out[0].kind).toBe('create folder');
    expect(out[0].newPath).toBe(dash);
    expect(out.slice(1).map((/** @type {any} */ o) => o.newPath)).toEqual([
      dash + '/Performance Management System.xlsx',
      dash + '/CAIDI deliverables .docx',
    ]);
  });

  it('does not create it twice when the library already has it', () => {
    const dash = REFS + '/Dashboards';
    const out = apply({
      cfg: config({
        applyRenames: false,
        applyMoves: true,
        applyStructure: true,
        approvedPlan: [{ from: 'Performance Management System.xlsx', into: dash }],
      }),
      folders: [LIB, KAREN, REFS, dash],
      proposals: [],
    });
    expect(out.some((/** @type {any} */ o) => o.kind === 'create folder')).toBe(false);
    expect(out[0].newPath).toBe(dash + '/Performance Management System.xlsx');
  });

  it('applies nothing at all with every switch off', () => {
    const out = apply({
      cfg: config({ applyRenames: false }),
      proposals: [
        { file: 'CAIDI deliverables .docx', project: 'NIPSCO', document: 'CAIDI-Deliverables', status: '', date: '2026-07-09', version: '', category: 'x', folder: '', why: 'x' },
      ],
    });
    expect(out.length).toBe(1);
    expect(out[0].nothingToDo).toBe(true);
    expect(out[0].proposed).toBe(1);
  });
});

/**
 * The folder structure the agent invents. Yancey asked for it in as many words
 * — *"Can it think and come up with its own folder structure?"* — and then
 * bounded it: *"I wanted to give it freedom a little bit."*
 */
describe('WF-F structure, through the node that applies it', () => {
  const HOST = 'https://upperformpartners.sharepoint.com';
  const SITE = '/sites/TaskManagement';
  const LIB = SITE + '/Shared Documents/Power BI Task Management';
  const REFS = LIB + '/Resources and References';

  const FILES = ['a.xlsx', 'b.xlsx', 'c.docx'].map((n) => ({
    name: n,
    path: REFS + '/' + n,
    modified: '2026-07-09T10:00:00Z',
  }));

  /** @param {Partial<any>} over @returns {any[]} */
  const apply = (over) =>
    run(organiserApplyCode, {
      $: nodeRef({
        Config: {
          host: HOST,
          sitePath: SITE,
          applyRenames: false,
          applyMoves: true,
          applyStructure: true,
          applyFolderRenames: false,
          maxChangesPerRun: 20,
          maxNewFolders: 6,
          minPerFolder: 2,
          renameReports: true,
          approvedPlan: [],
          assumeStatus: 'Draft',
          maxItems: 5000,
          maxNamed: 30,
          ...over,
        },
        'List The Library': {
          value: [
            ...FILES.map((f) => ({ FileLeafRef: f.name, FileRef: f.path })),
            ...[LIB, REFS].map((p) => ({ FileLeafRef: p, FileRef: p, FSObjType: '1' })),
          ],
        },
        'Review The Library': {
          library: 'Documents',
          site: SITE,
          files: FILES,
          folders: [LIB, REFS],
          rows: FILES.map((f) => ({
            current: f.name,
            path: f.path,
            proposed: '',
            blocked: '',
            compliant: true,
            assumptions: [],
            placement: 'settled',
            placementWhy: '',
            destination: '',
          })),
          dependencies: { byFile: {}, unreadable: [], read: [] },
        },
      }),
      $input: {
        first: () => ({
          json: {
            output: [
              {
                content: [
                  {
                    text: JSON.stringify({
                      folders: [{ from: REFS, to: 'Reference Library', why: 'says what it holds' }],
                      proposals: FILES.map((f) => ({
                        file: f.name,
                        project: '',
                        document: '',
                        status: '',
                        date: '',
                        version: '',
                        category: '',
                        folder: f.name === 'c.docx' ? 'Guides' : 'Dashboards',
                        why: 'x',
                      })),
                    }),
                  },
                ],
              },
            ],
          },
        }),
      },
      Buffer,
    }).map((/** @type {any} */ i) => i.json);

  it('creates the folder before anything moves into it', () => {
    const out = apply({});
    expect(out[0].kind).toBe('create folder');
    expect(out[0].newPath).toBe(REFS + '/Dashboards');
    expect(out.slice(1).every((/** @type {any} */ o) => o.kind === 'move')).toBe(true);
    expect(out.slice(1).map((/** @type {any} */ o) => o.newPath)).toEqual([
      REFS + '/Dashboards/a.xlsx',
      REFS + '/Dashboards/b.xlsx',
    ]);
  });

  /** A folder holding one file is the pile it came from with an extra click. */
  it('refuses the folder only one document was put in', () => {
    const out = apply({});
    const refusals = out[0].summary.folderRefusals;
    expect(refusals.map((/** @type {any} */ r) => r.folder)).toContain('Guides');
    expect(refusals[0].because).toContain('the pile it came from');
  });

  /** Everything inside a folder moves with it, so this is a third promise. */
  it('renames a folder last, and only when that switch is armed', () => {
    expect(apply({}).some((/** @type {any} */ o) => o.kind === 'rename folder')).toBe(false);
    const armed = apply({ applyFolderRenames: true });
    expect(armed[armed.length - 1].kind).toBe('rename folder');
    expect(armed[armed.length - 1].newPath).toBe(LIB + '/Reference Library');
  });

  /**
   * Moving into a folder nobody is creating would 404. The structure is still
   * PROPOSED — that is what the review page is for — but nothing is sent.
   */
  it('does not file into a new folder when creating them is switched off', () => {
    const out = apply({ applyStructure: false });
    expect(out.length).toBe(1);
    expect(out[0].nothingToDo).toBe(true);
    expect(out[0].newFolders, 'still proposed, just not applied').toEqual([
      REFS + '/Dashboards',
    ]);
  });
});

/**
 * The things that make it work on a library other than this one: a listing that
 * did not reach the end, a cap that falls between a folder and its documents,
 * and two documents composed onto the same path.
 */
describe('WF-F Apply Proposals on a library it has not seen before', () => {
  const HOST = 'https://upperformpartners.sharepoint.com';
  const SITE = '/sites/TaskManagement';
  const LIB = SITE + '/Shared Documents/Power BI Task Management';
  const REFS = LIB + '/Resources and References';
  const DASH = REFS + '/Dashboards';

  /**
   * @param {{ cfg?: any, files?: any[], folders?: string[], rows?: any[], moreToRead?: boolean, proposals?: any[] }} setup
   * @returns {any[]}
   */
  const apply = (setup) =>
    run(organiserApplyCode, {
      $: nodeRef({
        Config: {
          host: HOST,
          sitePath: SITE,
          applyRenames: true,
          applyMoves: true,
          applyStructure: true,
          applyFolderRenames: false,
          maxChangesPerRun: 12,
          maxNewFolders: 6,
          minPerFolder: 2,
          renameReports: true,
          approvedPlan: [],
          assumeStatus: 'Draft',
          maxItems: 5000,
          maxNamed: 30,
          ...(setup.cfg || {}),
        },
        'List The Library': {
          value: [
            ...(setup.files || []).map((/** @type {any} */ f) => ({ FileLeafRef: f.name, FileRef: f.path })),
            ...(setup.folders || [LIB, REFS]).map((/** @type {any} */ p) => ({ FileLeafRef: p, FileRef: p, FSObjType: '1' })),
          ],
        },
        'Review The Library': {
          library: 'Documents',
          site: SITE,
          files: setup.files || [],
          folders: setup.folders || [LIB, REFS],
          rows: setup.rows || [],
          listed: (setup.files || []).length + (setup.folders || [LIB, REFS]).length,
          pages: 1,
          moreToRead: setup.moreToRead === true,
          dependencies: { byFile: {}, unreadable: [], read: [] },
        },
      }),
      $input: {
        first: () => ({
          json: { output: [{ content: [{ text: JSON.stringify({ proposals: setup.proposals || [], folders: [] }) }] }] },
        }),
      },
      Buffer,
    }).map((/** @type {any} */ i) => i.json);

  /** @param {string} name @param {string} folder @param {Partial<any>} [over] */
  const doc = (name, folder, over) => ({
    current: name,
    path: folder + '/' + name,
    proposed: '',
    blocked: '',
    compliant: false,
    assumptions: [],
    placement: 'settled',
    placementWhy: '',
    destination: '',
    ...over,
  });

  /**
   * A library bigger than one page is a library whose collisions are only
   * partly known. Reviewing a slice and calling it the library is how a rename
   * lands on a document this run never saw.
   */
  /**
   * The last page it read still offered another, so it stopped at the page cap
   * rather than at the end of the library.
   */
  it('applies nothing at all when it has not read the whole library', () => {
    const rows = [doc('a.xlsx', REFS, { placement: 'move', destination: DASH + '/a.xlsx' })];
    const out = apply({ rows, files: [{ name: 'a.xlsx', path: REFS + '/a.xlsx' }], moreToRead: true });
    expect(out.length).toBe(1);
    expect(out[0].nothingToDo).toBe(true);
    expect(out[0].truncated).toBe(true);
    expect(out[0].planNote).toContain('NOTHING WAS APPLIED');
    expect(out[0].planNote).toContain('Raise maxPages');
  });

  it('applies the same plan once it has', () => {
    const rows = [doc('a.xlsx', REFS, { placement: 'move', destination: DASH + '/a.xlsx' })];
    const out = apply({ rows, files: [{ name: 'a.xlsx', path: REFS + '/a.xlsx' }], moreToRead: false });
    expect(out[0].nothingToDo).toBe(false);
    expect(out[0].summary.truncated).toBe(false);
  });

  /**
   * A library of exactly maxItems documents read in one page is complete, and
   * the old count-based guess called it partial and stalled for ever.
   */
  it('does not call a full page a partial library', () => {
    const rows = [doc('a.xlsx', REFS, { placement: 'move', destination: DASH + '/a.xlsx' })];
    const out = apply({ rows, files: [{ name: 'a.xlsx', path: REFS + '/a.xlsx' }], cfg: { maxItems: 3 } });
    expect(out[0].nothingToDo).toBe(false);
  });

  /** A cap that falls between a folder and its documents creates an empty one. */
  it('holds back a whole folder rather than half of it', () => {
    const rows = ['a.xlsx', 'b.xlsx', 'c.xlsx'].map((n) =>
      doc(n, REFS, { placement: 'move', destination: DASH + '/' + n }),
    );
    const files = rows.map((r) => ({ name: r.current, path: r.path }));
    const out = apply({ rows, files, cfg: { maxChangesPerRun: 2 } });
    expect(out.length).toBe(1);
    expect(out[0].nothingToDo).toBe(true);
    expect(out[0].heldBack).toEqual([{ group: 'folder:' + DASH, actions: 4, neverFits: true }]);
  });

  it('sends the folder and every document in it when they fit', () => {
    const rows = ['a.xlsx', 'b.xlsx'].map((n) =>
      doc(n, REFS, { placement: 'move', destination: DASH + '/' + n }),
    );
    const files = rows.map((r) => ({ name: r.current, path: r.path }));
    const out = apply({ rows, files, cfg: { maxChangesPerRun: 3 } });
    expect(out.map((/** @type {any} */ o) => o.kind)).toEqual(['create folder', 'move', 'move']);
    expect(out[0].summary.heldBack).toEqual([]);
    expect(out[0].newPath, 'the folder before anything goes into it').toBe(DASH);
  });

  /**
   * Two documents of the same name in different folders, both filed into one:
   * `flags=0` means the second call fails rather than overwriting, but planning
   * a call you know will fail is not a review.
   */
  it('refuses two documents that would land on one path, and sends neither', () => {
    const rows = [
      doc('Report.xlsx', LIB + '/One', { placement: 'move', destination: DASH + '/Report.xlsx' }),
      doc('Report.xlsx', LIB + '/Two', { placement: 'move', destination: DASH + '/Report.xlsx' }),
    ];
    const files = rows.map((r) => ({ name: r.current, path: r.path }));
    const out = apply({ rows, files, folders: [LIB, REFS, LIB + '/One', LIB + '/Two'] });
    expect(out.length).toBe(1);
    expect(out[0].nothingToDo).toBe(true);
    for (const row of out[0].rows) {
      expect(row.blocked).toContain('would land on the same path');
    }
  });

  /** And with nothing left to go in it, the folder is not created either. */
  it('does not create a folder whose documents were all refused', () => {
    const rows = [
      doc('Report.xlsx', LIB + '/One', { placement: 'move', destination: DASH + '/Report.xlsx' }),
      doc('Report.xlsx', LIB + '/Two', { placement: 'move', destination: DASH + '/Report.xlsx' }),
    ];
    const files = rows.map((r) => ({ name: r.current, path: r.path }));
    const out = apply({ rows, files, folders: [LIB, REFS, LIB + '/One', LIB + '/Two'] });
    expect(out[0].newFolders).toEqual([]);
  });

  /**
   * A folder two levels deeper than anything that exists needs both levels, and
   * they belong to the same decision as the document that goes in them. Grouped
   * separately, the cap could create the outer folder and move the document
   * while skipping the inner one — a move into a folder that is not there.
   */
  it('keeps a whole chain of new folders together with what goes in them', () => {
    const deep = REFS + '/Dashboards/Weekly';
    const rows = ['a.xlsx', 'b.xlsx'].map((n) =>
      doc(n, REFS, { placement: 'move', destination: deep + '/' + n }),
    );
    const files = rows.map((r) => ({ name: r.current, path: r.path }));
    const out = apply({ rows, files, cfg: { maxChangesPerRun: 12 } });
    expect(out.map((/** @type {any} */ o) => o.newPath)).toEqual([
      REFS + '/Dashboards',
      deep,
      deep + '/a.xlsx',
      deep + '/b.xlsx',
    ]);
    expect(new Set(out.map((/** @type {any} */ o) => o.summary && 1)).size).toBe(1);
  });

  it('holds the whole chain back rather than part of it', () => {
    const deep = REFS + '/Dashboards/Weekly';
    const rows = ['a.xlsx', 'b.xlsx'].map((n) =>
      doc(n, REFS, { placement: 'move', destination: deep + '/' + n }),
    );
    const files = rows.map((r) => ({ name: r.current, path: r.path }));
    const out = apply({ rows, files, cfg: { maxChangesPerRun: 3 } });
    expect(out.length).toBe(1);
    expect(out[0].nothingToDo).toBe(true);
    expect(out[0].heldBack).toEqual([{ group: 'folder:' + REFS + '/Dashboards', actions: 4, neverFits: true }]);
  });

  /**
   * A run can propose creating "Dashboards" and renaming another folder to
   * "Dashboards" in the same breath. Whichever ran second would fail, and the
   * rename check only knew about folders that existed before the run started.
   */
  it('will not rename a folder onto one it is about to create', () => {
    const rows = ['a.xlsx', 'b.xlsx'].map((n) =>
      doc(n, REFS, { placement: 'move', destination: DASH + '/' + n }),
    );
    const files = rows.map((r) => ({ name: r.current, path: r.path }));
    const out = run(organiserApplyCode, {
      $: nodeRef({
        Config: {
          host: HOST, sitePath: SITE, applyRenames: false, applyMoves: true,
          applyStructure: true, applyFolderRenames: true, maxChangesPerRun: 12,
          maxNewFolders: 6, minPerFolder: 2, renameReports: true, approvedPlan: [],
          assumeStatus: 'Draft', maxNamed: 30, maxItems: 5000,
        },
        'List The Library': {
          value: [
            ...files.map((/** @type {any} */ f) => ({ FileLeafRef: f.name, FileRef: f.path })),
            ...[LIB, REFS, REFS + '/Old'].map((p) => ({ FileLeafRef: p, FileRef: p, FSObjType: '1' })),
          ],
        },
        'Review The Library': {
          library: 'Documents', site: SITE, files, folders: [LIB, REFS, REFS + '/Old'],
          rows, truncated: false, listed: files.length,
          dependencies: { byFile: {}, unreadable: [], read: [] },
        },
      }),
      $input: {
        first: () => ({
          json: {
            output: [{ content: [{ text: JSON.stringify({
              proposals: [],
              folders: [{ from: REFS + '/Old', to: 'Dashboards', why: 'x' }],
            }) }] }],
          },
        }),
      },
      Buffer,
    }).map((/** @type {any} */ i) => i.json);
    expect(out.some((/** @type {any} */ o) => o.kind === 'rename folder')).toBe(false);
    const refused = out[0].summary.folderRenameRefusals;
    expect(refused[0]?.because).toContain('already sits beside it');
  });

  /**
   * Documents past maxNamed were never shown to the model, so nothing was
   * proposed for them — a fact about the run, not about them, and one that
   * reads as "nothing needed doing" unless it is said out loud.
   */
  it('counts the documents it never got to', () => {
    const rows = ['a.xlsx', 'b.xlsx', 'c.xlsx'].map((n) => doc(n, REFS));
    const files = rows.map((r) => ({ name: r.current, path: r.path }));
    const out = apply({ rows, files, cfg: { maxNamed: 2 } });
    expect(out[0].beyondModel).toBe(1);
  });

  it('counts none when it read everything it could', () => {
    const rows = ['a.xlsx', 'b.xlsx'].map((n) => doc(n, REFS));
    const files = rows.map((r) => ({ name: r.current, path: r.path }));
    expect(apply({ rows, files, cfg: { maxNamed: 30 } })[0].beyondModel).toBe(0);
  });

  /**
   * Two documents each taking the other's place cannot be done at all without
   * somewhere to put one of them first, and there is no such place here.
   */
  it('refuses a swap rather than sending half of it', () => {
    const rows = [
      doc('a.xlsx', LIB, { placement: 'move', destination: REFS + '/a.xlsx' }),
      doc('a.xlsx', REFS, { placement: 'move', destination: LIB + '/a.xlsx' }),
    ];
    const files = rows.map((r) => ({ name: r.current, path: r.path }));
    const out = apply({ rows, files, cfg: { applyRenames: false } });
    expect(out.length).toBe(1);
    expect(out[0].nothingToDo).toBe(true);
    expect(out[0].impossibleSwaps.length).toBe(2);
    expect(out[0].impossibleSwaps[0].because).toContain('trade places');
  });

  /** But a document taking a place another is vacating is fine, in that order. */
  it('sends the document that is leaving before the one taking its place', () => {
    const rows = [
      doc('a.xlsx', LIB, { placement: 'move', destination: REFS + '/a.xlsx' }),
      doc('a.xlsx', REFS, { placement: 'move', destination: LIB + '/Task 2/a.xlsx' }),
    ];
    const files = rows.map((r) => ({ name: r.current, path: r.path }));
    const out = apply({ rows, files, folders: [LIB, REFS, LIB + '/Task 2'], cfg: { applyRenames: false } });
    expect(out.map((/** @type {any} */ o) => o.oldPath)).toEqual([REFS + '/a.xlsx', LIB + '/a.xlsx']);
  });

  /**
   * Naming a destination as unreachable and then sending the move anyway is the
   * worst of both: the page says it cannot be done and the call goes out to fail.
   */
  it('cancels a move to a destination it cannot reach', () => {
    const rows = [doc('a.xlsx', REFS, { placement: 'move', destination: '/sites/Elsewhere/Somewhere/a.xlsx' })];
    const out = apply({ rows, files: [{ name: 'a.xlsx', path: REFS + '/a.xlsx' }] });
    expect(out.length).toBe(1);
    expect(out[0].nothingToDo).toBe(true);
    expect(out[0].unreachableFolders).toEqual(['/sites/Elsewhere/Somewhere']);
    expect(out[0].rows[0].placement).toBe('settled');
  });

  it('refuses a move onto a document already sitting there', () => {
    const rows = [doc('a.xlsx', LIB + '/One', { placement: 'move', destination: REFS + '/a.xlsx' })];
    const out = apply({
      rows,
      files: [
        { name: 'a.xlsx', path: LIB + '/One/a.xlsx' },
        { name: 'a.xlsx', path: REFS + '/a.xlsx' },
      ],
      folders: [LIB, REFS, LIB + '/One'],
    });
    expect(out[0].nothingToDo).toBe(true);
    expect(out[0].rows[0].blocked).toContain('already sits');
  });
});

/**
 * The node that turns a listing into a library.
 *
 * A library bigger than one request is read a page at a time, and the two ways
 * that goes wrong are opposite: reading only the first page and calling it the
 * library, or reading every page and still calling it partial. The nextLink on
 * the last page is what tells them apart, and it is exact rather than a guess
 * from how many items came back.
 */
const organiserReviewCode = await codeNode(HELPERS['wf-f/Review The Library'], ORGANISER_REVIEW);

describe('WF-F Review The Library node body', () => {
  const SITE = '/sites/TaskManagement';
  const LIB = SITE + '/Shared Documents/Power BI Task Management';

  /** @param {string} name */
  const file = (name) => ({ FileLeafRef: name, FileRef: LIB + '/' + name, Modified: '2026-04-15T10:00:00Z', FSObjType: 0 });
  /** @param {string} path */
  const folder = (path) => ({ FileLeafRef: path.slice(path.lastIndexOf('/') + 1), FileRef: path, FSObjType: '1' });

  /** @param {any[]} pages @returns {any} */
  const review = (pages) =>
    run(organiserReviewCode, {
      $: nodeRef({
        Config: { libraryTitle: 'Documents', sitePath: SITE, maxItems: 5000, assumeStatus: 'Draft', renameReports: true },
        'List The Library': pages,
      }),
      $input: { first: () => ({ json: { dependencies: { byFile: {}, unreadable: [], read: [] } } }) },
      Buffer,
    })[0].json;

  it('reads one page as the whole library', () => {
    const out = review([{ value: [file('a.xlsx'), folder(LIB + '/One')] }]);
    expect(out.files.map((/** @type {any} */ f) => f.name)).toEqual(['a.xlsx']);
    expect(out.folders).toEqual([LIB + '/One']);
    expect(out.listed).toBe(2);
    expect(out.moreToRead).toBe(false);
  });

  it('joins every page it was given into one library', () => {
    const out = review([
      { value: [file('a.xlsx')], 'odata.nextLink': 'https://next/1' },
      { value: [file('b.xlsx')], 'odata.nextLink': 'https://next/2' },
      { value: [file('c.xlsx')] },
    ]);
    expect(out.files.map((/** @type {any} */ f) => f.name)).toEqual(['a.xlsx', 'b.xlsx', 'c.xlsx']);
    expect(out.listed).toBe(3);
    expect(out.moreToRead, 'the last page offered no more').toBe(false);
  });

  /**
   * The page cap stopping short is the case that matters: everything read looks
   * exactly like a complete library unless the last page says otherwise.
   */
  it('says there is more when the last page it read offered more', () => {
    const out = review([
      { value: [file('a.xlsx')], 'odata.nextLink': 'https://next/1' },
      { value: [file('b.xlsx')], 'odata.nextLink': 'https://next/2' },
    ]);
    expect(out.listed).toBe(2);
    expect(out.moreToRead).toBe(true);
  });

  it('takes the OData 4 spelling of the same thing', () => {
    expect(review([{ value: [file('a.xlsx')], '@odata.nextLink': 'https://next/1' }]).moreToRead).toBe(true);
  });

  it('is not upset by a page that came back empty', () => {
    const out = review([{ value: [file('a.xlsx')] }, {}]);
    expect(out.listed).toBe(1);
    expect(out.moreToRead).toBe(false);
  });
});

/**
 * The Config node, which is where a run stops being what it was built to be and
 * becomes what it was asked to be.
 *
 * Everything downstream reads $('Config'), so this node is the single place a
 * form submission can reach - and the single place a sandbox refuses one. Both
 * halves are worth executing rather than assuming.
 */
const configMergeCode = await codeNode(HELPERS['wf-f/Config'], CONFIG_MERGE);

describe('WF-F Config node body', () => {
  const defaults = {
    host: 'https://upperformpartners.sharepoint.com',
    libraryTitle: 'Documents',
    maxChangesPerRun: 12,
    applyRenames: false,
    applyMoves: false,
    applyStructure: false,
    applyFolderRenames: false,
    applyProposed: false,
    approvedPlan: [{ from: 'baked.xlsx', to: 'Baked.xlsx' }],
    canArm: true,
  };

  /** @param {Record<string, any> | null} called @param {any} [base] */
  const merged = (called, base) => {
    const settings = base || defaults;
    const globals = {
      $: called === null
        ? nodeRef({})
        : nodeRef({ 'Called With': called }),
      $input: { first: () => ({ json: settings }), all: () => [{ json: settings }] },
    };
    return run(configMergeCode, globals)[0].json;
  };

  it('leaves a manual run exactly as it was built', () => {
    // No Called With node in the execution, which is what $() throwing means.
    const cfg = merged(null);
    expect(cfg.applyRenames).toBe(false);
    expect(cfg.host).toBe(defaults.host);
    expect(cfg.approvedPlan).toEqual(defaults.approvedPlan);
    expect(cfg.overrideRefusals).toEqual([]);
  });

  it('treats a call with no overrides as a manual run', () => {
    const cfg = merged({ overrides: '{}' });
    expect(cfg.applyRenames).toBe(false);
    expect(cfg.overrideRefusals).toEqual([]);
  });

  it('arms what the caller asked for when the deployment allows it', () => {
    const cfg = merged({ overrides: JSON.stringify({ applyRenames: true, approvedPlan: [{ from: 'a.xlsx', to: 'B.xlsx' }] }) });
    expect(cfg.applyRenames).toBe(true);
    expect(cfg.approvedPlan).toEqual([{ from: 'a.xlsx', to: 'B.xlsx' }]);
  });

  it('refuses to arm on a copy that may not write, and reports why', () => {
    const cfg = merged(
      { overrides: JSON.stringify({ applyRenames: true, applyMoves: true }) },
      Object.assign({}, defaults, { canArm: false })
    );
    expect(cfg.applyRenames).toBe(false);
    expect(cfg.applyMoves).toBe(false);
    expect(cfg.overrideRefusals).toHaveLength(2);
    expect(cfg.overrideRefusals[0]).toContain('not allowed to change the library');
  });

  it('will not let a caller move the run to a different library', () => {
    const cfg = merged({ overrides: JSON.stringify({ host: 'https://evil.example.com', libraryTitle: 'Other' }) });
    expect(cfg.host).toBe(defaults.host);
    expect(cfg.libraryTitle).toBe('Documents');
    expect(cfg.overrideRefusals).toHaveLength(2);
  });

  it('survives overrides that are not JSON at all', () => {
    const cfg = merged({ overrides: '{not json' });
    expect(cfg.applyRenames).toBe(false);
    expect(cfg.overrideRefusals[0]).toContain('not readable');
  });
});

/**
 * WF-G's two Code nodes — the whole of the web endpoint's logic.
 *
 * These matter more than most node bodies because what reaches them comes from
 * a browser. The decisions are in src/selfserve.js and tested there; what is
 * tested here is that a hostile or malformed request cannot talk this node into
 * asking WF-F for something it should not.
 */
const decideCode = await codeNode(HELPERS['wf-g/Decide'], DECIDE);
const answerCode = await codeNode(HELPERS['wf-g/Answer'], ANSWER);

describe('WF-G web endpoint node bodies', () => {
  const LIB = '/sites/TaskManagement/Shared Documents/Power BI Task Management';

  /** @param {any} body */
  const decide = (body) =>
    run(decideCode, {
      $input: { first: () => ({ json: { body } }), all: () => [{ json: { body } }] },
    })[0].json;

  it('treats an empty body as a look that changes nothing', () => {
    const asked = decide({});
    expect(asked.applying).toBe(false);
    expect(JSON.parse(asked.overrides)).toEqual({
      applyRenames: false,
      applyMoves: false,
      applyStructure: false,
      applyFolderRenames: false,
      applyProposed: false,
      approvedPlan: [],
    });
  });

  it('turns chosen changes into a plan that binds', () => {
    const asked = decide({
      chosen: [
        { index: 1, label: '1. Rename', kind: 'rename', from: LIB + '/a.xlsx', to: 'B_2026-01-01_Draft_v0.1.xlsx' },
      ],
    });
    expect(asked.applying).toBe(true);
    const sent = JSON.parse(asked.overrides);
    expect(sent.applyRenames).toBe(true);
    expect(sent.applyProposed).toBe(false);
    expect(sent.approvedPlan).toEqual([{ from: LIB + '/a.xlsx', to: 'B_2026-01-01_Draft_v0.1.xlsx' }]);
  });

  it('agrees to create a folder only when a chosen filing needs one', () => {
    const needs = decide({ chosen: [{ index: 1, label: '1. File', kind: 'move', from: LIB + '/a.pdf', into: LIB + '/New', createsFolder: true }] });
    expect(JSON.parse(needs.overrides).applyStructure).toBe(true);
    const settled = decide({ chosen: [{ index: 1, label: '1. File', kind: 'move', from: LIB + '/a.pdf', into: LIB + '/Old', createsFolder: false }] });
    expect(JSON.parse(settled.overrides).applyStructure).toBe(false);
  });

  it('cannot be talked into a setting that is not a change', () => {
    // A caller sending switches directly gets none of them: only `chosen` is
    // read, and the switches are derived from what it holds.
    const asked = decide({ applyRenames: true, applyProposed: true, host: 'https://evil.example.com', chosen: [] });
    const sent = JSON.parse(asked.overrides);
    expect(sent.applyRenames).toBe(false);
    expect(sent.applyProposed).toBe(false);
    expect(sent.host).toBeUndefined();
  });

  it('passes settings through to the organiser', () => {
    const asked = decide({ chosen: [], settings: { renameReports: false, readContents: false, maxChangesPerRun: 4 } });
    const sent = JSON.parse(asked.overrides);
    expect(sent.renameReports).toBe(false);
    expect(sent.readContents).toBe(false);
    expect(sent.maxChangesPerRun).toBe(4);
  });

  it('carries the newer settings through as well', () => {
    const asked = decide({ chosen: [], settings: { maxNewFolders: 3, minPerFolder: 4, assumeStatus: 'Review' } });
    const sent = JSON.parse(asked.overrides);
    expect(sent.maxNewFolders).toBe(3);
    expect(sent.minPerFolder).toBe(4);
    expect(sent.assumeStatus).toBe('Review');
  });

  it('will not let settings arm a run behind the ticks', () => {
    // The ticks decide what a run does. A caller hiding applyRenames in its
    // settings has it overwritten, because approvedFrom always writes all five.
    const asked = decide({ chosen: [], settings: { applyRenames: true, applyMoves: true, applyProposed: true } });
    const sent = JSON.parse(asked.overrides);
    expect(sent.applyRenames).toBe(false);
    expect(sent.applyMoves).toBe(false);
    expect(sent.applyProposed).toBe(false);
  });

  it('keeps the ticks when settings arrive alongside them', () => {
    const asked = decide({
      chosen: [{ index: 1, label: '1. Rename', kind: 'rename', from: LIB + '/a.xlsx', to: 'B_2026-01-01_Draft_v0.1.xlsx' }],
      settings: { maxChangesPerRun: 2 },
    });
    const sent = JSON.parse(asked.overrides);
    expect(sent.applyRenames).toBe(true);
    expect(sent.maxChangesPerRun).toBe(2);
    expect(sent.approvedPlan).toHaveLength(1);
  });

  it('ignores settings that are not an object', () => {
    expect(JSON.parse(decide({ chosen: [], settings: 'nope' }).overrides).renameReports).toBeUndefined();
    expect(JSON.parse(decide({ chosen: [], settings: [1, 2] }).overrides).renameReports).toBeUndefined();
  });

  it('counts entries it could not use rather than dropping them silently', () => {
    const asked = decide({ chosen: [{ label: 'nonsense' }, { nothing: true }] });
    expect(asked.chosenCount).toBe(0);
    expect(asked.unusable).toBeGreaterThan(0);
  });

  it('survives a body that is not an object at all', () => {
    expect(decide(null).applying).toBe(false);
    expect(decide([]).applying).toBe(false);
    expect(decide('nope').applying).toBe(false);
  });

  const review = {
    library: 'Documents',
    scanned: 27,
    compliant: 25,
    readContents: true,
    truncated: false,
    pages: 1,
    canArm: true,
    newFolders: [LIB + '/Reliability Reviews'],
    toRename: [{ file: LIB + '/DIMP_Dashboard.xlsx', becomes: '2026-08-14_DIMP_Upperform_v0.1_Draft.xlsx' }],
    toMove: [{ file: LIB + '/Loose.pdf', into: LIB + '/Reliability Reviews' }],
    folderRefusals: [
      { name: 'Dashboards', because: 'it sits inside a numbered stage folder' },
      { name: 'Task 1', because: 'it sits inside a numbered stage folder' },
      { name: 'Lonely', because: 'only one document was put in it' },
    ],
    applied: [],
    failed: [],
  };

  /** @param {any} summary @param {any} [asked] */
  const answer = (summary, asked) =>
    run(answerCode, {
      $: nodeRef({ Decide: asked || { applying: false, unusable: 0 } }),
      $input: { first: () => ({ json: summary }), all: () => [{ json: summary }] },
    })[0].json;

  it('gives the page every change it could make', () => {
    const said = answer(review);
    expect(said.ok).toBe(true);
    expect(said.offered).toHaveLength(3);
    expect(said.offered.map((/** @type {any} */ o) => o.kind)).toEqual(['folder', 'rename', 'move']);
  });

  it('says one refusal once and names what it applies to', () => {
    const said = answer(review);
    expect(said.folderRefusals).toHaveLength(2);
    expect(said.folderRefusals[0].names).toEqual(['Dashboards', 'Task 1']);
    expect(said.folderRefusals[1].names).toEqual(['Lonely']);
  });

  it('passes on what the run could not see', () => {
    const said = answer(Object.assign({}, review, { truncated: true, pages: 12, beyondModel: 8 }));
    expect(said.truncated).toBe(true);
    expect(said.pages).toBe(12);
    expect(said.beyondModel).toBe(8);
  });

  it('hands the page the whole configuration, so nobody opens n8n to read it', () => {
    const said = answer(Object.assign({}, review, {
      configuration: {
        host: 'https://upperformpartners.sharepoint.com',
        sitePath: '/sites/TaskManagement',
        libraryTitle: 'Documents',
        canArm: false,
        maxNewFolders: 6,
        minPerFolder: 2,
        assumeStatus: 'Draft',
        modelNaming: 'gpt-5.6-luna',
        pinnedPlan: 15,
      },
    }));
    expect(said.config.libraryTitle).toBe('Documents');
    expect(said.config.canArm).toBe(false);
    expect(said.config.pinnedPlan).toBe(15);
    expect(said.config.modelNaming).toBe('gpt-5.6-luna');
  });

  it('answers with an empty configuration rather than undefined', () => {
    expect(answer({}).config).toEqual({});
  });

  it('reports the settings the run actually used, not the ones asked for', () => {
    const said = answer(Object.assign({}, review, {
      renameReports: false,
      readContents: true,
      maxChangesPerRun: 4,
    }));
    expect(said.settings).toEqual({ renameReports: false, readContents: true, maxChangesPerRun: 4 });
  });

  it('passes on why a deployment did less than it was asked to', () => {
    const said = answer(Object.assign({}, review, {
      canArm: false,
      overrideRefusals: ['applyRenames was asked for, but this copy is not allowed to change the library'],
    }));
    expect(said.canArm).toBe(false);
    expect(said.overrideRefusals[0]).toContain('not allowed to change the library');
  });

  it('reports what an applying run actually did', () => {
    const said = answer(
      Object.assign({}, review, { applied: [{ what: 'renamed a.xlsx' }], failed: [{ what: 'b.xlsx', because: 'locked' }] }),
      { applying: true, unusable: 0 }
    );
    expect(said.applying).toBe(true);
    expect(said.applied).toHaveLength(1);
    expect(said.failed).toHaveLength(1);
  });

  it('answers at all when the organiser returned nothing useful', () => {
    const said = answer({});
    expect(said.ok).toBe(true);
    expect(said.scanned).toBe(0);
    expect(said.offered).toEqual([]);
  });
});


/**
 * The Summary node, which is the only thing the web page ever sees of a run.
 *
 * Its job is translation, and a translation that is subtly wrong in one field
 * is invisible until something downstream builds a path out of it.
 */
describe('WF-F Summary node body', () => {
  const LIB = '/sites/TaskManagement/Shared Documents/Power BI Task Management';
  const DEST = LIB + '/Reliability Reviews';

  /** @param {any[]} rows @param {any} [cfg] */
  const summarise = (rows, cfg) => {
    const settings = Object.assign(
      { canArm: false, readContents: true, renameReports: true, maxChangesPerRun: 12, overrideRefusals: [] },
      cfg || {}
    );
    const applyOut = { rows, library: 'Documents', scanned: rows.length, compliant: 0, proposed: 0, movesProposed: 0, awaiting: 0, blocked: 0, newFolders: [] };
    return run(ORGANISER_SUMMARY, {
      $: nodeRef({
        Config: settings,
        // Defaults is the built configuration, untouched by any caller.
        Defaults: { approvedPlan: [{ from: 'a.xlsx', to: 'B.xlsx' }], maxChangesPerRun: 12, maxNewFolders: 6, minPerFolder: 2 },
      }),
      $input: { first: () => ({ json: applyOut }), all: () => [{ json: applyOut }] },
    })[0].json;
  };

  it('reports the folder a document lands in, not the path it lands at', () => {
    // applyApprovedPlan appends the file name to `into`. Reporting the full
    // destination sent it back through the page and built
    // .../Reliability Reviews/Report.xlsx/Report.xlsx.
    const said = summarise([{
      current: 'Report.xlsx',
      path: LIB + '/Report.xlsx',
      placement: 'move',
      destination: DEST + '/Report.xlsx',
      placementWhy: 'grouped with others',
      compliant: false,
      proposed: '',
    }]);
    expect(said.toMove).toHaveLength(1);
    expect(said.toMove[0].into).toBe(DEST);
    expect(said.toMove[0].into).not.toContain('Report.xlsx');
  });

  it('survives a destination with no folder in it at all', () => {
    const said = summarise([{
      current: 'Report.xlsx', path: '/Report.xlsx', placement: 'move',
      destination: 'Report.xlsx', placementWhy: '', compliant: false, proposed: '',
    }]);
    expect(said.toMove[0].into).toBe('Report.xlsx');
  });

  it('reports the limits the deployment allows, not the ones this run used', () => {
    // A run where somebody chose 3 new folders reports 3. A page that reads its
    // ceiling off that offers 3 next time, then 2, then 1 - a choice narrowing
    // what may be chosen again. The limits come from Defaults instead.
    const said = summarise([], { maxChangesPerRun: 5, maxNewFolders: 3, minPerFolder: 4 });
    expect(said.configuration.maxNewFolders).toBe(3);
    expect(said.configuration.limits.maxNewFolders).toBe(6);
    expect(said.configuration.limits.maxChangesPerRun).toBe(12);
    expect(said.configuration.limits.minPerFolder).toBe(2);
  });

  it('reports the plan the deployment pins, not the one the caller sent', () => {
    const said = summarise([], { approvedPlan: [] });
    expect(said.configuration.pinnedPlan).toBe(1);
  });

  it('carries the whole configuration for the page to show', () => {
    const said = summarise([], { host: 'https://x.sharepoint.com', libraryTitle: 'Documents', modelNaming: 'gpt-5.6-luna', maxNewFolders: 6, minPerFolder: 2, assumeStatus: 'Draft' });
    expect(said.configuration.libraryTitle).toBe('Documents');
    expect(said.configuration.modelNaming).toBe('gpt-5.6-luna');
    expect(said.configuration.minPerFolder).toBe(2);
    expect(said.configuration.assumeStatus).toBe('Draft');
  });
});
