import { describe, expect, it } from 'vitest';
import { codeNode } from '../scripts/bundle.mjs';
import { HELPERS } from '../workflows/helpers-manifest.mjs';
import { VOICE_GUARD } from '../workflows/wf-b-draft.mjs';
import { APPLY_CLASSIFICATION } from '../workflows/wf-a-classify.mjs';
import { APPLY as ORGANISER_APPLY } from '../workflows/wf-f-organizer.mjs';

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
  return { first: () => ({ json: value }), all: () => [{ json: value }], item: { json: value } };
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
    return run(organiserApplyCode, {
      $: nodeRef({
        Config: setup.cfg || config({}),
        'Review The Library': {
          library: 'Documents',
          site: SITE,
          files,
          folders: setup.folders || [LIB, KAREN, REFS],
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
          ...over,
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
