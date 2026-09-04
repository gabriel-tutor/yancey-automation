import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

/**
 * The page Yancey opens.
 *
 * It had no tests until it had produced three bugs in one afternoon: a JSON
 * blob printed where a list of changes belonged, three controls wired to
 * nothing, and pickers that narrowed a little every time somebody chose less.
 * Two of the three were invisible to unit tests of the functions involved,
 * because the functions were right and the WIRING was not — so the wiring is
 * what most of this file is about.
 *
 * Everything here reads web/index.html, the artefact that actually ships, and
 * pulls the real functions out of it. A copy of them would only prove the copy
 * works.
 */
const PAGE = readFileSync(new URL('../web/index.html', import.meta.url), 'utf8');

/** The page's own script, without the surrounding markup. */
const SCRIPT = (() => {
  const at = PAGE.indexOf('<script>');
  const to = PAGE.indexOf('</script>', at);
  if (at === -1 || to === -1) throw new Error('web/index.html has no script');
  return PAGE.slice(at + '<script>'.length, to);
})();

/** Everything between `<details class="settings">` and the working panel. */
const PANEL = (() => {
  const at = PAGE.indexOf('<details class="settings"');
  const to = PAGE.indexOf('<!-- Working', at);
  return PAGE.slice(at, to);
})();

/**
 * Lift a top-level function or object literal out of the page by name.
 *
 * This is evaluated with `new Function`, which is worth being explicit about:
 * the only thing that ever reaches it is web/index.html, read off this repo's
 * own disk at test time, and nothing here ships. There is no untrusted input
 * and no interpolation of anything a caller supplies. The alternative — keeping
 * a second copy of these functions in the test — would prove the copy works and
 * say nothing about the page, which is the exact mistake that let three
 * controls ship unwired.
 *
 * @param {string} name
 * @param {'function' | 'var'} [kind]
 */
function source(name, kind) {
  const re = kind === 'var'
    ? new RegExp('\\n  var ' + name + ' = \\{[\\s\\S]*?\\n  \\};')
    : new RegExp('\\n  function ' + name + '\\([\\s\\S]*?\\n  \\}\\n');
  const found = SCRIPT.match(re);
  if (!found) throw new Error('web/index.html has no ' + (kind || 'function') + ' called ' + name);
  return found[0];
}

/* ------------------------------------------------------------- a small DOM */

/** @param {string} tag */
const node = (tag) => {
  /** @type {any} */
  const made = {
    tag,
    className: '',
    _text: '',
    /** @type {any[]} */ children: [],
    /** @param {any} child */
    appendChild(child) {
      made.children.push(child);
      return child;
    },
  };
  Object.defineProperty(made, 'textContent', {
    get: () => made._text,
    /** @param {any} value */
    set: (value) => {
      made._text = value == null ? '' : String(value);
      made.children = [];
    },
  });
  return made;
};

const documentStub = {
  createElement: node,
  /** @param {any} t */
  createTextNode: (t) => Object.assign(node('#text'), { _text: String(t), textContent: String(t) }),
};

/** @param {any} n @returns {string} */
const flatten = (n) => String(n._text || '') + (n.children || []).map(flatten).join('');

/** @param {string} tag @param {any} value @param {string} [className] */
const text = (tag, value, className) => {
  const made = node(tag);
  made.textContent = value == null ? '' : String(value);
  if (className) made.className = className;
  return made;
};

/** A localStorage that can be told to misbehave, since browsers do. */
function storage() {
  /** @type {Record<string, string>} */
  let held = {};
  let broken = false;
  return {
    raw: () => JSON.stringify(held),
    reset: () => {
      held = {};
      broken = false;
    },
    /** @param {boolean} on */
    breaks: (on) => {
      broken = on;
    },
    api: {
      /** @param {string} k @param {string} v */
      setItem(k, v) {
        if (broken) throw new Error('QuotaExceededError');
        held[k] = v;
      },
      /** @param {string} k */
      getItem(k) {
        if (broken) throw new Error('SecurityError');
        return k in held ? held[k] : null;
      },
      /** @param {string} k */
      removeItem(k) {
        if (broken) throw new Error('SecurityError');
        delete held[k];
      },
    },
  };
}

/* ---------------------------------------------------- the wiring invariants */

describe('the page is wired to itself', () => {
  const inMarkup = new Set([...PAGE.matchAll(/id="([a-zA-Z0-9-]+)"/g)].map((m) => m[1]));

  it('never reaches for an element that is not in the markup', () => {
    // el('typo') returns undefined and the next line throws, at the moment a
    // person clicks something. Cheap to prevent, and this is where.
    const asked = new Set([...SCRIPT.matchAll(/\bel\('([a-zA-Z0-9-]+)'\)/g)].map((m) => m[1]));
    const missing = [...asked].filter((id) => !inMarkup.has(id));
    expect(missing, 'the script reads elements that do not exist').toEqual([]);
  });

  it('gives every settings control a change listener', () => {
    // The bug this exists for: three controls were added to the panel and the
    // listener list was left at the three that came before them, so they
    // updated nothing and saved nothing. Both symptoms were downstream of one
    // missing line, and no test of the functions could have seen it.
    const controls = new Set([...PANEL.matchAll(/id="(set-[a-z]+)"/g)].map((m) => m[1]));
    const listed = SCRIPT.match(/\[([^\]]*'set-[^\]]*)\]\.forEach/);
    expect(listed, 'there is no list of settings controls to attach listeners to').toBeTruthy();
    const wired = new Set([...String(listed?.[1]).matchAll(/'(set-[a-z]+)'/g)].map((m) => m[1]));

    expect(controls.size).toBeGreaterThan(0);
    expect([...controls].filter((id) => !wired.has(id)), 'controls with no change listener').toEqual([]);
    expect([...wired].filter((id) => !controls.has(id)), 'listeners for controls that do not exist').toEqual([]);
  });

  it('saves whenever a control changes, not only after a run', () => {
    const handler = SCRIPT.match(/\[[^\]]*'set-[^\]]*\]\.forEach\([\s\S]{0,400}?\}\);/);
    expect(String(handler?.[0])).toContain('saveChoices()');
  });

  it('parses', () => {
    expect(() => new Function(SCRIPT)).not.toThrow();
  });
});

/* ------------------------------------------------------ what a change reads */

describe('rendering a change that has happened', () => {
  const { doneRow } = new Function(
    'document',
    'text',
    [source('DONE_WORD', 'var'), source('folderOf'), source('doneRow'), 'return { doneRow };'].join('\n')
  )(documentStub, text);

  const applied = {
    kind: 'rename and move',
    oldName: 'PD_0730_ReportOut_Template (1).xlsx',
    newName: 'Reliability-Report-Template_2026-09-04_Draft_v0.1.xlsx',
    newPath: '/sites/TaskManagement/Shared Documents/Power BI Task Management/Resources and References/Reliability-Report-Template_2026-09-04_Draft_v0.1.xlsx',
    assumptions: ['no project code was evidenced, so none was added', 'date taken from the file timestamp, not the document'],
  };

  it('never prints a JSON blob at somebody', () => {
    // The first production run reported both of its changes as raw JSON,
    // because the fallback looked for fields the apply side does not emit.
    for (const entry of [applied, { kind: 'move', oldName: 'a.pdf', newName: 'a.pdf', newPath: '/x/y/a.pdf' }, { kind: 'create folder', newName: 'Reviews', newPath: '/x/Reviews' }]) {
      const out = flatten(doneRow(entry, false));
      expect(out).not.toContain('{"');
      expect(out).not.toContain('"kind"');
    }
  });

  it('names the kind in the apply side\'s vocabulary, not the review page\'s', () => {
    // 'rename and move' here; 'rename+move' on the review page. One lookup for
    // both would miss quietly and print a kind nobody wrote.
    expect(flatten(doneRow(applied, false))).toContain('Renamed and filed');
    expect(flatten(doneRow({ kind: 'rename', oldName: 'a', newName: 'b' }, false))).toContain('Renamed');
    expect(flatten(doneRow({ kind: 'move', oldName: 'a', newName: 'a', newPath: '/x/a' }, false))).toContain('Filed');
    expect(flatten(doneRow({ kind: 'create folder', newName: 'X', newPath: '/x/X' }, false))).toContain('New folder');
    expect(flatten(doneRow({ kind: 'rename folder', oldName: 'a', newName: 'b' }, false))).toContain('Folder renamed');
  });

  it('shows the old name, the new name and the folder it landed in', () => {
    const out = flatten(doneRow(applied, false));
    expect(out).toContain('PD_0730_ReportOut_Template (1).xlsx');
    expect(out).toContain('Reliability-Report-Template_2026-09-04_Draft_v0.1.xlsx');
    expect(out).toContain('Resources and References');
    // The folder, not the path: the file name must not appear twice.
    expect(out.split('Reliability-Report-Template_2026-09-04_Draft_v0.1.xlsx')).toHaveLength(2);
  });

  it('says what it had to assume, which is what makes a name checkable', () => {
    const out = flatten(doneRow(applied, false));
    expect(out).toContain('Assumed');
    expect(out).toContain('date taken from the file timestamp, not the document');
  });

  it('gives a failure its reason and no assumptions', () => {
    const out = flatten(doneRow({ kind: 'rename', oldName: 'Locked.xlsx', newName: 'Tidy.xlsx', why: 'checked out by another user', assumptions: ['ignored'] }, true));
    expect(out).toContain('checked out by another user');
    expect(out).not.toContain('Assumed');
  });

  it('survives an entry with almost nothing on it', () => {
    expect(() => doneRow({}, false)).not.toThrow();
    expect(() => doneRow({ kind: 'rename' }, true)).not.toThrow();
    expect(flatten(doneRow('a plain sentence', false))).toContain('a plain sentence');
  });
});

/* ------------------------------------------------------------ the settings */

describe('the settings a person chose', () => {
  /** Build the settings half of the page against a fake set of controls. */
  function panel() {
    /** @type {Record<string, any>} */
    const controls = {
      'set-reports': { checked: true },
      'set-contents': { checked: true },
      'set-cap': { value: '12', options: /** @type {number[]} */ ([]) },
      'set-folders': { value: '6', options: /** @type {number[]} */ ([]) },
      'set-minper': { value: '2', options: /** @type {number[]} */ ([]) },
      'set-status': { value: 'Draft' },
    };
    for (const control of Object.values(controls)) {
      if (!('options' in control)) continue;
      const box = /** @type {any} */ (control);
      box.appendChild = (/** @type {any} */ opt) => box.options.push(Number(opt.value));
      Object.defineProperty(box, 'textContent', { set: () => { box.options = []; }, get: () => '' });
    }
    const store = storage();
    const api = new Function(
      'window',
      'el',
      'document',
      [
        "var CHOICES = 'organiser.settings.v1';",
        'var cap = 12;',
        'var bounds = { maxChangesPerRun: 12, maxNewFolders: 6, minPerFolder: 2 };',
        source('fillRange'),
        source('fillFloor'),
        source('settings'),
        source('saveChoices'),
        source('loadChoices'),
        source('useLimits'),
        source('applyChoices'),
        'return { settings, saveChoices, loadChoices, useLimits, applyChoices };',
      ].join('\n')
    )({ localStorage: store.api }, (/** @type {string} */ id) => controls[id], documentStub);
    return { api, controls, store };
  }

  const LIMITS = { maxChangesPerRun: 12, maxNewFolders: 6, minPerFolder: 2 };
  /** @param {any} c @param {string} id */
  const span = (c, id) => c[id].options[0] + '..' + c[id].options[c[id].options.length - 1];

  it('keeps what was chosen across a reload', () => {
    const { api, controls, store } = panel();
    api.useLimits(LIMITS);
    controls['set-reports'].checked = false;
    controls['set-cap'].value = '4';
    controls['set-status'].value = 'Review';
    api.saveChoices();

    // A reload: limits from the deployment, values from what was chosen.
    const fresh = panel();
    fresh.store.api.setItem('organiser.settings.v1', JSON.parse(store.raw())['organiser.settings.v1']);
    fresh.api.useLimits(LIMITS);
    fresh.api.applyChoices(fresh.api.loadChoices());
    expect(fresh.controls['set-reports'].checked).toBe(false);
    expect(fresh.controls['set-cap'].value).toBe('4');
    expect(fresh.controls['set-status'].value).toBe('Review');
  });

  it('does not let a choice narrow what may be chosen next time', () => {
    // Choose three folders, and the picker offered one to three ever after -
    // then two, then one. The ceiling was being read off the choice.
    const { api, controls } = panel();
    api.useLimits(LIMITS);
    controls['set-folders'].value = '3';
    controls['set-minper'].value = '5';
    api.saveChoices();

    api.useLimits(LIMITS);
    api.applyChoices(api.loadChoices());
    expect(span(controls, 'set-folders')).toBe('1..6');
    expect(span(controls, 'set-minper')).toBe('2..10');
    expect(controls['set-folders'].value).toBe('3');
    expect(controls['set-minper'].value).toBe('5');
  });

  it('clamps a choice the deployment would now refuse', () => {
    const { api, controls } = panel();
    api.useLimits(LIMITS);
    controls['set-cap'].value = '12';
    controls['set-minper'].value = '2';
    api.saveChoices();

    // The deployment has since tightened both ends.
    api.useLimits({ maxChangesPerRun: 6, maxNewFolders: 6, minPerFolder: 4 });
    api.applyChoices(api.loadChoices());
    expect(Number(controls['set-cap'].value)).toBeLessThanOrEqual(6);
    expect(Number(controls['set-minper'].value)).toBeGreaterThanOrEqual(4);
  });

  it('asks again rather than breaking when the browser will not store', () => {
    const { api, store } = panel();
    api.useLimits(LIMITS);
    store.breaks(true);
    expect(() => api.saveChoices()).not.toThrow();
    expect(api.loadChoices()).toBeNull();
    expect(() => api.applyChoices(null)).not.toThrow();
  });

  it('ignores stored rubbish', () => {
    const { api, store } = panel();
    store.api.setItem('organiser.settings.v1', '{not json');
    expect(api.loadChoices()).toBeNull();
    store.api.setItem('organiser.settings.v1', '"a string"');
    expect(api.loadChoices()).toBeNull();
  });
});

/* ------------------------------------------- what is kept about the client */

describe('what the browser is allowed to remember', () => {
  function memory() {
    const store = storage();
    const api = new Function(
      'window',
      [
        "var REMEMBERED = 'organiser.config.v1';",
        "var ARMING = ['applyRenames','applyMoves','applyStructure','applyFolderRenames','applyProposed'];",
        source('remember'),
        source('recall'),
        source('forget'),
        'return { remember, recall, forget };',
      ].join('\n')
    )({ localStorage: store.api });
    return { api, store };
  }

  const CONFIG = {
    libraryTitle: 'Documents',
    host: 'https://upperformpartners.sharepoint.com',
    sitePath: '/sites/TaskManagement',
    canArm: true,
    limits: { maxChangesPerRun: 12, maxNewFolders: 6, minPerFolder: 2 },
    pinnedPlan: 15,
    pinnedPlanLines: [{ from: 'TCO Pilot Program Oversight Dashboard Data.xlsx', to: 'TCO_x.xlsx', into: '' }],
    applyRenames: true,
    applyMoves: true,
    applyStructure: false,
    applyFolderRenames: false,
    applyProposed: false,
  };

  it('keeps the deployment facts a refresh needs', () => {
    const { api } = memory();
    api.remember(CONFIG);
    const held = api.recall();
    expect(held.config.libraryTitle).toBe('Documents');
    expect(held.config.limits.maxNewFolders).toBe(6);
    expect(held.config.pinnedPlan).toBe(15);
    expect(typeof held.at).toBe('number');
  });

  it('never keeps a client file name', () => {
    const { api, store } = memory();
    api.remember(CONFIG);
    expect(api.recall().config.pinnedPlanLines).toBeUndefined();
    expect(store.raw()).not.toContain('TCO Pilot');
  });

  it('never keeps per-run arming state', () => {
    // Kept, it would read as "this tool can do nothing" after a look and "this
    // is armed right now" after an apply. Neither describes the deployment.
    const { api, store } = memory();
    api.remember(CONFIG);
    for (const key of ['applyRenames', 'applyMoves', 'applyStructure', 'applyFolderRenames', 'applyProposed']) {
      expect(api.recall().config[key], key + ' should not be remembered').toBeUndefined();
    }
    expect(store.raw()).not.toContain('applyRenames');
  });

  it('is emptied on signing out, because the names are the client\'s', () => {
    const { api } = memory();
    api.remember(CONFIG);
    api.forget();
    expect(api.recall()).toBeNull();
  });

  it('degrades to no memory rather than a broken page', () => {
    const { api, store } = memory();
    store.api.setItem('organiser.config.v1', '{not json');
    expect(api.recall()).toBeNull();
    store.api.setItem('organiser.config.v1', JSON.stringify({ at: 1, config: {} }));
    expect(api.recall(), 'a config with no library is not worth showing').toBeNull();
    store.breaks(true);
    expect(api.recall()).toBeNull();
    expect(() => api.remember(CONFIG)).not.toThrow();
    expect(() => api.forget()).not.toThrow();
  });
});
