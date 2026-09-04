import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { COOKIE, issue } from '../web/api/_auth.js';

/**
 * The Vercel function the page talks to.
 *
 * It is deliberately thin — every decision lives in the workflow, where it is
 * tested — but thin is not the same as trivial, and it has already dropped a
 * field on the floor: the first version forwarded `chosen` and silently ignored
 * `settings`, so every toggle on the page did nothing and said nothing.
 *
 * It is also the gate. The page hiding a button is a page, not a guard; this is
 * what stands between the open internet and a workflow that renames a client's
 * documents, so most of what follows is about it saying no.
 */
const SECRET = 'a-secret-that-only-the-server-knows';
const ORGANISER = 'https://n8n.example.invalid/webhook/organise';

/** A response object that records what the handler did to it. */
function reply() {
  const sent = { status: 0, body: /** @type {any} */ (null), headers: /** @type {Record<string, any>} */ ({}) };
  const res = {
    /** @param {number} code */
    status(code) {
      sent.status = code;
      return res;
    },
    /** @param {any} body */
    json(body) {
      sent.body = body;
      return res;
    },
    /** @param {string} name @param {any} value */
    setHeader(name, value) {
      sent.headers[name] = value;
    },
    sent,
  };
  return res;
}

/** @param {{ method?: string, body?: any, cookie?: string }} [how] */
function request(how) {
  const it = how || {};
  return {
    method: it.method || 'POST',
    body: it.body === undefined ? {} : it.body,
    headers: it.cookie === undefined ? {} : { cookie: it.cookie },
  };
}

const signedIn = () => COOKIE + '=' + encodeURIComponent(issue(SECRET));

/**
 * Import the handler fresh.
 *
 * It reads its environment once, at module load, the way a serverless function
 * does — so a test that wants a different environment needs a different module
 * instance, not a different variable.
 *
 * @param {Record<string, string | undefined>} [env]
 */
async function load(env) {
  vi.resetModules();
  process.env.SESSION_SECRET = SECRET;
  process.env.N8N_ORGANISE_URL = ORGANISER;
  delete process.env.N8N_ORGANISE_SECRET;
  for (const [key, value] of Object.entries(env || {})) {
    if (value === undefined) delete process.env[key];
    else process.env[key] = value;
  }
  return (await import('../web/api/organise.js')).default;
}

/** @param {{ ok?: boolean, status?: number, text?: string }} [answer] */
function organiserAnswers(answer) {
  const it = answer || {};
  const calls = /** @type {any[]} */ ([]);
  const stub = vi.fn(async (/** @type {string} */ url, /** @type {any} */ init) => {
    calls.push({ url, init, body: JSON.parse(init.body) });
    return {
      ok: it.ok === undefined ? true : it.ok,
      status: it.status || 200,
      text: async () => (it.text === undefined ? JSON.stringify([{ json: { ok: true, scanned: 29 } }]) : it.text),
    };
  });
  vi.stubGlobal('fetch', stub);
  return calls;
}

const KEPT = { ...process.env };
beforeEach(() => vi.unstubAllGlobals());
afterEach(() => {
  vi.unstubAllGlobals();
  process.env = { ...KEPT };
});

describe('the gate on the endpoint', () => {
  it('turns away a request with no session, without touching the organiser', async () => {
    const calls = organiserAnswers();
    const handler = await load();
    const res = reply();
    await handler(request(), res);
    expect(res.sent.status).toBe(401);
    expect(res.sent.body.signedIn).toBe(false);
    expect(calls, 'an unauthenticated request must never reach SharePoint').toEqual([]);
  });

  it('turns away a forged cookie', async () => {
    const calls = organiserAnswers();
    const handler = await load();
    const res = reply();
    await handler(request({ cookie: COOKIE + '=eyJleHAiOjk5OTk5OTk5OTl9.deadbeef' }), res);
    expect(res.sent.status).toBe(401);
    expect(calls).toEqual([]);
  });

  it('turns away a session signed with a different secret', async () => {
    const calls = organiserAnswers();
    const handler = await load();
    const res = reply();
    await handler(request({ cookie: COOKIE + '=' + encodeURIComponent(issue('some-other-secret')) }), res);
    expect(res.sent.status).toBe(401);
    expect(calls).toEqual([]);
  });

  it('lets a real session through', async () => {
    const calls = organiserAnswers();
    const handler = await load();
    const res = reply();
    await handler(request({ cookie: signedIn() }), res);
    expect(res.sent.status).toBe(200);
    expect(calls).toHaveLength(1);
  });

  it('refuses anything but POST, and says which', async () => {
    const handler = await load();
    for (const method of ['GET', 'PUT', 'DELETE']) {
      const res = reply();
      await handler(request({ method }), res);
      expect(res.sent.status).toBe(405);
      expect(res.sent.headers.Allow).toBe('POST');
    }
  });

  it('fails closed when the deployment has nowhere to send it', async () => {
    const calls = organiserAnswers();
    const handler = await load({ N8N_ORGANISE_URL: undefined });
    const res = reply();
    await handler(request({ cookie: signedIn() }), res);
    expect(res.sent.status).toBe(500);
    expect(res.sent.body.error).toContain('N8N_ORGANISE_URL');
    expect(calls).toEqual([]);
  });

  it('never lets an answer be cached', async () => {
    organiserAnswers();
    const handler = await load();
    const res = reply();
    await handler(request({ cookie: signedIn() }), res);
    expect(res.sent.headers['Cache-Control']).toBe('no-store');
  });
});

describe('what it forwards', () => {
  it('sends the ticks and the settings, which it once dropped', async () => {
    const calls = organiserAnswers();
    const handler = await load();
    const chosen = [{ index: 1, label: '1. Rename', kind: 'rename', from: '/x/a.xlsx', to: 'b.xlsx' }];
    const settings = { renameReports: false, readContents: false, maxChangesPerRun: 4 };
    await handler(request({ cookie: signedIn(), body: { chosen, settings } }), reply());
    expect(calls[0].body.chosen).toEqual(chosen);
    expect(calls[0].body.settings).toEqual(settings);
  });

  it('sends empty rather than undefined when neither was given', async () => {
    const calls = organiserAnswers();
    const handler = await load();
    await handler(request({ cookie: signedIn(), body: {} }), reply());
    expect(calls[0].body).toEqual({ chosen: [], settings: {} });
  });

  it('refuses to forward a chosen or settings that is the wrong shape', async () => {
    const calls = organiserAnswers();
    const handler = await load();
    await handler(request({ cookie: signedIn(), body: { chosen: 'all of them', settings: [1, 2] } }), reply());
    expect(calls[0].body).toEqual({ chosen: [], settings: {} });
  });

  it('survives a body that is not an object at all', async () => {
    organiserAnswers();
    const handler = await load();
    for (const body of [null, undefined, 'nope', 42]) {
      const res = reply();
      await handler(request({ cookie: signedIn(), body }), res);
      expect(res.sent.status).toBe(200);
    }
  });

  it('carries the webhook key only when the deployment has one', async () => {
    let calls = organiserAnswers();
    let handler = await load();
    await handler(request({ cookie: signedIn() }), reply());
    expect(calls[0].init.headers['X-Organise-Key'], 'no key set, so no header').toBeUndefined();

    calls = organiserAnswers();
    handler = await load({ N8N_ORGANISE_SECRET: 'the-webhook-key' });
    await handler(request({ cookie: signedIn() }), reply());
    expect(calls[0].init.headers['X-Organise-Key']).toBe('the-webhook-key');
  });

  it('never puts the webhook key or the URL in what it answers with', async () => {
    organiserAnswers();
    const handler = await load({ N8N_ORGANISE_SECRET: 'the-webhook-key' });
    const res = reply();
    await handler(request({ cookie: signedIn() }), res);
    const said = JSON.stringify(res.sent.body) + JSON.stringify(res.sent.headers);
    expect(said).not.toContain('the-webhook-key');
    expect(said).not.toContain('n8n.example.invalid');
  });
});

describe('what it answers with', () => {
  it('unwraps the single item n8n replies with', async () => {
    organiserAnswers({ text: JSON.stringify([{ json: { ok: true, scanned: 29, offered: [] } }]) });
    const handler = await load();
    const res = reply();
    await handler(request({ cookie: signedIn() }), res);
    expect(res.sent.body).toEqual({ ok: true, scanned: 29, offered: [] });
  });

  it('takes a bare object just as happily', async () => {
    organiserAnswers({ text: JSON.stringify({ ok: true, scanned: 7 }) });
    const handler = await load();
    const res = reply();
    await handler(request({ cookie: signedIn() }), res);
    expect(res.sent.body).toEqual({ ok: true, scanned: 7 });
  });

  it('reports the organiser refusing, rather than pretending it answered', async () => {
    organiserAnswers({ ok: false, status: 403, text: 'Authorization is required' });
    const handler = await load();
    const res = reply();
    await handler(request({ cookie: signedIn() }), res);
    expect(res.sent.status).toBe(502);
    expect(res.sent.body.ok).toBe(false);
    expect(res.sent.body.error).toContain('403');
  });

  it('reports an answer that was not JSON', async () => {
    organiserAnswers({ text: '<html>504 Gateway Time-out</html>' });
    const handler = await load();
    const res = reply();
    await handler(request({ cookie: signedIn() }), res);
    expect(res.sent.status).toBe(502);
    expect(res.sent.body.error).toContain('not JSON');
  });

  it('says a run may still be going rather than that it failed', async () => {
    // The workflow does not stop because the page stopped listening, and
    // telling somebody it failed invites them to start a second one.
    vi.stubGlobal('fetch', vi.fn(async () => {
      throw new Error('The operation was aborted due to timeout');
    }));
    const handler = await load();
    const res = reply();
    await handler(request({ cookie: signedIn() }), res);
    expect(res.sent.status).toBe(504);
    expect(res.sent.body.error).toContain('may still be running');
  });

  it('keeps a long upstream answer from becoming the response', async () => {
    organiserAnswers({ ok: false, status: 500, text: 'x'.repeat(50000) });
    const handler = await load();
    const res = reply();
    await handler(request({ cookie: signedIn() }), res);
    expect(String(res.sent.body.detail).length).toBeLessThanOrEqual(500);
  });
});
