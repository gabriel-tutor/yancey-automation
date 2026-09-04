/**
 * The page talks to this; this talks to n8n.
 *
 * A browser calling the n8n webhook directly would need CORS on the n8n side
 * and would put the webhook URL — the only thing guarding a workflow that
 * renames a client's documents — into view-source on every laptop that opens
 * the page. Proxying from the same origin removes both problems at once: no
 * CORS preflight, and the URL stays on the server.
 *
 * It is deliberately thin. Everything that decides anything is in the workflow,
 * where it is tested; this forwards a request and forwards an answer.
 */

import { signedIn } from './_auth.js';

/**
 * A scan reads the whole library, opens what it can and asks a model to name
 * it, which takes about a minute. Vercel's default is far shorter, and a
 * function that gives up at ten seconds turns a working run into a page that
 * says it failed.
 */
export const config = { maxDuration: 300 };

const ORGANISE_URL = process.env.N8N_ORGANISE_URL;
const ORGANISE_SECRET = process.env.N8N_ORGANISE_SECRET;

/** @param {any} req @param {any} res */
export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');

  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ ok: false, error: 'Use POST.' });
  }

  // The gate is here and not only on the page. A page that hides a button is a
  // page, not a guard: the endpoint is what reaches SharePoint, so the endpoint
  // is what has to ask.
  if (!signedIn(req)) {
    return res.status(401).json({ ok: false, signedIn: false, error: 'Sign in first.' });
  }

  if (!ORGANISE_URL) {
    // Said plainly, because the person who sees this is the person who can fix
    // it, and "500" would not tell them which of a dozen things went wrong.
    return res.status(500).json({
      ok: false,
      error: 'This deployment has no N8N_ORGANISE_URL set, so it does not know where the organiser lives.',
    });
  }

  const chosen = Array.isArray(req.body && req.body.chosen) ? req.body.chosen : [];

  // Forwarded whole rather than picked apart. What a setting is allowed to do
  // is decided by mergeOverrides in the workflow, where it is tested; a proxy
  // that also had an opinion would be a second place to keep in step, and the
  // first version of this file forgot to forward settings at all.
  const asked = req.body && req.body.settings;
  const settings = asked && typeof asked === 'object' && !Array.isArray(asked) ? asked : {};

  /** @type {Record<string, string>} */
  const headers = { 'Content-Type': 'application/json' };
  // Optional today, and the one thing worth adding tomorrow. See docs/WEB.md.
  if (ORGANISE_SECRET) headers['X-Organise-Key'] = ORGANISE_SECRET;

  try {
    const answer = await fetch(ORGANISE_URL, {
      method: 'POST',
      headers,
      body: JSON.stringify({ chosen, settings }),
    });

    const text = await answer.text();

    if (!answer.ok) {
      return res.status(502).json({
        ok: false,
        error: 'The organiser answered ' + answer.status + '.',
        detail: text.slice(0, 500),
      });
    }

    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch {
      return res.status(502).json({
        ok: false,
        error: 'The organiser answered with something that was not JSON.',
        detail: text.slice(0, 500),
      });
    }

    // n8n answers with the items the last node emitted, which is an array of
    // one. Either shape is unwrapped so the page only ever sees the object.
    const body = Array.isArray(parsed) ? parsed[0] : parsed;
    return res.status(200).json(body && body.json ? body.json : body);
  } catch (err) {
    // A scan that runs past this function's own ceiling lands here, and saying
    // "it may still be running" is the truth: the workflow does not stop
    // because the page stopped listening.
    const message = err instanceof Error ? err.message : String(err);
    return res.status(504).json({
      ok: false,
      error: 'The organiser did not answer in time. It may still be running — check n8n before running it again.',
      detail: message.slice(0, 300),
    });
  }
}
