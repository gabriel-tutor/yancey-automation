/**
 * Signing in, signing out, and answering "am I signed in?".
 *
 *   GET     is this browser signed in
 *   POST    { passcode }  → a session cookie, or 401
 *   DELETE  end the session
 *
 * The passcode never leaves the server and is never written to a log, a
 * response or an error. What comes back on success is a cookie the browser
 * cannot read.
 */
import { ATTEMPT_DELAY_MS, clearCookie, issue, sameSecret, setCookie, signedIn } from './_auth.js';

export const config = { maxDuration: 10 };

/** @param {number} ms */
const pause = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/** @param {any} req @param {any} res */
export default async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');

  if (req.method === 'GET') {
    return res.status(200).json({ signedIn: signedIn(req) });
  }

  if (req.method === 'DELETE') {
    res.setHeader('Set-Cookie', clearCookie());
    return res.status(200).json({ signedIn: false });
  }

  if (req.method !== 'POST') {
    res.setHeader('Allow', 'GET, POST, DELETE');
    return res.status(405).json({ ok: false, error: 'Use GET, POST or DELETE.' });
  }

  const expected = process.env.ORGANISE_PASSCODE;
  const secret = process.env.SESSION_SECRET || expected;

  if (!expected) {
    // Fails closed, and says which of a dozen things is wrong to the person who
    // can fix it. There is no passcode here to leak.
    return res.status(503).json({
      ok: false,
      error: 'This deployment has no ORGANISE_PASSCODE set, so nobody can sign in. See docs/WEB.md.',
    });
  }

  // Every attempt costs the same, whether or not the passcode was close.
  await pause(ATTEMPT_DELAY_MS);

  const given = req.body && typeof req.body.passcode === 'string' ? req.body.passcode : '';
  if (!given || !sameSecret(given, expected)) {
    return res.status(401).json({ ok: false, error: 'That passcode was not right.' });
  }

  res.setHeader('Set-Cookie', setCookie(issue(secret)));
  return res.status(200).json({ ok: true, signedIn: true });
}
