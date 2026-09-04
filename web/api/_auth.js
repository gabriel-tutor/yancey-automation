/**
 * Who is allowed to use the organiser.
 *
 * One shared passcode, because there is one person. Anything more — accounts,
 * an identity provider, per-user roles — would be machinery for a population
 * that does not exist, and machinery nobody exercises is machinery nobody
 * notices breaking.
 *
 * The shape:
 *
 *   the passcode           lives in a Vercel environment variable, server side,
 *                          and is compared in constant time
 *   what the browser holds is a signed token in an httpOnly cookie — not the
 *                          passcode, and not readable from JavaScript
 *   what the token says    is only when it expires; there is nothing else to say
 *
 * A file under `api/` whose name starts with an underscore is a module rather
 * than a route, so this is never reachable over HTTP.
 */
import { createHmac, randomBytes, timingSafeEqual, createHash } from 'node:crypto';

export const COOKIE = 'organiser_session';

/** Twelve hours: long enough for a working day, short enough that a forgotten
 * laptop stops being a way in overnight. */
const LIFETIME_SECONDS = 12 * 60 * 60;

/**
 * Sign-in is deliberately slow.
 *
 * A serverless function keeps no state between requests, so there is nowhere to
 * count failed attempts. What is left is to make each attempt cost something. It
 * is not a rate limiter and should not be mistaken for one — it buys time, and
 * the length of the passcode does the rest.
 */
export const ATTEMPT_DELAY_MS = 600;

/** @param {Buffer | Uint8Array | string} buffer */
function b64url(buffer) {
  return Buffer.from(buffer).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

/** @param {string} value */
function fromB64url(value) {
  const padded = String(value).replace(/-/g, '+').replace(/_/g, '/');
  return Buffer.from(padded + '==='.slice((padded.length + 3) % 4), 'base64');
}

/**
 * Compare two secrets without leaking, through timing, how much of one was
 * right. Hashing first means the comparison is over two equal-length buffers
 * whatever the inputs were, which timingSafeEqual requires.
 *
 * @param {string} a @param {string} b
 */
export function sameSecret(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string') return false;
  const left = createHash('sha256').update(a, 'utf8').digest();
  const right = createHash('sha256').update(b, 'utf8').digest();
  return timingSafeEqual(left, right);
}

/**
 * A token is its expiry and a signature over it. There is nothing else worth
 * putting in one: it says a browser signed in and when that stops counting.
 *
 * @param {string} secret
 * @param {number} [now] milliseconds, for tests
 * @returns {string}
 */
export function issue(secret, now) {
  const at = typeof now === 'number' ? now : Date.now();
  const payload = b64url(JSON.stringify({ exp: Math.floor(at / 1000) + LIFETIME_SECONDS }));
  const mac = b64url(createHmac('sha256', secret).update(payload).digest());
  return payload + '.' + mac;
}

/**
 * @param {string} token
 * @param {string} secret
 * @param {number} [now] milliseconds, for tests
 * @returns {boolean}
 */
export function valid(token, secret, now) {
  if (typeof token !== 'string' || !secret) return false;
  const cut = token.indexOf('.');
  if (cut <= 0) return false;

  const payload = token.slice(0, cut);
  const mac = token.slice(cut + 1);
  const wanted = b64url(createHmac('sha256', secret).update(payload).digest());
  // Signature first, and in constant time. Reading the payload of something
  // that has not been shown to be ours is doing work an attacker chose.
  if (!sameSecret(mac, wanted)) return false;

  let claim;
  try {
    claim = JSON.parse(fromB64url(payload).toString('utf8'));
  } catch {
    return false;
  }
  const at = typeof now === 'number' ? now : Date.now();
  return typeof claim.exp === 'number' && claim.exp * 1000 > at;
}

/**
 * Read one cookie out of a Cookie header.
 *
 * @param {string | undefined} header
 * @param {string} name
 * @returns {string}
 */
export function cookieFrom(header, name) {
  if (typeof header !== 'string') return '';
  for (const part of header.split(';')) {
    const at = part.indexOf('=');
    if (at === -1) continue;
    if (part.slice(0, at).trim() !== name) continue;
    try {
      return decodeURIComponent(part.slice(at + 1).trim());
    } catch {
      return part.slice(at + 1).trim();
    }
  }
  return '';
}

/**
 * The Set-Cookie for a signed-in browser.
 *
 * httpOnly so page scripts cannot read it, Secure so it never crosses plain
 * HTTP, and SameSite=Lax so another site cannot make the browser spend it.
 *
 * @param {string} token
 * @returns {string}
 */
export function setCookie(token) {
  return COOKIE + '=' + encodeURIComponent(token) +
    '; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=' + LIFETIME_SECONDS;
}

/** The Set-Cookie that ends a session. */
export function clearCookie() {
  return COOKIE + '=; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=0';
}

/**
 * Whether this request carries a session this deployment issued.
 *
 * Fails closed. A deployment with no passcode set is not an open one — it is a
 * broken one, and letting everybody in would be the worst possible reading of
 * a missing setting.
 *
 * @param {{ headers?: Record<string, any> }} req
 * @returns {boolean}
 */
export function signedIn(req) {
  const secret = process.env.SESSION_SECRET || process.env.ORGANISE_PASSCODE;
  if (!secret) return false;
  const header = req && req.headers ? req.headers.cookie : '';
  return valid(cookieFrom(header, COOKIE), secret);
}

/** A fresh secret, for setting up a deployment. */
export function freshSecret() {
  return b64url(randomBytes(32));
}
