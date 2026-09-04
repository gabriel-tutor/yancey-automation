import { describe, expect, it } from 'vitest';
import { clearCookie, cookieFrom, freshSecret, issue, sameSecret, setCookie, valid } from '../web/api/_auth.js';

/**
 * The gate on the organiser page.
 *
 * It lives in web/api because that is where Vercel runs it, and it is tested
 * from here because a signature check that is subtly wrong looks exactly like
 * one that is right until somebody forges a cookie.
 */
const SECRET = 'a-secret-that-only-the-server-knows';

describe('comparing secrets', () => {
  it('accepts the same string', () => {
    expect(sameSecret('hunter2-but-longer', 'hunter2-but-longer')).toBe(true);
  });

  it('rejects a different one, including one that is merely a prefix', () => {
    expect(sameSecret('correct-horse', 'correct-horse-battery')).toBe(false);
    expect(sameSecret('correct-horse-battery', 'correct-horse')).toBe(false);
    expect(sameSecret('a', 'b')).toBe(false);
  });

  it('rejects anything that is not a string', () => {
    expect(sameSecret(/** @type {any} */ (undefined), 'x')).toBe(false);
    expect(sameSecret('x', /** @type {any} */ (null))).toBe(false);
    expect(sameSecret(/** @type {any} */ ({}), /** @type {any} */ ({}))).toBe(false);
  });
});

describe('the session token', () => {
  it('accepts one it just issued', () => {
    expect(valid(issue(SECRET), SECRET)).toBe(true);
  });

  it('refuses one signed with a different secret', () => {
    expect(valid(issue('some-other-secret'), SECRET)).toBe(false);
  });

  it('refuses one whose payload was edited after signing', () => {
    // The whole point: a browser can read its own cookie in devtools, so the
    // expiry has to be unforgeable rather than merely inconvenient to change.
    const token = issue(SECRET);
    const forged = Buffer.from(JSON.stringify({ exp: 9999999999 })).toString('base64url') +
      '.' + token.slice(token.indexOf('.') + 1);
    expect(valid(forged, SECRET)).toBe(false);
  });

  it('refuses one that has expired', () => {
    const issuedAt = Date.UTC(2026, 0, 1, 9, 0, 0);
    const token = issue(SECRET, issuedAt);
    expect(valid(token, SECRET, issuedAt + 60 * 1000)).toBe(true);
    expect(valid(token, SECRET, issuedAt + 13 * 60 * 60 * 1000)).toBe(false);
  });

  it('refuses rubbish rather than throwing on it', () => {
    for (const rubbish of ['', '.', 'nodot', 'a.b', '...', 'eyJ4Ijox.', /** @type {any} */ (null)]) {
      expect(valid(rubbish, SECRET)).toBe(false);
    }
  });

  it('refuses everything when there is no secret, rather than letting anyone in', () => {
    expect(valid(issue(SECRET), '')).toBe(false);
  });

  it('issues a different token each time it is asked', () => {
    expect(freshSecret()).not.toBe(freshSecret());
    expect(freshSecret().length).toBeGreaterThan(32);
  });
});

describe('reading the cookie back', () => {
  it('finds the one it is asked for among several', () => {
    expect(cookieFrom('a=1; organiser_session=abc.def; z=9', 'organiser_session')).toBe('abc.def');
  });

  it('is not fooled by a name that merely ends the same way', () => {
    expect(cookieFrom('not_organiser_session=wrong', 'organiser_session')).toBe('');
  });

  it('survives a header that is missing or malformed', () => {
    expect(cookieFrom(undefined, 'organiser_session')).toBe('');
    expect(cookieFrom('', 'organiser_session')).toBe('');
    expect(cookieFrom('novalue', 'organiser_session')).toBe('');
  });

  it('round-trips a token through the header it sets', () => {
    const token = issue(SECRET);
    const header = setCookie(token).split(';')[0];
    expect(valid(cookieFrom(header, 'organiser_session'), SECRET)).toBe(true);
  });
});

describe('the cookie itself', () => {
  const flags = setCookie(issue(SECRET));

  it('cannot be read by page scripts', () => {
    expect(flags).toContain('HttpOnly');
  });

  it('never crosses plain HTTP', () => {
    expect(flags).toContain('Secure');
  });

  it('cannot be spent by another site', () => {
    expect(flags).toContain('SameSite=Lax');
  });

  it('expires with the token rather than outliving it', () => {
    expect(flags).toContain('Max-Age=43200');
  });

  it('is emptied and expired when the session ends', () => {
    expect(clearCookie()).toContain('Max-Age=0');
    expect(clearCookie()).toContain('HttpOnly');
  });
});
