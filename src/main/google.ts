// Prihlásenie cez Google (OAuth 2.0 pre počítačové aplikácie: systémový prehliadač + spätná adresa 127.0.0.1 + PKCE).
// Program nikdy nevidí heslo Google. Obnovovací token sa ukladá zašifrovaný v trezore systému (safeStorage),
// prístupový token (platný ~1 h) len v pamäti. Ten istý token slúži pre poštu (IMAP/SMTP XOAUTH2),
// kalendár (CalDAV) a kontakty (CardDAV).
//
// Klienta OAuth si používateľ zaregistruje v Google Cloud (typ „Počítačová aplikácia“) a ID vloží do nastavení
// (settings.google). Do verejného repozitára sa nič nedáva.
import crypto from 'node:crypto';
import http from 'node:http';
import { net, shell } from 'electron';
import { oauthSecretKey, type AccountConfig } from '../core/account';
import { clearSecret, getSecret, setSecret } from './accounts';
import { log } from './diagnostics';
import { UserError } from './mail';
import { loadSettings } from './settings';

const AUTH_URL = 'https://accounts.google.com/o/oauth2/v2/auth';
const TOKEN_URL = 'https://oauth2.googleapis.com/token';
const REVOKE_URL = 'https://oauth2.googleapis.com/revoke';

export const SCOPE_MAIL = 'https://mail.google.com/';
export const SCOPE_CALENDAR = 'https://www.googleapis.com/auth/calendar';
export const SCOPE_CONTACTS = 'https://www.googleapis.com/auth/carddav';
const SCOPES = ['openid', 'email', 'profile', SCOPE_MAIL, SCOPE_CALENDAR, SCOPE_CONTACTS];

interface Stored {
  refresh: string;
  scope: string;
  email?: string;
}

export interface GoogleGrant {
  email: string;
  name?: string;
  mail: boolean;
  calendar: boolean;
  contacts: boolean;
}

const b64url = (b: Buffer) => b.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

function client() {
  const g = loadSettings().google;
  if (!g?.clientId?.trim())
    throw new UserError(
      'Chýba ID klienta Google.',
      'V nastaveniach konta Gmail → „Prihlásenie cez Google – nastavenie klienta“ vlož ID klienta (a tajný kľúč) z Google Cloud.',
    );
  return { id: g.clientId.trim(), secret: g.clientSecret?.trim() || undefined };
}

function stored(accountId: string): Stored | null {
  const raw = getSecret(oauthSecretKey(accountId));
  if (!raw) return null;
  try {
    return JSON.parse(raw) as Stored;
  } catch {
    return null;
  }
}

const grantOf = (s: Stored, name?: string): GoogleGrant => ({
  email: s.email ?? '',
  name,
  mail: s.scope.includes(SCOPE_MAIL),
  calendar: s.scope.includes(SCOPE_CALENDAR),
  contacts: s.scope.includes(SCOPE_CONTACTS),
});

/** Čo konto povolilo (Google pri prihlásení dovolí jednotlivé oprávnenia odškrtnúť). */
export function grant(accountId: string): GoogleGrant | null {
  const s = stored(accountId);
  return s ? grantOf(s) : null;
}

async function postForm(url: string, form: Record<string, string | undefined>) {
  const body = new URLSearchParams(Object.entries(form).filter(([, v]) => v !== undefined) as [string, string][]).toString();
  const r = await net.fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body });
  const text = await r.text();
  let json: Record<string, unknown> = {};
  try {
    json = JSON.parse(text) as Record<string, unknown>;
  } catch {
    /* nie JSON */
  }
  return { ok: r.ok, status: r.status, json, text };
}

function idTokenClaims(idToken: unknown): { email?: string; name?: string } {
  if (typeof idToken !== 'string') return {};
  try {
    const part = idToken.split('.')[1] ?? '';
    return JSON.parse(Buffer.from(part.replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf8')) as { email?: string; name?: string };
  } catch {
    return {};
  }
}

const page = (title: string, text: string) =>
  `<!doctype html><meta charset="utf-8"><title>${title}</title><body style="font-family:Segoe UI,system-ui,sans-serif;background:#f5f6f7;display:grid;place-items:center;height:100vh;margin:0"><div style="background:#fff;padding:32px 40px;border-radius:10px;box-shadow:0 2px 12px #0002;max-width:460px"><h2 style="margin-top:0;color:#2F6F5E">${title}</h2><p>${text}</p></div></body>`;

let pending: { cancel: () => void } | null = null;

/**
 * Prihlásenie: otvorí stránku Google v predvolenom prehliadači a počká na návrat na http://127.0.0.1:<port>.
 * Uloží obnovovací token pre konto. Najviac 5 minút.
 */
export async function signIn(accountId: string, loginHint?: string): Promise<GoogleGrant> {
  const c = client();
  pending?.cancel();
  const verifier = b64url(crypto.randomBytes(48));
  const challenge = b64url(crypto.createHash('sha256').update(verifier).digest());
  const state = b64url(crypto.randomBytes(16));

  const server = http.createServer();
  await new Promise<void>((res, rej) => {
    server.once('error', rej);
    server.listen(0, '127.0.0.1', () => res());
  });
  const port = (server.address() as { port: number }).port;
  const redirect = `http://127.0.0.1:${port}`;

  const code = await new Promise<string>((resolve, reject) => {
    const timer = setTimeout(() => finish(new UserError('Prihlásenie cez Google nebolo dokončené do 5 minút.')), 5 * 60_000);
    const finish = (err: Error | null, value?: string) => {
      clearTimeout(timer);
      pending = null;
      setTimeout(() => server.close(), 500);
      if (err) reject(err);
      else resolve(value!);
    };
    pending = { cancel: () => finish(new UserError('Prihlásenie bolo zrušené.')) };
    server.on('request', (req, res) => {
      const u = new URL(req.url ?? '/', redirect);
      if (u.pathname !== '/') {
        res.writeHead(404).end();
        return;
      }
      const err = u.searchParams.get('error');
      const got = u.searchParams.get('code');
      const okState = u.searchParams.get('state') === state;
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      if (got && okState) {
        res.end(page('Prihlásenie dokončené', 'Môžeš zavrieť túto kartu a vrátiť sa do programu Universum eMail.'));
        finish(null, got);
      } else {
        res.end(page('Prihlásenie sa nepodarilo', err === 'access_denied' ? 'Prístup nebol povolený.' : `Chyba: ${err ?? 'neplatná odpoveď'}`));
        finish(new UserError(err === 'access_denied' ? 'Prístup ku kontu Google nebol povolený.' : `Google vrátil chybu: ${err ?? 'neplatná odpoveď'}.`));
      }
    });
    const url = new URL(AUTH_URL);
    url.search = new URLSearchParams({
      client_id: c.id,
      redirect_uri: redirect,
      response_type: 'code',
      scope: SCOPES.join(' '),
      code_challenge: challenge,
      code_challenge_method: 'S256',
      state,
      access_type: 'offline',
      prompt: 'consent',
      include_granted_scopes: 'true',
      ...(loginHint?.includes('@') ? { login_hint: loginHint } : {}),
    }).toString();
    void shell.openExternal(url.toString()).catch((e) => finish(e as Error));
  });

  const r = await postForm(TOKEN_URL, {
    code,
    client_id: c.id,
    client_secret: c.secret,
    redirect_uri: redirect,
    grant_type: 'authorization_code',
    code_verifier: verifier,
  });
  if (!r.ok || typeof r.json.refresh_token !== 'string') {
    log('google: výmena kódu zlyhala', r.status, r.text.slice(0, 300));
    const desc = String(r.json.error_description ?? r.json.error ?? r.status);
    throw new UserError(
      `Google odmietol prihlásenie (${desc}).`,
      /secret/i.test(desc) ? 'Doplň tajný kľúč klienta (client secret) v nastavení klienta Google.' : undefined,
    );
  }
  const claims = idTokenClaims(r.json.id_token);
  const s: Stored = { refresh: r.json.refresh_token, scope: String(r.json.scope ?? ''), email: claims.email };
  if (!s.scope.includes(SCOPE_MAIL))
    throw new UserError('Pri prihlásení nebol povolený prístup k pošte Gmail.', 'Prihlás sa znova a nechaj zaškrtnuté „Čítanie, písanie, odosielanie a odstraňovanie e-mailov“.');
  setSecret(oauthSecretKey(accountId), JSON.stringify(s));
  tokens.set(accountId, { token: String(r.json.access_token), exp: Date.now() + Number(r.json.expires_in ?? 3600) * 1000 });
  log(`google: prihlásené ${s.email ?? ''} (${s.scope})`);
  return grantOf(s, claims.name);
}

export function cancelSignIn() {
  pending?.cancel();
}

const tokens = new Map<string, { token: string; exp: number }>();
const refreshing = new Map<string, Promise<string>>();

/** Platný prístupový token (obnoví sa automaticky). */
export async function accessToken(acc: AccountConfig): Promise<string> {
  const t = tokens.get(acc.id);
  if (t && t.exp - Date.now() > 90_000) return t.token;
  const running = refreshing.get(acc.id);
  if (running) return running;
  const p = (async () => {
    const s = stored(acc.id);
    if (!s)
      throw new UserError('Konto nie je prihlásené cez Google.', 'Otvor nastavenia konta a daj „Prihlásiť sa cez Google“.');
    const c = client();
    let r;
    try {
      r = await postForm(TOKEN_URL, { client_id: c.id, client_secret: c.secret, refresh_token: s.refresh, grant_type: 'refresh_token' });
    } catch (e) {
      throw new UserError(`Nedá sa spojiť so serverom Google (${(e as Error).message}).`);
    }
    if (!r.ok || typeof r.json.access_token !== 'string') {
      log('google: obnova tokenu zlyhala', r.status, r.text.slice(0, 300));
      if (r.json.error === 'invalid_grant')
        throw new UserError(
          'Prihlásenie Google vypršalo alebo bolo zrušené.',
          'Otvor nastavenia konta a daj „Prihlásiť sa cez Google“ znova. Ak sa to opakuje každý týždeň, klient Google je v režime Testovanie – prepni ho v Google Cloud na „V produkcii“.',
        );
      throw new UserError(`Google nevydal prístupový token (${String(r.json.error_description ?? r.json.error ?? r.status)}).`);
    }
    const token = r.json.access_token;
    tokens.set(acc.id, { token, exp: Date.now() + Number(r.json.expires_in ?? 3600) * 1000 });
    return token;
  })();
  refreshing.set(acc.id, p);
  try {
    return await p;
  } finally {
    refreshing.delete(acc.id);
  }
}

/** Pri odmietnutom tokene (401) zabudnúť ho, aby sa nabudúce obnovil. */
export function invalidate(accountId: string) {
  tokens.delete(accountId);
}

/** Odhlásiť: zrušiť token u Google a zmazať ho z trezoru. */
export async function signOut(accountId: string) {
  const s = stored(accountId);
  tokens.delete(accountId);
  clearSecret(oauthSecretKey(accountId));
  if (s) await net.fetch(`${REVOKE_URL}?token=${encodeURIComponent(s.refresh)}`, { method: 'POST' }).catch((): void => undefined);
}
