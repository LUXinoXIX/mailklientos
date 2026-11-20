// Prihlásenie cez Microsoft pre osobné kontá Outlook.com / Hotmail / Live (OAuth 2.0, Microsoft identity platform v2):
// systémový prehliadač + návrat na http://localhost:<port> + PKCE, verejný klient (bez tajného kľúča).
// Outlook.com od 16. 9. 2024 nepovoľuje IMAP/SMTP s heslom (ani heslom pre aplikácie) – len OAuth (XOAUTH2).
//
// Obnovovací token sa ukladá zašifrovaný v trezore systému (safeStorage) pod oauth:<id konta>, prístupový token
// (platný ~1 h) len v pamäti. Microsoft pri obnove vydáva nový obnovovací token – ukladá sa vždy ten najnovší.
//
// Aplikáciu si používateľ raz zaregistruje v Microsoft Entra (App registrations) a ID klienta vloží do nastavení
// (settings.microsoft). Do verejného repozitára sa nič nedáva.
import crypto from 'node:crypto';
import http from 'node:http';
import { net, shell } from 'electron';
import { oauthSecretKey, type AccountConfig } from '../core/account';
import { clearSecret, getSecret, setSecret } from './accounts';
import { log } from './diagnostics';
import { UserError } from './mail';
import { loadSettings } from './settings';

/** Len osobné kontá Microsoft (Outlook.com, Hotmail, Live, MSN). */
const TENANT = 'consumers';
const AUTH_URL = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/authorize`;
const TOKEN_URL = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`;

export const SCOPE_IMAP = 'https://outlook.office.com/IMAP.AccessAsUser.All';
export const SCOPE_SMTP = 'https://outlook.office.com/SMTP.Send';
const SCOPES = ['offline_access', 'openid', 'email', 'profile', SCOPE_IMAP, SCOPE_SMTP];

interface Stored {
  refresh: string;
  scope: string;
  email?: string;
}

import type { MicrosoftGrant } from '../core/api';
export type { MicrosoftGrant };

const b64url = (b: Buffer) => b.toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

function clientId() {
  const id = loadSettings().microsoft?.clientId?.trim();
  if (!id)
    throw new UserError(
      'Chýba ID aplikácie Microsoft.',
      'V nastaveniach konta Outlook.com → „Nastavenie aplikácie Microsoft“ vlož ID aplikácie (klienta) z Microsoft Entra.',
    );
  return id;
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

/** Microsoft niekedy vráti rozsahy bez predpony https://outlook.office.com/ – porovnávať podľa konca. */
const hasScope = (scope: string, s: string) => {
  const tail = s.split('/').pop()!.toLowerCase();
  return scope.toLowerCase().split(/\s+/).some((x) => x.endsWith(tail));
};

const grantOf = (s: Stored, name?: string): MicrosoftGrant => ({
  email: s.email ?? '',
  name,
  imap: hasScope(s.scope, SCOPE_IMAP),
  smtp: hasScope(s.scope, SCOPE_SMTP),
});

export function grant(accountId: string): MicrosoftGrant | null {
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

function idTokenClaims(idToken: unknown): { email?: string; preferred_username?: string; name?: string } {
  if (typeof idToken !== 'string') return {};
  try {
    const part = idToken.split('.')[1] ?? '';
    return JSON.parse(Buffer.from(part.replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf8'));
  } catch {
    return {};
  }
}

/** Zrozumiteľné vysvetlenie chýb AADSTS (Microsoft Entra). */
export function explainAad(text: string): { message: string; hint?: string } {
  const code = /AADSTS(\d+)/.exec(text)?.[1];
  switch (code) {
    case '50011':
    case '500113':
      return {
        message: 'Adresa návratu nie je v registrácii aplikácie povolená.',
        hint: 'V Microsoft Entra → tvoja aplikácia → Overovanie (Authentication) pridaj platformu „Mobilné a počítačové aplikácie“ s adresou http://localhost.',
      };
    case '700016':
    case '9002331':
    case '50194':
      return {
        message: 'Aplikácia nie je povolená pre osobné kontá Microsoft.',
        hint: 'Pri registrácii treba zvoliť typ kont „Iba osobné kontá Microsoft“ alebo „Kontá v ľubovoľnom adresári organizácie a osobné kontá Microsoft“. Skontroluj aj ID aplikácie.',
      };
    case '7000218':
      return {
        message: 'Microsoft požaduje tajný kľúč – aplikácia nie je nastavená ako verejný klient.',
        hint: 'V Microsoft Entra → Overovanie (Authentication) zapni „Povoliť toky verejného klienta“ (Allow public client flows) a adresu http://localhost zadaj pod „Mobilné a počítačové aplikácie“, nie pod „Web“.',
      };
    case '65001':
    case '65004':
      return { message: 'Prístup k pošte nebol povolený.', hint: 'Prihlás sa znova a na stránke Microsoftu potvrď „Prijať“.' };
    case '70000':
    case '70008':
    case '700082':
    case '700084':
      return { message: 'Prihlásenie Microsoft vypršalo alebo bolo zrušené.', hint: 'Otvor nastavenia konta a daj „Prihlásiť sa cez Microsoft“ znova.' };
    default:
      return { message: code ? `Microsoft odmietol prihlásenie (AADSTS${code}).` : 'Microsoft odmietol prihlásenie.' };
  }
}

const page = (title: string, text: string) =>
  `<!doctype html><meta charset="utf-8"><title>${title}</title><body style="font-family:Segoe UI,system-ui,sans-serif;background:#f5f6f7;display:grid;place-items:center;height:100vh;margin:0"><div style="background:#fff;padding:32px 40px;border-radius:10px;box-shadow:0 2px 12px #0002;max-width:460px"><h2 style="margin-top:0;color:#2F6F5E">${title}</h2><p>${text}</p></div></body>`;

let pending: { cancel: () => void } | null = null;

/** Lokálny server na návrat z prihlásenia: localhost môže v prehliadači znamenať 127.0.0.1 aj ::1 – počúva na oboch. */
async function loopback(handler: http.RequestListener) {
  const v4 = http.createServer(handler);
  await new Promise<void>((res, rej) => {
    v4.once('error', rej);
    v4.listen(0, '127.0.0.1', () => res());
  });
  const port = (v4.address() as { port: number }).port;
  const v6 = http.createServer(handler);
  await new Promise<void>((res) => {
    v6.once('error', () => res()); // bez IPv6 nevadí
    v6.listen(port, '::1', () => res());
  });
  return {
    port,
    close: () => {
      v4.close();
      if (v6.listening) v6.close();
    },
  };
}

/**
 * Prihlásenie: otvorí stránku Microsoftu v predvolenom prehliadači a počká na návrat na http://localhost:<port>.
 * Uloží obnovovací token pre konto. Najviac 5 minút.
 */
export async function signIn(accountId: string, loginHint?: string): Promise<MicrosoftGrant> {
  const id = clientId();
  pending?.cancel();
  const verifier = b64url(crypto.randomBytes(48));
  const challenge = b64url(crypto.createHash('sha256').update(verifier).digest());
  const state = b64url(crypto.randomBytes(16));

  let redirect = '';
  let srv: { port: number; close: () => void } | undefined;
  const code = await new Promise<string>((resolve, reject) => {
    let timer: NodeJS.Timeout | undefined;
    const finish = (err: Error | null, value?: string) => {
      if (timer) clearTimeout(timer);
      pending = null;
      setTimeout(() => srv?.close(), 500);
      if (err) reject(err);
      else resolve(value!);
    };
    pending = { cancel: () => finish(new UserError('Prihlásenie bolo zrušené.')) };
    const handler: http.RequestListener = (req, res) => {
      const u = new URL(req.url ?? '/', 'http://localhost');
      if (u.pathname !== '/') {
        res.writeHead(404).end();
        return;
      }
      const err = u.searchParams.get('error');
      const desc = u.searchParams.get('error_description') ?? '';
      const got = u.searchParams.get('code');
      const okState = u.searchParams.get('state') === state;
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      if (got && okState) {
        res.end(page('Prihlásenie dokončené', 'Môžeš zavrieť túto kartu a vrátiť sa do programu Universum eMail.'));
        finish(null, got);
      } else {
        const x = err === 'access_denied' ? { message: 'Prístup ku kontu Microsoft nebol povolený.' } : explainAad(`${err ?? ''} ${desc}`);
        res.end(page('Prihlásenie sa nepodarilo', x.message));
        log('microsoft: prihlásenie zlyhalo', err ?? '', desc.slice(0, 300));
        finish(new UserError(x.message, (x as { hint?: string }).hint));
      }
    };
    loopback(handler)
      .then((s) => {
        srv = s;
        redirect = `http://localhost:${s.port}`;
        timer = setTimeout(() => finish(new UserError('Prihlásenie cez Microsoft nebolo dokončené do 5 minút.')), 5 * 60_000);
        const url = new URL(AUTH_URL);
        url.search = new URLSearchParams({
          client_id: id,
          redirect_uri: redirect,
          response_type: 'code',
          response_mode: 'query',
          scope: SCOPES.join(' '),
          code_challenge: challenge,
          code_challenge_method: 'S256',
          state,
          prompt: 'select_account',
          ...(loginHint?.includes('@') ? { login_hint: loginHint } : {}),
        }).toString();
        return shell.openExternal(url.toString());
      })
      .catch((e) => finish(e as Error));
  });

  const r = await postForm(TOKEN_URL, {
    client_id: id,
    code,
    redirect_uri: redirect,
    grant_type: 'authorization_code',
    code_verifier: verifier,
    scope: SCOPES.join(' '),
  });
  if (!r.ok || typeof r.json.refresh_token !== 'string') {
    log('microsoft: výmena kódu zlyhala', r.status, r.text.slice(0, 400));
    const x = explainAad(String(r.json.error_description ?? r.json.error ?? r.status));
    throw new UserError(x.message, x.hint);
  }
  const claims = idTokenClaims(r.json.id_token);
  const email = claims.email ?? (claims.preferred_username?.includes('@') ? claims.preferred_username : undefined);
  const s: Stored = { refresh: r.json.refresh_token, scope: String(r.json.scope ?? ''), email };
  const g = grantOf(s, claims.name);
  if (!g.imap)
    throw new UserError('Pri prihlásení nebol povolený prístup k pošte (IMAP).', 'Prihlás sa znova a na stránke Microsoftu potvrď všetky oprávnenia.');
  setSecret(oauthSecretKey(accountId), JSON.stringify(s));
  tokens.set(accountId, { token: String(r.json.access_token), exp: Date.now() + Number(r.json.expires_in ?? 3600) * 1000 });
  log(`microsoft: prihlásené ${s.email ?? ''} (${s.scope})`);
  return g;
}

export function cancelSignIn() {
  pending?.cancel();
}

const tokens = new Map<string, { token: string; exp: number }>();
const refreshing = new Map<string, Promise<string>>();

/** Platný prístupový token pre IMAP a SMTP (obnoví sa automaticky). */
export async function accessToken(acc: AccountConfig): Promise<string> {
  const t = tokens.get(acc.id);
  if (t && t.exp - Date.now() > 90_000) return t.token;
  const running = refreshing.get(acc.id);
  if (running) return running;
  const p = (async () => {
    const s = stored(acc.id);
    if (!s) throw new UserError('Konto nie je prihlásené cez Microsoft.', 'Otvor nastavenia konta a daj „Prihlásiť sa cez Microsoft“.');
    let r;
    try {
      r = await postForm(TOKEN_URL, {
        client_id: clientId(),
        refresh_token: s.refresh,
        grant_type: 'refresh_token',
        scope: SCOPES.join(' '),
      });
    } catch (e) {
      throw new UserError(`Nedá sa spojiť so serverom Microsoft (${(e as Error).message}).`);
    }
    if (!r.ok || typeof r.json.access_token !== 'string') {
      log('microsoft: obnova tokenu zlyhala', r.status, r.text.slice(0, 400));
      if (r.json.error === 'invalid_grant' || r.json.error === 'interaction_required')
        throw new UserError('Prihlásenie Microsoft vypršalo alebo bolo zrušené.', 'Otvor nastavenia konta a daj „Prihlásiť sa cez Microsoft“ znova.');
      const x = explainAad(String(r.json.error_description ?? r.json.error ?? r.status));
      throw new UserError(x.message, x.hint);
    }
    // Microsoft obnovovací token pri obnove mení – uložiť nový
    if (typeof r.json.refresh_token === 'string' && r.json.refresh_token !== s.refresh)
      setSecret(oauthSecretKey(acc.id), JSON.stringify({ ...s, refresh: r.json.refresh_token, scope: String(r.json.scope ?? s.scope) }));
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

export function invalidate(accountId: string) {
  tokens.delete(accountId);
}

/** Odhlásiť: zmazať token z trezoru. (Microsoft nemá verejné zrušenie tokenu pre osobné kontá; prístup sa dá
    odobrať na account.microsoft.com → Ochrana osobných údajov → Aplikácie a služby.) */
export async function signOut(accountId: string) {
  tokens.delete(accountId);
  clearSecret(oauthSecretKey(accountId));
}
