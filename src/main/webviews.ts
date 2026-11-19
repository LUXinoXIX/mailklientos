// Vložené weby (Outlook na webe pre kontá M365, weby v pravom paneli) ako značka <webview> v stránke.
// Vykresľujú sa vnútri stránky (nie ako samostatná vrstva nad oknom) – na Windows tak nedochádza
// k blikaniu okna pri ich zobrazení/skrytí a dialógy aplikácie ich môžu prekryť.
// Každý web má vlastnú trvalú reláciu (partition): oddelené prihlásenie a cookies.
import { attachContextMenu } from './contextmenu';
import path from 'node:path';
import { app, BrowserWindow, session, shell, type Session, type WebContents } from 'electron';
import { getPassword, loadAccounts } from './accounts';
import type { WebEvent } from '../core/api';
import { log } from './diagnostics';

const prepared = new WeakSet<Session>();
let win: BrowserWindow | null = null;
let tabSeq = 0;

function emit(ev: WebEvent) {
  if (win && !win.isDestroyed()) win.webContents.send('web:event', ev);
}

/** Prihlasovacie okná (MFA, OAuth) musia zostať ako vyskakovacie okno s väzbou na otvárajúcu stránku. */
const AUTH_HOSTS = [
  /^login\.microsoftonline\.com$/i, /^login\.live\.com$/i, /^login\.microsoft\.com$/i, /^account\.live\.com$/i,
  /^accounts\.google\.com$/i, /(^|\.)msauth\.net$/i, /(^|\.)msftauth\.net$/i, /^device\.login\.microsoftonline\.com$/i,
];
const isAuth = (url: string) => {
  try {
    const u = new URL(url);
    return u.protocol === 'about:' || AUTH_HOSTS.some((r) => r.test(u.hostname));
  } catch {
    return false;
  }
};

/** Microsoft služby (Office na webe, Teams, OneDrive…), ktoré sa otvárajú ako karta v aplikácii. */
const MS_HOSTS = [
  /(^|\.)microsoftonline\.com$/i, /(^|\.)microsoft\.com$/i, /(^|\.)live\.com$/i, /(^|\.)office\.com$/i,
  /(^|\.)office365\.com$/i, /(^|\.)outlook\.com$/i, /(^|\.)sharepoint\.com$/i, /(^|\.)cloud\.microsoft$/i,
  /(^|\.)microsoft365\.com$/i, /(^|\.)skype\.com$/i,
];
const isMs = (url: string) => {
  try {
    return MS_HOSTS.some((r) => r.test(new URL(url).hostname));
  } catch {
    return false;
  }
};

/** User-agent bez „Electron/…“ a názvu aplikácie, aby weby (Outlook, WhatsApp) videli bežný Chrome. */
const chromeUserAgent = (ua: string) => ua.replace(/\s(MailKlient|mailklient)\/\S+/g, '').replace(/\sElectron\/\S+/g, '');

export function prepareSession(partition: string) {
  const ses = session.fromPartition(partition);
  if (prepared.has(ses)) return ses;
  prepared.add(ses);
  ses.setUserAgent(chromeUserAgent(ses.getUserAgent()));
  // Oznámenia (nová pošta), mikrofón a kamera (Teams), schránka, celá obrazovka.
  const allowed = new Set(['notifications', 'media', 'clipboard-sanitized-write', 'clipboard-read', 'fullscreen', 'display-capture']);
  ses.setPermissionRequestHandler((_wc, perm, cb) => cb(allowed.has(perm)));
  // Outlook na webe a Teams sa pred zobrazením upozornenia pýtajú na stav povolenia – odpovedať „povolené“,
  // upozornenia z webu potom Electron zobrazí ako upozornenia Windows
  ses.setPermissionCheckHandler((_wc, perm) => allowed.has(perm));
  ses.setPermissionCheckHandler((_wc, perm) => allowed.has(perm));
  return ses;
}

/** Rovnaký server/doména (mail.firma.sk a webmail.firma.sk sú tá istá stránka). */
const siteOf = (u: string) => {
  try {
    const h = new URL(u).hostname.toLowerCase();
    if (/^\d+\.\d+\.\d+\.\d+$/.test(h)) return h;
    return h.split('.').slice(-2).join('.');
  } catch {
    return '';
  }
};
const sameSite = (a: string, b: string) => !!siteOf(a) && siteOf(a) === siteOf(b);

/** Vložený webmail s neplatným certifikátom (interný server): dôverovať, len ak to konto povolí, len pre jeho server. */
let certHandler = false;
function installCertHandler() {
  if (certHandler) return;
  certHandler = true;
  app.on('certificate-error', (event, wc, url, error, _cert, callback) => {
    try {
      const host = new URL(url).hostname;
      const acc = loadAccounts().find(
        (a) => a.kind === 'm365' && 'host' in a && a.host === 'web' && a.webIgnoreCert && a.webUrl && new URL(a.webUrl).hostname === host && wc.session === session.fromPartition(`persist:acc-${a.id}`),
      );
      if (acc) {
        event.preventDefault();
        log('webmail: neplatný certifikát prijatý', { host, error });
        callback(true);
        return;
      }
    } catch {
      /* nič */
    }
    callback(false);
  });
}

/** Skupina kariet podľa relácie: „acc-<id>“ pre kontá M365, „web-<id>“ pre pravý panel. */
const groupOf = (guest: WebContents) => path.basename(guest.session.storagePath ?? '');

/** Skript, ktorý vyplní prihlasovací formulár webmailu (meno + heslo), voliteľne ho odošle.
    Podporuje aj prihlásenie v dvoch krokoch (najprv meno, potom heslo na ďalšej stránke). */
const FILL = `(function (u, p, submit) {
  const vis = (e) => !!e && e.offsetParent !== null && !e.disabled && !e.readOnly;
  const set = (el, v) => {
    const d = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
    d.set.call(el, v);
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  };
  const userLike = (x) => /^(text|email|)$/i.test(x.type || 'text') && !/search|captcha|otp|code/i.test((x.name || '') + (x.id || ''));
  const pw = [...document.querySelectorAll('input[type=password]')].find(vis);
  if (!pw) {
    const only = [...document.querySelectorAll('input')].filter(vis).find((x) => userLike(x) && /user|login|mail|meno|account|name/i.test((x.name || '') + (x.id || '') + (x.autocomplete || '')));
    if (only && u && !only.value) { set(only, u); return 'user'; }
    return 'none';
  }
  if (pw.value) return 'already';
  const scope = pw.form || document;
  const inputs = [...scope.querySelectorAll('input')].filter(vis);
  const i = inputs.indexOf(pw);
  const user = inputs.slice(0, i < 0 ? inputs.length : i).reverse().find(userLike);
  if (user && u && !user.value) set(user, u);
  set(pw, p);
  if (submit) {
    setTimeout(() => {
      const btn = scope.querySelector('button[type=submit], input[type=submit]') || [...scope.querySelectorAll('button')].find(vis);
      if (btn) btn.click(); else if (pw.form && pw.form.requestSubmit) pw.form.requestSubmit();
    }, 350);
    return 'submitted';
  }
  return 'filled';
})`;
const lastSubmit = new Map<string, number>();

/** Vložený webmail s uloženým heslom: po načítaní prihlasovacej stránky vyplniť meno a heslo. */
function autofill(guest: WebContents, group: string) {
  if (!group.startsWith('acc-')) return;
  const id = group.slice(4);
  const acc = loadAccounts().find((a) => a.id === id);
  if (!acc || acc.kind !== 'm365' || acc.host !== 'web' || !acc.webUrl) return;
  try {
    if (siteOf(guest.getURL()) !== siteOf(acc.webUrl)) return; // len na stránke webmailu, nie na cudzej
  } catch {
    return;
  }
  const pass = getPassword(id);
  if (!pass) return;
  // automatické odoslanie najviac raz za 2 minúty – pri zlom hesle sa nezacyklí
  const canSubmit = !!acc.webAutoLogin && Date.now() - (lastSubmit.get(id) ?? 0) > 120_000;
  const run = () =>
    guest
      .executeJavaScript(`${FILL}(${JSON.stringify(acc.webUser ?? '')}, ${JSON.stringify(pass)}, ${canSubmit})`, true)
      .then((r: string) => {
        if (r === 'submitted') lastSubmit.set(id, Date.now());
        return r;
      })
      .catch(() => 'error');
  void run().then((r) => {
    // prihlasovací formulár sa niekedy vykreslí až po načítaní (Zimbra, moderné webmaily)
    if (r === 'none') setTimeout(() => void run(), 1500);
  });
}

function setupGuest(guest: WebContents) {
  const group = groupOf(guest);
  guest.on('did-finish-load', () => autofill(guest, group));
  guest.on('did-navigate-in-page', () => setTimeout(() => autofill(guest, group), 600));
  guest.setWindowOpenHandler(({ url, disposition }) => {
    // 1) prihlásenie: malé okno (OAuth potrebuje window.opener)
    if (isAuth(url) || (disposition === 'new-window' && isMs(url) && /oauth|authorize|login/i.test(url)))
      return {
        action: 'allow',
        overrideBrowserWindowOptions: { autoHideMenuBar: true, width: 520, height: 720, parent: win ?? undefined },
      };
    // 2) Office na webe, Kalendár, Teams… z konta M365: nová karta vo vloženom okne (rovnaké prihlásenie)
    if (isMs(url) && group.startsWith('acc-')) {
      emit({ type: 'tab-opened', group, key: `${group}~${++tabSeq}`, url });
      return { action: 'deny' };
    }
    // 2b) vložený webmail (Zimbra, Roundcube…): odkazy na ten istý server (nové okno správy, písanie) ako karta
    if (group.startsWith('acc-') && sameSite(url, guest.getURL())) {
      emit({ type: 'tab-opened', group, key: `${group}~${++tabSeq}`, url });
      return { action: 'deny' };
    }
    // 3) ostatné odkazy: bežný prehliadač
    if (/^(https?|mailto):/i.test(url)) void shell.openExternal(url);
    return { action: 'deny' };
  });
}

/** Prepojí hlavné okno: bezpečnostné nastavenia pre <webview> a obsluha nových okien. */
export function attachWindow(w: BrowserWindow) {
  win = w;
  installCertHandler();
  w.on('closed', () => {
    win = null;
  });
  w.webContents.on('will-attach-webview', (e, prefs, params) => {
    // len webové adresy, bez Node a bez vlastného preloadu
    if (!/^https?:/i.test(params.src ?? '') || !/^persist:(acc|web)-/.test(params.partition ?? '')) {
      log('webview odmietnutý', { src: (params.src ?? '').slice(0, 80), partition: params.partition });
      e.preventDefault();
      return;
    }
    delete prefs.preload;
    prefs.nodeIntegration = false;
    prefs.contextIsolation = true;
    prefs.sandbox = true;
    prefs.backgroundThrottling = false;
    prepareSession(params.partition!);
  });
  w.webContents.on('did-attach-webview', (_e, guest) => {
    setupGuest(guest);
    attachContextMenu(guest, { web: true });
  });
}

/** Otvorí stránku v Edge (Windows) alebo v predvolenom prehliadači. */
export async function openOutside(url: string) {
  if (!/^https?:/i.test(url)) return;
  if (process.platform === 'win32') {
    try {
      await shell.openExternal(`microsoft-edge:${url}`);
      return;
    } catch {
      /* Edge nie je: predvolený prehliadač */
    }
  }
  await shell.openExternal(url);
}

/** Odhlásenie / zmazanie konta: vymaže uložené prihlásenie relácie. */
export async function clearPartition(partition: string) {
  await session.fromPartition(partition).clearStorageData();
}
