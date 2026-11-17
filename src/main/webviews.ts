// Vložené weby (Outlook na webe pre kontá M365, weby v pravom paneli) ako WebContentsView.
// Každý web má vlastnú trvalú reláciu (partition): oddelené prihlásenie a cookies.
// Renderer posiela, kde má web ležať (obdĺžnik v px okna); viditeľných môže byť viac naraz.
import { BrowserWindow, session, shell, WebContentsView, type Session } from 'electron';
import type { WebEvent } from '../core/api';

export interface Bounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface Entry {
  view: WebContentsView;
  partition: string;
  attached: boolean;
  /** Skupina kariet (napr. konto M365); nové Office aplikácie sa otvárajú ako karty v nej. */
  group: string;
}


let tabSeq = 0;
function emit(ev: WebEvent) {
  if (win && !win.isDestroyed()) win.webContents.send('web:event', ev);
}

const views = new Map<string, Entry>();
const prepared = new WeakSet<Session>();
let win: BrowserWindow | null = null;

export function attachWindow(w: BrowserWindow) {
  win = w;
  w.on('closed', () => {
    win = null;
    views.clear();
  });
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

/** Microsoft služby (Office na webe, Teams, OneDrive…) a weby, ktoré necháme v aplikácii. */
const POPUP_HOSTS = [
  /(^|\.)microsoftonline\.com$/i, /(^|\.)microsoft\.com$/i, /(^|\.)live\.com$/i, /(^|\.)office\.com$/i,
  /(^|\.)office365\.com$/i, /(^|\.)outlook\.com$/i, /(^|\.)sharepoint\.com$/i, /(^|\.)cloud\.microsoft$/i,
  /(^|\.)microsoft365\.com$/i, /(^|\.)skype\.com$/i, /(^|\.)msauth\.net$/i, /(^|\.)msftauth\.net$/i,
  /(^|\.)whatsapp\.com$/i, /(^|\.)google\.com$/i,
];
const allowPopup = (url: string) => {
  try {
    const u = new URL(url);
    if (u.protocol === 'about:') return true;
    return POPUP_HOSTS.some((r) => r.test(u.hostname));
  } catch {
    return false;
  }
};

/** User-agent bez „Electron/…“ a názvu aplikácie, aby weby (Outlook, WhatsApp) videli bežný Chrome. */
function chromeUserAgent(ua: string) {
  return ua.replace(/\s(MailKlient|mailklient)\/\S+/g, '').replace(/\sElectron\/\S+/g, '');
}

function prepareSession(partition: string) {
  const ses = session.fromPartition(partition);
  if (prepared.has(ses)) return ses;
  prepared.add(ses);
  ses.setUserAgent(chromeUserAgent(ses.getUserAgent()));
  // Oznámenia (nová pošta), mikrofón a kamera (Teams), schránka, celá obrazovka.
  const allowed = new Set(['notifications', 'media', 'clipboard-sanitized-write', 'clipboard-read', 'fullscreen', 'display-capture']);
  ses.setPermissionRequestHandler((_wc, perm, cb) => cb(allowed.has(perm)));
  ses.setPermissionCheckHandler((_wc, perm) => allowed.has(perm));
  return ses;
}

function create(key: string, url: string, partition: string, group = key): Entry {
  prepareSession(partition);
  const view = new WebContentsView({
    webPreferences: { partition, contextIsolation: true, nodeIntegration: false, sandbox: true, backgroundThrottling: false },
  });
  view.setBackgroundColor('#ffffff');
  const wc = view.webContents;
  wc.setWindowOpenHandler(({ url: target, disposition }) => {
    // 1) prihlásenie: skutočné malé okno (OAuth potrebuje window.opener)
    if (isAuth(target) || (disposition === 'new-window' && allowPopup(target) && /oauth|authorize|login/i.test(target)))
      return {
        action: 'allow',
        overrideBrowserWindowOptions: { autoHideMenuBar: true, width: 520, height: 720, parent: win ?? undefined },
      };
    // 2) Office na webe, Kalendár, Teams…: nová karta vo vloženom okne (rovnaké prihlásenie)
    if (allowPopup(target) && group.startsWith('acc-')) {
      const tabKey = `${group}~${++tabSeq}`;
      create(tabKey, target, partition, group);
      emit({ type: 'tab-opened', group, key: tabKey, url: target });
      return { action: 'deny' };
    }
    // 3) ostatné odkazy: bežný prehliadač
    if (/^(https?|mailto):/i.test(target)) void shell.openExternal(target);
    return { action: 'deny' };
  });
  wc.on('page-title-updated', (_e, title) => emit({ type: 'title', key, title }));
  wc.on('did-fail-load', (_e, code, desc, failedUrl, isMain) => {
    if (!isMain || code === -3) return; // -3 = zrušené (presmerovanie)
    const html = `<!doctype html><meta charset="utf-8"><body style="font-family:Segoe UI,system-ui,sans-serif;padding:32px;color:#242424">
      <h2 style="font-weight:600">Stránku sa nepodarilo načítať</h2><p>${desc} (${code})</p><p style="color:#616161">${failedUrl}</p>
      <p>Skontroluj pripojenie na internet a skús tlačidlo Obnoviť.</p></body>`;
    void wc.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
  });
  const entry: Entry = { view, partition, attached: false, group };
  views.set(key, entry);
  void wc.loadURL(url);
  return entry;
}

/** Kľúče: „acc-<id>“ (Outlook), „acc-<id>:<modul>“, „acc-<id>~<n>“ (karta z Office), „web-<id>“ (pravý panel). */
const groupOf = (key: string) => key.split(/[~:]/)[0]!;

function clamp(b: Bounds): Bounds {
  return {
    x: Math.max(0, Math.round(b.x)),
    y: Math.max(0, Math.round(b.y)),
    width: Math.max(0, Math.round(b.width)),
    height: Math.max(0, Math.round(b.height)),
  };
}

/** Zobrazí web na danom mieste (vytvorí ho pri prvom použití). */
export function show(key: string, url: string, partition: string, bounds: Bounds) {
  if (!win) return;
  let e = views.get(key);
  if (e && e.partition !== partition) {
    destroy(key);
    e = undefined;
  }
  if (!e) e = create(key, url, partition, groupOf(key));
  if (!e.attached) {
    win.contentView.addChildView(e.view);
    e.attached = true;
  }
  e.view.setBounds(clamp(bounds));
  e.view.setVisible(true);
}

export function setBounds(key: string, bounds: Bounds) {
  views.get(key)?.view.setBounds(clamp(bounds));
}

/** Skryje web (relácia a stav stránky zostanú, aby sa pri návrate nenačítaval znova). */
export function hide(key: string) {
  const e = views.get(key);
  if (e) e.view.setVisible(false);
}

export function hideAll() {
  for (const e of views.values()) e.view.setVisible(false);
}

export function navigate(key: string, url: string) {
  const e = views.get(key);
  if (e) void e.view.webContents.loadURL(url);
}

export function reload(key: string) {
  views.get(key)?.view.webContents.reload();
}

export function back(key: string) {
  const nav = views.get(key)?.view.webContents.navigationHistory;
  if (nav?.canGoBack()) nav.goBack();
}

export function currentUrl(key: string): string | null {
  return views.get(key)?.view.webContents.getURL() ?? null;
}

export function destroy(key: string) {
  const e = views.get(key);
  if (!e) return;
  views.delete(key);
  try {
    if (win && e.attached) win.contentView.removeChildView(e.view);
    e.view.webContents.close();
  } catch {
    /* okno už neexistuje */
  }
}

/** Odhlásenie / zmazanie konta: zatvorí jeho weby a vymaže uložené prihlásenie. */
export async function clearPartition(partition: string) {
  for (const [k, e] of views) if (e.partition === partition) destroy(k);
  await session.fromPartition(partition).clearStorageData();
}
