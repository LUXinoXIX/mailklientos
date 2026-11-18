// Doplnky: tlač a PDF, náhľad príloh, odhlásenie z newslettera, pripomenutie bez odpovede, záloha nastavení,
// ikona v oznamovacej oblasti s počtom neprečítaných.
import { getMainWindow } from './mainwin';
import fs from 'node:fs';
import https from 'node:https';
import os from 'node:os';
import path from 'node:path';
import { app, BrowserWindow, dialog, Menu, nativeImage, Notification, shell, Tray } from 'electron';
import type { AccountConfig } from '../core/account';
import type { MessageFull } from '../core/api';
import * as store from './accounts';
import { log } from './diagnostics';
import * as mail from './mail';
import { loadSettings, saveSettings } from './settings';

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const win = () => BrowserWindow.getFocusedWindow() ?? BrowserWindow.getAllWindows()[0] ?? null;
const emit = (ev: unknown) => {
  for (const w of BrowserWindow.getAllWindows()) if (!w.isDestroyed()) w.webContents.send('app:event', ev);
};

// ---------------------------------------------------------------------------
// tlač a PDF

function printable(m: MessageFull) {
  const head = [
    ['Od', m.fromName ? `${m.fromName} <${m.fromAddress}>` : m.fromAddress],
    ['Odoslané', m.date ? new Date(m.date).toLocaleString('sk-SK') : ''],
    ['Komu', m.to],
    ...(m.cc ? [['Kópia', m.cc]] : []),
    ['Predmet', m.subject],
    ...(m.attachments.length ? [['Prílohy', m.attachments.map((a) => a.filename).join(', ')]] : []),
  ];
  const body = m.html ? m.html.replace(/^[\s\S]*<body[^>]*>/i, '').replace(/<\/body>[\s\S]*$/i, '').replace(/<script[\s\S]*?<\/script>/gi, '') : `<pre style="white-space:pre-wrap;font-family:inherit">${esc(m.text)}</pre>`;
  return `<!doctype html><html><head><meta charset="utf-8"><style>
body{font-family:'Segoe UI',Arial,sans-serif;font-size:11pt;color:#111;margin:24px}
h1{font-size:16pt;margin:0 0 8px}table.h{border-collapse:collapse;margin-bottom:14px;font-size:10pt}
table.h td{padding:2px 10px 2px 0;vertical-align:top}table.h td:first-child{font-weight:600;color:#444}
hr{border:0;border-top:1px solid #bbb;margin:10px 0 16px}img{max-width:100%}
</style></head><body><h1>${esc(m.subject)}</h1><table class="h">${head.map(([k, v]) => `<tr><td>${k}:</td><td>${esc(v ?? '')}</td></tr>`).join('')}</table><hr>${body}</body></html>`;
}

export async function printMessage(acc: AccountConfig, folder: string, uid: number, pdf: boolean) {
  const m = await mail.getMessage(acc, folder, uid);
  const w = new BrowserWindow({ show: false, webPreferences: { javascript: false, sandbox: true } });
  try {
    await w.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(printable(m))}`);
    if (pdf) {
      const parent = win();
      const opts = { title: 'Uložiť ako PDF', defaultPath: path.join(os.homedir(), 'Documents', `${m.subject.replace(/[\\/:*?"<>|]/g, '_').slice(0, 100) || 'sprava'}.pdf`), filters: [{ name: 'PDF', extensions: ['pdf'] }] };
      const r = parent ? await dialog.showSaveDialog(parent, opts) : await dialog.showSaveDialog(opts);
      if (r.canceled || !r.filePath) return '';
      fs.writeFileSync(r.filePath, await w.webContents.printToPDF({ pageSize: 'A4', printBackground: true }));
      return r.filePath;
    }
    await new Promise<void>((resolve) => w.webContents.print({ silent: false, printBackground: true }, () => resolve()));
    return 'print';
  } finally {
    if (!w.isDestroyed()) w.destroy();
  }
}

// ---------------------------------------------------------------------------
// náhľad príloh

const IMG = /\.(png|jpe?g|gif|webp|bmp|svg)$/i;

export async function previewAttachment(acc: AccountConfig, folder: string, uid: number, index: number) {
  const a = await mail.getAttachment(acc, folder, uid, index);
  if (IMG.test(a.filename)) {
    const ext = path.extname(a.filename).slice(1).toLowerCase().replace('jpg', 'jpeg').replace('svg', 'svg+xml');
    return { kind: 'image' as const, src: `data:image/${ext};base64,${a.content.toString('base64')}`, filename: a.filename };
  }
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'universum-nahlad-'));
  const file = path.join(dir, a.filename.replace(/[\\/:*?"<>|]/g, '_'));
  fs.writeFileSync(file, a.content);
  if (/\.(pdf|txt|csv|log|html?)$/i.test(a.filename)) {
    const w = new BrowserWindow({ width: 900, height: 1000, title: `Náhľad – ${a.filename}`, autoHideMenuBar: true, webPreferences: { plugins: true, javascript: /\.pdf$/i.test(a.filename), sandbox: true } });
    void w.loadFile(file);
    return { kind: 'window' as const, filename: a.filename };
  }
  // Office a ostatné: otvoriť v programe, ktorý ich vie zobraziť (Word, Excel…)
  const err = await shell.openPath(file);
  if (err) throw new mail.UserError(`Prílohu sa nepodarilo otvoriť: ${err}`);
  return { kind: 'external' as const, filename: a.filename };
}

// ---------------------------------------------------------------------------
// odhlásenie z newslettera (RFC 2369 / RFC 8058)

export async function unsubscribe(acc: AccountConfig, u: NonNullable<MessageFull['unsubscribe']>) {
  if (u.url && u.oneClick) {
    const ok = await new Promise<boolean>((resolve) => {
      const body = 'List-Unsubscribe=One-Click';
      const req = https.request(u.url!, { method: 'POST', timeout: 15000, headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': String(body.length) } }, (res) => {
        res.resume();
        resolve((res.statusCode ?? 500) < 400);
      });
      req.on('error', () => resolve(false));
      req.on('timeout', () => req.destroy());
      req.end(body);
    });
    if (ok) return 'Odber zrušený.';
  }
  if (u.mailto) {
    const m = /^mailto:([^?]+)(?:\?(.*))?$/i.exec(u.mailto);
    if (m) {
      const q = new URLSearchParams(m[2] ?? '');
      await mail.sendMail(acc, { to: decodeURIComponent(m[1]!), subject: q.get('subject') ?? 'unsubscribe', text: q.get('body') ?? 'unsubscribe' });
      return 'Žiadosť o zrušenie odberu odoslaná e-mailom.';
    }
  }
  if (u.url) {
    await shell.openExternal(u.url);
    return 'Otvorená stránka na zrušenie odberu – dokončite ho v prehliadači.';
  }
  throw new mail.UserError('Správa neobsahuje spôsob odhlásenia.');
}

// ---------------------------------------------------------------------------
// pripomenutie, ak niekto neodpovie

interface FollowUp {
  accountId: string;
  messageId: string;
  subject: string;
  to: string;
  due: string;
}
const fuFile = () => path.join(app.getPath('userData'), 'followups.json');
const fuLoad = (): FollowUp[] => {
  try {
    return JSON.parse(fs.readFileSync(fuFile(), 'utf8')) as FollowUp[];
  } catch {
    return [];
  }
};
const fuSave = (l: FollowUp[]) => fs.writeFileSync(fuFile(), JSON.stringify(l, null, 2), 'utf8');

export function addFollowUp(f: FollowUp) {
  fuSave([...fuLoad().filter((x) => x.messageId !== f.messageId), f]);
}

async function checkFollowUps() {
  const now = Date.now();
  const list = fuLoad();
  const keep: FollowUp[] = [];
  for (const f of list) {
    try {
      const acc = store.getImap(f.accountId);
      const replied = await mail.withImap(acc, async (c) => {
        const lock = await c.getMailboxLock('INBOX');
        try {
          const a: number[] = (await c.search({ header: { 'in-reply-to': f.messageId } }, { uid: true })) || [];
          if (a.length) return true;
          const b: number[] = (await c.search({ header: { references: f.messageId } }, { uid: true })) || [];
          return b.length > 0;
        } finally {
          lock.release();
        }
      });
      if (replied) continue; // odpoveď prišla – pripomienka netreba
      if (Date.parse(f.due) > now) {
        keep.push(f);
        continue;
      }
      if (Notification.isSupported()) new Notification({ title: 'Bez odpovede', body: `${f.to} zatiaľ neodpovedal(a) na „${f.subject}“.` }).show();
      emit({ type: 'followup', subject: f.subject, to: f.to, accountId: f.accountId });
    } catch (e) {
      log('pripomenutie: kontrola zlyhala', String((e as Error).message));
      keep.push(f);
    }
  }
  fuSave(keep);
}

export function startFollowUps() {
  setTimeout(() => void checkFollowUps(), 120_000);
  setInterval(() => void checkFollowUps(), 30 * 60_000);
}

// ---------------------------------------------------------------------------
// záloha a obnova nastavení (bez hesiel a súkromných kľúčov)

export async function backupExport() {
  const parent = win();
  const opts = { title: 'Uložiť zálohu nastavení', defaultPath: path.join(os.homedir(), 'Documents', `universum-email-zaloha-${new Date().toISOString().slice(0, 10)}.json`), filters: [{ name: 'Záloha', extensions: ['json'] }] };
  const r = parent ? await dialog.showSaveDialog(parent, opts) : await dialog.showSaveDialog(opts);
  if (r.canceled || !r.filePath) return '';
  const data = { app: 'Universum eMail', version: app.getVersion(), created: new Date().toISOString(), accounts: store.loadAccounts(), settings: loadSettings() };
  fs.writeFileSync(r.filePath, JSON.stringify(data, null, 2), 'utf8');
  return r.filePath;
}

export async function backupImport() {
  const parent = win();
  const opts = { title: 'Obnoviť nastavenia zo zálohy', properties: ['openFile'] as Array<'openFile'>, filters: [{ name: 'Záloha', extensions: ['json'] }] };
  const r = parent ? await dialog.showOpenDialog(parent, opts) : await dialog.showOpenDialog(opts);
  if (r.canceled || !r.filePaths[0]) return { accounts: 0, restored: false };
  const data = JSON.parse(fs.readFileSync(r.filePaths[0], 'utf8')) as { accounts?: Parameters<typeof store.saveAccount>[0][]; settings?: ReturnType<typeof loadSettings> };
  if (!data || (!data.accounts && !data.settings)) throw new mail.UserError('Súbor nie je záloha Universum eMail.');
  let n = 0;
  const existing = store.loadAccounts();
  for (const a of data.accounts ?? []) {
    if (existing.some((x) => x.id === a.id)) continue;
    store.saveAccount(a);
    n++;
  }
  if (data.settings) saveSettings({ ...loadSettings(), ...data.settings });
  return { accounts: n, restored: true };
}

// ---------------------------------------------------------------------------
// ikona v oznamovacej oblasti a počet neprečítaných

let tray: Tray | null = null;
let quitting = false;
export const isQuitting = () => quitting;
export function markQuitting() {
  quitting = true;
}

function trayIcon() {
  const p = app.isPackaged ? path.join(process.resourcesPath, 'tray.png') : path.join(app.getAppPath(), 'assets', 'icon-32.png');
  const img = nativeImage.createFromPath(p);
  return img.isEmpty() ? nativeImage.createEmpty() : img.resize({ width: 16, height: 16 });
}

export function setupTray(showMain: () => void) {
  if (tray || process.platform === 'darwin') return;
  try {
    tray = new Tray(trayIcon());
  } catch (e) {
    log('tray: nepodarilo sa vytvoriť', String(e));
    return;
  }
  tray.setToolTip('Universum eMail');
  tray.setContextMenu(
    Menu.buildFromTemplate([
      { label: 'Otvoriť Universum eMail', click: showMain },
      { type: 'separator' },
      { label: 'Ukončiť', click: () => { quitting = true; app.quit(); } },
    ]),
  );
  tray.on('click', showMain);
}

/** Počet neprečítaných: popis ikony pri hodinách a odznak na paneli úloh (obrázok z rozhrania). */
export function setUnread(total: number, badgeDataUrl?: string) {
  tray?.setToolTip(total ? `Universum eMail – ${total} neprečítaných` : 'Universum eMail');
  const w = getMainWindow();
  if (w && process.platform === 'win32') {
    w.setOverlayIcon(total && badgeDataUrl ? nativeImage.createFromDataURL(badgeDataUrl) : null, total ? `${total} neprečítaných` : '');
  }
  if (process.platform !== 'win32') app.setBadgeCount(total);
}
