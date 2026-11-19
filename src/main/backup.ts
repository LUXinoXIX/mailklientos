// Záloha / prenos nastavení medzi počítačmi (aj Windows ↔ Mac): jeden súbor JSON.
// Heslá sa NEexportujú (sú v systémovom úložisku kľúčov) – po importe ich treba zadať znova.
// Programy v pravom paneli sa pri importe prevezmú len vtedy, ak na tomto počítači existujú.
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { app, BrowserWindow, dialog } from 'electron';
import type { AnyAccount } from '../core/account';
import type { AppSettings, LauncherItem } from '../core/settings';
import { loadAccounts } from './accounts';
import { isUrlScheme, loadSettings, saveSettings } from './settings';
import { UserError } from './mail';

export interface BackupParts {
  accounts: boolean;
  contacts: boolean;
  calendars: boolean;
}

interface BackupFile {
  format: 'universum-email-backup';
  version: 1;
  app: string;
  platform: string;
  exported: string;
  settings: AppSettings;
  accounts?: AnyAccount[];
  contacts?: unknown;
  recent?: unknown;
  calendars?: unknown;
}

export interface BackupInfo {
  token: string;
  file: string;
  exported: string;
  platform: string;
  app: string;
  has: { accounts: number; contacts: number; calendars: number };
  /** kontá zo zálohy, ktoré tu už sú (podľa id alebo adresy) */
  accountsExisting: number;
  launchers: { name: string; ok: boolean }[];
}

const ud = (f: string) => path.join(app.getPath('userData'), f);
const readJson = (f: string): unknown => {
  try {
    return JSON.parse(fs.readFileSync(ud(f), 'utf8'));
  } catch {
    return undefined;
  }
};
const writeJson = (f: string, data: unknown) => {
  const tmp = ud(f) + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(data, null, 1), 'utf8');
  fs.renameSync(tmp, ud(f));
};
const pending = new Map<string, BackupFile>();
const accKey = (a: AnyAccount) => `${a.kind}|${(('email' in a && a.email) || ('user' in a && (a as { user?: string }).user) || '').toLowerCase()}`;
const launcherOk = (l: LauncherItem) => isUrlScheme(l.target.trim()) || fs.existsSync(l.target.trim());

export async function exportBackup(win: BrowserWindow | null, parts: BackupParts): Promise<string | null> {
  const data: BackupFile = {
    format: 'universum-email-backup',
    version: 1,
    app: app.getVersion(),
    platform: process.platform,
    exported: new Date().toISOString(),
    settings: loadSettings(),
  };
  if (parts.accounts) data.accounts = loadAccounts();
  if (parts.contacts) {
    data.contacts = readJson('contacts.json');
    data.recent = readJson('recent.json');
  }
  if (parts.calendars) data.calendars = readJson('calendars.json');
  const opts: Electron.SaveDialogOptions = {
    title: 'Exportovať nastavenia',
    defaultPath: path.join(app.getPath('documents'), `Universum-nastavenia-${new Date().toISOString().slice(0, 10)}.json`),
    filters: [{ name: 'Nastavenia Universum eMail', extensions: ['json'] }],
  };
  const r = win ? await dialog.showSaveDialog(win, opts) : await dialog.showSaveDialog(opts);
  if (r.canceled || !r.filePath) return null;
  fs.writeFileSync(r.filePath, JSON.stringify(data, null, 1), 'utf8');
  return r.filePath;
}

export async function pickBackup(win: BrowserWindow | null): Promise<BackupInfo | null> {
  const opts: Electron.OpenDialogOptions = {
    title: 'Importovať nastavenia',
    properties: ['openFile'],
    filters: [{ name: 'Nastavenia Universum eMail', extensions: ['json'] }],
  };
  const r = win ? await dialog.showOpenDialog(win, opts) : await dialog.showOpenDialog(opts);
  if (r.canceled || !r.filePaths[0]) return null;
  let data: BackupFile;
  try {
    data = JSON.parse(fs.readFileSync(r.filePaths[0], 'utf8')) as BackupFile;
  } catch {
    throw new UserError('Súbor sa nedá prečítať (nie je to JSON).');
  }
  if (data?.format !== 'universum-email-backup' || !data.settings) throw new UserError('Toto nie je súbor s nastaveniami Universum eMail.');
  const token = crypto.randomUUID();
  pending.set(token, data);
  const mine = loadAccounts();
  const ids = new Set(mine.map((a) => a.id));
  const keys = new Set(mine.map(accKey));
  const allLaunchers = [...(data.settings.launchers ?? []), ...(data.settings.rightPanels ?? []).flatMap((p) => p.launchers ?? [])];
  const cal = data.calendars as { calendars?: unknown[] } | undefined;
  const con = data.contacts as { contacts?: unknown[] } | undefined;
  return {
    token,
    file: path.basename(r.filePaths[0]),
    exported: data.exported,
    platform: data.platform,
    app: data.app,
    has: { accounts: data.accounts?.length ?? 0, contacts: con?.contacts?.length ?? 0, calendars: cal?.calendars?.length ?? 0 },
    accountsExisting: (data.accounts ?? []).filter((a) => ids.has(a.id) || keys.has(accKey(a))).length,
    launchers: allLaunchers.map((l) => ({ name: l.name, ok: launcherOk(l) })),
  };
}

/** Spojí programy: tie, čo tu už sú, + zo zálohy len existujúce (bez duplicít podľa cesty). */
function mergeLaunchers(local: LauncherItem[] | undefined, imported: LauncherItem[] | undefined): LauncherItem[] {
  const out = [...(local ?? [])];
  for (const l of imported ?? []) if (launcherOk(l) && !out.some((x) => x.target === l.target)) out.push(l);
  return out;
}

export function applyBackup(token: string, parts: BackupParts & { settings: boolean }): void {
  const data = pending.get(token);
  if (!data) throw new UserError('Import vypršal – vyber súbor znova.');
  if (parts.settings) {
    const cur = loadSettings();
    const s: AppSettings = { ...structuredClone(data.settings) };
    s.launchers = mergeLaunchers(cur.launchers, data.settings.launchers);
    s.rightPanels = (data.settings.rightPanels ?? []).map((p) => ({
      ...p,
      launchers: mergeLaunchers(cur.rightPanels?.find((x) => x.id === p.id)?.launchers, p.launchers),
    }));
    // vlastné panely, ktoré sú len tu, ponechať
    for (const p of cur.rightPanels ?? []) if (!s.rightPanels.some((x) => x.id === p.id)) s.rightPanels.push(p);
    saveSettings(s);
  }
  if (parts.accounts && data.accounts?.length) {
    const mine = loadAccounts();
    const ids = new Set(mine.map((a) => a.id));
    const keys = new Set(mine.map(accKey));
    const add = data.accounts.filter((a) => !ids.has(a.id) && !keys.has(accKey(a)));
    if (add.length) writeJson('accounts.json', [...mine, ...add]);
  }
  if (parts.contacts && data.contacts) {
    const cur = (readJson('contacts.json') as { contacts?: { uid?: string }[] } | undefined)?.contacts ?? [];
    const inc = (data.contacts as { contacts?: { uid?: string }[] }).contacts ?? [];
    const uids = new Set(cur.map((c) => c.uid));
    writeJson('contacts.json', { version: 1, contacts: [...cur, ...inc.filter((c) => !uids.has(c.uid))] });
    if (data.recent && !fs.existsSync(ud('recent.json'))) writeJson('recent.json', data.recent);
  }
  if (parts.calendars && data.calendars) {
    const cur = (readJson('calendars.json') as { calendars?: { id: string }[]; events?: Record<string, unknown> } | undefined) ?? {};
    const inc = data.calendars as { calendars?: { id: string }[]; events?: Record<string, unknown> };
    const have = new Set((cur.calendars ?? []).map((c) => c.id));
    const newCals = (inc.calendars ?? []).filter((c) => !have.has(c.id));
    const events = { ...(cur.events ?? {}) };
    for (const c of newCals) if (inc.events?.[c.id]) events[c.id] = inc.events[c.id];
    writeJson('calendars.json', { ...cur, calendars: [...(cur.calendars ?? []), ...newCals], events });
  }
  pending.delete(token);
  // reštart – všetky moduly načítajú nové súbory
  setTimeout(() => {
    app.relaunch();
    app.exit(0);
  }, 300);
}
