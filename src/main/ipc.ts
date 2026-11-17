import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { BrowserWindow, dialog, ipcMain, shell } from 'electron';
import { accountPartition, type AccountConfig, type AnyAccount, type ServerConfig } from '../core/account';
import type { DraftRef, PickedFile, Result, SendInput } from '../core/api';
import type { AppSettings, LauncherItem, RuleAction } from '../core/settings';
import { autodetect, probeServer, type Protocol } from '../core/tls';
import * as store from './accounts';
import { autoconfig } from './autoconfig';
import * as mail from './mail';
import * as settings from './settings';
import * as web from './webviews';
import * as diag from './diagnostics';
import { syncWatchers } from './watch';
import * as cache from './cache';
import { applyActions, runRulesOnFolder } from './actions';
import { snooze } from './snooze';

/** Každá odpoveď ide ako { ok, data } alebo { ok:false, error, hint } – žiadne surové výnimky do UI. */
function handle<A extends unknown[], T>(channel: string, fn: (...args: A) => Promise<T> | T) {
  ipcMain.handle(channel, async (_e, ...args): Promise<Result<T>> => {
    try {
      return { ok: true, data: await fn(...(args as A)) };
    } catch (e) {
      const u = mail.toUserError(e);
      console.error(`[ipc ${channel}]`, e);
      return { ok: false, error: u.message, hint: u.hint, needsTrust: u.needsTrust };
    }
  });
}

/** Jednoduché volanie bez návratovej hodnoty (vložené weby). */
function on<A extends unknown[]>(channel: string, fn: (...args: A) => unknown) {
  ipcMain.handle(channel, async (_e, ...args) => {
    try {
      await fn(...(args as A));
    } catch (e) {
      console.error(`[ipc ${channel}]`, e);
    }
  });
}

const winOf = () => BrowserWindow.getFocusedWindow() ?? BrowserWindow.getAllWindows()[0] ?? null;
const safeName = (n: string) => n.replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_').slice(0, 180) || 'priloha';

export function registerIpc() {
  // kontá
  handle('accounts:list', () => ({ accounts: store.publicAccounts(), canStoreSecrets: store.canStoreSecrets() }));
  handle('accounts:save', (acc: AnyAccount, password?: string) => {
    mail.closeAllConnections(acc.id);
    const r = store.saveAccount(acc, password);
    syncWatchers();
    return r;
  });
  handle('accounts:delete', async (id: string) => {
    const acc = store.loadAccounts().find((a) => a.id === id);
    mail.closeAllConnections(id);
    store.deleteAccount(id);
    cache.removeFolder(id);
    syncWatchers();
    if (acc?.kind === 'm365') await web.clearPartition(accountPartition(id));
    return true as const;
  });
  handle('accounts:reorder', (ids: string[]) => {
    store.reorderAccounts(ids);
    return true as const;
  });
  handle('server:probe', (cfg: ServerConfig, proto: Protocol) => probeServer(cfg, proto));
  handle('server:autodetect', (host: string, proto: Protocol) => autodetect(host, proto));
  handle('server:autoconfig', (email: string) => autoconfig(email));
  handle('account:test', (acc: AccountConfig, password?: string) => mail.testAccount(acc, password));

  // pošta
  handle('mail:folders', (id: string) => mail.listFolders(store.getImap(id)));
  handle('mail:list', (id: string, folder: string, limit?: number) => mail.listMessages(store.getImap(id), folder, limit));
  handle('mail:listCached', (id: string, folder: string, limit?: number) => mail.listCached(id, folder, limit));
  handle('mail:search', (q: string, accountId?: string, folder?: string) => cache.search(q, { account: accountId, folder }));
  handle('mail:searchServer', (id: string, folder: string, q: string) => mail.searchServer(store.getImap(id), folder, q));
  handle('mail:keywords', (id: string, folder: string, uids: number[], add: string[], remove: string[]) =>
    mail.setKeywords(store.getImap(id), folder, uids, add, remove));
  handle('mail:thread', (id: string, messageId: string) => cache.threadOf(id, messageId));
  handle('mail:actions', async (id: string, folder: string, uids: number[], actions: RuleAction[]) => {
    await applyActions(store.getImap(id), folder, uids, actions);
    return true as const;
  });
  handle('mail:snooze', (id: string, folder: string, uid: number, messageId: string, subject: string, until: string) =>
    snooze(id, folder, uid, messageId, subject, until));
  handle('rules:run', (id: string, folder: string) => runRulesOnFolder(store.getImap(id), folder, settings.loadSettings().rules ?? []));
  handle('mail:get', (id: string, folder: string, uid: number) => mail.getMessage(store.getImap(id), folder, uid));
  handle('mail:send', (id: string, msg: SendInput) => mail.sendMail(store.getImap(id), msg));
  handle('mail:flags', async (id: string, folder: string, uids: number[], flags: { seen?: boolean; flagged?: boolean }) => {
    await mail.setFlags(store.getImap(id), folder, uids, flags);
    return true as const;
  });
  handle('mail:delete', (id: string, folder: string, uids: number[]) => mail.deleteMessages(store.getImap(id), folder, uids));
  handle('mail:move', async (id: string, folder: string, uids: number[], dest: string) => {
    await mail.moveMessages(store.getImap(id), folder, uids, dest);
    return true as const;
  });
  handle('mail:saveDraft', (id: string, msg: SendInput, prev?: DraftRef) => mail.saveDraft(store.getImap(id), msg, prev));
  handle('mail:deleteDraft', (id: string, d: DraftRef) => mail.deleteDraft(store.getImap(id), d));
  handle('mail:unread', (id: string) => mail.unreadCount(store.getImap(id)));
  handle('mail:archive', (id: string, folder: string, uids: number[]) => mail.archiveMessages(store.getImap(id), folder, uids));
  handle('mail:attachment-open', async (id: string, folder: string, uid: number, index: number) => {
    const a = await mail.getAttachment(store.getImap(id), folder, uid, index);
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'mailklient-'));
    const file = path.join(dir, safeName(a.filename));
    fs.writeFileSync(file, a.content);
    const err = await shell.openPath(file);
    if (err) throw new mail.UserError(`Prílohu sa nepodarilo otvoriť: ${err}`);
    return true as const;
  });
  handle('mail:attachment-save', async (id: string, folder: string, uid: number, index: number) => {
    const a = await mail.getAttachment(store.getImap(id), folder, uid, index);
    const w = winOf();
    const opts = { title: 'Uložiť prílohu', defaultPath: path.join(os.homedir(), 'Downloads', safeName(a.filename)) };
    const r = w ? await dialog.showSaveDialog(w, opts) : await dialog.showSaveDialog(opts);
    if (r.canceled || !r.filePath) return null;
    fs.writeFileSync(r.filePath, a.content);
    return r.filePath;
  });
  handle('files:pick', async (): Promise<PickedFile[]> => {
    const w = winOf();
    const opts: Electron.OpenDialogOptions = { title: 'Priložiť súbory', properties: ['openFile', 'multiSelections'] };
    const r = w ? await dialog.showOpenDialog(w, opts) : await dialog.showOpenDialog(opts);
    if (r.canceled) return [];
    return r.filePaths.map((p) => ({ path: p, filename: path.basename(p), size: fs.statSync(p).size }));
  });

  // nastavenia a spúšťače
  handle('settings:get', () => settings.loadSettings());
  handle('settings:save', (s: AppSettings) => settings.saveSettings(s));
  handle('signatures:outlook', () => settings.outlookSignatures());
  handle('launcher:pick', () => settings.pickProgram(winOf()));
  handle('launcher:run', (l: LauncherItem) => settings.runLauncher(l));
  handle('launcher:icon', (target: string) => settings.programIcon(target));

  ipcMain.handle('app:openExternal', async (_e, url: string) => {
    if (/^(https?|mailto):/i.test(url)) await shell.openExternal(url);
  });

  // stabilita vykresľovania a záznam chýb
  handle('diag:gpuDisabled', () => diag.gpuDisabled());
  handle('diag:gpuForced', () => diag.gpuForced());
  on('diag:setGpuForced', (on: boolean) => diag.setGpuForced(on));
  on('diag:setGpuDisabled', (off: boolean) => diag.setGpuDisabled(off));
  on('diag:openLog', () => diag.openLog());
  on('diag:log', (msg: string) => diag.log('renderer:', String(msg).slice(0, 2000)));

  // vložené weby
  on('web:clear', (partition: string) => web.clearPartition(partition));
  on('web:openOutside', (url: string) => web.openOutside(url));
}
