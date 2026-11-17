import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { BrowserWindow, dialog, ipcMain, shell } from 'electron';
import { accountPartition, type AccountConfig, type AnyAccount, type ServerConfig } from '../core/account';
import type { PickedFile, Result, SendInput, WebBounds } from '../core/api';
import type { AppSettings, LauncherItem } from '../core/settings';
import { autodetect, probeServer, type Protocol } from '../core/tls';
import * as store from './accounts';
import { autoconfig } from './autoconfig';
import * as mail from './mail';
import * as settings from './settings';
import * as web from './webviews';
import * as diag from './diagnostics';

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
  handle('accounts:save', (acc: AnyAccount, password?: string) => store.saveAccount(acc, password));
  handle('accounts:delete', async (id: string) => {
    const acc = store.loadAccounts().find((a) => a.id === id);
    store.deleteAccount(id);
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
  on('diag:setGpuDisabled', (off: boolean) => diag.setGpuDisabled(off));
  on('diag:openLog', () => diag.openLog());
  on('diag:log', (msg: string) => diag.log('renderer:', String(msg).slice(0, 2000)));

  // vložené weby
  on('web:show', (key: string, url: string, partition: string, b: WebBounds) => web.show(key, url, partition, b));
  on('web:bounds', (key: string, b: WebBounds) => web.setBounds(key, b));
  on('web:hide', (key: string) => web.hide(key));
  on('web:hideAll', () => web.hideAll());
  on('web:navigate', (key: string, url: string) => web.navigate(key, url));
  on('web:reload', (key: string) => web.reload(key));
  on('web:back', (key: string) => web.back(key));
  on('web:close', (key: string) => web.destroy(key));
  on('web:clear', (partition: string) => web.clearPartition(partition));
  on('web:openOutside', async (key: string, fallback: string) => {
    const url = web.currentUrl(key) || fallback;
    if (!/^https?:/i.test(url)) return;
    // Windows: Edge cez jeho URL schému (funguje aj keď je predvolený iný prehliadač).
    if (process.platform === 'win32') {
      try {
        await shell.openExternal(`microsoft-edge:${url}`);
        return;
      } catch {
        /* Edge nie je: predvolený prehliadač */
      }
    }
    await shell.openExternal(url);
  });
}
