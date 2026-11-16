import { ipcMain, shell } from 'electron';
import type { AccountConfig, ServerConfig } from '../core/account';
import type { Result, SendInput } from '../core/api';
import { autodetect, probeServer, type Protocol } from '../core/tls';
import * as store from './accounts';
import * as mail from './mail';

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

export function registerIpc() {
  handle('accounts:list', () => ({ accounts: store.publicAccounts(), canStoreSecrets: store.canStoreSecrets() }));
  handle('accounts:save', (acc: AccountConfig, password?: string) => store.saveAccount(acc, password));
  handle('accounts:delete', (id: string) => {
    store.deleteAccount(id);
    return true as const;
  });
  handle('server:probe', (cfg: ServerConfig, proto: Protocol) => probeServer(cfg, proto));
  handle('server:autodetect', (host: string, proto: Protocol) => autodetect(host, proto));
  handle('account:test', (acc: AccountConfig, password?: string) => mail.testAccount(acc, password));
  handle('mail:folders', (id: string) => mail.listFolders(store.getAccount(id)));
  handle('mail:list', (id: string, folder: string, limit?: number) => mail.listMessages(store.getAccount(id), folder, limit));
  handle('mail:get', (id: string, folder: string, uid: number) => mail.getMessage(store.getAccount(id), folder, uid));
  handle('mail:send', (id: string, msg: SendInput) => mail.sendMail(store.getAccount(id), msg));

  ipcMain.handle('app:openExternal', async (_e, url: string) => {
    if (/^(https?|mailto):/i.test(url)) await shell.openExternal(url);
  });
}
