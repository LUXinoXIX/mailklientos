import { contextBridge, ipcRenderer } from 'electron';
import type { MailApi } from './core/api';

const inv = (channel: string, ...args: unknown[]) => ipcRenderer.invoke(channel, ...args);

const api: MailApi = {
  platform: process.platform,
  listAccounts: () => inv('accounts:list'),
  saveAccount: (acc, password) => inv('accounts:save', acc, password),
  deleteAccount: (id) => inv('accounts:delete', id),
  probe: (cfg, proto) => inv('server:probe', cfg, proto),
  autodetect: (host, proto) => inv('server:autodetect', host, proto),
  testAccount: (acc, password) => inv('account:test', acc, password),
  listFolders: (id) => inv('mail:folders', id),
  listMessages: (id, folder, limit) => inv('mail:list', id, folder, limit),
  getMessage: (id, folder, uid) => inv('mail:get', id, folder, uid),
  send: (id, msg) => inv('mail:send', id, msg),
  openExternal: (url) => inv('app:openExternal', url),
};

contextBridge.exposeInMainWorld('mailApi', api);
