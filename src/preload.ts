import { contextBridge, ipcRenderer, webUtils, type IpcRendererEvent } from 'electron';
import type { MailApi, WebEvent } from './core/api';

const inv = (channel: string, ...args: unknown[]) => ipcRenderer.invoke(channel, ...args);

const api: MailApi = {
  platform: process.platform,
  listAccounts: () => inv('accounts:list'),
  saveAccount: (acc, password) => inv('accounts:save', acc, password),
  deleteAccount: (id) => inv('accounts:delete', id),
  reorderAccounts: (ids) => inv('accounts:reorder', ids),
  probe: (cfg, proto) => inv('server:probe', cfg, proto),
  autodetect: (host, proto) => inv('server:autodetect', host, proto),
  testAccount: (acc, password) => inv('account:test', acc, password),
  autoconfig: (email) => inv('server:autoconfig', email),
  listFolders: (id) => inv('mail:folders', id),
  listMessages: (id, folder, limit) => inv('mail:list', id, folder, limit),
  getMessage: (id, folder, uid) => inv('mail:get', id, folder, uid),
  send: (id, msg) => inv('mail:send', id, msg),
  setFlags: (id, folder, uids, flags) => inv('mail:flags', id, folder, uids, flags),
  deleteMessages: (id, folder, uids) => inv('mail:delete', id, folder, uids),
  moveMessages: (id, folder, uids, dest) => inv('mail:move', id, folder, uids, dest),
  unreadCount: (id) => inv('mail:unread', id),
  saveDraft: (id, msg, prev) => inv('mail:saveDraft', id, msg, prev),
  deleteDraft: (id, d) => inv('mail:deleteDraft', id, d),
  pathForFile: (file) => webUtils.getPathForFile(file),
  archiveMessages: (id, folder, uids) => inv('mail:archive', id, folder, uids),
  openAttachment: (id, folder, uid, index) => inv('mail:attachment-open', id, folder, uid, index),
  saveAttachment: (id, folder, uid, index) => inv('mail:attachment-save', id, folder, uid, index),
  pickFiles: () => inv('files:pick'),
  openExternal: (url) => inv('app:openExternal', url),
  getSettings: () => inv('settings:get'),
  saveSettings: (s) => inv('settings:save', s),
  outlookSignatures: () => inv('signatures:outlook'),
  pickProgram: () => inv('launcher:pick'),
  runLauncher: (item) => inv('launcher:run', item),
  programIcon: (target) => inv('launcher:icon', target),
  webShow: (key, url, partition, bounds) => inv('web:show', key, url, partition, bounds),
  webBounds: (key, bounds) => inv('web:bounds', key, bounds),
  webHide: (key) => inv('web:hide', key),
  webHideAll: () => inv('web:hideAll'),
  webNavigate: (key, url) => inv('web:navigate', key, url),
  webReload: (key) => inv('web:reload', key),
  webBack: (key) => inv('web:back', key),
  webClose: (key) => inv('web:close', key),
  webOpenOutside: (key, fallbackUrl) => inv('web:openOutside', key, fallbackUrl),
  webClear: (partition) => inv('web:clear', partition),
  gpuDisabled: () => inv('diag:gpuDisabled'),
  gpuForced: () => inv('diag:gpuForced'),
  setGpuForced: (on) => inv('diag:setGpuForced', on),
  setGpuDisabled: (off) => inv('diag:setGpuDisabled', off),
  openLog: () => inv('diag:openLog'),
  logError: (msg) => inv('diag:log', msg),
  onAppEvent: (cb) => {
    const listener = (_e: IpcRendererEvent, ev: { type: 'gpu-crash'; count: number }) => cb(ev);
    ipcRenderer.on('app:event', listener);
    return () => ipcRenderer.removeListener('app:event', listener);
  },
  onWebEvent: (cb) => {
    const listener = (_e: IpcRendererEvent, ev: WebEvent) => cb(ev);
    ipcRenderer.on('web:event', listener);
    return () => ipcRenderer.removeListener('web:event', listener);
  },
};

contextBridge.exposeInMainWorld('mailApi', api);
