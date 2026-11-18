import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { app, BrowserWindow, dialog, ipcMain, shell } from 'electron';
import { accountPartition, type AccountConfig, type AnyAccount, type ServerConfig } from '../core/account';
import type { DraftRef, PickedFile, Result, SendInput } from '../core/api';
import type { AppSettings, LauncherItem, RuleAction, Task } from '../core/settings';
import { autodetect, probeServer, type Protocol } from '../core/tls';
import * as store from './accounts';
import { autoconfig } from './autoconfig';
import * as mail from './mail';
import * as settings from './settings';
import * as web from './webviews';
import * as diag from './diagnostics';
import { reconnectWatchers, syncWatchers, testNotification } from './watch';
import * as cache from './cache';
import { applyActions, runRulesOnFolder } from './actions';
import { snooze } from './snooze';
import * as outbox from './outbox';
import * as updater from './updater';
import * as io from './importexport';
import * as cstore from './cryptostore';
import * as zimbra from './zimbra';
import * as tasksync from './tasksync';
import * as extras from './extras';
import * as cal from './calendar';
import * as dav from './dav';
import type { CalEvent, Contact } from '../core/ical';

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

let openComposeWindow: ((q: Record<string, string>) => void) | undefined;
/** main.ts odovzdá funkciu na otvorenie okna písania (samostatné okno konceptu). */
export function setComposeWindowFactory(f: (q: Record<string, string>) => void) {
  openComposeWindow = f;
}

export function registerIpc() {
  // kontá
  handle('accounts:list', () => ({ accounts: store.publicAccounts(), canStoreSecrets: store.canStoreSecrets() }));
  handle('accounts:save', (acc: AnyAccount, password?: string) => {
    mail.closeAllConnections(acc.id);
    cal.forget(acc.id);
    zimbra.forget(acc.id);
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
  handle('mail:list', (id: string, folder: string, limit?: number, preview?: boolean) => mail.listMessages(store.getImap(id), folder, limit, preview));
  handle('mail:listCached', (id: string, folder: string, limit?: number) => mail.listCached(id, folder, limit));
  handle('mail:search', (q: string, accountId?: string, folder?: string, field?: string) => cache.search(q, { account: accountId, folder, field }));
  handle('mail:searchServer', (id: string, folder: string, q: string, field?: string) => mail.searchServer(store.getImap(id), folder, q, 200, field));
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
  handle('mail:openDraft', (id: string, folder: string, uid: number) => mail.openDraft(store.getImap(id), folder, uid));
  handle('outbox:queue', (id: string, input: SendInput, sendAt: string, kind: 'undo' | 'later') => outbox.queue(id, input, new Date(sendAt), kind));
  handle('outbox:cancel', (outId: string) => outbox.cancel(outId));
  handle('outbox:list', () => outbox.list());
  handle('window:message', (id: string, folder: string, uid: number) => {
    openComposeWindow?.({ view: '1', acc: id, folder, uid: String(uid) });
    return true as const;
  });
  handle('window:reply', (id: string, folder: string, uid: number, mode: string) => {
    openComposeWindow?.({ compose: '1', mode, acc: id, folder, uid: String(uid) });
    return true as const;
  });
  // zmena správy v inom okne → ostatné okná obnovia zoznam
  handle('mail:changed', (id: string, folder: string) => {
    for (const w of BrowserWindow.getAllWindows()) if (!w.isDestroyed()) w.webContents.send('app:event', { type: 'mail-changed', accountId: id, folder, newCount: 0 });
    return true as const;
  });
  handle('mail:sendReceive', async () => {
    reconnectWatchers();
    return { retried: await outbox.flushRetries() };
  });
  handle('window:compose', (id: string, folder: string, uid: number) => {
    openComposeWindow?.({ compose: '1', acc: id, folder, uid: String(uid) });
    return true as const;
  });
  handle('dav:collections', (id: string) => cal.collections(store.getImap(id), true));
  handle('cal:events', (id: string, ids: string[], from: string, to: string) => cal.events(store.getImap(id), ids, from, to));
  handle('cal:save', (id: string, calendarId: string, e: CalEvent, invite: boolean) => cal.save(store.getImap(id), calendarId, e, invite));
  handle('cal:delete', (id: string, e: CalEvent, notify: boolean) => cal.remove(store.getImap(id), e, notify));
  handle('cal:respond', (id: string, folder: string, uid: number, ps: 'ACCEPTED' | 'TENTATIVE' | 'DECLINED') => cal.respond(store.getImap(id), folder, uid, ps));
  handle('card:list', (id: string) => cal.contactList(store.getImap(id)));
  handle('card:save', (id: string, bookId: string, c: Contact) => cal.contactSave(store.getImap(id), bookId, c));
  handle('card:delete', async (id: string, href: string) => {
    await dav.deleteContact(store.getImap(id), href);
    return true as const;
  });
  handle('io:exportEml', (id: string, folder: string, uid: number, subject: string) => io.exportEml(store.getImap(id), folder, uid, subject));
  handle('io:exportMbox', (id: string, folder: string) => io.exportMbox(store.getImap(id), folder));
  handle('io:importFiles', (id: string, folder: string) => io.importFiles(store.getImap(id), folder));
  handle('io:importThunderbird', (id: string) => io.importThunderbird(store.getImap(id)));
  handle('crypto:summary', () => cstore.summary());
  handle('crypto:pgpGenerate', (name: string, email: string, pass: string) => cstore.pgpGenerate(name, email, pass));
  handle('crypto:pick', async (kind: 'pgp' | 'p12' | 'cert') => {
    const w = winOf();
    const filters =
      kind === 'pgp' ? [{ name: 'Kľúč PGP', extensions: ['asc', 'gpg', 'pgp', 'key', 'txt'] }]
      : kind === 'p12' ? [{ name: 'Certifikát s kľúčom', extensions: ['p12', 'pfx'] }]
      : [{ name: 'Certifikát', extensions: ['cer', 'crt', 'pem', 'der'] }];
    const opts = { title: 'Vyberte súbor', properties: ['openFile'] as Array<'openFile'>, filters };
    const r = w ? await dialog.showOpenDialog(w, opts) : await dialog.showOpenDialog(opts);
    return r.canceled ? '' : r.filePaths[0] ?? '';
  });
  handle('crypto:pgpImportFile', (file: string, pass?: string) => cstore.pgpImport(fs.readFileSync(file, 'utf8'), pass));
  handle('crypto:pgpImportText', (text: string, pass?: string) => cstore.pgpImport(text, pass));
  handle('crypto:smimeImportFile', (file: string, pass: string) => {
    try {
      return cstore.smimeImportP12(fs.readFileSync(file), pass);
    } catch (e) {
      throw new mail.UserError(/password|MAC|Invalid/i.test(String((e as Error).message)) ? 'Nesprávne heslo k súboru .p12.' : (e as Error).message);
    }
  });
  handle('crypto:certImportFile', (file: string) => cstore.smimeImportCert(fs.readFileSync(file)));
  handle('crypto:pgpExport', async (fp: string) => {
    const w = winOf();
    const opts = { title: 'Uložiť verejný kľúč', defaultPath: path.join(os.homedir(), 'Documents', `verejny-kluc-${fp.slice(-8)}.asc`) };
    const r = w ? await dialog.showSaveDialog(w, opts) : await dialog.showSaveDialog(opts);
    if (r.canceled || !r.filePath) return '';
    fs.writeFileSync(r.filePath, cstore.pgpExportPublic(fp), 'utf8');
    return r.filePath;
  });
  handle('crypto:remove', (kind: 'pgpOwn' | 'pgpContacts' | 'smimeOwn' | 'smimeContacts', id: string) => {
    cstore.remove(kind, id);
    return true as const;
  });
  // Zimbra
  const zacc = (id: string) => store.getImap(id);
  handle('zimbra:check', (id: string) => zimbra.isZimbra(zacc(id)));
  handle('zimbra:gal', (id: string, q: string) => zimbra.galSearch(zacc(id), q));
  handle('zimbra:galAuto', (id: string, q: string) => zimbra.galAutocomplete(zacc(id), q));
  handle('zimbra:freeBusy', (id: string, emails: string[], from: number, to: number) => zimbra.freeBusy(zacc(id), emails, from, to));
  handle('zimbra:resources', (id: string) => zimbra.resources(zacc(id)));
  handle('zimbra:shares', (id: string) => zimbra.shares(zacc(id)));
  handle('zimbra:mount', (id: string, sh: zimbra.Share, name: string) => zimbra.mountShare(zacc(id), sh, name));
  handle('zimbra:getOoo', (id: string) => zimbra.getOoo(zacc(id)));
  handle('zimbra:setOoo', (id: string, o: zimbra.ServerOoo) => zimbra.setOoo(zacc(id), o));
  handle('zimbra:getFilters', (id: string) => zimbra.getFilters(zacc(id)));
  handle('zimbra:setFilters', (id: string, r: zimbra.ServerRule[]) => zimbra.setFilters(zacc(id), r));
  handle('zimbra:signatures', (id: string) => zimbra.getSignatures(zacc(id)));
  handle('zimbra:saveSignature', (id: string, sig: { id?: string; name: string; html: string; text: string }) => zimbra.saveSignature(zacc(id), sig));
  handle('zimbra:identities', (id: string) => zimbra.getIdentities(zacc(id)));
  handle('zimbra:tags', (id: string) => zimbra.tags(zacc(id)));
  handle('zimbra:briefcase', (id: string, q: string) => zimbra.briefcase(zacc(id), q));
  handle('zimbra:download', (id: string, d: zimbra.BriefcaseDoc) => zimbra.downloadDoc(zacc(id), d));
  handle('zimbra:quota', async (id: string) => {
    const acc = zacc(id);
    if (await zimbra.isZimbra(acc)) {
      const q = await zimbra.quota(acc).catch(() => null);
      if (q && q.limit) return q;
    }
    return mail.imapQuota(acc).catch(() => null);
  });
  handle('zimbra:changePassword', async (id: string, oldPass: string, newPass: string) => {
    const acc = zacc(id);
    await zimbra.changePassword(acc, oldPass, newPass);
    mail.closeAllConnections(id);
    store.saveAccount(store.getAccount(id), newPass);
    syncWatchers();
    return true as const;
  });
  handle('zimbra:getWB', (id: string) => zimbra.getWhiteBlack(zacc(id)));
  handle('zimbra:setWB', (id: string, w: string[], b: string[]) => zimbra.setWhiteBlack(zacc(id), w, b));
  handle('zimbra:block', (id: string, email: string) => zimbra.blockSender(zacc(id), email));
  handle('tasks:sync', (tasks: Task[]) => tasksync.syncTasks(tasks));
  handle('tasks:delete', (t: Task) => tasksync.deleteRemote(t));
  handle('app:info', () => ({ version: app.getVersion(), electron: process.versions.electron, chrome: process.versions.chrome, node: process.versions.node, platform: `${process.platform} ${process.arch}` }));
  handle('mail:print', (id: string, folder: string, uid: number, pdf: boolean) => extras.printMessage(store.getImap(id), folder, uid, pdf));
  handle('mail:preview', (id: string, folder: string, uid: number, index: number) => extras.previewAttachment(store.getImap(id), folder, uid, index));
  handle('mail:unsubscribe', (id: string, u: { url?: string; mailto?: string; oneClick?: boolean }) => extras.unsubscribe(store.getImap(id), u));
  handle('backup:export', () => extras.backupExport());
  handle('backup:import', () => extras.backupImport());
  handle('app:unread', (total: number, badge?: string) => {
    extras.setUnread(total, badge);
    return true as const;
  });
  handle('web:openOffice', async (url: string) => {
    if (!/^https:\/\//i.test(url)) throw new mail.UserError('Neplatná adresa dokumentu.');
    const app = /\.xlsx?(\b|$)|excel|:x:/i.test(url) ? 'ms-excel' : /\.pptx?(\b|$)|powerpoint|:p:/i.test(url) ? 'ms-powerpoint' : 'ms-word';
    await shell.openExternal(`${app}:ofe|u|${url}`);
    return true as const;
  });
  handle('diag:testNotification', () => testNotification());
  handle('update:check', () => updater.check());
  handle('update:install', () => {
    updater.installNow();
    return true as const;
  });
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
