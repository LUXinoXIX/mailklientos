// Synchronizácia úloh MailKlienta s úlohami na serveri (CalDAV VTODO – Úlohy v Zimbre, Nextcloud…).
import crypto from 'node:crypto';
import type { AccountConfig } from '../core/account';
import type { Task } from '../core/settings';
import * as store from './accounts';
import * as cal from './calendar';
import * as dav from './dav';
import { log } from './diagnostics';

async function todoList(acc: AccountConfig) {
  const cols = await cal.collections(acc);
  return cols.find((c) => c.kind === 'calendar' && c.todos && !c.readOnly) ?? null;
}

/** Zlúči lokálne úlohy so serverom; vráti nový zoznam (lokálne zmeny majú prednosť). */
export async function syncTasks(local: Task[]): Promise<{ tasks: Task[]; accounts: number }> {
  const accs = store.loadAccounts().filter((a) => a.kind === 'imap') as AccountConfig[];
  let tasks = [...local];
  let used = 0;
  for (const acc of accs) {
    let list;
    try {
      list = await todoList(acc);
    } catch {
      continue;
    }
    if (!list) continue;
    used++;
    try {
      const server = await dav.todos(acc, list);
      const byUid = new Map(server.map((t) => [t.uid, t]));
      // 1) lokálne úlohy tohto konta: odoslať zmenené, prevziať zmeny zo servera, zmazané na serveri odstrániť
      const next: Task[] = [];
      for (const t of tasks) {
        if (t.remote?.accountId !== acc.id) {
          next.push(t);
          continue;
        }
        const srv = byUid.get(t.remote.uid);
        byUid.delete(t.remote.uid);
        if (!srv) {
          if (t.dirty) {
            const saved = await dav.saveTodo(acc, list, { uid: t.remote.uid, summary: t.title, due: t.due, done: t.done, description: t.note });
            next.push({ ...t, dirty: false, remote: { ...t.remote, href: saved.href, etag: saved.etag } });
          }
          continue; // zmazaná na serveri
        }
        if (t.dirty) {
          const saved = await dav.saveTodo(acc, list, { uid: t.remote.uid, summary: t.title, due: t.due, done: t.done, description: t.note, href: srv.href });
          next.push({ ...t, dirty: false, remote: { ...t.remote, href: saved.href, etag: saved.etag } });
        } else next.push({ ...t, title: srv.summary || t.title, due: srv.due, done: srv.done, note: srv.description ?? t.note, remote: { ...t.remote, href: srv.href, etag: srv.etag } });
      }
      // 2) nové úlohy zo servera
      for (const srv of byUid.values())
        next.push({ id: crypto.randomUUID().slice(0, 8), title: srv.summary || '(bez názvu)', due: srv.due, done: srv.done, note: srv.description, created: new Date().toISOString(), remote: { accountId: acc.id, uid: srv.uid, href: srv.href, etag: srv.etag } });
      tasks = next;
    } catch (e) {
      log(`úlohy: synchronizácia ${acc.name} zlyhala`, String((e as Error).message));
    }
  }
  // 3) lokálne úlohy bez servera: na konto správy, inak na prvé konto s úlohami
  for (let i = 0; i < tasks.length; i++) {
    const t = tasks[i]!;
    if (t.remote) continue;
    const accId = t.message?.accountId;
    const acc = accs.find((a) => a.id === accId) ?? null;
    const target = acc && (await todoList(acc).catch(() => null)) ? acc : null;
    let chosen = target;
    if (!chosen) for (const a of accs) if (await todoList(a).catch(() => null)) { chosen = a; break; }
    if (!chosen) break;
    const list = await todoList(chosen);
    if (!list) continue;
    try {
      const uid = `${crypto.randomUUID()}@mailklient`;
      const saved = await dav.saveTodo(chosen, list, { uid, summary: t.title, due: t.due, done: t.done, description: t.note });
      tasks[i] = { ...t, dirty: false, remote: { accountId: chosen.id, uid, href: saved.href, etag: saved.etag } };
    } catch (e) {
      log('úlohy: odoslanie novej úlohy zlyhalo', String((e as Error).message));
    }
  }
  return { tasks, accounts: used };
}

export async function deleteRemote(t: Task) {
  if (!t.remote?.href) return true as const;
  await dav.deleteTodo(store.getImap(t.remote.accountId), t.remote.href);
  return true as const;
}
