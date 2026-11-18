// Pošta na odoslanie: 10 s na späťvzatie („Zrušiť odoslanie“) a naplánované odoslanie („Odoslať neskôr“).
// Fronta je v userData/outbox.json, takže naplánované správy prežijú reštart (odošlú sa, keď MailKlient beží).
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { app, BrowserWindow, Notification } from 'electron';
import type { OutboxItem, SendInput } from '../core/api';
import * as store from './accounts';
import { log } from './diagnostics';
import * as mail from './mail';
import { addFollowUp } from './extras';

const file = () => path.join(app.getPath('userData'), 'outbox.json');
const load = (): OutboxItem[] => {
  try {
    return JSON.parse(fs.readFileSync(file(), 'utf8')) as OutboxItem[];
  } catch {
    return [];
  }
};
const save = (l: OutboxItem[]) => fs.writeFileSync(file(), JSON.stringify(l, null, 2), 'utf8');
const emit = (ev: unknown) => {
  for (const w of BrowserWindow.getAllWindows()) if (!w.isDestroyed()) w.webContents.send('app:event', ev);
};

export function queue(accountId: string, input: SendInput, sendAt: Date, kind: 'undo' | 'later'): OutboxItem {
  const item: OutboxItem = { id: crypto.randomUUID(), accountId, input, sendAt: sendAt.toISOString(), subject: input.subject, to: input.to, kind };
  save([...load(), item]);
  emit({ type: 'outbox-changed' });
  return item;
}

/** Zruší odoslanie; vráti správu, aby sa dala znova otvoriť na úpravu. */
export function cancel(id: string): OutboxItem | null {
  const l = load();
  const it = l.find((x) => x.id === id) ?? null;
  if (it) save(l.filter((x) => x.id !== id));
  emit({ type: 'outbox-changed' });
  return it;
}

export const list = () => load().filter((x) => x.kind === 'later');

const sending = new Set<string>();
async function tick() {
  const now = Date.now();
  for (const it of load()) {
    if (Date.parse(it.sendAt) > now || sending.has(it.id)) continue;
    sending.add(it.id);
    try {
      const acc = store.getImap(it.accountId);
      // pred odoslaním vyradiť z fronty (aby sa pri páde neodoslala dvakrát)
      save(load().filter((x) => x.id !== it.id));
      const r = await mail.sendMail(acc, it.input);
      if (it.input.followUpDays && r.messageId) {
        addFollowUp({ accountId: it.accountId, messageId: r.messageId, subject: it.subject, to: it.to, due: new Date(Date.now() + it.input.followUpDays * 86400_000).toISOString() });
      }
      emit({ type: 'outbox-sent', id: it.id, subject: it.subject, savedToSent: r.savedToSent });
    } catch (e) {
      const msg = String((e as Error)?.message ?? e);
      log('odoslanie z fronty zlyhalo', it.subject, msg);
      // vrátiť do fronty o 5 minút, používateľa upozorniť
      save([...load(), { ...it, sendAt: new Date(Date.now() + 5 * 60_000).toISOString(), kind: 'later', retry: true }]);
      emit({ type: 'outbox-failed', id: it.id, subject: it.subject, error: msg });
      if (Notification.isSupported()) new Notification({ title: 'Správu sa nepodarilo odoslať', body: `${it.subject}\n${msg}\nSkúsim znova o 5 minút.` }).show();
    } finally {
      sending.delete(it.id);
      emit({ type: 'outbox-changed' });
    }
  }
}

/** „Odoslať a prijať“: neodoslané správy (po chybe) skúsiť hneď, naplánované nechať na svoj čas. */
export async function flushRetries(): Promise<number> {
  const l = load();
  const n = l.filter((x) => x.retry).length;
  if (n) save(l.map((x) => (x.retry ? { ...x, sendAt: new Date().toISOString() } : x)));
  await tick();
  return n;
}

export function startOutbox() {
  setInterval(() => void tick(), 2000);
}
