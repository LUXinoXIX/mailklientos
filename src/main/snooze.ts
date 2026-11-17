// Odložené správy: presun do priečinka „Odložené“ a v zadaný čas späť do Doručenej pošty ako neprečítané.
import fs from 'node:fs';
import path from 'node:path';
import { app, BrowserWindow } from 'electron';
import * as store from './accounts';
import { log } from './diagnostics';
import * as mail from './mail';

export const SNOOZE_FOLDER = 'Odložené';

interface Snoozed {
  accountId: string;
  messageId: string;
  subject: string;
  until: string;
}

const file = () => path.join(app.getPath('userData'), 'snoozed.json');
const load = (): Snoozed[] => {
  try {
    return JSON.parse(fs.readFileSync(file(), 'utf8')) as Snoozed[];
  } catch {
    return [];
  }
};
const save = (l: Snoozed[]) => fs.writeFileSync(file(), JSON.stringify(l, null, 2), 'utf8');

export async function snooze(accountId: string, folder: string, uid: number, messageId: string, subject: string, until: string) {
  if (!messageId) throw new mail.UserError('Správa nemá Message-ID, nedá sa odložiť.');
  const acc = store.getImap(accountId);
  const dest = await mail.ensureFolder(acc, SNOOZE_FOLDER);
  await mail.moveMessages(acc, folder, [uid], dest);
  save([...load().filter((x) => !(x.accountId === accountId && x.messageId === messageId)), { accountId, messageId, subject, until }]);
  return true as const;
}

let running = false;
async function tick() {
  if (running) return;
  running = true;
  try {
    const now = Date.now();
    const list = load();
    const due = list.filter((x) => Date.parse(x.until) <= now);
    if (!due.length) return;
    const left = [...list];
    for (const x of due) {
      try {
        const acc = store.getImap(x.accountId);
        const uid = await mail.findByMessageId(acc, SNOOZE_FOLDER, x.messageId);
        if (uid) {
          await mail.setFlags(acc, SNOOZE_FOLDER, [uid], { seen: false });
          await mail.moveMessages(acc, SNOOZE_FOLDER, [uid], 'INBOX');
        }
        left.splice(left.indexOf(x), 1);
        for (const w of BrowserWindow.getAllWindows()) w.webContents.send('app:event', { type: 'mail-changed', accountId: x.accountId, folder: 'INBOX', newCount: 1 });
      } catch (e) {
        log('odloženie: návrat zlyhal (skúsim neskôr)', x.subject, String((e as Error)?.message ?? e));
      }
    }
    save(left);
  } finally {
    running = false;
  }
}

export function startSnoozeTimer() {
  setTimeout(() => void tick(), 15_000);
  setInterval(() => void tick(), 60_000);
}
