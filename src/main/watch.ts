// Okamžité doručenie: pre každé IMAP konto samostatné spojenie, ktoré drží otvorenú Doručenú poštu v režime IDLE.
// Pri novej správe: udalosť pre rozhranie (obnoviť zoznam a počty) a upozornenie Windows s odosielateľom a predmetom.
import { BrowserWindow, Notification } from 'electron';
import type { AccountConfig } from '../core/account';
import * as store from './accounts';
import { log } from './diagnostics';
import { closeClient, connectWatcher } from './mail';
import { loadSettings } from './settings';
import { runRules } from './actions';
import { maybeAutoReply } from './vacation';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Client = any;

interface Watch {
  id: string;
  key: string;
  client?: Client;
  stopped: boolean;
  lastUid: number;
  retry?: ReturnType<typeof setTimeout>;
  delay: number;
}

const watches = new Map<string, Watch>();

function send(ev: unknown) {
  for (const w of BrowserWindow.getAllWindows()) if (!w.isDestroyed()) w.webContents.send('app:event', ev);
}

function mainWindow() {
  return BrowserWindow.getAllWindows().find((w) => !w.isDestroyed()) ?? null;
}

function notify(acc: AccountConfig, items: Array<{ uid: number; from: string; subject: string }>) {
  const s = loadSettings();
  if (s.notifications === false || !Notification.isSupported()) return;
  const win = mainWindow();
  if (win && win.isFocused()) return; // pri práci v aplikácii stačí obnovený zoznam
  win?.flashFrame(true);
  const shown = items.slice(-3);
  for (const m of shown) {
    const n = new Notification({
      title: m.from || acc.name,
      subtitle: acc.name,
      body: `${m.subject || '(bez predmetu)'}\n${acc.name}`,
      silent: false,
    });
    n.on('click', () => {
      const w = mainWindow();
      if (w) {
        if (w.isMinimized()) w.restore();
        w.show();
        w.focus();
      }
      send({ type: 'open-message', accountId: acc.id, folder: 'INBOX', uid: m.uid });
    });
    n.show();
  }
  if (items.length > shown.length) {
    new Notification({ title: acc.name, body: `Ďalšie nové správy: ${items.length - shown.length}` }).show();
  }
}

async function fetchNew(w: Watch, acc: AccountConfig) {
  const c = w.client;
  if (!c) return;
  const items: Array<{ uid: number; from: string; subject: string }> = [];
  const forRules: Array<{ uid: number; from: string; to: string; subject: string }> = [];
  const lock = await c.getMailboxLock('INBOX');
  try {
    for await (const m of c.fetch(`${w.lastUid + 1}:*`, { uid: true, envelope: true, flags: true, headers: ['list-id', 'list-unsubscribe', 'precedence', 'auto-submitted'] }, { uid: true })) {
      if (m.uid <= w.lastUid) continue;
      w.lastUid = Math.max(w.lastUid, m.uid);
      const f = m.envelope?.from?.[0];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const to = [...(m.envelope?.to ?? []), ...(m.envelope?.cc ?? [])].map((x: any) => `${x.name ?? ''} ${x.address ?? ''}`).join(' ');
      forRules.push({ uid: m.uid, from: `${f?.name ?? ''} ${f?.address ?? ''}`, to, subject: m.envelope?.subject ?? '' });
      if (m.flags?.has('\\Seen')) continue;
      void maybeAutoReply(acc, { fromAddress: f?.address ?? '', subject: m.envelope?.subject ?? '', messageId: m.envelope?.messageId, headers: m.headers ? m.headers.toString() : '' });
      items.push({ uid: m.uid, from: f?.name || f?.address || '', subject: m.envelope?.subject ?? '' });
    }
  } finally {
    lock.release();
  }
  const rules = loadSettings().rules ?? [];
  if (forRules.length && rules.length) {
    const hits = await runRules(acc, 'INBOX', rules, forRules).catch(() => 0);
    if (hits) log(`pravidlá: použité na ${hits} nových správ (${acc.name})`);
  }
  send({ type: 'mail-changed', accountId: acc.id, folder: 'INBOX', newCount: items.length });
  if (items.length) notify(acc, items);
}

function schedule(w: Watch) {
  if (w.stopped) return;
  if (w.retry) clearTimeout(w.retry);
  w.retry = setTimeout(() => void start(w), w.delay);
  w.delay = Math.min(w.delay * 2, 5 * 60_000);
}

async function start(w: Watch) {
  if (w.stopped) return;
  let acc: AccountConfig;
  try {
    acc = store.getImap(w.id);
  } catch {
    return;
  }
  try {
    const c = await connectWatcher(acc);
    if (w.stopped) return void closeClient(c);
    w.client = c;
    const box = await c.mailboxOpen('INBOX', { readOnly: true });
    // pri prvom spojení začíname od aktuálneho stavu (staré správy neohlasovať)
    if (!w.lastUid) w.lastUid = Math.max(0, Number(box.uidNext ?? 1) - 1);
    else await fetchNew(w, acc); // po výpadku dohnať zmeškané
    w.delay = 5000;
    log(`sledovanie pošty: ${acc.name} pripojené (IDLE)`);
    let busy = Promise.resolve();
    const later = (fn: () => Promise<void>) => {
      busy = busy.then(fn).catch((e) => log('sledovanie pošty: chyba', String(e)));
    };
    c.on('exists', (d: { path: string; count: number; prevCount: number }) => {
      if (d.count > d.prevCount) later(() => fetchNew(w, acc));
    });
    c.on('expunge', () => send({ type: 'mail-changed', accountId: acc.id, folder: 'INBOX', newCount: 0 }));
    c.on('flags', () => send({ type: 'mail-changed', accountId: acc.id, folder: 'INBOX', newCount: 0 }));
    c.on('close', () => {
      if (w.client === c) w.client = undefined;
      if (!w.stopped) {
        log(`sledovanie pošty: ${acc.name} odpojené, nové spojenie o ${Math.round(w.delay / 1000)} s`);
        schedule(w);
      }
    });
  } catch (e) {
    log(`sledovanie pošty: ${acc.name} sa nepripojilo`, String((e as Error)?.message ?? e));
    w.client = undefined;
    schedule(w);
  }
}

function stop(w: Watch) {
  w.stopped = true;
  if (w.retry) clearTimeout(w.retry);
  if (w.client) void closeClient(w.client);
  w.client = undefined;
}

const keyOf = (a: AccountConfig) => JSON.stringify([a.imap, a.username, store.getPassword(a.id) ? 1 : 0]);

/** Zladí sledovanie s aktuálnym zoznamom kont (pri štarte a po každej zmene kont). */
export function syncWatchers() {
  const accs = store.loadAccounts().filter((a) => a.kind === 'imap') as AccountConfig[];
  const ids = new Set(accs.map((a) => a.id));
  for (const [id, w] of watches) {
    const a = accs.find((x) => x.id === id);
    if (!ids.has(id) || !a || keyOf(a) !== w.key) {
      stop(w);
      watches.delete(id);
    }
  }
  for (const a of accs) {
    if (watches.has(a.id) || !store.getPassword(a.id)) continue;
    const w: Watch = { id: a.id, key: keyOf(a), stopped: false, lastUid: 0, delay: 5000 };
    watches.set(a.id, w);
    void start(w);
  }
}

export function stopAllWatchers() {
  for (const w of watches.values()) stop(w);
  watches.clear();
}

/** Po prebudení počítača: spojenia bývajú mŕtve, nadviazať nové hneď. */
export function reconnectWatchers() {
  for (const w of watches.values()) {
    if (w.client) void closeClient(w.client);
    w.client = undefined;
    w.delay = 2000;
    schedule(w);
  }
}
