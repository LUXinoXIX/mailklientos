// Export a import pošty: správa ako .eml, priečinok ako .mbox; import .eml/.mbox súborov a priečinkov Thunderbirdu
// (mbox súbory profilu) do IMAP priečinkov.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import readline from 'node:readline';
import { BrowserWindow, dialog } from 'electron';
import type { AccountConfig } from '../core/account';
import { log } from './diagnostics';
import { UserError, withImap } from './mail';

const safe = (n: string) => n.replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_').slice(0, 150) || 'sprava';
const win = () => BrowserWindow.getFocusedWindow() ?? BrowserWindow.getAllWindows()[0] ?? null;
const progress = (text: string) => {
  for (const w of BrowserWindow.getAllWindows()) if (!w.isDestroyed()) w.webContents.send('app:event', { type: 'progress', text });
};

export async function exportEml(acc: AccountConfig, folder: string, uid: number, subject: string) {
  const w = win();
  const opts = { title: 'Uložiť správu', defaultPath: path.join(os.homedir(), 'Documents', `${safe(subject)}.eml`), filters: [{ name: 'E-mail', extensions: ['eml'] }] };
  const r = w ? await dialog.showSaveDialog(w, opts) : await dialog.showSaveDialog(opts);
  if (r.canceled || !r.filePath) return '';
  const src: Buffer = await withImap(acc, async (c) => {
    const lock = await c.getMailboxLock(folder);
    try {
      const m = await c.fetchOne(String(uid), { source: true }, { uid: true });
      if (!m?.source) throw new UserError('Správa sa nenašla.');
      return m.source as Buffer;
    } finally {
      lock.release();
    }
  });
  fs.writeFileSync(r.filePath, src);
  return r.filePath;
}

/** Celý priečinok do jedného súboru .mbox (formát mboxrd, otvorí ho aj Thunderbird). */
export async function exportMbox(acc: AccountConfig, folder: string) {
  const w = win();
  const opts = { title: 'Exportovať priečinok', defaultPath: path.join(os.homedir(), 'Documents', `${safe(folder)}.mbox`), filters: [{ name: 'mbox', extensions: ['mbox'] }] };
  const r = w ? await dialog.showSaveDialog(w, opts) : await dialog.showSaveDialog(opts);
  if (r.canceled || !r.filePath) return { file: '', count: 0 };
  const out = fs.createWriteStream(r.filePath);
  let count = 0;
  await withImap(acc, async (c) => {
    const lock = await c.getMailboxLock(folder);
    try {
      if (!c.mailbox?.exists) return;
      for await (const m of c.fetch('1:*', { uid: true, source: true, internalDate: true, envelope: true })) {
        const from = (m.envelope?.from?.[0]?.address as string) || 'MAILER-DAEMON';
        const date = (m.internalDate ? new Date(m.internalDate) : new Date()).toUTCString().replace(',', '');
        const body = (m.source as Buffer).toString('binary').replace(/\r\n/g, '\n').replace(/^(>*From )/gm, '>$1');
        out.write(`From ${from} ${date}\n`, 'binary');
        out.write(body.endsWith('\n') ? body : `${body}\n`, 'binary');
        out.write('\n', 'binary');
        if (++count % 50 === 0) progress(`Export: ${count} správ…`);
      }
    } finally {
      lock.release();
    }
  });
  await new Promise<void>((res) => out.end(res));
  return { file: r.filePath, count };
}

/** Prečíta správy z mbox súboru (po jednej, aj pri veľkých súboroch). */
async function* readMbox(file: string): AsyncGenerator<Buffer> {
  const rl = readline.createInterface({ input: fs.createReadStream(file, { encoding: 'latin1' }), crlfDelay: Infinity });
  let cur: string[] | null = null;
  let prevBlank = true;
  for await (const line of rl) {
    if (line.startsWith('From ') && prevBlank) {
      if (cur && cur.length) yield Buffer.from(cur.join('\r\n').replace(/\r\n$/, ''), 'latin1');
      cur = [];
    } else if (cur) cur.push(line.replace(/^>(>*From )/, '$1'));
    prevBlank = line === '';
  }
  if (cur && cur.length) yield Buffer.from(cur.join('\r\n'), 'latin1');
}

async function ensure(c: { list(): Promise<Array<{ path: string; delimiter?: string }>>; mailboxCreate(p: string): Promise<unknown> }, name: string) {
  const list = await c.list();
  const delim = list[0]?.delimiter ?? '/';
  const target = name.split('/').join(delim);
  if (!list.some((m) => m.path === target)) await c.mailboxCreate(target).catch((): void => undefined);
  return target;
}

async function appendAll(acc: AccountConfig, folder: string, messages: AsyncIterable<Buffer> | Iterable<Buffer>) {
  let count = 0;
  await withImap(acc, async (c) => {
    const target = await ensure(c, folder);
    for await (const raw of messages) {
      if (raw.length < 20) continue;
      await c.append(target, raw, ['\\Seen']);
      if (++count % 25 === 0) progress(`Import do „${folder}“: ${count} správ…`);
    }
  });
  return count;
}

/** Import súborov .eml a .mbox do priečinka (predvolene „Import“). */
export async function importFiles(acc: AccountConfig, folder: string) {
  const w = win();
  const opts = { title: 'Importovať poštu', properties: ['openFile', 'multiSelections'] as Array<'openFile' | 'multiSelections'>, filters: [{ name: 'E-maily', extensions: ['eml', 'mbox', 'mbx'] }, { name: 'Všetky súbory', extensions: ['*'] }] };
  const r = w ? await dialog.showOpenDialog(w, opts) : await dialog.showOpenDialog(opts);
  if (r.canceled || !r.filePaths.length) return 0;
  let total = 0;
  for (const f of r.filePaths) {
    if (/\.eml$/i.test(f)) total += await appendAll(acc, folder, [fs.readFileSync(f)]);
    else total += await appendAll(acc, folder, readMbox(f));
  }
  log(`import: ${total} správ do ${folder}`);
  return total;
}

/** Mbox súbory v profile Thunderbirdu (súbor bez prípony, vedľa ktorého je .msf). */
function findThunderbirdMboxes(root: string, base = root, out: Array<{ file: string; name: string }> = []) {
  let entries: fs.Dirent[] = [];
  try {
    entries = fs.readdirSync(root, { withFileTypes: true });
  } catch {
    return out;
  }
  for (const e of entries) {
    const p = path.join(root, e.name);
    if (e.isDirectory()) findThunderbirdMboxes(p, base, out);
    else if (!path.extname(e.name) && fs.existsSync(`${p}.msf`) && fs.statSync(p).size > 0) {
      const rel = path.relative(base, p).split(path.sep).map((x) => x.replace(/\.sbd$/i, '')).join('/');
      out.push({ file: p, name: rel });
    }
  }
  return out;
}

/** Import celého profilu Thunderbirdu (lokálne priečinky aj kópie IMAP) do „Thunderbird/…“. */
export async function importThunderbird(acc: AccountConfig) {
  const w = win();
  const defaultPath = process.env.APPDATA ? path.join(process.env.APPDATA, 'Thunderbird', 'Profiles') : os.homedir();
  const opts = { title: 'Vyberte profil Thunderbirdu (priečinok …default-release)', defaultPath, properties: ['openDirectory'] as Array<'openDirectory'> };
  const r = w ? await dialog.showOpenDialog(w, opts) : await dialog.showOpenDialog(opts);
  if (r.canceled || !r.filePaths[0]) return { folders: 0, messages: 0 };
  const root = r.filePaths[0];
  const roots = ['Mail', 'ImapMail'].map((x) => path.join(root, x)).filter((x) => fs.existsSync(x));
  const boxes = (roots.length ? roots : [root]).flatMap((x) => findThunderbirdMboxes(x));
  if (!boxes.length) throw new UserError('V zvolenom priečinku sa nenašli priečinky pošty Thunderbirdu.', 'Vyberte priečinok profilu, napr. …\\Thunderbird\\Profiles\\abcd.default-release.');
  let messages = 0;
  for (const b of boxes) messages += await appendAll(acc, `Thunderbird/${b.name}`, readMbox(b.file));
  log(`import Thunderbird: ${boxes.length} priečinkov, ${messages} správ`);
  return { folders: boxes.length, messages };
}
