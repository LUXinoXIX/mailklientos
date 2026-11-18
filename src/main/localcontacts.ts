// Lokálny adresár (kontakty uložené len v programe, bez servera): userData/contacts.json.
// Súbor sa zapisuje atomicky (dočasný súbor + premenovanie), aby sa pri páde nestratil.
import fs from 'node:fs';
import path from 'node:path';
import { app, BrowserWindow, dialog } from 'electron';
import type { Contact } from '../core/ical';
import { newUid } from '../core/contactsio';
import { UserError } from './mail';

const file = () => path.join(app.getPath('userData'), 'contacts.json');
let cache: Contact[] | null = null;

function load(): Contact[] {
  if (cache) return cache;
  try {
    const j = JSON.parse(fs.readFileSync(file(), 'utf8')) as { contacts?: Contact[] };
    cache = Array.isArray(j.contacts) ? j.contacts : [];
  } catch {
    cache = [];
  }
  return cache;
}

function persist() {
  const tmp = `${file()}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify({ version: 1, contacts: cache ?? [] }, null, 1), 'utf8');
  fs.renameSync(tmp, file());
}

export function list(): Contact[] {
  return [...load()].sort((a, b) => (a.name || a.emails[0] || '').localeCompare(b.name || b.emails[0] || '', 'sk'));
}

/** Pridá alebo prepíše kontakty (podľa uid). */
export function put(items: Contact[]): Contact[] {
  const all = load();
  const out: Contact[] = [];
  for (const raw of items) {
    const c: Contact = { ...raw, uid: raw.uid || newUid(), href: undefined, etag: undefined, bookId: 'local', emails: raw.emails.filter(Boolean), phones: raw.phones.filter(Boolean) };
    const i = all.findIndex((x) => x.uid === c.uid);
    if (i >= 0) all[i] = c;
    else all.push(c);
    out.push(c);
  }
  persist();
  return out;
}

export function remove(uids: string[]) {
  cache = load().filter((c) => !uids.includes(c.uid));
  persist();
  return true as const;
}

/** Výber súboru na import; text sa dekóduje (UTF-8, UTF-16 s BOM, inak Windows-1250 ako slovenský Outlook). */
export async function pickImport(win: BrowserWindow | null) {
  const opts = {
    title: 'Import kontaktov',
    properties: ['openFile'] as Array<'openFile'>,
    filters: [{ name: 'Kontakty (vCard, CSV)', extensions: ['vcf', 'vcard', 'csv', 'txt'] }],
  };
  const r = win ? await dialog.showOpenDialog(win, opts) : await dialog.showOpenDialog(opts);
  const f = r.canceled ? undefined : r.filePaths[0];
  if (!f) return null;
  const buf = fs.readFileSync(f);
  if (buf.length > 50 * 1024 * 1024) throw new UserError('Súbor je väčší ako 50 MB – nevyzerá ako súbor s kontaktmi.');
  return { filename: path.basename(f), text: decode(buf) };
}

export function decode(buf: Buffer): string {
  if (buf[0] === 0xff && buf[1] === 0xfe) return new TextDecoder('utf-16le').decode(buf.subarray(2));
  if (buf[0] === 0xfe && buf[1] === 0xff) return new TextDecoder('utf-16be').decode(buf.subarray(2));
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(buf).replace(/^﻿/, '');
  } catch {
    return new TextDecoder('windows-1250').decode(buf);
  }
}

export async function saveExport(win: BrowserWindow | null, name: string, text: string) {
  const ext = path.extname(name).slice(1) || 'vcf';
  const opts = { title: 'Export kontaktov', defaultPath: path.join(app.getPath('documents'), name), filters: [{ name: ext === 'csv' ? 'CSV' : 'vCard', extensions: [ext] }] };
  const r = win ? await dialog.showSaveDialog(win, opts) : await dialog.showSaveDialog(opts);
  if (r.canceled || !r.filePath) return '';
  fs.writeFileSync(r.filePath, text, 'utf8');
  return r.filePath;
}
