// Lokálne kalendáre a odbery (ICS/webcal) pre kontá – aj pre IMAP kontá bez CalDAV servera.
// userData/calendars.json: kalendáre (lokálne = vlastné udalosti, odber = len na čítanie z adresy)
// a ich udalosti (hlavné záznamy s opakovaním; výskyty sa rozvinú pri čítaní).
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { app, BrowserWindow, dialog, net } from 'electron';
import type { DavCollection } from '../core/api';
import { expand, parseEvents, type CalEvent } from '../core/ical';
import { UserError } from './mail';

interface LocalCal {
  id: string;
  accountId: string;
  name: string;
  color?: string;
  kind: 'local' | 'sub';
  /** adresa odberu (https:// alebo webcal://) */
  url?: string;
  refreshedAt?: string;
  error?: string;
}
type Stored = CalEvent & { exdates?: number[] };
interface Db {
  calendars: LocalCal[];
  events: Record<string, Stored[]>;
}

const file = () => path.join(app.getPath('userData'), 'calendars.json');
let db: Db | null = null;
function load(): Db {
  if (db) return db;
  try {
    const j = JSON.parse(fs.readFileSync(file(), 'utf8')) as Db;
    db = { calendars: Array.isArray(j.calendars) ? j.calendars : [], events: j.events ?? {} };
  } catch {
    db = { calendars: [], events: {} };
  }
  return db;
}
function persist() {
  const tmp = `${file()}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(db), 'utf8');
  fs.renameSync(tmp, file());
}

export const isLocalId = (id?: string) => !!id && (id.startsWith('local:') || id.startsWith('sub:'));

/** Kalendáre konta ako DavCollection (rovnaký tvar ako kalendáre zo servera). */
export function collections(accountId: string): DavCollection[] {
  return load()
    .calendars.filter((c) => c.accountId === accountId)
    .map((c) => ({ id: c.id, href: c.id, name: c.kind === 'sub' ? `${c.name} (odber)` : c.name, color: c.color, kind: 'calendar' as const, events: true, readOnly: c.kind === 'sub' }));
}

/** Konto bez kalendára na serveri dostane jeden lokálny kalendár automaticky. */
export function ensureDefault(accountId: string) {
  const d = load();
  if (d.calendars.some((c) => c.accountId === accountId && c.kind === 'local')) return;
  d.calendars.push({ id: `local:${crypto.randomUUID()}`, accountId, name: 'Môj kalendár', color: '#2F6F5E', kind: 'local' });
  persist();
}

export function addLocal(accountId: string, name: string, color?: string): DavCollection {
  const d = load();
  const c: LocalCal = { id: `local:${crypto.randomUUID()}`, accountId, name: name.trim() || 'Kalendár', color: color || '#1d4ed8', kind: 'local' };
  d.calendars.push(c);
  persist();
  return collections(accountId).find((x) => x.id === c.id)!;
}

export async function subscribe(accountId: string, url: string, name?: string, color?: string): Promise<DavCollection> {
  const u = url.trim().replace(/^webcals?:\/\//i, 'https://');
  if (!/^https?:\/\/\S+$/i.test(u)) throw new UserError('Zadaj adresu kalendára (https://… alebo webcal://…).');
  const d = load();
  const c: LocalCal = { id: `sub:${crypto.randomUUID()}`, accountId, name: name?.trim() || new URL(u).hostname, color: color || '#b45309', kind: 'sub', url: u };
  d.calendars.push(c);
  await refresh(c);
  persist();
  return collections(accountId).find((x) => x.id === c.id)!;
}

export function rename(id: string, name: string, color?: string) {
  const c = load().calendars.find((x) => x.id === id);
  if (!c) throw new UserError('Kalendár sa nenašiel.');
  if (name.trim()) c.name = name.trim().replace(/ \(odber\)$/, '');
  if (color) c.color = color;
  persist();
  return true as const;
}

export function removeCal(id: string) {
  const d = load();
  d.calendars = d.calendars.filter((c) => c.id !== id);
  delete d.events[id];
  persist();
  return true as const;
}

async function refresh(c: LocalCal) {
  try {
    const r = await net.fetch(c.url!, { headers: { Accept: 'text/calendar, */*' } });
    if (!r.ok) throw new Error(`server vrátil ${r.status}`);
    const text = await r.text();
    if (!/BEGIN:VCALENDAR/i.test(text)) throw new Error('na adrese nie je kalendár (iCalendar)');
    load().events[c.id] = parseEvents(text);
    c.refreshedAt = new Date().toISOString();
    c.error = undefined;
  } catch (e) {
    c.error = (e as Error).message;
    if (!load().events[c.id]) throw new UserError(`Odber kalendára sa nepodaril: ${c.error}`);
  }
}

/** Výskyty udalostí v období; odbery sa obnovia, ak sú staršie ako hodina. */
export async function events(accountId: string, ids: string[], from: string, to: string): Promise<CalEvent[]> {
  const d = load();
  const cals = d.calendars.filter((c) => c.accountId === accountId && (!ids.length || ids.includes(c.id)));
  const stale = cals.filter((c) => c.kind === 'sub' && (!c.refreshedAt || Date.now() - Date.parse(c.refreshedAt) > 3600_000));
  if (stale.length) {
    await Promise.allSettled(stale.map(refresh));
    persist();
  }
  const f = Date.parse(from);
  const t = Date.parse(to);
  const out: CalEvent[] = [];
  for (const c of cals) {
    const list = d.events[c.id] ?? [];
    const overrides = list.filter((e) => e.recurrenceId);
    for (const e of list.filter((x) => !x.recurrenceId)) {
      const exdates = [...(e.exdates ?? []), ...overrides.filter((o) => o.uid === e.uid).map((o) => Date.parse(o.recurrenceId!))];
      for (const occ of expand({ ...e, exdates }, f, t))
        out.push({ ...occ, href: c.kind === 'local' ? `${c.id}#${e.uid}` : undefined, etag: 'local', calendarId: c.id, color: c.color });
    }
    for (const o of overrides) if (Date.parse(o.start) < t && Date.parse(o.end) > f) out.push({ ...o, calendarId: c.id, color: c.color });
  }
  return out;
}

const calOfHref = (href: string) => href.split('#')[0]!;

export function save(calendarId: string, e: CalEvent): CalEvent {
  const d = load();
  const c = d.calendars.find((x) => x.id === calendarId);
  if (!c) throw new UserError('Kalendár sa nenašiel.');
  if (c.kind === 'sub') throw new UserError('Kalendár z odberu je len na čítanie.');
  // presun do iného lokálneho kalendára: odstrániť z pôvodného
  if (e.href && calOfHref(e.href) !== calendarId) removeEvent(e.href);
  const list = (d.events[calendarId] ??= []);
  const uid = e.uid || `${crypto.randomUUID()}@mailklient`;
  const prev = list.find((x) => x.uid === uid && !x.recurrenceId);
  const ev: Stored = { ...e, uid, href: undefined, etag: undefined, calendarId: undefined, color: undefined, exdates: prev?.exdates, recurrenceId: undefined };
  // pri opakovanej udalosti ostáva pôvodný začiatok série (mení sa celá séria), ak sa čas výskytu nezmenil
  d.events[calendarId] = [...list.filter((x) => x.uid !== uid || x.recurrenceId), ev];
  persist();
  return { ...ev, href: `${calendarId}#${uid}`, calendarId, color: c.color };
}

export function removeEvent(href: string) {
  const [calId, uid] = href.split('#');
  const d = load();
  if (calId && uid && d.events[calId]) {
    d.events[calId] = d.events[calId]!.filter((x) => x.uid !== uid);
    persist();
  }
  return true as const;
}

/** Import .ics do lokálneho kalendára (duplicity podľa UID sa prepíšu). */
export function importToLocal(calendarId: string, list: Stored[]) {
  const d = load();
  const c = d.calendars.find((x) => x.id === calendarId);
  if (!c || c.kind !== 'local') throw new UserError('Importovať sa dá do vlastného (nie odoberaného) kalendára.');
  const cur = d.events[calendarId] ?? [];
  const keys = new Set(list.map((e) => `${e.uid}|${e.recurrenceId ?? ''}`));
  d.events[calendarId] = [...cur.filter((e) => !keys.has(`${e.uid}|${e.recurrenceId ?? ''}`)), ...list.map((e) => ({ ...e, href: undefined, etag: undefined, calendarId: undefined }))];
  persist();
  return list.length;
}

export async function pickIcs(win: BrowserWindow | null): Promise<string | null> {
  const opts = { title: 'Import kalendára', properties: ['openFile'] as Array<'openFile'>, filters: [{ name: 'Kalendár (iCalendar)', extensions: ['ics', 'ical', 'ifb', 'icalendar'] }] };
  const r = win ? await dialog.showOpenDialog(win, opts) : await dialog.showOpenDialog(opts);
  const f = r.canceled ? undefined : r.filePaths[0];
  if (!f) return null;
  const text = fs.readFileSync(f, 'utf8');
  if (!/BEGIN:VCALENDAR/i.test(text)) throw new UserError('Súbor nie je kalendár (iCalendar .ics).');
  return text;
}

export function forgetAccount(accountId: string) {
  const d = load();
  const ids = d.calendars.filter((c) => c.accountId === accountId).map((c) => c.id);
  d.calendars = d.calendars.filter((c) => c.accountId !== accountId);
  for (const id of ids) delete d.events[id];
  persist();
}

export function subInfo(id: string) {
  const c = load().calendars.find((x) => x.id === id);
  return c ? { url: c.url, refreshedAt: c.refreshedAt, error: c.error, kind: c.kind } : null;
}

/** Najbližšie udalosti z kalendára na adrese (dlaždica v Súhrne); výsledok sa drží 1 hodinu. */
const feedCache = new Map<string, { at: number; list: Stored[] }>();
export async function upcoming(url: string, count: number): Promise<CalEvent[]> {
  const u = url.trim().replace(/^webcals?:\/\//i, 'https://');
  if (!/^https?:\/\/\S+$/i.test(u)) throw new UserError('Neplatná adresa kalendára.');
  let c = feedCache.get(u);
  if (!c || Date.now() - c.at > 3600_000) {
    const r = await net.fetch(u, { headers: { Accept: 'text/calendar, */*' } });
    if (!r.ok) throw new UserError(`Kalendár na adrese vrátil chybu ${r.status}.`);
    const text = await r.text();
    if (!/BEGIN:VCALENDAR/i.test(text)) throw new UserError('Na adrese nie je kalendár (iCalendar).');
    c = { at: Date.now(), list: parseEvents(text) };
    feedCache.set(u, c);
  }
  const from = Date.now();
  const to = from + 120 * 86400_000;
  const out: CalEvent[] = [];
  for (const e of c.list.filter((x) => !x.recurrenceId)) for (const occ of expand(e, from, to, 50)) out.push(occ);
  return out.filter((e) => Date.parse(e.end) > from).sort((a, b) => a.start.localeCompare(b.start)).slice(0, Math.max(1, Math.min(30, count)));
}

/** Text stránky/kanála RSS pre dlaždicu Správy (spracuje ho rozhranie); 15 min v pamäti, max. 3 MB. */
const rssCache = new Map<string, { at: number; text: string; url: string }>();
export async function fetchFeedText(url: string): Promise<{ text: string; url: string }> {
  const u = url.trim().replace(/^feed:\/\//i, 'https://');
  if (!/^https?:\/\/\S+$/i.test(u)) throw new UserError('Neplatná adresa (https://…).');
  const c = rssCache.get(u);
  if (c && Date.now() - c.at < 15 * 60_000) return { text: c.text, url: c.url };
  const r = await net.fetch(u, { headers: { Accept: 'application/rss+xml, application/atom+xml, application/xml, text/xml, text/html;q=0.8, */*;q=0.5' } });
  if (!r.ok) throw new UserError(`Stránka vrátila chybu ${r.status}.`);
  const buf = Buffer.from(await r.arrayBuffer());
  if (buf.length > 3 * 1024 * 1024) throw new UserError('Kanál je príliš veľký.');
  // kódovanie z hlavičky XML (napr. windows-1250), inak UTF-8
  const head = buf.subarray(0, 300).toString('latin1');
  const enc = /encoding=["']([\w-]+)["']/i.exec(head)?.[1] ?? /charset=([\w-]+)/i.exec(r.headers.get('content-type') ?? '')?.[1] ?? 'utf-8';
  let text: string;
  try {
    text = new TextDecoder(enc.toLowerCase()).decode(buf);
  } catch {
    text = new TextDecoder('utf-8').decode(buf);
  }
  rssCache.set(u, { at: Date.now(), text, url: r.url || u });
  return { text, url: r.url || u };
}
