// Naposledy použité adresy (ručne zadaní aj vybraní príjemcovia odoslaných správ) – na dopĺňanie pri písaní.
// Uložené v userData/recent.json, aby ich mali všetky okná (aj samostatné okno písania) a prežili reinštaláciu rozhrania.
import fs from 'node:fs';
import path from 'node:path';
import { app, BrowserWindow } from 'electron';

export interface RecentAddr {
  name: string;
  email: string;
  /** koľkokrát sa na adresu písalo */
  count: number;
  /** naposledy (ISO) */
  last: string;
}

const MAX = 1000;
const file = () => path.join(app.getPath('userData'), 'recent.json');
let cache: RecentAddr[] | null = null;

function load(): RecentAddr[] {
  if (cache) return cache;
  try {
    const j = JSON.parse(fs.readFileSync(file(), 'utf8')) as RecentAddr[];
    cache = Array.isArray(j) ? j.filter((x) => x && typeof x.email === 'string') : [];
  } catch {
    cache = [];
  }
  return cache;
}
function persist() {
  const tmp = `${file()}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(cache ?? []), 'utf8');
  fs.renameSync(tmp, file());
  for (const w of BrowserWindow.getAllWindows()) if (!w.isDestroyed()) w.webContents.send('app:event', { type: 'contacts-changed' });
}

/** najčastejšie a najnovšie navrchu */
export function list(): RecentAddr[] {
  return [...load()].sort((a, b) => b.last.localeCompare(a.last) || b.count - a.count);
}

const EMAIL = /^[^\s@<>,;"]+@[^\s@<>,;"]+\.[^\s@<>,;"]+$/;

/** Zapamätá adresy; `migrate` = prevzatie starého zoznamu bez zvýšenia počtu. */
export function add(items: { name?: string; email: string }[], migrate = false) {
  const all = load();
  const now = new Date().toISOString();
  for (const it of items) {
    const email = (it.email ?? '').trim();
    if (!EMAIL.test(email)) continue;
    const name = (it.name ?? '').trim();
    const hit = all.find((x) => x.email.toLowerCase() === email.toLowerCase());
    if (hit) {
      if (!migrate) {
        hit.count++;
        hit.last = now;
      }
      if (name && name !== email && (!hit.name || hit.name === hit.email)) hit.name = name;
    } else all.push({ name: name && name !== email ? name : '', email, count: migrate ? 0 : 1, last: migrate ? '1970-01-01T00:00:00.000Z' : now });
  }
  cache = list().slice(0, MAX);
  persist();
  return true as const;
}

export function remove(emails: string[]) {
  const low = emails.map((e) => e.toLowerCase());
  cache = load().filter((x) => !low.includes(x.email.toLowerCase()));
  persist();
  return true as const;
}
