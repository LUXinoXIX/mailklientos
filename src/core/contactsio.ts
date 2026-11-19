// Import a export kontaktov: vCard (.vcf – Apple, Android, Google, Outlook) a CSV (Outlook, Google, všeobecné).
// Bez Electronu – používa sa v rozhraní (náhľad, mapovanie stĺpcov) a v testoch.
import { buildVCard, parseVCards, type Contact } from './ical';

export type CsvFormat = 'outlook' | 'google' | 'generic';
export type Field = 'skip' | 'fullName' | 'firstName' | 'middleName' | 'lastName' | 'email' | 'phone' | 'org' | 'title' | 'note';

export const FIELD_LABELS: Record<Field, string> = {
  skip: '— nepoužiť —',
  fullName: 'Celé meno',
  firstName: 'Meno',
  middleName: 'Stredné meno',
  lastName: 'Priezvisko',
  email: 'E-mail',
  phone: 'Telefón',
  org: 'Firma',
  title: 'Funkcia',
  note: 'Poznámka',
};

// ---------------------------------------------------------------------------
// vCard

/** vCard 2.1 z Outlooku: hodnoty v quoted-printable (aj zalomené „=“ na konci riadka) a CHARSET. */
function decodeQuotedPrintable(text: string): string {
  const lines = text.split(/\r?\n/);
  const out: string[] = [];
  for (let i = 0; i < lines.length; i++) {
    let line = lines[i]!;
    const head = line.slice(0, line.indexOf(':') + 1);
    if (!/ENCODING=QUOTED-PRINTABLE|;QUOTED-PRINTABLE/i.test(head)) {
      out.push(line);
      continue;
    }
    while (line.endsWith('=') && i + 1 < lines.length) line = line.slice(0, -1) + lines[++i]!;
    const charset = /CHARSET=([^;:]+)/i.exec(head)?.[1] ?? 'utf-8';
    const raw = line.slice(head.length);
    const bytes: number[] = [];
    for (let j = 0; j < raw.length; j++) {
      if (raw[j] === '=' && /^[0-9A-F]{2}$/i.test(raw.slice(j + 1, j + 3))) {
        bytes.push(parseInt(raw.slice(j + 1, j + 3), 16));
        j += 2;
      } else bytes.push(raw.charCodeAt(j) & 0xff);
    }
    let value: string;
    try {
      value = new TextDecoder(charset.toLowerCase()).decode(new Uint8Array(bytes));
    } catch {
      value = new TextDecoder('utf-8').decode(new Uint8Array(bytes));
    }
    const cleanHead = head.replace(/;ENCODING=QUOTED-PRINTABLE|;QUOTED-PRINTABLE/gi, '').replace(/;CHARSET=[^;:]+/gi, '');
    out.push(cleanHead + value.replace(/\r?\n/g, '\\n'));
  }
  return out.join('\r\n');
}

export function parseVcf(text: string): Contact[] {
  return parseVCards(decodeQuotedPrintable(text.replace(/^﻿/, ''))).filter((c) => c.name || c.emails.length || c.phones.length);
}

export function toVcf(list: Contact[]): string {
  return list.map((c) => buildVCard({ ...c, uid: c.uid || newUid() })).join('');
}

// ---------------------------------------------------------------------------
// CSV

/** RFC 4180: úvodzovky, zdvojené úvodzovky, zalomenie riadka v úvodzovkách; oddeľovač , ; alebo tabulátor. */
export function parseCsv(text: string): string[][] {
  text = text.replace(/^﻿/, '');
  const first = text.slice(0, text.search(/\r?\n/) >>> 0 || text.length);
  const count = (ch: string) => first.split(ch).length;
  const sep = count('\t') > count(';') && count('\t') > count(',') ? '\t' : count(';') > count(',') ? ';' : ',';
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let inQ = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]!;
    if (inQ) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          cell += '"';
          i++;
        } else inQ = false;
      } else cell += ch;
    } else if (ch === '"') inQ = true;
    else if (ch === sep) {
      row.push(cell);
      cell = '';
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = '';
    } else cell += ch;
  }
  if (cell || row.length) {
    row.push(cell);
    rows.push(row);
  }
  return rows.filter((r) => r.some((c) => c.trim()));
}

const norm = (s: string) => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();

export function detectCsvFormat(headers: string[]): CsvFormat {
  const h = headers.map(norm);
  if (h.some((x) => /^e mail \d+ value$/.test(x))) return 'google';
  if (h.includes('e mail address') || h.includes('e mailova adresa') || h.includes('e mail adresa')) return 'outlook';
  return 'generic';
}

/** Automatické priradenie stĺpcov (anglický aj slovenský Outlook, nový aj starý Google CSV, bežné názvy). */
export function guessField(header: string): Field {
  const h = norm(header);
  if (!h) return 'skip';
  if (/label$|type$|typ$/.test(h)) return 'skip';
  if (/^(first name|given name|meno|krstne meno|vorname|firstname)$/.test(h)) return 'firstName';
  if (/^(middle name|additional name|stredne meno|druhe meno)$/.test(h)) return 'middleName';
  if (/^(last name|family name|surname|priezvisko|nachname|lastname)$/.test(h)) return 'lastName';
  if (/^(name|full name|display name|zobrazovane meno|cele meno|meno a priezvisko|nazov)$/.test(h)) return 'fullName';
  if (/e ?mail|mail address|^email/.test(h) && !/display|zobraz|type|typ/.test(h)) return 'email';
  if (/phone|telef|mobil|mobile|fax|pager/.test(h) && !/type|typ|label/.test(h)) return /fax|pager/.test(h) ? 'skip' : 'phone';
  if (/^(company|organization( 1)? name|organization|firma|spolocnost|organizacia)$/.test(h)) return 'org';
  if (/^(job title|organization( 1)? title|title|funkcia|pracovna pozicia|pozicia)$/.test(h)) return 'title';
  if (/^(notes|note|poznamky|poznamka)$/.test(h)) return 'note';
  return 'skip';
}

const EMAIL = /[^\s<>,;:"]+@[^\s<>,;:"]+\.[^\s<>,;:"]+/g;

/** Riadky CSV → kontakty podľa mapovania stĺpcov. Google dáva viac hodnôt do bunky oddelených „ ::: “. */
export function csvToContacts(rows: string[][], mapping: Field[]): Contact[] {
  const out: Contact[] = [];
  for (const r of rows) {
    const parts: Partial<Record<Field, string[]>> = {};
    r.forEach((v, i) => {
      const f = mapping[i] ?? 'skip';
      const val = v.replace(/\r\n?/g, '\n').trim();
      if (f === 'skip' || !val) return;
      (parts[f] ??= []).push(val);
    });
    const one = (f: Field) => parts[f]?.join(' ').trim() || '';
    const name = one('fullName') || [one('firstName'), one('middleName'), one('lastName')].filter(Boolean).join(' ');
    const emails = [...new Set((parts.email ?? []).flatMap((v) => v.match(EMAIL) ?? []).map((e) => e.toLowerCase()))];
    const phones = [...new Set((parts.phone ?? []).flatMap((v) => v.split(/\s*:::\s*/)).map((p) => p.trim()).filter(Boolean))];
    if (!name && !emails.length && !phones.length) continue;
    out.push({ uid: '', name: name || emails[0] || phones[0] || '', emails, phones, org: one('org') || undefined, title: one('title') || undefined, note: parts.note?.join('\n') || undefined });
  }
  return out;
}

function csvCell(v: string): string {
  return /[",;\r\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v;
}
function splitName(name: string): [string, string, string] {
  const p = name.trim().split(/\s+/);
  if (p.length === 1) return [p[0] ?? '', '', ''];
  const last = p.pop()!;
  const first = p.shift()!;
  return [first, p.join(' '), last];
}

/** CSV pre Outlook (anglické hlavičky – Outlook ich pri importe rozpozná v každom jazyku), UTF-8 s BOM. */
export function toOutlookCsv(list: Contact[]): string {
  const head = ['First Name', 'Middle Name', 'Last Name', 'Company', 'Job Title', 'E-mail Address', 'E-mail 2 Address', 'E-mail 3 Address', 'Mobile Phone', 'Business Phone', 'Home Phone', 'Notes'];
  const rows = list.map((c) => {
    const [f, m, l] = splitName(c.name);
    return [f, m, l, c.org ?? '', c.title ?? '', c.emails[0] ?? '', c.emails[1] ?? '', c.emails[2] ?? '', c.phones[0] ?? '', c.phones[1] ?? '', c.phones[2] ?? '', c.note ?? ''];
  });
  return '﻿' + [head, ...rows].map((r) => r.map(csvCell).join(',')).join('\r\n') + '\r\n';
}

/** CSV pre Google Kontakty (formát z contacts.google.com, 2023+). */
export function toGoogleCsv(list: Contact[]): string {
  const ne = Math.max(1, ...list.map((c) => c.emails.length));
  const np = Math.max(1, ...list.map((c) => c.phones.length));
  const head = ['First Name', 'Middle Name', 'Last Name', 'Organization Name', 'Organization Title', 'Notes'];
  for (let i = 1; i <= ne; i++) head.push(`E-mail ${i} - Label`, `E-mail ${i} - Value`);
  for (let i = 1; i <= np; i++) head.push(`Phone ${i} - Label`, `Phone ${i} - Value`);
  const rows = list.map((c) => {
    const [f, m, l] = splitName(c.name);
    const r = [f, m, l, c.org ?? '', c.title ?? '', c.note ?? ''];
    for (let i = 0; i < ne; i++) r.push(c.emails[i] ? (i ? '* Other' : '* Work') : '', c.emails[i] ?? '');
    for (let i = 0; i < np; i++) r.push(c.phones[i] ? 'Mobile' : '', c.phones[i] ?? '');
    return r;
  });
  return '﻿' + [head, ...rows].map((r) => r.map(csvCell).join(',')).join('\r\n') + '\r\n';
}

// ---------------------------------------------------------------------------
// zlúčenie a duplicity

export type DupMode = 'merge' | 'skip' | 'add';

export function newUid(): string {
  return globalThis.crypto?.randomUUID?.() ?? `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
}

/** Kľúč na zisťovanie rovnakého človeka: e-maily (malými písmenami), inak meno + telefón. */
export function contactKeys(c: Contact): string[] {
  const k = c.emails.map((e) => `e:${e.toLowerCase()}`);
  if (!k.length && c.name) k.push(`n:${norm(c.name)}|${(c.phones[0] ?? '').replace(/\D/g, '')}`);
  return k;
}

/** Doplní chýbajúce údaje z `b` do `a` (nič existujúce neprepíše). */
export function mergeContact(a: Contact, b: Contact): Contact {
  const uniq = (x: string[], f: (s: string) => string) => x.filter((v, i) => x.findIndex((y) => f(y) === f(v)) === i);
  return {
    ...a,
    name: a.name || b.name,
    emails: uniq([...a.emails, ...b.emails], (s) => s.toLowerCase()),
    phones: uniq([...a.phones, ...b.phones], (s) => s.replace(/\D/g, '')),
    org: a.org || b.org,
    title: a.title || b.title,
    note: a.note && b.note && !a.note.includes(b.note) ? `${a.note}\n${b.note}` : a.note || b.note,
  };
}

export interface ImportPlan {
  /** kontakty na uloženie (nové aj zmenené existujúce) */
  save: Contact[];
  added: number;
  merged: number;
  skipped: number;
  noEmail: number;
}

/** Porovná importované kontakty s cieľovým adresárom a podľa režimu duplicít pripraví, čo sa uloží. */
export function planImport(incoming: Contact[], existing: Contact[], mode: DupMode): ImportPlan {
  const index = new Map<string, Contact>();
  for (const c of existing) for (const k of contactKeys(c)) index.set(k, c);
  const changed = new Map<Contact, Contact>();
  const added: Contact[] = [];
  let merged = 0;
  let skipped = 0;
  for (const raw of incoming) {
    const c = { ...raw, uid: '', href: undefined, etag: undefined, bookId: undefined };
    const hit = mode === 'add' ? undefined : contactKeys(c).map((k) => index.get(k)).find(Boolean);
    if (!hit) {
      const n = { ...c, uid: newUid() };
      added.push(n);
      for (const k of contactKeys(n)) index.set(k, n);
      continue;
    }
    if (mode === 'skip') {
      skipped++;
      continue;
    }
    if (added.includes(hit)) {
      // ten istý človek dvakrát v importovanom súbore → spojiť do jedného nového
      Object.assign(hit, mergeContact(hit, c));
      for (const k of contactKeys(hit)) index.set(k, hit);
    } else {
      const base = changed.get(hit) ?? hit;
      const m = mergeContact(base, c);
      if (JSON.stringify(m) !== JSON.stringify(base)) changed.set(hit, m);
      for (const k of contactKeys(m)) index.set(k, hit);
    }
    merged++;
  }
  return { save: [...added, ...changed.values()], added: added.length, merged, skipped, noEmail: incoming.filter((c) => !c.emails.length).length };
}
