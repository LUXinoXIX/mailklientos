// iCalendar (RFC 5545) a vCard (RFC 6350) bez externých knižníc: čítanie udalostí a kontaktov,
// rozvinutie opakovaní v zadanom rozsahu, zápis udalosti, odpoveď na pozvánku (iMIP REPLY).

export interface CalAttendee {
  email: string;
  name?: string;
  /** ROOM / RESOURCE pri zasadačkách a zariadeniach */
  cutype?: string;
  partstat?: 'NEEDS-ACTION' | 'ACCEPTED' | 'DECLINED' | 'TENTATIVE' | string;
  role?: string;
}

export interface CalEvent {
  uid: string;
  summary: string;
  /** ISO UTC; pri celodennej udalosti polnoc miestneho času */
  start: string;
  end: string;
  allDay: boolean;
  location?: string;
  description?: string;
  /** odkaz na online schôdzu (Teams, Jitsi, …) */
  onlineUrl?: string;
  organizer?: { email: string; name?: string };
  attendees?: CalAttendee[];
  status?: string;
  rrule?: string;
  /** pri rozvinutom opakovaní: pôvodný začiatok výskytu */
  recurrenceId?: string;
  sequence?: number;
  /** odkaz na zdroj na serveri */
  href?: string;
  etag?: string;
  calendarId?: string;
  color?: string;
}

export interface Contact {
  uid: string;
  name: string;
  emails: string[];
  phones: string[];
  org?: string;
  title?: string;
  note?: string;
  href?: string;
  etag?: string;
  bookId?: string;
}

// ---------------------------------------------------------------------------
// riadky a vlastnosti

interface Prop {
  name: string;
  params: Record<string, string>;
  value: string;
}

/** Rozbalí zalomené riadky (RFC 5545 3.1) a rozdelí na vlastnosti. */
function parseLines(text: string): Prop[] {
  const raw = text.replace(/\r\n[ \t]/g, '').replace(/\n[ \t]/g, '').split(/\r?\n/);
  const out: Prop[] = [];
  for (const line of raw) {
    if (!line.trim()) continue;
    // názov;param=hodnota;param="hodnota:s dvojbodkou":hodnota
    let i = 0;
    let inQ = false;
    for (; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') inQ = !inQ;
      else if (ch === ':' && !inQ) break;
    }
    const head = line.slice(0, i);
    const value = line.slice(i + 1);
    const parts = head.match(/(?:[^;"]|"[^"]*")+/g) ?? [head];
    // skupina pred názvom (item1.EMAIL z Apple) sa ignoruje
    const name = (parts[0] ?? '').toUpperCase().replace(/^[A-Z0-9-]+\./, '');
    const params: Record<string, string> = {};
    for (const p of parts.slice(1)) {
      const eq = p.indexOf('=');
      if (eq > 0) params[p.slice(0, eq).toUpperCase()] = p.slice(eq + 1).replace(/^"|"$/g, '');
    }
    out.push({ name, params, value });
  }
  return out;
}

const unescapeText = (v: string) => v.replace(/\\n/gi, '\n').replace(/\\([,;\\])/g, '$1');
const escapeText = (v: string) => v.replace(/\\/g, '\\\\').replace(/\n/g, '\\n').replace(/([,;])/g, '\\$1');

interface Comp {
  type: string;
  props: Prop[];
  children: Comp[];
}

function toTree(props: Prop[]): Comp[] {
  const root: Comp = { type: 'ROOT', props: [], children: [] };
  const stack: Comp[] = [root];
  for (const p of props) {
    const top = stack[stack.length - 1]!;
    if (p.name === 'BEGIN') {
      const c: Comp = { type: p.value.toUpperCase(), props: [], children: [] };
      top.children.push(c);
      stack.push(c);
    } else if (p.name === 'END') {
      if (stack.length > 1) stack.pop();
    } else top.props.push(p);
  }
  return root.children;
}

const all = (c: Comp, type: string): Comp[] => c.children.flatMap((x) => (x.type === type ? [x] : all(x, type)));
const get = (c: Comp, name: string) => c.props.find((p) => p.name === name);

// ---------------------------------------------------------------------------
// časové pásma: TZID → posun pomocou Intl (aj názvy pásiem z Windows/Outlooku)

const WINDOWS_TZ: Record<string, string> = {
  'Central Europe Standard Time': 'Europe/Bratislava',
  'Central European Standard Time': 'Europe/Warsaw',
  'W. Europe Standard Time': 'Europe/Berlin',
  'Romance Standard Time': 'Europe/Paris',
  'GMT Standard Time': 'Europe/London',
  'E. Europe Standard Time': 'Europe/Bucharest',
  'FLE Standard Time': 'Europe/Kiev',
  'UTC': 'UTC',
  'Eastern Standard Time': 'America/New_York',
  'Pacific Standard Time': 'America/Los_Angeles',
};

function zoneOf(tzid?: string): string | undefined {
  if (!tzid) return undefined;
  const t = tzid.replace(/^\/+/, '').replace(/^.*?(?=[A-Z][a-z]+\/[A-Z])/, '');
  if (WINDOWS_TZ[tzid]) return WINDOWS_TZ[tzid];
  try {
    new Intl.DateTimeFormat('en-US', { timeZone: t });
    return t;
  } catch {
    return undefined;
  }
}

/** Posun pásma v minútach pre daný okamih (UTC → miestny čas v pásme). */
function tzOffset(zone: string, utcMs: number): number {
  const f = new Intl.DateTimeFormat('en-US', {
    timeZone: zone, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit',
  });
  const p = Object.fromEntries(f.formatToParts(new Date(utcMs)).map((x) => [x.type, x.value]));
  const asUtc = Date.UTC(+p.year!, +p.month! - 1, +p.day!, +p.hour!, +p.minute!, +p.second!);
  return Math.round((asUtc - utcMs) / 60000);
}

/** Miestny čas v pásme → UTC milisekundy. */
function zonedToUtc(y: number, mo: number, d: number, h: number, mi: number, s: number, zone: string): number {
  const guess = Date.UTC(y, mo, d, h, mi, s);
  let off = tzOffset(zone, guess);
  let utc = guess - off * 60000;
  const off2 = tzOffset(zone, utc);
  if (off2 !== off) utc = guess - off2 * 60000;
  return utc;
}

/** Hodnota DATE alebo DATE-TIME → { ms (UTC), allDay } */
export function parseIcalDate(value: string, tzid?: string): { ms: number; allDay: boolean } {
  const m = /^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})?(Z)?)?$/.exec(value.trim());
  if (!m) return { ms: Date.parse(value), allDay: false };
  const [, y, mo, d, h, mi, s, z] = m;
  if (h === undefined) return { ms: new Date(+y!, +mo! - 1, +d!).getTime(), allDay: true };
  if (z) return { ms: Date.UTC(+y!, +mo! - 1, +d!, +h, +mi!, +(s ?? 0)), allDay: false };
  const zone = zoneOf(tzid);
  if (zone) return { ms: zonedToUtc(+y!, +mo! - 1, +d!, +h, +mi!, +(s ?? 0), zone), allDay: false };
  return { ms: new Date(+y!, +mo! - 1, +d!, +h, +mi!, +(s ?? 0)).getTime(), allDay: false };
}

function parseDuration(v: string): number {
  const m = /^([+-])?P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?$/.exec(v.trim());
  if (!m) return 0;
  const [, sign, w, d, h, mi, s] = m;
  const ms = ((+(w ?? 0) * 7 + +(d ?? 0)) * 86400 + +(h ?? 0) * 3600 + +(mi ?? 0) * 60 + +(s ?? 0)) * 1000;
  return sign === '-' ? -ms : ms;
}

const mailto = (v: string) => v.replace(/^mailto:/i, '').trim();

function eventFrom(c: Comp): CalEvent | null {
  const uid = get(c, 'UID')?.value ?? '';
  const ds = get(c, 'DTSTART');
  if (!ds) return null;
  const start = parseIcalDate(ds.value, ds.params.TZID);
  const de = get(c, 'DTEND');
  const dur = get(c, 'DURATION');
  let endMs = de ? parseIcalDate(de.value, de.params.TZID).ms : dur ? start.ms + parseDuration(dur.value) : start.ms + (start.allDay ? 86400000 : 3600000);
  if (endMs < start.ms) endMs = start.ms;
  const org = get(c, 'ORGANIZER');
  const rid = get(c, 'RECURRENCE-ID');
  return {
    uid,
    summary: unescapeText(get(c, 'SUMMARY')?.value ?? '(bez názvu)'),
    start: new Date(start.ms).toISOString(),
    end: new Date(endMs).toISOString(),
    allDay: start.allDay,
    location: get(c, 'LOCATION') ? unescapeText(get(c, 'LOCATION')!.value) : undefined,
    description: get(c, 'DESCRIPTION') ? unescapeText(get(c, 'DESCRIPTION')!.value) : undefined,
    onlineUrl: (() => {
      const v = (get(c, 'X-MICROSOFT-SKYPETEAMSMEETINGURL') ?? get(c, 'X-GOOGLE-CONFERENCE') ?? get(c, 'URL'))?.value;
      return v && /^https:\/\//i.test(v) ? v : undefined;
    })(),
    organizer: org ? { email: mailto(org.value), name: org.params.CN } : undefined,
    attendees: c.props
      .filter((p) => p.name === 'ATTENDEE')
      .map((p) => ({ email: mailto(p.value), name: p.params.CN, partstat: p.params.PARTSTAT, role: p.params.ROLE, cutype: p.params.CUTYPE })),
    status: get(c, 'STATUS')?.value,
    rrule: get(c, 'RRULE')?.value,
    recurrenceId: rid ? new Date(parseIcalDate(rid.value, rid.params.TZID).ms).toISOString() : undefined,
    sequence: Number(get(c, 'SEQUENCE')?.value ?? 0),
  };
}

/** Udalosti z textu .ics (bez rozvinutia opakovaní). */
export function parseEvents(ics: string): Array<CalEvent & { exdates: number[] }> {
  const tree = toTree(parseLines(ics));
  return tree.flatMap((cal) => all(cal, 'VEVENT')).flatMap((c) => {
    // udalosť s neplatným dátumom sa preskočí – nesmie zhodiť celý kalendár („Invalid time value“)
    let e: ReturnType<typeof eventFrom>;
    try {
      e = eventFrom(c);
    } catch {
      return [];
    }
    if (!e) return [];
    const exdates = c.props
      .filter((p) => p.name === 'EXDATE')
      .flatMap((p) => p.value.split(',').map((v) => parseIcalDate(v, p.params.TZID).ms));
    return [{ ...e, exdates }];
  });
}

export function parseMethod(ics: string): string | undefined {
  return toTree(parseLines(ics))
    .map((c) => get(c, 'METHOD')?.value)
    .find(Boolean)
    ?.toUpperCase();
}

// ---------------------------------------------------------------------------
// opakovania (RRULE): DAILY, WEEKLY (BYDAY), MONTHLY (BYMONTHDAY / BYDAY s poradím), YEARLY; INTERVAL, COUNT, UNTIL

const DAYS = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];

export function expand(ev: CalEvent & { exdates?: number[] }, fromMs: number, toMs: number, max = 500): CalEvent[] {
  const startMs = Date.parse(ev.start);
  const dur = Date.parse(ev.end) - startMs;
  if (!ev.rrule) return startMs < toMs && startMs + Math.max(dur, 1) > fromMs ? [ev] : [];
  const r = Object.fromEntries(ev.rrule.split(';').map((x) => x.split('=') as [string, string]));
  const freq = r.FREQ;
  const interval = Math.max(1, Number(r.INTERVAL ?? 1));
  const count = r.COUNT ? Number(r.COUNT) : Infinity;
  const until = r.UNTIL ? parseIcalDate(r.UNTIL).ms : Infinity;
  const byday = r.BYDAY ? r.BYDAY.split(',') : [];
  const bymonthday = r.BYMONTHDAY ? r.BYMONTHDAY.split(',').map(Number) : [];
  const ex = new Set(ev.exdates ?? []);
  const out: CalEvent[] = [];
  const s0 = new Date(startMs);
  // výskyty sa počítajú v miestnom čase (rovnaká hodina aj po zmene letného času)
  const mk = (d: Date) => new Date(d.getFullYear(), d.getMonth(), d.getDate(), s0.getHours(), s0.getMinutes(), s0.getSeconds());
  let n = 0;
  const push = (d: Date) => {
    const t = d.getTime();
    if (t < startMs || t > until || n >= count) return false;
    n++;
    if (!ex.has(t) && t < toMs && t + Math.max(dur, 1) > fromMs) {
      out.push({ ...ev, start: new Date(t).toISOString(), end: new Date(t + dur).toISOString(), recurrenceId: new Date(t).toISOString() });
    }
    return true;
  };
  const nthWeekday = (y: number, m: number, spec: string): Date | null => {
    const mm = /^([+-]?\d+)?(SU|MO|TU|WE|TH|FR|SA)$/.exec(spec);
    if (!mm) return null;
    const ord = mm[1] ? Number(mm[1]) : 1;
    const wd = DAYS.indexOf(mm[2]!);
    if (ord > 0) {
      const first = new Date(y, m, 1);
      const day = 1 + ((wd - first.getDay() + 7) % 7) + (ord - 1) * 7;
      const d = new Date(y, m, day);
      return d.getMonth() === m ? d : null;
    }
    const last = new Date(y, m + 1, 0);
    const day = last.getDate() - ((last.getDay() - wd + 7) % 7) + (ord + 1) * 7;
    const d = new Date(y, m, day);
    return d.getMonth() === m ? d : null;
  };
  for (let i = 0, guard = 0; guard < 5000 && out.length < max; i += interval, guard++) {
    let batch: Date[] = [];
    if (freq === 'DAILY') batch = [mk(new Date(s0.getFullYear(), s0.getMonth(), s0.getDate() + i))];
    else if (freq === 'WEEKLY') {
      const wkStart = new Date(s0.getFullYear(), s0.getMonth(), s0.getDate() - ((s0.getDay() + 6) % 7) + i * 7);
      const days = byday.length ? byday.map((d) => DAYS.indexOf(d.replace(/^[+-]?\d+/, ''))) : [s0.getDay()];
      batch = days.map((wd) => mk(new Date(wkStart.getFullYear(), wkStart.getMonth(), wkStart.getDate() + ((wd + 6) % 7)))).sort((a, b) => a.getTime() - b.getTime());
    } else if (freq === 'MONTHLY') {
      const base = new Date(s0.getFullYear(), s0.getMonth() + i, 1);
      const y = base.getFullYear();
      const m = base.getMonth();
      if (byday.length) batch = byday.map((b) => nthWeekday(y, m, b)).filter((d): d is Date => !!d).map(mk);
      else batch = (bymonthday.length ? bymonthday : [s0.getDate()])
        .map((md) => (md > 0 ? new Date(y, m, md) : new Date(y, m + 1, md + 1)))
        .filter((d) => d.getMonth() === m)
        .map(mk);
      batch.sort((a, b) => a.getTime() - b.getTime());
    } else if (freq === 'YEARLY') {
      const d = new Date(s0.getFullYear() + i, s0.getMonth(), s0.getDate());
      if (d.getMonth() === s0.getMonth()) batch = [mk(d)];
    } else break;
    if (batch.length && batch[0]!.getTime() > Math.min(until, toMs)) break;
    for (const d of batch) push(d);
    if (n >= count) break;
  }
  return out;
}

// ---------------------------------------------------------------------------
// zápis

const pad = (n: number) => String(n).padStart(2, '0');
const utcStamp = (ms: number) => {
  const d = new Date(ms);
  return `${d.getUTCFullYear()}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}T${pad(d.getUTCHours())}${pad(d.getUTCMinutes())}${pad(d.getUTCSeconds())}Z`;
};
const dateStamp = (ms: number) => {
  const d = new Date(ms);
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;
};

const enc = new TextEncoder();
/** Zalomenie riadkov na 75 bajtov (RFC 5545). */
function fold(line: string): string {
  const out: string[] = [];
  let cur = '';
  for (const ch of line) {
    if (enc.encode(cur + ch).length > 74) {
      out.push(cur);
      cur = ' ';
    }
    cur += ch;
  }
  out.push(cur);
  return out.join('\r\n');
}

function eventLines(e: CalEvent, extra: string[] = []): string[] {
  const s = Date.parse(e.start);
  const en = Date.parse(e.end);
  const L = [
    'BEGIN:VEVENT',
    `UID:${e.uid}`,
    `DTSTAMP:${utcStamp(Date.now())}`,
    e.allDay ? `DTSTART;VALUE=DATE:${dateStamp(s)}` : `DTSTART:${utcStamp(s)}`,
    e.allDay ? `DTEND;VALUE=DATE:${dateStamp(Math.max(en, s + 86400000))}` : `DTEND:${utcStamp(en)}`,
    `SUMMARY:${escapeText(e.summary)}`,
  ];
  if (e.location) L.push(`LOCATION:${escapeText(e.location)}`);
  if (e.description) L.push(`DESCRIPTION:${escapeText(e.description)}`);
  if (e.onlineUrl) {
    L.push(`URL:${e.onlineUrl}`);
    // Outlook a Teams podľa tejto vlastnosti ukážu tlačidlo „Pripojiť sa“
    if (/teams\.(microsoft|live)\.com/i.test(e.onlineUrl)) L.push(`X-MICROSOFT-SKYPETEAMSMEETINGURL:${e.onlineUrl}`);
  }
  if (e.rrule) L.push(`RRULE:${e.rrule}`);
  if (e.sequence) L.push(`SEQUENCE:${e.sequence}`);
  if (e.status) L.push(`STATUS:${e.status}`);
  if (e.organizer) L.push(`ORGANIZER${e.organizer.name ? `;CN="${e.organizer.name}"` : ''}:mailto:${e.organizer.email}`);
  for (const a of e.attendees ?? [])
    L.push(
      `ATTENDEE${a.name ? `;CN="${a.name}"` : ''}${a.cutype ? `;CUTYPE=${a.cutype}` : ''};ROLE=${a.role ?? (a.cutype ? 'NON-PARTICIPANT' : 'REQ-PARTICIPANT')};PARTSTAT=${a.partstat ?? 'NEEDS-ACTION'};RSVP=TRUE:mailto:${a.email}`,
    );
  L.push(...extra, 'END:VEVENT');
  return L;
}

const wrap = (lines: string[], method?: string) =>
  ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//MailKlient//SK', 'CALSCALE:GREGORIAN', ...(method ? [`METHOD:${method}`] : []), ...lines, 'END:VCALENDAR']
    .map(fold)
    .join('\r\n') + '\r\n';

export function buildEvent(e: CalEvent): string {
  return wrap(eventLines(e));
}

/** Pozvánka pre účastníkov (METHOD:REQUEST) alebo zrušenie (CANCEL). */
export function buildInvite(e: CalEvent, method: 'REQUEST' | 'CANCEL'): string {
  return wrap(eventLines({ ...e, status: method === 'CANCEL' ? 'CANCELLED' : e.status }), method);
}

/** Odpoveď na pozvánku (METHOD:REPLY) – len vlastný účastník s novým PARTSTAT. */
export function buildReply(e: CalEvent, me: { email: string; name?: string }, partstat: 'ACCEPTED' | 'DECLINED' | 'TENTATIVE'): string {
  const s = Date.parse(e.start);
  const lines = [
    'BEGIN:VEVENT',
    `UID:${e.uid}`,
    `DTSTAMP:${utcStamp(Date.now())}`,
    e.allDay ? `DTSTART;VALUE=DATE:${dateStamp(s)}` : `DTSTART:${utcStamp(s)}`,
    e.allDay ? `DTEND;VALUE=DATE:${dateStamp(Date.parse(e.end))}` : `DTEND:${utcStamp(Date.parse(e.end))}`,
    `SUMMARY:${escapeText(e.summary)}`,
    ...(e.recurrenceId ? [`RECURRENCE-ID:${utcStamp(Date.parse(e.recurrenceId))}`] : []),
    `SEQUENCE:${e.sequence ?? 0}`,
    ...(e.organizer ? [`ORGANIZER${e.organizer.name ? `;CN="${e.organizer.name}"` : ''}:mailto:${e.organizer.email}`] : []),
    `ATTENDEE${me.name ? `;CN="${me.name}"` : ''};PARTSTAT=${partstat}:mailto:${me.email}`,
    'END:VEVENT',
  ];
  return wrap(lines, 'REPLY');
}

// ---------------------------------------------------------------------------
// vCard

export function parseVCards(text: string): Contact[] {
  const tree = toTree(parseLines(text));
  return tree
    .filter((c) => c.type === 'VCARD')
    .map((c) => {
      const fn = get(c, 'FN')?.value;
      const n = get(c, 'N')?.value.split(';');
      const name = unescapeText(fn || [n?.[1], n?.[0]].filter(Boolean).join(' ') || '');
      return {
        uid: get(c, 'UID')?.value ?? '',
        name,
        emails: c.props.filter((p) => p.name === 'EMAIL').map((p) => p.value.trim()).filter(Boolean),
        phones: c.props.filter((p) => p.name === 'TEL').map((p) => p.value.replace(/^tel:/i, '').trim()).filter(Boolean),
        org: get(c, 'ORG') ? unescapeText(get(c, 'ORG')!.value.split(';')[0] ?? '') : undefined,
        title: get(c, 'TITLE') ? unescapeText(get(c, 'TITLE')!.value) : undefined,
        note: get(c, 'NOTE') ? unescapeText(get(c, 'NOTE')!.value) : undefined,
      };
    });
}

export function buildVCard(c: Contact): string {
  const parts = c.name.trim().split(/\s+/);
  const last = parts.length > 1 ? parts.pop()! : '';
  const L = ['BEGIN:VCARD', 'VERSION:3.0', `UID:${c.uid}`, `FN:${escapeText(c.name)}`, `N:${escapeText(last)};${escapeText(parts.join(' '))};;;`];
  for (const e of c.emails) L.push(`EMAIL;TYPE=INTERNET:${e}`);
  for (const p of c.phones) L.push(`TEL:${p}`);
  if (c.org) L.push(`ORG:${escapeText(c.org)}`);
  if (c.title) L.push(`TITLE:${escapeText(c.title)}`);
  if (c.note) L.push(`NOTE:${escapeText(c.note)}`);
  L.push('END:VCARD');
  return L.map(fold).join('\r\n') + '\r\n';
}

// ---------------------------------------------------------------------------
// úlohy (VTODO)

export interface CalTodo {
  uid: string;
  summary: string;
  /** YYYY-MM-DD */
  due?: string;
  done: boolean;
  description?: string;
  href?: string;
  etag?: string;
  lastModified?: string;
}

export function parseTodos(ics: string): CalTodo[] {
  const tree = toTree(parseLines(ics));
  return tree.flatMap((cal) => all(cal, 'VTODO')).map((c) => {
    const due = get(c, 'DUE');
    const st = get(c, 'STATUS')?.value?.toUpperCase();
    const lm = get(c, 'LAST-MODIFIED');
    let dueIso: string | undefined;
    if (due) {
      const d = new Date(parseIcalDate(due.value, due.params.TZID).ms);
      dueIso = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
    }
    return {
      uid: get(c, 'UID')?.value ?? '',
      summary: unescapeText(get(c, 'SUMMARY')?.value ?? ''),
      due: dueIso,
      done: st === 'COMPLETED' || Number(get(c, 'PERCENT-COMPLETE')?.value ?? 0) >= 100,
      description: get(c, 'DESCRIPTION') ? unescapeText(get(c, 'DESCRIPTION')!.value) : undefined,
      lastModified: lm ? new Date(parseIcalDate(lm.value).ms).toISOString() : undefined,
    };
  });
}

export function buildTodo(t: CalTodo): string {
  const L = ['BEGIN:VTODO', `UID:${t.uid}`, `DTSTAMP:${utcStamp(Date.now())}`, `LAST-MODIFIED:${utcStamp(Date.now())}`, `SUMMARY:${escapeText(t.summary)}`];
  if (t.due) L.push(`DUE;VALUE=DATE:${t.due.replace(/-/g, '')}`);
  if (t.description) L.push(`DESCRIPTION:${escapeText(t.description)}`);
  L.push(t.done ? 'STATUS:COMPLETED' : 'STATUS:NEEDS-ACTION', `PERCENT-COMPLETE:${t.done ? 100 : 0}`);
  if (t.done) L.push(`COMPLETED:${utcStamp(Date.now())}`);
  L.push('END:VTODO');
  return wrap(L);
}
