// CalDAV a CardDAV (Zimbra, Nextcloud, iCloud, Google cez heslo aplikácie…): zistenie kalendárov a adresárov,
// čítanie udalostí v rozsahu, uloženie a odstránenie udalosti, kontakty. Prihlásenie rovnaké ako pri IMAP,
// pravidlá pre certifikát rovnaké ako pri IMAP serveri (aj staré Zimbry s neplatným certifikátom).
import https from 'node:https';
import http from 'node:http';
import tls from 'node:tls';
import { DOMParser } from '@xmldom/xmldom';
import { isGmail, type AccountConfig } from '../core/account';
import { buildEvent, buildTodo, buildVCard, expand, parseEvents, parseTodos, parseVCards, type CalEvent, type CalTodo, type Contact } from '../core/ical';
import { buildTlsOptions, verifySocket } from '../core/tls';
import { getPassword } from './accounts';
import * as google from './google';
import { log } from './diagnostics';
import { UserError } from './mail';

import type { DavCollection } from '../core/api';
export type { DavCollection };

interface DavResponse {
  status: number;
  body: string;
  headers: http.IncomingHttpHeaders;
}

/** Hlavička Authorization: heslo (Basic) alebo token Google (Bearer). */
async function authHeader(acc: AccountConfig): Promise<string> {
  if (acc.auth === 'oauth') return `Bearer ${await google.accessToken(acc)}`;
  if (isGmail(acc))
    throw new UserError('Kalendár a kontakty Google vyžadujú prihlásenie cez Google.', 'Heslo pre aplikácie platí len pre poštu. V nastaveniach konta zvoľ „Prihlásiť sa cez Google“.');
  const pass = getPassword(acc.id);
  if (!pass) throw new UserError('Pre konto chýba heslo.');
  return `Basic ${Buffer.from(`${acc.username}:${pass}`).toString('base64')}`;
}

async function request(acc: AccountConfig, method: string, url: string, body?: string, headers: Record<string, string> = {}, redirects = 3): Promise<DavResponse> {
  const r = await requestOnce(acc, await authHeader(acc), method, url, body, headers, redirects);
  // token Google mohol medzičasom vypršať: raz obnoviť a skúsiť znova
  if (r.status === 401 && acc.auth === 'oauth') {
    google.invalidate(acc.id);
    return requestOnce(acc, await authHeader(acc), method, url, body, headers, redirects);
  }
  return r;
}

function requestOnce(acc: AccountConfig, auth: string, method: string, url: string, body: string | undefined, headers: Record<string, string>, redirects: number): Promise<DavResponse> {
  const u = new URL(url);
  const sameHost = u.hostname.toLowerCase() === acc.imap.host.toLowerCase();
  const policy = sameHost ? acc.imap.tls : { certMode: 'strict' as const, legacy: false };
  const lib = u.protocol === 'http:' ? http : https;
  return new Promise((resolve, reject) => {
    const req = lib.request(
      u,
      {
        method,
        ...(u.protocol === 'https:' ? buildTlsOptions(u.hostname, policy) : {}),
        headers: {
          Authorization: auth,
          'User-Agent': 'MailKlient',
          ...(body ? { 'Content-Type': headers['Content-Type'] ?? 'application/xml; charset=utf-8', 'Content-Length': String(Buffer.byteLength(body)) } : {}),
          ...headers,
        },
        timeout: 20000,
      },
      (res) => {
        if (res.socket instanceof tls.TLSSocket && policy.certMode === 'pinned') {
          const bad = verifySocket(res.socket, u.hostname, policy);
          if (bad) {
            res.destroy();
            return reject(new UserError(bad));
          }
        }
        const chunks: Buffer[] = [];
        res.on('data', (c: Buffer) => chunks.push(c));
        res.on('end', () => {
          const status = res.statusCode ?? 0;
          if ([301, 302, 307, 308].includes(status) && res.headers.location && redirects > 0) {
            resolve(requestOnce(acc, auth, method, new URL(res.headers.location, u).toString(), body, headers, redirects - 1));
            return;
          }
          resolve({ status, body: Buffer.concat(chunks).toString('utf8'), headers: res.headers });
        });
      },
    );
    req.on('timeout', () => req.destroy(new Error('Server neodpovedá (časový limit).')));
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}

// ---------------------------------------------------------------------------
// XML

const DAV = 'DAV:';
const CAL = 'urn:ietf:params:xml:ns:caldav';
const CARD = 'urn:ietf:params:xml:ns:carddav';
const APPLE = 'http://apple.com/ns/ical/';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type El = any;

function parseXml(text: string): El {
  return new DOMParser().parseFromString(text, 'text/xml');
}
const els = (n: El, ns: string, name: string): El[] => (n ? Array.from(n.getElementsByTagNameNS(ns, name)) : []);
const first = (n: El, ns: string, name: string): El | undefined => els(n, ns, name)[0];
const text = (n: El | undefined) => (n?.textContent ?? '').trim();

function responses(xml: string) {
  const doc = parseXml(xml);
  return els(doc, DAV, 'response').map((r) => {
    const okProps = els(r, DAV, 'propstat').filter((ps) => /\s2\d\d\s/.test(` ${text(first(ps, DAV, 'status'))} `));
    return { href: text(first(r, DAV, 'href')), props: okProps.map((ps) => first(ps, DAV, 'prop')).filter(Boolean) as El[] };
  });
}
const propEl = (props: El[], ns: string, name: string): El | undefined => props.map((p) => first(p, ns, name)).find(Boolean);

async function propfind(acc: AccountConfig, url: string, depth: '0' | '1', props: string): Promise<ReturnType<typeof responses>> {
  const body = `<?xml version="1.0" encoding="utf-8"?><d:propfind xmlns:d="DAV:" xmlns:c="${CAL}" xmlns:cr="${CARD}" xmlns:a="${APPLE}"><d:prop>${props}</d:prop></d:propfind>`;
  const r = await request(acc, 'PROPFIND', url, body, { Depth: depth });
  if (r.status === 401 || (r.status === 403 && acc.auth === 'oauth'))
    throw new UserError(
      'Server kalendára odmietol prihlásenie.',
      acc.auth === 'oauth'
        ? 'Prihlás sa cez Google znova a povoľ kalendár aj kontakty. V Google Cloud musia byť zapnuté rozhrania CalDAV API a CardDAV API.'
        : 'Skontroluj meno a heslo konta.',
    );
  if (r.status >= 400) throw new UserError(`Server kalendára vrátil chybu ${r.status}.`);
  return responses(r.body);
}

// ---------------------------------------------------------------------------
// zistenie kalendárov a adresárov

const homes = new Map<string, { cal?: string; card?: string; at: number }>();

function candidates(acc: AccountConfig): string[] {
  if (acc.davUrl?.trim()) return [acc.davUrl.trim()];
  const host = acc.imap.host;
  const domain = acc.email.split('@')[1] ?? host;
  const user = encodeURIComponent(acc.username);
  return [
    `https://${host}/.well-known/caldav`,
    `https://${host}/dav/${user}/`, // Zimbra
    `https://${host}/remote.php/dav/`, // Nextcloud
    `https://${domain}/.well-known/caldav`,
  ];
}

/** Google: kalendáre cez CalDAV v2 (principal → calendar-home-set), kontakty cez CardDAV. Záloha: pevné adresy. */
async function discoverGoogle(acc: AccountConfig) {
  const email = encodeURIComponent(acc.email);
  const g = google.grant(acc.id);
  const res: { cal?: string; card?: string; at: number } = { at: Date.now() };
  if (g?.calendar !== false) {
    try {
      const p = await propfind(acc, `https://apidata.googleusercontent.com/caldav/v2/${email}/user`, '0', '<c:calendar-home-set/>');
      const href = text(first(propEl(p[0]?.props ?? [], CAL, 'calendar-home-set'), DAV, 'href'));
      res.cal = href ? new URL(href, 'https://apidata.googleusercontent.com/').toString() : `https://apidata.googleusercontent.com/caldav/v2/${email}/events/`;
    } catch (e) {
      // len hlavný kalendár (zbierka sama sa pri PROPFIND Depth 1 vráti tiež)
      log(`kalendár Google (principal): ${String((e as Error).message)}`);
      res.cal = `https://apidata.googleusercontent.com/caldav/v2/${email}/events/`;
    }
  }
  if (g?.contacts !== false) res.card = `https://www.googleapis.com/carddav/v1/principals/${email}/lists/`;
  if (!res.cal && !res.card) throw new UserError('Kalendár ani kontakty Google nie sú dostupné.', 'Prihlás sa cez Google znova a povoľ kalendár a kontakty.');
  homes.set(acc.id, res);
  log(`kalendár: ${acc.name} (Google) – ${res.cal ?? '-'} / kontakty ${res.card ?? '-'}`);
  return res;
}

async function discover(acc: AccountConfig) {
  const cached = homes.get(acc.id);
  if (cached && Date.now() - cached.at < 3600_000) return cached;
  if (isGmail(acc) && !acc.davUrl?.trim()) {
    if (acc.auth !== 'oauth') await authHeader(acc); // vyhodí zrozumiteľnú chybu
    return discoverGoogle(acc);
  }
  let lastErr: unknown;
  for (const base of candidates(acc)) {
    try {
      const r = await propfind(acc, base, '0', '<d:current-user-principal/><c:calendar-home-set/><cr:addressbook-home-set/>');
      const props = r[0]?.props ?? [];
      let cal = text(first(propEl(props, CAL, 'calendar-home-set'), DAV, 'href'));
      let card = text(first(propEl(props, CARD, 'addressbook-home-set'), DAV, 'href'));
      const principal = text(first(propEl(props, DAV, 'current-user-principal'), DAV, 'href'));
      if ((!cal || !card) && principal) {
        const p = await propfind(acc, new URL(principal, base).toString(), '0', '<c:calendar-home-set/><cr:addressbook-home-set/>');
        const pp = p[0]?.props ?? [];
        cal ||= text(first(propEl(pp, CAL, 'calendar-home-set'), DAV, 'href'));
        card ||= text(first(propEl(pp, CARD, 'addressbook-home-set'), DAV, 'href'));
      }
      if (cal || card) {
        const res = { cal: cal ? new URL(cal, base).toString() : undefined, card: card ? new URL(card, base).toString() : undefined, at: Date.now() };
        homes.set(acc.id, res);
        log(`kalendár: ${acc.name} – ${res.cal ?? '-'} / kontakty ${res.card ?? '-'}`);
        return res;
      }
    } catch (e) {
      lastErr = e;
    }
  }
  if (lastErr instanceof UserError) throw lastErr;
  throw new UserError(
    'Kalendár ani kontakty sa na serveri nenašli.',
    'V nastaveniach konta doplň adresu CalDAV (pri Zimbre zvyčajne https://server/dav/meno@firma.sk/).',
  );
}

export async function collections(acc: AccountConfig): Promise<DavCollection[]> {
  const h = await discover(acc);
  const out: DavCollection[] = [];
  const props =
    '<d:resourcetype/><d:displayname/><a:calendar-color/><c:supported-calendar-component-set/><d:current-user-privilege-set/>';
  for (const [home, kind] of [[h.cal, 'calendar'], [h.card, 'addressbook']] as const) {
    if (!home) continue;
    const list = await propfind(acc, home, '1', props);
    for (const r of list) {
      const rt = propEl(r.props, DAV, 'resourcetype');
      const isCal = !!first(rt, CAL, 'calendar');
      const isBook = !!first(rt, CARD, 'addressbook');
      if ((kind === 'calendar' && !isCal) || (kind === 'addressbook' && !isBook)) continue;
      const comps = els(propEl(r.props, CAL, 'supported-calendar-component-set'), CAL, 'comp').map((c: El) => c.getAttribute('name'));
      const privs = text(propEl(r.props, DAV, 'current-user-privilege-set'));
      const href = new URL(r.href, home).toString();
      out.push({
        id: href,
        href,
        kind,
        name: text(propEl(r.props, DAV, 'displayname')) || decodeURIComponent(r.href.split('/').filter(Boolean).pop() ?? ''),
        color: text(propEl(r.props, APPLE, 'calendar-color')).slice(0, 7) || undefined,
        events: kind === 'calendar' ? !comps.length || comps.includes('VEVENT') : undefined,
        todos: kind === 'calendar' ? comps.includes('VTODO') : undefined,
        readOnly: privs ? !/write/i.test(privs) : false,
      });
    }
  }
  return out;
}

const icalStamp = (ms: number) => new Date(ms).toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');

/** Udalosti kalendára v rozsahu (opakovania rozvinuté). */
export async function events(acc: AccountConfig, cal: DavCollection, fromIso: string, toIso: string): Promise<CalEvent[]> {
  const from = Date.parse(fromIso);
  const to = Date.parse(toIso);
  const body = `<?xml version="1.0" encoding="utf-8"?>
<c:calendar-query xmlns:d="DAV:" xmlns:c="${CAL}">
  <d:prop><d:getetag/><c:calendar-data/></d:prop>
  <c:filter><c:comp-filter name="VCALENDAR"><c:comp-filter name="VEVENT">
    <c:time-range start="${icalStamp(from)}" end="${icalStamp(to)}"/>
  </c:comp-filter></c:comp-filter></c:filter>
</c:calendar-query>`;
  const r = await request(acc, 'REPORT', cal.href, body, { Depth: '1' });
  if (r.status >= 400) throw new UserError(`Kalendár „${cal.name}“ vrátil chybu ${r.status}.`);
  const out: CalEvent[] = [];
  for (const x of responses(r.body)) {
    const data = text(propEl(x.props, CAL, 'calendar-data'));
    if (!data) continue;
    const etag = text(propEl(x.props, DAV, 'getetag'));
    const href = new URL(x.href, cal.href).toString();
    const parsed = parseEvents(data);
    // výnimky z opakovania (RECURRENCE-ID) nahradia príslušný výskyt
    const overrides = parsed.filter((e) => e.recurrenceId);
    for (const e of parsed.filter((p) => !p.recurrenceId)) {
      const exdates = [...e.exdates, ...overrides.filter((o) => o.uid === e.uid).map((o) => Date.parse(o.recurrenceId!))];
      for (const occ of expand({ ...e, exdates }, from, to)) out.push({ ...occ, href, etag, calendarId: cal.id, color: cal.color });
    }
    for (const o of overrides) if (Date.parse(o.start) < to && Date.parse(o.end) > from) out.push({ ...o, href, etag, calendarId: cal.id, color: cal.color });
  }
  return out;
}

/** Uloží (nová alebo upravená) udalosť. Pri opakovanej udalosti sa upravuje celá séria. */
export async function saveEvent(acc: AccountConfig, cal: DavCollection, e: CalEvent): Promise<CalEvent> {
  const href = e.href ?? new URL(`${encodeURIComponent(e.uid)}.ics`, cal.href.endsWith('/') ? cal.href : `${cal.href}/`).toString();
  const r = await request(acc, 'PUT', href, buildEvent({ ...e, recurrenceId: undefined }), {
    'Content-Type': 'text/calendar; charset=utf-8',
    ...(e.href && e.etag ? { 'If-Match': e.etag } : { 'If-None-Match': '*' }),
  });
  if (r.status === 412) throw new UserError('Udalosť medzitým zmenil niekto iný. Obnov kalendár a skús znova.');
  if (r.status >= 400) throw new UserError(`Udalosť sa nepodarilo uložiť (chyba ${r.status}).`);
  return { ...e, href, etag: String(r.headers.etag ?? ''), calendarId: cal.id };
}

/** Uloží pozvánku tak, ako prišla (zachová všetky údaje organizátora), len s vlastným PARTSTAT. */
export async function saveRawEvent(acc: AccountConfig, cal: DavCollection, uid: string, ics: string) {
  const href = new URL(`${encodeURIComponent(uid)}.ics`, cal.href.endsWith('/') ? cal.href : `${cal.href}/`).toString();
  const clean = ics.replace(/^METHOD:.*\r?\n/m, '');
  const r = await request(acc, 'PUT', href, clean, { 'Content-Type': 'text/calendar; charset=utf-8' });
  if (r.status >= 400) throw new UserError(`Udalosť sa nepodarilo uložiť do kalendára (chyba ${r.status}).`);
}

export async function deleteEvent(acc: AccountConfig, href: string, etag?: string) {
  const r = await request(acc, 'DELETE', href, undefined, etag ? { 'If-Match': etag } : {});
  if (r.status >= 400 && r.status !== 404) throw new UserError(`Udalosť sa nepodarilo odstrániť (chyba ${r.status}).`);
}

// ---------------------------------------------------------------------------
// úlohy (VTODO)

export async function todos(acc: AccountConfig, cal: DavCollection): Promise<CalTodo[]> {
  const body = `<?xml version="1.0" encoding="utf-8"?>
<c:calendar-query xmlns:d="DAV:" xmlns:c="${CAL}">
  <d:prop><d:getetag/><c:calendar-data/></d:prop>
  <c:filter><c:comp-filter name="VCALENDAR"><c:comp-filter name="VTODO"/></c:comp-filter></c:filter>
</c:calendar-query>`;
  const r = await request(acc, 'REPORT', cal.href, body, { Depth: '1' });
  if (r.status >= 400) throw new UserError(`Úlohy „${cal.name}“ vrátili chybu ${r.status}.`);
  return responses(r.body).flatMap((x) => {
    const data = text(propEl(x.props, CAL, 'calendar-data'));
    if (!data) return [];
    return parseTodos(data).map((t) => ({ ...t, href: new URL(x.href, cal.href).toString(), etag: text(propEl(x.props, DAV, 'getetag')) }));
  });
}

export async function saveTodo(acc: AccountConfig, cal: DavCollection, t: CalTodo): Promise<CalTodo> {
  const href = t.href ?? new URL(`${encodeURIComponent(t.uid)}.ics`, cal.href.endsWith('/') ? cal.href : `${cal.href}/`).toString();
  const r = await request(acc, 'PUT', href, buildTodo(t), { 'Content-Type': 'text/calendar; charset=utf-8' });
  if (r.status >= 400) throw new UserError(`Úlohu sa nepodarilo uložiť (chyba ${r.status}).`);
  return { ...t, href, etag: String(r.headers.etag ?? '') };
}

export async function deleteTodo(acc: AccountConfig, href: string) {
  const r = await request(acc, 'DELETE', href);
  if (r.status >= 400 && r.status !== 404) throw new UserError(`Úlohu sa nepodarilo odstrániť (chyba ${r.status}).`);
}

// ---------------------------------------------------------------------------
// kontakty

export async function contacts(acc: AccountConfig, book: DavCollection): Promise<Contact[]> {
  const body = `<?xml version="1.0" encoding="utf-8"?>
<cr:addressbook-query xmlns:d="DAV:" xmlns:cr="${CARD}"><d:prop><d:getetag/><cr:address-data/></d:prop></cr:addressbook-query>`;
  const r = await request(acc, 'REPORT', book.href, body, { Depth: '1' });
  if (r.status >= 400) throw new UserError(`Adresár „${book.name}“ vrátil chybu ${r.status}.`);
  return responses(r.body).flatMap((x) => {
    const data = text(propEl(x.props, CARD, 'address-data'));
    if (!data) return [];
    return parseVCards(data).map((c) => ({ ...c, href: new URL(x.href, book.href).toString(), etag: text(propEl(x.props, DAV, 'getetag')), bookId: book.id }));
  });
}

export async function saveContact(acc: AccountConfig, book: DavCollection, c: Contact): Promise<Contact> {
  const href = c.href ?? new URL(`${encodeURIComponent(c.uid)}.vcf`, book.href.endsWith('/') ? book.href : `${book.href}/`).toString();
  const r = await request(acc, 'PUT', href, buildVCard(c), {
    'Content-Type': 'text/vcard; charset=utf-8',
    ...(c.href && c.etag ? { 'If-Match': c.etag } : { 'If-None-Match': '*' }),
  });
  if (r.status >= 400) throw new UserError(`Kontakt sa nepodarilo uložiť (chyba ${r.status}).`);
  return { ...c, href, etag: String(r.headers.etag ?? ''), bookId: book.id };
}

export async function deleteContact(acc: AccountConfig, href: string) {
  const r = await request(acc, 'DELETE', href);
  if (r.status >= 400 && r.status !== 404) throw new UserError(`Kontakt sa nepodarilo odstrániť (chyba ${r.status}).`);
}

export function forgetDiscovery(accountId: string) {
  homes.delete(accountId);
}
