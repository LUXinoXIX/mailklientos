// Kalendár a kontakty nad CalDAV/CardDAV: výber kalendárov, pozvánky (iMIP) a odpovede na ne.
import * as cache from './cache';
import crypto from 'node:crypto';
import type { AccountConfig } from '../core/account';
import type { DavCollection } from '../core/api';
import { buildInvite, buildReply, parseEvents, type CalEvent, type Contact } from '../core/ical';
import * as dav from './dav';
import * as mail from './mail';
import * as local from './localcal';

const colCache = new Map<string, { at: number; list: DavCollection[] }>();

/** Kalendáre a adresáre zo servera (CalDAV/CardDAV) + lokálne kalendáre a odbery v programe.
    Konto bez kalendára na serveri (bežný IMAP) dostane lokálny „Môj kalendár“. */
export async function collections(acc: AccountConfig, fresh = false): Promise<DavCollection[]> {
  const c = colCache.get(acc.id);
  let server: DavCollection[];
  if (c && !fresh && Date.now() - c.at < 10 * 60_000) server = c.list;
  else {
    try {
      server = await dav.collections(acc);
    } catch (e) {
      // server bez CalDAV/CardDAV: ostanú lokálne kalendáre; chybu hlásiť len ak nie je ani lokálny
      server = [];
      if (!local.collections(acc.id).length) local.ensureDefault(acc.id);
      if (!local.collections(acc.id).length) throw e;
    }
    colCache.set(acc.id, { at: Date.now(), list: server });
  }
  if (!server.some((x) => x.kind === 'calendar')) local.ensureDefault(acc.id);
  return [...server, ...local.collections(acc.id)];
}
const serverOnly = async (acc: AccountConfig) => (await collections(acc)).filter((x) => !local.isLocalId(x.id));

async function calendar(acc: AccountConfig, id?: string): Promise<DavCollection> {
  const list = (await serverOnly(acc)).filter((x) => x.kind === 'calendar' && x.events !== false);
  const cal = (id && list.find((x) => x.id === id)) || list.find((x) => !x.readOnly) || list[0];
  if (!cal) throw new mail.UserError('Na serveri sa nenašiel žiadny kalendár.');
  return cal;
}

export async function events(acc: AccountConfig, ids: string[], from: string, to: string): Promise<CalEvent[]> {
  const all = await collections(acc);
  const cals = all.filter((x) => x.kind === 'calendar' && x.events !== false && !local.isLocalId(x.id) && (!ids.length || ids.includes(x.id)));
  const localIds = all.filter((x) => local.isLocalId(x.id) && (!ids.length || ids.includes(x.id))).map((x) => x.id);
  const res = await Promise.allSettled([...cals.map((c) => dav.events(acc, c, from, to)), ...(localIds.length ? [local.events(acc.id, localIds, from, to)] : [])]);
  const ok = res.flatMap((r) => (r.status === 'fulfilled' ? r.value : []));
  if (!ok.length && res.some((r) => r.status === 'rejected')) throw (res.find((r) => r.status === 'rejected') as PromiseRejectedResult).reason;
  return ok;
}

const myAddresses = (acc: AccountConfig) => [acc.email, ...(acc.identities ?? []).map((x) => x.email)].map((x) => x.toLowerCase());

export async function save(acc: AccountConfig, calendarId: string, e: CalEvent, invite: boolean): Promise<CalEvent> {
  const ev: CalEvent = { ...e, uid: e.uid || `${crypto.randomUUID()}@mailklient`, sequence: e.href ? (e.sequence ?? 0) + 1 : 0 };
  const others = (ev.attendees ?? []).filter((a) => !myAddresses(acc).includes(a.email.toLowerCase()));
  if (others.length && !ev.organizer) ev.organizer = { email: acc.email, name: acc.name };
  const saved = local.isLocalId(calendarId) ? local.save(calendarId, ev) : await dav.saveEvent(acc, await calendar(acc, calendarId), ev);
  if (invite && others.length && ev.organizer && myAddresses(acc).includes(ev.organizer.email.toLowerCase())) {
    await mail.sendMail(acc, {
      to: others.map((a) => (a.name ? `"${a.name}" <${a.email}>` : a.email)).join(', '),
      subject: ev.summary,
      text: `${ev.summary}\n${new Date(ev.start).toLocaleString('sk-SK')}${ev.location ? `\n${ev.location}` : ''}${ev.description ? `\n\n${ev.description}` : ''}`,
      // formátovaný popis aj v tele pozvánky
      html: ev.descriptionHtml
        ? `<p><b>${esc(ev.summary)}</b><br>${esc(new Date(ev.start).toLocaleString('sk-SK'))}${ev.location ? `<br>${esc(ev.location)}` : ''}</p>${ev.descriptionHtml}`
        : undefined,
      ical: { method: 'REQUEST', content: buildInvite(ev, 'REQUEST') },
    });
  }
  return saved;
}

export async function remove(acc: AccountConfig, e: CalEvent, notify: boolean) {
  if (!e.href) return true as const;
  if (local.isLocalId(e.href)) local.removeEvent(e.href);
  else await dav.deleteEvent(acc, e.href, e.etag);
  const others = (e.attendees ?? []).filter((a) => !myAddresses(acc).includes(a.email.toLowerCase()));
  if (notify && others.length && e.organizer && myAddresses(acc).includes(e.organizer.email.toLowerCase())) {
    await mail.sendMail(acc, {
      to: others.map((a) => a.email).join(', '),
      subject: `Zrušené: ${e.summary}`,
      text: `Udalosť „${e.summary}“ (${new Date(e.start).toLocaleString('sk-SK')}) bola zrušená.`,
      ical: { method: 'CANCEL', content: buildInvite({ ...e, sequence: (e.sequence ?? 0) + 1 }, 'CANCEL') },
    });
  }
  return true as const;
}

/** Odpoveď na pozvánku: e-mail organizátorovi (METHOD:REPLY) a pri prijatí uloženie do kalendára. */
export async function respond(acc: AccountConfig, folder: string, uid: number, partstat: 'ACCEPTED' | 'TENTATIVE' | 'DECLINED'): Promise<string> {
  const msg = await mail.getMessage(acc, folder, uid);
  const inv = msg.invite;
  if (!inv) throw new mail.UserError('Správa neobsahuje pozvánku.');
  const mine = myAddresses(acc);
  const me = inv.event.attendees?.find((a) => mine.includes(a.email.toLowerCase())) ?? { email: acc.email, name: acc.name };
  const label = partstat === 'ACCEPTED' ? 'Prijaté' : partstat === 'TENTATIVE' ? 'Predbežne prijaté' : 'Odmietnuté';
  if (inv.event.organizer?.email) {
    await mail.sendMail(acc, {
      to: inv.event.organizer.email,
      subject: `${label}: ${inv.event.summary}`,
      text: `${label}: ${inv.event.summary}`,
      replyMessageId: msg.messageId,
      from: me.email.toLowerCase() !== acc.email.toLowerCase() ? { name: me.name ?? acc.name, email: me.email } : undefined,
      ical: { method: 'REPLY', content: buildReply(inv.event, { email: me.email, name: me.name }, partstat) },
    });
  }
  let saved = '';
  if (partstat !== 'DECLINED') {
    try {
      const cal = await calendar(acc);
      // vlastný PARTSTAT v pôvodnej pozvánke
      const re = new RegExp(`(ATTENDEE[^:\\r\\n]*?)PARTSTAT=[A-Z-]+([^:\\r\\n]*:mailto:${me.email.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, 'i');
      const ics = re.test(inv.ics) ? inv.ics.replace(re, `$1PARTSTAT=${partstat}$2`) : inv.ics;
      await dav.saveRawEvent(acc, cal, inv.event.uid, ics);
      saved = ` Udalosť je v kalendári „${cal.name}“.`;
    } catch (e) {
      saved = ` Do kalendára sa ju nepodarilo uložiť: ${(e as Error).message}`;
    }
  }
  // stav odpovede natrvalo pri správe (kľúčové slovo na serveri + v cache), aby karta pozvánky ukazovala prijaté aj po návrate
  const KW = { ACCEPTED: '$InviteAccepted', TENTATIVE: '$InviteTentative', DECLINED: '$InviteDeclined' } as const;
  try {
    await mail.setKeywords(acc, folder, [uid], [KW[partstat], '\\Answered'], Object.values(KW).filter((k) => k !== KW[partstat]));
  } catch {
    // server nedovolí vlastné kľúčové slová – aspoň v cache programu
    const cur = cache.cachedSummary(acc.id, folder, uid)?.keywords ?? [];
    cache.updateKeywords(acc.id, folder, uid, [...cur.filter((k) => !Object.values(KW).includes(k as never)), KW[partstat]]);
  }
  return `${label} – odpoveď odoslaná organizátorovi.${saved}`;
}

export async function contactList(acc: AccountConfig): Promise<Contact[]> {
  const books = (await collections(acc)).filter((x) => x.kind === 'addressbook');
  const res = await Promise.allSettled(books.map((b) => dav.contacts(acc, b)));
  return res.flatMap((r) => (r.status === 'fulfilled' ? r.value : [])).sort((a, b) => a.name.localeCompare(b.name, 'sk'));
}

export async function contactSave(acc: AccountConfig, bookId: string, c: Contact) {
  const books = (await collections(acc)).filter((x) => x.kind === 'addressbook');
  const book = books.find((b) => b.id === bookId) ?? books[0];
  if (!book) throw new mail.UserError('Na serveri sa nenašiel adresár kontaktov.');
  return dav.saveContact(acc, book, { ...c, uid: c.uid || crypto.randomUUID() });
}

export function forget(accountId: string) {
  colCache.delete(accountId);
  dav.forgetDiscovery(accountId);
}

/** Import súboru .ics do zvoleného kalendára (lokálny naraz, na server po jednej udalosti). */
export async function importIcs(acc: AccountConfig, calendarId: string, text: string): Promise<{ ok: number; failed: number }> {
  const list = parseEvents(text);
  if (!list.length) throw new mail.UserError('V súbore nie sú žiadne udalosti.');
  if (local.isLocalId(calendarId)) return { ok: local.importToLocal(calendarId, list), failed: 0 };
  const cal = await calendar(acc, calendarId);
  let ok = 0;
  let failed = 0;
  // výnimky z opakovania patria k hlavnej udalosti – na server sa posiela len hlavný záznam
  for (const e of list.filter((x) => !x.recurrenceId)) {
    try {
      await dav.saveEvent(acc, cal, { ...e, href: undefined, etag: undefined });
      ok++;
    } catch {
      failed++;
    }
  }
  return { ok, failed };
}

const esc = (t: string) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
