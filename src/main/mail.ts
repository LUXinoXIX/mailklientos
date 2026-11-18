// IMAP (imapflow) a SMTP (nodemailer) nad nastaveniami konta z TLS jadra.
// Operácie zdieľajú jedno trvalé spojenie na konto (pool nižšie).
import crypto from 'node:crypto';
import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';
import nodemailer from 'nodemailer';
import MailComposer from 'nodemailer/lib/mail-composer';
import type { AccountConfig, ServerConfig } from '../core/account';
import type { AccountTestResult, DraftRef, FolderInfo, MessageFull, MessageSummary, SendInput } from '../core/api';
import {
  explainError, explainVerify, imapflowConnection, nodemailerConnection, probeServer, verifySocket,
  type Protocol,
} from '../core/tls';
import { getPassword } from './accounts';
import * as cache from './cache';
import * as secure from './securemail';
import { parseEvents, parseMethod } from '../core/ical';

export class UserError extends Error {
  constructor(message: string, public hint?: string, public needsTrust?: boolean) {
    super(message);
  }
}

/** Preloží chybu knižnice na zrozumiteľnú správu. */
export function toUserError(e: unknown, cfg?: ServerConfig): UserError {
  if (e instanceof UserError) return e;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const err = e as any;
  const text = `${err?.responseText ?? ''} ${err?.response ?? ''} ${err?.message ?? ''}`;
  if (err?.authenticationFailed || /AUTHENTICATIONFAILED|authentication failed|invalid credentials|EAUTH|535/i.test(`${text} ${err?.code ?? ''}`))
    return new UserError('Prihlásenie zlyhalo: nesprávne meno alebo heslo.', 'Pri Zimbre býva meno celý e-mail alebo len časť pred @.');
  if (/LOGINDISABLED|PRIVACYREQUIRED/i.test(text))
    return new UserError('Server nepovoľuje prihlásenie bez šifrovania.', 'Nastav šifrovanie STARTTLS alebo SSL/TLS.');
  const code: string | undefined = err?.code;
  if (code && explainVerify(code) && !explainVerify(code).startsWith('Chyba certifikátu'))
    return new UserError(explainVerify(code), 'V nastaveniach konta daj „Otestovať“ a ak serveru dôveruješ, ulož výnimku pre certifikát.', true);
  const x = explainError(e, cfg?.security ?? 'ssl');
  return new UserError(x.message, x.hint);
}

/** Pri pinovanom certifikáte overí odtlačok pred spojením (knižnica sama overenie vypína). */
async function ensureTrusted(cfg: ServerConfig, proto: Protocol) {
  if (cfg.tls.certMode !== 'pinned') return;
  const r = await probeServer(cfg, proto);
  if (!r.ok) throw new UserError(r.message, r.hint, r.needsTrust);
}

function passwordFor(acc: AccountConfig, override?: string) {
  const pw = override || getPassword(acc.id);
  if (!pw) throw new UserError('Pre konto chýba heslo.', 'Otvor nastavenia konta a zadaj heslo.');
  return pw;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Client = any;

function newClient(acc: AccountConfig, pass: string): Client {
  const client: Client = new ImapFlow({
    ...imapflowConnection(acc.imap),
    auth: { user: acc.username, pass },
    logger: false,
    connectionTimeout: 20000,
    greetingTimeout: 15000,
    socketTimeout: 300000,
  });
  client.on('error', () => {
    /* bez tohto by chyba na pozadí zhodila hlavný proces */
  });
  return client;
}

async function closeClient(client: Client) {
  try {
    await client.logout();
  } catch {
    try {
      client.close();
    } catch {
      /* nič */
    }
  }
}

/** Jednorazové spojenie (test konta s novým heslom). */
async function withOneShot<T>(acc: AccountConfig, fn: (c: Client) => Promise<T>, pass: string): Promise<T> {
  await ensureTrusted(acc.imap, 'imap');
  const client = newClient(acc, pass);
  try {
    await client.connect();
    const bad = verifySocket(client.socket ?? client._socket, acc.imap.host, acc.imap.tls);
    if (bad) throw new UserError(bad, undefined, true);
    return await fn(client);
  } catch (e) {
    throw toUserError(e, acc.imap);
  } finally {
    await closeClient(client);
  }
}

/** Samostatné spojenie na sledovanie schránky (IDLE). Volajúci ho sám zatvára. */
export async function connectWatcher(acc: AccountConfig): Promise<Client> {
  const pass = passwordFor(acc);
  await ensureTrusted(acc.imap, 'imap');
  const client = newClient(acc, pass);
  try {
    await client.connect();
    const bad = verifySocket(client.socket ?? client._socket, acc.imap.host, acc.imap.tls);
    if (bad) throw new UserError(bad, undefined, true);
    return client;
  } catch (e) {
    await closeClient(client);
    throw toUserError(e, acc.imap);
  }
}

export { closeClient };

// ---------------------------------------------------------------------------
// Trvalé spojenia: jedno na konto, znovu použité pre všetky operácie (rýchle otváranie správ).
// Po 5 minútach nečinnosti sa zavrie; pri zmene nastavení alebo hesla sa vytvorí nové.

interface Pooled {
  key: string;
  client: Client;
  ready: Promise<Client>;
  /** spojenie je nadviazané (od tej chvíle sa dá kontrolovať client.usable) */
  ok: boolean;
  idle?: ReturnType<typeof setTimeout>;
}

const pool = new Map<string, Pooled>();
const IDLE_MS = 5 * 60_000;

const poolKey = (acc: AccountConfig, pass: string) =>
  JSON.stringify([acc.imap, acc.username, crypto.createHash('sha256').update(pass).digest('hex')]);

function drop(id: string, p?: Pooled) {
  const cur = pool.get(id);
  if (!cur || (p && cur !== p)) return;
  pool.delete(id);
  if (cur.idle) clearTimeout(cur.idle);
  void closeClient(cur.client);
}

async function pooledClient(acc: AccountConfig): Promise<Pooled> {
  const pass = passwordFor(acc);
  const key = poolKey(acc, pass);
  let p = pool.get(acc.id);
  if (p && (p.key !== key || (p.ok && p.client.usable === false))) {
    drop(acc.id, p);
    p = undefined;
  }
  if (!p) {
    await ensureTrusted(acc.imap, 'imap');
    const client = newClient(acc, pass);
    const entry: Pooled = { key, client, ok: false, ready: Promise.resolve(client) };
    entry.ready = (async () => {
      await client.connect();
      const bad = verifySocket(client.socket ?? client._socket, acc.imap.host, acc.imap.tls);
      if (bad) throw new UserError(bad, undefined, true);
      entry.ok = true;
      return client;
    })();
    client.on('close', () => drop(acc.id, entry));
    pool.set(acc.id, entry);
    p = entry;
  }
  if (p.idle) clearTimeout(p.idle);
  try {
    await p.ready;
  } catch (e) {
    drop(acc.id, p);
    throw e;
  }
  return p;
}

const isConnectionError = (e: unknown) => {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const err = e as any;
  return /NoConnection|ECONNRESET|EPIPE|ETIMEDOUT|Socket|Connection not available|closed/i.test(`${err?.code ?? ''} ${err?.message ?? ''}`);
};

export async function withImap<T>(acc: AccountConfig, fn: (c: Client) => Promise<T>, pwOverride?: string): Promise<T> {
  if (pwOverride) return withOneShot(acc, fn, pwOverride);
  for (let attempt = 0; ; attempt++) {
    let p: Pooled | undefined;
    try {
      p = await pooledClient(acc);
      return await fn(p.client);
    } catch (e) {
      // spojenie medzičasom spadlo (uspatý počítač, server ho zavrel): skúsiť raz znova s novým
      if (p && attempt === 0 && p.ok && (isConnectionError(e) || p.client.usable === false)) {
        drop(acc.id, p);
        continue;
      }
      throw toUserError(e, acc.imap);
    } finally {
      const cur = pool.get(acc.id);
      if (cur) {
        if (cur.idle) clearTimeout(cur.idle);
        cur.idle = setTimeout(() => drop(acc.id, cur), IDLE_MS);
      }
    }
  }
}

/** Zavrie spojenia (pri ukončení aplikácie alebo zmene konta). */
export function closeAllConnections(id?: string) {
  for (const k of [...pool.keys()]) if (!id || k === id) drop(k);
}

const SPECIAL_ORDER = ['\\Inbox', '\\Drafts', '\\Sent', '\\Archive', '\\Junk', '\\Trash'];

export async function listFolders(acc: AccountConfig): Promise<FolderInfo[]> {
  return withImap(acc, async (c) => {
    const list: Array<{ path: string; name: string; specialUse?: string; delimiter?: string }> = await c.list();
    const out = list.map((m) => ({
      path: m.path,
      name: m.path.toUpperCase() === 'INBOX' ? 'Doručená pošta' : m.name,
      specialUse: m.path.toUpperCase() === 'INBOX' ? '\\Inbox' : m.specialUse ?? null,
      depth: m.delimiter ? m.path.split(m.delimiter).length - 1 : 0,
    }));
    const rank = (f: FolderInfo) => {
      const i = f.specialUse ? SPECIAL_ORDER.indexOf(f.specialUse) : -1;
      return i < 0 ? 99 : i;
    };
    return out.sort((a, b) => rank(a) - rank(b) || a.path.localeCompare(b.path));
  });
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const addr = (a: any) => ({ name: a?.name || a?.address || '', address: a?.address || '' });

const SYSTEM_FLAGS = new Set(['\\Seen', '\\Flagged', '\\Answered', '\\Deleted', '\\Draft', '\\Recent']);

/** Krátky náhľad textu z prvých bajtov tela správy (text/plain, inak HTML bez značiek). */
export function previewFrom(head: string, chunk: Buffer): string {
  const decode = (enc: string, body: string, charset: string) => {
    let buf: Buffer;
    if (/base64/i.test(enc)) buf = Buffer.from(body.replace(/[^A-Za-z0-9+/=]/g, '').replace(/=+[^=]*$/, (m) => m.replace(/[^=]/g, '')), 'base64');
    else if (/quoted-printable/i.test(enc)) buf = Buffer.from(body.replace(/=\r?\n/g, '').replace(/=([0-9A-F]{2})/gi, (_m, h: string) => String.fromCharCode(parseInt(h, 16))), 'latin1');
    else buf = Buffer.from(body, 'latin1');
    try {
      return new TextDecoder(charset || 'utf-8').decode(buf);
    } catch {
      return buf.toString('utf8');
    }
  };
  const clean = (t: string, html: boolean) =>
    (html ? t.replace(/<style[\s\S]*?<\/style>/gi, ' ').replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&[a-z]+;/gi, ' ') : t)
      .replace(/^>.*$/gm, '')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 160);
  const raw = chunk.toString('latin1');
  const ct = /content-type:\s*([^;\r\n]+)([^\r\n]*(?:\r?\n[ \t][^\r\n]*)*)/i.exec(head);
  const type = (ct?.[1] ?? 'text/plain').toLowerCase();
  const charset = /charset="?([\w-]+)/i.exec(ct?.[2] ?? '')?.[1] ?? 'utf-8';
  if (type.startsWith('multipart/')) {
    // prvá textová časť
    const parts = raw.split(/\r?\n--[^\r\n]+\r?\n/);
    let html = '';
    for (const p of parts) {
      const i = p.search(/\r?\n\r?\n/);
      if (i < 0) continue;
      const ph = p.slice(0, i);
      const pb = p.slice(i).replace(/^\r?\n\r?\n/, '');
      const pt = /content-type:\s*([^;\r\n]+)/i.exec(ph)?.[1]?.toLowerCase() ?? '';
      const enc = /content-transfer-encoding:\s*([\w-]+)/i.exec(ph)?.[1] ?? '';
      const cs = /charset="?([\w-]+)/i.exec(ph)?.[1] ?? 'utf-8';
      if (pt === 'text/plain') return clean(decode(enc, pb, cs), false);
      if (pt === 'text/html' && !html) html = clean(decode(enc, pb, cs), true);
    }
    return html;
  }
  const enc = /content-transfer-encoding:\s*([\w-]+)/i.exec(head)?.[1] ?? '';
  return clean(decode(enc, raw, charset), type === 'text/html');
}

export async function listMessages(acc: AccountConfig, folder: string, limit = 50, preview = false): Promise<MessageSummary[]> {
  return withImap(acc, async (c) => {
    const lock = await c.getMailboxLock(folder);
    try {
      cache.checkUidValidity(acc.id, folder, c.mailbox?.uidValidity);
      const total: number = c.mailbox?.exists ?? 0;
      if (!total) {
        cache.removeFolder(acc.id, folder);
        return [];
      }
      const from = Math.max(1, total - limit + 1);
      const out: MessageSummary[] = [];
      const recipients = new Map<number, string>();
      const q = preview
        ? { uid: true, envelope: true, flags: true, internalDate: true, headers: ['content-type', 'content-transfer-encoding'], bodyParts: [{ key: 'TEXT', maxLength: 3000 }] }
        : { uid: true, envelope: true, flags: true, internalDate: true };
      for await (const m of c.fetch(`${from}:*`, q)) {
        const f = addr(m.envelope?.from?.[0]);
        const d: Date | undefined = m.envelope?.date ?? m.internalDate;
        const flags: Set<string> = m.flags ?? new Set();
        out.push({
          accountId: acc.id,
          folder,
          uid: m.uid,
          subject: m.envelope?.subject || '(bez predmetu)',
          fromName: f.name,
          fromAddress: f.address,
          date: d ? new Date(d).toISOString() : null,
          seen: flags.has('\\Seen'),
          flagged: flags.has('\\Flagged'),
          answered: flags.has('\\Answered'),
          messageId: m.envelope?.messageId || undefined,
          inReplyTo: m.envelope?.inReplyTo || undefined,
          keywords: [...flags].filter((x) => !SYSTEM_FLAGS.has(x)),
          preview: preview ? previewFrom(m.headers ? m.headers.toString('latin1') : '', (m.bodyParts?.get('text') ?? m.bodyParts?.get('TEXT') ?? Buffer.alloc(0)) as Buffer) || undefined : undefined,
        });
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const rc = [...(m.envelope?.to ?? []), ...(m.envelope?.cc ?? [])].map((x: any) => `${x.name ?? ''} ${x.address ?? ''}`).join(' ');
        recipients.set(m.uid, rc);
      }
      out.reverse();
      cache.storeList(acc.id, folder, out, recipients);
      return out;
    } finally {
      lock.release();
    }
  });
}

/** Koncept na úpravu: obsah + prílohy uložené do dočasných súborov (koncept sa pri ďalšom ukladaní nahrádza). */
export async function openDraft(acc: AccountConfig, folder: string, uid: number) {
  const msg = await getMessage(acc, folder, uid);
  const files: Array<{ path: string; filename: string; size: number }> = [];
  if (msg.attachments.length) {
    const fs = await import('node:fs');
    const os = await import('node:os');
    const path = await import('node:path');
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'mailklient-draft-'));
    for (let i = 0; i < msg.attachments.length; i++) {
      const a = await getAttachment(acc, folder, uid, i);
      const file = path.join(dir, a.filename.replace(/[\\/:*?"<>|]/g, '_'));
      fs.writeFileSync(file, a.content);
      files.push({ path: file, filename: a.filename, size: a.content.length });
    }
  }
  return { msg, files };
}

/** Hľadanie na serveri (IMAP SEARCH): predmet, odosielateľ, príjemca, text. Pre správy, ktoré ešte nie sú v cache. */
export async function searchServer(acc: AccountConfig, folder: string, q: string, limit = 200, field = 'all'): Promise<MessageSummary[]> {
  return withImap(acc, async (c) => {
    const lock = await c.getMailboxLock(folder);
    try {
      const crit: Record<string, object> = {
        all: { or: [{ subject: q }, { from: q }, { to: q }, { cc: q }, { body: q }] },
        people: { or: [{ subject: q }, { from: q }, { to: q }, { cc: q }] },
        subject: { subject: q },
        body: { body: q },
        from: { from: q },
        to: { or: [{ to: q }, { cc: q }] },
      };
      const uids: number[] = (await c.search(crit[field] ?? crit.all!, { uid: true })) || [];
      const pick = uids.sort((a, b) => b - a).slice(0, limit);
      if (!pick.length) return [];
      const out: MessageSummary[] = [];
      for await (const m of c.fetch(uidSet(pick), { uid: true, envelope: true, flags: true, internalDate: true }, { uid: true })) {
        const f = addr(m.envelope?.from?.[0]);
        const d: Date | undefined = m.envelope?.date ?? m.internalDate;
        const flags: Set<string> = m.flags ?? new Set();
        out.push({
          accountId: acc.id, folder, uid: m.uid, subject: m.envelope?.subject || '(bez predmetu)',
          fromName: f.name, fromAddress: f.address, date: d ? new Date(d).toISOString() : null,
          seen: flags.has('\\Seen'), flagged: flags.has('\\Flagged'), answered: flags.has('\\Answered'),
          messageId: m.envelope?.messageId || undefined, inReplyTo: m.envelope?.inReplyTo || undefined,
          keywords: [...flags].filter((x) => !SYSTEM_FLAGS.has(x)),
        });
      }
      return out.sort((a, b) => (b.date ?? '').localeCompare(a.date ?? ''));
    } finally {
      lock.release();
    }
  });
}

/** Zoznam z lokálnej cache (okamžite, aj bez siete). */
export function listCached(accountId: string, folder: string, limit = 50) {
  return cache.cachedList(accountId, folder, limit);
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const addrText = (x: any): string => (Array.isArray(x) ? x.map((y) => y.text).join(', ') : x?.text ?? '');

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function inviteOf(p: any): MessageFull['invite'] {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const part = (p.attachments ?? []).find((a: any) => /^text\/calendar/i.test(a.contentType ?? '') || /\.ics$/i.test(a.filename ?? ''));
  if (!part) return undefined;
  try {
    const ics = (part.content as Buffer).toString('utf8');
    const ev = parseEvents(ics)[0];
    if (!ev) return undefined;
    const method = parseMethod(ics) ?? (/method=([A-Z]+)/i.exec(part.contentType ?? '')?.[1]?.toUpperCase() ?? 'PUBLISH');
    const { exdates: _x, ...event } = ev;
    return { method, event, ics };
  } catch {
    return undefined;
  }
}

const domainOf = (addr?: string) => (addr ?? '').split('@')[1]?.toLowerCase().replace(/>$/, '') ?? '';
const baseDomain = (d: string) => d.split('.').slice(-2).join('.');

/** Overenie odosielateľa: výsledky SPF/DKIM/DMARC zapísané serverom + jednoduché znaky phishingu. */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function authOf(p: any): MessageFull['auth'] {
  const h = p.headers as Map<string, unknown>;
  const raw = ([] as unknown[]).concat(h.get('authentication-results') ?? [], h.get('arc-authentication-results') ?? []).map(String).join(' ; ');
  const pick = (k: string) => new RegExp(`\\b${k}=(\\w+)`, 'i').exec(raw)?.[1]?.toLowerCase();
  let spf = pick('spf');
  if (!spf) spf = /^\s*(pass|fail|softfail|neutral|none)/i.exec(String(h.get('received-spf') ?? ''))?.[1]?.toLowerCase();
  const dkim = pick('dkim');
  const dmarc = pick('dmarc');
  const from = p.from?.value?.[0];
  const fromDomain = domainOf(from?.address);
  const warnings: string[] = [];
  if (dmarc === 'fail') warnings.push('Server overil, že správa neprišla od skutočného vlastníka domény odosielateľa (DMARC).');
  else if (spf === 'fail' && dkim !== 'pass') warnings.push('Server odosielateľa nie je oprávnený posielať poštu za túto doménu (SPF).');
  if (dkim === 'fail' && dmarc !== 'pass') warnings.push('Podpis správy (DKIM) nesedí – správa mohla byť cestou zmenená.');
  const name = String(from?.name ?? '');
  const inName = /[\w.+-]+@[\w-]+\.[\w.-]+/.exec(name)?.[0];
  if (inName && inName.toLowerCase() !== String(from?.address ?? '').toLowerCase())
    warnings.push(`Meno odosielateľa obsahuje inú adresu (${inName}) než skutočný odosielateľ (${from?.address}).`);
  const replyTo = p.replyTo?.value?.[0]?.address as string | undefined;
  if (replyTo && fromDomain && baseDomain(domainOf(replyTo)) !== baseDomain(fromDomain))
    warnings.push(`Odpoveď pôjde na inú doménu (${replyTo}) než od koho správa prišla.`);
  // odkazy, ktorých text ukazuje inú doménu než skutočný cieľ
  if (typeof p.html === 'string') {
    const re = /<a\b[^>]*href=["']?(https?:\/\/[^"' >]+)[^>]*>([\s\S]*?)<\/a>/gi;
    let m: RegExpExecArray | null;
    let n = 0;
    while ((m = re.exec(p.html)) && n < 300) {
      n++;
      const shown = m[2]!.replace(/<[^>]+>/g, '').trim();
      const sd = /^(?:https?:\/\/)?(?:www\.)?([a-z0-9-]+(?:\.[a-z0-9-]+)+)/i.exec(shown)?.[1]?.toLowerCase();
      let real = '';
      try {
        real = new URL(m[1]!).hostname.toLowerCase().replace(/^www\./, '');
      } catch {
        continue;
      }
      if (sd && /\.[a-z]{2,}$/.test(sd) && baseDomain(sd) !== baseDomain(real)) {
        warnings.push(`Odkaz „${shown.slice(0, 60)}“ v skutočnosti vedie na ${real}.`);
        break;
      }
    }
  }
  const verified = dmarc === 'pass' || (dkim === 'pass' && spf === 'pass');
  return { spf, dkim, dmarc, verified, warnings };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function unsubscribeOf(p: any): MessageFull['unsubscribe'] {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const line = (p.headerLines ?? []).find((h: any) => h.key === 'list-unsubscribe')?.line as string | undefined;
  if (!line) return undefined;
  const v = line.replace(/^list-unsubscribe:\s*/i, '').replace(/\r?\n[ \t]+/g, ' ');
  const items = [...v.matchAll(/<([^>]+)>/g)].map((m) => m[1]!.trim());
  const url = items.find((x) => /^https:\/\//i.test(x));
  const mailto = items.find((x) => /^mailto:/i.test(x));
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const post = (p.headerLines ?? []).find((h: any) => h.key === 'list-unsubscribe-post')?.line as string | undefined;
  if (!url && !mailto) return undefined;
  return { url, mailto, oneClick: !!post && /one-click/i.test(post) };
}

async function parseFull(acc: AccountConfig, folder: string, uid: number, source: Buffer, flagged: boolean): Promise<MessageFull> {
  // šifrovaná alebo podpísaná správa (PGP/MIME, S/MIME): dešifrovať a overiť podpis
  const sec = await secure.unwrap(source).catch(() => ({ source, info: undefined }));
  const p = await simpleParser(sec.source);
  if (!sec.info && typeof p.text === 'string' && p.text.includes('-----BEGIN PGP MESSAGE-----')) {
    const plain = await secure.decryptInline(p.text);
    if (plain) {
      p.text = plain;
      p.html = false;
    }
  }
  const f = p.from?.value?.[0];
  return {
    security: sec.info,
    accountId: acc.id,
    folder,
    uid,
    subject: p.subject || '(bez predmetu)',
    fromName: f?.name || f?.address || '',
    fromAddress: f?.address || '',
    date: p.date ? p.date.toISOString() : null,
    seen: true,
    flagged,
    messageId: p.messageId || undefined,
    inReplyTo: p.inReplyTo || undefined,
    to: addrText(p.to),
    cc: addrText(p.cc),
    html: typeof p.html === 'string' ? p.html : null,
    text: p.text ?? '',
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    invite: inviteOf(p),
    unsubscribe: unsubscribeOf(p),
    auth: authOf(p),
    // index = poradie v pôvodnej správe (getAttachment); inline = obrázok vložený v texte (logo v podpise a pod.)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    attachments: (p.attachments ?? []).map((a: any, index: number) => ({ a, index }))
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .filter(({ a }: any) => !/^text\/calendar/i.test(a.contentType ?? ''))
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .map(({ a, index }: any) => ({
        filename: a.filename || 'príloha',
        size: a.size ?? 0,
        contentType: a.contentType,
        index,
        inline: isInlineImage(a) || undefined,
      })),
  };
}

export async function getMessage(acc: AccountConfig, folder: string, uid: number): Promise<MessageFull> {
  // otvorená už predtým: z cache (okamžite aj offline); prečítané sa na serveri označí na pozadí
  const cached = cache.cachedSource(acc.id, folder, uid);
  if (cached) {
    const row = cache.cachedSummary(acc.id, folder, uid);
    void setFlags(acc, folder, [uid], { seen: true }).catch(() => undefined);
    return parseFull(acc, folder, uid, cached, !!row?.flagged);
  }
  return withImap(acc, async (c) => {
    const lock = await c.getMailboxLock(folder);
    try {
      const msg = await c.fetchOne(String(uid), { source: true, flags: true }, { uid: true });
      if (!msg?.source) throw new UserError('Správa sa nenašla (mohla byť presunutá alebo zmazaná).');
      await c.messageFlagsAdd(String(uid), ['\\Seen'], { uid: true }).catch((): void => undefined);
      const full = await parseFull(acc, folder, uid, msg.source, !!msg.flags?.has('\\Flagged'));
      cache.updateFlags(acc.id, folder, [uid], { seen: true });
      // text šifrovaných správ sa do lokálneho fulltextu neukladá
      cache.storeBody(acc.id, folder, uid, msg.source, full.security?.encrypted ? '' : `${full.text}`, `${full.to} ${full.cc}`);
      return full;
    } finally {
      lock.release();
    }
  });
}

/** Správa pre nodemailer (odoslanie aj koncept). Obrázky vložené v texte (data:) sa prevedú na prílohy cid. */
function buildMail(acc: AccountConfig, input: SendInput) {
  const from = input.from?.email ? { name: input.from.name || acc.name, address: input.from.email } : { name: acc.name, address: acc.email };
  const headers: Record<string, string> = {};
  if (input.readReceipt && !input.autoReply) headers['Disposition-Notification-To'] = from.address;
  if (input.autoReply) {
    headers['Auto-Submitted'] = 'auto-replied';
    headers['X-Auto-Response-Suppress'] = 'All';
  }
  return {
    from,
    to: input.to || undefined,
    cc: input.cc || undefined,
    bcc: input.bcc || undefined,
    subject: input.subject,
    text: input.text,
    html: input.html || undefined,
    attachDataUrls: true,
    attachments: input.attachments?.map((a) => ({ path: a.path, filename: a.filename })),
    inReplyTo: input.replyMessageId || undefined,
    references: input.replyMessageId || undefined,
    priority: input.priority ?? 'normal',
    headers: Object.keys(headers).length ? headers : undefined,
    icalEvent: input.ical ? { method: input.ical.method, content: input.ical.content } : undefined,
  };
}

async function findFolder(c: Client, use: string, names: RegExp): Promise<string | undefined> {
  const list: Array<{ path: string; specialUse?: string }> = await c.list();
  return (list.find((m) => m.specialUse === use) ?? list.find((m) => names.test(m.path)))?.path;
}

/** Uloží koncept do Konceptov (IMAP APPEND) a zmaže predchádzajúcu verziu. */
export async function saveDraft(acc: AccountConfig, input: SendInput, prev?: DraftRef): Promise<DraftRef | null> {
  // MailComposer vie data: obrázky previesť len pri odosielaní, preto ich tu nechávame v HTML
  const raw: Buffer = await new MailComposer({ ...buildMail(acc, input), attachDataUrls: true, date: new Date() }).compile().build();
  return withImap(acc, async (c) => {
    let folder = await findFolder(c, '\\Drafts', /^(drafts|koncept|draft|inbox[./]drafts)/i);
    if (!folder) {
      await c.mailboxCreate('Drafts').catch((): void => undefined);
      folder = 'Drafts';
    }
    const res = await c.append(folder, raw, ['\\Seen', '\\Draft']);
    if (prev) await deleteDraftIn(c, prev);
    return res?.uid ? { folder, uid: Number(res.uid) } : null;
  });
}

async function deleteDraftIn(c: Client, d: DraftRef) {
  const lock = await c.getMailboxLock(d.folder);
  try {
    await c.messageDelete(String(d.uid), { uid: true });
  } catch {
    /* koncept už neexistuje */
  } finally {
    lock.release();
  }
}

export async function deleteDraft(acc: AccountConfig, d: DraftRef) {
  await withImap(acc, (c) => deleteDraftIn(c, d));
  return true as const;
}

/** Obrázok, ktorý je súčasťou HTML textu (cid:, multipart/related) – logo v podpise, vložený obrázok. Nie je to „skutočná“ príloha. */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function isInlineImage(a: any): boolean {
  if (!/^image\//i.test(a.contentType ?? '')) return false;
  if (!a.cid) return false;
  return a.related === true || a.contentDisposition === 'inline';
}

/** Prílohy pôvodnej správy pri preposlaní. */
async function forwardedAttachments(acc: AccountConfig, f: { folder: string; uid: number }) {
  return withImap(acc, async (c) => {
    const lock = await c.getMailboxLock(f.folder);
    try {
      const msg = await c.fetchOne(String(f.uid), { source: true }, { uid: true });
      if (!msg?.source) return [];
      const p = await simpleParser(msg.source);
      return (p.attachments ?? [])
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        .filter((a: any) => !isInlineImage(a))
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        .map((a: any) => ({ filename: (a.filename as string) || 'priloha', content: a.content as Buffer, contentType: a.contentType }));
    } finally {
      lock.release();
    }
  });
}

export async function sendMail(acc: AccountConfig, input: SendInput) {
  const pass = passwordFor(acc);
  await ensureTrusted(acc.smtp, 'smtp');
  const base = buildMail(acc, input);
  const fw = input.forward ? await forwardedAttachments(acc, input.forward) : [];
  const mail = { ...base, attachments: [...(base.attachments ?? []), ...fw] };
  const transport = nodemailer.createTransport({
    ...nodemailerConnection(acc.smtp),
    auth: { user: acc.username, pass },
    connectionTimeout: 20000,
    greetingTimeout: 15000,
  });
  let info;
  let secureRaw: Buffer | null = null;
  try {
    if (input.security?.sign || input.security?.encrypt) {
      const messageId = `<${crypto.randomUUID()}@${(mail.from as { address: string }).address.split('@')[1] ?? 'mailklient'}>`;
      const built = await secure.buildSecure({ ...mail, messageId, date: new Date() } as Record<string, unknown>, input.security, (mail.from as { address: string }).address);
      secureRaw = built.raw;
      info = await transport.sendMail({ envelope: { from: (mail.from as { address: string }).address, to: built.recipients }, raw: built.raw });
      info.messageId = messageId;
    } else info = await transport.sendMail(mail);
  } catch (e) {
    if (e instanceof secure.SecureError) throw new UserError(e.message);
    throw toUserError(e, acc.smtp);
  } finally {
    transport.close();
  }
  // Uloženie do Odoslaných (IMAP servery to zvyčajne nerobia samy).
  let savedToSent = false;
  try {
    const raw: Buffer = secureRaw ?? (await new MailComposer({ ...mail, messageId: info.messageId, date: new Date() }).compile().build());
    await withImap(acc, async (c) => {
      const list: Array<{ path: string; specialUse?: string }> = await c.list();
      const sent = list.find((m) => m.specialUse === '\\Sent') ?? list.find((m) => /^(sent|odoslan)/i.test(m.path));
      if (sent) {
        await c.append(sent.path, raw, ['\\Seen']);
        savedToSent = true;
      }
    });
  } catch {
    /* odoslanie prebehlo, uloženie kópie je len doplnok */
  }
  if (input.draft) await deleteDraft(acc, input.draft).catch((): void => undefined);
  return { messageId: String(info.messageId ?? ''), savedToSent };
}

/** Test pred uložením konta: sonda IMAP a SMTP a skúšobné prihlásenie IMAP. */
export async function testAccount(acc: AccountConfig, password?: string): Promise<AccountTestResult> {
  const [imap, smtp] = await Promise.all([probeServer(acc.imap, 'imap'), probeServer(acc.smtp, 'smtp')]);
  const res: AccountTestResult = { imap, smtp };
  if (imap.ok) {
    try {
      await withImap(acc, async () => true, password);
      res.login = { ok: true, message: 'Prihlásenie do IMAP prebehlo úspešne.' };
    } catch (e) {
      const u = toUserError(e, acc.imap);
      res.login = { ok: false, message: u.message, hint: u.hint };
    }
  }
  return res;
}

// ---------------------------------------------------------------------------
// Štandardné operácie so správami: príznaky, mazanie, presun, prílohy.

const uidSet = (uids: number[]) => uids.map(String).join(',');

/** Nastaví príznaky Prečítané / Označené (vlajka). */
export async function setFlags(acc: AccountConfig, folder: string, uids: number[], flags: { seen?: boolean; flagged?: boolean }) {
  if (!uids.length) return true;
  return withImap(acc, async (c) => {
    const lock = await c.getMailboxLock(folder);
    try {
      const set = uidSet(uids);
      if (flags.seen === true) await c.messageFlagsAdd(set, ['\\Seen'], { uid: true });
      if (flags.seen === false) await c.messageFlagsRemove(set, ['\\Seen'], { uid: true });
      if (flags.flagged === true) await c.messageFlagsAdd(set, ['\\Flagged'], { uid: true });
      if (flags.flagged === false) await c.messageFlagsRemove(set, ['\\Flagged'], { uid: true });
      cache.updateFlags(acc.id, folder, uids, flags);
      return true;
    } finally {
      lock.release();
    }
  });
}

/** Kľúčové slová (kategórie, pripnutie) na serveri a v cache. */
export async function setKeywords(acc: AccountConfig, folder: string, uids: number[], add: string[], remove: string[]) {
  if (!uids.length) return true;
  return withImap(acc, async (c) => {
    const lock = await c.getMailboxLock(folder);
    try {
      const set = uidSet(uids);
      if (add.length) await c.messageFlagsAdd(set, add, { uid: true });
      if (remove.length) await c.messageFlagsRemove(set, remove, { uid: true });
      for (const uid of uids) {
        const cur = cache.cachedSummary(acc.id, folder, uid)?.keywords ?? [];
        cache.updateKeywords(acc.id, folder, uid, [...new Set([...cur.filter((k) => !remove.includes(k)), ...add])]);
      }
      return true;
    } finally {
      lock.release();
    }
  });
}

/** Vytvorí priečinok, ak neexistuje; vráti jeho cestu. */
export async function ensureFolder(acc: AccountConfig, name: string): Promise<string> {
  return withImap(acc, async (c) => {
    const list: Array<{ path: string; name: string }> = await c.list();
    const hit = list.find((m) => m.path === name || m.name === name);
    if (hit) return hit.path;
    await c.mailboxCreate(name).catch((): void => undefined);
    return name;
  });
}

/** Nájde UID správy podľa Message-ID v priečinku (po presune sa UID mení). */
export async function findByMessageId(acc: AccountConfig, folder: string, messageId: string): Promise<number | null> {
  return withImap(acc, async (c) => {
    const lock = await c.getMailboxLock(folder);
    try {
      const uids: number[] = (await c.search({ header: { 'message-id': messageId } }, { uid: true })) || [];
      return uids.length ? uids[uids.length - 1]! : null;
    } finally {
      lock.release();
    }
  });
}

async function findSpecial(c: Client, use: string, names: RegExp): Promise<string | undefined> {
  const list: Array<{ path: string; specialUse?: string }> = await c.list();
  return (list.find((m) => m.specialUse === use) ?? list.find((m) => names.test(m.path)))?.path;
}

/** Odstráni správy: presun do Odstránených; ak už sú v Odstránených (alebo kôš neexistuje), zmaže natrvalo. */
export async function deleteMessages(acc: AccountConfig, folder: string, uids: number[]) {
  if (!uids.length) return { movedTo: null as string | null };
  return withImap(acc, async (c) => {
    const trash = await findSpecial(c, '\\Trash', /^(trash|deleted|odstr|kôš|kos|inbox[./]trash)/i);
    const lock = await c.getMailboxLock(folder);
    try {
      cache.removeMessages(acc.id, folder, uids);
      if (trash && trash !== folder) {
        await c.messageMove(uidSet(uids), trash, { uid: true });
        return { movedTo: trash as string | null };
      }
      await c.messageDelete(uidSet(uids), { uid: true });
      return { movedTo: null as string | null };
    } finally {
      lock.release();
    }
  });
}

/** Presunie správy do iného priečinka. */
export async function moveMessages(acc: AccountConfig, folder: string, uids: number[], dest: string) {
  if (!uids.length || dest === folder) return true;
  return withImap(acc, async (c) => {
    const lock = await c.getMailboxLock(folder);
    try {
      await c.messageMove(uidSet(uids), dest, { uid: true });
      cache.removeMessages(acc.id, folder, uids);
      return true;
    } finally {
      lock.release();
    }
  });
}

/** Archivovať: presun do priečinka Archív (vytvorí ho, ak neexistuje). */
export async function archiveMessages(acc: AccountConfig, folder: string, uids: number[]) {
  if (!uids.length) return { movedTo: '' };
  return withImap(acc, async (c) => {
    let archive = await findSpecial(c, '\\Archive', /^(archive|archív|archiv)$/i);
    if (!archive) {
      await c.mailboxCreate('Archive').catch((): void => undefined);
      archive = 'Archive';
    }
    const lock = await c.getMailboxLock(folder);
    try {
      await c.messageMove(uidSet(uids), archive, { uid: true });
      cache.removeMessages(acc.id, folder, uids);
      return { movedTo: archive };
    } finally {
      lock.release();
    }
  });
}

/** Obsah prílohy podľa poradia v správe. */
export async function getAttachment(acc: AccountConfig, folder: string, uid: number, index: number) {
  const cached = cache.cachedSource(acc.id, folder, uid);
  if (cached) {
    const p = await simpleParser(cached);
    const a = p.attachments?.[index];
    if (a) return { filename: (a.filename as string) || `priloha-${index + 1}`, content: a.content as Buffer };
  }
  return withImap(acc, async (c) => {
    const lock = await c.getMailboxLock(folder);
    try {
      const msg = await c.fetchOne(String(uid), { source: true }, { uid: true });
      if (!msg?.source) throw new UserError('Správa sa nenašla.');
      const p = await simpleParser(msg.source);
      const a = p.attachments?.[index];
      if (!a) throw new UserError('Príloha sa nenašla.');
      return { filename: (a.filename as string) || `priloha-${index + 1}`, content: a.content as Buffer };
    } finally {
      lock.release();
    }
  });
}

/** Počet neprečítaných v Doručenej pošte (pre pás kont). */
export async function unreadCount(acc: AccountConfig): Promise<number> {
  return withImap(acc, async (c) => {
    const st = await c.status('INBOX', { unseen: true });
    return Number(st?.unseen ?? 0);
  });
}

/** Obsadenosť schránky cez IMAP QUOTA (ak ju server podporuje). */
export async function imapQuota(acc: AccountConfig): Promise<{ used: number; limit: number } | null> {
  return withImap(acc, async (c) => {
    const q = await c.getQuota('INBOX').catch((): null => null);
    if (!q?.storage) return null;
    return { used: Number(q.storage.used ?? q.storage.usage ?? 0), limit: Number(q.storage.limit ?? 0) };
  });
}

/** Priečinky: nový, premenovať, odstrániť, všetko prečítané, vyprázdniť. */
export async function folderOp(
  acc: AccountConfig,
  op: 'create' | 'rename' | 'delete' | 'markRead' | 'empty',
  path: string,
  name?: string,
): Promise<string> {
  return withImap(acc, async (c) => {
    const list: Array<{ path: string; delimiter?: string }> = await c.list();
    const delim = list.find((m) => m.path === path)?.delimiter ?? list.find((m) => m.delimiter)?.delimiter ?? '/';
    if (op === 'create') {
      const target = path ? `${path}${delim}${name}` : name!;
      await c.mailboxCreate(target);
      return target;
    }
    if (op === 'rename') {
      const parts = path.split(delim);
      parts[parts.length - 1] = name!;
      const target = parts.join(delim);
      await c.mailboxRename(path, target);
      cache.removeFolder(acc.id, path);
      return target;
    }
    if (op === 'delete') {
      await c.mailboxDelete(path);
      cache.removeFolder(acc.id, path);
      return path;
    }
    const lock = await c.getMailboxLock(path);
    try {
      if (!c.mailbox || !c.mailbox.exists) return path;
      if (op === 'markRead') await c.messageFlagsAdd('1:*', ['\\Seen']);
      else {
        await c.messageDelete('1:*');
        cache.removeFolder(acc.id, path);
      }
      return path;
    } finally {
      lock.release();
    }
  });
}
