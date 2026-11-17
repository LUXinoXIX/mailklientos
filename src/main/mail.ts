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

async function withImap<T>(acc: AccountConfig, fn: (c: Client) => Promise<T>, pwOverride?: string): Promise<T> {
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

export async function listMessages(acc: AccountConfig, folder: string, limit = 50): Promise<MessageSummary[]> {
  return withImap(acc, async (c) => {
    const lock = await c.getMailboxLock(folder);
    try {
      const total: number = c.mailbox?.exists ?? 0;
      if (!total) return [];
      const from = Math.max(1, total - limit + 1);
      const out: MessageSummary[] = [];
      for await (const m of c.fetch(`${from}:*`, { uid: true, envelope: true, flags: true, internalDate: true })) {
        const f = addr(m.envelope?.from?.[0]);
        const d: Date | undefined = m.envelope?.date ?? m.internalDate;
        out.push({
          accountId: acc.id,
          folder,
          uid: m.uid,
          subject: m.envelope?.subject || '(bez predmetu)',
          fromName: f.name,
          fromAddress: f.address,
          date: d ? new Date(d).toISOString() : null,
          seen: !!m.flags?.has('\\Seen'),
          flagged: !!m.flags?.has('\\Flagged'),
        });
      }
      return out.reverse();
    } finally {
      lock.release();
    }
  });
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const addrText = (x: any): string => (Array.isArray(x) ? x.map((y) => y.text).join(', ') : x?.text ?? '');

export async function getMessage(acc: AccountConfig, folder: string, uid: number): Promise<MessageFull> {
  return withImap(acc, async (c) => {
    const lock = await c.getMailboxLock(folder);
    try {
      const msg = await c.fetchOne(String(uid), { source: true, flags: true }, { uid: true });
      if (!msg?.source) throw new UserError('Správa sa nenašla (mohla byť presunutá alebo zmazaná).');
      const p = await simpleParser(msg.source);
      await c.messageFlagsAdd(String(uid), ['\\Seen'], { uid: true }).catch((): void => undefined);
      const f = p.from?.value?.[0];
      return {
        accountId: acc.id,
        folder,
        uid,
        subject: p.subject || '(bez predmetu)',
        fromName: f?.name || f?.address || '',
        fromAddress: f?.address || '',
        date: p.date ? p.date.toISOString() : null,
        seen: true,
        flagged: !!msg.flags?.has('\\Flagged'),
        messageId: p.messageId || undefined,
        to: addrText(p.to),
        cc: addrText(p.cc),
        html: typeof p.html === 'string' ? p.html : null,
        text: p.text ?? '',
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        attachments: (p.attachments ?? []).map((a: any) => ({
          filename: a.filename || 'príloha',
          size: a.size ?? 0,
          contentType: a.contentType,
        })),
      };
    } finally {
      lock.release();
    }
  });
}

/** Správa pre nodemailer (odoslanie aj koncept). Obrázky vložené v texte (data:) sa prevedú na prílohy cid. */
function buildMail(acc: AccountConfig, input: SendInput) {
  const fromName = acc.name;
  return {
    from: { name: fromName, address: acc.email },
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
    headers: input.readReceipt ? { 'Disposition-Notification-To': acc.email } : undefined,
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
        .filter((a: any) => a.contentDisposition !== 'inline' || !a.cid)
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
  try {
    info = await transport.sendMail(mail);
  } catch (e) {
    throw toUserError(e, acc.smtp);
  } finally {
    transport.close();
  }
  // Uloženie do Odoslaných (IMAP servery to zvyčajne nerobia samy).
  let savedToSent = false;
  try {
    const raw: Buffer = await new MailComposer({ ...mail, messageId: info.messageId, date: new Date() }).compile().build();
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
      return true;
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
      return { movedTo: archive };
    } finally {
      lock.release();
    }
  });
}

/** Obsah prílohy podľa poradia v správe. */
export async function getAttachment(acc: AccountConfig, folder: string, uid: number, index: number) {
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
