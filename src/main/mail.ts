// IMAP (imapflow) a SMTP (nodemailer) nad nastaveniami konta z TLS jadra.
// Každá operácia si zatiaľ otvára vlastné spojenie (jednoduché a odolné);
// trvalé spojenie s IDLE a lokálna cache prídu v ďalšom kroku.
import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';
import nodemailer from 'nodemailer';
import MailComposer from 'nodemailer/lib/mail-composer';
import type { AccountConfig, ServerConfig } from '../core/account';
import type { AccountTestResult, FolderInfo, MessageFull, MessageSummary, SendInput } from '../core/api';
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

async function withImap<T>(acc: AccountConfig, fn: (c: Client) => Promise<T>, pwOverride?: string): Promise<T> {
  const pass = passwordFor(acc, pwOverride);
  await ensureTrusted(acc.imap, 'imap');
  const client: Client = new ImapFlow({
    ...imapflowConnection(acc.imap),
    auth: { user: acc.username, pass },
    logger: false,
    connectionTimeout: 20000,
    greetingTimeout: 15000,
    socketTimeout: 120000,
  });
  client.on('error', () => {
    /* bez tohto by chyba na pozadí zhodila hlavný proces */
  });
  try {
    await client.connect();
    const bad = verifySocket(client.socket ?? client._socket, acc.imap.host, acc.imap.tls);
    if (bad) throw new UserError(bad, undefined, true);
    return await fn(client);
  } catch (e) {
    throw toUserError(e, acc.imap);
  } finally {
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

export async function sendMail(acc: AccountConfig, input: SendInput) {
  const pass = passwordFor(acc);
  await ensureTrusted(acc.smtp, 'smtp');
  const mail = {
    from: { name: acc.name, address: acc.email },
    to: input.to,
    cc: input.cc || undefined,
    subject: input.subject,
    text: input.text,
  };
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
