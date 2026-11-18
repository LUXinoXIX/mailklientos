// Lokálna cache pošty (SQLite vstavané v Node/Electrone, bez natívneho modulu):
// zoznamy správ pre okamžité zobrazenie a offline režim, zdrojové texty otvorených správ, fulltext (FTS5, bez diakritiky).
import path from 'node:path';
import fs from 'node:fs';
import { DatabaseSync } from 'node:sqlite';
import { app } from 'electron';
import type { MessageSummary } from '../core/api';
import { log } from './diagnostics';

let db: DatabaseSync | null = null;
let broken = false;

function open(): DatabaseSync | null {
  if (db || broken) return db;
  try {
    const file = path.join(app.getPath('userData'), 'mailcache.db');
    fs.mkdirSync(path.dirname(file), { recursive: true });
    db = new DatabaseSync(file);
    db.exec(`
      PRAGMA journal_mode = WAL;
      PRAGMA synchronous = NORMAL;
      CREATE TABLE IF NOT EXISTS folders (
        account TEXT NOT NULL, folder TEXT NOT NULL, uidvalidity TEXT, synced_at TEXT,
        PRIMARY KEY (account, folder)
      );
      CREATE TABLE IF NOT EXISTS messages (
        id INTEGER PRIMARY KEY,
        account TEXT NOT NULL, folder TEXT NOT NULL, uid INTEGER NOT NULL,
        subject TEXT, from_name TEXT, from_addr TEXT, recipients TEXT, date TEXT,
        seen INTEGER, flagged INTEGER, answered INTEGER,
        message_id TEXT, in_reply_to TEXT, keywords TEXT,
        source BLOB, body_text TEXT,
        UNIQUE (account, folder, uid)
      );
      CREATE INDEX IF NOT EXISTS messages_list ON messages (account, folder, date DESC);
      CREATE INDEX IF NOT EXISTS messages_msgid ON messages (message_id);
      CREATE VIRTUAL TABLE IF NOT EXISTS fts USING fts5 (
        subject, sender, recipients, body,
        tokenize = 'unicode61 remove_diacritics 2'
      );
    `);
    try {
      db.exec('ALTER TABLE messages ADD COLUMN preview TEXT');
    } catch {
      /* stĺpec už existuje */
    }
    try {
      db.exec('ALTER TABLE messages ADD COLUMN has_att INTEGER');
    } catch {
      /* stĺpec už existuje */
    }
    return db;
  } catch (e) {
    broken = true;
    log('cache: nedá sa otvoriť databáza', String(e));
    return null;
  }
}

function safe<T>(fn: (d: DatabaseSync) => T, fallback: T): T {
  const d = open();
  if (!d) return fallback;
  try {
    return fn(d);
  } catch (e) {
    log('cache: chyba', String(e));
    return fallback;
  }
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Row = any;

const toSummary = (r: Row): MessageSummary => ({
  accountId: r.account,
  folder: r.folder,
  uid: Number(r.uid),
  subject: r.subject ?? '',
  fromName: r.from_name ?? '',
  fromAddress: r.from_addr ?? '',
  date: r.date ?? null,
  seen: !!r.seen,
  flagged: !!r.flagged,
  answered: !!r.answered,
  messageId: r.message_id ?? undefined,
  inReplyTo: r.in_reply_to ?? undefined,
  keywords: r.keywords ? String(r.keywords).split(' ').filter(Boolean) : [],
  preview: r.preview ?? (r.body_text ? String(r.body_text).replace(/\s+/g, ' ').slice(0, 160) : undefined),
  hasAttachments: r.has_att == null ? undefined : !!r.has_att,
});

/** Pri zmene UIDVALIDITY sú staré UID neplatné: priečinok v cache vymazať. */
export function checkUidValidity(account: string, folder: string, uidvalidity: string | number | bigint | undefined) {
  if (uidvalidity === undefined) return;
  const v = String(uidvalidity);
  safe((d) => {
    const r = d.prepare('SELECT uidvalidity FROM folders WHERE account = ? AND folder = ?').get(account, folder) as Row;
    if (r && r.uidvalidity && r.uidvalidity !== v) removeFolder(account, folder);
    d.prepare(
      `INSERT INTO folders (account, folder, uidvalidity, synced_at) VALUES (?, ?, ?, ?)
       ON CONFLICT (account, folder) DO UPDATE SET uidvalidity = excluded.uidvalidity, synced_at = excluded.synced_at`,
    ).run(account, folder, v, new Date().toISOString());
  }, undefined);
}

function ftsWrite(d: DatabaseSync, id: number | bigint, m: { subject?: string; sender?: string; recipients?: string; body?: string }) {
  d.prepare('DELETE FROM fts WHERE rowid = ?').run(id);
  d.prepare('INSERT INTO fts (rowid, subject, sender, recipients, body) VALUES (?, ?, ?, ?, ?)').run(
    id, m.subject ?? '', m.sender ?? '', m.recipients ?? '', m.body ?? '',
  );
}

/**
 * Uloží zoznam zo servera. Správy v rozsahu UID, ktoré server už nemá, sa z cache odstránia (zmazané inde).
 */
export function storeList(account: string, folder: string, list: MessageSummary[], recipients?: Map<number, string>) {
  safe((d) => {
    d.exec('BEGIN');
    try {
      const up = d.prepare(`
        INSERT INTO messages (account, folder, uid, subject, from_name, from_addr, recipients, date, seen, flagged, answered, message_id, in_reply_to, keywords, preview, has_att)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT (account, folder, uid) DO UPDATE SET
          subject = excluded.subject, from_name = excluded.from_name, from_addr = excluded.from_addr,
          recipients = COALESCE(excluded.recipients, recipients), date = excluded.date,
          seen = excluded.seen, flagged = excluded.flagged, answered = excluded.answered,
          message_id = excluded.message_id, in_reply_to = excluded.in_reply_to, keywords = excluded.keywords,
          preview = COALESCE(excluded.preview, preview),
          has_att = COALESCE(excluded.has_att, has_att)
        RETURNING id, body_text`);
      for (const m of list) {
        const rc = recipients?.get(m.uid) ?? null;
        const r = up.get(
          account, folder, m.uid, m.subject, m.fromName, m.fromAddress, rc, m.date,
          m.seen ? 1 : 0, m.flagged ? 1 : 0, m.answered ? 1 : 0, m.messageId ?? null, m.inReplyTo ?? null,
          (m.keywords ?? []).join(' '),
          m.preview ?? null,
          m.hasAttachments == null ? null : m.hasAttachments ? 1 : 0,
        ) as Row;
        ftsWrite(d, r.id, { subject: m.subject, sender: `${m.fromName} ${m.fromAddress}`, recipients: rc ?? '', body: r.body_text ?? '' });
      }
      if (list.length) {
        const minUid = Math.min(...list.map((m) => m.uid));
        const keep = new Set(list.map((m) => m.uid));
        const rows = d.prepare('SELECT id, uid FROM messages WHERE account = ? AND folder = ? AND uid >= ?').all(account, folder, minUid) as Row[];
        for (const r of rows) if (!keep.has(Number(r.uid))) deleteRow(d, r.id);
      }
      d.exec('COMMIT');
    } catch (e) {
      d.exec('ROLLBACK');
      throw e;
    }
  }, undefined);
}

function deleteRow(d: DatabaseSync, id: number | bigint) {
  d.prepare('DELETE FROM fts WHERE rowid = ?').run(id);
  d.prepare('DELETE FROM messages WHERE id = ?').run(id);
}

/** Posledné správy priečinka z cache (okamžite, aj offline). */
export function cachedList(account: string, folder: string, limit = 50): MessageSummary[] {
  return safe(
    (d) =>
      (d.prepare('SELECT * FROM messages WHERE account = ? AND folder = ? ORDER BY date DESC LIMIT ?').all(account, folder, limit) as Row[]).map(toSummary),
    [],
  );
}

export function cachedSummary(account: string, folder: string, uid: number): MessageSummary | null {
  return safe((d) => {
    const r = d.prepare('SELECT * FROM messages WHERE account = ? AND folder = ? AND uid = ?').get(account, folder, uid) as Row;
    return r ? toSummary(r) : null;
  }, null);
}

export function recipientsOf(account: string, folder: string, uid: number): string {
  return safe((d) => {
    const r = d.prepare('SELECT recipients FROM messages WHERE account = ? AND folder = ? AND uid = ?').get(account, folder, uid) as Row;
    return (r?.recipients as string) ?? '';
  }, '');
}

/** Zdrojový text správy (ak bola otvorená), pre otvorenie offline. */
export function cachedSource(account: string, folder: string, uid: number): Buffer | null {
  return safe((d) => {
    const r = d.prepare('SELECT source FROM messages WHERE account = ? AND folder = ? AND uid = ?').get(account, folder, uid) as Row;
    return r?.source ? Buffer.from(r.source as Uint8Array) : null;
  }, null);
}

/** Uloží zdroj otvorenej správy a jej text do fulltextu. Veľké správy (> 25 MB) len text. */
export function storeBody(account: string, folder: string, uid: number, source: Buffer, text: string, recipients: string) {
  safe((d) => {
    const r = d.prepare('SELECT id, subject, from_name, from_addr FROM messages WHERE account = ? AND folder = ? AND uid = ?').get(account, folder, uid) as Row;
    if (!r) return;
    const body = text.slice(0, 200_000);
    d.prepare('UPDATE messages SET source = ?, body_text = ?, recipients = ? WHERE id = ?').run(
      source.length <= 25 * 1024 * 1024 ? source : null, body, recipients, r.id,
    );
    ftsWrite(d, r.id, { subject: r.subject, sender: `${r.from_name} ${r.from_addr}`, recipients, body });
  }, undefined);
}

export function updateFlags(account: string, folder: string, uids: number[], patch: { seen?: boolean; flagged?: boolean; answered?: boolean }) {
  safe((d) => {
    for (const [col, v] of Object.entries(patch)) {
      if (v === undefined || !['seen', 'flagged', 'answered'].includes(col)) continue;
      const st = d.prepare(`UPDATE messages SET ${col} = ? WHERE account = ? AND folder = ? AND uid = ?`);
      for (const uid of uids) st.run(v ? 1 : 0, account, folder, uid);
    }
  }, undefined);
}

export function updateKeywords(account: string, folder: string, uid: number, keywords: string[]) {
  safe((d) => {
    d.prepare('UPDATE messages SET keywords = ? WHERE account = ? AND folder = ? AND uid = ?').run(keywords.join(' '), account, folder, uid);
  }, undefined);
}

export function removeMessages(account: string, folder: string, uids: number[]) {
  safe((d) => {
    const st = d.prepare('SELECT id FROM messages WHERE account = ? AND folder = ? AND uid = ?');
    for (const uid of uids) {
      const r = st.get(account, folder, uid) as Row;
      if (r) deleteRow(d, r.id);
    }
  }, undefined);
}

export function removeFolder(account: string, folder?: string) {
  safe((d) => {
    const rows = (folder
      ? d.prepare('SELECT id FROM messages WHERE account = ? AND folder = ?').all(account, folder)
      : d.prepare('SELECT id FROM messages WHERE account = ?').all(account)) as Row[];
    for (const r of rows) d.prepare('DELETE FROM fts WHERE rowid = ?').run(r.id);
    if (folder) {
      d.prepare('DELETE FROM messages WHERE account = ? AND folder = ?').run(account, folder);
      d.prepare('DELETE FROM folders WHERE account = ? AND folder = ?').run(account, folder);
    } else {
      d.prepare('DELETE FROM messages WHERE account = ?').run(account);
      d.prepare('DELETE FROM folders WHERE account = ?').run(account);
    }
  }, undefined);
}

/** Fulltext: slová ako predpony („rozpo“ nájde „Rozpočet“), bez ohľadu na diakritiku a veľkosť písmen. */
export function search(query: string, opts: { account?: string; folder?: string; limit?: number; field?: string } = {}): MessageSummary[] {
  const words = query
    .normalize('NFC')
    .split(/[\s"'()*:^+-]+/)
    .map((w) => w.trim())
    .filter((w) => w.length > 0)
    .slice(0, 8);
  if (!words.length) return [];
  const cols: Record<string, string> = { all: '', people: '{subject sender recipients}', subject: '{subject}', body: '{body}', from: '{sender}', to: '{recipients}' };
  const col = cols[opts.field ?? 'all'] ?? '';
  const terms = words.map((w) => `"${w.replace(/"/g, '')}"*`).join(' AND ');
  const match = col ? `${col} : (${terms})` : terms;
  return safe((d) => {
    const where = ['fts MATCH ?'];
    const args: (string | number)[] = [match];
    if (opts.account) {
      where.push('m.account = ?');
      args.push(opts.account);
    }
    if (opts.folder) {
      where.push('m.folder = ?');
      args.push(opts.folder);
    }
    args.push(opts.limit ?? 200);
    const rows = d
      .prepare(`SELECT m.* FROM fts JOIN messages m ON m.id = fts.rowid WHERE ${where.join(' AND ')} ORDER BY m.date DESC LIMIT ?`)
      .all(...args) as Row[];
    return rows.map(toSummary);
  }, []);
}

/** Správy s rovnakým vláknom (Message-ID / In-Reply-To) naprieč priečinkami konta – pre konverzácie. */
export function threadOf(account: string, messageId: string): MessageSummary[] {
  return safe((d) => {
    const seen = new Set<string>([messageId]);
    const queue = [messageId];
    const out = new Map<number, Row>();
    while (queue.length && seen.size < 200) {
      const id = queue.shift()!;
      const rows = d
        .prepare('SELECT * FROM messages WHERE account = ? AND (message_id = ? OR in_reply_to = ?)')
        .all(account, id, id) as Row[];
      for (const r of rows) {
        out.set(Number(r.id), r);
        for (const x of [r.message_id, r.in_reply_to]) if (x && !seen.has(x)) (seen.add(x), queue.push(x));
      }
    }
    return [...out.values()].map(toSummary).sort((a, b) => (a.date ?? '').localeCompare(b.date ?? ''));
  }, []);
}

export function close() {
  try {
    db?.close();
  } catch {
    /* nič */
  }
  db = null;
}
