// Automatická odpoveď (neprítomnosť), keď je MailKlient spustený: každému odosielateľovi najviac raz za 4 dni,
// nie na hromadnú poštu, konferencie, automatické správy ani na adresy typu noreply.
import fs from 'node:fs';
import path from 'node:path';
import { app } from 'electron';
import type { AccountConfig } from '../core/account';
import { log } from './diagnostics';
import * as mail from './mail';
import { loadSettings } from './settings';

const file = () => path.join(app.getPath('userData'), 'vacation-replied.json');
const load = (): Record<string, string> => {
  try {
    return JSON.parse(fs.readFileSync(file(), 'utf8')) as Record<string, string>;
  } catch {
    return {};
  }
};

export interface IncomingHeaders {
  fromAddress: string;
  subject: string;
  messageId?: string;
  /** surové hlavičky List-Id, Precedence, Auto-Submitted */
  headers: string;
}

export async function maybeAutoReply(acc: AccountConfig, m: IncomingHeaders) {
  const v = loadSettings().vacation;
  if (!v?.enabled) return;
  if (v.accountIds?.length && !v.accountIds.includes(acc.id)) return;
  const now = Date.now();
  if (v.from && now < Date.parse(v.from)) return;
  if (v.until && now > Date.parse(v.until) + 86_400_000) return;
  const addr = m.fromAddress.toLowerCase();
  if (!addr || addr === acc.email.toLowerCase()) return;
  if (/(no-?reply|mailer-daemon|postmaster|bounce|notification|newsletter)/i.test(addr)) return;
  if (/^(list-id|list-unsubscribe):/im.test(m.headers)) return;
  if (/^precedence:\s*(bulk|list|junk)/im.test(m.headers)) return;
  if (/^auto-submitted:\s*(?!no\b)/im.test(m.headers)) return;
  const replied = load();
  const key = `${acc.id}|${addr}`;
  if (replied[key] && now - Date.parse(replied[key]) < 4 * 86_400_000) return;
  try {
    await mail.sendMail(acc, {
      to: m.fromAddress,
      subject: `${v.subject || 'Automatická odpoveď'}: ${m.subject}`,
      text: v.text || '',
      html: v.html || undefined,
      replyMessageId: m.messageId,
      autoReply: true,
    });
    replied[key] = new Date().toISOString();
    fs.writeFileSync(file(), JSON.stringify(replied), 'utf8');
    log(`automatická odpoveď odoslaná: ${addr}`);
  } catch (e) {
    log('automatická odpoveď zlyhala', String((e as Error)?.message ?? e));
  }
}
