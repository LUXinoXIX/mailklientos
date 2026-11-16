// Ukážkové dáta, kým nie je pridané žiadne konto (aby bolo rozloženie hneď vidieť).
import type { AccountPublic } from '../core/account';
import type { FolderInfo, MessageFull } from '../core/api';

export const DEMO_ID = 'demo';

export const demoAccount: AccountPublic = {
  id: DEMO_ID,
  kind: 'imap',
  name: 'Ukážka',
  email: 'ukazka@firma.sk',
  color: '#14b8a6',
  username: 'ukazka',
  hasPassword: false,
  imap: { host: 'mail.firma.sk', port: 993, security: 'ssl', tls: { certMode: 'strict', legacy: false } },
  smtp: { host: 'mail.firma.sk', port: 465, security: 'ssl', tls: { certMode: 'strict', legacy: false } },
};

export const demoFolders: FolderInfo[] = [
  { path: 'INBOX', name: 'Doručená pošta', specialUse: '\\Inbox', depth: 0 },
  { path: 'Drafts', name: 'Koncepty', specialUse: '\\Drafts', depth: 0 },
  { path: 'Sent', name: 'Odoslané', specialUse: '\\Sent', depth: 0 },
  { path: 'Trash', name: 'Odstránené', specialUse: '\\Trash', depth: 0 },
];

const d = (h: number) => new Date(Date.now() - h * 3600_000).toISOString();
const m = (uid: number, fromName: string, subject: string, h: number, text: string, seen = true): MessageFull => ({
  accountId: DEMO_ID, folder: 'INBOX', uid, subject, fromName, fromAddress: `${fromName.split(' ')[0].toLowerCase()}@firma.sk`,
  date: d(h), seen, flagged: false, to: 'ukazka@firma.sk', cc: '', html: null, text, attachments: [],
});

export const demoMessages: MessageFull[] = [
  m(5, 'Jana Kováčová', 'Porada v utorok o 10:00', 1, 'Ahoj,\n\nposielam pozvánku na utorkovú poradu. Prosím priprav stav projektu.\n\nJana', false),
  m(4, 'Zimbra Admin', 'Certifikát servera bude obnovený', 5, 'Dobrý deň,\n\ncertifikát poštového servera bude v piatok obnovený. Klienti s uloženou výnimkou certifikátu sa vás opýtajú na nový odtlačok.\n\nIT oddelenie', false),
  m(3, 'Peter Horváth', 'Re: Rozpočet Q4', 26, 'Vďaka, tabuľku som skontroloval, je to v poriadku.'),
  m(2, 'Microsoft 365', 'Týždenný prehľad aktivity', 50, 'Toto je ukážková správa. Pridaj vlastné IMAP konto cez Domov → Pridať konto.'),
  m(1, 'Lucia Nagyová', 'Fotky z teambuildingu', 75, 'Fotky sú na zdieľanom disku.'),
];
