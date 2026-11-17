// Zmluva medzi rendererom a hlavným procesom (cez preload). Len typy.
import type { RuleAction } from './settings';
import type { CalEvent, Contact } from './ical';
import type { AccountConfig, AccountPublic, AnyAccount, ServerConfig } from './account';
import type { AppSettings, LauncherItem } from './settings';
import type { Protocol, ProbeResult } from './tls-types';

export type Result<T> =
  | { ok: true; data: T }
  | { ok: false; error: string; hint?: string; needsTrust?: boolean };

export interface FolderInfo {
  path: string;
  name: string;
  specialUse: string | null;
  depth: number;
}

export interface MessageSummary {
  accountId: string;
  folder: string;
  uid: number;
  subject: string;
  fromName: string;
  fromAddress: string;
  date: string | null;
  seen: boolean;
  flagged: boolean;
  /** odpovedané (\\Answered) */
  answered?: boolean;
  messageId?: string;
  inReplyTo?: string;
  /** IMAP kľúčové slová (kategórie, $Pin …) */
  keywords?: string[];
}

export interface MessageFull extends MessageSummary {
  to: string;
  cc: string;
  html: string | null;
  text: string;
  attachments: { filename: string; size: number; contentType: string }[];
  messageId?: string;
  /** pozvánka na udalosť (iMIP, text/calendar v správe) */
  invite?: { method: string; event: CalEvent; ics: string };
  /** overenie odosielateľa podľa servera (Authentication-Results) a varovania pred phishingom */
  auth?: { spf?: string; dkim?: string; dmarc?: string; verified: boolean; warnings: string[] };
}

export interface SendInput {
  to: string;
  cc?: string;
  subject: string;
  text: string;
  inReplyTo?: { folder: string; uid: number };
  /** Message-ID správy, na ktorú sa odpovedá (vlákna v klientoch). */
  replyMessageId?: string;
  attachments?: { path: string; filename?: string }[];
  bcc?: string;
  /** Formátovaný obsah (HTML z editora); text je verzia v čistom texte. */
  html?: string;
  priority?: 'high' | 'normal' | 'low';
  /** Žiadosť o potvrdenie o prečítaní. */
  readReceipt?: boolean;
  /** Koncept, ktorý sa po odoslaní zmaže. */
  draft?: DraftRef;
  /** Preposlanie: prílohy pôvodnej správy sa pripoja. */
  forward?: { folder: string; uid: number };
  /** Odosielateľ (identita/alias); prázdne = hlavná adresa konta */
  from?: { name: string; email: string };
  /** automatická odpoveď (hlavička Auto-Submitted, bez potvrdení) */
  autoReply?: boolean;
  /** pozvánka / odpoveď na pozvánku (text/calendar) */
  ical?: { method: string; content: string };
}

/** Správa čakajúca na odoslanie (odložené odoslanie alebo 10 s na späťvzatie). */
export interface OutboxItem {
  id: string;
  accountId: string;
  sendAt: string;
  subject: string;
  to: string;
  input: SendInput;
  /** 'undo' = krátke čakanie na späťvzatie, 'later' = naplánované */
  kind: 'undo' | 'later';
}

export interface DraftRef {
  folder: string;
  uid: number;
}

export interface AccountTestResult {
  imap: ProbeResult;
  smtp: ProbeResult;
  login?: { ok: boolean; message: string; hint?: string };
}

export interface WebBounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** Udalosti vložených webov: nová karta (Office z Outlooku), zmena názvu stránky. */
export type WebEvent =
  | { type: 'tab-opened'; group: string; key: string; url: string }
  | { type: 'title'; key: string; title: string };

/** Nastavenie servera z Autoconfig / ISPDB (Thunderbird). */
export interface AutoServer {
  host: string;
  port: number;
  security: import('./account').Security;
  /** email = celý e-mail, local = časť pred @ */
  username: 'email' | 'local' | 'other';
}

export interface AutoConfigResult {
  source: string;
  name?: string;
  imap?: AutoServer;
  smtp?: AutoServer;
}

export interface PickedFile {
  path: string;
  filename: string;
  size: number;
}

export interface MailApi {
  platform: string;
  // kontá
  listAccounts(): Promise<Result<{ accounts: AccountPublic[]; canStoreSecrets: boolean }>>;
  saveAccount(acc: AnyAccount, password?: string): Promise<Result<AccountPublic>>;
  deleteAccount(id: string): Promise<Result<true>>;
  reorderAccounts(ids: string[]): Promise<Result<true>>;
  probe(cfg: ServerConfig, proto: Protocol): Promise<Result<ProbeResult>>;
  autodetect(host: string, proto: Protocol): Promise<Result<{ config?: ServerConfig; result: ProbeResult; tried: string[] }>>;
  testAccount(acc: AccountConfig, password?: string): Promise<Result<AccountTestResult>>;
  autoconfig(email: string): Promise<Result<AutoConfigResult | null>>;
  // pošta (IMAP)
  listFolders(accountId: string): Promise<Result<FolderInfo[]>>;
  listMessages(accountId: string, folder: string, limit?: number): Promise<Result<MessageSummary[]>>;
  getMessage(accountId: string, folder: string, uid: number): Promise<Result<MessageFull>>;
  /** posledné správy z lokálnej cache (okamžite, offline) */
  listCached(accountId: string, folder: string, limit?: number): Promise<Result<MessageSummary[]>>;
  openDraft(accountId: string, folder: string, uid: number): Promise<Result<{ msg: MessageFull; files: PickedFile[] }>>;
  outboxQueue(accountId: string, input: SendInput, sendAt: string, kind: 'undo' | 'later'): Promise<Result<OutboxItem>>;
  outboxCancel(id: string): Promise<Result<OutboxItem | null>>;
  outboxList(): Promise<Result<OutboxItem[]>>;
  /** otvorí koncept v samostatnom okne */
  composeWindow(accountId: string, folder: string, uid: number): Promise<Result<true>>;
  davCollections(accountId: string): Promise<Result<DavCollection[]>>;
  updateCheck(): Promise<Result<boolean>>;
  exportEml(accountId: string, folder: string, uid: number, subject: string): Promise<Result<string>>;
  exportMbox(accountId: string, folder: string): Promise<Result<{ file: string; count: number }>>;
  importFiles(accountId: string, folder: string): Promise<Result<number>>;
  importThunderbird(accountId: string): Promise<Result<{ folders: number; messages: number }>>;
  updateInstall(): Promise<Result<true>>;
  calEvents(accountId: string, calendarIds: string[], from: string, to: string): Promise<Result<CalEvent[]>>;
  calSave(accountId: string, calendarId: string, event: CalEvent, invite: boolean): Promise<Result<CalEvent>>;
  calDelete(accountId: string, event: CalEvent, notify: boolean): Promise<Result<true>>;
  calRespond(accountId: string, folder: string, uid: number, partstat: 'ACCEPTED' | 'TENTATIVE' | 'DECLINED'): Promise<Result<string>>;
  cardList(accountId: string): Promise<Result<Contact[]>>;
  cardSave(accountId: string, bookId: string, contact: Contact): Promise<Result<Contact>>;
  cardDelete(accountId: string, href: string): Promise<Result<true>>;
  setKeywords(accountId: string, folder: string, uids: number[], add: string[], remove: string[]): Promise<Result<true>>;
  /** správy toho istého vlákna z cache (všetky priečinky konta) */
  thread(accountId: string, messageId: string): Promise<Result<MessageSummary[]>>;
  runActions(accountId: string, folder: string, uids: number[], actions: RuleAction[]): Promise<Result<true>>;
  snooze(accountId: string, folder: string, uid: number, messageId: string, subject: string, until: string): Promise<Result<true>>;
  /** spustí pravidlá na uložené správy priečinka; vráti počet zasiahnutých */
  runRules(accountId: string, folder: string): Promise<Result<number>>;
  /** fulltext v lokálnej cache */
  searchLocal(q: string, accountId?: string, folder?: string): Promise<Result<MessageSummary[]>>;
  /** hľadanie na serveri (IMAP SEARCH) */
  searchServer(accountId: string, folder: string, q: string): Promise<Result<MessageSummary[]>>;
  send(accountId: string, msg: SendInput): Promise<Result<{ messageId: string; savedToSent: boolean }>>;
  setFlags(accountId: string, folder: string, uids: number[], flags: { seen?: boolean; flagged?: boolean }): Promise<Result<true>>;
  deleteMessages(accountId: string, folder: string, uids: number[]): Promise<Result<{ movedTo: string | null }>>;
  moveMessages(accountId: string, folder: string, uids: number[], dest: string): Promise<Result<true>>;
  unreadCount(accountId: string): Promise<Result<number>>;
  /** Uloží koncept do priečinka Koncepty (nahradí predchádzajúcu verziu). */
  saveDraft(accountId: string, msg: SendInput, prev?: DraftRef): Promise<Result<DraftRef | null>>;
  deleteDraft(accountId: string, draft: DraftRef): Promise<Result<true>>;
  /** Cesta k súboru pretiahnutému z Prieskumníka. */
  pathForFile(file: File): string;
  archiveMessages(accountId: string, folder: string, uids: number[]): Promise<Result<{ movedTo: string }>>;
  openAttachment(accountId: string, folder: string, uid: number, index: number): Promise<Result<true>>;
  saveAttachment(accountId: string, folder: string, uid: number, index: number): Promise<Result<string | null>>;
  pickFiles(): Promise<Result<PickedFile[]>>;
  openExternal(url: string): Promise<void>;
  // nastavenia a spúšťače
  getSettings(): Promise<Result<AppSettings>>;
  saveSettings(s: AppSettings): Promise<Result<AppSettings>>;
  /** Podpisy klasického Outlooku na tomto počítači (len Windows). */
  outlookSignatures(): Promise<Result<Array<{ name: string; text: string; html?: string }>>>;
  pickProgram(): Promise<Result<{ target: string; name: string } | null>>;
  runLauncher(item: LauncherItem): Promise<Result<string>>;
  programIcon(target: string): Promise<Result<string | null>>;
  // stabilita vykresľovania
  gpuDisabled(): Promise<Result<boolean>>;
  /** Uloží nastavenie a reštartuje aplikáciu. */
  setGpuDisabled(off: boolean): Promise<void>;
  gpuForced(): Promise<Result<boolean>>;
  /** Vynúti grafickú kartu napriek zoznamu blokovaných ovládačov; reštart. */
  setGpuForced(on: boolean): Promise<void>;
  openLog(): Promise<void>;
  logError(msg: string): Promise<void>;
  onAppEvent(cb: (ev: AppEvent) => void): () => void;
  // vložené weby
  /** Otvorí stránku v Edge (Windows) alebo v predvolenom prehliadači. */
  webOpenOutside(url: string): Promise<void>;
  onWebEvent(cb: (ev: WebEvent) => void): () => void;
  /** Vymaže prihlásenie webovej relácie (odhlásiť konto M365 / web). */
  webClear(partition: string): Promise<void>;
}

export type AppEvent =
  | { type: 'gpu-crash'; count: number }
  | { type: 'mail-changed'; accountId: string; folder: string; newCount: number }
  | { type: 'open-message'; accountId: string; folder: string; uid: number }
  | { type: 'outbox-changed' }
  | { type: 'progress'; text: string }
  | { type: 'update'; state: 'downloading' | 'none' | 'ready'; version?: string }
  | { type: 'outbox-sent'; id: string; subject: string; savedToSent: boolean }
  | { type: 'outbox-failed'; id: string; subject: string; error: string };

export interface DavCollection {
  id: string;
  href: string;
  name: string;
  color?: string;
  kind: 'calendar' | 'addressbook';
  events?: boolean;
  todos?: boolean;
  readOnly?: boolean;
}
