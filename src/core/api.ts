// Zmluva medzi rendererom a hlavným procesom (cez preload). Len typy.
import type { RuleAction, Task } from './settings';
import type { CalEvent, Contact } from './ical';
import type { AccountConfig, AccountPublic, AnyAccount, ServerConfig } from './account';
import type { AppSettings, LauncherItem } from './settings';
import type { Protocol, ProbeResult } from './tls-types';

export type Result<T> =
  | { ok: true; data: T }
  | { ok: false; error: string; hint?: string; needsTrust?: boolean };

export interface BackupInfo {
  token: string;
  file: string;
  exported: string;
  platform: string;
  app: string;
  has: { accounts: number; contacts: number; calendars: number };
  accountsExisting: number;
  launchers: { name: string; ok: boolean }[];
}

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
  /** začiatok textu správy (riadok s náhľadom v zozname) */
  preview?: string;
  /** správa má skutočné prílohy (spinka v zozname; logá z podpisu sa nerátajú) */
  hasAttachments?: boolean;
  /** Gmail: ID vlákna (X-GM-THRID) – konverzácie presne ako v Gmaile */
  threadId?: string;
}

/** odpoveď alebo preposlanie nájdené v Odoslaných */
export interface ReplyInfo {
  kind: 'reply' | 'forward';
  date: string | null;
  to: string;
  folder: string;
  uid: number;
}

export interface MessageFull extends MessageSummary {

  to: string;
  cc: string;
  html: string | null;
  text: string;
  /** index = poradie v pôvodnej správe (pre otvorenie/uloženie); inline = obrázok z textu správy (logo v podpise) */
  attachments: { filename: string; size: number; contentType: string; index?: number; inline?: boolean }[];
  messageId?: string;
  /** pozvánka na udalosť (iMIP, text/calendar v správe) */
  invite?: { method: string; event: CalEvent; ics: string };
  /** odhlásenie z newslettera (List-Unsubscribe) */
  unsubscribe?: { url?: string; mailto?: string; oneClick?: boolean };
  /** šifrovanie / podpis (PGP, S/MIME) */
  security?: {
    encrypted?: 'pgp' | 'smime';
    decryptError?: string;
    signed?: 'pgp' | 'smime';
    signatureValid?: boolean;
    signer?: string;
    trusted?: boolean;
    signatureNote?: string;
  };
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
  forward?: { folder: string; uid: number; /** celá pôvodná správa ako príloha .eml */ asAttachment?: boolean };
  /** pôvodná správa, na ktorú sa odpovedá / ktorá sa preposiela – po odoslaní sa označí (\\Answered / $Forwarded) */
  origin?: { accountId: string; folder: string; uid: number; messageId?: string; kind: 'reply' | 'forward' };
  /** Odosielateľ (identita/alias); prázdne = hlavná adresa konta */
  from?: { name: string; email: string };
  /** automatická odpoveď (hlavička Auto-Submitted, bez potvrdení) */
  autoReply?: boolean;
  /** pozvánka / odpoveď na pozvánku (text/calendar) */
  ical?: { method: string; content: string };
  /** pripomenúť, ak príjemca neodpovie do toľkých dní */
  followUpDays?: number;
  /** podpísať / zašifrovať (PGP alebo S/MIME) */
  security?: { sign?: 'pgp' | 'smime'; encrypt?: 'pgp' | 'smime' };
}

/** Správa čakajúca na odoslanie (odložené odoslanie alebo 10 s na späťvzatie). */
export interface OutboxItem {
  /** vrátená do fronty po chybe odoslania */
  retry?: boolean;
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

/** Prihlásenie cez Google: čo konto povolilo. */
export interface GoogleGrant {
  email: string;
  name?: string;
  mail: boolean;
  calendar: boolean;
  contacts: boolean;
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
  listMessages(accountId: string, folder: string, limit?: number, preview?: boolean): Promise<Result<MessageSummary[]>>;
  /** správa v samostatnom okne */
  messageWindow(accountId: string, folder: string, uid: number): Promise<Result<true>>;
  /** odpoveď / preposlanie v samostatnom okne písania */
  replyWindow(accountId: string, folder: string, uid: number, mode: 'reply' | 'replyAll' | 'forward' | 'forwardAttach'): Promise<Result<true>>;
  /** oznámiť ostatným oknám zmenu v priečinku */
  notifyChanged(accountId: string, folder: string): Promise<Result<true>>;
  /** Poslať ďalej ako originál (presmerovanie): pôvodná správa bez zmeny, hlavičky Resent-* */
  redirectMessage(accountId: string, folder: string, uid: number, to: string): Promise<Result<void>>;
  /** ďalšie okno s celým programom */
  newWindow(): Promise<Result<void>>;
  /** odpovede a preposlania správy (hľadá v Odoslaných podľa In-Reply-To / References) */
  replyInfo(accountId: string, messageId: string): Promise<Result<ReplyInfo[]>>;
  /** Odoslať a prijať: znovu pripojiť sledovanie pošty, odoslať správy čakajúce po chybe */
  sendReceive(): Promise<Result<{ retried: number }>>;
  /** priečinky: nový podpriečinok, premenovať, odstrániť, označiť všetko ako prečítané, vyprázdniť */
  folderOp(accountId: string, op: 'create' | 'rename' | 'delete' | 'markRead' | 'empty', path: string, name?: string): Promise<Result<string>>;
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
  testNotification(): Promise<Result<string>>;
  printMessage(accountId: string, folder: string, uid: number, pdf: boolean): Promise<Result<string>>;
  previewAttachment(accountId: string, folder: string, uid: number, index: number): Promise<Result<{ kind: 'image' | 'window' | 'external'; src?: string; filename: string }>>;
  /** náhľad súboru priloženého k rozpísanej správe */
  previewFile(file: string, filename?: string): Promise<Result<{ kind: 'image' | 'window' | 'external'; src?: string; filename: string }>>;
  unsubscribe(accountId: string, u: { url?: string; mailto?: string; oneClick?: boolean }): Promise<Result<string>>;
  setUnread(total: number, badge?: string): Promise<Result<true>>;
  openInOffice(url: string): Promise<Result<true>>;
  appInfo(): Promise<Result<{ version: string; electron: string; chrome: string; node: string; platform: string }>>;
  zimbraCheck(accountId: string): Promise<Result<boolean>>;
  zimbraGal(accountId: string, q: string): Promise<Result<GalEntry[]>>;
  zimbraGalAuto(accountId: string, q: string): Promise<Result<GalEntry[]>>;
  zimbraFreeBusy(accountId: string, emails: string[], from: number, to: number): Promise<Result<FreeBusy[]>>;
  zimbraResources(accountId: string): Promise<Result<ZResource[]>>;
  zimbraShares(accountId: string): Promise<Result<ZShare[]>>;
  zimbraMount(accountId: string, share: ZShare, name: string): Promise<Result<true>>;
  zimbraGetOoo(accountId: string): Promise<Result<ServerOoo>>;
  zimbraSetOoo(accountId: string, o: ServerOoo): Promise<Result<true>>;
  zimbraGetFilters(accountId: string): Promise<Result<ServerRule[]>>;
  zimbraSetFilters(accountId: string, rules: ServerRule[]): Promise<Result<true>>;
  zimbraSignatures(accountId: string): Promise<Result<Array<{ id: string; name: string; html?: string; text?: string }>>>;
  zimbraSaveSignature(accountId: string, s: { id?: string; name: string; html: string; text: string }): Promise<Result<string>>;
  zimbraIdentities(accountId: string): Promise<Result<Array<{ name: string; email: string; signatureId?: string }>>>;
  zimbraTags(accountId: string): Promise<Result<Array<{ id: string; name: string; color: string }>>>;
  zimbraBriefcase(accountId: string, q: string): Promise<Result<BriefcaseDoc[]>>;
  zimbraDownload(accountId: string, d: BriefcaseDoc): Promise<Result<PickedFile>>;
  zimbraQuota(accountId: string): Promise<Result<{ used: number; limit: number } | null>>;
  zimbraChangePassword(accountId: string, oldPass: string, newPass: string): Promise<Result<true>>;
  zimbraGetWB(accountId: string): Promise<Result<{ white: string[]; black: string[] }>>;
  zimbraSetWB(accountId: string, white: string[], black: string[]): Promise<Result<true>>;
  zimbraBlock(accountId: string, email: string): Promise<Result<true>>;
  tasksSync(tasks: Task[]): Promise<Result<{ tasks: Task[]; accounts: number }>>;
  tasksDelete(task: Task): Promise<Result<true>>;
  cryptoSummary(): Promise<Result<CryptoSummary>>;
  pgpGenerate(name: string, email: string, pass: string): Promise<Result<string>>;
  cryptoPick(kind: 'pgp' | 'p12' | 'cert'): Promise<Result<string>>;
  pgpImportFile(file: string, pass?: string): Promise<Result<{ own: number; pub: number }>>;
  pgpImportText(text: string, pass?: string): Promise<Result<{ own: number; pub: number }>>;
  smimeImportFile(file: string, pass: string): Promise<Result<CertSummary>>;
  certImportFile(file: string): Promise<Result<CertSummary>>;
  pgpExport(fingerprint: string): Promise<Result<string>>;
  cryptoRemove(kind: 'pgpOwn' | 'pgpContacts' | 'smimeOwn' | 'smimeContacts', id: string): Promise<Result<true>>;
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
  /** lokálny kalendár v programe (aj pre IMAP bez CalDAV) */
  calAddLocal(accountId: string, name: string, color?: string): Promise<Result<DavCollection>>;
  /** odber kalendára z adresy (ICS/webcal), len na čítanie, obnovuje sa každú hodinu */
  calSubscribe(accountId: string, url: string, name?: string, color?: string): Promise<Result<DavCollection>>;
  calRename(calendarId: string, name: string, color?: string): Promise<Result<true>>;
  calRemove(calendarId: string): Promise<Result<true>>;
  calSubInfo(calendarId: string): Promise<Result<{ url?: string; refreshedAt?: string; error?: string; kind: 'local' | 'sub' } | null>>;
  /** najbližšie udalosti z kalendára na adrese (ICS/webcal) – dlaždica v Súhrne */
  icsUpcoming(url: string, count: number): Promise<Result<CalEvent[]>>;
  /** text RSS/Atom kanála alebo webovej stránky (na nájdenie kanála) – dlaždica Správy v Súhrne */
  rssFetch(url: string): Promise<Result<{ text: string; url: string }>>;
  /** import súboru .ics do kalendára; null = zrušené */
  calImport(accountId: string, calendarId: string): Promise<Result<{ ok: number; failed: number } | null>>;
  /** otvorí Teams na webe s prihlásením Outlook konta – na vytvorenie odkazu na schôdzu */
  meetTeams(accountId: string): Promise<Result<true>>;
  meetTeamsClose(): Promise<Result<true>>;
  /** Jitsi schôdza v okne programu */
  meetJitsi(url: string): Promise<Result<true>>;
  /** text v schránke (zachytenie skopírovaného odkazu na Teams) */
  clipboardRead(): Promise<Result<string>>;
  /** naposledy použité adresy príjemcov (userData/recent.json, spoločné pre všetky okná) */
  recentList(): Promise<Result<{ name: string; email: string; count: number; last: string }[]>>;
  recentAdd(items: { name?: string; email: string }[], migrate?: boolean): Promise<Result<true>>;
  recentRemove(emails: string[]): Promise<Result<true>>;
  /** lokálny adresár v programe (userData/contacts.json) */
  localContacts(): Promise<Result<Contact[]>>;
  localPut(items: Contact[]): Promise<Result<Contact[]>>;
  localRemove(uids: string[]): Promise<Result<true>>;
  /** výber súboru .vcf/.csv na import kontaktov; text už dekódovaný; null = zrušené */
  contactsPickImport(): Promise<Result<{ filename: string; text: string } | null>>;
  /** uloženie exportu kontaktov; '' = zrušené */
  contactsSaveExport(name: string, text: string): Promise<Result<string>>;
  cardSave(accountId: string, bookId: string, contact: Contact): Promise<Result<Contact>>;
  cardDelete(accountId: string, href: string): Promise<Result<true>>;
  setKeywords(accountId: string, folder: string, uids: number[], add: string[], remove: string[]): Promise<Result<true>>;
  /** Gmail: štítky správy a ich zmena */
  gmailLabels(accountId: string, folder: string, uid: number): Promise<Result<string[]>>;
  gmailSetLabels(accountId: string, folder: string, uids: number[], add: string[], remove: string[]): Promise<Result<true>>;
  /** Prihlásenie cez Google (OAuth v prehliadači) pre konto s daným id (aj ešte neuložené) */
  googleSignIn(accountId: string, loginHint?: string): Promise<Result<GoogleGrant>>;
  googleCancel(): Promise<Result<true>>;
  googleSignOut(accountId: string): Promise<Result<true>>;
  googleGrant(accountId: string): Promise<Result<GoogleGrant | null>>;
  /** správy toho istého vlákna z cache (všetky priečinky konta) */
  thread(accountId: string, messageId: string): Promise<Result<MessageSummary[]>>;
  runActions(accountId: string, folder: string, uids: number[], actions: RuleAction[]): Promise<Result<true>>;
  snooze(accountId: string, folder: string, uid: number, messageId: string, subject: string, until: string): Promise<Result<true>>;
  /** spustí pravidlá na uložené správy priečinka; vráti počet zasiahnutých */
  runRules(accountId: string, folder: string): Promise<Result<number>>;
  /** fulltext v lokálnej cache */
  searchLocal(q: string, accountId?: string, folder?: string, field?: string): Promise<Result<MessageSummary[]>>;
  /** správy s vlajkou alebo pripnuté zo všetkých kont (z lokálnej cache) */
  flaggedAll(limit?: number): Promise<Result<MessageSummary[]>>;
  /** hľadanie na serveri (IMAP SEARCH) */
  searchServer(accountId: string, folder: string, q: string, field?: string): Promise<Result<MessageSummary[]>>;
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
  /** Záloha/prenos nastavení (bez hesiel). */
  settingsExport(parts: { accounts: boolean; contacts: boolean; calendars: boolean }): Promise<Result<string | null>>;
  settingsPick(): Promise<Result<BackupInfo | null>>;
  /** Použije zálohu a reštartuje program. */
  settingsApply(token: string, parts: { settings: boolean; accounts: boolean; contacts: boolean; calendars: boolean }): Promise<Result<void>>;
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
  /** Zámok programu heslom / PIN-om. */
  lockStatus(): Promise<Result<LockStatus>>;
  lockNow(): Promise<Result<void>>;
  lockUnlock(secret: string): Promise<Result<void>>;
  lockTouchId(): Promise<Result<void>>;
  /** zapnúť alebo zmeniť; vráti záchranný kód */
  lockSet(kind: 'pin' | 'password', secret: string, current?: string): Promise<Result<string>>;
  lockDisable(current: string): Promise<Result<void>>;
  lockOptions(o: Partial<LockOptions>): Promise<Result<void>>;
  /** obnova: vráti jednorazový token na nastavenie nového PIN-u/hesla */
  lockRecoverAccount(accountId: string, password: string): Promise<Result<string>>;
  lockRecoverCode(code: string): Promise<Result<string>>;
  lockReset(token: string, kind: 'pin' | 'password', secret: string): Promise<Result<string>>;
  // vložené weby
  /** Otvorí stránku v Edge (Windows) alebo v predvolenom prehliadači. */
  webOpenOutside(url: string): Promise<void>;
  onWebEvent(cb: (ev: WebEvent) => void): () => void;
  /** Vymaže prihlásenie webovej relácie (odhlásiť konto M365 / web). */
  webClear(partition: string): Promise<void>;
}

export interface LockOptions {
  /** minúty nečinnosti do zamknutia (0 = vypnuté) */
  idleMin: number;
  /** zamknúť pri zamknutí obrazovky / uspaní */
  onScreenLock: boolean;
  touchId: boolean;
  /** pri zamknutí upozornenia bez obsahu */
  hideNotif: boolean;
}
export interface LockStatus extends LockOptions {
  enabled: boolean;
  locked: boolean;
  kind: 'pin' | 'password';
  touchIdAvailable: boolean;
  /** sekundy do ďalšieho pokusu po zlých pokusoch */
  blockedFor: number;
  hasRecoveryCode: boolean;
  /** kontá, ktorých heslom sa dá overiť obnova */
  recoveryAccounts: { id: string; name: string; email: string }[];
}

export type AppEvent =
  | { type: 'lock-changed'; locked: boolean }
  | { type: 'settings-changed' }
  | { type: 'gpu-crash'; count: number }
  | { type: 'mail-changed'; accountId: string; folder: string; newCount: number }
  | { type: 'open-message'; accountId: string; folder: string; uid: number }
  | { type: 'outbox-changed' }
  | { type: 'outbox-queued'; item: OutboxItem }
  | { type: 'contacts-changed' }
  | { type: 'followup'; subject: string; to: string; accountId: string }
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

export interface CertSummary {
  emails: string[];
  subject: string;
  issuer: string;
  validTo: string;
}
export interface CryptoSummary {
  pgpOwn: Array<{ fingerprint: string; userIds: string[]; created: string }>;
  pgpContacts: Array<{ fingerprint: string; userIds: string[] }>;
  smimeOwn: CertSummary[];
  smimeContacts: CertSummary[];
}

export interface GalEntry { name: string; email: string; company?: string; title?: string; department?: string; phone?: string; mobile?: string; type?: 'account' | 'resource' | 'group' }
export interface FreeBusy { email: string; slots: Array<{ s: number; e: number; kind: 'b' | 't' | 'u' | 'n' }> }
export interface ZResource { name: string; email: string; kind: 'Location' | 'Equipment'; capacity?: number; location?: string; description?: string }
export interface ZShare { ownerEmail: string; ownerName?: string; ownerId: string; folderId: string; folderPath: string; view: string; rights: string; mounted: boolean }
export interface ServerOoo { enabled: boolean; reply: string; from?: string; until?: string; external?: string; externalEnabled?: boolean }
export interface ServerRule {
  name: string;
  active: boolean;
  match: 'all' | 'any';
  conds: Array<{ field: 'from' | 'to' | 'subject' | 'any'; contains: string }>;
  actions: Array<{ type: 'move' | 'read' | 'flag' | 'tag' | 'delete' | 'stop'; folder?: string; tag?: string }>;
  readOnly?: boolean;
  raw?: unknown;
}
export interface BriefcaseDoc { id: string; name: string; size: number; contentType: string; modified: string; folder: string }
