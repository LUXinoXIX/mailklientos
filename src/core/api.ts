// Zmluva medzi rendererom a hlavným procesom (cez preload). Len typy.
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
}

export interface MessageFull extends MessageSummary {
  to: string;
  cc: string;
  html: string | null;
  text: string;
  attachments: { filename: string; size: number; contentType: string }[];
  messageId?: string;
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
  outlookSignatures(): Promise<Result<Array<{ name: string; text: string }>>>;
  pickProgram(): Promise<Result<{ target: string; name: string } | null>>;
  runLauncher(item: LauncherItem): Promise<Result<string>>;
  programIcon(target: string): Promise<Result<string | null>>;
  // vložené weby
  webShow(key: string, url: string, partition: string, bounds: WebBounds): Promise<void>;
  webBounds(key: string, bounds: WebBounds): Promise<void>;
  webHide(key: string): Promise<void>;
  webHideAll(): Promise<void>;
  webNavigate(key: string, url: string): Promise<void>;
  webReload(key: string): Promise<void>;
  webBack(key: string): Promise<void>;
  /** Zavrie kartu (zničí web). */
  webClose(key: string): Promise<void>;
  onWebEvent(cb: (ev: WebEvent) => void): () => void;
  // stabilita vykresľovania
  gpuDisabled(): Promise<Result<boolean>>;
  /** Uloží nastavenie a reštartuje aplikáciu. */
  setGpuDisabled(off: boolean): Promise<void>;
  openLog(): Promise<void>;
  logError(msg: string): Promise<void>;
  onAppEvent(cb: (ev: { type: 'gpu-crash'; count: number }) => void): () => void;
  /** Otvorí aktuálnu stránku webu v Edge (Windows) alebo v predvolenom prehliadači. */
  webOpenOutside(key: string, fallbackUrl: string): Promise<void>;
  /** Vymaže prihlásenie webovej relácie (odhlásiť konto M365 / web). */
  webClear(partition: string): Promise<void>;
}
