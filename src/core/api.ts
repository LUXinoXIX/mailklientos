// Zmluva medzi rendererom a hlavným procesom (cez preload). Len typy.
import type { AccountConfig, AccountPublic, ServerConfig } from './account';
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
}

export interface SendInput {
  to: string;
  cc?: string;
  subject: string;
  text: string;
  inReplyTo?: { folder: string; uid: number };
}

export interface AccountTestResult {
  imap: ProbeResult;
  smtp: ProbeResult;
  login?: { ok: boolean; message: string; hint?: string };
}

export interface MailApi {
  platform: string;
  listAccounts(): Promise<Result<{ accounts: AccountPublic[]; canStoreSecrets: boolean }>>;
  saveAccount(acc: AccountConfig, password?: string): Promise<Result<AccountPublic>>;
  deleteAccount(id: string): Promise<Result<true>>;
  probe(cfg: ServerConfig, proto: Protocol): Promise<Result<ProbeResult>>;
  autodetect(host: string, proto: Protocol): Promise<Result<{ config?: ServerConfig; result: ProbeResult; tried: string[] }>>;
  testAccount(acc: AccountConfig, password?: string): Promise<Result<AccountTestResult>>;
  listFolders(accountId: string): Promise<Result<FolderInfo[]>>;
  listMessages(accountId: string, folder: string, limit?: number): Promise<Result<MessageSummary[]>>;
  getMessage(accountId: string, folder: string, uid: number): Promise<Result<MessageFull>>;
  send(accountId: string, msg: SendInput): Promise<Result<{ messageId: string; savedToSent: boolean }>>;
  openExternal(url: string): Promise<void>;
}
