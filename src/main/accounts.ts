// Úložisko kont (JSON v priečinku používateľských dát) a hesiel (Electron safeStorage:
// DPAPI na Windows, Keychain na macOS, libsecret/kwallet na Linuxe).
import { app, safeStorage } from 'electron';
import fs from 'node:fs';
import path from 'node:path';
import type { AccountConfig, AccountPublic, AnyAccount } from '../core/account';

const dir = () => app.getPath('userData');
const accountsFile = () => path.join(dir(), 'accounts.json');
const secretsFile = () => path.join(dir(), 'secrets.json');

/** Heslá, ktoré sa nedajú bezpečne uložiť (Linux bez keyringu), držíme len v pamäti. */
const memorySecrets = new Map<string, string>();

function readJson<T>(file: string, fallback: T): T {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8')) as T;
  } catch {
    return fallback;
  }
}

function writeJson(file: string, data: unknown) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = `${file}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(data, null, 2), 'utf8');
  fs.renameSync(tmp, file);
}

export function canStoreSecrets(): boolean {
  if (!safeStorage.isEncryptionAvailable()) return false;
  // Na Linuxe bez keyringu Electron použije "basic_text" = v podstate nešifrované.
  const backend = (safeStorage as unknown as { getSelectedStorageBackend?: () => string }).getSelectedStorageBackend?.();
  return !(process.platform === 'linux' && backend === 'basic_text');
}

export function loadAccounts(): AnyAccount[] {
  // Staršie záznamy (krok 2) nemali kind = vždy IMAP.
  return readJson<AnyAccount[]>(accountsFile(), []).map((a) => ({ ...a, kind: a.kind ?? 'imap' }) as AnyAccount);
}

export function getAccount(id: string): AnyAccount {
  const a = loadAccounts().find((x) => x.id === id);
  if (!a) throw new Error('Konto neexistuje.');
  return a;
}

/** IMAP konto (pre poštové operácie). */
export function getImap(id: string): AccountConfig {
  const a = getAccount(id);
  if (a.kind !== 'imap') throw new Error('Toto konto nie je IMAP.');
  return a;
}

export function getPassword(id: string): string | undefined {
  if (memorySecrets.has(id)) return memorySecrets.get(id);
  const s = readJson<Record<string, string>>(secretsFile(), {});
  if (!s[id] || !canStoreSecrets()) return undefined;
  try {
    return safeStorage.decryptString(Buffer.from(s[id], 'base64'));
  } catch {
    return undefined;
  }
}

function setPassword(id: string, password: string) {
  if (!canStoreSecrets()) {
    memorySecrets.set(id, password);
    return;
  }
  const s = readJson<Record<string, string>>(secretsFile(), {});
  s[id] = safeStorage.encryptString(password).toString('base64');
  writeJson(secretsFile(), s);
}

export const toPublic = (a: AnyAccount): AccountPublic =>
  a.kind === 'imap' ? { ...a, hasPassword: !!getPassword(a.id) } : { ...a, hasPassword: false };

export function publicAccounts(): AccountPublic[] {
  return loadAccounts().map(toPublic);
}

export function saveAccount(acc: AnyAccount, password?: string): AccountPublic {
  const list = loadAccounts();
  const i = list.findIndex((x) => x.id === acc.id);
  if (i >= 0) list[i] = acc;
  else list.push(acc);
  writeJson(accountsFile(), list);
  if (password && acc.kind === 'imap') setPassword(acc.id, password);
  return toPublic(acc);
}

/** Nové poradie kont v páse kont (podľa zoznamu id). */
export function reorderAccounts(ids: string[]) {
  const list = loadAccounts();
  const rank = (id: string) => {
    const i = ids.indexOf(id);
    return i < 0 ? 999 : i;
  };
  writeJson(accountsFile(), [...list].sort((a, b) => rank(a.id) - rank(b.id)));
}

export function deleteAccount(id: string) {
  writeJson(accountsFile(), loadAccounts().filter((x) => x.id !== id));
  const s = readJson<Record<string, string>>(secretsFile(), {});
  delete s[id];
  writeJson(secretsFile(), s);
  memorySecrets.delete(id);
}
