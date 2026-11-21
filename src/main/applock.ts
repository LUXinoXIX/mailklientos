// Zámok programu heslom alebo PIN-om. Chráni rozhranie (nie súbory na disku – šifrovanie používateľ nechcel).
// Heslo sa neukladá – len scrypt odtlačok so soľou v userData/lock.json (mimo settings.json, takže nejde do exportu).
// Obnova bez straty čohokoľvek: heslo k niektorému e-mailovému kontu (overí sa prihlásením na IMAP) alebo záchranný kód.
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { app, BrowserWindow, powerMonitor, systemPreferences } from 'electron';
import type { AccountConfig } from '../core/account';
import type { LockOptions, LockStatus } from '../core/api';
import * as store from './accounts';
import { log } from './diagnostics';
import { UserError, withImap } from './mail';

interface Secret {
  salt: string;
  hash: string;
}
interface LockFile extends LockOptions {
  enabled: boolean;
  kind: 'pin' | 'password';
  secret?: Secret;
  recovery?: Secret;
}

const DEFAULTS: LockOptions = { idleMin: 15, onScreenLock: true, touchId: false, hideNotif: true };
const file = () => path.join(app.getPath('userData'), 'lock.json');
let cfg: LockFile | null = null;
let locked = false;
let fails = 0;
let blockedUntil = 0;
const resetTokens = new Map<string, number>();

function load(): LockFile {
  if (cfg) return cfg;
  try {
    cfg = { ...DEFAULTS, enabled: false, kind: 'pin', ...(JSON.parse(fs.readFileSync(file(), 'utf8')) as Partial<LockFile>) };
  } catch {
    cfg = { ...DEFAULTS, enabled: false, kind: 'pin' };
  }
  return cfg;
}
function save() {
  const tmp = file() + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(cfg, null, 1), 'utf8');
  fs.renameSync(tmp, file());
}

const scrypt = (text: string, salt: string) => crypto.scryptSync(text.normalize('NFC'), Buffer.from(salt, 'hex'), 32, { N: 16384, r: 8, p: 1 }).toString('hex');
function makeSecret(text: string): Secret {
  const salt = crypto.randomBytes(16).toString('hex');
  return { salt, hash: scrypt(text, salt) };
}
function matches(s: Secret | undefined, text: string) {
  if (!s) return false;
  const a = Buffer.from(scrypt(text, s.salt), 'hex');
  const b = Buffer.from(s.hash, 'hex');
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}
/** záchranný kód bez zameniteľných znakov (0/O, 1/I/L) */
function newRecoveryCode() {
  const abc = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  const bytes = crypto.randomBytes(16);
  const chars = [...bytes].map((b) => abc[b % abc.length]).join('');
  return chars.match(/.{4}/g)!.join('-');
}
const normCode = (c: string) => c.toUpperCase().replace(/[^A-Z0-9]/g, '');

function validate(kind: 'pin' | 'password', secret: string) {
  if (kind === 'pin' && !/^\d{4,12}$/.test(secret)) throw new UserError('PIN musí mať 4 až 12 číslic.');
  if (kind === 'password' && secret.length < 6) throw new UserError('Heslo musí mať aspoň 6 znakov.');
}

const touchIdAvailable = () => process.platform === 'darwin' && systemPreferences.canPromptTouchID();

export function status(): LockStatus {
  const c = load();
  return {
    enabled: c.enabled,
    locked: c.enabled && locked,
    kind: c.kind,
    idleMin: c.idleMin,
    onScreenLock: c.onScreenLock,
    touchId: c.touchId && touchIdAvailable(),
    touchIdAvailable: touchIdAvailable(),
    hideNotif: c.hideNotif,
    blockedFor: Math.max(0, Math.ceil((blockedUntil - Date.now()) / 1000)),
    hasRecoveryCode: !!c.recovery,
    recoveryAccounts: store
      .loadAccounts()
      .filter((a): a is AccountConfig => a.kind === 'imap' && a.auth !== 'oauth' && !!store.getPassword(a.id))
      .map((a) => ({ id: a.id, name: a.name, email: a.email })),
  };
}

export const isLocked = () => load().enabled && locked;
/** pri zamknutom programe upozornenia bez obsahu */
export const hideNotificationContent = () => isLocked() && load().hideNotif;

function broadcast() {
  const ev = { type: 'lock-changed', locked: isLocked() };
  for (const w of BrowserWindow.getAllWindows()) if (!w.isDestroyed()) w.webContents.send('app:event', ev);
}

export function lockNow() {
  if (!load().enabled || locked) return;
  locked = true;
  log('zámok: program zamknutý');
  broadcast();
}

function unlocked() {
  locked = false;
  fails = 0;
  blockedUntil = 0;
  broadcast();
}

export function unlock(secret: string): void {
  const c = load();
  if (!c.enabled) return unlocked();
  if (Date.now() < blockedUntil) throw new UserError(`Príliš veľa pokusov. Skús znova o ${Math.ceil((blockedUntil - Date.now()) / 1000)} s.`);
  if (matches(c.secret, secret)) return unlocked();
  fails++;
  // po 5 zlých pokusoch čakanie 30 s, ďalej sa predlžuje (max. 5 min)
  if (fails >= 5) blockedUntil = Date.now() + Math.min(300, 30 * 2 ** (fails - 5)) * 1000;
  throw new UserError(c.kind === 'pin' ? 'Nesprávny PIN.' : 'Nesprávne heslo.');
}

export async function unlockTouchId(): Promise<void> {
  const c = load();
  if (!c.touchId || !touchIdAvailable()) throw new UserError('Touch ID nie je zapnuté.');
  try {
    await systemPreferences.promptTouchID('odomknúť Universum eMail');
  } catch {
    throw new UserError('Touch ID neoverilo odtlačok.');
  }
  unlocked();
}

/** Zapnutie alebo zmena PIN-u/hesla. Pri zmene treba súčasné. Vráti nový záchranný kód. */
export function setSecret(kind: 'pin' | 'password', secret: string, current?: string): string {
  const c = load();
  if (c.enabled && !matches(c.secret, current ?? '')) throw new UserError('Súčasný PIN alebo heslo nesedí.');
  validate(kind, secret);
  const code = newRecoveryCode();
  cfg = { ...c, enabled: true, kind, secret: makeSecret(secret), recovery: makeSecret(normCode(code)) };
  save();
  locked = false;
  broadcast();
  return code;
}

export function disable(current: string) {
  const c = load();
  if (c.enabled && !matches(c.secret, current)) throw new UserError('Súčasný PIN alebo heslo nesedí.');
  cfg = { ...c, enabled: false, secret: undefined, recovery: undefined };
  save();
  unlocked();
}

export function setOptions(o: Partial<LockOptions>) {
  const c = load();
  cfg = {
    ...c,
    idleMin: o.idleMin ?? c.idleMin,
    onScreenLock: o.onScreenLock ?? c.onScreenLock,
    touchId: o.touchId ?? c.touchId,
    hideNotif: o.hideNotif ?? c.hideNotif,
  };
  save();
  broadcast();
}

// --- obnova bez straty údajov -------------------------------------------------

function issueResetToken() {
  const t = crypto.randomUUID();
  resetTokens.set(t, Date.now() + 10 * 60_000);
  return t;
}

/** Overí heslo k e-mailovému kontu prihlásením na IMAP server. */
export async function recoverWithAccount(accountId: string, password: string): Promise<string> {
  if (Date.now() < blockedUntil) throw new UserError(`Príliš veľa pokusov. Skús znova o ${Math.ceil((blockedUntil - Date.now()) / 1000)} s.`);
  const acc = store.loadAccounts().find((a) => a.id === accountId);
  if (!acc || acc.kind !== 'imap') throw new UserError('Konto sa nenašlo.');
  try {
    await withImap(acc, async () => true, password);
  } catch (e) {
    fails++;
    if (fails >= 5) blockedUntil = Date.now() + Math.min(300, 30 * 2 ** (fails - 5)) * 1000;
    log('zámok: obnova cez konto zlyhala', String(e));
    throw new UserError('Server heslo neprijal (alebo nie je dostupný).');
  }
  return issueResetToken();
}

export function recoverWithCode(code: string): string {
  if (Date.now() < blockedUntil) throw new UserError(`Príliš veľa pokusov. Skús znova o ${Math.ceil((blockedUntil - Date.now()) / 1000)} s.`);
  if (!matches(load().recovery, normCode(code))) {
    fails++;
    if (fails >= 5) blockedUntil = Date.now() + Math.min(300, 30 * 2 ** (fails - 5)) * 1000;
    throw new UserError('Záchranný kód nesedí.');
  }
  return issueResetToken();
}

/** Po úspešnej obnove nastaví nový PIN/heslo (bez súčasného). Vráti nový záchranný kód. */
export function resetWith(token: string, kind: 'pin' | 'password', secret: string): string {
  const until = resetTokens.get(token);
  if (!until || until < Date.now()) throw new UserError('Obnova vypršala – začni znova.');
  validate(kind, secret);
  resetTokens.delete(token);
  const code = newRecoveryCode();
  cfg = { ...load(), enabled: true, kind, secret: makeSecret(secret), recovery: makeSecret(normCode(code)) };
  save();
  fails = 0;
  blockedUntil = 0;
  // odomkne sa až po zobrazení nového záchranného kódu (rozhranie zavolá unlock s novým PIN-om/heslom)
  return code;
}

// --- automatické zamykanie ----------------------------------------------------

export function startAutoLock() {
  locked = load().enabled; // pri spustení zamknuté
  const onScreen = () => load().onScreenLock && lockNow();
  powerMonitor.on('lock-screen', onScreen);
  powerMonitor.on('suspend', onScreen);
  setInterval(() => {
    const c = load();
    if (!c.enabled || locked || !c.idleMin) return;
    try {
      if (powerMonitor.getSystemIdleTime() >= c.idleMin * 60) lockNow();
    } catch {
      /* */
    }
  }, 20_000);
}
