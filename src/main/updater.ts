// Automatické aktualizácie (Windows, Squirrel): MailKlient si pri štarte a každé 4 hodiny skontroluje
// release „latest-build“ na GitHube, novú verziu stiahne na pozadí a ponúkne reštart.
import { app, autoUpdater, BrowserWindow } from 'electron';
import { log } from './diagnostics';

const FEED = 'https://github.com/LUXinoXIX/mailklientos/releases/download/latest-build';
let ready = false;
let started = false;

const emit = (ev: unknown) => {
  for (const w of BrowserWindow.getAllWindows()) if (!w.isDestroyed()) w.webContents.send('app:event', ev);
};

export function startUpdater() {
  if (started || process.platform !== 'win32' || !app.isPackaged) return;
  started = true;
  try {
    autoUpdater.setFeedURL({ url: FEED });
  } catch (e) {
    log('aktualizácie: nastavenie zlyhalo', String(e));
    return;
  }
  autoUpdater.on('error', (e) => log('aktualizácie: chyba', String(e?.message ?? e)));
  autoUpdater.on('update-available', () => {
    log('aktualizácie: sťahujem novú verziu');
    emit({ type: 'update', state: 'downloading' });
  });
  autoUpdater.on('update-not-available', () => emit({ type: 'update', state: 'none' }));
  autoUpdater.on('update-downloaded', (_e, notes, name) => {
    ready = true;
    log('aktualizácie: nová verzia pripravená', String(name ?? ''));
    emit({ type: 'update', state: 'ready', version: String(name ?? '') });
  });
  // po štarte chvíľu počkať (Squirrel pri prvom spustení ešte dokončuje inštaláciu)
  setTimeout(() => check(), 60_000);
  setInterval(() => check(), 4 * 3600_000);
}

export function check() {
  if (!started) return false;
  try {
    autoUpdater.checkForUpdates();
    return true;
  } catch (e) {
    log('aktualizácie: kontrola zlyhala', String(e));
    return false;
  }
}

export function installNow() {
  if (ready) autoUpdater.quitAndInstall();
}

export const isReady = () => ready;
