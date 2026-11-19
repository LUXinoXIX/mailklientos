// Automatické aktualizácie (Windows, Squirrel): MailKlient si pri štarte a každé 4 hodiny skontroluje
// release „latest-build“ na GitHube, novú verziu stiahne na pozadí a ponúkne reštart.
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { app, autoUpdater, BrowserWindow, net } from 'electron';
import { log } from './diagnostics';

const FEED = 'https://github.com/LUXinoXIX/mailklientos/releases/download/latest-build';
let ready = false;
let started = false;

const emit = (ev: unknown) => {
  for (const w of BrowserWindow.getAllWindows()) if (!w.isDestroyed()) w.webContents.send('app:event', ev);
};

/** Inštalácia klasickým inštalátorom (Inno Setup): vedľa programu je unins000.exe, nie je tam Squirrel Update.exe. */
const innoMode = () =>
  process.platform === 'win32' &&
  fs.existsSync(path.join(path.dirname(process.execPath), 'unins000.exe')) &&
  !fs.existsSync(path.resolve(path.dirname(process.execPath), '..', 'Update.exe'));
let innoFile = '';

const newer = (a: string, b: string) => {
  const x = a.split('.').map(Number);
  const y = b.split('.').map(Number);
  for (let i = 0; i < 3; i++) if ((x[i] ?? 0) !== (y[i] ?? 0)) return (x[i] ?? 0) > (y[i] ?? 0);
  return false;
};

/** Inno: verziu zistí zo súboru RELEASES (ten istý release), nový UniversumEmail-Setup.exe stiahne na pozadí. */
async function checkInno() {
  try {
    const r = await net.fetch(`${FEED}/RELEASES`);
    if (!r.ok) throw new Error(`RELEASES ${r.status}`);
    const m = /mailklient-(\d+\.\d+\.\d+)-full\.nupkg/i.exec(await r.text());
    const latest = m?.[1];
    if (!latest || !newer(latest, app.getVersion())) return emit({ type: 'update', state: 'none' });
    if (ready) return;
    emit({ type: 'update', state: 'downloading' });
    log('aktualizácie (Inno): sťahujem', latest);
    const f = path.join(os.tmpdir(), `UniversumEmail-Setup-${latest}.exe`);
    const d = await net.fetch(`${FEED}/UniversumEmail-Setup.exe`);
    if (!d.ok) throw new Error(`inštalátor ${d.status}`);
    fs.writeFileSync(f, Buffer.from(await d.arrayBuffer()));
    innoFile = f;
    ready = true;
    emit({ type: 'update', state: 'ready', version: latest });
  } catch (e) {
    log('aktualizácie (Inno): chyba', String((e as Error)?.message ?? e));
  }
}

export function startUpdater() {
  if (started || process.platform !== 'win32' || !app.isPackaged) return;
  started = true;
  if (innoMode()) {
    setTimeout(() => void checkInno(), 60_000);
    setInterval(() => void checkInno(), 4 * 3600_000);
    return;
  }
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
  if (innoMode()) {
    void checkInno();
    return true;
  }
  try {
    autoUpdater.checkForUpdates();
    return true;
  } catch (e) {
    log('aktualizácie: kontrola zlyhala', String(e));
    return false;
  }
}

export function installNow() {
  if (!ready) return;
  if (innoFile) {
    // tichá inštalácia na to isté miesto (pre všetkých alebo len pre mňa podľa doterajšej inštalácie), potom sa program spustí sám
    const allUsers = /\\program files/i.test(process.execPath);
    spawn(innoFile, ['/SILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/CLOSEAPPLICATIONS', allUsers ? '/ALLUSERS' : '/CURRENTUSER'], { detached: true, stdio: 'ignore' }).unref();
    setTimeout(() => app.exit(0), 500);
    return;
  }
  autoUpdater.quitAndInstall();
}

export const isReady = () => ready;
