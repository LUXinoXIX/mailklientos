import { app, BrowserWindow, Menu, powerMonitor, session, shell } from 'electron';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { registerIpc } from './main/ipc';
import { closeAllConnections } from './main/mail';
import { startSnoozeTimer } from './main/snooze';
import { reconnectWatchers, stopAllWatchers, syncWatchers } from './main/watch';
import { attachWindow } from './main/webviews';
import { applyRenderingSwitches, log, watchProcesses } from './main/diagnostics';
import { handleSquirrelEvent } from './main/squirrel';

// Dymový test v CI (MAILKLIENT_SMOKE=m365): dočasné dáta s ukážkovým kontom M365 a otvoreným pravým panelom.
const SMOKE_MODE = process.env.MAILKLIENT_SMOKE ?? '';
const SMOKE = SMOKE_MODE === 'm365' || SMOKE_MODE === 'compose';
if (SMOKE) {
  const dir = path.join(os.tmpdir(), `mailklient-smoke-${SMOKE_MODE}`);
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(
    path.join(dir, 'accounts.json'),
    JSON.stringify(
      SMOKE_MODE === 'm365' ? [{ id: 'smoke', kind: 'm365', name: 'Microsoft 365', email: 'test@contoso.com', color: '#1d4ed8', host: 'office' }] : [],
    ),
  );
  app.setPath('userData', dir);
}

// Prepínače vykresľovania (proti blikaniu) musia byť nastavené pred štartom.
applyRenderingSwitches();

// Windows inštalátor Squirrel: pri inštalácii len vytvoriť zástupcov a skončiť.
const squirrelEvent = handleSquirrelEvent();

const createWindow = () => {
  const win = new BrowserWindow({
    autoHideMenuBar: true,
    backgroundColor: '#f5f5f5',
    webPreferences: {
      backgroundThrottling: false,
      webviewTag: true,
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
    width: 1440,
    height: 900,
    minWidth: 960,
    minHeight: 600,
    title: 'MailKlient',
    show: false,
  });
  win.once('ready-to-show', () => win.show());
  win.on('unresponsive', () => log('okno neodpovedá'));
  attachWindow(win);

  // Odkazy z e-mailov a webov otvárať v systémovom prehliadači, nie v aplikácii.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^(https?|mailto):/i.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (e, url) => {
    if (MAIN_WINDOW_VITE_DEV_SERVER_URL && url.startsWith(MAIN_WINDOW_VITE_DEV_SERVER_URL)) return;
    e.preventDefault();
    if (/^(https?|mailto):/i.test(url)) shell.openExternal(url);
  });

  if (MAIN_WINDOW_VITE_DEV_SERVER_URL) {
    win.loadURL(MAIN_WINDOW_VITE_DEV_SERVER_URL + (SMOKE ? `?smoke=${SMOKE_MODE}` : ''));
  } else {
    win.loadFile(path.join(__dirname, `../renderer/${MAIN_WINDOW_VITE_NAME}/index.html`), SMOKE ? { query: { smoke: SMOKE_MODE } } : undefined);
  }
};

if (!squirrelEvent && !app.requestSingleInstanceLock()) app.quit();

// Druhé spustenie: zobraziť už otvorené okno.
app.on('second-instance', () => {
  const win = BrowserWindow.getAllWindows()[0];
  if (!win) return;
  if (win.isMinimized()) win.restore();
  win.focus();
});

app.whenReady().then(() => {
  if (squirrelEvent) return;
  watchProcesses();
  // žiadne predvolené menu Electronu (okno má vlastný ribbon)
  Menu.setApplicationMenu(null);
  // kontrola pravopisu: slovenčina + angličtina (Windows a macOS používajú jazyky systému)
  if (process.platform === 'linux') {
    try {
      session.defaultSession.setSpellCheckerLanguages(['sk', 'en-US']);
    } catch {
      /* slovník nie je k dispozícii */
    }
  }
  registerIpc();
  createWindow();
  // Windows: ID aplikácie pre upozornenia (inak sa nezobrazia pri Squirrel inštalácii)
  if (process.platform === 'win32') app.setAppUserModelId('com.squirrel.mailklient.MailKlient');
  syncWatchers();
  startSnoozeTimer();
  powerMonitor.on('resume', () => reconnectWatchers());
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('before-quit', () => {
  stopAllWatchers();
  closeAllConnections();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
