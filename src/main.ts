import { app, BrowserWindow, shell } from 'electron';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { registerIpc } from './main/ipc';
import { attachWindow } from './main/webviews';
import { handleSquirrelEvent } from './main/squirrel';

// Dymový test v CI (MAILKLIENT_SMOKE=m365): dočasné dáta s ukážkovým kontom M365 a otvoreným pravým panelom.
const SMOKE = process.env.MAILKLIENT_SMOKE === 'm365';
if (SMOKE) {
  const dir = path.join(os.tmpdir(), 'mailklient-smoke');
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(
    path.join(dir, 'accounts.json'),
    JSON.stringify([{ id: 'smoke', kind: 'm365', name: 'Microsoft 365', email: 'test@contoso.com', color: '#1d4ed8', host: 'office' }]),
  );
  app.setPath('userData', dir);
}

// Windows inštalátor Squirrel: pri inštalácii len vytvoriť zástupcov a skončiť.
const squirrelEvent = handleSquirrelEvent();

const createWindow = () => {
  const win = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 960,
    minHeight: 600,
    title: 'MailKlient',
    autoHideMenuBar: true,
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });
  win.once('ready-to-show', () => win.show());
  attachWindow(win);
  if (SMOKE)
    win.webContents.once('did-finish-load', () => {
      // počkať, kým React zapíše svoje predvoľby, potom prepísať a načítať znova
      setTimeout(() => void win.webContents.executeJavaScript(
        "localStorage.setItem('rightOpen','true');localStorage.setItem('activeAccount',JSON.stringify('smoke'));location.reload();",
      ), 3000);
    });

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
    win.loadURL(MAIN_WINDOW_VITE_DEV_SERVER_URL);
  } else {
    win.loadFile(path.join(__dirname, `../renderer/${MAIN_WINDOW_VITE_NAME}/index.html`));
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
  registerIpc();
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
