import { app, BrowserWindow, shell } from 'electron';
import path from 'node:path';
import { registerIpc } from './main/ipc';
import { handleSquirrelEvent } from './main/squirrel';

// Windows inštalátor Squirrel: pri inštalácii len vytvoriť zástupcov a skončiť.
const squirrelEvent = handleSquirrelEvent();

const createWindow = () => {
  const win = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 960,
    minHeight: 600,
    title: 'Mail klient',
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
