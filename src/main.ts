import { app, BrowserWindow, Menu, powerMonitor, session, shell } from 'electron';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { registerIpc, setComposeWindowFactory } from './main/ipc';
import { startOutbox } from './main/outbox';
import { startUpdater } from './main/updater';
import { showSplash } from './main/splash';
import { startAutoLock } from './main/applock';
import { isQuitting, markQuitting, setupTray, startFollowUps } from './main/extras';
import { loadSettings } from './main/settings';
import { closeAllConnections } from './main/mail';
import { startSnoozeTimer } from './main/snooze';
import { reconnectWatchers, stopAllWatchers, syncWatchers } from './main/watch';
import { attachWindow } from './main/webviews';
import { applyRenderingSwitches, log, watchProcesses } from './main/diagnostics';
import { handleSquirrelEvent } from './main/squirrel';
import { getMainWindow, setMainWindow } from './main/mainwin';
import { attachContextMenu } from './main/contextmenu';

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

const boundsFile = () => path.join(app.getPath('userData'), 'msgwin.json');
function savedBounds(): Electron.Rectangle | null {
  try {
    return JSON.parse(fs.readFileSync(boundsFile(), 'utf8')) as Electron.Rectangle;
  } catch {
    return null;
  }
}

const createWindow = (query?: Record<string, string>) => {
  const compose = !!query?.compose;
  const view = !!query?.view;
  /** ďalšie hlavné okno (ozubené koliesko → Nové okno): celý program, ale nie je „hlavné“ (zatvorí sa normálne) */
  const second = !!query?.second;
  const mainWin = !compose && !view && !second;
  const main = getMainWindow();
  const sb = second && main ? main.getBounds() : null;
  // okno správy: posledný rozmer, každé ďalšie posunuté, aby sa neprekrývali presne
  const vb = view ? savedBounds() : null;
  const shift = view ? (BrowserWindow.getAllWindows().length % 6) * 28 : 0;
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
    width: view ? vb?.width ?? 900 : compose ? 980 : 1440,
    height: view ? vb?.height ?? 820 : compose ? 820 : 900,
    ...(view && vb ? { x: vb.x + shift, y: vb.y + shift } : {}),
    ...(sb ? { x: sb.x + 40, y: sb.y + 40, width: sb.width, height: sb.height } : {}),
    minWidth: view ? 480 : compose ? 640 : 960,
    minHeight: view ? 400 : 600,
    title: view ? 'Správa – Universum eMail' : compose ? 'Nová správa – Universum eMail' : 'Universum eMail',
    show: false,
  });
  // štartovací banner len pri hlavnom okne (nie pri okne písania a dymovom teste)
  const closeSplash = mainWin && !SMOKE && BrowserWindow.getAllWindows().length <= 1 ? showSplash() : null;
  win.once('ready-to-show', () => {
    if (closeSplash) closeSplash(() => win.show());
    else win.show();
  });
  win.on('unresponsive', () => log('okno neodpovedá'));
  // zatvorenie hlavného okna: aplikácia beží ďalej pri hodinách (pošta, upozornenia, odloženie, plánované odoslanie)
  if (view) {
    win.on('close', () => {
      try {
        fs.writeFileSync(boundsFile(), JSON.stringify(win.getNormalBounds()), 'utf8');
      } catch {
        /* nevadí */
      }
    });
  }
  if (mainWin) {
    win.on('close', (e) => {
      if (isQuitting() || SMOKE || loadSettings().runInBackground === false) return;
      e.preventDefault();
      win.hide();
    });
  }
  attachWindow(win);
  attachContextMenu(win.webContents);
  if (mainWin) setMainWindow(win);

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
    const q = { ...(SMOKE ? { smoke: SMOKE_MODE } : {}), ...query };
    win.loadFile(path.join(__dirname, `../renderer/${MAIN_WINDOW_VITE_NAME}/index.html`), Object.keys(q).length ? { query: q } : undefined);
  }
};
setComposeWindowFactory((q) => createWindow(q));

if (!squirrelEvent && !app.requestSingleInstanceLock()) app.quit();

// Druhé spustenie: zobraziť už otvorené okno.
app.on('second-instance', () => {
  const win = getMainWindow();
  if (!win) return;
  if (win.isMinimized()) win.restore();
  win.show();
  win.focus();
});

app.whenReady().then(() => {
  if (squirrelEvent) return;
  watchProcesses();
  // Windows/Linux: žiadne predvolené menu Electronu (okno má vlastný ribbon).
  // macOS: horné menu je nutné – bez neho nefungujú Cmd+C/V/X/A/Z, Cmd+Q, Cmd+W, Cmd+M.
  if (process.platform === 'darwin') {
    app.setAboutPanelOptions({ applicationName: 'Universum eMail', applicationVersion: app.getVersion(), copyright: 'Universum' });
    Menu.setApplicationMenu(
      Menu.buildFromTemplate([
        {
          label: 'Universum eMail',
          submenu: [
            { role: 'about', label: 'O programe Universum eMail' },
            { type: 'separator' },
            { role: 'hide', label: 'Skryť Universum eMail' },
            { role: 'hideOthers', label: 'Skryť ostatné' },
            { role: 'unhide', label: 'Zobraziť všetko' },
            { type: 'separator' },
            { role: 'quit', label: 'Ukončiť Universum eMail' },
          ],
        },
        {
          label: 'Upraviť',
          submenu: [
            { role: 'undo', label: 'Späť' },
            { role: 'redo', label: 'Znova' },
            { type: 'separator' },
            { role: 'cut', label: 'Vystrihnúť' },
            { role: 'copy', label: 'Kopírovať' },
            { role: 'paste', label: 'Prilepiť' },
            { role: 'pasteAndMatchStyle', label: 'Prilepiť ako čistý text' },
            { role: 'selectAll', label: 'Vybrať všetko' },
          ],
        },
        {
          label: 'Zobraziť',
          submenu: [
            { role: 'resetZoom', label: 'Pôvodná veľkosť' },
            { role: 'zoomIn', label: 'Zväčšiť' },
            { role: 'zoomOut', label: 'Zmenšiť' },
            { type: 'separator' },
            { role: 'togglefullscreen', label: 'Celá obrazovka' },
          ],
        },
        {
          label: 'Okno',
          submenu: [
            { role: 'minimize', label: 'Minimalizovať' },
            { role: 'zoom', label: 'Zväčšiť okno' },
            { role: 'close', label: 'Zavrieť okno' },
            { label: 'Nové okno', accelerator: 'Cmd+Shift+N', click: () => createWindow({ second: '1' }) },
            { type: 'separator' },
            { role: 'front', label: 'Všetko dopredu' },
          ],
        },
      ]),
    );
  } else {
    Menu.setApplicationMenu(null);
  }
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
  startOutbox();
  startFollowUps();
  startAutoLock();
  setupTray(() => {
    const w = getMainWindow();
    if (!w || w.webContents.getURL().includes('view=1') || w.webContents.getURL().includes('compose=1')) return createWindow();
    if (w.isMinimized()) w.restore();
    w.show();
    w.focus();
  });
  if (!squirrelEvent) startUpdater();
  powerMonitor.on('resume', () => reconnectWatchers());
  // macOS: klik na ikonu v Docku znova ukáže hlavné okno (aj keď beží skryté na pozadí)
  app.on('activate', () => {
    const w = getMainWindow();
    if (w && !w.isDestroyed()) {
      if (w.isMinimized()) w.restore();
      w.show();
      w.focus();
    } else if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('before-quit', () => {
  markQuitting();
  stopAllWatchers();
  closeAllConnections();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
