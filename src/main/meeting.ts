// Online schôdze: okno Teams (prihlásenie konta Outlook) na vytvorenie odkazu a Jitsi schôdza priamo v programe.
import { BrowserWindow, clipboard, shell } from 'electron';
import { loadAccounts } from './accounts';
import { getMainWindow } from './mainwin';
import { UserError } from './mail';
import { prepareSession } from './webviews';

let teamsWin: BrowserWindow | null = null;

/** Otvorí Teams na webe s prihlásením zvoleného Outlook konta (rovnaká relácia ako vložený Outlook). */
export function openTeamsForLink(accountId: string) {
  const acc = loadAccounts().find((a) => a.id === accountId && a.kind === 'm365');
  if (!acc) throw new UserError('Konto Outlook pre Teams sa nenašlo. Nastav ho v udalosti → Online schôdza.');
  const personal = 'host' in acc && acc.host === 'live';
  const url = personal ? 'https://teams.live.com/v2/' : 'https://teams.microsoft.com/v2/';
  if (teamsWin && !teamsWin.isDestroyed()) {
    teamsWin.focus();
    return true as const;
  }
  const ses = prepareSession(`persist:acc-${acc.id}`);
  const parent = getMainWindow() ?? undefined;
  teamsWin = new BrowserWindow({
    width: 1150,
    height: 780,
    parent,
    title: `Teams – ${acc.name || acc.email} · vytvor odkaz na schôdzu`,
    autoHideMenuBar: true,
    webPreferences: { session: ses, sandbox: true, contextIsolation: true },
  });
  teamsWin.webContents.setWindowOpenHandler(({ url: u }) => {
    if (/^https:\/\/(login\.(microsoftonline|live)\.com|teams\.(microsoft|live)\.com)\//i.test(u)) return { action: 'allow' };
    void shell.openExternal(u);
    return { action: 'deny' };
  });
  teamsWin.on('closed', () => (teamsWin = null));
  void teamsWin.loadURL(url);
  return true as const;
}

export function closeTeamsWindow() {
  if (teamsWin && !teamsWin.isDestroyed()) teamsWin.close();
  return true as const;
}

export async function readClipboardText() {
  return (await Promise.resolve(clipboard.readText())).slice(0, 4000);
}

/** Jitsi schôdza vo vlastnom okne programu (mikrofón, kamera a zdieľanie obrazovky povolené). */
export function openJitsi(url: string) {
  if (!/^https:\/\//i.test(url)) throw new UserError('Neplatný odkaz na schôdzu.');
  const ses = prepareSession('persist:meet');
  const w = new BrowserWindow({ width: 1200, height: 800, title: 'Jitsi Meet', autoHideMenuBar: true, webPreferences: { session: ses, sandbox: true, contextIsolation: true } });
  w.webContents.setWindowOpenHandler(({ url: u }) => {
    // prihlásenie moderátora na meet.jit.si (Google / GitHub / Facebook) v malom okne
    if (/^https:\/\/(accounts\.google\.com|github\.com|www\.facebook\.com|meet\.jit\.si|8x8\.vc)\//i.test(u)) return { action: 'allow' };
    void shell.openExternal(u);
    return { action: 'deny' };
  });
  void w.loadURL(url);
  return true as const;
}
