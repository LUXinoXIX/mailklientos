// Nastavenia aplikácie (weby a spúšťače v pravom paneli) v JSON súbore v používateľských dátach.
// Spúšťanie programov a načítanie ich ikon.
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { app, BrowserWindow, dialog, shell } from 'electron';
import { DEFAULT_SETTINGS, type AppSettings, type LauncherItem } from '../core/settings';

const file = () => path.join(app.getPath('userData'), 'settings.json');

export function loadSettings(): AppSettings {
  try {
    const s = JSON.parse(fs.readFileSync(file(), 'utf8')) as Partial<AppSettings>;
    return {
      webs: Array.isArray(s.webs) ? s.webs : DEFAULT_SETTINGS.webs,
      launchers: Array.isArray(s.launchers) ? s.launchers : DEFAULT_SETTINGS.launchers,
      signatures: Array.isArray(s.signatures) ? s.signatures : [],
      signatureDefaults: s.signatureDefaults && typeof s.signatureDefaults === 'object' ? s.signatureDefaults : {},
    };
  } catch {
    return structuredClone(DEFAULT_SETTINGS);
  }
}

export function saveSettings(s: AppSettings): AppSettings {
  fs.mkdirSync(path.dirname(file()), { recursive: true });
  const tmp = `${file()}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(s, null, 2), 'utf8');
  fs.renameSync(tmp, file());
  return s;
}

const isUrlScheme = (t: string) => /^[a-z][a-z0-9+.-]+:/i.test(t) && !/^[a-z]:[\\/]/i.test(t);

/** Spustí program alebo otvorí URL schému. Vráti text chyby alebo prázdny reťazec. */
export async function runLauncher(l: LauncherItem): Promise<string> {
  const t = l.target.trim();
  if (isUrlScheme(t)) {
    await shell.openExternal(t);
    return '';
  }
  if (!fs.existsSync(t)) return `Program sa nenašiel: ${t}`;
  if (l.args?.length) {
    const child = spawn(t, l.args, { detached: true, stdio: 'ignore', cwd: path.dirname(t) });
    child.on('error', () => undefined);
    child.unref();
    return '';
  }
  return shell.openPath(t); // '' = v poriadku, inak text chyby
}

/** Dialóg na výber programu. */
export async function pickProgram(win: BrowserWindow | null): Promise<{ target: string; name: string } | null> {
  const filters =
    process.platform === 'win32'
      ? [{ name: 'Programy', extensions: ['exe', 'lnk', 'bat', 'cmd', 'msc'] }, { name: 'Všetky súbory', extensions: ['*'] }]
      : [{ name: 'Všetky súbory', extensions: ['*'] }];
  const opts: Electron.OpenDialogOptions = {
    title: 'Vyber program',
    properties: process.platform === 'darwin' ? ['openFile', 'treatPackageAsDirectory'] : ['openFile'],
    filters,
    defaultPath: process.platform === 'win32' ? 'C:\\Program Files' : process.platform === 'darwin' ? '/Applications' : '/usr/bin',
  };
  const r = win ? await dialog.showOpenDialog(win, opts) : await dialog.showOpenDialog(opts);
  if (r.canceled || !r.filePaths[0]) return null;
  let target = r.filePaths[0];
  // Zástupca .lnk: použiť cieľ, ak sa dá prečítať.
  if (process.platform === 'win32' && target.toLowerCase().endsWith('.lnk')) {
    try {
      target = shell.readShortcutLink(target).target || target;
    } catch {
      /* ponechať .lnk */
    }
  }
  const name = path.basename(target).replace(/\.(exe|lnk|bat|cmd|app)$/i, '');
  return { target, name: name.charAt(0).toUpperCase() + name.slice(1) };
}

/** Ikona programu ako data URL (alebo null pre URL schému / nedostupný súbor). */
export async function programIcon(target: string): Promise<string | null> {
  if (isUrlScheme(target) || !fs.existsSync(target)) return null;
  try {
    const img = await app.getFileIcon(target, { size: 'large' });
    return img.isEmpty() ? null : img.toDataURL();
  } catch {
    return null;
  }
}

/** Podpisy klasického Outlooku na tomto počítači (%APPDATA%\\Microsoft\\Signatures\\*.txt). */
export function outlookSignatures(): Array<{ name: string; text: string }> {
  if (process.platform !== 'win32' || !process.env.APPDATA) return [];
  const dir = path.join(process.env.APPDATA, 'Microsoft', 'Signatures');
  try {
    return fs
      .readdirSync(dir)
      .filter((f) => f.toLowerCase().endsWith('.txt'))
      .map((f) => {
        const buf = fs.readFileSync(path.join(dir, f));
        // Outlook ukladá .txt ako UTF-16 LE s BOM alebo v kódovaní Windows; UTF-8 ako záloha.
        const text = buf[0] === 0xff && buf[1] === 0xfe ? buf.subarray(2).toString('utf16le') : buf.toString('utf8');
        return { name: f.replace(/\.txt$/i, ''), text: text.replace(/\r\n/g, '\n').trim() };
      })
      .filter((x) => x.text);
  } catch {
    return [];
  }
}
