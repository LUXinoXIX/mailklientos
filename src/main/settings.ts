// Nastavenia aplikácie (weby a spúšťače v pravom paneli) v JSON súbore v používateľských dátach.
// Spúšťanie programov a načítanie ich ikon.
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { app, nativeImage, BrowserWindow, dialog, shell } from 'electron';
import { DEFAULT_SETTINGS, type AppSettings, type LauncherItem } from '../core/settings';

const file = () => path.join(app.getPath('userData'), 'settings.json');

export function loadSettings(): AppSettings {
  try {
    const s = JSON.parse(fs.readFileSync(file(), 'utf8')) as Partial<AppSettings>;
    return {
      // všetky ostatné nastavenia (pravidlá, kategórie, šablóny, úlohy, témy…) ostávajú tak, ako boli uložené
      ...structuredClone(DEFAULT_SETTINGS),
      ...s,
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
  const err = await shell.openPath(t); // '' = v poriadku, inak text chyby
  if (err && process.platform === 'darwin') {
    // Mac: záložne cez príkaz open (aplikácie .app aj súbory)
    return new Promise((res) => {
      const child = spawn('open', [t], { detached: true, stdio: 'ignore' });
      child.on('error', (e) => res(`${err} (${e.message})`));
      child.on('exit', (code) => res(code === 0 ? '' : err));
      child.unref();
    });
  }
  return err;
}

/** Dialóg na výber programu. */
export async function pickProgram(win: BrowserWindow | null): Promise<{ target: string; name: string } | null> {
  const filters =
    process.platform === 'win32'
      ? [{ name: 'Programy', extensions: ['exe', 'lnk', 'bat', 'cmd', 'msc'] }, { name: 'Všetky súbory', extensions: ['*'] }]
      : process.platform === 'darwin'
        ? [{ name: 'Aplikácie', extensions: ['app'] }, { name: 'Všetky súbory', extensions: ['*'] }]
        : [{ name: 'Všetky súbory', extensions: ['*'] }];
  const opts: Electron.OpenDialogOptions = {
    title: 'Vyber program',
    // Mac: aplikácia (.app) je balík – vyberá sa ako súbor (s treatPackageAsDirectory by sa otvorila ako priečinok)
    properties: ['openFile'],
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
/** Mac: ikona aplikácie priamo zo súboru .icns v balíku – PNG vo vnútri sa vráti ako data URL bez akéhokoľvek
    natívneho spracovania obrázka (app.getFileIcon / nativeImage na macOS 27 zhadzovali program). */
function macAppIcon(appPath: string): string | null {
  try {
    const plist = fs.readFileSync(path.join(appPath, 'Contents', 'Info.plist'), 'utf8');
    let name = /<key>CFBundleIconFile<\/key>\s*<string>([^<]+)<\/string>/.exec(plist)?.[1];
    if (!name) return null;
    if (!name.endsWith('.icns')) name += '.icns';
    const buf = fs.readFileSync(path.join(appPath, 'Contents', 'Resources', name));
    if (buf.length < 16 || buf.subarray(0, 4).toString('latin1') !== 'icns') return null;
    const PNG = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    let best: Buffer | null = null;
    let bw = 0;
    for (let off = 8; off + 8 <= buf.length; ) {
      const len = buf.readUInt32BE(off + 4);
      if (len < 8 || off + len > buf.length) break;
      const data = buf.subarray(off + 8, off + len);
      if (data.length > 24 && data.subarray(0, 8).equals(PNG)) {
        const w = data.readUInt32BE(16);
        // najmenšia aspoň 64 px, inak najväčšia
        if (!best || (w >= 64 && (bw < 64 || w < bw)) || (bw < 64 && w > bw)) {
          best = data;
          bw = w;
        }
      }
      off += len;
    }
    return best ? 'data:image/png;base64,' + best.toString('base64') : null;
  } catch {
    return null;
  }
}

/** Poistka: značka sa zapíše pred načítaním ikony a zmaže po ňom. Ak ostala z minulého behu (program pritom spadol),
    ikony programov sa už nenačítavajú – zobrazí sa všeobecná ikona a program sa dá normálne spustiť. */
const iconGuard = () => path.join(app.getPath('userData'), 'icon-guard');
let iconsOff: boolean | null = null;

export async function programIcon(target: string): Promise<string | null> {
  if (isUrlScheme(target) || !fs.existsSync(target)) return null;
  if (iconsOff === null) iconsOff = fs.existsSync(iconGuard());
  if (iconsOff) return null;
  if (process.platform === 'darwin' && !/\.app\/?$/i.test(target)) return null;
  try {
    fs.writeFileSync(iconGuard(), target);
  } catch {
    /* bez poistky */
  }
  try {
    if (process.platform === 'darwin') return macAppIcon(target);
    const img = await app.getFileIcon(target, { size: 'large' });
    return img.isEmpty() ? null : img.toDataURL();
  } catch {
    return null;
  } finally {
    try {
      fs.rmSync(iconGuard(), { force: true });
    } catch {
      /* */
    }
  }
}

/** Text súboru podpisu Outlooku: UTF-16 s BOM, UTF-8 s BOM, inak kódovanie z <meta charset> (zvyčajne windows-1250). */
function decodeOutlook(buf: Buffer, html: boolean): string {
  if (buf[0] === 0xff && buf[1] === 0xfe) return buf.subarray(2).toString('utf16le');
  if (buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) return buf.subarray(3).toString('utf8');
  const head = buf.subarray(0, 2048).toString('latin1');
  const cs = html ? /charset=["']?([\w-]+)/i.exec(head)?.[1] : undefined;
  try {
    return new TextDecoder(cs ?? 'utf-8').decode(buf);
  } catch {
    return buf.toString('utf8');
  }
}

const MIME: Record<string, string> = { '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif', '.bmp': 'image/bmp', '.svg': 'image/svg+xml', '.webp': 'image/webp' };

/** Podpisy klasického Outlooku z %APPDATA%\Microsoft\Signatures: .htm (formát a obrázky ako data: URL) + .txt (text). */
export function outlookSignatures(): Array<{ name: string; text: string; html?: string }> {
  if (process.platform !== 'win32' || !process.env.APPDATA) return [];
  const dir = path.join(process.env.APPDATA, 'Microsoft', 'Signatures');
  let files: string[];
  try {
    files = fs.readdirSync(dir);
  } catch {
    return [];
  }
  const names = new Set(files.filter((f) => /\.(htm|html|txt)$/i.test(f)).map((f) => f.replace(/\.(htm|html|txt)$/i, '')));
  const out: Array<{ name: string; text: string; html?: string }> = [];
  for (const name of names) {
    const find = (ext: string) => files.find((f) => f.toLowerCase() === `${name}${ext}`.toLowerCase());
    let text = '';
    let html: string | undefined;
    const txt = find('.txt');
    if (txt) {
      try {
        text = decodeOutlook(fs.readFileSync(path.join(dir, txt)), false).replace(/\r\n/g, '\n').trim();
      } catch {
        /* nečitateľný súbor */
      }
    }
    const htm = find('.htm') ?? find('.html');
    if (htm) {
      try {
        html = decodeOutlook(fs.readFileSync(path.join(dir, htm)), true);
        // obrázky z priečinka „<názov>_files“ vložiť priamo (data: URL), aby podpis fungoval aj bez súborov
        html = html.replace(/(<img[^>]*?\ssrc=)(["'])([^"']+)\2/gi, (m, pre: string, q: string, src: string) => {
          if (/^(data|https?|cid):/i.test(src)) return m;
          const file = path.join(dir, decodeURIComponent(src).replace(/\//g, path.sep));
          const mime = MIME[path.extname(file).toLowerCase()];
          if (!mime || !file.startsWith(dir)) return m;
          try {
            return `${pre}${q}data:${mime};base64,${fs.readFileSync(file).toString('base64')}${q}`;
          } catch {
            return m;
          }
        });
        // VML náhrady obrázkov (v:imagedata) sú len pre Word; HTML <img> stačí
      } catch {
        html = undefined;
      }
    }
    if (text || html) out.push({ name, text, html });
  }
  return out;
}
