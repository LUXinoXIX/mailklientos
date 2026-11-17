// Stabilita vykresľovania a záznam chýb.
// Biela obrazovka / blikanie na Windows: najčastejšie pád grafického procesu (GPU) alebo chybné
// vyhodnotenie zakrytia okna (Native Window Occlusion). Prvé riešime prepínačom „bez hardvérovej
// akcelerácie“, druhé vypnutím tejto funkcie Chromia. Všetko sa zapisuje do mailklient.log.
import fs from 'node:fs';
import path from 'node:path';
import { app, BrowserWindow, shell } from 'electron';

interface Prefs {
  /** true = vykresľovať bez GPU (softvérovo). */
  disableGpu?: boolean;
}

const prefsFile = () => path.join(app.getPath('userData'), 'prefs.json');
export const logFile = () => path.join(app.getPath('userData'), 'mailklient.log');

export function loadPrefs(): Prefs {
  try {
    return JSON.parse(fs.readFileSync(prefsFile(), 'utf8')) as Prefs;
  } catch {
    return {};
  }
}

export function savePrefs(p: Prefs) {
  fs.mkdirSync(path.dirname(prefsFile()), { recursive: true });
  fs.writeFileSync(prefsFile(), JSON.stringify(p, null, 2), 'utf8');
}

export function log(...parts: unknown[]) {
  const line = `${new Date().toISOString()} ${parts.map((p) => (typeof p === 'string' ? p : JSON.stringify(p))).join(' ')}\n`;
  try {
    const f = logFile();
    fs.mkdirSync(path.dirname(f), { recursive: true });
    // log nenechať rásť donekonečna
    if (fs.existsSync(f) && fs.statSync(f).size > 2_000_000) fs.renameSync(f, `${f}.old`);
    fs.appendFileSync(f, line, 'utf8');
  } catch {
    /* zápis logu nesmie zhodiť aplikáciu */
  }
}

/** Volať PRED app.whenReady(): prepínače Chromia podľa nastavení. */
export function applyRenderingSwitches() {
  // Windows: chybné vyhodnotenie zakrytia okna spôsobuje biele/prázdne okno a blikanie.
  app.commandLine.appendSwitch('disable-features', 'CalculateNativeWinOcclusion');
  if (loadPrefs().disableGpu) {
    app.disableHardwareAcceleration();
    log('štart: hardvérová akcelerácia VYPNUTÁ (nastavenie)');
  }
}

let gpuCrashes = 0;

/** Sledovanie pádov procesov; pri páde GPU dá vedieť rendereru (ponuka vypnúť akceleráciu). */
export function watchProcesses() {
  log(`štart: verzia ${app.getVersion()}, Electron ${process.versions.electron}, ${process.platform} ${process.arch}`);
  app.on('child-process-gone', (_e, d) => {
    log('pád procesu', { type: d.type, reason: d.reason, exitCode: d.exitCode, name: d.name });
    if (d.type === 'GPU' && d.reason !== 'clean-exit') {
      gpuCrashes++;
      for (const w of BrowserWindow.getAllWindows())
        if (!w.isDestroyed()) w.webContents.send('app:event', { type: 'gpu-crash', count: gpuCrashes });
    }
  });
  app.on('render-process-gone', (_e, wc, d) => {
    log('pád stránky', { url: wc.getURL().slice(0, 120), reason: d.reason, exitCode: d.exitCode });
  });
  app.on('web-contents-created', (_e, wc) => {
    wc.on('unresponsive', () => log('stránka neodpovedá', wc.getURL().slice(0, 120)));
    wc.on('responsive', () => log('stránka opäť odpovedá', wc.getURL().slice(0, 120)));
  });
  try {
    log('GPU', app.getGPUFeatureStatus());
  } catch {
    /* nič */
  }
}

export function gpuDisabled() {
  return !!loadPrefs().disableGpu;
}

/** Zapne/vypne hardvérovú akceleráciu a reštartuje aplikáciu. */
export function setGpuDisabled(off: boolean) {
  savePrefs({ ...loadPrefs(), disableGpu: off });
  log(`nastavenie: hardvérová akcelerácia ${off ? 'vypnutá' : 'zapnutá'}, reštart`);
  app.relaunch();
  app.exit(0);
}

export function openLog() {
  if (!fs.existsSync(logFile())) log('záznam vytvorený');
  shell.showItemInFolder(logFile());
}
