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
  /** true = použiť grafickú kartu aj keď je jej ovládač v Chromiu na zozname blokovaných. */
  forceGpu?: boolean;
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
  const p = loadPrefs();
  if (p.disableGpu) {
    app.disableHardwareAcceleration();
    log('štart: hardvérová akcelerácia VYPNUTÁ (nastavenie)');
  } else if (p.forceGpu) {
    // Grafický proces beží v sandboxe; bezpečnostný softvér (firemné pravidlá, antivírus vkladajúci DLL)
    // ho často zablokuje a Chromium potom vykresľuje bez grafickej karty. Bez sandboxu sa spustí.
    app.commandLine.appendSwitch('ignore-gpu-blocklist');
    app.commandLine.appendSwitch('disable-gpu-sandbox');
    log('štart: grafická karta VYNÚTENÁ (ignore-gpu-blocklist, disable-gpu-sandbox)');
  }
}

let gpuCrashes = 0;

/** Sledovanie pádov procesov; pri páde GPU dá vedieť rendereru (ponuka vypnúť akceleráciu). */
/** Stav grafiky až po spustení okna (pri štarte býva ešte neúplný). */
export function logGpuLater() {
  setTimeout(() => {
    try {
      const st = app.getGPUFeatureStatus();
      log('GPU po štarte', { gpu_compositing: st.gpu_compositing, rasterization: st.rasterization, webgl: st.webgl });
    } catch {
      /* nič */
    }
  }, 8000);
}

export function watchProcesses() {
  log(`štart: verzia ${app.getVersion()}, Electron ${process.versions.electron}, ${process.platform} ${process.arch}`);
  logGpuLater();
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
  // ovládač grafiky a dôvod zablokovania (ako chrome://gpu)
  void app
    .getGPUInfo('complete')
    .then((info) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const i = info as any;
      const dev = (i.gpuDevice ?? []).map((d: any) => ({
        vendor: d.vendorString ?? d.vendorId,
        device: d.deviceString ?? d.deviceId,
        driver: d.driverVersion,
        active: d.active,
      }));
      log('GPU zariadenia', dev);
      log('GPU dôvody', { blocklist: i.auxAttributes?.gpuBlocklistReasons ?? i.featureStatus, workarounds: i.driverBugWorkarounds?.length, disabled: i.auxAttributes?.softwareRendering });
    })
    .catch((e: unknown) => log('GPU info chyba', String(e)));
}

export function gpuDisabled() {
  return !!loadPrefs().disableGpu;
}

export function gpuForced() {
  return !!loadPrefs().forceGpu;
}

/** Vynúti grafickú kartu (ignoruje zoznam blokovaných ovládačov) a reštartuje. */
export function setGpuForced(on: boolean) {
  savePrefs({ ...loadPrefs(), forceGpu: on, disableGpu: on ? false : loadPrefs().disableGpu });
  log(`nastavenie: vynútená grafická karta ${on ? 'zapnutá' : 'vypnutá'}, reštart`);
  app.relaunch();
  app.exit(0);
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
