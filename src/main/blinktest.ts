// Test blikania: okná, ktoré krok po kroku pridávajú vlastnosti hlavného okna MailKlienta.
// Holé okno nebliká, MailKlient áno → hľadáme, ktorá vlastnosť blikanie spôsobuje.
import path from 'node:path';
import { BrowserWindow, type BrowserWindowConstructorOptions } from 'electron';
import { log } from './diagnostics';
import { attachWindow } from './webviews';

export type BlinkTestKind = 'bare' | 'prefs' | 'webview' | 'app-noweb' | 'app-noframe' | 'app-nodiag' | 'app-demo' | 'app-min';

/** Parametre adresy pre okná s celým MailKlientom (každé vypína niečo navyše oproti testu 4). */
const APP_QUERY: Partial<Record<BlinkTestKind, Record<string, string>>> = {
  'app-noweb': { noweb: '1' },
  'app-noframe': { noweb: '1', noframe: '1' },
  'app-nodiag': { noweb: '1', nodiag: '1' },
  'app-demo': { noweb: '1', demo: '1' },
  'app-min': { noweb: '1', noframe: '1', nodiag: '1', demo: '1' },
};

let appWindowFactory: ((query?: Record<string, string>) => BrowserWindow) | null = null;
export function setAppWindowFactory(f: (query?: Record<string, string>) => BrowserWindow) {
  appWindowFactory = f;
}

/** Rovnaké nastavenia ako hlavné okno MailKlienta. */
export function mainWindowOptions(): BrowserWindowConstructorOptions {
  return {
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
  };
}

const TITLES: Record<BlinkTestKind, string> = {
  bare: '1 – holé okno',
  prefs: '2 – nastavenia okna ako MailKlient',
  webview: '3 – okno s vloženým webom',
  'app-noweb': '4 – MailKlient bez vložených webov',
  'app-noframe': '5 – ako 4, správa bez iframe',
  'app-nodiag': '6 – ako 4, bez diagnostiky',
  'app-demo': '7 – ako 4, ukážkové dáta bez kont',
  'app-min': '8 – ako 4, bez iframe, diagnostiky aj kont',
};

function page(kind: BlinkTestKind) {
  const web =
    kind === 'webview'
      ? `<div style="display:flex;gap:8px;height:220px">
           <webview src="https://www.bing.com" partition="persist:web-blinktest" style="flex:1;border:1px solid #ccc"></webview>
           <webview src="https://www.bing.com" partition="persist:web-blinktest2" style="flex:1;visibility:hidden"></webview>
         </div>`
      : '';
  return `<!doctype html><html><head><meta charset="utf-8"><title>Test ${TITLES[kind]}</title>
<style>body{font:15px Segoe UI,sans-serif;margin:20px;background:#fff;color:#222}
button{font-size:15px;padding:10px 18px;margin:6px;border:1px solid #888;border-radius:4px;background:#f3f3f3}
#b{font-size:18px;font-weight:600}</style></head><body>
<h2>Test ${TITLES[kind]}</h2>
<p>Klikajte na tlačidlá a do textu. Bliká toto okno?</p>
<button>Tlačidlo 1</button><button>Tlačidlo 2</button><input style="font-size:15px;padding:6px;width:220px" placeholder="sem niečo napíšte">
<p>Kliknutia: <span id="c">0</span> &nbsp; · &nbsp; Okno stratilo zameranie: <span id="b">0</span></p>
${web}
<script>let c=0;document.addEventListener('mousedown',()=>{document.getElementById('c').textContent=++c});
window.setBlurs=n=>{const e=document.getElementById('b');e.textContent=n;e.style.color=n?'#c00':'#222'};</script>
</body></html>`;
}

function watchBlur(w: BrowserWindow, kind: BlinkTestKind) {
  let blurs = 0;
  const report = () =>
    w.webContents.executeJavaScript(`window.setBlurs && window.setBlurs(${blurs})`).catch(() => undefined);
  w.on('blur', () => {
    blurs++;
    log(`test blikania ${kind}: blur`);
    void report();
  });
  w.on('focus', () => log(`test blikania ${kind}: focus`));
  w.webContents.on('did-finish-load', () => void report());
}

export function openBlinkTest(kind: BlinkTestKind) {
  log(`test blikania ${kind}: otvorené`);
  const query = APP_QUERY[kind];
  if (query) {
    if (!appWindowFactory) return;
    const w = appWindowFactory(query);
    w.setTitle(`MailKlient – test ${TITLES[kind]}`);
    watchBlur(w, kind);
    return;
  }
  const opts: BrowserWindowConstructorOptions =
    kind === 'bare' ? { autoHideMenuBar: true } : mainWindowOptions();
  const w = new BrowserWindow({ ...opts, width: 720, height: kind === 'webview' ? 600 : 420, title: `Test ${TITLES[kind]}` });
  if (kind === 'webview') attachWindow(w);
  watchBlur(w, kind);
  void w.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(page(kind)));
}
