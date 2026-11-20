// Štartovací banner „Universum eMail“ (súhvezdie v tvare obálky), kým sa nenačíta hlavné okno.
import { app, BrowserWindow } from 'electron';

const html = (version: string) => `<!doctype html><html><head><meta charset="utf-8"><style>
html,body{margin:0;height:100%;background:#05081a;overflow:hidden;-webkit-user-select:none;cursor:default}
svg{display:block;width:100%;height:100%}
text{font-family:'Segoe UI Variable Display','Segoe UI',-apple-system,BlinkMacSystemFont,'Ubuntu',Arial,sans-serif}
@keyframes draw{0%{stroke-dashoffset:1}60%,100%{stroke-dashoffset:0}}
@keyframes pulse{0%,100%{opacity:.25}50%{opacity:.7}}
@keyframes tw{0%,100%{opacity:.25}50%{opacity:.9}}
@keyframes load{0%{transform:translateX(-70px)}100%{transform:translateX(250px)}}
.ln{stroke-dasharray:1;stroke-dashoffset:1;animation:draw 2.4s ease-out forwards}
.l2{animation-delay:.4s}.l3{animation-delay:.8s}
.halo{animation:pulse 2.6s ease-in-out infinite}
.s{animation:tw 2.4s ease-in-out infinite}
.bar{animation:load 1.6s ease-in-out infinite}
</style></head><body>
<svg viewBox="0 0 720 405">
<defs>
<radialGradient id="bg" cx="0.75" cy="0.5" r="0.55"><stop offset="0" stop-color="#121a45"/><stop offset="1" stop-color="#121a45" stop-opacity="0"/></radialGradient>
<clipPath id="track"><rect x="56" y="359" width="180" height="2" rx="1"/></clipPath>
</defs>
<rect width="720" height="405" fill="url(#bg)"/>
<g fill="#fff"><circle class="s" cx="40" cy="40" r="1.2"/><circle class="s" style="animation-delay:.6s" cx="140" cy="300" r="1"/><circle class="s" style="animation-delay:1.1s" cx="260" cy="40" r="1.4"/><circle class="s" style="animation-delay:.3s" cx="320" cy="350" r="1"/><circle class="s" style="animation-delay:1.5s" cx="370" cy="70" r="1.2"/><circle class="s" style="animation-delay:.9s" cx="700" cy="60" r="1"/><circle class="s" style="animation-delay:1.8s" cx="690" cy="330" r="1.4"/><circle class="s" style="animation-delay:.4s" cx="480" cy="360" r="1"/><circle class="s" style="animation-delay:1.2s" cx="590" cy="40" r="1.2"/><circle class="s" style="animation-delay:.7s" cx="380" cy="220" r="1"/><circle class="s" style="animation-delay:1.4s" cx="200" cy="200" r="1"/></g>
<g fill="none" stroke="#7dd3fc" stroke-width="1.2" stroke-linecap="round" opacity=".6">
<path class="ln" pathLength="1" d="M430 110 H650 V280 H430 Z"/>
<path class="ln l2" pathLength="1" d="M430 110 L540 200 L650 110"/>
<path class="ln l3" pathLength="1" d="M430 280 L360 330 L300 318" opacity=".6"/>
</g>
<g fill="#7dd3fc"><circle class="halo" cx="430" cy="110" r="12"/><circle class="halo" cx="650" cy="110" r="12"/><circle class="halo" cx="650" cy="280" r="10"/><circle class="halo" cx="430" cy="280" r="10"/><circle class="halo" cx="540" cy="200" r="16" fill="#f5c76b"/></g>
<g fill="#fff"><circle cx="430" cy="110" r="3.6"/><circle cx="650" cy="110" r="3.6"/><circle cx="650" cy="280" r="3"/><circle cx="430" cy="280" r="3"/><circle cx="360" cy="330" r="2.2" opacity=".8"/><circle cx="300" cy="318" r="1.8" opacity=".7"/></g>
<circle cx="540" cy="200" r="5" fill="#fde68a"/>
<text x="54" y="170" font-size="54" font-weight="600" fill="#fff" letter-spacing="-0.8">Universum</text>
<text x="57" y="200" font-size="17" font-weight="500" fill="#7dd3fc" letter-spacing="7">EMAIL</text>
<text x="56" y="346" font-size="12" fill="#a9b4d0">Pripájam kontá…</text>
<rect x="56" y="359" width="180" height="2" rx="1" fill="#1a2244"/>
<g clip-path="url(#track)"><rect class="bar" x="56" y="359" width="70" height="2" rx="1" fill="#f5c76b"/></g>
<text x="692" y="386" text-anchor="end" font-size="11" fill="#8d98b8" letter-spacing=".4">verzia ${version}</text>
</svg></body></html>`;

/** Zobrazí banner; vráti funkciu, ktorá ho zavrie (najskôr po minDuration ms). */
export function showSplash(minDuration = 1400): (after: () => void) => void {
  const started = Date.now();
  const w = new BrowserWindow({
    width: 560,
    height: 315,
    frame: false,
    resizable: false,
    movable: true,
    skipTaskbar: true,
    // na Macu nie navrchu – inak by zakryl otázku Kľúčenky na prístup k uloženým heslám a program by vyzeral zaseknutý
    alwaysOnTop: process.platform !== 'darwin',
    show: false,
    center: true,
    backgroundColor: '#05081a',
    webPreferences: { sandbox: true, contextIsolation: true },
  });
  w.once('ready-to-show', () => w.show());
  void w.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html(app.getVersion()))}`);
  let closed = false;
  // poistka: ak hlavné okno dlho nie je pripravené (napr. čaká na Kľúčenku), banner po 12 s zmizne sám
  const guard = setTimeout(() => {
    if (!closed && !w.isDestroyed()) w.close();
  }, 12_000);
  return (after) => {
    clearTimeout(guard);
    if (closed) return after();
    closed = true;
    setTimeout(() => {
      after();
      if (!w.isDestroyed()) w.close();
    }, Math.max(0, minDuration - (Date.now() - started)));
  };
}
