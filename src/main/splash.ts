// Štartovací banner „Universum eMail“ (galaxia a let obálky), kým sa nenačíta hlavné okno.
import { BrowserWindow } from 'electron';

const HTML = `<!doctype html><html><head><meta charset="utf-8"><style>
html,body{margin:0;height:100%;background:#05060b;overflow:hidden;font-family:'Segoe UI',Arial,sans-serif;-webkit-user-select:none;cursor:default}
svg{display:block;width:100%;height:100%}
@keyframes fly{0%{transform:translate(0,0) scale(.9);opacity:0}15%{opacity:1}100%{transform:translate(172px,-60px) scale(.35);opacity:0}}
@keyframes glow{0%,100%{opacity:.15}50%{opacity:.32}}
@keyframes tw{0%,100%{opacity:.25}50%{opacity:1}}
@keyframes bar{0%{width:0}100%{width:120px}}
.env{animation:fly 1.6s ease-in infinite;transform-origin:59px 200px}
.halo{animation:glow 2s ease-in-out infinite}
.s{animation:tw 2.4s ease-in-out infinite}
.bar{animation:bar 1.8s ease-out forwards}
</style></head><body>
<svg viewBox="0 0 480 270">
<g fill="#fff"><circle class="s" cx="30" cy="40" r="1"/><circle class="s" style="animation-delay:.6s" cx="90" cy="20" r=".8"/><circle class="s" style="animation-delay:1.1s" cx="440" cy="30" r="1.2"/><circle class="s" style="animation-delay:.3s" cx="410" cy="230" r="1"/><circle class="s" style="animation-delay:1.5s" cx="70" cy="240" r="1.2"/><circle class="s" style="animation-delay:.9s" cx="200" cy="230" r=".8"/><circle class="s" style="animation-delay:1.8s" cx="330" cy="20" r=".9"/><circle class="s" style="animation-delay:.4s" cx="150" cy="110" r=".7"/><circle class="s" style="animation-delay:1.2s" cx="360" cy="180" r=".8"/></g>
<g transform="rotate(-20 240 135)"><ellipse cx="240" cy="135" rx="210" ry="38" fill="#5aa2ff" opacity=".06"/><ellipse cx="240" cy="135" rx="140" ry="22" fill="#5aa2ff" opacity=".1"/><ellipse cx="240" cy="135" rx="70" ry="11" fill="#9cc8ff" opacity=".25"/></g>
<circle cx="240" cy="135" r="9" fill="#ffd166"/><circle class="halo" cx="240" cy="135" r="24" fill="#ffd166"/>
<path d="M60 200 Q 160 150 232 137" fill="none" stroke="#5aa2ff" stroke-width="1.5" stroke-dasharray="3 5" opacity=".7"/>
<g class="env"><rect x="44" y="190" width="30" height="21" rx="3" fill="#e8f1ff"/><path d="M45 192 L59 203 L73 192" fill="none" stroke="#13233f" stroke-width="2"/></g>
<text x="240" y="70" text-anchor="middle" font-size="30" font-weight="600" fill="#fff">Universum eMail</text>
<rect x="180" y="232" width="120" height="3" rx="1.5" fill="#1b2433"/><rect class="bar" x="180" y="232" width="0" height="3" rx="1.5" fill="#5aa2ff"/>
<text x="240" y="252" text-anchor="middle" font-size="10" fill="#7c8db0">Pripájam kontá…</text>
</svg></body></html>`;

/** Zobrazí banner; vráti funkciu, ktorá ho zavrie (najskôr po minDuration ms). */
export function showSplash(minDuration = 1400): (after: () => void) => void {
  const started = Date.now();
  const w = new BrowserWindow({
    width: 480,
    height: 270,
    frame: false,
    resizable: false,
    movable: true,
    skipTaskbar: true,
    alwaysOnTop: true,
    show: false,
    center: true,
    backgroundColor: '#05060b',
    webPreferences: { sandbox: true, contextIsolation: true },
  });
  w.once('ready-to-show', () => w.show());
  void w.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(HTML)}`);
  let closed = false;
  return (after) => {
    if (closed) return after();
    closed = true;
    setTimeout(() => {
      after();
      if (!w.isDestroyed()) w.close();
    }, Math.max(0, minDuration - (Date.now() - started)));
  };
}
