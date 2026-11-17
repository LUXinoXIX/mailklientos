// Miesto v rozložení, kam hlavný proces položí vložený web (WebContentsView).
// Web leží NAD stránkou: keď je otvorený dialóg, web sa skryje (hidden), inak by ho prekryl.
import { useEffect, useRef } from 'react';

export function WebSlot(props: { viewKey: string; url: string; partition: string; hidden?: boolean; className?: string }) {
  const ref = useRef<HTMLDivElement>(null);
  const { viewKey, url, partition, hidden } = props;

  useEffect(() => {
    const api = window.mailApi;
    const el = ref.current;
    if (!api || !el) return;
    if (hidden) {
      void api.webHide(viewKey);
      return;
    }
    let last = '';
    const rect = () => {
      const r = el.getBoundingClientRect();
      return { x: r.left, y: r.top, width: r.width, height: r.height };
    };
    const b0 = rect();
    last = JSON.stringify(b0);
    void api.webShow(viewKey, url, partition, b0);
    // Poloha sa môže zmeniť aj bez zmeny veľkosti (banner, ribbon), preto aj krátky interval.
    const update = () => {
      const b = rect();
      const s = JSON.stringify(b);
      if (s !== last) {
        last = s;
        void api.webBounds(viewKey, b);
      }
    };
    const ro = new ResizeObserver(update);
    ro.observe(el);
    window.addEventListener('resize', update);
    const t = window.setInterval(update, 250);
    return () => {
      ro.disconnect();
      window.removeEventListener('resize', update);
      window.clearInterval(t);
      void api.webHide(viewKey);
    };
    // url a partition sa pri existujúcom webe ignorujú (stránka si drží vlastnú navigáciu)
  }, [viewKey, url, partition, hidden]);

  return (
    <div ref={ref} className={`webslot ${props.className ?? ''}`}>
      <div className="webslot-hint">{window.mailApi ? 'Načítavam…' : 'Vložený web sa zobrazí v aplikácii.'}</div>
    </div>
  );
}
