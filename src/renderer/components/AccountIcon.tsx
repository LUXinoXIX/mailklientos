// Ikonka konta v páse kont: automaticky iniciály, vlastný text (2–3 znaky),
// pripravený symbol (pošta, cloud, Zimbra, IMAP …) alebo vlastný obrázok.
import { Radio, RadioGroup } from '@fluentui/react-components';
import type { ReactNode } from 'react';

export interface IconCfg {
  /** auto = iniciály z názvu konta */
  iconKind?: 'auto' | 'text' | 'preset' | 'image';
  iconText?: string;
  iconPreset?: string;
  /** malý obrázok ako data: URL (64 × 64) */
  iconImage?: string;
}

const P = (d: ReactNode) => (
  <svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    {d}
  </svg>
);

/** Vlastné jednoduché symboly (nie logá značiek). */
export const ICON_PRESETS: Array<{ id: string; label: string; svg: ReactNode }> = [
  { id: 'mail', label: 'Pošta (IMAP)', svg: P(<><rect x="3" y="5" width="18" height="14" rx="2" /><path d="m3.5 6.5 8.5 6.5 8.5-6.5" /></>) },
  { id: 'cloudmail', label: 'Pošta v cloude (365)', svg: P(<><path d="M7 18a4 4 0 0 1-.6-7.95A5.5 5.5 0 0 1 17 8.5a3.75 3.75 0 0 1 .5 7.47" /><rect x="8" y="13" width="9" height="7" rx="1.2" /><path d="m8.4 13.6 4.1 3 4.1-3" /></>) },
  { id: 'grid', label: 'Aplikácie (Office)', svg: P(<><rect x="4" y="4" width="7" height="7" rx="1.5" /><rect x="13" y="4" width="7" height="7" rx="1.5" /><rect x="4" y="13" width="7" height="7" rx="1.5" /><rect x="13" y="13" width="7" height="7" rx="1.5" /></>) },
  { id: 'z', label: 'Z (Zimbra)', svg: P(<><path d="M7 6h10L7 18h10" strokeWidth="2.4" /></>) },
  { id: 'at', label: 'Zavináč', svg: P(<><circle cx="12" cy="12" r="3.5" /><path d="M15.5 12v1.5a2.5 2.5 0 0 0 5 0V12a8.5 8.5 0 1 0-3.4 6.8" /></>) },
  { id: 'server', label: 'Server', svg: P(<><rect x="4" y="4" width="16" height="7" rx="1.5" /><rect x="4" y="13" width="16" height="7" rx="1.5" /><path d="M8 7.5h.01M8 16.5h.01" strokeWidth="2.6" /></>) },
  { id: 'briefcase', label: 'Práca', svg: P(<><rect x="3" y="7" width="18" height="13" rx="2" /><path d="M9 7V5.5A1.5 1.5 0 0 1 10.5 4h3A1.5 1.5 0 0 1 15 5.5V7M3 12.5h18" /></>) },
  { id: 'home', label: 'Domov / osobné', svg: P(<><path d="m4 11 8-7 8 7" /><path d="M6 9.5V20h12V9.5" /><path d="M10 20v-5h4v5" /></>) },
  { id: 'building', label: 'Firma', svg: P(<><rect x="5" y="3" width="14" height="18" rx="1.5" /><path d="M9 7h.01M15 7h.01M9 11h.01M15 11h.01M9 15h.01M15 15h.01" strokeWidth="2.6" /><path d="M11 21v-3h2v3" /></>) },
  { id: 'heart', label: 'Rodina', svg: P(<path d="M12 20s-7-4.4-7-10a4 4 0 0 1 7-2.6A4 4 0 0 1 19 10c0 5.6-7 10-7 10Z" />) },
  { id: 'star', label: 'Dôležité', svg: P(<path d="m12 3.5 2.6 5.3 5.9.9-4.3 4.1 1 5.8L12 16.9l-5.2 2.7 1-5.8-4.3-4.1 5.9-.9Z" />) },
  { id: 'globe', label: 'Web', svg: P(<><circle cx="12" cy="12" r="8.5" /><path d="M3.5 12h17M12 3.5c2.5 2.6 2.5 14.4 0 17M12 3.5c-2.5 2.6-2.5 14.4 0 17" /></>) },
];

/** Obsah štvorčeka konta (bez rámika). */
export function AccountIconContent({ cfg, fallback }: { cfg: IconCfg; fallback: string }) {
  const kind = cfg.iconKind ?? 'auto';
  if (kind === 'image' && cfg.iconImage) return <img src={cfg.iconImage} alt="" className="acc-img" />;
  if (kind === 'preset' && cfg.iconPreset) {
    const p = ICON_PRESETS.find((x) => x.id === cfg.iconPreset);
    if (p) return <>{p.svg}</>;
  }
  if (kind === 'text' && cfg.iconText?.trim()) {
    const t = cfg.iconText.trim().slice(0, 3);
    return <span className={`acc-txt l${t.length}`}>{t}</span>;
  }
  return <>{fallback}</>;
}

async function toSmallDataUrl(file: File): Promise<string> {
  const url = URL.createObjectURL(file);
  try {
    const img = await new Promise<HTMLImageElement>((res, rej) => {
      const i = new Image();
      i.onload = () => res(i);
      i.onerror = rej;
      i.src = url;
    });
    const c = document.createElement('canvas');
    c.width = c.height = 64;
    const x = c.getContext('2d')!;
    const s = Math.min(img.width, img.height);
    x.drawImage(img, (img.width - s) / 2, (img.height - s) / 2, s, s, 0, 0, 64, 64);
    return c.toDataURL('image/png');
  } finally {
    URL.revokeObjectURL(url);
  }
}

/** Výber ikonky v nastaveniach konta. */
export function AccountIconPicker(props: { value: IconCfg; color: string; fallback: string; onChange: (v: IconCfg) => void }) {
  const v = props.value;
  const kind = v.iconKind ?? 'auto';
  const set = (p: Partial<IconCfg>) => props.onChange({ ...v, ...p });
  return (
    <div className="icon-picker">
      <div className="icon-picker-row">
        <span className="acc-sq preview" style={{ background: props.color, color: '#fff', borderColor: props.color }}>
          <AccountIconContent cfg={v} fallback={props.fallback} />
        </span>
        <RadioGroup layout="horizontal" value={kind} onChange={(_, d) => set({ iconKind: d.value as IconCfg['iconKind'] })}>
          <Radio value="auto" label="Iniciály (automaticky)" />
          <Radio value="text" label="Text" />
          <Radio value="preset" label="Symbol" />
          <Radio value="image" label="Obrázok" />
        </RadioGroup>
      </div>
      {kind === 'text' && (
        <input
          className="icon-text-input"
          maxLength={3}
          placeholder="napr. PDG"
          value={v.iconText ?? ''}
          onChange={(e) => set({ iconText: e.target.value.toUpperCase() })}
          aria-label="Text na ikonke (2–3 znaky)"
        />
      )}
      {kind === 'preset' && (
        <div className="icon-presets">
          {ICON_PRESETS.map((p) => (
            <button
              key={p.id}
              type="button"
              className={`icon-preset ${v.iconPreset === p.id ? 'on' : ''}`}
              style={v.iconPreset === p.id ? { background: props.color, color: '#fff' } : undefined}
              title={p.label}
              aria-label={p.label}
              onClick={() => set({ iconPreset: p.id })}
            >
              {p.svg}
            </button>
          ))}
        </div>
      )}
      {kind === 'image' && (
        <div className="row gap">
          <label className="icon-upload">
            <input
              type="file"
              accept="image/*"
              onChange={async (e) => {
                const f = e.target.files?.[0];
                if (f) set({ iconImage: await toSmallDataUrl(f) });
              }}
            />
            <span className="icon-upload-btn">Vybrať obrázok…</span>
          </label>
          <span className="muted small">PNG, JPG alebo SVG; oreže sa na štvorec 64 × 64.</span>
        </div>
      )}
    </div>
  );
}
