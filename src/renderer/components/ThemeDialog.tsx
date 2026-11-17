// Vlastné dizajnové témy: farby (akcent, horný pruh, pás kont, panely, čítanie, text), svetlá/tmavá, zaoblenie, živý náhľad.
import { Button, Checkbox, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Input } from '@fluentui/react-components';
import { useState } from 'react';
import type { ThemeDef } from '../../core/settings';
import { BUILTIN_THEMES, onColor } from '../theme';

const newId = () => `custom-${Math.random().toString(36).slice(2, 8)}`;
const FIELDS: Array<{ k: 'accent' | 'top' | 'rail' | 'pane' | 'read' | 'text'; label: string }> = [
  { k: 'accent', label: 'Akcent (tlačidlá, výber)' },
  { k: 'top', label: 'Horný pruh' },
  { k: 'rail', label: 'Pás kont' },
  { k: 'pane', label: 'Panely (priečinky, zoznam)' },
  { k: 'read', label: 'Čítanie (pozadie)' },
  { k: 'text', label: 'Text' },
];

function Preview({ t }: { t: ThemeDef }) {
  const top = t.top ?? '#2F6F5E';
  const r = t.radius;
  return (
    <div className="theme-preview" style={{ background: t.read, color: t.text }}>
      <div style={{ gridRow: '1 / 3', background: t.rail, display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 6, paddingTop: 8 }}>
        <span style={{ width: 24, height: 24, borderRadius: '50%', background: t.accent }} />
        <span style={{ width: 24, height: 24, borderRadius: '50%', background: '#1d4ed8' }} />
      </div>
      <div style={{ gridColumn: '2 / 4', background: top, color: onColor(top), display: 'flex', alignItems: 'center', gap: 10, padding: '0 10px' }}>
        <b>Martin</b>
        <span style={{ flex: 1, background: `${onColor(top)}22`, borderRadius: r, padding: '3px 8px' }}>Hľadať v pošte</span>
      </div>
      <div style={{ background: t.pane, padding: 8 }}>
        <div style={{ background: t.accent, color: onColor(t.accent), borderRadius: r, padding: '4px 8px', marginBottom: 8 }}>Nová správa</div>
        <div style={{ padding: '3px 6px', borderRadius: r, background: `${t.accent}26` }}>Doručená pošta</div>
        <div style={{ padding: '3px 6px' }}>Koncepty</div>
        <div style={{ padding: '3px 6px' }}>Odoslané</div>
      </div>
      <div style={{ padding: 12 }}>
        <div style={{ fontSize: 15, fontWeight: 600 }}>Porada v utorok o 10:00</div>
        <div style={{ opacity: 0.6, margin: '4px 0 10px' }}>Jana Kováčová · 10:40</div>
        <div>Ahoj, posielam pozvánku na utorkovú poradu.</div>
        <div style={{ display: 'flex', gap: 6, marginTop: 14 }}>
          <span style={{ background: t.accent, color: onColor(t.accent), padding: '4px 12px', borderRadius: r }}>Odpovedať</span>
          <span style={{ border: `1px solid ${t.text}44`, padding: '4px 12px', borderRadius: r }}>Preposlať</span>
        </div>
      </div>
    </div>
  );
}

export function ThemeDialog(props: { custom: ThemeDef[]; current: string; onSave: (list: ThemeDef[], select?: string) => void; onClose: () => void }) {
  const [list, setList] = useState<ThemeDef[]>(props.custom);
  const [sel, setSel] = useState<string | null>(props.custom.find((t) => t.id === props.current)?.id ?? props.custom[0]?.id ?? null);
  const cur = list.find((t) => t.id === sel);
  const patch = (p: Partial<ThemeDef>) => setList((l) => l.map((t) => (t.id === sel ? { ...t, ...p } : t)));
  const addFrom = (base: ThemeDef) => {
    const t: ThemeDef = { ...base, id: newId(), name: `${base.name.replace(/ \(tmavá\)$/, '')} – vlastná`, top: base.top ?? base.accent };
    setList((l) => [...l, t]);
    setSel(t.id);
  };

  return (
    <Dialog open onOpenChange={(_, d) => !d.open && props.onClose()}>
      <DialogSurface style={{ maxWidth: 1060, width: '92vw' }}>
        <DialogBody style={{ display: 'flex', flexDirection: 'column', maxHeight: '88vh' }}>
          <DialogTitle>Vlastné témy</DialogTitle>
          <DialogContent className="sig-grid">
            <aside className="sig-list">
              <div className="sig-list-items">
                {list.map((t) => (
                  <button key={t.id} className={`folder ${t.id === sel ? 'active' : ''}`} onClick={() => setSel(t.id)}>
                    <span className="theme-dot" style={{ background: t.accent, borderColor: t.rail }} />
                    <span className="grow">{t.name}</span>
                  </button>
                ))}
                {!list.length && <div className="muted small pad">Zatiaľ žiadna vlastná téma.</div>}
              </div>
              <div className="sig-list-actions">
                <div className="muted small">Nová podľa:</div>
                {BUILTIN_THEMES.map((b) => (
                  <Button key={b.id} size="small" onClick={() => addFrom(b)}>{b.name}</Button>
                ))}
              </div>
            </aside>
            <section className="sig-main">
              {cur ? (
                <>
                  <div className="sig-name-row">
                    <label className="sig-label" htmlFor="th-name">Názov</label>
                    <Input id="th-name" className="grow" value={cur.name} onChange={(_, d) => patch({ name: d.value })} />
                    <Button size="small" appearance="subtle" onClick={() => {
                      setList((l) => l.filter((t) => t.id !== cur.id));
                      setSel(list.find((t) => t.id !== cur.id)?.id ?? null);
                    }}>Odstrániť</Button>
                  </div>
                  <div className="theme-edit">
                    <div className="theme-fields">
                      {FIELDS.map((f) => (
                        <label key={f.k} className="theme-field">
                          <span>{f.label}</span>
                          <input
                            type="color"
                            value={(f.k === 'top' ? cur.top ?? cur.accent : cur[f.k]) as string}
                            onChange={(e) => patch({ [f.k]: e.target.value } as Partial<ThemeDef>)}
                          />
                        </label>
                      ))}
                      <Checkbox label="Horný pruh vo farbe konta" checked={cur.top === null} onChange={(_, d) => patch({ top: d.checked ? null : cur.accent })} />
                      <Checkbox label="Tmavá téma (ovládacie prvky)" checked={cur.dark} onChange={(_, d) => patch({ dark: !!d.checked })} />
                      <label className="theme-field">
                        <span>Zaoblenie rohov: {cur.radius} px</span>
                        <input type="range" min={0} max={12} step={1} value={cur.radius} onChange={(e) => patch({ radius: Number(e.target.value) })} />
                      </label>
                    </div>
                    <Preview t={cur} />
                  </div>
                </>
              ) : (
                <div className="empty">Vytvorte tému podľa niektorej zo vstavaných (vľavo) a upravte farby.</div>
              )}
            </section>
          </DialogContent>
          <DialogActions style={{ justifyContent: 'flex-end', alignSelf: 'flex-end' }}>
            <Button appearance="primary" onClick={() => props.onSave(list, cur?.id)}>Uložiť a použiť</Button>
            <Button onClick={props.onClose}>Zrušiť</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
