// Šablóny správ a rýchle texty: formátovaný obsah (aj s obrázkami), voliteľný predmet.
// Vkladajú sa pri písaní cez Správa → Šablóny.
import { Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Input } from '@fluentui/react-components';
import { useState } from 'react';
import type { AppSettings, Template } from '../../core/settings';
import { SignatureEditor } from './SignatureEditor';

const newId = () => Math.random().toString(36).slice(2, 10);

export function TemplatesDialog(props: { settings: AppSettings; onSave: (s: AppSettings) => void; onClose: () => void }) {
  const [list, setList] = useState<Template[]>(props.settings.templates ?? []);
  const [sel, setSel] = useState<string | null>(list[0]?.id ?? null);
  const cur = list.find((t) => t.id === sel);
  const patch = (p: Partial<Template>) => setList((l) => l.map((t) => (t.id === sel ? { ...t, ...p } : t)));

  return (
    <Dialog open onOpenChange={(_, d) => !d.open && props.onClose()}>
      <DialogSurface style={{ maxWidth: 1100, width: '92vw' }}>
        <DialogBody style={{ display: 'flex', flexDirection: 'column', maxHeight: '86vh' }}>
          <DialogTitle>Šablóny a rýchle texty</DialogTitle>
          <DialogContent className="sig-grid">
            <aside className="sig-list">
              <div className="sig-list-items">
                {list.map((t) => (
                  <button key={t.id} className={`folder ${t.id === sel ? 'active' : ''}`} onClick={() => setSel(t.id)}>
                    <span className="grow">{t.name || '(bez názvu)'}</span>
                  </button>
                ))}
                {!list.length && <div className="muted small pad">Zatiaľ žiadna šablóna.</div>}
              </div>
              <div className="sig-list-actions">
                <Button
                  size="small"
                  appearance="primary"
                  onClick={() => {
                    const t: Template = { id: newId(), name: 'Nová šablóna', html: '' };
                    setList((l) => [...l, t]);
                    setSel(t.id);
                  }}
                >
                  + Nová šablóna
                </Button>
              </div>
            </aside>
            <section className="sig-main">
              {cur ? (
                <>
                  <div className="sig-name-row">
                    <label className="sig-label" htmlFor="tpl-name">Názov</label>
                    <Input id="tpl-name" className="grow" value={cur.name} onChange={(_, d) => patch({ name: d.value })} />
                    <Button size="small" appearance="subtle" onClick={() => {
                      setList((l) => l.filter((t) => t.id !== cur.id));
                      setSel(list.find((t) => t.id !== cur.id)?.id ?? null);
                    }}>Odstrániť</Button>
                  </div>
                  <div className="sig-name-row">
                    <label className="sig-label" htmlFor="tpl-subj">Predmet</label>
                    <Input id="tpl-subj" className="grow" placeholder="nepovinné – vyplní sa, ak správa ešte nemá predmet" value={cur.subject ?? ''} onChange={(_, d) => patch({ subject: d.value })} />
                  </div>
                  <SignatureEditor id={cur.id} html={cur.html} onChange={(html) => patch({ html })} />
                  <div className="muted small">Šablónu vložíte pri písaní: karta Správa → Šablóny. Text sa vloží na miesto kurzora.</div>
                </>
              ) : (
                <div className="empty">Vytvorte šablónu, alebo pri písaní zvoľte Šablóny → Uložiť túto správu ako šablónu.</div>
              )}
            </section>
          </DialogContent>
          <DialogActions style={{ justifyContent: 'flex-end', alignSelf: 'flex-end' }}>
            <Button appearance="primary" onClick={() => props.onSave({ ...props.settings, templates: list })}>Uložiť</Button>
            <Button onClick={props.onClose}>Zrušiť</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
