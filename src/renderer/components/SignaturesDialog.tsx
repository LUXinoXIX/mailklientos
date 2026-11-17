// Nastavenia → Podpisy: zoznam podpisov, úprava textu, predvolené podpisy pre každé IMAP konto, import z Outlooku.
import {
  Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input, MessageBar,
  MessageBarBody, Select, Textarea,
} from '@fluentui/react-components';
import { useState } from 'react';
import type { ImapAccountPublic } from '../../core/account';
import type { AppSettings, Signature } from '../../core/settings';
import { accountColor } from '../theme';

const newId = () => Math.random().toString(36).slice(2, 10);

export function SignaturesDialog(props: {
  settings: AppSettings;
  accounts: ImapAccountPublic[];
  onSave: (s: AppSettings) => void;
  onClose: () => void;
}) {
  const [sigs, setSigs] = useState<Signature[]>(props.settings.signatures);
  const [defs, setDefs] = useState(props.settings.signatureDefaults);
  const [sel, setSel] = useState<string | null>(props.settings.signatures[0]?.id ?? null);
  const [info, setInfo] = useState<string | null>(null);
  const cur = sigs.find((s) => s.id === sel);

  const patch = (p: Partial<Signature>) => setSigs((list) => list.map((s) => (s.id === sel ? { ...s, ...p } : s)));
  const add = (name = 'Nový podpis', text = '') => {
    const s = { id: newId(), name, text };
    setSigs((list) => [...list, s]);
    setSel(s.id);
    return s;
  };

  async function importOutlook() {
    const r = await window.mailApi?.outlookSignatures();
    if (!r?.ok || !r.data.length) {
      setInfo('Na tomto počítači sa nenašli podpisy klasického Outlooku. Podpis z Outlooku na webe skopíruj a vlož do nového podpisu.');
      return;
    }
    const added = r.data.filter((x) => !sigs.some((s) => s.name === x.name)).map((x) => ({ id: newId(), name: x.name, text: x.text }));
    setSigs((list) => [...list, ...added]);
    if (added[0]) setSel(added[0].id);
    setInfo(`Importované podpisy: ${added.length}.`);
  }

  function remove() {
    if (!cur) return;
    setSigs((list) => list.filter((s) => s.id !== cur.id));
    setDefs((d) =>
      Object.fromEntries(
        Object.entries(d).map(([k, v]) => [k, { newMsg: v.newMsg === cur.id ? undefined : v.newMsg, reply: v.reply === cur.id ? undefined : v.reply }]),
      ),
    );
    setSel(sigs.find((s) => s.id !== cur.id)?.id ?? null);
  }

  return (
    <Dialog open onOpenChange={(_, d) => !d.open && props.onClose()}>
      <DialogSurface className="sig-dialog">
        <DialogBody>
          <DialogTitle>Podpisy</DialogTitle>
          <DialogContent className="sig-grid">
            <div className="sig-list">
              {sigs.map((s) => (
                <button key={s.id} className={`folder ${s.id === sel ? 'active' : ''}`} onClick={() => setSel(s.id)}>
                  <span className="grow">{s.name || '(bez názvu)'}</span>
                </button>
              ))}
              {!sigs.length && <div className="muted small pad">Zatiaľ žiadny podpis.</div>}
              <div className="sig-list-actions">
                <Button size="small" appearance="primary" onClick={() => add()}>+ Nový podpis</Button>
                <Button size="small" onClick={() => void importOutlook()}>Import z Outlooku</Button>
              </div>
            </div>
            <div className="sig-edit">
              {cur ? (
                <>
                  <Field label="Názov podpisu">
                    <Input value={cur.name} onChange={(_, d) => patch({ name: d.value })} />
                  </Field>
                  <Field label="Text podpisu">
                    <Textarea className="sig-text" resize="vertical" value={cur.text} onChange={(_, d) => patch({ text: d.value })} placeholder={'S pozdravom\nJán Novák\nvedúci oddelenia · Firma s.r.o.\n+421 900 000 000'} />
                  </Field>
                  <div>
                    <Button size="small" appearance="subtle" onClick={remove}>Odstrániť podpis</Button>
                  </div>
                </>
              ) : (
                <div className="empty">Vytvor podpis tlačidlom „+ Nový podpis“ alebo ho importuj z Outlooku.</div>
              )}
              {info && <MessageBar intent="info"><MessageBarBody>{info}</MessageBarBody></MessageBar>}
            </div>
            <div className="sig-defaults">
              <div className="sig-defaults-title">Predvolené podpisy podľa konta</div>
              {props.accounts.map((a) => (
                <div key={a.id} className="sig-acc">
                  <div className="row gap0"><span className="dot" style={{ background: accountColor(a.color) }} /> <b>{a.name}</b></div>
                  <div className="grid2">
                    <Field label="Nové správy" size="small">
                      <Select size="small" value={defs[a.id]?.newMsg ?? ''} onChange={(_, d) => setDefs({ ...defs, [a.id]: { ...defs[a.id], newMsg: d.value || undefined } })}>
                        <option value="">(žiadny)</option>
                        {sigs.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                      </Select>
                    </Field>
                    <Field label="Odpovede" size="small">
                      <Select size="small" value={defs[a.id]?.reply ?? ''} onChange={(_, d) => setDefs({ ...defs, [a.id]: { ...defs[a.id], reply: d.value || undefined } })}>
                        <option value="">(žiadny)</option>
                        {sigs.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                      </Select>
                    </Field>
                  </div>
                </div>
              ))}
              {!props.accounts.length && <div className="muted small">Pridaj IMAP konto. Kontá Microsoft 365 majú podpisy priamo v Outlooku na webe.</div>}
            </div>
          </DialogContent>
          <DialogActions>
            <Button appearance="primary" onClick={() => props.onSave({ ...props.settings, signatures: sigs, signatureDefaults: defs })}>Uložiť</Button>
            <Button onClick={props.onClose}>Zrušiť</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
