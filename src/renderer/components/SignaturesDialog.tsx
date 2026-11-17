// Nastavenia → Podpisy: zoznam podpisov, formátovaný editor (logo, farby, odkazy), predvolené podpisy pre každé IMAP konto,
// import z klasického Outlooku (.htm s obrázkami).
import {
  Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Input, MessageBar,
  MessageBarBody, Select,
} from '@fluentui/react-components';
import DOMPurify from 'dompurify';
import { useState } from 'react';
import type { ImapAccountPublic } from '../../core/account';
import { signatureToHtml, type AppSettings, type Signature } from '../../core/settings';
import { accountColor } from '../theme';
import { SignatureEditor } from './SignatureEditor';

const newId = () => Math.random().toString(36).slice(2, 10);

/** HTML z Outlooku/Wordu: bez skriptov, podmienených komentárov a hlavičky; obrázky ostávajú (data: URL). */
function cleanImported(html: string) {
  const body = html.replace(/<!--\[if[\s\S]*?<!\[endif\]-->/gi, '').replace(/^[\s\S]*<body[^>]*>/i, '').replace(/<\/body>[\s\S]*$/i, '');
  return DOMPurify.sanitize(body, { FORBID_TAGS: ['style', 'meta', 'link', 'xml'], ADD_DATA_URI_TAGS: ['img'] });
}

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
  const add = () => {
    const s: Signature = { id: newId(), name: 'Nový podpis', text: '', html: '' };
    setSigs((list) => [...list, s]);
    setSel(s.id);
  };
  const duplicate = () => {
    if (!cur) return;
    const s: Signature = { ...cur, id: newId(), name: `${cur.name} (kópia)` };
    setSigs((list) => [...list, s]);
    setSel(s.id);
  };

  async function importOutlook() {
    const r = await window.mailApi?.outlookSignatures();
    if (!r?.ok || !r.data.length) {
      setInfo('Na tomto počítači sa nenašli podpisy klasického Outlooku. Podpis z Outlooku na webe označte, skopírujte (Ctrl+C) a vložte do nového podpisu (Ctrl+V) – formát aj logo sa prenesú.');
      return;
    }
    const added = r.data
      .filter((x) => !sigs.some((s) => s.name === x.name))
      .map((x) => ({ id: newId(), name: x.name, text: x.text, html: x.html ? cleanImported(x.html) : undefined }));
    setSigs((list) => [...list, ...added]);
    if (added[0]) setSel(added[0].id);
    setInfo(added.length ? `Importované podpisy z Outlooku: ${added.length}.` : 'Všetky podpisy z Outlooku už v zozname sú.');
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

  const setDef = (acc: string, key: 'newMsg' | 'reply', v: string) => setDefs({ ...defs, [acc]: { ...defs[acc], [key]: v || undefined } });

  return (
    <Dialog open onOpenChange={(_, d) => !d.open && props.onClose()}>
      <DialogSurface className="sig-dialog" style={{ maxWidth: 1180, width: '94vw' }}>
        <DialogBody style={{ display: 'flex', flexDirection: 'column', maxHeight: '88vh' }}>
          <DialogTitle>Podpisy</DialogTitle>
          <DialogContent className="sig-grid">
            <aside className="sig-list">
              <div className="sig-list-items">
                {sigs.map((s) => (
                  <button key={s.id} className={`folder ${s.id === sel ? 'active' : ''}`} onClick={() => setSel(s.id)}>
                    <span className="grow">{s.name || '(bez názvu)'}</span>
                  </button>
                ))}
                {!sigs.length && <div className="muted small pad">Zatiaľ žiadny podpis.</div>}
              </div>
              <div className="sig-list-actions">
                <Button size="small" appearance="primary" onClick={add}>+ Nový podpis</Button>
                <Button size="small" onClick={() => void importOutlook()}>Import z Outlooku</Button>
              </div>
            </aside>

            <section className="sig-main">
              {info && (
                <MessageBar intent="info">
                  <MessageBarBody>{info}</MessageBarBody>
                </MessageBar>
              )}
              {cur ? (
                <>
                  <div className="sig-name-row">
                    <label className="sig-label" htmlFor="sig-name">Názov</label>
                    <Input id="sig-name" className="grow" value={cur.name} onChange={(_, d) => patch({ name: d.value })} />
                    <Button size="small" appearance="subtle" onClick={duplicate}>Duplikovať</Button>
                    <Button size="small" appearance="subtle" onClick={remove}>Odstrániť</Button>
                  </div>
                  <SignatureEditor id={cur.id} html={signatureToHtml(cur)} onChange={(html, text) => patch({ html, text })} />
                  <div className="muted small">Tip: logo vložíte tlačidlom „Obrázok“, pretiahnutím alebo Ctrl+V; veľkosť zmeníte ťahaním za roh. Podpis z Outlooku na webe skopírujte a vložte.</div>
                </>
              ) : (
                <div className="empty">Vytvorte podpis tlačidlom „+ Nový podpis“ alebo ho importujte z Outlooku.</div>
              )}

              <div className="sig-defaults">
                <div className="sig-defaults-title">Predvolené podpisy podľa konta</div>
                {props.accounts.length ? (
                  <table className="sig-def-table">
                    <thead>
                      <tr><th>Konto</th><th>Nové správy</th><th>Odpovede a preposlania</th></tr>
                    </thead>
                    <tbody>
                      {props.accounts.map((a) => (
                        <tr key={a.id}>
                          <td><span className="dot" style={{ background: accountColor(a.color) }} /> {a.name} <span className="muted small">{a.email}</span></td>
                          {(['newMsg', 'reply'] as const).map((k) => (
                            <td key={k}>
                              <Select size="small" value={defs[a.id]?.[k] ?? ''} onChange={(_, d) => setDef(a.id, k, d.value)}>
                                <option value="">(žiadny)</option>
                                {sigs.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                              </Select>
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                ) : (
                  <div className="muted small">Pridajte IMAP konto. Kontá Microsoft 365 majú podpisy priamo v Outlooku na webe.</div>
                )}
              </div>
            </section>
          </DialogContent>
          <DialogActions style={{ justifyContent: 'flex-end', alignSelf: 'flex-end' }}>
            <Button appearance="primary" onClick={() => props.onSave({ ...props.settings, signatures: sigs, signatureDefaults: defs })}>Uložiť</Button>
            <Button onClick={props.onClose}>Zrušiť</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
