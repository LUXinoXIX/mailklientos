// Sprievodca importom kontaktov: vCard (.vcf) a CSV z Outlooku / Googlu / iného programu.
// Rozpozná formát, ponúkne priradenie stĺpcov, cieľový adresár a čo s duplicitami; pred importom ukáže súhrn.
import { Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Radio, RadioGroup, Select, Spinner } from '@fluentui/react-components';
import { useMemo, useState } from 'react';
import { csvToContacts, detectCsvFormat, FIELD_LABELS, guessField, parseCsv, parseVcf, planImport, type DupMode, type Field } from '../../core/contactsio';
import type { Contact } from '../../core/ical';

export interface ImportTarget {
  id: string;
  label: string;
  color: string;
  /** súčasné kontakty v cieli (na zistenie duplicít) */
  existing: Contact[];
  /** uloží jeden kontakt / všetky (lokálny adresár naraz) */
  save: (list: Contact[], progress: (done: number) => void) => Promise<{ ok: number; failed: string[] }>;
}

const FORMAT_LABEL = { outlook: 'CSV z Outlooku', google: 'CSV z Google Kontaktov', generic: 'CSV (všeobecné)' } as const;

export function ContactImport(props: { targets: ImportTarget[]; defaultTarget: string; onClose: () => void; onDone: (msg: string) => void }) {
  const api = window.mailApi;
  const [file, setFile] = useState<{ filename: string; text: string } | null>(null);
  const [err, setErr] = useState('');
  const [mapping, setMapping] = useState<Field[]>([]);
  const [target, setTarget] = useState(props.defaultTarget);
  const [dup, setDup] = useState<DupMode>('merge');
  const [busy, setBusy] = useState<{ done: number; total: number } | null>(null);

  const isVcf = !!file && (/\.(vcf|vcard)$/i.test(file.filename) || /^\s*﻿?BEGIN:VCARD/i.test(file.text));
  const csv = useMemo(() => (file && !isVcf ? parseCsv(file.text) : null), [file, isVcf]);
  const headers = csv?.[0] ?? [];
  const format = csv ? detectCsvFormat(headers) : null;

  const incoming = useMemo<Contact[]>(() => {
    if (!file) return [];
    if (isVcf) return parseVcf(file.text);
    return csv ? csvToContacts(csv.slice(1), mapping) : [];
  }, [file, isVcf, csv, mapping]);

  const tgt = props.targets.find((t) => t.id === target) ?? props.targets[0];
  const plan = useMemo(() => planImport(incoming, tgt?.existing ?? [], dup), [incoming, tgt, dup]);

  async function pick() {
    setErr('');
    const r = await api?.contactsPickImport();
    if (!r) return;
    if (!r.ok) return setErr(r.error);
    if (!r.data) return;
    setFile(r.data);
    const isV = /\.(vcf|vcard)$/i.test(r.data.filename) || /^\s*﻿?BEGIN:VCARD/i.test(r.data.text);
    if (!isV) setMapping((parseCsv(r.data.text)[0] ?? []).map(guessField));
  }

  async function run() {
    if (!tgt || !plan.save.length) return;
    setBusy({ done: 0, total: plan.save.length });
    const res = await tgt.save(plan.save, (done) => setBusy({ done, total: plan.save.length }));
    setBusy(null);
    const parts = [`Import do „${tgt.label}“: nové ${plan.added}`, `doplnené ${plan.save.length - plan.added}`];
    if (plan.skipped) parts.push(`preskočené ${plan.skipped}`);
    if (res.failed.length) parts.push(`nepodarilo sa ${res.failed.length} (${res.failed.slice(0, 3).join(', ')}${res.failed.length > 3 ? '…' : ''})`);
    props.onDone(parts.join(', ') + '.');
  }

  const mapped = new Set(mapping);
  const mapWarn = csv && !mapped.has('email') && !mapped.has('fullName') && !mapped.has('firstName') && !mapped.has('lastName');

  return (
    <Dialog open onOpenChange={(_, d) => !d.open && !busy && props.onClose()}>
      <DialogSurface style={{ maxWidth: 760 }}>
        <DialogBody>
          <DialogTitle>Import kontaktov</DialogTitle>
          <DialogContent className="dialog-content ci">
            <div className="row gap" style={{ alignItems: 'center' }}>
              <Button appearance={file ? 'secondary' : 'primary'} onClick={() => void pick()} disabled={!!busy}>{file ? 'Vybrať iný súbor…' : 'Vybrať súbor…'}</Button>
              {file ? (
                <span>
                  <b>{file.filename}</b> <span className="ci-badge">{isVcf ? 'vCard (Apple, Android, Google, Outlook)' : FORMAT_LABEL[format!]}</span> · {incoming.length} kontaktov
                </span>
              ) : (
                <span className="muted">vCard .vcf (Apple, Android, Google) alebo CSV (Outlook, Google, Excel)</span>
              )}
            </div>
            {err && <div className="pane-error">⚠ {err}</div>}

            {csv && (
              <section>
                <div className="ci-h">Stĺpce zo súboru → údaje kontaktu {format !== 'generic' && <span className="muted small">(priradené automaticky, dajú sa zmeniť)</span>}</div>
                <div className="ci-map">
                  {headers.map((h, i) => (
                    <div key={i} className="ci-map-row">
                      <span className="ci-col" title={h}>{h || `stĺpec ${i + 1}`}</span>
                      <span className="ci-ex muted small" title={csv[1]?.[i] ?? ''}>{csv[1]?.[i] ?? ''}</span>
                      <Select size="small" value={mapping[i] ?? 'skip'} onChange={(_, d) => setMapping((m) => headers.map((_, j) => (j === i ? (d.value as Field) : m[j] ?? 'skip')))}>
                        {(Object.keys(FIELD_LABELS) as Field[]).map((f) => <option key={f} value={f}>{FIELD_LABELS[f]}</option>)}
                      </Select>
                    </div>
                  ))}
                </div>
                {mapWarn && <div className="pane-error">⚠ Priraď aspoň meno alebo e-mail, inak sa nič neimportuje.</div>}
              </section>
            )}

            {file && (
              <div className="grid2 ci-opts">
                <section>
                  <div className="ci-h">Kam uložiť</div>
                  <RadioGroup value={target} onChange={(_, d) => setTarget(d.value)}>
                    {props.targets.map((t) => (
                      <Radio key={t.id} value={t.id} label={<span><span className="dot" style={{ background: t.color }} /> {t.label}{t.id !== 'local' ? ' (odošle sa na server)' : ''}</span>} />
                    ))}
                  </RadioGroup>
                </section>
                <section>
                  <div className="ci-h">Rovnaký kontakt už existuje (rovnaký e-mail)</div>
                  <RadioGroup value={dup} onChange={(_, d) => setDup(d.value as DupMode)}>
                    <Radio value="merge" label="Zlúčiť – doplniť chýbajúce údaje" />
                    <Radio value="skip" label="Preskočiť" />
                    <Radio value="add" label="Pridať ako nový" />
                  </RadioGroup>
                </section>
              </div>
            )}

            {file && (
              <div className="ci-sum">
                Výsledok: <b>nové {plan.added}</b> · zlúčené {dup === 'merge' ? plan.merged : 0} · preskočené {plan.skipped}
                {plan.noEmail > 0 && <> · bez e-mailu {plan.noEmail} (uložia sa s telefónom)</>}
              </div>
            )}
            {file && incoming.length > 0 && (
              <section>
                <div className="ci-h">Ukážka</div>
                <div className="ci-preview">
                  {incoming.slice(0, 5).map((c, i) => (
                    <div key={i} className="ci-prev-row">
                      <b>{c.name || '(bez mena)'}</b>
                      <span className="muted">{c.emails.join(', ') || 'bez e-mailu'}</span>
                      <span className="muted">{[c.phones[0], c.org].filter(Boolean).join(' · ')}</span>
                    </div>
                  ))}
                  {incoming.length > 5 && <div className="muted small">… a ďalších {incoming.length - 5}</div>}
                </div>
              </section>
            )}
          </DialogContent>
          <DialogActions>
            {busy && <span className="row gap"><Spinner size="tiny" /> {busy.done} / {busy.total}</span>}
            <Button onClick={props.onClose} disabled={!!busy}>Zrušiť</Button>
            <Button appearance="primary" disabled={!file || !plan.save.length || !!busy} onClick={() => void run()}>Importovať</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
