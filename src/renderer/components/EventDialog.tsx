// Nová / upravená udalosť kalendára: čas, celodenná, miesto, účastníci (pozvánky e-mailom), opakovanie, popis.
import {
  Button, Checkbox, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input, Select, Textarea,
} from '@fluentui/react-components';
import { useState } from 'react';
import type { DavCollection } from '../../core/api';
import type { CalEvent } from '../../core/ical';

const pad = (n: number) => String(n).padStart(2, '0');
export const toLocalInput = (iso: string) => {
  const d = new Date(iso);
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
};
const toDateInput = (iso: string) => toLocalInput(iso).slice(0, 10);

const REPEAT: Array<{ v: string; label: string }> = [
  { v: '', label: 'Neopakuje sa' },
  { v: 'FREQ=DAILY', label: 'Každý deň' },
  { v: 'FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR', label: 'Každý pracovný deň' },
  { v: 'FREQ=WEEKLY', label: 'Každý týždeň' },
  { v: 'FREQ=MONTHLY', label: 'Každý mesiac' },
  { v: 'FREQ=YEARLY', label: 'Každý rok' },
];

export function EventDialog(props: {
  event: CalEvent;
  calendars: DavCollection[];
  onSave: (e: CalEvent, calendarId: string, invite: boolean) => Promise<string | null>;
  onDelete?: (e: CalEvent, notify: boolean) => Promise<string | null>;
  onClose: () => void;
}) {
  const writable = props.calendars.filter((c) => !c.readOnly && c.events !== false);
  const [e, setE] = useState<CalEvent>(props.event);
  const [calId, setCalId] = useState(props.event.calendarId ?? writable[0]?.id ?? '');
  const [people, setPeople] = useState((props.event.attendees ?? []).map((a) => (a.name ? `${a.name} <${a.email}>` : a.email)).join(', '));
  const [invite, setInvite] = useState(true);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const isNew = !props.event.href;
  const repeatKnown = REPEAT.some((r) => r.v === (e.rrule ?? '').replace(/;?(COUNT|UNTIL)=[^;]+/g, ''));

  const attendees = () =>
    people
      .split(/[,;\n]+/)
      .map((x) => x.trim())
      .filter(Boolean)
      .map((x) => {
        const m = /^(.*?)\s*<([^>]+)>$/.exec(x);
        const email = (m ? m[2]! : x).trim();
        const prev = props.event.attendees?.find((a) => a.email.toLowerCase() === email.toLowerCase());
        return { email, name: m ? m[1]!.replace(/^"|"$/g, '').trim() || undefined : undefined, partstat: prev?.partstat ?? 'NEEDS-ACTION' };
      })
      .filter((a) => a.email.includes('@'));

  async function save() {
    if (!e.summary.trim()) return setError('Zadaj názov udalosti.');
    if (Date.parse(e.end) < Date.parse(e.start)) return setError('Koniec je pred začiatkom.');
    setBusy(true);
    const err = await props.onSave({ ...e, attendees: attendees() }, calId, invite);
    setBusy(false);
    if (err) setError(err);
  }

  return (
    <Dialog open onOpenChange={(_, d) => !d.open && props.onClose()}>
      <DialogSurface style={{ maxWidth: 640 }}>
        <DialogBody>
          <DialogTitle>{isNew ? 'Nová udalosť' : 'Udalosť'}</DialogTitle>
          <DialogContent className="dialog-content">
            <Field label="Názov" required>
              <Input autoFocus value={e.summary === '(bez názvu)' ? '' : e.summary} onChange={(_, d) => setE({ ...e, summary: d.value })} placeholder="Porada, obed, termín…" />
            </Field>
            <div className="grid2">
              <Field label="Začiatok">
                {e.allDay ? (
                  <input type="date" className="date-input" value={toDateInput(e.start)} onChange={(x) => {
                    const s = new Date(`${x.target.value}T00:00`);
                    const len = Date.parse(e.end) - Date.parse(e.start);
                    setE({ ...e, start: s.toISOString(), end: new Date(s.getTime() + Math.max(len, 86400000)).toISOString() });
                  }} />
                ) : (
                  <input type="datetime-local" className="date-input" value={toLocalInput(e.start)} onChange={(x) => {
                    const s = new Date(x.target.value);
                    const len = Date.parse(e.end) - Date.parse(e.start);
                    if (!isNaN(s.getTime())) setE({ ...e, start: s.toISOString(), end: new Date(s.getTime() + len).toISOString() });
                  }} />
                )}
              </Field>
              <Field label="Koniec">
                {e.allDay ? (
                  <input type="date" className="date-input" value={toDateInput(new Date(Date.parse(e.end) - 86400000).toISOString())} onChange={(x) => {
                    const d = new Date(`${x.target.value}T00:00`);
                    setE({ ...e, end: new Date(d.getTime() + 86400000).toISOString() });
                  }} />
                ) : (
                  <input type="datetime-local" className="date-input" value={toLocalInput(e.end)} onChange={(x) => {
                    const d = new Date(x.target.value);
                    if (!isNaN(d.getTime())) setE({ ...e, end: d.toISOString() });
                  }} />
                )}
              </Field>
            </div>
            <div className="row gap">
              <Checkbox
                label="Celý deň"
                checked={e.allDay}
                onChange={(_, d) => {
                  const s = new Date(e.start);
                  if (d.checked) {
                    const day = new Date(s.getFullYear(), s.getMonth(), s.getDate());
                    setE({ ...e, allDay: true, start: day.toISOString(), end: new Date(day.getTime() + 86400000).toISOString() });
                  } else {
                    const st = new Date(s.getFullYear(), s.getMonth(), s.getDate(), 9);
                    setE({ ...e, allDay: false, start: st.toISOString(), end: new Date(st.getTime() + 3600000).toISOString() });
                  }
                }}
              />
              <Select size="small" value={repeatKnown ? (e.rrule ?? '') : 'custom'} onChange={(_, d) => setE({ ...e, rrule: d.value === 'custom' ? e.rrule : d.value || undefined })} aria-label="Opakovanie">
                {REPEAT.map((r) => <option key={r.v} value={r.v}>{r.label}</option>)}
                {!repeatKnown && <option value="custom">Vlastné opakovanie</option>}
              </Select>
              {writable.length > 1 && (
                <Select size="small" value={calId} onChange={(_, d) => setCalId(d.value)} aria-label="Kalendár" disabled={!isNew}>
                  {writable.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
                </Select>
              )}
            </div>
            <Field label="Miesto">
              <Input value={e.location ?? ''} onChange={(_, d) => setE({ ...e, location: d.value || undefined })} placeholder="zasadačka, adresa alebo odkaz na Teams" />
            </Field>
            <Field label="Účastníci" hint="E-maily oddelené čiarkou. Pozvánka príde ako bežná pozvánka (funguje v Outlooku aj Zimbre).">
              <Input value={people} onChange={(_, d) => setPeople(d.value)} placeholder="jana@firma.sk, Peter <peter@firma.sk>" />
            </Field>
            {(e.attendees?.length ?? 0) > 0 && (
              <div className="attendees">
                {e.attendees!.map((a) => (
                  <span key={a.email} className={`att att-${(a.partstat ?? 'NEEDS-ACTION').toLowerCase()}`}>
                    {a.name || a.email} · {a.partstat === 'ACCEPTED' ? 'prijal(a)' : a.partstat === 'DECLINED' ? 'odmietol(a)' : a.partstat === 'TENTATIVE' ? 'predbežne' : 'bez odpovede'}
                  </span>
                ))}
              </div>
            )}
            {people.trim() && <Checkbox label="Poslať účastníkom pozvánku e-mailom" checked={invite} onChange={(_, d) => setInvite(!!d.checked)} />}
            <Field label="Popis">
              <Textarea resize="vertical" value={e.description ?? ''} onChange={(_, d) => setE({ ...e, description: d.value || undefined })} />
            </Field>
            {e.rrule && !isNew && <div className="muted small">Zmena opakovanej udalosti sa uloží pre celú sériu.</div>}
            {error && <div className="pane-error">⚠ {error}</div>}
          </DialogContent>
          <DialogActions fluid className="dialog-actions">
            {props.onDelete && !isNew && (
              <Button appearance="subtle" disabled={busy} onClick={async () => {
                setBusy(true);
                const err = await props.onDelete!(props.event, true);
                setBusy(false);
                if (err) setError(err);
              }}>Odstrániť</Button>
            )}
            <div className="grow" />
            <Button appearance="primary" disabled={busy || !writable.length} onClick={() => void save()}>{busy ? 'Ukladám…' : 'Uložiť'}</Button>
            <Button onClick={props.onClose}>Zrušiť</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
