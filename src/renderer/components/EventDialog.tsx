// Nová / upravená udalosť kalendára: čas, celodenná, miesto, účastníci (pozvánky e-mailom), opakovanie, popis.
import {
  Button, Checkbox, Spinner, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input, Select,
} from '@fluentui/react-components';
import { useEffect, useState } from 'react';
import type { DavCollection, FreeBusy, ZResource } from '../../core/api';
import type { ImapAccountPublic } from '../../core/account';
import type { CalAttendee, CalEvent } from '../../core/ical';
import { RichTextField } from './RichTextField';
import { applyMeeting, findMeeting, stripMeetingBlock, type MeetingSettings } from '../../core/meeting';
import { MeetingPicker, type MeetOutlook } from './MeetingPicker';

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

/** Asistent plánovania: obsadenosť účastníkov a zasadačiek v deň udalosti (Zimbra). */
function Scheduler(props: { accountId: string; me: string; emails: string[]; start: string; end: string; onPick: (start: Date) => void }) {
  const api = window.mailApi!;
  const day = new Date(props.start);
  const from = new Date(day.getFullYear(), day.getMonth(), day.getDate(), 7).getTime();
  const to = from + 12 * 3600_000;
  const [fb, setFb] = useState<FreeBusy[] | null>(null);
  const [err, setErr] = useState<string | null>(null);
  const key = [props.me, ...props.emails].join(',');
  useEffect(() => {
    let alive = true;
    setFb(null);
    void api.zimbraFreeBusy(props.accountId, [props.me, ...props.emails], from, to).then((r) => {
      if (!alive) return;
      if (r.ok) setFb(r.data);
      else setErr(r.error);
    });
    return () => {
      alive = false;
    };
  }, [key, from]); // eslint-disable-line react-hooks/exhaustive-deps
  const len = Date.parse(props.end) - Date.parse(props.start);
  const busyAt = (t: number) => (fb ?? []).some((u) => u.slots.some((x) => x.kind !== 'n' && x.s < t + len && x.e > t));
  const firstFree = () => {
    for (let t = Math.max(from, Math.ceil(Date.now() / 1800000) * 1800000); t + len <= to; t += 1800000) if (!busyAt(t)) return t;
    return null;
  };
  const pct = (t: number) => `${((t - from) / (to - from)) * 100}%`;
  const evS = Date.parse(props.start);
  return (
    <div className="sched">
      <div className="sched-head">
        <b>Asistent plánovania</b>
        <span className="muted small">{day.toLocaleDateString('sk-SK', { weekday: 'long', day: 'numeric', month: 'numeric' })}, 7:00–19:00</span>
        <span className="grow" />
        <Button size="small" disabled={!fb} onClick={() => { const t = firstFree(); if (t) props.onPick(new Date(t)); }}>Nájsť voľný čas</Button>
      </div>
      {err && <div className="muted small">⚠ {err}</div>}
      {!fb && !err && <Spinner size="extra-tiny" />}
      <div className="sched-hours">{Array.from({ length: 12 }, (_, i) => <span key={i}>{7 + i}</span>)}</div>
      {fb?.map((u) => (
        <div key={u.email} className="sched-row">
          <span className="sched-name" title={u.email}>{u.email === props.me ? 'Ja' : u.email}</span>
          <div className="sched-bar">
            {u.slots.filter((x) => x.e > from && x.s < to).map((x, i) => (
              <span key={i} className={`sched-slot ${x.kind}`} style={{ left: pct(Math.max(x.s, from)), width: `calc(${pct(Math.min(x.e, to))} - ${pct(Math.max(x.s, from))})` }} />
            ))}
            {evS >= from && evS < to && <span className="sched-event" style={{ left: pct(evS), width: `calc(${pct(Math.min(evS + len, to))} - ${pct(evS)})` }} />}
          </div>
        </div>
      ))}
      <div className="muted small">Tmavé = obsadený, šrafované = predbežne, fialové = mimo kancelárie.</div>
    </div>
  );
}

export function EventDialog(props: {
  event: CalEvent;
  calendars: DavCollection[];
  account?: ImapAccountPublic;
  zimbra?: boolean;
  onSave: (e: CalEvent, calendarId: string, invite: boolean) => Promise<string | null>;
  onDelete?: (e: CalEvent, notify: boolean) => Promise<string | null>;
  onClose: () => void;
  /** online schôdze (Teams cez Outlook konto, Jitsi) */
  meet?: { settings: MeetingSettings; outlooks: MeetOutlook[]; onSettings: (s: MeetingSettings) => void };
}) {
  const meetSettings = props.meet?.settings ?? {};
  const [meetUrl, setMeetUrl] = useState<string | null>(() => findMeeting(props.event, meetSettings.jitsiServer)?.url ?? null);
  const writable = props.calendars.filter((c) => !c.readOnly && c.events !== false);
  // blok s odkazom sa v popise neukazuje (pri uložení sa doplní nanovo)
  const [e, setE] = useState<CalEvent>(() => ({ ...props.event, description: props.event.description ? stripMeetingBlock(props.event.description) || undefined : undefined }));
  const [calId, setCalId] = useState(props.event.calendarId ?? writable[0]?.id ?? '');
  const [people, setPeople] = useState((props.event.attendees ?? []).filter((a) => a.cutype !== 'ROOM' && a.cutype !== 'RESOURCE').map((a) => (a.name ? `${a.name} <${a.email}>` : a.email)).join(', '));
  const [invite, setInvite] = useState(true);
  const [busy, setBusy] = useState(false);
  const [rooms, setRooms] = useState<ZResource[]>([]);
  const [showSched, setShowSched] = useState(false);
  const [roomList, setRoomList] = useState<CalAttendee[]>((props.event.attendees ?? []).filter((a) => a.cutype === 'ROOM' || a.cutype === 'RESOURCE'));
  useEffect(() => {
    if (!props.zimbra || !props.account) return;
    void window.mailApi?.zimbraResources(props.account.id).then((r) => r.ok && setRooms(r.data));
  }, [props.zimbra, props.account]);
  const [error, setError] = useState<string | null>(null);
  const isNew = !props.event.href;
  const repeatKnown = REPEAT.some((r) => r.v === (e.rrule ?? '').replace(/;?(COUNT|UNTIL)=[^;]+/g, ''));

  const attendees = (): CalAttendee[] => [
    ...people
      .split(/[,;\n]+/)
      .map((x) => x.trim())
      .filter(Boolean)
      .map((x) => {
        const m = /^(.*?)\s*<([^>]+)>$/.exec(x);
        const email = (m ? m[2]! : x).trim();
        const prev = props.event.attendees?.find((a) => a.email.toLowerCase() === email.toLowerCase());
        return { email, name: m ? m[1]!.replace(/^"|"$/g, '').trim() || undefined : undefined, partstat: prev?.partstat ?? 'NEEDS-ACTION' };
      })
      .filter((a) => a.email.includes('@')),
    ...roomList,
  ];

  async function save() {
    if (!e.summary.trim()) return setError('Zadaj názov udalosti.');
    if (Date.parse(e.end) < Date.parse(e.start)) return setError('Koniec je pred začiatkom.');
    setBusy(true);
    const withMeet = applyMeeting({ ...e, onlineUrl: findMeeting(props.event, meetSettings.jitsiServer)?.url }, meetUrl, meetSettings.jitsiServer);
    // formátovaný popis: odkaz na online stretnutie pridať aj do HTML (čistý text ho už má z applyMeeting)
    if (withMeet.descriptionHtml && withMeet.onlineUrl && !withMeet.descriptionHtml.includes(withMeet.onlineUrl)) {
      const u = withMeet.onlineUrl.replace(/"/g, '%22');
      withMeet.descriptionHtml += `<p>${withMeet.location ? `${withMeet.location} – ` : ''}pripojiť sa: <a href="${u}">${u}</a></p>`;
    }
    const err = await props.onSave({ ...withMeet, attendees: attendees() }, calId, invite);
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
            <MeetingPicker
              url={meetUrl}
              title={e.summary === '(bez názvu)' ? '' : e.summary}
              isNew={isNew}
              settings={meetSettings}
              outlooks={props.meet?.outlooks ?? []}
              onUrl={setMeetUrl}
              onSettings={(s) => props.meet?.onSettings(s)}
            />
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
            {props.zimbra && props.account && (
              <>
                {rooms.length > 0 && (
                  <div className="row gap">
                    <Select
                      size="small"
                      value=""
                      onChange={(_, d) => {
                        const r = rooms.find((x) => x.email === d.value);
                        if (!r || roomList.some((x) => x.email === r.email)) return;
                        setRoomList([...roomList, { email: r.email, name: r.name, cutype: r.kind === 'Location' ? 'ROOM' : 'RESOURCE', partstat: 'NEEDS-ACTION', role: 'NON-PARTICIPANT' }]);
                        if (r.kind === 'Location' && !e.location) setE({ ...e, location: r.location || r.name });
                      }}
                      aria-label="Pridať zasadačku alebo zariadenie"
                    >
                      <option value="">+ Zasadačka / zariadenie…</option>
                      {rooms.map((r) => (
                        <option key={r.email} value={r.email}>{r.name}{r.capacity ? ` (${r.capacity} miest)` : ''}{r.kind === 'Equipment' ? ' – zariadenie' : ''}</option>
                      ))}
                    </Select>
                    {roomList.map((r) => (
                      <span key={r.email} className="att">
                        {r.name || r.email}{' '}
                        <button className="cat-x" aria-label="Odobrať" onClick={() => setRoomList(roomList.filter((x) => x.email !== r.email))}>×</button>
                      </span>
                    ))}
                  </div>
                )}
                {!e.allDay && (attendees().length > 0) && (
                  <Button size="small" appearance="subtle" onClick={() => setShowSched(!showSched)}>{showSched ? 'Skryť asistenta plánovania' : 'Asistent plánovania (kedy majú všetci čas)'}</Button>
                )}
                {showSched && !e.allDay && (
                  <Scheduler
                    accountId={props.account.id}
                    me={props.account.email}
                    emails={attendees().map((a) => a.email)}
                    start={e.start}
                    end={e.end}
                    onPick={(d) => {
                      const len = Date.parse(e.end) - Date.parse(e.start);
                      setE({ ...e, start: d.toISOString(), end: new Date(d.getTime() + len).toISOString() });
                    }}
                  />
                )}
              </>
            )}
            {people.trim() && <Checkbox label="Poslať účastníkom pozvánku e-mailom" checked={invite} onChange={(_, d) => setInvite(!!d.checked)} />}
            <Field label="Popis">
              {/* formátovaný popis: ide ako X-ALT-DESC (HTML) a zároveň čistý text do DESCRIPTION – Zimbra aj Outlook ho zobrazia formátovaný */}
              <RichTextField
                html={e.descriptionHtml ?? (e.description ? e.description.split('\n').map((l) => `<p>${l.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</p>`).join('') : '')}
                placeholder="Program stretnutia, poznámky…"
                onChange={(html, text) => setE({ ...e, description: text || undefined, descriptionHtml: html || undefined })}
              />
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
