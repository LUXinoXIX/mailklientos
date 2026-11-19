// Karta „Kalendáre (ICS)“ v pravom paneli: viac kalendárov z odkazu (iCal/webcal) pod sebou,
// poradie a počet zobrazených udalostí pre každý kalendár nastaviteľné.
import { Button, Input, Spinner } from '@fluentui/react-components';
import { useEffect, useState } from 'react';
import type { CalEvent } from '../../core/ical';
import { findMeeting, MEETING_LABEL } from '../../core/meeting';
import type { IcsFeed } from '../../core/settings';

const COLORS = ['#2F6F5E', '#b45309', '#6d28d9', '#1d4ed8', '#be123c', '#0e7490', '#3f6212', '#4B5563'];
const COUNTS = [3, 5, 8, 10, 15, 20];
const dayKey = (d: Date) => `${d.getFullYear()}-${d.getMonth()}-${d.getDate()}`;

function range(e: CalEvent) {
  const st = new Date(e.start);
  const en = new Date(e.end);
  const date = (d: Date) => d.toLocaleDateString('sk-SK', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
  const time = (d: Date) => d.toLocaleTimeString('sk-SK', { hour: '2-digit', minute: '2-digit' });
  if (e.allDay) {
    const last = new Date(en.getTime() - 86400000);
    return dayKey(last) === dayKey(st) || last < st ? `${date(st)} · celý deň` : `${date(st)} – ${date(last)}`;
  }
  return dayKey(st) === dayKey(en) ? `${date(st)}, ${time(st)} – ${time(en)}` : `${date(st)} ${time(st)} – ${date(en)} ${time(en)}`;
}

/** Popis bez dlhých technických blokov (pozvánky Teams obsahujú čiary a ID schôdze). */
const cleanDesc = (s: string) =>
  s
    .replace(/_{5,}|-{5,}|={5,}/g, '')
    .replace(/<[^>]+>/g, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

function when(e: CalEvent, now: Date) {
  const st = new Date(e.start);
  const tomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
  const day = dayKey(st) === dayKey(now) ? 'dnes' : dayKey(st) === dayKey(tomorrow) ? 'zajtra' : st.toLocaleDateString('sk-SK', { weekday: 'short', day: 'numeric', month: 'numeric' });
  return e.allDay ? `${day} · celý deň` : `${day} ${st.toLocaleTimeString('sk-SK', { hour: '2-digit', minute: '2-digit' })}`;
}

export function IcsPane(props: { feeds: IcsFeed[]; onFeeds: (f: IcsFeed[]) => void; onStatus: (s: string) => void }) {
  const api = window.mailApi;
  const [data, setData] = useState<Record<string, { list?: CalEvent[]; error?: string }>>({});
  const [edit, setEdit] = useState(props.feeds.length === 0);
  const [form, setForm] = useState({ url: '', name: '', count: 5 });
  const [tick, setTick] = useState(0);
  const [open, setOpen] = useState<string | null>(null);
  const now = new Date();
  const key = props.feeds.map((f) => `${f.url}|${f.count}`).join(',');

  useEffect(() => {
    if (!api) return;
    let alive = true;
    for (const f of props.feeds) {
      setData((d) => ({ ...d, [f.id]: d[f.id]?.list ? d[f.id]! : {} }));
      void api.icsUpcoming(f.url, f.count).then((r) => alive && setData((d) => ({ ...d, [f.id]: r.ok ? { list: r.data } : { error: r.error } })));
    }
    return () => {
      alive = false;
    };
  }, [api, key, tick]); // eslint-disable-line react-hooks/exhaustive-deps

  const move = (i: number, d: number) => {
    const f = [...props.feeds];
    const j = i + d;
    if (j < 0 || j >= f.length) return;
    [f[i], f[j]] = [f[j]!, f[i]!];
    props.onFeeds(f);
  };
  const add = () => {
    const url = form.url.trim();
    if (!/^(https?|webcals?):\/\/\S+$/i.test(url)) return props.onStatus('Zadaj adresu kalendára (https://… alebo webcal://…).');
    let name = form.name.trim();
    try {
      name ||= new URL(url.replace(/^webcals?:/i, 'https:')).hostname;
    } catch {
      name ||= 'Kalendár';
    }
    props.onFeeds([...props.feeds, { id: Math.random().toString(36).slice(2, 10), name, url, count: form.count, color: COLORS[props.feeds.length % COLORS.length] }]);
    setForm({ url: '', name: '', count: 5 });
  };

  return (
    <div className="ics-pane">
      <div className="ics-tools">
        <span className="muted small grow">{props.feeds.length ? `${props.feeds.length} ${props.feeds.length === 1 ? 'kalendár' : props.feeds.length < 5 ? 'kalendáre' : 'kalendárov'}` : 'Kalendáre z odkazu'}</span>
        <Button size="small" appearance="subtle" onClick={() => setTick((t) => t + 1)} disabled={!props.feeds.length} title="Obnoviť">↻</Button>
        <Button size="small" appearance={edit ? 'primary' : 'subtle'} onClick={() => setEdit((x) => !x)}>{edit ? 'Hotovo' : 'Upraviť'}</Button>
      </div>

      {edit && (
        <div className="ics-edit">
          {props.feeds.map((f, i) => (
            <div key={f.id} className="ics-edit-row">
              <input type="color" className="ics-color" value={f.color ?? COLORS[i % COLORS.length]} onChange={(ev) => props.onFeeds(props.feeds.map((x) => (x.id === f.id ? { ...x, color: ev.target.value } : x)))} aria-label={`Farba ${f.name}`} />
              <input className="ics-name grow" value={f.name} onChange={(ev) => props.onFeeds(props.feeds.map((x) => (x.id === f.id ? { ...x, name: ev.target.value } : x)))} aria-label="Názov" title={f.url} />
              <select value={f.count} onChange={(ev) => props.onFeeds(props.feeds.map((x) => (x.id === f.id ? { ...x, count: Number(ev.target.value) } : x)))} aria-label="Počet udalostí" title="Počet zobrazených udalostí">
                {COUNTS.map((n) => <option key={n} value={n}>{n}</option>)}
              </select>
              <button className="ics-mini" disabled={!i} onClick={() => move(i, -1)} aria-label="Vyššie" title="Vyššie">↑</button>
              <button className="ics-mini" disabled={i === props.feeds.length - 1} onClick={() => move(i, 1)} aria-label="Nižšie" title="Nižšie">↓</button>
              <button className="ics-mini" onClick={() => props.onFeeds(props.feeds.filter((x) => x.id !== f.id))} aria-label={`Odstrániť ${f.name}`} title="Odstrániť">×</button>
            </div>
          ))}
          <div className="ics-add">
            <Input size="small" placeholder="https://… alebo webcal://…" value={form.url} onChange={(_, d) => setForm({ ...form, url: d.value })} />
            <div className="ics-edit-row">
              <Input size="small" className="grow" placeholder="Názov (napr. Sviatky)" value={form.name} onChange={(_, d) => setForm({ ...form, name: d.value })} />
              <select value={form.count} onChange={(ev) => setForm({ ...form, count: Number(ev.target.value) })} aria-label="Počet udalostí">
                {COUNTS.map((n) => <option key={n} value={n}>{n} udal.</option>)}
              </select>
              <Button size="small" appearance="primary" onClick={add} disabled={!form.url.trim()}>Pridať</Button>
            </div>
            <span className="muted small">Verejná adresa kalendára vo formáte iCal (Google, iCloud, Outlook, Zimbra, sviatky, rozvrh). Obnovuje sa každú hodinu.</span>
          </div>
        </div>
      )}

      <div className="ics-list">
        {!props.feeds.length && !edit && <div className="empty">Pridaj kalendár tlačidlom Upraviť.</div>}
        {props.feeds.map((f, i) => {
          const d = data[f.id];
          const color = f.color ?? COLORS[i % COLORS.length];
          return (
            <section key={f.id} className="ics-sec">
              <div className="ics-sec-h"><span className="dot" style={{ background: color }} />{f.name}</div>
              {(!d || (!d.list && !d.error)) && <Spinner size="extra-tiny" />}
              {d?.error && <div className="muted small">⚠ {d.error}</div>}
              {d?.list && !d.list.length && <div className="muted small">Žiadne nadchádzajúce udalosti.</div>}
              {d?.list?.map((e, k) => {
                const m = findMeeting(e);
                const live = !e.allDay && Date.parse(e.start) <= now.getTime() && Date.parse(e.end) > now.getTime();
                const id = `${f.id}:${e.uid}:${e.start}:${k}`;
                const isOpen = open === id;
                const desc = isOpen && e.description ? cleanDesc(e.description) : '';
                return (
                  <div
                    key={k}
                    className={`ics-ev ${live ? 'live' : ''} ${isOpen ? 'open' : ''}`}
                    style={{ borderLeftColor: color }}
                    role="button"
                    tabIndex={0}
                    aria-expanded={isOpen}
                    title={isOpen ? undefined : 'Kliknutím zobrazíš podrobnosti'}
                    onClick={(ev) => {
                      if ((ev.target as HTMLElement).closest('button, a')) return;
                      setOpen(isOpen ? null : id);
                    }}
                    onKeyDown={(ev) => {
                      if (ev.key === 'Enter' || ev.key === ' ') {
                        ev.preventDefault();
                        setOpen(isOpen ? null : id);
                      }
                    }}
                  >
                    <span className="ics-when">{when(e, now)}</span>
                    <span className="ics-sum">{e.summary}</span>
                    {(e.location || m) && (
                      <span className="ics-loc muted small">
                        {e.location && !m ? e.location : ''}
                        {m && (
                          <button className="link" onClick={() => (m.kind === 'jitsi' ? void api?.meetJitsi(m.url) : void api?.openExternal(m.url))}>
                            Pripojiť sa – {MEETING_LABEL[m.kind]}
                          </button>
                        )}
                      </span>
                    )}
                    {isOpen && (
                      <div className="ics-detail">
                        <div>🕒 {range(e)}</div>
                        {e.location && <div>📍 {e.location}</div>}
                        {e.organizer && <div>👤 {e.organizer.name || e.organizer.email}</div>}
                        {!!e.attendees?.length && (
                          <div className="muted">
                            👥 {e.attendees.slice(0, 5).map((a) => a.name || a.email).join(', ')}
                            {e.attendees.length > 5 ? ` a ďalších ${e.attendees.length - 5}` : ''}
                          </div>
                        )}
                        {desc && <div className="ics-desc">{desc.length > 1500 ? desc.slice(0, 1500) + '…' : desc}</div>}
                        {!e.location && !desc && !e.organizer && <div className="muted">Bez ďalších podrobností.</div>}
                      </div>
                    )}
                  </div>
                );
              })}
            </section>
          );
        })}
      </div>
    </div>
  );
}
