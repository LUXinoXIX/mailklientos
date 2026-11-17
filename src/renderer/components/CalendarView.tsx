// Kalendár IMAP konta (CalDAV): deň, pracovný týždeň, týždeň, mesiac; viac kalendárov s farbami; nová udalosť kliknutím do mriežky.
import { Button, Checkbox, Spinner, Tab, TabList } from '@fluentui/react-components';
import { useEffect, useMemo, useRef, useState } from 'react';
import type { ImapAccountPublic } from '../../core/account';
import type { DavCollection } from '../../core/api';
import type { CalEvent } from '../../core/ical';
import { accountColor } from '../theme';
import { EventDialog } from './EventDialog';

type View = 'day' | 'workweek' | 'week' | 'month';
const HOUR = 44; // výška hodiny v px
const DAY_MS = 86400000;

const startOfDay = (d: Date) => new Date(d.getFullYear(), d.getMonth(), d.getDate());
const addDays = (d: Date, n: number) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
const mondayOf = (d: Date) => addDays(startOfDay(d), -((d.getDay() + 6) % 7));
const sameDay = (a: Date, b: Date) => a.toDateString() === b.toDateString();
const fmtTime = (iso: string) => new Date(iso).toLocaleTimeString('sk-SK', { hour: '2-digit', minute: '2-digit' });

export function CalendarView(props: { account: ImapAccountPublic; onStatus: (s: string) => void; onSettings: () => void }) {
  const api = window.mailApi;
  const [view, setView] = useState<View>(() => (localStorage.getItem('calView') as View) || 'workweek');
  const [anchor, setAnchor] = useState(() => startOfDay(new Date()));
  const [cals, setCals] = useState<DavCollection[] | null>(null);
  const [hidden, setHidden] = useState<string[]>(() => JSON.parse(localStorage.getItem(`calHidden:${props.account.id}`) ?? '[]') as string[]);
  const [events, setEvents] = useState<CalEvent[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [edit, setEdit] = useState<CalEvent | null>(null);
  const [tick, setTick] = useState(0);
  const scroller = useRef<HTMLDivElement>(null);

  useEffect(() => localStorage.setItem('calView', view), [view]);
  useEffect(() => localStorage.setItem(`calHidden:${props.account.id}`, JSON.stringify(hidden)), [hidden, props.account.id]);

  // rozsah zobrazenia
  const range = useMemo(() => {
    if (view === 'day') return { from: anchor, days: 1 };
    if (view === 'workweek') return { from: mondayOf(anchor), days: 5 };
    if (view === 'week') return { from: mondayOf(anchor), days: 7 };
    const first = new Date(anchor.getFullYear(), anchor.getMonth(), 1);
    return { from: mondayOf(first), days: 42 };
  }, [view, anchor]);
  const days = useMemo(() => Array.from({ length: range.days }, (_, i) => addDays(range.from, i)), [range]);

  useEffect(() => {
    if (!api) return;
    setCals(null);
    setError(null);
    void api.davCollections(props.account.id).then((r) => {
      if (r.ok) setCals(r.data.filter((c) => c.kind === 'calendar' && c.events !== false));
      else {
        setCals([]);
        setError(`${r.error}${r.hint ? ` ${r.hint}` : ''}`);
      }
    });
  }, [api, props.account.id]);

  useEffect(() => {
    if (!api || !cals?.length) return;
    let alive = true;
    setLoading(true);
    const ids = cals.filter((c) => !hidden.includes(c.id)).map((c) => c.id);
    const to = addDays(range.from, range.days);
    void api.calEvents(props.account.id, ids, range.from.toISOString(), to.toISOString()).then((r) => {
      if (!alive) return;
      setLoading(false);
      if (r.ok) {
        setEvents(r.data);
        setError(null);
      } else setError(r.error);
    });
    return () => {
      alive = false;
    };
  }, [api, cals, hidden, range, props.account.id, tick]);

  // ráno na začiatok pracovného dňa
  useEffect(() => {
    if (scroller.current) scroller.current.scrollTop = 7 * HOUR;
  }, [view, cals]);

  const colorOf = (e: CalEvent) => e.color ?? cals?.find((c) => c.id === e.calendarId)?.color ?? accountColor(props.account.color);
  const move = (dir: number) => {
    if (view === 'month') setAnchor(new Date(anchor.getFullYear(), anchor.getMonth() + dir, 1));
    else setAnchor(addDays(anchor, dir * (view === 'day' ? 1 : 7)));
  };
  const title =
    view === 'month'
      ? anchor.toLocaleDateString('sk-SK', { month: 'long', year: 'numeric' })
      : view === 'day'
        ? anchor.toLocaleDateString('sk-SK', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })
        : `${days[0]!.toLocaleDateString('sk-SK', { day: 'numeric', month: 'numeric' })} – ${days[days.length - 1]!.toLocaleDateString('sk-SK', { day: 'numeric', month: 'numeric', year: 'numeric' })}`;

  const newAt = (d: Date, allDay = false) => {
    const start = allDay ? startOfDay(d) : d;
    setEdit({ uid: '', summary: '', start: start.toISOString(), end: new Date(start.getTime() + (allDay ? DAY_MS : 3600000)).toISOString(), allDay });
  };

  async function saveEvent(e: CalEvent, calendarId: string, invite: boolean) {
    const r = await api!.calSave(props.account.id, calendarId, e, invite);
    if (!r.ok) return `${r.error}${r.hint ? ` ${r.hint}` : ''}`;
    setEdit(null);
    setTick((t) => t + 1);
    props.onStatus(invite && e.attendees?.length ? 'Udalosť uložená, pozvánky odoslané.' : 'Udalosť uložená.');
    return null;
  }
  async function deleteEvent(e: CalEvent, notify: boolean) {
    const r = await api!.calDelete(props.account.id, e, notify);
    if (!r.ok) return r.error;
    setEdit(null);
    setTick((t) => t + 1);
    props.onStatus('Udalosť odstránená.');
    return null;
  }

  const dayEvents = (d: Date) => {
    const s = d.getTime();
    const e = s + DAY_MS;
    return events.filter((x) => Date.parse(x.start) < e && Date.parse(x.end) > s);
  };

  /** Rozloženie prekrývajúcich sa udalostí do stĺpcov. */
  const layout = (list: CalEvent[]) => {
    const sorted = [...list].sort((a, b) => Date.parse(a.start) - Date.parse(b.start));
    const cols: number[] = [];
    const out = sorted.map((ev) => {
      const s = Date.parse(ev.start);
      let c = cols.findIndex((end) => end <= s);
      if (c < 0) c = cols.length;
      cols[c] = Date.parse(ev.end);
      return { ev, col: c };
    });
    return out.map((x) => ({ ...x, cols: Math.max(1, cols.length) }));
  };

  const now = new Date();
  return (
    <section className="cal">
      <aside className="cal-side">
        <MiniMonth value={anchor} onPick={(d) => setAnchor(d)} />
        <div className="folder-section">Kalendáre</div>
        {cals === null ? (
          <Spinner size="extra-tiny" className="pad" />
        ) : (
          cals.map((c) => (
            <div key={c.id} className="cal-check">
              <Checkbox
                checked={!hidden.includes(c.id)}
                onChange={(_, d) => setHidden((h) => (d.checked ? h.filter((x) => x !== c.id) : [...h, c.id]))}
                label={
                  <span className="row gap0">
                    <span className="dot" style={{ background: c.color ?? accountColor(props.account.color) }} /> {c.name}
                  </span>
                }
              />
            </div>
          ))
        )}
      </aside>
      <div className="cal-main">
        <div className="cal-toolbar">
          <Button appearance="primary" onClick={() => newAt(new Date(Math.ceil(Date.now() / 1800000) * 1800000))} disabled={!cals?.length}>
            Nová udalosť
          </Button>
          <Button onClick={() => setAnchor(startOfDay(new Date()))}>Dnes</Button>
          <Button appearance="subtle" aria-label="Predchádzajúce" onClick={() => move(-1)}>‹</Button>
          <Button appearance="subtle" aria-label="Nasledujúce" onClick={() => move(1)}>›</Button>
          <span className="cal-title">{title}</span>
          {loading && <Spinner size="extra-tiny" />}
          <span className="grow" />
          <TabList size="small" selectedValue={view} onTabSelect={(_, d) => setView(d.value as View)}>
            <Tab value="day">Deň</Tab>
            <Tab value="workweek">Pracovný týždeň</Tab>
            <Tab value="week">Týždeň</Tab>
            <Tab value="month">Mesiac</Tab>
          </TabList>
        </div>
        {error && (
          <div className="pane-error pad">
            ⚠ {error} <button className="link" onClick={props.onSettings}>Nastavenia konta</button>
          </div>
        )}
        {view === 'month' ? (
          <div className="cal-month">
            {['Po', 'Ut', 'St', 'Št', 'Pi', 'So', 'Ne'].map((d) => <div key={d} className="cal-month-head">{d}</div>)}
            {days.map((d) => {
              const list = dayEvents(d).sort((a, b) => Number(b.allDay) - Number(a.allDay) || Date.parse(a.start) - Date.parse(b.start));
              return (
                <div
                  key={d.toISOString()}
                  className={`cal-month-day ${d.getMonth() !== anchor.getMonth() ? 'other' : ''} ${sameDay(d, now) ? 'today' : ''}`}
                  onDoubleClick={() => newAt(new Date(d.getFullYear(), d.getMonth(), d.getDate(), 9))}
                >
                  <button className="cal-daynum" onClick={() => { setAnchor(d); setView('day'); }}>{d.getDate()}</button>
                  {list.slice(0, 4).map((e) => (
                    <button key={`${e.uid}-${e.start}`} className="cal-chip" style={{ borderLeftColor: colorOf(e) }} onClick={() => setEdit(e)}>
                      {!e.allDay && <span className="muted">{fmtTime(e.start)} </span>}
                      {e.summary}
                    </button>
                  ))}
                  {list.length > 4 && <span className="muted small">+{list.length - 4} ďalšie</span>}
                </div>
              );
            })}
          </div>
        ) : (
          <div className="cal-week" style={{ ['--days' as string]: days.length }}>
            <div className="cal-head">
              <div className="cal-gutter" />
              {days.map((d) => (
                <div key={d.toISOString()} className={`cal-dayhead ${sameDay(d, now) ? 'today' : ''}`}>
                  <span className="cal-wd">{d.toLocaleDateString('sk-SK', { weekday: 'short' })}</span>
                  <span className="cal-dn">{d.getDate()}</span>
                </div>
              ))}
            </div>
            <div className="cal-allday">
              <div className="cal-gutter muted small">celý deň</div>
              {days.map((d) => (
                <div key={d.toISOString()} className="cal-allday-cell" onDoubleClick={() => newAt(d, true)}>
                  {dayEvents(d).filter((e) => e.allDay || Date.parse(e.end) - Date.parse(e.start) >= DAY_MS).map((e) => (
                    <button key={`${e.uid}-${e.start}`} className="cal-chip" style={{ borderLeftColor: colorOf(e) }} onClick={() => setEdit(e)}>{e.summary}</button>
                  ))}
                </div>
              ))}
            </div>
            <div className="cal-scroll" ref={scroller}>
              <div className="cal-grid" style={{ height: 24 * HOUR }}>
                <div className="cal-gutter">
                  {Array.from({ length: 24 }, (_, h) => <div key={h} className="cal-hour-label" style={{ top: h * HOUR }}>{h}:00</div>)}
                </div>
                {days.map((d) => {
                  const list = dayEvents(d).filter((e) => !e.allDay && Date.parse(e.end) - Date.parse(e.start) < DAY_MS);
                  return (
                    <div
                      key={d.toISOString()}
                      className={`cal-col ${sameDay(d, now) ? 'today' : ''}`}
                      onClick={(ev) => {
                        if (ev.target !== ev.currentTarget) return;
                        const y = ev.nativeEvent.offsetY;
                        const mins = Math.floor((y / HOUR) * 2) * 30;
                        newAt(new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, mins));
                      }}
                    >
                      {layout(list).map(({ ev, col, cols }) => {
                        const s = Math.max(Date.parse(ev.start), d.getTime());
                        const e = Math.min(Date.parse(ev.end), d.getTime() + DAY_MS);
                        const top = ((s - d.getTime()) / 3600000) * HOUR;
                        const h = Math.max(((e - s) / 3600000) * HOUR, 20);
                        return (
                          <button
                            key={`${ev.uid}-${ev.start}`}
                            className="cal-event"
                            style={{ top, height: h - 2, left: `${(col / cols) * 100}%`, width: `${100 / cols}%`, borderLeftColor: colorOf(ev), background: `${colorOf(ev)}1f` }}
                            onClick={() => setEdit(ev)}
                            title={`${ev.summary}${ev.location ? ` · ${ev.location}` : ''}`}
                          >
                            <b>{ev.summary}</b>
                            <span className="muted">{fmtTime(ev.start)}–{fmtTime(ev.end)}{ev.location ? ` · ${ev.location}` : ''}</span>
                          </button>
                        );
                      })}
                      {sameDay(d, now) && <div className="cal-now" style={{ top: ((now.getHours() * 60 + now.getMinutes()) / 60) * HOUR }} />}
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}
      </div>
      {edit && cals && (
        <EventDialog event={edit} calendars={cals} onClose={() => setEdit(null)} onSave={saveEvent} onDelete={deleteEvent} />
      )}
    </section>
  );
}

/** Malý mesačný kalendár na výber dňa. */
function MiniMonth(props: { value: Date; onPick: (d: Date) => void }) {
  const [month, setMonth] = useState(() => new Date(props.value.getFullYear(), props.value.getMonth(), 1));
  useEffect(() => setMonth(new Date(props.value.getFullYear(), props.value.getMonth(), 1)), [props.value]);
  const start = mondayOf(month);
  const today = new Date();
  return (
    <div className="mini-month">
      <div className="mini-head">
        <button className="linkbtn" aria-label="Predchádzajúci mesiac" onClick={() => setMonth(new Date(month.getFullYear(), month.getMonth() - 1, 1))}>‹</button>
        <span>{month.toLocaleDateString('sk-SK', { month: 'long', year: 'numeric' })}</span>
        <button className="linkbtn" aria-label="Nasledujúci mesiac" onClick={() => setMonth(new Date(month.getFullYear(), month.getMonth() + 1, 1))}>›</button>
      </div>
      <div className="mini-grid">
        {['P', 'U', 'S', 'Š', 'P', 'S', 'N'].map((d, i) => <span key={i} className="mini-wd">{d}</span>)}
        {Array.from({ length: 42 }, (_, i) => addDays(start, i)).map((d) => (
          <button
            key={d.toISOString()}
            className={`mini-day ${d.getMonth() !== month.getMonth() ? 'other' : ''} ${sameDay(d, today) ? 'today' : ''} ${sameDay(d, props.value) ? 'sel' : ''}`}
            onClick={() => props.onPick(d)}
          >
            {d.getDate()}
          </button>
        ))}
      </div>
    </div>
  );
}
