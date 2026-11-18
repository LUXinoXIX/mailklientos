// Voliteľný náhľad kalendára a úloh vedľa čítacieho panela (pošta Zimbra/IMAP), podobne ako „Môj deň“ v Outlooku.
import { useMemo, useState } from 'react';
import type { CalEvent } from '../../core/ical';
import type { Task } from '../../core/settings';
import { AddIcon, CalendarSmallIcon, ChevronLeftSmallIcon, ChevronRightSmallIcon, CloseSmallIcon, TasksSmallIcon } from '../icons';

export type MyDayEvent = CalEvent & { accountColor: string };

const DAYS = ['Po', 'Ut', 'St', 'Št', 'Pi', 'So', 'Ne'];
const iso = (d: Date) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const sameDay = (a: Date, b: Date) => iso(a) === iso(b);

/** deň patrí udalosti, ak sa s ňou prekrýva (celodenná končí o polnoci nasledujúceho dňa) */
function onDay(e: CalEvent, d: Date) {
  const from = new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
  const to = from + 86400000;
  return Date.parse(e.start) < to && Date.parse(e.end) > from;
}

export function MyDayPane(props: {
  calendar: boolean;
  tasks: boolean;
  /** udalosti v zobrazenom mesiaci (pre zvolené konto) */
  events: MyDayEvent[];
  loading: boolean;
  error?: string;
  month: Date;
  onMonth: (d: Date) => void;
  taskList: Task[];
  onToggleTask: (t: Task) => void;
  onOpenTask: (t: Task) => void;
  onAddTask: (title: string, due?: string) => void;
  onOpenCalendar: () => void;
  onOpenTasks: () => void;
  onHide: (part: 'calendar' | 'tasks') => void;
}) {
  const [day, setDay] = useState(() => new Date());
  const [newTask, setNewTask] = useState('');
  const today = new Date();
  const m = props.month;

  const cells = useMemo(() => {
    const first = new Date(m.getFullYear(), m.getMonth(), 1);
    const shift = (first.getDay() + 6) % 7;
    const start = new Date(first.getFullYear(), first.getMonth(), 1 - shift);
    return Array.from({ length: 42 }, (_, i) => new Date(start.getFullYear(), start.getMonth(), start.getDate() + i));
  }, [m]);
  const busy = useMemo(() => new Set(cells.filter((c) => props.events.some((e) => onDay(e, c))).map(iso)), [cells, props.events]);
  const dayEvents = props.events
    .filter((e) => onDay(e, day))
    .sort((a, b) => Number(b.allDay) - Number(a.allDay) || Date.parse(a.start) - Date.parse(b.start));
  const time = (s: string) => new Date(s).toLocaleTimeString('sk-SK', { hour: '2-digit', minute: '2-digit' });
  const now = Date.now();
  const todayIso = iso(today);
  const open = props.taskList
    .filter((t) => !t.done)
    .sort((a, b) => (a.due ?? '9999').localeCompare(b.due ?? '9999') || a.created.localeCompare(b.created));

  return (
    <aside className="myday" aria-label="Kalendár a úlohy">
      {props.calendar && (
        <section className="myday-sec">
          <div className="myday-head">
            <CalendarSmallIcon />
            <button className="link myday-title" onClick={props.onOpenCalendar} title="Otvoriť kalendár">Kalendár</button>
            <button className="icon-x" aria-label="Skryť kalendár" title="Skryť kalendár (zapneš v karte Zobraziť)" onClick={() => props.onHide('calendar')}>
              <CloseSmallIcon />
            </button>
          </div>
          <div className="mini-cal">
            <div className="mini-cal-nav">
              <button aria-label="Predchádzajúci mesiac" onClick={() => props.onMonth(new Date(m.getFullYear(), m.getMonth() - 1, 1))}><ChevronLeftSmallIcon /></button>
              <span>{m.toLocaleDateString('sk-SK', { month: 'long', year: 'numeric' })}</span>
              <button aria-label="Ďalší mesiac" onClick={() => props.onMonth(new Date(m.getFullYear(), m.getMonth() + 1, 1))}><ChevronRightSmallIcon /></button>
              <button
                className="mini-today"
                onClick={() => {
                  setDay(new Date());
                  props.onMonth(new Date(today.getFullYear(), today.getMonth(), 1));
                }}
              >
                Dnes
              </button>
            </div>
            <div className="mini-grid" role="grid">
              {DAYS.map((d) => <span key={d} className="mini-dow">{d}</span>)}
              {cells.map((c) => (
                <button
                  key={iso(c)}
                  className={`mini-day ${c.getMonth() !== m.getMonth() ? 'out' : ''} ${sameDay(c, today) ? 'today' : ''} ${sameDay(c, day) ? 'sel' : ''} ${busy.has(iso(c)) ? 'busy' : ''}`}
                  onClick={() => setDay(c)}
                  aria-label={c.toLocaleDateString('sk-SK', { weekday: 'long', day: 'numeric', month: 'long' })}
                >
                  {c.getDate()}
                </button>
              ))}
            </div>
          </div>
          <div className="myday-day">{sameDay(day, today) ? 'Dnes' : day.toLocaleDateString('sk-SK', { weekday: 'long', day: 'numeric', month: 'long' })}</div>
          {props.loading && <div className="muted small pad">Načítavam…</div>}
          {props.error && <div className="muted small pad">⚠ {props.error}</div>}
          {!props.loading && !props.error && !dayEvents.length && <div className="muted small pad">Žiadne stretnutia.</div>}
          {dayEvents.map((e) => {
            const past = Date.parse(e.end) < now;
            const live = Date.parse(e.start) <= now && Date.parse(e.end) > now;
            return (
              <button
                key={`${e.uid}-${e.start}`}
                className={`agenda-item ${past ? 'past' : ''} ${live ? 'live' : ''}`}
                style={{ borderLeftColor: e.color ?? e.accountColor }}
                onClick={props.onOpenCalendar}
              >
                <span className="agenda-time">{e.allDay ? 'celý deň' : `${time(e.start)}–${time(e.end)}`}</span>
                <b>{e.summary}</b>
                {e.location && <span className="muted small">{e.location}</span>}
              </button>
            );
          })}
        </section>
      )}
      {props.tasks && (
        <section className="myday-sec">
          <div className="myday-head">
            <TasksSmallIcon />
            <button className="link myday-title" onClick={props.onOpenTasks} title="Otvoriť úlohy">Úlohy</button>
            <span className="muted small">{open.length || ''}</span>
            <button className="icon-x" aria-label="Skryť úlohy" title="Skryť úlohy (zapneš v karte Zobraziť)" onClick={() => props.onHide('tasks')}>
              <CloseSmallIcon />
            </button>
          </div>
          <form
            className="myday-add"
            onSubmit={(e) => {
              e.preventDefault();
              const t = newTask.trim();
              if (!t) return;
              props.onAddTask(t, props.calendar && !sameDay(day, today) ? iso(day) : undefined);
              setNewTask('');
            }}
          >
            <AddIcon />
            <input value={newTask} onChange={(e) => setNewTask(e.target.value)} placeholder="Pridať úlohu" aria-label="Nová úloha" />
          </form>
          {!open.length && <div className="muted small pad">Žiadne otvorené úlohy.</div>}
          {open.map((t) => (
            <div key={t.id} className="agenda-task">
              <input type="checkbox" checked={t.done} onChange={() => props.onToggleTask(t)} aria-label="Hotové" />
              <button className="link" onClick={() => props.onOpenTask(t)} title={t.title}>{t.title}</button>
              {t.due && (
                <span className={`small ${t.due < todayIso ? 'overdue' : 'muted'}`}>
                  {t.due < todayIso ? 'po termíne' : t.due === todayIso ? 'dnes' : new Date(`${t.due}T00:00`).toLocaleDateString('sk-SK', { day: 'numeric', month: 'numeric' })}
                </span>
              )}
            </div>
          ))}
        </section>
      )}
    </aside>
  );
}
