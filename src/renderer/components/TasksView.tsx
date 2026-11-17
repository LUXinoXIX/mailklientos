// Úlohy: vlastné aj zo správ (vlajka s termínom). Skupiny Po termíne / Dnes / Zajtra / Neskôr / Bez termínu / Hotové.
import { Button, Checkbox, Input } from '@fluentui/react-components';
import { useEffect, useMemo, useRef, useState } from 'react';
import type { AppSettings, Task } from '../../core/settings';

const iso = (d: Date) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
export const todayIso = () => iso(new Date());
export const dueLabel = (due?: string) => {
  if (!due) return '';
  const t = todayIso();
  const tm = iso(new Date(Date.now() + 86400000));
  if (due < t) return `po termíne (${new Date(`${due}T00:00`).toLocaleDateString('sk-SK')})`;
  if (due === t) return 'dnes';
  if (due === tm) return 'zajtra';
  return new Date(`${due}T00:00`).toLocaleDateString('sk-SK', { weekday: 'short', day: 'numeric', month: 'numeric' });
};

export function TasksView(props: {
  settings: AppSettings;
  onSettings: (s: AppSettings) => void;
  onOpenMessage: (m: NonNullable<Task['message']>) => void;
}) {
  const tasks = props.settings.tasks ?? [];
  const [title, setTitle] = useState('');
  const [due, setDue] = useState('');
  const [showDone, setShowDone] = useState(false);
  const save = (list: Task[]) => props.onSettings({ ...props.settings, tasks: list });
  const patch = (id: string, p: Partial<Task>) => save(tasks.map((t) => (t.id === id ? { ...t, ...p, dirty: true } : t)));
  const [sync, setSync] = useState<string>('');

  // synchronizácia so serverom (Úlohy v Zimbre / CalDAV): pri otvorení a 3 s po zmene
  const latest = useRef(props.settings);
  latest.current = props.settings;
  const running = useRef(false);
  const runSync = async () => {
    const api = window.mailApi;
    if (!api || running.current) return;
    running.current = true;
    setSync('Synchronizujem úlohy…');
    const before = latest.current.tasks ?? [];
    const r = await api.tasksSync(before);
    running.current = false;
    if (!r.ok) return setSync(`Synchronizácia zlyhala: ${r.error}`);
    if (!r.data.accounts) return setSync('');
    // ak sa medzitým niečo zmenilo, synchronizovať znova neskôr
    if (latest.current.tasks !== before) return setSync('');
    if (JSON.stringify(r.data.tasks) !== JSON.stringify(before)) props.onSettings({ ...latest.current, tasks: r.data.tasks });
    setSync(`Synchronizované so serverom ${new Date().toLocaleTimeString('sk-SK', { hour: '2-digit', minute: '2-digit' })}`);
  };
  useEffect(() => {
    void runSync();
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => {
    if (!tasks.some((t) => t.dirty || !t.remote)) return;
    const t = window.setTimeout(() => void runSync(), 3000);
    return () => window.clearTimeout(t);
  }, [tasks]); // eslint-disable-line react-hooks/exhaustive-deps

  const groups = useMemo(() => {
    const t = todayIso();
    const tm = iso(new Date(Date.now() + 86400000));
    const open = tasks.filter((x) => !x.done).sort((a, b) => (a.due ?? '9999').localeCompare(b.due ?? '9999'));
    return [
      { name: 'Po termíne', list: open.filter((x) => x.due && x.due < t), warn: true },
      { name: 'Dnes', list: open.filter((x) => x.due === t) },
      { name: 'Zajtra', list: open.filter((x) => x.due === tm) },
      { name: 'Neskôr', list: open.filter((x) => x.due && x.due > tm) },
      { name: 'Bez termínu', list: open.filter((x) => !x.due) },
    ];
  }, [tasks]);
  const done = tasks.filter((x) => x.done);

  const add = () => {
    if (!title.trim()) return;
    save([...tasks, { id: Math.random().toString(36).slice(2, 10), title: title.trim(), due: due || undefined, done: false, created: new Date().toISOString() }]);
    setTitle('');
    setDue('');
  };

  const row = (t: Task) => (
    <div key={t.id} className={`task ${t.done ? 'done' : ''}`}>
      <Checkbox checked={t.done} onChange={(_, d) => patch(t.id, { done: !!d.checked })} aria-label="Hotové" />
      <div className="task-main">
        <input className="task-title" value={t.title} onChange={(e) => patch(t.id, { title: e.target.value })} aria-label="Názov úlohy" />
        <div className="task-meta">
          <input type="date" className="date-input small" value={t.due ?? ''} onChange={(e) => patch(t.id, { due: e.target.value || undefined })} aria-label="Termín" />
          {t.due && <span className={`muted small ${t.due < todayIso() && !t.done ? 'overdue' : ''}`}>{dueLabel(t.due)}</span>}
          {t.message && (
            <button className="link small" onClick={() => props.onOpenMessage(t.message!)}>
              ✉ {t.message.from}: {t.message.subject}
            </button>
          )}
        </div>
      </div>
      <Button size="small" appearance="subtle" aria-label="Odstrániť úlohu" onClick={() => {
        if (t.remote) void window.mailApi?.tasksDelete(t);
        save(tasks.filter((x) => x.id !== t.id));
      }}>×</Button>
    </div>
  );

  return (
    <section className="tasks">
      <div className="tasks-add">
        <Input className="grow" placeholder="Nová úloha… (Enter)" value={title} onChange={(_, d) => setTitle(d.value)} onKeyDown={(e) => e.key === 'Enter' && add()} />
        <input type="date" className="date-input" value={due} onChange={(e) => setDue(e.target.value)} aria-label="Termín" />
        <Button appearance="primary" onClick={add}>Pridať</Button>
      </div>
      {sync && <div className="muted small pad">{sync}</div>}
      <div className="tasks-list">
        {groups.map((g) =>
          g.list.length ? (
            <div key={g.name}>
              <div className={`folder-section ${g.warn ? 'overdue' : ''}`}>{g.name} ({g.list.length})</div>
              {g.list.map(row)}
            </div>
          ) : null,
        )}
        {!tasks.some((x) => !x.done) && <div className="empty">Žiadne otvorené úlohy. Úlohu vytvoríte aj zo správy: Domov → Úloha.</div>}
        {done.length > 0 && (
          <>
            <button className="link small pad" onClick={() => setShowDone(!showDone)}>{showDone ? 'Skryť hotové' : `Zobraziť hotové (${done.length})`}</button>
            {showDone && done.map(row)}
          </>
        )}
      </div>
    </section>
  );
}
