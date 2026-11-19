// Úvodná stránka „Súhrn“ (ako úvodná stránka Thunderbirdu): kontá s neprečítanými, dnešné stretnutia,
// úlohy, správy s vlajkou/pripnuté a upozornenia. Vlastné pozadie (vstavané tapety alebo obrázok),
// bloky s nastaviteľnou priehľadnosťou a rozmazaním, aby boli čitateľné na farebnom pozadí.
import { Button, Checkbox, Radio, RadioGroup, Slider } from '@fluentui/react-components';
import { useEffect, useMemo, useRef, useState } from 'react';
import type { AccountPublic } from '../../core/account';
import type { MessageSummary, OutboxItem } from '../../core/api';
import type { CalEvent } from '../../core/ical';
import { findMeeting, MEETING_LABEL } from '../../core/meeting';
import type { SummarySettings, SummaryTile, Task } from '../../core/settings';

export const SUMMARY = '__summary__';

const TILE_LABEL: Record<SummaryTile, string> = {
  accounts: 'Kontá',
  today: 'Dnes a zajtra',
  tasks: 'Úlohy',
  flagged: 'S vlajkou a pripnuté',
  notice: 'Treba si všimnúť',
};
const DEFAULT_TILES: { id: SummaryTile; on: boolean }[] = (['accounts', 'today', 'tasks', 'flagged', 'notice'] as SummaryTile[]).map((id) => ({ id, on: true }));

// Vstavané tapety – vlastné jednoduché kresby (vlny), žiadne cudzie obrázky.
const WALLS: Record<string, { name: string; dark: boolean; bg: [string, string, string]; waves: [string, string, string] }> = {
  sonoma: { name: 'Sonoma', dark: false, bg: ['#3b5bdb', '#7b5bd6', '#e8590c'], waves: ['#ffd43b', '#f76707', '#4dabf7'] },
  sunset: { name: 'Západ slnka', dark: false, bg: ['#ff922b', '#e64980', '#5f3dc4'], waves: ['#ffe066', '#f03e3e', '#ff8787'] },
  sage: { name: 'Šalvia', dark: false, bg: ['#d8e6df', '#8fb5a6', '#2f6f5e'], waves: ['#f1f5f2', '#6f9b8d', '#c3d9cf'] },
  ocean: { name: 'Oceán', dark: false, bg: ['#a5d8ff', '#339af0', '#1864ab'], waves: ['#e7f5ff', '#74c0fc', '#1c7ed6'] },
  night: { name: 'Noc', dark: true, bg: ['#0b1026', '#1e2a6b', '#3b2c7a'], waves: ['#4263eb', '#7048e8', '#1c7ed6'] },
  plain: { name: 'Bez obrázka', dark: false, bg: ['#eef1f4', '#eef1f4', '#e3e6ea'], waves: ['#e3e6ea', '#dde1e6', '#e9ecef'] },
};

function Wallpaper({ id }: { id: string }) {
  const w = WALLS[id] ?? WALLS.sonoma!;
  return (
    <svg className="sum-wall" viewBox="0 0 1600 1000" preserveAspectRatio="xMidYMid slice" aria-hidden="true">
      <defs>
        <linearGradient id={`sw-${id}`} x1="0" y1="0" x2="1" y2="1">
          <stop offset="0" stopColor={w.bg[0]} />
          <stop offset=".5" stopColor={w.bg[1]} />
          <stop offset="1" stopColor={w.bg[2]} />
        </linearGradient>
      </defs>
      <rect width="1600" height="1000" fill={`url(#sw-${id})`} />
      <path d="M0 640 C320 470 600 810 940 640 S1440 450 1600 560 L1600 1000 L0 1000Z" fill={w.waves[0]} opacity=".55" />
      <path d="M0 770 C380 640 660 940 1040 770 S1460 690 1600 750 L1600 1000 L0 1000Z" fill={w.waves[1]} opacity=".6" />
      <path d="M0 300 C400 130 720 450 1120 260 S1520 170 1600 230 L1600 0 L0 0Z" fill={w.waves[2]} opacity=".45" />
    </svg>
  );
}

function greeting(d: Date) {
  const h = d.getHours();
  return h < 5 ? 'Dobrú noc' : h < 9 ? 'Dobré ráno' : h < 12 ? 'Dobré dopoludnie' : h < 18 ? 'Dobré popoludnie' : h < 22 ? 'Dobrý večer' : 'Dobrú noc';
}
const dayKey = (d: Date) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

/** Zmenší obrázok na max. 2560 px (JPEG) a zistí, či je tmavý (podľa priemerného jasu). */
async function prepareImage(file: File): Promise<{ url: string; dark: boolean }> {
  const src = await new Promise<string>((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(String(r.result));
    r.onerror = () => rej(new Error('Obrázok sa nepodarilo načítať.'));
    r.readAsDataURL(file);
  });
  const img = await new Promise<HTMLImageElement>((res, rej) => {
    const i = new Image();
    i.onload = () => res(i);
    i.onerror = () => rej(new Error('Súbor nie je obrázok.'));
    i.src = src;
  });
  const k = Math.min(1, 2560 / Math.max(img.width, img.height));
  const c = document.createElement('canvas');
  c.width = Math.round(img.width * k);
  c.height = Math.round(img.height * k);
  const x = c.getContext('2d')!;
  x.drawImage(img, 0, 0, c.width, c.height);
  const s = document.createElement('canvas');
  s.width = s.height = 32;
  const sx = s.getContext('2d')!;
  sx.drawImage(c, 0, 0, 32, 32);
  const px = sx.getImageData(0, 0, 32, 32).data;
  let lum = 0;
  for (let i = 0; i < px.length; i += 4) lum += 0.2126 * px[i]! + 0.7152 * px[i + 1]! + 0.0722 * px[i + 2]!;
  return { url: c.toDataURL('image/jpeg', 0.85), dark: lum / (px.length / 4) < 120 };
}

export function SummaryPage(props: {
  accounts: AccountPublic[];
  unread: Record<string, number>;
  colorOf: (id: string) => string;
  /** IMAP kontá s kalendárom */
  calendarAccounts: { id: string; name: string }[];
  tasks: Task[];
  quota: Record<string, { used: number; limit: number } | null>;
  settings: SummarySettings;
  onSettings: (s: SummarySettings) => void;
  onOpenAccount: (id: string) => void;
  onOpenMessage: (m: MessageSummary) => void;
  onOpenTasks: () => void;
  onToggleTask: (t: Task) => void;
  onOpenCalendar: (accountId: string) => void;
  /** úplne vypnúť Súhrn (zapína sa späť v ozubenom kolieske) */
  onDisable: () => void;
}) {
  const api = window.mailApi;
  const s = props.settings;
  const [edit, setEdit] = useState(false);
  const [now, setNow] = useState(() => new Date());
  const [events, setEvents] = useState<(CalEvent & { accountId: string })[] | null>(null);
  const [flagged, setFlagged] = useState<MessageSummary[] | null>(null);
  const [outbox, setOutbox] = useState<OutboxItem[]>([]);
  const [imgErr, setImgErr] = useState('');
  const fileRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const t = window.setInterval(() => setNow(new Date()), 60_000);
    return () => window.clearInterval(t);
  }, []);
  const today = dayKey(now);
  const calKey = props.calendarAccounts.map((a) => a.id).join(',');
  useEffect(() => {
    if (!api) return;
    let alive = true;
    const from = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const to = new Date(from.getTime() + 2 * 86400000);
    void Promise.all(props.calendarAccounts.map((a) => api.calEvents(a.id, [], from.toISOString(), to.toISOString()).then((r) => (r.ok ? r.data.map((e) => ({ ...e, accountId: a.id })) : [])))).then(
      (rs) => {
        if (!alive) return;
        // ten istý kalendár zdieľaný vo viacerých kontách → udalosť len raz
        const seen = new Set<string>();
        setEvents(rs.flat().filter((e) => {
          const k = `${e.uid}|${e.start}`;
          if (seen.has(k)) return false;
          seen.add(k);
          return true;
        }).sort((a, b) => a.start.localeCompare(b.start)));
      },
    );
    void api.flaggedAll(8).then((r) => alive && setFlagged(r.ok ? r.data : []));
    void api.outboxList().then((r) => alive && r.ok && setOutbox(r.data));
    return () => {
      alive = false;
    };
  }, [api, calKey, today]); // eslint-disable-line react-hooks/exhaustive-deps

  const tiles = useMemo(() => {
    const t = s.tiles ?? DEFAULT_TILES;
    return [...t, ...DEFAULT_TILES.filter((d) => !t.some((x) => x.id === d.id))];
  }, [s.tiles]);
  const wallId = s.wallpaper ?? 'sonoma';
  const custom = wallId === 'custom' && s.customImage;
  const dark = s.dark ?? (custom ? !!s.customDark : (WALLS[wallId] ?? WALLS.sonoma!).dark);
  const opacity = s.opacity ?? 0.7;
  const cardStyle: React.CSSProperties = {
    background: dark ? `rgba(22, 24, 32, ${opacity})` : `rgba(255, 255, 255, ${opacity})`,
    backdropFilter: s.blur === false ? undefined : 'blur(14px) saturate(1.4)',
    WebkitBackdropFilter: s.blur === false ? undefined : 'blur(14px) saturate(1.4)',
  };

  const todayEv = (events ?? []).filter((e) => !e.allDay ? dayKey(new Date(e.start)) === today : e.start.slice(0, 10) <= today && e.end.slice(0, 10) > today || dayKey(new Date(e.start)) === today);
  const tomorrow = (events ?? []).filter((e) => !todayEv.includes(e));
  const openTasks = props.tasks.filter((t) => !t.done).sort((a, b) => (a.due ?? '9999').localeCompare(b.due ?? '9999')).slice(0, 6);
  const failed = outbox.filter((o) => o.retry);
  const later = outbox.filter((o) => o.kind === 'later' && !o.retry);
  const fullBoxes = props.accounts.filter((a) => {
    const q = props.quota[a.id];
    return q && q.limit > 0 && q.used / q.limit > 0.9;
  });
  const totalUnread = props.accounts.reduce((n, a) => n + (props.unread[a.id] ?? 0), 0);
  const accName = (id: string) => props.accounts.find((a) => a.id === id)?.name ?? '';

  const tile = (id: SummaryTile) => {
    switch (id) {
      case 'accounts':
        return (
          <>
            {props.accounts.map((a) => (
              <button key={a.id} className="sum-row sum-click" onClick={() => props.onOpenAccount(a.id)}>
                <span className="sum-sw" style={{ background: props.colorOf(a.id) }} />
                <span className="grow sum-ell">{a.name || a.email}</span>
                <b className={props.unread[a.id] ? '' : 'sum-dim'}>{props.unread[a.id] ?? 0}</b>
              </button>
            ))}
            {!props.accounts.length && <div className="sum-dim">Zatiaľ žiadne kontá.</div>}
          </>
        );
      case 'today':
        return (
          <>
            {events === null && <div className="sum-dim">Načítavam kalendáre…</div>}
            {events && !todayEv.length && <div className="sum-dim">Dnes nemáš žiadne stretnutia.</div>}
            {todayEv.map((e, i) => {
              const m = findMeeting(e);
              const past = !e.allDay && Date.parse(e.end) < now.getTime();
              return (
                <div key={i} className={`sum-row ${past ? 'sum-past' : ''}`}>
                  <b className="sum-time">{e.allDay ? 'celý deň' : new Date(e.start).toLocaleTimeString('sk-SK', { hour: '2-digit', minute: '2-digit' })}</b>
                  <button className="grow sum-ell sum-link" onClick={() => props.onOpenCalendar(e.accountId)} title={`${e.summary} · ${accName(e.accountId)}`}>{e.summary}</button>
                  {m && !past && (
                    <button className="sum-pill sum-click" title={m.url} onClick={() => (m.kind === 'jitsi' ? void api?.meetJitsi(m.url) : void api?.openExternal(m.url))}>
                      {MEETING_LABEL[m.kind].replace('Microsoft ', '')} ▸
                    </button>
                  )}
                </div>
              );
            })}
            {events && tomorrow.length > 0 && (
              <div className="sum-foot">
                Zajtra: {tomorrow.slice(0, 3).map((e) => `${e.allDay ? '' : new Date(e.start).toLocaleTimeString('sk-SK', { hour: '2-digit', minute: '2-digit' }) + ' '}${e.summary}`).join(' · ')}
                {tomorrow.length > 3 ? ` (+${tomorrow.length - 3})` : ''}
              </div>
            )}
          </>
        );
      case 'tasks':
        return (
          <>
            {!openTasks.length && <div className="sum-dim">Žiadne otvorené úlohy.</div>}
            {openTasks.map((t) => {
              const late = t.due && t.due < today;
              return (
                <div key={t.id} className="sum-row">
                  <input type="checkbox" checked={t.done} onChange={() => props.onToggleTask(t)} aria-label={`Hotovo: ${t.title}`} />
                  <button className="grow sum-ell sum-link" onClick={props.onOpenTasks}>{t.title}</button>
                  {t.due && <span className={`sum-pill ${late ? 'sum-late' : ''}`}>{late ? 'po termíne' : t.due === today ? 'dnes' : new Date(`${t.due}T12:00`).toLocaleDateString('sk-SK', { weekday: 'short', day: 'numeric', month: 'numeric' })}</span>}
                </div>
              );
            })}
          </>
        );
      case 'flagged':
        return (
          <>
            {flagged === null && <div className="sum-dim">Načítavam…</div>}
            {flagged && !flagged.length && <div className="sum-dim">Žiadne správy s vlajkou.</div>}
            {(flagged ?? []).map((m) => (
              <button key={`${m.accountId}/${m.folder}/${m.uid}`} className="sum-row sum-click" onClick={() => props.onOpenMessage(m)} title={`${m.fromName || m.fromAddress} – ${m.subject}`}>
                <span className="sum-sw" style={{ background: props.colorOf(m.accountId) }} />
                <span className="grow sum-ell"><b>{m.fromName || m.fromAddress}</b> – {m.subject}</span>
              </button>
            ))}
          </>
        );
      case 'notice': {
        const items: string[] = [];
        if (failed.length) items.push(`${failed.length} ${failed.length === 1 ? 'správa sa nepodarila odoslať' : 'správy sa nepodarilo odoslať'} – skúsi sa znova (Odoslať a prijať)`);
        if (later.length) items.push(`${later.length} naplánované na neskoršie odoslanie`);
        for (const a of fullBoxes) items.push(`Schránka ${a.name} je plná na ${Math.round((props.quota[a.id]!.used / props.quota[a.id]!.limit) * 100)} %`);
        if (totalUnread) items.push(`Spolu ${totalUnread} neprečítaných správ`);
        return items.length ? <>{items.map((t, i) => <div key={i} className="sum-row">{t}</div>)}</> : <div className="sum-dim">Všetko v poriadku.</div>;
      }
    }
  };

  const moveTile = (i: number, d: number) => {
    const t = [...tiles];
    const j = i + d;
    if (j < 0 || j >= t.length) return;
    [t[i], t[j]] = [t[j]!, t[i]!];
    props.onSettings({ ...s, tiles: t });
  };

  return (
    <section className={`summary ${dark ? 'dark' : ''}`} aria-label="Súhrn">
      {custom ? <div className="sum-wall sum-img" style={{ backgroundImage: `url(${s.customImage})` }} /> : <Wallpaper id={WALLS[wallId] ? wallId : 'sonoma'} />}
      <div className="sum-content">
        <div className="sum-head">
          <span className="sum-hello">{greeting(now)}</span>
          <span className="sum-date">{now.toLocaleDateString('sk-SK', { weekday: 'long', day: 'numeric', month: 'long' })}</span>
          <span className="grow" />
          <Button size="small" className="sum-edit-btn" style={cardStyle} onClick={() => setEdit((x) => !x)}>{edit ? 'Hotovo' : 'Upraviť súhrn'}</Button>
        </div>

        {edit && (
          <div className="sum-card sum-settings" style={cardStyle}>
            <div className="sum-set-grid">
              <section>
                <div className="sum-h">Po spustení otvoriť</div>
                <RadioGroup value={s.startup ?? 'summary'} onChange={(_, d) => props.onSettings({ ...s, startup: d.value as SummarySettings['startup'] })}>
                  <Radio value="summary" label="Súhrn" />
                  <Radio value="all" label="Všetky doručené" />
                  <Radio value="last" label="Naposledy otvorené konto" />
                </RadioGroup>
                <Button size="small" appearance="subtle" onClick={props.onDisable} title="Súhrn zmizne z pásu kont; zapneš ho späť v ozubenom kolieske dole vľavo">
                  Vypnúť Súhrn úplne
                </Button>
              </section>
              <section>
                <div className="sum-h">Dlaždice</div>
                {tiles.map((t, i) => (
                  <div key={t.id} className="sum-tile-set">
                    <Checkbox label={TILE_LABEL[t.id]} checked={t.on} onChange={(_, d) => props.onSettings({ ...s, tiles: tiles.map((x) => (x.id === t.id ? { ...x, on: !!d.checked } : x)) })} />
                    <span className="grow" />
                    <Button size="small" appearance="subtle" disabled={!i} onClick={() => moveTile(i, -1)} aria-label="Vyššie">↑</Button>
                    <Button size="small" appearance="subtle" disabled={i === tiles.length - 1} onClick={() => moveTile(i, 1)} aria-label="Nižšie">↓</Button>
                  </div>
                ))}
              </section>
              <section>
                <div className="sum-h">Pozadie</div>
                <div className="sum-walls">
                  {Object.entries(WALLS).map(([id, w]) => (
                    <button key={id} className={`sum-wall-pick ${wallId === id ? 'on' : ''}`} onClick={() => props.onSettings({ ...s, wallpaper: id })} title={w.name}>
                      <span style={{ background: `linear-gradient(135deg, ${w.bg[0]}, ${w.bg[1]}, ${w.bg[2]})` }} />
                      {w.name}
                    </button>
                  ))}
                  <button className={`sum-wall-pick ${wallId === 'custom' ? 'on' : ''}`} onClick={() => (s.customImage ? props.onSettings({ ...s, wallpaper: 'custom' }) : fileRef.current?.click())} title="Vlastný obrázok">
                    <span style={s.customImage ? { backgroundImage: `url(${s.customImage})`, backgroundSize: 'cover' } : undefined}>{s.customImage ? '' : '+'}</span>
                    Vlastný
                  </button>
                </div>
                <Button size="small" onClick={() => fileRef.current?.click()}>Vybrať obrázok…</Button>
                <input
                  ref={fileRef}
                  type="file"
                  accept="image/*"
                  hidden
                  onChange={async (ev) => {
                    const f = ev.target.files?.[0];
                    ev.target.value = '';
                    if (!f) return;
                    try {
                      setImgErr('');
                      const r = await prepareImage(f);
                      props.onSettings({ ...s, wallpaper: 'custom', customImage: r.url, customDark: r.dark, dark: undefined });
                    } catch (e) {
                      setImgErr((e as Error).message);
                    }
                  }}
                />
                {imgErr && <div className="pane-error">{imgErr}</div>}
              </section>
              <section>
                <div className="sum-h">Čitateľnosť blokov</div>
                <div className="sum-dim small">Nepriehľadnosť: {Math.round(opacity * 100)} %</div>
                <Slider min={30} max={100} step={5} value={Math.round(opacity * 100)} onChange={(_, d) => props.onSettings({ ...s, opacity: d.value / 100 })} aria-label="Nepriehľadnosť blokov" />
                <Checkbox label="Rozmazať pozadie (matné sklo)" checked={s.blur !== false} onChange={(_, d) => props.onSettings({ ...s, blur: !!d.checked })} />
                <RadioGroup layout="horizontal" value={s.dark === undefined ? 'auto' : s.dark ? 'dark' : 'light'} onChange={(_, d) => props.onSettings({ ...s, dark: d.value === 'auto' ? undefined : d.value === 'dark' })}>
                  <Radio value="auto" label="Podľa pozadia" />
                  <Radio value="light" label="Svetlé" />
                  <Radio value="dark" label="Tmavé" />
                </RadioGroup>
              </section>
            </div>
          </div>
        )}

        <div className="sum-grid">
          {tiles.filter((t) => t.on).map((t) => (
            <div key={t.id} className={`sum-card sum-${t.id}`} style={cardStyle}>
              <div className="sum-h">{TILE_LABEL[t.id]}</div>
              {tile(t.id)}
            </div>
          ))}
        </div>
      </div>
    </section>
  );
}
