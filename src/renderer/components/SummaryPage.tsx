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
import { loadFeed, type Feed } from '../rss';

export const SUMMARY = '__summary__';

const TILE_LABEL: Record<SummaryTile, string> = {
  accounts: 'Kontá',
  today: 'Dnes a zajtra',
  tasks: 'Úlohy',
  flagged: 'S vlajkou a pripnuté',
  notice: 'Treba si všimnúť',
  feeds: 'Kalendár z odkazu',
  news: 'Správy',
};
const DEFAULT_TILES: { id: SummaryTile; on: boolean }[] = (['accounts', 'today', 'tasks', 'flagged', 'notice', 'feeds', 'news'] as SummaryTile[]).map((id) => ({ id, on: id !== 'feeds' && id !== 'news' }));

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
  const [feedData, setFeedData] = useState<Record<string, { list?: CalEvent[]; error?: string }>>({});
  const [feedForm, setFeedForm] = useState({ url: '', name: '', count: 5 });
  // správy z RSS
  const news = s.news ?? [];
  const [newsData, setNewsData] = useState<Record<string, { feed?: Feed; error?: string }>>({});
  const [newsForm, setNewsForm] = useState({ url: '', name: '', count: 5 });
  const newsKey = news.map((n) => n.url).join(',');
  useEffect(() => {
    if (!news.length) return;
    let alive = true;
    const load = () => {
      for (const n of news) void loadFeed(n.url).then((feed) => alive && setNewsData((d) => ({ ...d, [n.id]: { feed } }))).catch((e) => alive && setNewsData((d) => ({ ...d, [n.id]: { error: (e as Error).message } })));
    };
    load();
    const t = window.setInterval(load, 20 * 60_000);
    return () => {
      alive = false;
      window.clearInterval(t);
    };
  }, [newsKey]); // eslint-disable-line react-hooks/exhaustive-deps
  const feeds = s.feeds ?? [];
  // kalendáre z odkazu na celú šírku: každý vo vlastnom stĺpci, vždy pod ostatnými dlaždicami
  const wideFeeds = !!s.feedsWide && feeds.length > 0 && (s.tiles ?? []).find((t) => t.id === 'feeds')?.on !== false;
  const feedKey = feeds.map((f) => `${f.url}|${f.count}`).join(',');
  useEffect(() => {
    if (!api || !feeds.length) return;
    let alive = true;
    for (const f of feeds)
      void api.icsUpcoming(f.url, f.count).then((r) => alive && setFeedData((d) => ({ ...d, [f.id]: r.ok ? { list: r.data } : { error: r.error } })));
    return () => {
      alive = false;
    };
  }, [api, feedKey, dayKey(now)]); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    const t = window.setInterval(() => setNow(new Date()), 15_000);
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
      case 'feeds':
        return (
          <>
            {!feeds.length && <div className="sum-dim">Pridaj kalendár z odkazu v „Upraviť súhrn“ (napr. sviatky, rozvrh, zdieľaný kalendár).</div>}
            {feeds.map((f) => {
              const d = feedData[f.id];
              return (
                <div key={f.id} className="sum-feed">
                  {(feeds.length > 1 || wideFeeds) && <div className="sum-feed-name">{f.name}</div>}
                  {!d && <div className="sum-dim">Načítavam…</div>}
                  {d?.error && <div className="sum-dim">⚠ {d.error}</div>}
                  {d?.list && !d.list.length && <div className="sum-dim">Žiadne nadchádzajúce udalosti.</div>}
                  {d?.list?.map((e, i) => {
                    const st = new Date(e.start);
                    const isToday = dayKey(st) === today;
                    return (
                      <div key={i} className="sum-row">
                        <b className="sum-time sum-time-wide">
                          {isToday ? 'dnes' : st.toLocaleDateString('sk-SK', { weekday: 'short', day: 'numeric', month: 'numeric' })}
                          {!e.allDay && ` ${st.toLocaleTimeString('sk-SK', { hour: '2-digit', minute: '2-digit' })}`}
                        </b>
                        <span className="grow sum-ell" title={[e.summary, e.location].filter(Boolean).join(' · ')}>{e.summary}</span>
                      </div>
                    );
                  })}
                </div>
              );
            })}
          </>
        );
      case 'news':
        return (
          <>
            {!news.length && <div className="sum-dim">Pridaj RSS kanál v „Upraviť súhrn“ – napr. dennikn.sk/feed alebo adresu webu, kanál sa nájde sám.</div>}
            {news.map((n) => {
              const d = newsData[n.id];
              return (
                <div key={n.id} className="sum-feed">
                  {news.length > 1 && <div className="sum-feed-name">{n.name}</div>}
                  {!d && <div className="sum-dim">Načítavam…</div>}
                  {d?.error && <div className="sum-dim">⚠ {d.error}</div>}
                  {d?.feed?.items.slice(0, n.count).map((it, i) => (
                    <button key={i} className="sum-row sum-click sum-news" onClick={() => void api?.openExternal(it.link)} title={it.title}>
                      {it.image && <img src={it.image} alt="" className="sum-news-img" loading="lazy" referrerPolicy="no-referrer" />}
                      <span className="grow sum-news-t">{it.title}</span>
                      {it.date && !Number.isNaN(Date.parse(it.date)) && (
                        <span className="sum-pill">
                          {dayKey(new Date(it.date)) === today ? new Date(it.date).toLocaleTimeString('sk-SK', { hour: '2-digit', minute: '2-digit' }) : new Date(it.date).toLocaleDateString('sk-SK', { day: 'numeric', month: 'numeric' })}
                        </span>
                      )}
                    </button>
                  ))}
                </div>
              );
            })}
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
          <span className="sum-date">
            {now.toLocaleDateString('sk-SK', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}{' '}
            <b className="sum-clock">{now.toLocaleTimeString('sk-SK', { hour: '2-digit', minute: '2-digit' })}</b>
          </span>
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
                <div className="sum-h">Kalendár z odkazu (iCal)</div>
                <Checkbox label="Na celú šírku – každý kalendár vo vlastnom stĺpci, pod ostatnými dlaždicami" checked={!!s.feedsWide} onChange={(_, d) => props.onSettings({ ...s, feedsWide: !!d.checked })} />
                {feeds.map((f) => (
                  <div key={f.id} className="sum-tile-set">
                    <span className="grow sum-ell" title={f.url}>{f.name}</span>
                    <select className="sum-count" value={f.count} onChange={(ev) => props.onSettings({ ...s, feeds: feeds.map((x) => (x.id === f.id ? { ...x, count: Number(ev.target.value) } : x)) })} aria-label="Počet udalostí">
                      {[3, 5, 8, 10, 15].map((n) => <option key={n} value={n}>{n}</option>)}
                    </select>
                    <Button size="small" appearance="subtle" onClick={() => props.onSettings({ ...s, feeds: feeds.filter((x) => x.id !== f.id) })} aria-label={`Odstrániť ${f.name}`}>×</Button>
                  </div>
                ))}
                <input className="sum-input" placeholder="https://… alebo webcal://…" value={feedForm.url} onChange={(ev) => setFeedForm({ ...feedForm, url: ev.target.value })} />
                <div className="sum-tile-set">
                  <input className="sum-input grow" placeholder="Názov (napr. Sviatky)" value={feedForm.name} onChange={(ev) => setFeedForm({ ...feedForm, name: ev.target.value })} />
                  <select className="sum-count" value={feedForm.count} onChange={(ev) => setFeedForm({ ...feedForm, count: Number(ev.target.value) })} aria-label="Počet udalostí">
                    {[3, 5, 8, 10, 15].map((n) => <option key={n} value={n}>{n} udal.</option>)}
                  </select>
                  <Button
                    size="small"
                    disabled={!/^(https?|webcals?):\/\/\S+$/i.test(feedForm.url.trim())}
                    onClick={() => {
                      const url = feedForm.url.trim();
                      let name = feedForm.name.trim();
                      try {
                        name ||= new URL(url.replace(/^webcals?:/i, 'https:')).hostname;
                      } catch {
                        name ||= 'Kalendár';
                      }
                      const tilesNow = tiles.map((t) => (t.id === 'feeds' ? { ...t, on: true } : t));
                      props.onSettings({ ...s, tiles: tilesNow, feeds: [...feeds, { id: Math.random().toString(36).slice(2, 10), name, url, count: feedForm.count }] });
                      setFeedForm({ url: '', name: '', count: 5 });
                    }}
                  >
                    Pridať
                  </Button>
                </div>
              </section>
              <section>
                <div className="sum-h">Správy (RSS)</div>
                {news.map((n) => (
                  <div key={n.id} className="sum-tile-set">
                    <span className="grow sum-ell" title={n.url}>{newsData[n.id]?.feed?.title && n.name === n.url ? newsData[n.id]!.feed!.title : n.name}</span>
                    <select className="sum-count" value={n.count} onChange={(ev) => props.onSettings({ ...s, news: news.map((x) => (x.id === n.id ? { ...x, count: Number(ev.target.value) } : x)) })} aria-label="Počet správ">
                      {[3, 5, 8, 10, 15].map((k) => <option key={k} value={k}>{k}</option>)}
                    </select>
                    <Button size="small" appearance="subtle" onClick={() => props.onSettings({ ...s, news: news.filter((x) => x.id !== n.id) })} aria-label={`Odstrániť ${n.name}`}>×</Button>
                  </div>
                ))}
                <input className="sum-input" placeholder="dennikn.sk/feed alebo adresa webu" value={newsForm.url} onChange={(ev) => setNewsForm({ ...newsForm, url: ev.target.value })} />
                <div className="sum-tile-set">
                  <input className="sum-input grow" placeholder="Názov (nepovinné)" value={newsForm.name} onChange={(ev) => setNewsForm({ ...newsForm, name: ev.target.value })} />
                  <select className="sum-count" value={newsForm.count} onChange={(ev) => setNewsForm({ ...newsForm, count: Number(ev.target.value) })} aria-label="Počet správ">
                    {[3, 5, 8, 10, 15].map((k) => <option key={k} value={k}>{k} správ</option>)}
                  </select>
                  <Button
                    size="small"
                    disabled={!newsForm.url.trim()}
                    onClick={async () => {
                      let url = newsForm.url.trim();
                      if (!/^https?:\/\//i.test(url)) url = `https://${url}`;
                      let name = newsForm.name.trim();
                      if (!name) {
                        try {
                          name = (await loadFeed(url)).title || new URL(url).hostname.replace(/^www\./, '');
                        } catch {
                          name = new URL(url).hostname.replace(/^www\./, '');
                        }
                      }
                      const tilesNow = tiles.map((t) => (t.id === 'news' ? { ...t, on: true } : t));
                      props.onSettings({ ...s, tiles: tilesNow, news: [...news, { id: Math.random().toString(36).slice(2, 10), name, url, count: newsForm.count }] });
                      setNewsForm({ url: '', name: '', count: 5 });
                    }}
                  >
                    Pridať
                  </Button>
                </div>
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
          {tiles.filter((t) => t.on && !(t.id === 'feeds' && wideFeeds)).map((t) => (
            <div key={t.id} className={`sum-card sum-${t.id}`} style={cardStyle}>
              <div className="sum-h">
                {TILE_LABEL[t.id]}
                {t.id === 'feeds' && feeds.length > 1 && (
                  <button className="sum-wide-btn" title="Na celú šírku – každý kalendár vo vlastnom stĺpci, pod ostatnými dlaždicami" onClick={() => props.onSettings({ ...s, feedsWide: true })}>⇔</button>
                )}
              </div>
              {tile(t.id)}
            </div>
          ))}
        </div>
        {wideFeeds && (
          <div className="sum-card sum-feeds-wide" style={cardStyle}>
            <div className="sum-h">
              {TILE_LABEL.feeds}
              <button className="sum-wide-btn" title="Späť medzi ostatné dlaždice" onClick={() => props.onSettings({ ...s, feedsWide: false })}>⇤⇥</button>
            </div>
            <div className="sum-feed-cols" style={{ gridTemplateColumns: `repeat(${Math.max(1, feeds.length)}, minmax(0, 1fr))` }}>{tile('feeds')}</div>
          </div>
        )}
      </div>
    </section>
  );
}
