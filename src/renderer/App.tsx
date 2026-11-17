// Rozloženie C: pás kont úplne vľavo → pás modulov podľa konta → obsah (IMAP klient alebo vložený Outlook) → pravý panel.
import { FluentProvider } from '@fluentui/react-components';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { isImap, isM365, type AccountPublic, type ImapAccountPublic, type M365Account } from '../core/account';
import type { FolderInfo, MailApi, MessageFull, MessageSummary, OutboxItem } from '../core/api';
import { DEFAULT_SETTINGS, PIN_KEYWORD, categoriesOf, type AppSettings, type QuickStep, type SavedSearch, type Task } from '../core/settings';
import { ALL, AccountRail } from './components/AccountRail';
import { AccountDialog } from './components/AccountDialog';
import { Composer, type ComposeInit, type ComposeKind, type ComposerHandle, type ComposerState } from './components/Composer';
import type { Contact } from './components/RecipientField';
import {
  M365View, MODULE_TITLES, baseKey, initialTabs, moduleUrl, tabTitle, type TabsState,
} from './components/M365View';
import { AddAccountChooser, M365Dialog } from './components/M365Dialog';
import { SignaturesDialog } from './components/SignaturesDialog';
import { TemplatesDialog } from './components/TemplatesDialog';
import { ThemeDialog } from './components/ThemeDialog';
import { CryptoDialog } from './components/CryptoDialog';
import { AboutDialog, ShortcutsDialog } from './components/AboutDialog';
import { ZimbraDialog, type ZimbraTab } from './components/ZimbraDialog';
import { OrganizeDialog, type OrganizeTab } from './components/OrganizeDialog';
import { FolderPane, MessageList, ReadingPane, type Selection } from './components/MailPanes';
import { RightPanel, type AgendaData } from './components/RightPanel';
import { CalendarView } from './components/CalendarView';
import { ContactsView } from './components/ContactsView';
import { TasksView, dueLabel, todayIso } from './components/TasksView';
import {
  IMAP_MODULES, M365_MODULES, ModuleRail, Ribbon, TopBar, type ImapModule, type M365Module, type RibbonTab,
} from './components/Shell';
import { DEMO_ID, demoAccount, demoFolders, demoMessages } from './mock';
import { accountColor, getTheme, themeById, themeVars, type ThemeKey } from './theme';

declare global {
  interface Window {
    mailApi?: MailApi;
  }
}

function usePersisted<T>(key: string, initial: T) {
  const [v, setV] = useState<T>(() => {
    try {
      const s = localStorage.getItem(key);
      return s ? (JSON.parse(s) as T) : initial;
    } catch {
      return initial;
    }
  });
  useEffect(() => {
    try {
      localStorage.setItem(key, JSON.stringify(v));
    } catch {
      /* nevadí */
    }
  }, [key, v]);
  return [v, setV] as const;
}

type FolderState = FolderInfo[] | 'loading' | { error: string };
type Dialogs =
  | { kind: 'none' }
  | { kind: 'chooser' }
  | { kind: 'imap'; edit?: ImapAccountPublic }
  | { kind: 'm365'; edit?: M365Account }
  | { kind: 'signatures' }
  | { kind: 'organize'; tab?: OrganizeTab }
  | { kind: 'templates' }
  | { kind: 'themes' }
  | { kind: 'crypto' }
  | { kind: 'about' }
  | { kind: 'shortcuts' }
  | { kind: 'zimbra'; tab?: ZimbraTab };

const splitAddrs = (s: string) =>
  s
    .split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/)
    .map((x) => x.trim())
    .filter(Boolean);

/** Prototyp v okne Tauri (WebView2). */
const IS_TAURI = '__TAURI_INTERNALS__' in window;

/** Dymový test v CI: pravý panel otvorený, aby bolo vidieť aj vložené weby. */
const SMOKE_MODE = new URLSearchParams(window.location.search).get('smoke') ?? '';
const SMOKE = SMOKE_MODE === 'm365';
// samostatné okno písania (koncept otvorený z hlavného okna)
const POP = new URLSearchParams(window.location.search);
const POP_COMPOSE = POP.get('compose') === '1' ? { acc: POP.get('acc') ?? '', folder: POP.get('folder') ?? '', uid: Number(POP.get('uid')) } : null;

export function App() {
  const api = window.mailApi;
  const [theme, setTheme] = usePersisted<ThemeKey>('theme', 'blue');
  const [compact, setCompact] = usePersisted('ribbonCompact', false);
  const [rightOpen, setRightOpen] = usePersisted('rightOpen', SMOKE);
  const [rightActive, setRightActive] = usePersisted('rightActive', 'copilot');
  const [activeId, setActiveId] = usePersisted<string>('activeAccount', '');
  const [ribbonTab, setRibbonTab] = useState<RibbonTab>('home');
  const [imapModule, setImapModule] = useState<Record<string, ImapModule>>({});
  const [tabs, setTabs] = useState<Record<string, TabsState>>({});
  const [settings, setSettings] = useState<AppSettings>(DEFAULT_SETTINGS);
  const themeDef = themeById(theme, settings.customThemes);

  const [accounts, setAccounts] = useState<AccountPublic[]>([]);
  const [loaded, setLoaded] = useState(false);
  const [secretsWarning, setSecretsWarning] = useState(false);
  const [unread, setUnread] = useState<Record<string, number>>({});
  const [folders, setFolders] = useState<Record<string, FolderState>>({});
  const [selection, setSelection] = useState<Selection>({ accountId: ALL, folder: 'INBOX' });
  const [messages, setMessages] = useState<MessageSummary[]>([]);
  const [listState, setListState] = useState<{ loading: boolean; error?: string }>({ loading: false });
  const [current, setCurrent] = useState<MessageSummary | undefined>();
  const [full, setFull] = useState<{ msg?: MessageFull; loading: boolean; error?: string }>({ loading: false });
  const [search, setSearch] = useState('');
  const [refreshKey, setRefreshKey] = useState(0);
  // nová pošta (IDLE): obnoviť zoznam a počty bez načítania priečinkov
  const [mailTick, setMailTick] = useState(0);
  const pendingOpen = useRef<{ accountId: string; uid: number } | null>(null);
  // odoslanie čakajúce 10 s na späťvzatie; naplánované správy
  const [undo, setUndo] = useState<OutboxItem | null>(null);
  const [updateReady, setUpdateReady] = useState<string | null>(null);
  const [followupEv, setFollowupEv] = useState<{ subject: string; to: string; accountId: string } | null>(null);
  const [searchAll, setSearchAll] = useState(false);
  const [outboxTick, setOutboxTick] = useState(0);
  const [scheduled, setScheduled] = useState<OutboxItem[]>([]);
  const [status, setStatus] = useState('');
  const [dialog, setDialog] = useState<Dialogs>({ kind: 'none' });
  const [dialogKey, setDialogKey] = useState(0);
  const [panelOverlay, setPanelOverlay] = useState(false);
  const [menuOverlay, setMenuOverlay] = useState(false);
  const msgCache = useRef(new Map<string, MessageFull>());
  const [composeView, setCompose] = useState<{ init: ComposeInit; key: number } | null>(null);
  const composerRef = useRef<ComposerHandle>(null);
  const [composerState, setComposerState] = useState<ComposerState>({ importance: false, receipt: false, savedAt: null, sending: false });
  const openKey = useRef('');
  const [gpuOff, setGpuOff] = useState(false);
  const [gpuForced, setGpuForced] = useState(false);
  const [gpuCrash, setGpuCrash] = useState(0);

  // stabilita vykresľovania: stav akcelerácie, pády GPU, chyby rozhrania do záznamu
  useEffect(() => {
    if (!api) return;
    void api.gpuDisabled().then((r) => r.ok && setGpuOff(r.data));
    void api.gpuForced().then((r) => r.ok && setGpuForced(r.data));
    const onErr = (e: ErrorEvent) => void api.logError(`${e.message} @ ${e.filename}:${e.lineno}`);
    const onRej = (e: PromiseRejectionEvent) => void api.logError(`promise: ${String(e.reason)}`);
    window.addEventListener('error', onErr);
    window.addEventListener('unhandledrejection', onRej);
    let tickTimer = 0;
    const off = api.onAppEvent((ev) => {
      if (ev.type === 'gpu-crash') setGpuCrash(ev.count);
      else if (ev.type === 'mail-changed') {
        window.clearTimeout(tickTimer);
        tickTimer = window.setTimeout(() => setMailTick((t) => t + 1), 800);
      } else if (ev.type === 'outbox-sent') {
        setUndo((u) => (u?.id === ev.id ? null : u));
        setStatus(ev.savedToSent ? `Odoslané: ${ev.subject} (uložené do Odoslaných).` : `Odoslané: ${ev.subject}.`);
        setOutboxTick((t) => t + 1);
      } else if (ev.type === 'outbox-failed') {
        setUndo((u) => (u?.id === ev.id ? null : u));
        setStatus(`Správu „${ev.subject}“ sa nepodarilo odoslať: ${ev.error} Skúsim znova o 5 minút.`);
        setOutboxTick((t) => t + 1);
      } else if (ev.type === 'update') {
        if (ev.state === 'ready') setUpdateReady(ev.version || 'nová verzia');
        else if (ev.state === 'downloading') setStatus('Sťahujem novú verziu MailKlienta na pozadí…');
      } else if (ev.type === 'followup') {
        setFollowupEv(ev);
      } else if (ev.type === 'progress') {
        setStatus(ev.text);
      } else if (ev.type === 'outbox-changed') {
        setOutboxTick((t) => t + 1);
      } else if (ev.type === 'open-message') {
        pendingOpen.current = { accountId: ev.accountId, uid: ev.uid };
        setActiveId(ev.accountId);
        setSelection({ accountId: ev.accountId, folder: 'INBOX' });
        setMailTick((t) => t + 1);
      }
    });
    // dlhé zaseknutia rozhrania (> 300 ms) do záznamu, aby sa dalo dohľadať blikanie
    let po: PerformanceObserver | undefined;
    try {
      po = new PerformanceObserver((list) => {
        for (const e of list.getEntries()) if (e.duration > 300) void api.logError(`zaseknutie rozhrania ${Math.round(e.duration)} ms`);
      });
      po.observe({ type: 'longtask', buffered: false });
    } catch {
      /* nepodporované */
    }
    return () => {
      window.removeEventListener('error', onErr);
      window.removeEventListener('unhandledrejection', onRej);
      po?.disconnect();
      off();
    };
  }, [api]);

  const demo = loaded && accounts.length === 0;
  const railAccounts: AccountPublic[] = demo ? [demoAccount] : accounts;
  const imapAccounts = useMemo(() => (demo ? [demoAccount] : accounts.filter(isImap)), [accounts, demo]);
  const overlay = dialog.kind !== 'none' || panelOverlay || menuOverlay;

  // aktívne konto (alebo „Všetky“)
  const active: AccountPublic | null = railAccounts.find((a) => a.id === activeId) ?? null;
  const activeKey = active ? active.id : imapAccounts.length > 1 && activeId === ALL ? ALL : railAccounts[0]?.id ?? ALL;
  const activeAcc = railAccounts.find((a) => a.id === activeKey) ?? null;
  const activeImap = activeAcc && isImap(activeAcc) ? activeAcc : null;
  const activeM365 = activeAcc && isM365(activeAcc) ? activeAcc : null;
  const colorOf = useCallback((id: string) => accountColor(railAccounts.find((a) => a.id === id)?.color), [railAccounts]);

  const openDialog = (d: Dialogs) => {
    setDialog(d);
    setDialogKey((k) => k + 1);
  };
  const closeDialog = () => setDialog({ kind: 'none' });

  // kontá a nastavenia
  const loadAccounts = useCallback(async () => {
    if (!api) {
      setLoaded(true);
      return;
    }
    const r = await api.listAccounts();
    if (r.ok) {
      setAccounts(r.data.accounts);
      setSecretsWarning(!r.data.canStoreSecrets && r.data.accounts.some(isImap));
    } else setStatus(r.error);
    setLoaded(true);
  }, [api]);
  useEffect(() => {
    void loadAccounts();
    void api?.getSettings().then((r) => r.ok && setSettings(r.data));
  }, [loadAccounts, api]);

  const saveSettings = (s: AppSettings) => {
    setSettings(s);
    void api?.saveSettings(s);
  };

  // pri zmene konta: výber priečinka
  useEffect(() => {
    if (activeImap) setSelection({ accountId: activeImap.id, folder: 'INBOX' });
    else if (activeKey === ALL) setSelection({ accountId: ALL, folder: 'INBOX' });
    setCurrent(undefined);
    setFull({ loading: false });
  }, [activeKey]); // eslint-disable-line react-hooks/exhaustive-deps

  // karty kont M365
  useEffect(() => {
    if (activeM365 && !tabs[activeM365.id]) setTabs((t) => ({ ...t, [activeM365.id]: initialTabs(activeM365) }));
  }, [activeM365, tabs]);

  // udalosti vložených webov: nová karta z Outlooku (Office aplikácie), názvy stránok
  useEffect(() => {
    if (!api) return;
    return api.onWebEvent((ev) => {
      if (ev.type === 'tab-opened') {
        const accId = ev.group.replace(/^acc-/, '');
        let host = 'Office';
        try {
          host = new URL(ev.url).hostname.replace(/^www\./, '');
        } catch {
          /* nič */
        }
        setTabs((all) => {
          const st = all[accId];
          if (!st) return all;
          return { ...all, [accId]: { tabs: [...st.tabs, { key: ev.key, title: host, url: ev.url }], active: ev.key } };
        });
        setActiveId(accId);
      } else if (ev.type === 'title') {
        setTabs((all) => {
          let changed = false;
          const next: Record<string, TabsState> = {};
          for (const [id, st] of Object.entries(all)) {
            next[id] = {
              ...st,
              tabs: st.tabs.map((t) => {
                if (t.key !== ev.key) return t;
                changed = true;
                return { ...t, title: tabTitle(ev.title, t.module ? MODULE_TITLES[t.module] : t.title) };
              }),
            };
          }
          return changed ? next : all;
        });
      }
    });
  }, [api, setActiveId]);

  // neprečítané pre pás kont (každé 2 minúty)
  useEffect(() => {
    if (!api || demo) return;
    let alive = true;
    const run = () => {
      for (const a of accounts.filter(isImap))
        void api.unreadCount(a.id).then((r) => alive && r.ok && setUnread((u) => ({ ...u, [a.id]: r.data })));
    };
    run();
    // počty sa menia hlavne udalosťou z IDLE; pravidelná kontrola len ako záloha
    const t = window.setInterval(run, 600_000);
    return () => {
      alive = false;
      window.clearInterval(t);
    };
  }, [api, accounts, demo, refreshKey, mailTick]);

  // priečinky aktívneho IMAP konta
  useEffect(() => {
    if (demo) {
      setFolders({ [DEMO_ID]: demoFolders });
      return;
    }
    if (!activeImap || !api) return;
    const id = activeImap.id;
    setFolders((f) => (Array.isArray(f[id]) ? f : { ...f, [id]: 'loading' }));
    void api.listFolders(id).then((r) => setFolders((f) => ({ ...f, [id]: r.ok ? r.data : { error: r.error } })));
  }, [activeImap?.id, demo, api, refreshKey]); // eslint-disable-line react-hooks/exhaustive-deps

  // správy
  const showMail = !activeM365 && (activeImap ? (imapModule[activeImap.id] ?? 'mail') === 'mail' : true);
  useEffect(() => {
    if (!showMail) return;
    let cancelled = false;
    (async () => {
      if (demo) {
        setMessages(selection.folder === 'INBOX' ? demoMessages : []);
        return;
      }
      if (!api) return;
      const ids = accounts.filter(isImap).map((a) => a.id);
      const targets =
        selection.accountId === ALL ? ids.map((id) => ({ id, folder: 'INBOX' })) : ids.includes(selection.accountId) ? [{ id: selection.accountId, folder: selection.folder }] : [];
      if (!targets.length) {
        setMessages([]);
        return;
      }
      const byDate = (a: MessageSummary, b: MessageSummary) => (b.date ?? '').localeCompare(a.date ?? '');
      // 1) okamžite z lokálnej cache (aj offline)
      const cached = await Promise.all(targets.map((t) => api.listCached(t.id, t.folder, 50)));
      if (cancelled) return;
      const fromCache = cached.flatMap((r) => (r.ok ? r.data : [])).sort(byDate);
      if (fromCache.length && !mailTick) setMessages(fromCache);
      if (!mailTick || !messages.length) setListState({ loading: true });
      // 2) aktuálny stav zo servera (uloží sa aj do cache)
      const results = await Promise.all(targets.map((t) => api.listMessages(t.id, t.folder, 50, !!settings.previewLine)));
      if (cancelled) return;
      const errors = results.flatMap((r, i) => (r.ok ? [] : [`${accounts.find((a) => a.id === targets[i]!.id)?.name}: ${r.error}`]));
      const all = results.flatMap((r, i) => (r.ok ? r.data : cached[i]?.ok ? cached[i].data : [])).sort(byDate);
      setMessages(all);
      const offline = errors.length && all.length;
      setListState({ loading: false, error: errors.length ? (offline ? `Offline – zobrazené uložené správy. ${errors.join(' · ')}` : errors.join(' · ')) : undefined });
    })();
    return () => {
      cancelled = true;
    };
  }, [selection, accounts, demo, api, refreshKey, showMail, mailTick]); // eslint-disable-line react-hooks/exhaustive-deps

  // otvorenie správy z upozornenia Windows
  useEffect(() => {
    const p = pendingOpen.current;
    if (!p) return;
    const m = messages.find((x) => x.accountId === p.accountId && x.uid === p.uid);
    if (m) {
      pendingOpen.current = null;
      void open(m);
    }
  }, [messages]); // eslint-disable-line react-hooks/exhaustive-deps

  const patchMsg = (m: MessageSummary, patch: Partial<MessageSummary>) => {
    const same = (x: MessageSummary) => x.uid === m.uid && x.accountId === m.accountId && x.folder === m.folder;
    setMessages((list) => list.map((x) => (same(x) ? { ...x, ...patch } : x)));
    setCurrent((c) => (c && same(c) ? { ...c, ...patch } : c));
    setFull((f) => (f.msg && same(f.msg) ? { ...f, msg: { ...f.msg, ...patch } } : f));
    const ck = `${m.accountId}|${m.folder}|${m.uid}`;
    const c = msgCache.current.get(ck);
    if (c) msgCache.current.set(ck, { ...c, ...patch });
  };

  async function open(m: MessageSummary) {
    if (composeView) {
      setCompose(null); // koncept sa priebežne ukladá do Konceptov
      setRibbonTab('home');
    }
    setCurrent(m);
    if (m.accountId === DEMO_ID) {
      setFull({ loading: false, msg: demoMessages.find((x) => x.uid === m.uid) });
      return;
    }
    const ck = `${m.accountId}|${m.folder}|${m.uid}`;
    openKey.current = ck;
    const cached = msgCache.current.get(ck);
    if (cached) {
      setFull({ loading: false, msg: { ...cached, keywords: m.keywords, answered: m.answered } });
      if (m.seen) return;
    } else setFull((f) => ({ loading: true, msg: f.msg })); // predchádzajúca správa ostáva viditeľná
    const r = cached ? ({ ok: true, data: cached } as const) : await api!.getMessage(m.accountId, m.folder, m.uid);
    if (r.ok) {
      msgCache.current.set(ck, r.data);
      if (msgCache.current.size > 200) msgCache.current.delete(msgCache.current.keys().next().value!);
    }
    if (openKey.current !== ck) return; // medzitým klikol na inú správu
    setFull(r.ok ? { loading: false, msg: { ...r.data, keywords: m.keywords, answered: m.answered } } : { loading: false, error: `${r.error}${'hint' in r && r.hint ? ` ${r.hint}` : ''}` });
    if (r.ok && !m.seen) {
      patchMsg(m, { seen: true });
      setUnread((u) => ({ ...u, [m.accountId]: Math.max(0, (u[m.accountId] ?? 1) - 1) }));
    }
  }

  // operácie so správou
  const removeFromList = (m: MessageSummary) => {
    setMessages((list) => list.filter((x) => !(x.uid === m.uid && x.accountId === m.accountId && x.folder === m.folder)));
    setCurrent(undefined);
    setFull({ loading: false });
  };
  const guardDemo = (m: MessageSummary) => {
    if (m.accountId === DEMO_ID) {
      setStatus('V ukážkovom režime sa správy nemenia. Pridaj konto (+ v páse kont).');
      return true;
    }
    return false;
  };
  async function del() {
    const m = current;
    if (!m || !api || guardDemo(m)) return;
    const r = await api.deleteMessages(m.accountId, m.folder, [m.uid]);
    if (!r.ok) return setStatus(r.error);
    removeFromList(m);
    setStatus(r.data.movedTo ? 'Správa presunutá do Odstránených.' : 'Správa natrvalo odstránená.');
  }
  async function archive() {
    const m = current;
    if (!m || !api || guardDemo(m)) return;
    const r = await api.archiveMessages(m.accountId, m.folder, [m.uid]);
    if (!r.ok) return setStatus(r.error);
    removeFromList(m);
    setStatus(`Správa archivovaná (${r.data.movedTo}).`);
  }
  async function move(dest: string) {
    const m = current;
    if (!m || !api || guardDemo(m)) return;
    const r = await api.moveMessages(m.accountId, m.folder, [m.uid], dest);
    if (!r.ok) return setStatus(r.error);
    removeFromList(m);
    setStatus('Správa presunutá.');
  }
  async function toggleSeen() {
    const m = current;
    if (!m || !api || guardDemo(m)) return;
    const seen = !m.seen;
    const r = await api.setFlags(m.accountId, m.folder, [m.uid], { seen });
    if (!r.ok) return setStatus(r.error);
    patchMsg(m, { seen });
    setUnread((u) => ({ ...u, [m.accountId]: Math.max(0, (u[m.accountId] ?? 0) + (seen ? -1 : 1)) }));
  }
  async function toggleFlag() {
    const m = current;
    if (!m || !api || guardDemo(m)) return;
    const flagged = !m.flagged;
    const r = await api.setFlags(m.accountId, m.folder, [m.uid], { flagged });
    if (!r.ok) return setStatus(r.error);
    patchMsg(m, { flagged });
  }

  async function setCategory(id: string, on: boolean) {
    const m = current;
    if (!m || !api || guardDemo(m)) return;
    const r = await api.setKeywords(m.accountId, m.folder, [m.uid], on ? [id] : [], on ? [] : [id]);
    if (!r.ok) return setStatus(r.error);
    const kw = new Set(m.keywords ?? []);
    if (on) kw.add(id);
    else kw.delete(id);
    patchMsg(m, { keywords: [...kw] });
  }
  async function togglePin() {
    const m = current;
    if (!m) return;
    await setCategory(PIN_KEYWORD, !m.keywords?.includes(PIN_KEYWORD));
  }
  async function snoozeCurrent(until: Date) {
    const m = current;
    const mf = full.msg;
    if (!m || !api || guardDemo(m)) return;
    const r = await api.snooze(m.accountId, m.folder, m.uid, mf?.messageId ?? m.messageId ?? '', m.subject, until.toISOString());
    if (!r.ok) return setStatus(r.error);
    removeFromList(m);
    setStatus(`Správa odložená do ${until.toLocaleString('sk-SK', { weekday: 'short', day: 'numeric', month: 'numeric', hour: '2-digit', minute: '2-digit' })} (priečinok Odložené).`);
  }
  async function runQuickStep(q: QuickStep) {
    const m = current;
    if (!m || !api || guardDemo(m)) return;
    const r = await api.runActions(m.accountId, m.folder, [m.uid], q.actions);
    if (!r.ok) return setStatus(r.error);
    if (q.actions.some((a) => ['move', 'archive', 'delete'].includes(a.type))) removeFromList(m);
    else setMailTick((t) => t + 1);
    setStatus(`Rýchly krok „${q.name}“ vykonaný.`);
  }
  async function dropMessages(dest: string, d: { accountId: string; folder: string; uids: number[] }) {
    if (!api || d.accountId === DEMO_ID) return;
    const r = await api.moveMessages(d.accountId, d.folder, d.uids, dest);
    if (!r.ok) return setStatus(r.error);
    setMessages((list) => list.filter((x) => !(x.accountId === d.accountId && x.folder === d.folder && d.uids.includes(x.uid))));
    if (current && current.accountId === d.accountId && d.uids.includes(current.uid)) {
      setCurrent(undefined);
      setFull({ loading: false });
    }
    setStatus('Správa presunutá.');
  }
  function saveCurrentSearch() {
    const q = search.trim();
    if (!q) return;
    const s2: SavedSearch = { id: Math.random().toString(36).slice(2, 10), name: q, query: q, accountId: selection.accountId === ALL ? undefined : selection.accountId };
    saveSettings({ ...settings, savedSearches: [...(settings.savedSearches ?? []), s2] });
    setStatus(`Hľadanie „${q}“ uložené medzi priečinky.`);
  }

  // konverzácia otvorenej správy (z cache naprieč priečinkami, napr. aj Odoslané)
  const [thread, setThread] = useState<MessageSummary[]>([]);
  useEffect(() => {
    const id = full.msg?.messageId;
    if (!api || !full.msg || !id || full.msg.accountId === DEMO_ID) return setThread([]);
    let alive = true;
    void api.thread(full.msg.accountId, id).then((r) => alive && setThread(r.ok ? r.data : []));
    return () => {
      alive = false;
    };
  }, [full.msg?.messageId, full.msg?.accountId]); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    if (!api) return;
    void api.outboxList().then((r) => r.ok && setScheduled(r.data));
  }, [api, outboxTick]);

  /** Vrátené alebo naplánované odoslanie znova na úpravu. */
  function reopenFromOutbox(it: OutboxItem) {
    const i = it.input;
    const init: ComposeInit = {
        kind: 'new',
        accountId: it.accountId,
        to: i.to,
        cc: i.cc,
        bcc: i.bcc,
        subject: i.subject,
        html: i.html ?? `<p>${i.text}</p>`,
        draft: i.draft,
        files: (i.attachments ?? []).map((a) => ({ path: a.path, filename: a.filename ?? a.path.split(/[\\/]/).pop() ?? 'príloha', size: 0 })),
        importance: i.priority === 'high',
        receipt: i.readReceipt,
        fromEmail: i.from?.email,
        security: i.security,
    };
    if (POP_COMPOSE) {
      setPopInit(init);
      setPopDone(false);
      return;
    }
    setCompose({ init, key: Date.now() });
    setRibbonTab('message');
  }
  async function undoSend() {
    if (!undo || !api) return;
    const r = await api.outboxCancel(undo.id);
    setUndo(null);
    if (r.ok && r.data) {
      reopenFromOutbox(r.data);
      setStatus('Odoslanie zrušené, správa je znova otvorená.');
    } else setStatus('Správa už bola odoslaná.');
  }
  async function cancelScheduled(it: OutboxItem, edit: boolean) {
    if (!api) return;
    const r = await api.outboxCancel(it.id);
    if (r.ok && r.data && edit) reopenFromOutbox(r.data);
    setStatus(edit ? 'Naplánovaná správa otvorená na úpravu.' : 'Naplánované odoslanie zrušené (koncept ostal v Konceptoch).');
  }
  useEffect(() => {
    if (!undo) return;
    const t = window.setTimeout(() => setUndo(null), 11_000);
    return () => window.clearTimeout(t);
  }, [undo]);

  /** Koncept z priečinka Koncepty na úpravu (s prílohami). */
  async function editDraft(m: MessageSummary) {
    if (!api) return;
    setStatus('Otváram koncept…');
    const r = await api.openDraft(m.accountId, m.folder, m.uid);
    if (!r.ok) return setStatus(r.error);
    const d = r.data.msg;
    setCompose({
      init: { kind: 'new', accountId: m.accountId, to: d.to, cc: d.cc, subject: d.subject === '(bez predmetu)' ? '' : d.subject, html: d.html ?? `<p>${d.text}</p>`, draft: { folder: m.folder, uid: m.uid }, files: r.data.files },
      key: Date.now(),
    });
    setRibbonTab('message');
    setStatus('');
  }
  const isDrafts = (m: MessageSummary) => {
    const fs = folders[m.accountId];
    const f = Array.isArray(fs) ? fs.find((x) => x.path === m.folder) : undefined;
    return f?.specialUse === '\\Drafts' || /^(drafts|koncept)/i.test(m.folder);
  };

  function saveTemplate(t: { name: string; subject: string; html: string }) {
    const tpl = { id: Math.random().toString(36).slice(2, 10), ...t };
    saveSettings({ ...settings, templates: [...(settings.templates ?? []), tpl] });
    setStatus(`Šablóna „${t.name}“ uložená.`);
  }

  /** Úloha zo správy (ako vlajka s termínom v Outlooku); správa sa zároveň označí vlajkou. */
  async function taskFromMessage(due?: string) {
    const m = current;
    if (!m) return;
    const t: Task = {
      id: Math.random().toString(36).slice(2, 10),
      title: m.subject,
      due,
      done: false,
      created: new Date().toISOString(),
      message: { accountId: m.accountId, folder: m.folder, uid: m.uid, messageId: full.msg?.messageId ?? m.messageId, subject: m.subject, from: m.fromName || m.fromAddress },
    };
    saveSettings({ ...settings, tasks: [...(settings.tasks ?? []), t] });
    if (!m.flagged && m.accountId !== DEMO_ID) await toggleFlag();
    setStatus(`Úloha vytvorená${due ? ` s termínom ${dueLabel(due)}` : ''}.`);
  }
  function openTaskMessage(ref: NonNullable<Task['message']>) {
    setActiveId(ref.accountId);
    setImapModule((x) => ({ ...x, [ref.accountId]: 'mail' }));
    pendingOpen.current = { accountId: ref.accountId, uid: ref.uid };
    setSelection({ accountId: ref.accountId, folder: ref.folder });
  }

  // Agenda (pravý panel): dnešné stretnutia zo všetkých IMAP kont + úlohy na dnes a po termíne
  const [agendaTick, setAgendaTick] = useState(0);
  const [agendaEvents, setAgendaEvents] = useState<{ loading: boolean; error?: string; list: AgendaData['events'] }>({ loading: false, list: [] });
  const agendaVisible = rightOpen && rightActive === 'agenda';
  useEffect(() => {
    if (!api || demo || !agendaVisible) return;
    let alive = true;
    const run = async () => {
      const d = new Date();
      const from = new Date(d.getFullYear(), d.getMonth(), d.getDate());
      const to = new Date(from.getTime() + 86400000);
      setAgendaEvents((x) => ({ ...x, loading: true }));
      const accs = accounts.filter(isImap);
      const rs = await Promise.all(accs.map((a) => api.calEvents(a.id, [], from.toISOString(), to.toISOString())));
      if (!alive) return;
      const list = rs
        .flatMap((r, i) => (r.ok ? r.data.map((e) => ({ ...e, accountName: accs[i]!.name, accountColor: accountColor(accs[i]!.color) })) : []))
        .sort((a, b) => Number(b.allDay) - Number(a.allDay) || Date.parse(a.start) - Date.parse(b.start));
      const errs = rs.filter((r) => !r.ok).length;
      setAgendaEvents({ loading: false, list, error: errs && !list.length ? 'Kalendár sa nepodarilo načítať.' : undefined });
    };
    void run();
    const t = window.setInterval(() => void run(), 10 * 60_000);
    return () => {
      alive = false;
      window.clearInterval(t);
    };
  }, [api, accounts, demo, agendaVisible, agendaTick]);
  const agenda: AgendaData = {
    loading: agendaEvents.loading,
    error: agendaEvents.error,
    events: agendaEvents.list,
    tasks: (settings.tasks ?? []).filter((t) => !t.done && t.due && t.due <= todayIso()),
    onOpenTask: (t) => (t.message ? openTaskMessage(t.message) : activeImap && setImapModule((x) => ({ ...x, [activeImap.id]: 'tasks' }))),
    onOpenCalendar: () => activeImap && setImapModule((x) => ({ ...x, [activeImap.id]: 'calendar' })),
    onToggleTask: (t) => saveSettings({ ...settings, tasks: (settings.tasks ?? []).map((x) => (x.id === t.id ? { ...x, done: !x.done } : x)) }),
  };

  // pripomenutie bez odpovede → úloha na dnes
  useEffect(() => {
    if (!followupEv) return;
    const t: Task = { id: Math.random().toString(36).slice(2, 10), title: `Bez odpovede: ${followupEv.subject} (${followupEv.to})`, due: todayIso(), done: false, created: new Date().toISOString() };
    saveSettings({ ...settings, tasks: [...(settings.tasks ?? []), t] });
    setStatus(`${followupEv.to} neodpovedal(a) na „${followupEv.subject}“ – pridaná úloha na dnes.`);
    setFollowupEv(null);
  }, [followupEv]); // eslint-disable-line react-hooks/exhaustive-deps

  // počet neprečítaných: ikona pri hodinách a odznak na paneli úloh
  const totalUnread = Object.values(unread).reduce((a, b) => a + (b || 0), 0);
  useEffect(() => {
    if (!api || POP_COMPOSE) return;
    let badge: string | undefined;
    if (totalUnread > 0) {
      const c = document.createElement('canvas');
      c.width = c.height = 32;
      const x = c.getContext('2d');
      if (x) {
        x.fillStyle = '#c50f1f';
        x.beginPath();
        x.arc(16, 16, 15, 0, Math.PI * 2);
        x.fill();
        x.fillStyle = '#fff';
        x.font = `bold ${totalUnread > 99 ? 13 : 18}px Segoe UI, Arial`;
        x.textAlign = 'center';
        x.textBaseline = 'middle';
        x.fillText(totalUnread > 99 ? '99+' : String(totalUnread), 16, 17);
        badge = c.toDataURL('image/png');
      }
    }
    void api.setUnread(totalUnread, badge);
  }, [api, totalUnread]);

  // klávesové skratky ako v Outlooku
  const keys = useRef<(e: KeyboardEvent) => void>(() => undefined);
  keys.current = (e: KeyboardEvent) => {
    const t = e.target as HTMLElement | null;
    const typing = !!t && (t.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName));
    const ctrl = e.ctrlKey || e.metaKey;
    const k = e.key.toLowerCase();
    if (ctrl && e.key === 'Enter' && composeView) return e.preventDefault(), composerRef.current?.send();
    if (ctrl && k === 'e') return e.preventDefault(), document.querySelector<HTMLInputElement>('.topbar-search input')?.focus();
    if (typing || dialog.kind !== 'none') return;
    if (ctrl && k === 'n') return e.preventDefault(), compose('new');
    if (ctrl && e.shiftKey && k === 'r') return e.preventDefault(), full.msg && compose('replyAll');
    if (ctrl && k === 'r') return e.preventDefault(), full.msg && compose('reply');
    if (ctrl && k === 'f') return e.preventDefault(), full.msg && compose('forward');
    if (ctrl && k === 'q') return e.preventDefault(), current && !current.seen && void toggleSeen();
    if (ctrl && k === 'u') return e.preventDefault(), current && current.seen && void toggleSeen();
    if (ctrl && k === 'p') return e.preventDefault(), current && void api?.printMessage(current.accountId, current.folder, current.uid, false);
    if (e.key === 'Insert') return e.preventDefault(), current && void toggleFlag();
    if (e.key === 'Delete' && current) return e.preventDefault(), void del();
    if (ctrl && e.shiftKey && k === 'v') return e.preventDefault(), setStatus('Presun: použite Domov → Presunúť alebo pretiahnite správu na priečinok.');
    if (e.key === 'F9') return e.preventDefault(), setRefreshKey((x) => x + 1);
    if (ctrl && ['1', '2', '3', '4'].includes(e.key) && activeImap) {
      e.preventDefault();
      const mod = (['mail', 'calendar', 'contacts', 'tasks'] as ImapModule[])[Number(e.key) - 1]!;
      setImapModule((x) => ({ ...x, [activeImap.id]: mod }));
    }
  };
  useEffect(() => {
    const h = (e: KeyboardEvent) => keys.current(e);
    window.addEventListener('keydown', h);
    return () => window.removeEventListener('keydown', h);
  }, []);

  /** Export a import pošty, aktualizácie (karta Domov → Súbor). */
  async function runIo(k: 'eml' | 'mbox' | 'import' | 'thunderbird' | 'update' | 'print' | 'pdf' | 'backup' | 'restore' | 'shortcuts') {
    if (!api) return;
    if (k === 'shortcuts') return openDialog({ kind: 'shortcuts' });
    if (k === 'backup') {
      const r = await api.backupExport();
      return setStatus(r.ok ? (r.data ? `Záloha uložená: ${r.data} (bez hesiel).` : '') : r.error);
    }
    if (k === 'restore') {
      const r = await api.backupImport();
      if (!r.ok) return setStatus(r.error);
      if (!r.data.restored) return;
      await loadAccounts();
      void api.getSettings().then((x) => x.ok && setSettings(x.data));
      return setStatus(`Nastavenia obnovené${r.data.accounts ? `, pridané kontá: ${r.data.accounts} – zadajte k nim heslá (Nastavenia konta)` : ''}.`);
    }
    if (k === 'print' || k === 'pdf') {
      if (!current || current.accountId === DEMO_ID) return setStatus('Vyberte správu.');
      const r = await api.printMessage(current.accountId, current.folder, current.uid, k === 'pdf');
      return setStatus(r.ok ? (k === 'pdf' && r.data ? `PDF uložené: ${r.data}` : '') : r.error);
    }
    const accId = current?.accountId ?? activeImap?.id;
    if (k === 'update') {
      const r = await api.updateCheck();
      return setStatus(r.ok && r.data ? 'Hľadám novú verziu…' : 'Aktualizácie fungujú len v nainštalovanej verzii pre Windows.');
    }
    if (!accId || accId === DEMO_ID) return setStatus('Najprv pridajte IMAP konto.');
    if (k === 'eml') {
      if (!current) return setStatus('Vyberte správu.');
      const r = await api.exportEml(current.accountId, current.folder, current.uid, current.subject);
      return setStatus(r.ok ? (r.data ? `Správa uložená: ${r.data}` : '') : r.error);
    }
    if (k === 'mbox') {
      const folder = selection.accountId === ALL ? 'INBOX' : selection.folder;
      setStatus('Exportujem priečinok…');
      const r = await api.exportMbox(accId, folder);
      return setStatus(r.ok ? (r.data.file ? `Exportované ${r.data.count} správ do ${r.data.file}` : '') : r.error);
    }
    if (k === 'import') {
      const r = await api.importFiles(accId, 'Import');
      if (r.ok && r.data) setRefreshKey((x) => x + 1);
      return setStatus(r.ok ? (r.data ? `Importované ${r.data} správ do priečinka Import.` : '') : r.error);
    }
    const r = await api.importThunderbird(accId);
    if (r.ok && r.data.messages) setRefreshKey((x) => x + 1);
    setStatus(r.ok ? (r.data.folders ? `Z Thunderbirdu importovaných ${r.data.messages} správ v ${r.data.folders} priečinkoch (priečinok Thunderbird).` : '') : `${r.error}${r.hint ? ` ${r.hint}` : ''}`);
  }

  /** Písanie v čítacom paneli (ako Outlook). */
  function compose(kind: ComposeKind) {
    const m = full.msg;
    const accountId = m && m.accountId !== DEMO_ID ? m.accountId : activeImap && activeImap.id !== DEMO_ID ? activeImap.id : imapAccounts.find((a) => a.id !== DEMO_ID)?.id;
    if (activeImap) setImapModule((x) => ({ ...x, [activeImap.id]: 'mail' }));
    setRibbonTab('message');
    if (kind === 'new' || !m) {
      setCompose({ init: { kind: 'new', accountId }, key: Date.now() });
      return;
    }
    const me = (imapAccounts.find((a) => a.id === m.accountId)?.email ?? '').toLowerCase();
    const subject =
      kind === 'forward'
        ? /^(fw|fwd|pd):/i.test(m.subject) ? m.subject : `PD: ${m.subject}`
        : /^(re|odp):/i.test(m.subject) ? m.subject : `RE: ${m.subject}`;
    const cc =
      kind === 'replyAll'
        ? [...splitAddrs(m.to), ...splitAddrs(m.cc)]
            .filter((a) => !a.toLowerCase().includes(me) && !a.toLowerCase().includes(m.fromAddress.toLowerCase()))
            .join(', ')
        : '';
    setCompose({
      init: {
        kind,
        accountId,
        original: m,
        to: kind === 'forward' ? '' : m.fromName ? `"${m.fromName.replace(/"/g, '')}" <${m.fromAddress}>` : m.fromAddress,
        cc,
        subject,
      },
      key: Date.now(),
    });
  }

  // dymový test v CI: otvorená odpoveď na ukážkovú správu
  useEffect(() => {
    if (SMOKE_MODE !== 'compose' || !demo) return;
    const m = demoMessages[0]!;
    setCurrent(m);
    setFull({ loading: false, msg: m });
    setRibbonTab('message');
    setCompose({ init: { kind: 'reply', accountId: DEMO_ID, original: m, to: `"${m.fromName}" <${m.fromAddress}>`, subject: `RE: ${m.subject}` }, key: 1 });
  }, [demo]);

  // kontakty na dopĺňanie: odosielatelia načítaných správ + adresy, na ktoré si písal
  // ktoré kontá sú na serveri Zimbra (rozšírené funkcie cez jej rozhranie)
  const [zimbraOk, setZimbraOk] = useState<Record<string, boolean>>({});
  useEffect(() => {
    if (!api || demo) return;
    for (const a of accounts.filter(isImap)) void api.zimbraCheck(a.id).then((r) => setZimbraOk((z) => ({ ...z, [a.id]: r.ok && r.data })));
  }, [api, accounts, demo]);
  // obsadenosť schránky (Zimbra alebo IMAP QUOTA)
  const [quota, setQuota] = useState<Record<string, { used: number; limit: number } | null>>({});
  useEffect(() => {
    if (!api || demo) return;
    const run = () => {
      for (const a of accounts.filter(isImap)) void api.zimbraQuota(a.id).then((r) => r.ok && setQuota((q) => ({ ...q, [a.id]: r.data })));
    };
    run();
    const t = window.setInterval(run, 30 * 60_000);
    return () => window.clearInterval(t);
  }, [api, accounts, demo, refreshKey]);

  // adresáre CardDAV (napr. Zimbra) na dopĺňanie adries
  const [davContacts, setDavContacts] = useState<Record<string, Contact[]>>({});
  useEffect(() => {
    if (!api || demo) return;
    for (const a of accounts.filter(isImap))
      void api.cardList(a.id).then((r) => r.ok && setDavContacts((x) => ({ ...x, [a.id]: r.data.flatMap((c) => c.emails.map((email) => ({ name: c.name, email }))) })));
  }, [api, accounts, demo]);

  const contacts = useMemo<Contact[]>(() => {
    let known: Contact[] = [...Object.values(davContacts).flat()];
    try {
      known = [...(JSON.parse(localStorage.getItem('contacts') ?? '[]') as Contact[]), ...known];
    } catch {
      /* nič */
    }
    const all = [...known, ...messages.map((m) => ({ name: m.fromName, email: m.fromAddress }))].filter((c) => c.email);
    return all.filter((c, i) => all.findIndex((x) => x.email.toLowerCase() === c.email.toLowerCase()) === i);
  }, [messages, composeView, davContacts]);

  // hľadanie: fulltext v lokálnej cache (všetky uložené priečinky konta), na požiadanie aj na serveri
  const [found, setFound] = useState<{ q: string; list: MessageSummary[]; server?: 'loading' | 'done' } | null>(null);
  const searchScope = selection.accountId === ALL || searchAll ? undefined : selection.accountId;
  useEffect(() => {
    const q = search.trim();
    if (!api || demo || q.length < 2) {
      setFound(null);
      return;
    }
    const t = window.setTimeout(() => {
      void api.searchLocal(q, searchScope).then((r) => r.ok && setFound({ q, list: r.data }));
    }, 250);
    return () => window.clearTimeout(t);
  }, [search, searchScope, api, demo, mailTick]);

  async function searchOnServer() {
    if (!api || !found) return;
    const q = found.q;
    setFound({ ...found, server: 'loading' });
    const targets = searchScope ? [{ id: searchScope, folder: selection.folder }] : accounts.filter(isImap).map((a) => ({ id: a.id, folder: 'INBOX' }));
    const rs = await Promise.all(targets.map((t) => api.searchServer(t.id, t.folder, q)));
    setFound((f) => {
      if (!f || f.q !== q) return f;
      const key = (m: MessageSummary) => `${m.accountId}|${m.folder}|${m.uid}`;
      const map = new Map(f.list.map((m) => [key(m), m]));
      for (const r of rs) if (r.ok) for (const m of r.data) map.set(key(m), m);
      return { q, list: [...map.values()].sort((a, b) => (b.date ?? '').localeCompare(a.date ?? '')), server: 'done' };
    });
  }

  const filtered = useMemo(() => {
    if (found) return found.list;
    const q = search.trim().toLowerCase();
    if (!q) return messages;
    return messages.filter((m) => `${m.subject} ${m.fromName} ${m.fromAddress}`.toLowerCase().includes(q));
  }, [messages, search, found]);

  const folderList = activeImap && Array.isArray(folders[activeImap.id]) ? (folders[activeImap.id] as FolderInfo[]) : [];
  const listTitle = useMemo(() => {
    if (selection.accountId === ALL) return 'Všetky doručené';
    const fs = folders[selection.accountId];
    const f = Array.isArray(fs) ? fs.find((x) => x.path === selection.folder) : undefined;
    return f?.name ?? (selection.folder === 'INBOX' ? 'Doručená pošta' : selection.folder);
  }, [selection, folders]);

  // modul M365: otvorí alebo prepne kartu
  function selectM365Module(acc: M365Account, mod: M365Module) {
    const st = tabs[acc.id] ?? initialTabs(acc);
    if (mod === 'mail') return setTabs((t) => ({ ...t, [acc.id]: { ...st, active: baseKey(acc.id) } }));
    const key = `${baseKey(acc.id)}:${mod}`;
    const exists = st.tabs.some((x) => x.key === key);
    setTabs((t) => ({
      ...t,
      [acc.id]: {
        tabs: exists ? st.tabs : [...st.tabs, { key, title: MODULE_TITLES[mod], url: moduleUrl(mod, acc.host), module: mod }],
        active: key,
      },
    }));
  }
  function closeTab(acc: M365Account, key: string) {
    setTabs((t) => {
      const st = t[acc.id];
      if (!st) return t;
      const rest = st.tabs.filter((x) => x.key !== key);
      return { ...t, [acc.id]: { tabs: rest, active: st.active === key ? baseKey(acc.id) : st.active } };
    });
  }

  const editAccount = (a: AccountPublic) => {
    if (a.id === DEMO_ID) return openDialog({ kind: 'chooser' });
    if (isImap(a)) openDialog({ kind: 'imap', edit: a });
    else if (isM365(a)) openDialog({ kind: 'm365', edit: a });
  };

  // horný pruh
  const accent = themeDef.accent;
  const topColor = themeDef.top ?? (activeAcc ? accountColor(activeAcc.color) : accent);
  const topTitle = activeAcc ? activeAcc.name : 'Všetky doručené';
  const topSub = activeM365
    ? `${activeM365.email || 'Microsoft 365'} · Outlook na webe`
    : activeImap
      ? `${activeImap.email} · IMAP${activeImap.imap.host ? ` · ${activeImap.imap.host}` : ''}`
      : `${imapAccounts.length} IMAP ${imapAccounts.length === 1 ? 'konto' : imapAccounts.length < 5 ? 'kontá' : 'kont'}`;

  const curImapModule: ImapModule = activeImap ? imapModule[activeImap.id] ?? 'mail' : 'mail';
  // samostatné okno písania: načítať koncept a zobraziť len editor
  const [popInit, setPopInit] = useState<ComposeInit | null>(null);
  const [popDone, setPopDone] = useState(false);
  useEffect(() => {
    if (POP_COMPOSE && popDone && !undo && !composeView) window.setTimeout(() => window.close(), 300);
  }, [popDone, undo, composeView]);
  useEffect(() => {
    if (!POP_COMPOSE || !api) return;
    void api.openDraft(POP_COMPOSE.acc, POP_COMPOSE.folder, POP_COMPOSE.uid).then((r) => {
      if (!r.ok) return setStatus(r.error);
      const d = r.data.msg;
      setPopInit({ kind: 'new', accountId: POP_COMPOSE.acc, to: d.to, cc: d.cc, subject: d.subject === '(bez predmetu)' ? '' : d.subject, html: d.html ?? `<p>${d.text}</p>`, draft: { folder: POP_COMPOSE.folder, uid: POP_COMPOSE.uid }, files: r.data.files });
    });
  }, [api]);
  if (POP_COMPOSE) {
    const imapOnly = accounts.filter(isImap);
    return (
      <FluentProvider theme={getTheme(themeDef)}>
        <div className={`app theme-${theme} pop-compose`} style={themeVars(themeDef)}>
          {popInit && imapOnly.length ? (
            <Composer
              init={popInit}
              accounts={imapOnly}
              settings={settings}
              contacts={contacts}
              zimbra={zimbraOk}
              standalone
              onState={() => undefined}
              onQueued={(it) => it.kind === 'undo' && setUndo(it)}
              onSaveTemplate={saveTemplate}
              onClose={(st) => {
                if (st) setStatus(st);
                setPopInit(null);
                setPopDone(true);
              }}
            />
          ) : popDone && undo ? (
            <div className="undo-bar pop" role="status">
              <span>Odosiela sa „{undo.subject || '(bez predmetu)'}“…</span>
              <button className="undo-btn" onClick={() => void undoSend()}>Späť</button>
              <span className="undo-timer" />
            </div>
          ) : (
            <div className="empty">{status || 'Otváram koncept…'}</div>
          )}
        </div>
      </FluentProvider>
    );
  }

  const m365State = activeM365 ? tabs[activeM365.id] ?? initialTabs(activeM365) : null;
  const m365ActiveModule = m365State ? m365State.tabs.find((t) => t.key === m365State.active)?.module ?? null : null;

  return (
    // Pozor: className FluentProvidera sa kopíruje aj na kontajnery vyskakovacích prvkov (tooltip, menu, dialóg)
    // pripojené k <body>. Trieda .app (celé okno, pozadie) tam spôsobovala zakrytie celého okna = blikanie.
    <FluentProvider theme={getTheme(themeDef)}>
      <div className={`app theme-${theme} ${themeDef.dark ? 'dark' : ''}`} style={themeVars(themeDef)}>
      <div className="shell">
        <AccountRail
          accounts={railAccounts}
          active={activeKey}
          unread={unread}
          showAll={imapAccounts.length > 1}
          onSelect={setActiveId}
          onAdd={() => openDialog({ kind: 'chooser' })}
          onEdit={editAccount}
          onReorder={(ids) => {
            setAccounts((list) => [...list].sort((a, b) => ids.indexOf(a.id) - ids.indexOf(b.id)));
            void api?.reorderAccounts(ids);
          }}
          onSettings={() => (activeAcc ? editAccount(activeAcc) : openDialog({ kind: 'chooser' }))}
          onSignatures={() => openDialog({ kind: 'signatures' })}
          gpuOff={gpuOff}
          notifications={settings.notifications !== false}
          onCrypto={() => openDialog({ kind: 'crypto' })}
          onAbout={() => openDialog({ kind: 'about' })}
          kindOf={(a) => (a.id === DEMO_ID ? '' : isM365(a) ? '365' : zimbraOk[a.id] ? 'Z' : 'IMAP')}
          onTestNotification={() => void api?.testNotification().then((r) => setStatus(r.ok ? r.data : r.error))}
          onZimbra={activeImap && zimbraOk[activeImap.id] ? () => openDialog({ kind: 'zimbra' }) : undefined}
          onToggleNotifications={() => saveSettings({ ...settings, notifications: settings.notifications === false })}
          onToggleGpu={() => void api?.setGpuDisabled(!gpuOff)}
          gpuForced={gpuForced}
          onToggleGpuForced={() => void api?.setGpuForced(!gpuForced)}
          onOpenLog={() => void api?.openLog()}
          onMenuOpen={setMenuOverlay}
        />
        <div className="workspace">
          <TopBar
            kind={activeM365 ? 'OUTLOOK WEB' : activeImap && activeImap.id !== DEMO_ID ? (zimbraOk[activeImap.id] ? 'ZIMBRA' : 'IMAP') : undefined}
            title={topTitle}
            subtitle={topSub}
            color={topColor}
            search={search}
            onSearch={activeM365 ? undefined : setSearch}
            rightOpen={rightOpen}
            onToggleRight={() => setRightOpen(!rightOpen)}
          />
          {!activeM365 && curImapModule === 'mail' && (
            <Ribbon
              tab={ribbonTab}
              onTab={setRibbonTab}
              hasMessage={!!full.msg}
              seen={current?.seen ?? true}
              flagged={current?.flagged ?? false}
              canMove={!!current && current.accountId !== DEMO_ID && folderList.length > 0}
              folders={folderList.filter((f) => f.path !== current?.folder)}
              onCompose={() => compose('new')}
              onReply={() => compose('reply')}
              onReplyAll={() => compose('replyAll')}
              onForward={() => compose('forward')}
              onDelete={() => void del()}
              onArchive={() => void archive()}
              onMove={(d) => void move(d)}
              onToggleSeen={() => void toggleSeen()}
              onToggleFlag={() => void toggleFlag()}
              categories={categoriesOf(settings)}
              keywords={current?.keywords ?? []}
              onCategory={(id, on) => void setCategory(id, on)}
              pinned={!!current?.keywords?.includes(PIN_KEYWORD)}
              onTogglePin={() => void togglePin()}
              onTask={(due) => void taskFromMessage(due)}
              onIo={(k) => void runIo(k)}
              onZimbra={activeImap && zimbraOk[activeImap.id] ? (tab) => openDialog({ kind: 'zimbra', tab }) : undefined}
              onSnooze={(d) => void snoozeCurrent(d)}
              quickSteps={settings.quickSteps ?? []}
              onQuickStep={(q) => void runQuickStep(q)}
              onOrganize={(tab) => openDialog({ kind: 'organize', tab })}
              conversations={!!settings.conversations}
              onConversations={(v) => saveSettings({ ...settings, conversations: v })}
              density={settings.density ?? 'comfortable'}
              onDensity={(d) => saveSettings({ ...settings, density: d })}
              previewLine={!!settings.previewLine}
              onPreviewLine={(v) => {
                saveSettings({ ...settings, previewLine: v });
                setRefreshKey((x) => x + 1);
              }}
              runInBackground={settings.runInBackground !== false}
              onRunInBackground={(v) => saveSettings({ ...settings, runInBackground: v })}
              onRefresh={() => setRefreshKey((k) => k + 1)}
              onEditAccount={activeImap && activeImap.id !== DEMO_ID ? () => editAccount(activeImap) : undefined}
              onSignatures={() => openDialog({ kind: 'signatures' })}
              theme={theme}
              onTheme={setTheme}
              customThemes={settings.customThemes ?? []}
              onEditThemes={() => openDialog({ kind: 'themes' })}
              rightOpen={rightOpen}
              onToggleRight={() => setRightOpen(!rightOpen)}
              compact={compact}
              onCompact={setCompact}
              gpuOff={gpuOff}
              onToggleGpu={() => void api?.setGpuDisabled(!gpuOff)}
              onOpenLog={() => void api?.openLog()}
              composer={
                composeView
                  ? {
                      state: composerState,
                      signatures: settings.signatures,
                      templates: settings.templates ?? [],
                      onManageTemplates: () => openDialog({ kind: 'templates' }),
                      onManageCrypto: () => openDialog({ kind: 'crypto' }),
                      run: (fn) => composerRef.current && fn(composerRef.current),
                    }
                  : undefined
              }
            />
          )}
          {demo && (
            <div className="banner">
              Zobrazujú sa ukážkové dáta. Pridaj konto tlačidlom <b>+</b> v páse kont vľavo (IMAP alebo Microsoft 365).
            </div>
          )}
          {Object.entries(quota).some(([, q]) => q && q.limit > 0 && q.used / q.limit > 0.95) && (
            <div className="banner warn">Schránka na serveri je takmer plná – nové správy nemusia byť doručené. Odstráňte staré správy alebo veľké prílohy (aj z koša).</div>
          )}
          {updateReady && (
            <div className="banner">
              Je pripravená nová verzia MailKlienta.{' '}
              <button className="link" onClick={() => void api?.updateInstall()}>Reštartovať a aktualizovať</button>{' '}
              <button className="link" onClick={() => setUpdateReady(null)}>Neskôr</button>
            </div>
          )}
          {gpuCrash > 0 && !gpuOff && (
            <div className="banner warn">
              Grafická karta mala problém s vykresľovaním (preto mohol obraz zbelieť alebo blikať).{' '}
              <button className="link" onClick={() => void api?.setGpuDisabled(true)}>
                Vypnúť hardvérovú akceleráciu a reštartovať
              </button>
            </div>
          )}
          {secretsWarning && (
            <div className="banner warn">Systémové úložisko hesiel nie je dostupné: heslá sa držia len do zatvorenia aplikácie.</div>
          )}
          <main className={`body ${rightOpen ? 'with-right' : ''}`}>
            {activeM365 ? (
              <ModuleRail modules={M365_MODULES} active={m365ActiveModule} color={topColor} onSelect={(k) => selectM365Module(activeM365, k)} />
            ) : (
              <ModuleRail
                modules={IMAP_MODULES}
                active={curImapModule}
                color={topColor}
                onSelect={(k) => activeImap && setImapModule((m) => ({ ...m, [activeImap.id]: k }))}
              />
            )}
            <div className="content">
              {accounts.filter(isM365).filter((a) => tabs[a.id]).map((a) => (
                <M365View
                  key={a.id}
                  account={a}
                  state={tabs[a.id]!}
                  color={accountColor(a.color)}
                  visible={activeM365?.id === a.id}
                  onActivate={(key) => setTabs((t) => ({ ...t, [a.id]: { ...t[a.id]!, active: key } }))}
                  onClose={(key) => closeTab(a, key)}
                  onTitle={(key, title) =>
                    setTabs((all) => {
                      const st = all[a.id];
                      if (!st) return all;
                      const tabsNext = st.tabs.map((x) => (x.key === key ? { ...x, title: tabTitle(title, x.module ? MODULE_TITLES[x.module] : x.title) } : x));
                      return { ...all, [a.id]: { ...st, tabs: tabsNext } };
                    })
                  }
                  onEditAccount={() => editAccount(a)}
                  onUnread={(n) => setUnread((u) => (u[a.id] === n ? u : { ...u, [a.id]: n }))}
                />
              ))}
              {activeM365 ? null : showMail ? (
                <div className="mail">
                  <FolderPane
                    account={activeImap}
                    accounts={imapAccounts}
                    folders={folders}
                    unread={unread}
                    selection={selection}
                    onSelect={(s) => {
                      setSelection(s);
                      setSearch('');
                      setCurrent(undefined);
                      setFull({ loading: false });
                    }}
                    onDropMessages={(dest, d) => void dropMessages(dest, d)}
                    scheduled={scheduled.filter((x) => selection.accountId === ALL || x.accountId === selection.accountId)}
                    onScheduled={(it, edit) => void cancelScheduled(it, edit)}
                    quota={activeImap ? quota[activeImap.id] : null}
                    savedSearches={settings.savedSearches}
                    activeSearch={search.trim()}
                    onSavedSearch={(x) => {
                      if (x.accountId && x.accountId !== selection.accountId) setSelection({ accountId: x.accountId, folder: 'INBOX' });
                      setSearch(x.query);
                    }}
                    onRemoveSearch={(x) => saveSettings({ ...settings, savedSearches: (settings.savedSearches ?? []).filter((y) => y.id !== x.id) })}
                  />
                  <MessageList
                    title={found ? `Hľadanie: ${found.q}` : listTitle}
                    headerExtra={
                      found ? (
                        <span className="list-extra">
                          <button className="link small" disabled={found.server === 'loading'} onClick={() => void searchOnServer()}>
                            {found.server === 'loading' ? 'Hľadám na serveri…' : found.server === 'done' ? 'Prehľadané aj na serveri' : `Nájdené: ${found.list.length} · hľadať aj na serveri`}
                          </button>
                          {selection.accountId !== ALL && (
                            <button className="link small" onClick={() => setSearchAll(!searchAll)}>{searchAll ? 'len toto konto' : 'vo všetkých kontách'}</button>
                          )}
                          {!(settings.savedSearches ?? []).some((x) => x.query === found.q) && (
                            <button className="link small" onClick={saveCurrentSearch}>uložiť hľadanie</button>
                          )}
                        </span>
                      ) : undefined
                    }
                    messages={filtered}
                    categories={categoriesOf(settings)}
                    conversations={!!settings.conversations && !found}
                    density={settings.density ?? 'comfortable'}
                    previewLine={!!settings.previewLine}
                    loading={listState.loading}
                    error={listState.error}
                    selected={current}
                    colorOf={colorOf}
                    onOpen={(m) => void (isDrafts(m) ? editDraft(m) : open(m))}
                    onDeleteKey={() => void del()}
                  />
                  {composeView ? (
                    <Composer
                      key={composeView.key}
                      ref={composerRef}
                      init={composeView.init}
                      accounts={imapAccounts.filter((a) => a.id !== DEMO_ID).length ? imapAccounts.filter((a) => a.id !== DEMO_ID) : imapAccounts}
                      settings={settings}
                      contacts={contacts}
                      zimbra={zimbraOk}
                      onState={setComposerState}
                      onQueued={(it) => it.kind === 'undo' && setUndo(it)}
                      onSaveTemplate={saveTemplate}
                      onClose={(st) => {
                        setCompose(null);
                        setRibbonTab('home');
                        if (st) setStatus(st);
                      }}
                    />
                  ) : (
                  <ReadingPane
                    message={full.msg}
                    loading={full.loading}
                    error={full.error}
                    color={colorOf(full.msg?.accountId ?? '')}
                    onReply={() => compose('reply')}
                    onReplyAll={() => compose('replyAll')}
                    onForward={() => compose('forward')}
                    thread={thread}
                    onOpenThread={(t) => void open(t)}
                    categories={categoriesOf(settings)}
                    onRemoveCategory={(id) => void setCategory(id, false)}
                    onBlockSender={
                      full.msg && zimbraOk[full.msg.accountId]
                        ? async () => {
                            const m = full.msg!;
                            const r = await api!.zimbraBlock(m.accountId, m.fromAddress);
                            setStatus(r.ok ? `${m.fromAddress} je blokovaný – ďalšia pošta od neho pôjde do spamu.` : r.error);
                          }
                        : undefined
                    }
                    onPrint={(pdf) => void runIo(pdf ? 'pdf' : 'print')}
                    onUnsubscribe={
                      full.msg?.unsubscribe && full.msg.accountId !== DEMO_ID
                        ? async () => {
                            const m = full.msg!;
                            const r = await api!.unsubscribe(m.accountId, m.unsubscribe!);
                            setStatus(r.ok ? r.data : r.error);
                          }
                        : undefined
                    }
                    onRespond={async (ps) => {
                      const m = full.msg;
                      if (!m || !api || m.accountId === DEMO_ID) return;
                      const r = await api.calRespond(m.accountId, m.folder, m.uid, ps);
                      setStatus(r.ok ? r.data : r.error);
                      if (r.ok) setAgendaTick((t) => t + 1);
                    }}
                  />
                  )}
                </div>
              ) : (
                curImapModule === 'tasks' ? (
                  <TasksView settings={settings} onSettings={saveSettings} onOpenMessage={openTaskMessage} />
                ) : activeImap && activeImap.id !== DEMO_ID && curImapModule === 'calendar' ? (
                  <CalendarView key={activeImap.id} account={activeImap} zimbra={!!zimbraOk[activeImap.id]} onStatus={setStatus} onSettings={() => editAccount(activeImap)} />
                ) : activeImap && activeImap.id !== DEMO_ID && curImapModule === 'contacts' ? (
                  <ContactsView
                    key={activeImap.id}
                    account={activeImap}
                    zimbra={!!zimbraOk[activeImap.id]}
                    onStatus={setStatus}
                    onCompose={(to) => {
                      setImapModule((m) => ({ ...m, [activeImap.id]: 'mail' }));
                      setCompose({ init: { kind: 'new', accountId: activeImap.id, to }, key: Date.now() });
                      setRibbonTab('message');
                    }}
                  />
                ) : (
                  <section className="placeholder">
                    <h2>{IMAP_MODULES.find((x) => x.key === curImapModule)?.label}</h2>
                    <p className="muted">Pridajte IMAP konto – kalendár a kontakty sa načítajú zo servera (CalDAV/CardDAV, napr. Zimbra).</p>
                  </section>
                )
              )}
            </div>
            {rightOpen && (
              <RightPanel
                settings={settings}
                onSettings={saveSettings}
                active={rightActive}
                onActive={setRightActive}
                overlay={overlay}
                onOverlay={setPanelOverlay}
                onStatus={setStatus}
                agenda={api && !demo ? agenda : undefined}
              />
            )}
          </main>
          {undo && (
            <div className="undo-bar" role="status">
              <span>Odosiela sa „{undo.subject || '(bez predmetu)'}“…</span>
              <button className="undo-btn" onClick={() => void undoSend()}>Späť</button>
              <span className="undo-timer" />
            </div>
          )}
          <footer className="statusbar">
            {status ||
              (IS_TAURI
                ? 'Tauri (WebView2 / jadro Edge) – test vykresľovania s ukážkovými dátami'
                : demo
                ? 'Ukážkový režim'
                : `${accounts.length} ${accounts.length === 1 ? 'konto' : accounts.length < 5 ? 'kontá' : 'kont'}`)}
          </footer>
        </div>
      </div>

      {dialog.kind === 'signatures' && (
        <SignaturesDialog
          settings={settings}
          accounts={accounts.filter(isImap)}
          onClose={closeDialog}
          onSave={(s) => {
            saveSettings(s);
            closeDialog();
            setStatus('Podpisy uložené.');
          }}
        />
      )}
      {dialog.kind === 'zimbra' && activeImap && activeImap.id !== DEMO_ID && (
        <ZimbraDialog
          account={activeImap}
          folders={folderList}
          settings={settings}
          tab={dialog.tab}
          onSettings={saveSettings}
          onStatus={setStatus}
          onAccountIdentities={(ids) => {
            const cur = activeImap.identities ?? [];
            const merged = [...cur, ...ids.filter((x) => !cur.some((c) => c.email.toLowerCase() === x.email.toLowerCase()))];
            const { hasPassword: _h, ...acc } = activeImap;
            void api?.saveAccount({ ...acc, identities: merged }).then(() => void loadAccounts());
          }}
          onClose={closeDialog}
        />
      )}
      {dialog.kind === 'about' && <AboutDialog onClose={closeDialog} />}
      {dialog.kind === 'shortcuts' && <ShortcutsDialog onClose={closeDialog} />}
      {dialog.kind === 'crypto' && <CryptoDialog accounts={accounts.filter(isImap)} onClose={closeDialog} />}
      {dialog.kind === 'themes' && (
        <ThemeDialog
          custom={settings.customThemes ?? []}
          current={theme}
          onClose={closeDialog}
          onSave={(list, select) => {
            saveSettings({ ...settings, customThemes: list });
            if (select) setTheme(select);
            else if (!list.some((t) => t.id === theme) && theme.startsWith('custom-')) setTheme('sage');
            closeDialog();
          }}
        />
      )}
      {dialog.kind === 'templates' && (
        <TemplatesDialog
          settings={settings}
          onClose={closeDialog}
          onSave={(s2) => {
            saveSettings(s2);
            closeDialog();
            setStatus('Šablóny uložené.');
          }}
        />
      )}
      {dialog.kind === 'organize' && (
        <OrganizeDialog
          key={dialogKey}
          settings={settings}
          accounts={accounts.filter(isImap)}
          folders={Object.fromEntries(Object.entries(folders).filter(([, v]) => Array.isArray(v))) as Record<string, FolderInfo[]>}
          defaultAccountId={activeImap?.id}
          tab={dialog.tab}
          onRunRules={
            api && selection.accountId !== ALL
              ? async () => {
                  const r = await api.runRules(selection.accountId, selection.folder);
                  setMailTick((t) => t + 1);
                  return r.ok ? `Pravidlá použité na ${r.data} správ.` : r.error;
                }
              : undefined
          }
          onClose={closeDialog}
          onSave={(s2) => {
            saveSettings(s2);
            closeDialog();
            setStatus('Nastavenia organizácie uložené.');
          }}
        />
      )}
      {dialog.kind === 'chooser' && (
        <AddAccountChooser
          onClose={closeDialog}
          onImap={() => openDialog({ kind: 'imap' })}
          onM365={() => openDialog({ kind: 'm365' })}
        />
      )}
      {dialog.kind === 'imap' && (
        <AccountDialog
          key={dialogKey}
          open
          edit={dialog.edit}
          usedColors={accounts.map((a) => accountColor(a.color))}
          onClose={closeDialog}
          onSaved={(a) => {
            closeDialog();
            setStatus(`Konto ${a.name} uložené.`);
            setActiveId(a.id);
            void loadAccounts();
            setRefreshKey((k) => k + 1);
          }}
          onDeleted={() => {
            closeDialog();
            setActiveId('');
            void loadAccounts();
          }}
        />
      )}
      {dialog.kind === 'm365' && (
        <M365Dialog
          key={dialogKey}
          edit={dialog.edit}
          usedColors={accounts.map((a) => accountColor(a.color))}
          onClose={closeDialog}
          onSaved={(a) => {
            closeDialog();
            setStatus(`Konto ${a.name} uložené.`);
            setActiveId(a.id);
            void loadAccounts();
          }}
          onDeleted={() => {
            closeDialog();
            setActiveId('');
            void loadAccounts();
          }}
        />
      )}
      </div>
    </FluentProvider>
  );
}
