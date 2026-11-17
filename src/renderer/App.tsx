// Rozloženie C: pás kont úplne vľavo → pás modulov podľa konta → obsah (IMAP klient alebo vložený Outlook) → pravý panel.
import { FluentProvider } from '@fluentui/react-components';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { isImap, isM365, type AccountPublic, type ImapAccountPublic, type M365Account } from '../core/account';
import type { FolderInfo, MailApi, MessageFull, MessageSummary } from '../core/api';
import { DEFAULT_SETTINGS, type AppSettings } from '../core/settings';
import { ALL, AccountRail } from './components/AccountRail';
import { AccountDialog } from './components/AccountDialog';
import { Composer, type ComposeInit, type ComposeKind, type ComposerHandle, type ComposerState } from './components/Composer';
import type { Contact } from './components/RecipientField';
import {
  M365View, MODULE_TITLES, baseKey, initialTabs, moduleUrl, tabTitle, type TabsState,
} from './components/M365View';
import { AddAccountChooser, M365Dialog } from './components/M365Dialog';
import { SignaturesDialog } from './components/SignaturesDialog';
import { FolderPane, MessageList, ReadingPane, type Selection } from './components/MailPanes';
import { RightPanel } from './components/RightPanel';
import {
  IMAP_MODULES, M365_MODULES, ModuleRail, Ribbon, TopBar, type ImapModule, type M365Module, type RibbonTab,
} from './components/Shell';
import { DEMO_ID, demoAccount, demoFolders, demoMessages } from './mock';
import { THEMES, accountColor, getTheme, type ThemeKey } from './theme';

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
  | { kind: 'signatures' };

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
    // kliknutia do záznamu (čas a cieľ), aby sa dalo spárovať s blikaním
    const onClick = (e: MouseEvent) => {
      const t = e.target as HTMLElement | null;
      const label = t?.closest('button,[role=tab],a')?.textContent?.trim().slice(0, 40) ?? t?.tagName ?? '';
      void api.logError(`klik: ${label}`);
    };
    window.addEventListener('click', onClick, true);
    const onVis = () => void api.logError(`viditeľnosť stránky: ${document.visibilityState}`);
    document.addEventListener('visibilitychange', onVis);
    const onErr = (e: ErrorEvent) => void api.logError(`${e.message} @ ${e.filename}:${e.lineno}`);
    const onRej = (e: PromiseRejectionEvent) => void api.logError(`promise: ${String(e.reason)}`);
    window.addEventListener('error', onErr);
    window.addEventListener('unhandledrejection', onRej);
    const off = api.onAppEvent((ev) => ev.type === 'gpu-crash' && setGpuCrash(ev.count));
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
      window.removeEventListener('click', onClick, true);
      document.removeEventListener('visibilitychange', onVis);
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
    const t = window.setInterval(run, 120_000);
    return () => {
      alive = false;
      window.clearInterval(t);
    };
  }, [api, accounts, demo, refreshKey]);

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
      setListState({ loading: true });
      const results = await Promise.all(targets.map((t) => api.listMessages(t.id, t.folder, 50)));
      if (cancelled) return;
      const errors = results.flatMap((r, i) => (r.ok ? [] : [`${accounts.find((a) => a.id === targets[i]!.id)?.name}: ${r.error}`]));
      const all = results.flatMap((r) => (r.ok ? r.data : []));
      all.sort((a, b) => (b.date ?? '').localeCompare(a.date ?? ''));
      setMessages(all);
      setListState({ loading: false, error: errors.length ? errors.join(' · ') : undefined });
    })();
    return () => {
      cancelled = true;
    };
  }, [selection, accounts, demo, api, refreshKey, showMail]);

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
      setFull({ loading: false, msg: cached });
      if (m.seen) return;
    } else setFull((f) => ({ loading: true, msg: f.msg })); // predchádzajúca správa ostáva viditeľná
    const r = cached ? ({ ok: true, data: cached } as const) : await api!.getMessage(m.accountId, m.folder, m.uid);
    if (r.ok) {
      msgCache.current.set(ck, r.data);
      if (msgCache.current.size > 200) msgCache.current.delete(msgCache.current.keys().next().value!);
    }
    if (openKey.current !== ck) return; // medzitým klikol na inú správu
    setFull(r.ok ? { loading: false, msg: r.data } : { loading: false, error: `${r.error}${'hint' in r && r.hint ? ` ${r.hint}` : ''}` });
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
  const contacts = useMemo<Contact[]>(() => {
    let known: Contact[] = [];
    try {
      known = JSON.parse(localStorage.getItem('contacts') ?? '[]') as Contact[];
    } catch {
      /* nič */
    }
    const all = [...known, ...messages.map((m) => ({ name: m.fromName, email: m.fromAddress }))].filter((c) => c.email);
    return all.filter((c, i) => all.findIndex((x) => x.email.toLowerCase() === c.email.toLowerCase()) === i);
  }, [messages, composeView]);

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return messages;
    return messages.filter((m) => `${m.subject} ${m.fromName} ${m.fromAddress}`.toLowerCase().includes(q));
  }, [messages, search]);

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
  const accent = THEMES[theme].accent;
  const topColor = activeAcc ? accountColor(activeAcc.color) : accent;
  const topTitle = activeAcc ? activeAcc.name : 'Všetky doručené';
  const topSub = activeM365
    ? `${activeM365.email || 'Microsoft 365'} · Outlook na webe`
    : activeImap
      ? `${activeImap.email} · IMAP${activeImap.imap.host ? ` · ${activeImap.imap.host}` : ''}`
      : `${imapAccounts.length} IMAP ${imapAccounts.length === 1 ? 'konto' : imapAccounts.length < 5 ? 'kontá' : 'kont'}`;

  const curImapModule: ImapModule = activeImap ? imapModule[activeImap.id] ?? 'mail' : 'mail';
  const m365State = activeM365 ? tabs[activeM365.id] ?? initialTabs(activeM365) : null;
  const m365ActiveModule = m365State ? m365State.tabs.find((t) => t.key === m365State.active)?.module ?? null : null;

  return (
    <FluentProvider theme={getTheme(theme)} className={`app theme-${theme}`}>
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
          onToggleGpu={() => void api?.setGpuDisabled(!gpuOff)}
          gpuForced={gpuForced}
          onToggleGpuForced={() => void api?.setGpuForced(!gpuForced)}
          onOpenLog={() => void api?.openLog()}
              onBlinkTest={(k) => void api?.blinkTest(k)}
          onMenuOpen={setMenuOverlay}
        />
        <div className="workspace">
          <TopBar
            title={topTitle}
            subtitle={topSub}
            color={topColor}
            search={search}
            onSearch={activeM365 ? undefined : setSearch}
            rightOpen={rightOpen}
            onToggleRight={() => setRightOpen(!rightOpen)}
          />
          {!activeM365 && (
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
              onRefresh={() => setRefreshKey((k) => k + 1)}
              onEditAccount={activeImap && activeImap.id !== DEMO_ID ? () => editAccount(activeImap) : undefined}
              onSignatures={() => openDialog({ kind: 'signatures' })}
              theme={theme}
              onTheme={setTheme}
              rightOpen={rightOpen}
              onToggleRight={() => setRightOpen(!rightOpen)}
              compact={compact}
              onCompact={setCompact}
              gpuOff={gpuOff}
              onToggleGpu={() => void api?.setGpuDisabled(!gpuOff)}
              onOpenLog={() => void api?.openLog()}
              composer={
                composeView
                  ? { state: composerState, signatures: settings.signatures, run: (fn) => composerRef.current && fn(composerRef.current) }
                  : undefined
              }
            />
          )}
          {demo && (
            <div className="banner">
              Zobrazujú sa ukážkové dáta. Pridaj konto tlačidlom <b>+</b> v páse kont vľavo (IMAP alebo Microsoft 365).
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
                      setCurrent(undefined);
                      setFull({ loading: false });
                    }}
                  />
                  <MessageList
                    title={listTitle}
                    messages={filtered}
                    loading={listState.loading}
                    error={listState.error}
                    selected={current}
                    colorOf={colorOf}
                    onOpen={(m) => void open(m)}
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
                      onState={setComposerState}
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
                  />
                  )}
                </div>
              ) : (
                <section className="placeholder">
                  <h2>{IMAP_MODULES.find((x) => x.key === curImapModule)?.label}</h2>
                  <p className="muted">
                    Pripravuje sa: kalendár, úlohy a kontakty cez CalDAV/CardDAV (Zimbra, Google, Nextcloud) pre konto {activeImap?.name}.
                  </p>
                </section>
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
              />
            )}
          </main>
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
    </FluentProvider>
  );
}
