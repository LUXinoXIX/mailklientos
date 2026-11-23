// Rozloženie C: pás kont úplne vľavo → pás modulov podľa konta → obsah (IMAP klient alebo vložený Outlook) → pravý panel.
import { Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, FluentProvider, Input, Tooltip } from '@fluentui/react-components';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { isGmail, isImap, isM365, type AccountPublic, type ImapAccountPublic, type M365Account } from '../core/account';
import type { FolderInfo, MailApi, MessageFull, MessageSummary, OutboxItem } from '../core/api';
import { DEFAULT_SETTINGS, PIN_KEYWORD, categoriesOf, SEARCH_FIELDS, type SearchField, type AppSettings, type QuickStep, type SavedSearch, type Task } from '../core/settings';
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
import { EditSmallIcon, MailIcon } from './icons';
import { MessageView, type MessageRef } from './components/MessageView';
import { Splitter } from './components/Splitter';
import { ContextMenu, type CtxItem, type CtxState } from './components/ContextMenu';
import { ZimbraDialog, type ZimbraTab } from './components/ZimbraDialog';
import { OrganizeDialog, type OrganizeTab } from './components/OrganizeDialog';
import { FolderPane, MessageList, ReadingPane, type Selection } from './components/MailPanes';
import { RightPanel, type AgendaData, type PanelChoice } from './components/RightPanel';
import { MyDayPane, type MyDayEvent } from './components/MyDayPane';
import { useAttachmentGuard } from './components/AttachmentGuard';
import { attachmentRisk } from '../core/attrisk';
import { CalendarView } from './components/CalendarView';
import { AddressBook } from './components/AddressBook';
import { SettingsTransfer } from './components/SettingsTransfer';
import { LockSettings } from './components/LockSettings';
import { RedirectDialog } from './components/RedirectDialog';
import { SUMMARY, SummaryPage } from './components/SummaryPage';
import { TasksView, dueLabel, todayIso } from './components/TasksView';
import {
  IMAP_MODULES, M365_PERSONAL_MODULES, M365_WORK_MODULES, WEBMAIL_MODULES, AccountRailFooter, ModuleRail, RailEditor, Ribbon, TopBar, customizeRail, type ImapModule, type M365Module, type ModuleDef, type RibbonTab,
} from './components/Shell';
import { DEMO_ID, demoAccount, demoFolders, demoMessages } from './mock';

/** kľúč modulu pri voľbe „Všetky“ kontá (Pošta / Úlohy / Kontakty) */
const ALL_MODULES = '__all__';
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
  | { kind: 'imap'; edit?: ImapAccountPublic; preset?: 'gmail' }
  | { kind: 'm365'; edit?: M365Account; host?: 'office' | 'live' | 'web' }
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
const POP_COMPOSE =
  POP.get('compose') === '1'
    ? { acc: POP.get('acc') ?? '', folder: POP.get('folder') ?? '', uid: Number(POP.get('uid')), mode: (POP.get('mode') ?? '') as '' | 'reply' | 'replyAll' | 'forward' | 'forwardAttach' }
    : null;
// samostatné okno so správou
const POP_VIEW = POP.get('view') === '1' ? { accountId: POP.get('acc') ?? '', folder: POP.get('folder') ?? '', uid: Number(POP.get('uid')) } : null;
/** počet správ načítaných naraz do zoznamu (ďalšie po posunutí na koniec) */
const PAGE = 200;

/** Odpoveď / preposlanie: adresáti a predmet ako v Outlooku. */
function replyInit(m: MessageFull, kind: 'reply' | 'replyAll' | 'forward', accountId: string | undefined, myEmail: string): ComposeInit {
  const me = myEmail.toLowerCase();
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
  return {
    kind,
    accountId,
    original: m,
    to: kind === 'forward' ? '' : m.fromName ? `"${m.fromName.replace(/"/g, '')}" <${m.fromAddress}>` : m.fromAddress,
    cc,
    subject,
  };
}

export function App() {
  const api = window.mailApi;
  const [theme, setTheme] = usePersisted<ThemeKey>('theme', 'blue');
  const [compact, setCompact] = usePersisted('ribbonCompact', false);
  const [rightOpen, setRightOpen] = usePersisted('rightOpen', SMOKE);
  const [rightActive, setRightActive] = usePersisted('rightActive', 'copilot');
  const [activeId, setActiveId] = usePersisted<string>('activeAccount', '');
  const [ribbonTab, setRibbonTab] = useState<RibbonTab>('home');
  const { guard: attGuard, dialog: attGuardDialog } = useAttachmentGuard();
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
  // šírky stĺpcov (ťahadlá medzi stĺpcami), uložené v nastaveniach
  const PW = { folders: 230, list: 340, right: 360, myday: 290 };
  const [widths, setWidths] = useState(PW);
  useEffect(() => setWidths({ ...PW, ...(settings.paneWidths ?? {}) }), [settings.paneWidths]); // eslint-disable-line react-hooks/exhaustive-deps
  const saveWidth = (k: keyof typeof PW, v: number) => saveSettings({ ...settings, paneWidths: { ...(settings.paneWidths ?? {}), [k]: v } });
  const [msgTabs, setMsgTabs] = useState<Array<{ key: string; ref: MessageRef; title: string }>>([]);
  const [activeTab, setActiveTab] = useState<string | null>(null);
  const [limits, setLimits] = useState<Record<string, number>>({});
  const [hasMore, setHasMore] = useState(false);
  const [followupEv, setFollowupEv] = useState<{ subject: string; to: string; accountId: string } | null>(null);
  const [searchAll, setSearchAll] = useState(false);
  const [outboxTick, setOutboxTick] = useState(0);
  const [scheduled, setScheduled] = useState<OutboxItem[]>([]);
  const [status, setStatus] = useState('');
  const [transfer, setTransfer] = useState<null | 'export' | 'import'>(null);
  const [lockDlg, setLockDlg] = useState(false);
  const [redirect, setRedirect] = useState<MessageSummary | null>(null);
  const [dialog, setDialog] = useState<Dialogs>({ kind: 'none' });
  const [dialogKey, setDialogKey] = useState(0);
  const [panelOverlay, setPanelOverlay] = useState(false);
  const [menuOverlay, setMenuOverlay] = useState(false);
  const msgCache = useRef(new Map<string, MessageFull>());
  // rozpísané správy ako karty dole (ako Outlook): viac naraz, prepínanie medzi nimi a správami
  const [drafts, setDrafts] = useState<Array<{ init: ComposeInit; key: number }>>([]);
  const [activeDraft, setActiveDraft] = useState<number | null>(null);
  const composeView = drafts.find((d) => d.key === activeDraft) ?? null;
  const setCompose = (v: { init: ComposeInit; key: number } | null) => {
    if (!v) return setActiveDraft(null);
    setDrafts((d) => [...d, v]);
    setActiveDraft(v.key);
  };
  const removeDraft = (key: number) => {
    setDrafts((d) => d.filter((x) => x.key !== key));
    setActiveDraft((a) => (a === key ? null : a));
  };
  const composerRefs = useRef<Record<number, ComposerHandle | null>>({});
  const composerRef = { current: activeDraft != null ? composerRefs.current[activeDraft] ?? null : null };
  const [composerStates, setComposerStates] = useState<Record<number, ComposerState>>({});
  const composerState: ComposerState = (activeDraft != null && composerStates[activeDraft]) || { importance: false, receipt: false, savedAt: null, sending: false };
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
      else if (ev.type === 'settings-changed') void api.getSettings().then((r) => r.ok && setSettings(r.data));
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
        else if (ev.state === 'downloading') setStatus('Sťahujem novú verziu programu Universum eMail na pozadí…');
      } else if (ev.type === 'followup') {
        setFollowupEv(ev);
      } else if (ev.type === 'progress') {
        setStatus(ev.text);
      } else if (ev.type === 'outbox-changed') {
        setOutboxTick((t) => t + 1);
      } else if (ev.type === 'outbox-queued') {
        // odoslané zo samostatného okna (to sa hneď zavrie) → „Odosiela sa… Späť“ v hlavnom okne
        if (POP_COMPOSE || POP_VIEW) return;
        if (ev.item.kind === 'undo') setUndo((u) => (u?.id === ev.item.id ? u : ev.item));
        else setStatus(`Správa „${ev.item.subject || '(bez predmetu)'}“ sa odošle ${new Date(ev.item.sendAt).toLocaleString('sk-SK', { weekday: 'short', day: 'numeric', month: 'numeric', hour: '2-digit', minute: '2-digit' })}.`);
      } else if (ev.type === 'contacts-changed') {
        setRecentTick((t) => t + 1);
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
    void api?.getSettings().then((r) => {
      if (!r.ok) return;
      setSettings(r.data);
      // úvodná stránka po spustení (len hlavné okno)
      if (!POP_COMPOSE && !POP_VIEW && !startupDone.current) {
        startupDone.current = true;
        const st = r.data.summary?.startup ?? 'summary';
        if (st === 'summary' && r.data.summary?.enabled !== false) setActiveId(SUMMARY);
        else if (st === 'all') setActiveId(ALL);
        else if (activeId === SUMMARY) setActiveId(localStorage.getItem('lastAccount') ?? '');
      }
    });
  }, [loadAccounts, api]); // eslint-disable-line react-hooks/exhaustive-deps
  const startupDone = useRef(false);
  const summaryEnabled = settings.summary?.enabled !== false;
  const toggleAllInboxes = () => {
    const on = settings.showAllInboxes === false;
    saveSettings({ ...settings, showAllInboxes: on });
    // pri skrytí, ak bolo zvolené „Všetky“, prepnúť na prvé konto
    if (!on && activeId === ALL) setActiveId(imapAccounts[0]?.id ?? activeId);
  };
  const toggleSummary = () => {
    const on = !summaryEnabled;
    saveSettings({ ...settings, summary: { ...(settings.summary ?? {}), enabled: on } });
    if (on) setActiveId(SUMMARY);
    else if (activeId === SUMMARY) setActiveId(localStorage.getItem('lastAccount') ?? '');
    setStatus(on ? 'Súhrn je zapnutý.' : 'Súhrn je vypnutý. Zapneš ho v karte Zobrazenie alebo v ozubenom kolieske.');
  };
  const summaryOn = summaryEnabled && activeId === SUMMARY && !POP_COMPOSE && !POP_VIEW;
  useEffect(() => {
    if (activeId && activeId !== SUMMARY) localStorage.setItem('lastAccount', activeId);
  }, [activeId]);

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
  const showMail = !activeM365 && (imapModule[activeImap?.id ?? ALL_MODULES] ?? 'mail') === 'mail';
  useEffect(() => {
    if (!showMail || POP_VIEW || POP_COMPOSE) return;
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
      const limit = limits[`${selection.accountId}/${selection.folder}`] ?? PAGE;
      const cached = await Promise.all(targets.map((t) => api.listCached(t.id, t.folder, limit)));
      if (cancelled) return;
      const fromCache = cached.flatMap((r) => (r.ok ? r.data : [])).sort(byDate);
      if (fromCache.length && !mailTick) setMessages(fromCache);
      if (!mailTick || !messages.length) setListState({ loading: true });
      // 2) aktuálny stav zo servera (uloží sa aj do cache)
      const results = await Promise.all(targets.map((t) => api.listMessages(t.id, t.folder, limit, !!settings.previewLine)));
      if (cancelled) return;
      setHasMore(results.some((r) => r.ok && r.data.length >= limit));
      const errors = results.flatMap((r, i) => (r.ok ? [] : [`${accounts.find((a) => a.id === targets[i]!.id)?.name}: ${r.error}`]));
      const all = results.flatMap((r, i) => (r.ok ? r.data : cached[i]?.ok ? cached[i].data : [])).sort(byDate);
      setMessages(all);
      const offline = errors.length && all.length;
      setListState({ loading: false, error: errors.length ? (offline ? `Offline – zobrazené uložené správy. ${errors.join(' · ')}` : errors.join(' · ')) : undefined });
    })();
    return () => {
      cancelled = true;
    };
  }, [selection, accounts, demo, api, refreshKey, showMail, mailTick, limits]); // eslint-disable-line react-hooks/exhaustive-deps

  /** Odoslať a prijať: všetky kontá, neodoslané správy znova. */
  async function sendReceive() {
    if (!api) return;
    setStatus('Odosielam a prijímam…');
    const r = await api.sendReceive();
    setRefreshKey((x) => x + 1);
    setMailTick((t) => t + 1);
    setStatus(r.ok ? (r.data.retried ? `Prijaté; znova odosielam ${r.data.retried} neodoslan${r.data.retried === 1 ? 'ú správu' : 'é správy'}.` : 'Pošta skontrolovaná na všetkých kontách.') : r.error);
  }

  function loadMore() {
    const k = `${selection.accountId}/${selection.folder}`;
    setLimits((l) => ({ ...l, [k]: (l[k] ?? PAGE) + 500 }));
    setMailTick((t) => t + 1);
  }

  /** Dvojklik: samostatné okno (predvolené) alebo karta; Shift = opačne. */
  function openFull(m: MessageSummary, alt = false) {
    if (!api || m.accountId === DEMO_ID) return;
    if (isDrafts(m)) return void editDraft(m);
    const mode = settings.openMode ?? 'window';
    const asWindow = alt ? mode !== 'window' : mode === 'window';
    if (asWindow) return void api.messageWindow(m.accountId, m.folder, m.uid);
    const key = `${m.accountId}/${m.folder}/${m.uid}`;
    setMsgTabs((t) => (t.some((x) => x.key === key) ? t : [...t, { key, ref: { accountId: m.accountId, folder: m.folder, uid: m.uid }, title: m.subject }]));
    setActiveTab(key);
  }
  function closeMsgTab(key: string) {
    setMsgTabs((t) => {
      const rest = t.filter((x) => x.key !== key);
      if (activeTab === key) {
        const i = t.findIndex((x) => x.key === key);
        setActiveTab(rest[i]?.key ?? rest[i - 1]?.key ?? null);
      }
      return rest;
    });
  }

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

  // prepnutie priečinka, konta alebo modulu počas písania: rozpísaná správa ostane ako karta dole
  // a horný pás sa vráti na Domov (Odoslať a prijať, Nová správa…)
  const navKey = `${selection.accountId}|${selection.folder}|${activeKey}|${imapModule[activeKey] ?? 'mail'}`;
  const prevNav = useRef(navKey);
  useEffect(() => {
    if (prevNav.current === navKey) return;
    prevNav.current = navKey;
    if (composeView) setCompose(null);
    setRibbonTab((t) => (t === 'message' || t === 'insert' ? 'home' : t));
  }, [navKey]); // eslint-disable-line react-hooks/exhaustive-deps

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
  // viac vybraných správ v zozname (Ctrl/Shift + klik alebo Shift+šípky)
  const [multiSel, setMultiSel] = useState<MessageSummary[]>([]);
  const removeFromList = (m: MessageSummary) => {
    const same = (x?: MessageSummary) => !!x && x.uid === m.uid && x.accountId === m.accountId && x.folder === m.folder;
    setMessages((list) => list.filter((x) => !same(x)));
    // výber zrušiť len vtedy, keď ukazuje odstránenú správu (po Delete v zozname sa medzitým vybrala ďalšia)
    setCurrent((c) => (same(c) ? undefined : c));
    setFull((f) => (f.msg && same(f.msg) ? { loading: false } : f));
  };
  // ---- menu na pravý klik a hromadné akcie ----
  const [ctx, setCtx] = useState<CtxState | null>(null);
  const [ask, setAsk] = useState<{ title: string; value: string; ok: (v: string) => void } | null>(null);
  const showCtx = (x: number, y: number, items: CtxItem[]) => {
    setMenuOverlay(true);
    setCtx({ x, y, items });
  };
  const closeCtx = () => {
    setCtx(null);
    setMenuOverlay(false);
  };
  type BulkOp = 'seen' | 'unseen' | 'flag' | 'unflag' | 'pin' | 'unpin' | 'delete' | 'archive' | { move: string } | { cat: string; on: boolean };
  async function bulk(msgs: MessageSummary[], op: BulkOp) {
    if (!api || !msgs.length) return;
    if (msgs.some((m) => m.accountId === DEMO_ID)) return guardDemo(msgs[0]!);
    const groups = new Map<string, MessageSummary[]>();
    for (const m of msgs) groups.set(`${m.accountId}\n${m.folder}`, [...(groups.get(`${m.accountId}\n${m.folder}`) ?? []), m]);
    let err = '';
    for (const list of groups.values()) {
      const { accountId: a, folder: f } = list[0]!;
      const uids = list.map((m) => m.uid);
      let r: { ok: boolean; error?: string } = { ok: true };
      if (op === 'seen' || op === 'unseen') r = await api.setFlags(a, f, uids, { seen: op === 'seen' });
      else if (op === 'flag' || op === 'unflag') r = await api.setFlags(a, f, uids, { flagged: op === 'flag' });
      else if (op === 'pin' || op === 'unpin') r = await api.setKeywords(a, f, uids, op === 'pin' ? [PIN_KEYWORD] : [], op === 'pin' ? [] : [PIN_KEYWORD]);
      else if (op === 'delete') r = await api.deleteMessages(a, f, uids);
      else if (op === 'archive') r = await api.archiveMessages(a, f, uids);
      else if ('move' in op) r = await api.moveMessages(a, f, uids, op.move);
      else r = await api.setKeywords(a, f, uids, op.on ? [op.cat] : [], op.on ? [] : [op.cat]);
      if (!r.ok) {
        err = r.error ?? 'Chyba';
        continue;
      }
      for (const m of list) {
        if (op === 'seen' || op === 'unseen') patchMsg(m, { seen: op === 'seen' });
        else if (op === 'flag' || op === 'unflag') patchMsg(m, { flagged: op === 'flag' });
        else if (op === 'pin' || op === 'unpin' || (typeof op === 'object' && 'cat' in op)) {
          const id = typeof op === 'object' && 'cat' in op ? op.cat : PIN_KEYWORD;
          const on = typeof op === 'object' && 'cat' in op ? op.on : op === 'pin';
          const kw = new Set(m.keywords ?? []);
          if (on) kw.add(id);
          else kw.delete(id);
          patchMsg(m, { keywords: [...kw] });
        } else removeFromList(m);
      }
      if (op === 'seen' || op === 'unseen') setMailTick((t) => t + 1);
    }
    const n = msgs.length;
    const word = n === 1 ? 'správa' : n < 5 ? 'správy' : 'správ';
    setStatus(err || (op === 'delete' ? `Odstránené: ${n} ${word}.` : op === 'archive' ? `Archivované: ${n} ${word}.` : typeof op === 'object' && 'move' in op ? `Presunuté: ${n} ${word}.` : ''));
  }

  function messageMenu(msgs: MessageSummary[], x: number, y: number) {
    const one = msgs.length === 1 ? msgs[0]! : null;
    const accs = new Set(msgs.map((m) => m.accountId));
    const fs = accs.size === 1 ? folders[[...accs][0]!] : undefined;
    const fl = Array.isArray(fs) ? fs : [];
    const anyUnread = msgs.some((m) => !m.seen);
    const anyFlag = msgs.some((m) => m.flagged);
    const allPinned = msgs.every((m) => m.keywords?.includes(PIN_KEYWORD));
    const later = (h: number, d = 0) => {
      const t = new Date();
      t.setDate(t.getDate() + d);
      t.setHours(h, 0, 0, 0);
      return t;
    };
    const snoozeOne = async (until: Date) => {
      if (!one || !api || guardDemo(one)) return;
      const r = await api.snooze(one.accountId, one.folder, one.uid, one.messageId ?? '', one.subject, until.toISOString());
      if (!r.ok) return setStatus(r.error);
      removeFromList(one);
      setStatus(`Správa odložená do ${until.toLocaleString('sk-SK', { weekday: 'short', day: 'numeric', month: 'numeric', hour: '2-digit', minute: '2-digit' })}.`);
    };
    const items: CtxItem[] = [];
    if (one) {
      items.push(
        { label: 'Otvoriť v novom okne', onClick: () => openFull(one, (settings.openMode ?? 'window') !== 'window') },
        { label: 'Otvoriť na karte', onClick: () => openFull(one, (settings.openMode ?? 'window') === 'window') },
        { divider: true },
        { label: 'Odpovedať', hint: 'Ctrl+R', onClick: () => void api?.replyWindow(one.accountId, one.folder, one.uid, 'reply') },
        { label: 'Odpovedať všetkým', hint: 'Ctrl+Shift+R', onClick: () => void api?.replyWindow(one.accountId, one.folder, one.uid, 'replyAll') },
        {
          label: 'Poslať ďalej',
          sub: [
            { label: 'Poslať ďalej', hint: 'Ctrl+F', onClick: () => void api?.replyWindow(one.accountId, one.folder, one.uid, 'forward') },
            { label: 'Poslať ďalej ako prílohu…', onClick: () => void api?.replyWindow(one.accountId, one.folder, one.uid, 'forwardAttach') },
            { label: 'Poslať ďalej ako originál…', onClick: () => setRedirect(one) },
          ],
        },
        { divider: true },
      );
    } else items.push({ label: `Označených správ: ${msgs.length}`, disabled: true }, { divider: true });
    items.push(
      { label: anyUnread ? 'Označiť ako prečítané' : 'Označiť ako neprečítané', hint: anyUnread ? 'Ctrl+Q' : 'Ctrl+U', onClick: () => void bulk(msgs, anyUnread ? 'seen' : 'unseen') },
      { label: anyFlag ? 'Zrušiť vlajku' : 'Vlajka', hint: 'Insert', onClick: () => void bulk(msgs, anyFlag ? 'unflag' : 'flag') },
      { label: allPinned ? 'Odopnúť' : 'Pripnúť navrch', onClick: () => void bulk(msgs, allPinned ? 'unpin' : 'pin') },
      {
        label: 'Kategória',
        sub: categoriesOf(settings).map((c) => {
          const on = msgs.every((m) => m.keywords?.includes(c.id));
          return { label: `${on ? '✓ ' : ''}${c.name}`, onClick: () => void bulk(msgs, { cat: c.id, on: !on }) };
        }),
      },
    );
    if (one) {
      items.push({
        label: 'Odložiť',
        sub: [
          { label: 'Dnes o 17:00', onClick: () => void snoozeOne(later(17)), disabled: new Date().getHours() >= 17 },
          { label: 'Zajtra o 8:00', onClick: () => void snoozeOne(later(8, 1)) },
          { label: 'O týždeň', onClick: () => void snoozeOne(later(8, 7)) },
        ],
      });
    }
    items.push(
      { divider: true },
      {
        label: 'Presunúť do',
        disabled: accs.size !== 1,
        sub: fl.filter((f) => !msgs.every((m) => m.folder === f.path)).map((f) => ({ label: `${'\u2003'.repeat(f.depth)}${f.name}`, onClick: () => void bulk(msgs, { move: f.path }) })),
      },
      { label: 'Archivovať', onClick: () => void bulk(msgs, 'archive') },
      { label: 'Odstrániť', hint: 'Delete', onClick: () => void bulk(msgs, 'delete') },
    );
    if (one) {
      items.push(
        { divider: true },
        { label: 'Tlačiť…', hint: 'Ctrl+P', onClick: () => void api?.printMessage(one.accountId, one.folder, one.uid, false) },
        { label: 'Uložiť ako PDF…', onClick: () => void api?.printMessage(one.accountId, one.folder, one.uid, true) },
        { label: 'Uložiť ako .eml…', onClick: () => void api?.exportEml(one.accountId, one.folder, one.uid, one.subject).then((r) => setStatus(r.ok ? (r.data ? `Uložené: ${r.data}` : '') : r.error)) },
      );
      if (zimbraOk[one.accountId]) {
        items.push({
          label: `Blokovať ${one.fromAddress}`,
          onClick: () => {
            if (!window.confirm(`Blokovať odosielateľa ${one.fromAddress}? Ďalšie správy od neho Zimbra presunie do nevyžiadanej pošty.`)) return;
            void api?.zimbraBlock(one.accountId, one.fromAddress).then((r) => setStatus(r.ok ? `Odosielateľ ${one.fromAddress} blokovaný.` : r.error));
          },
        });
      }
    }
    showCtx(x, y, items);
  }

  function folderMenu(accountId: string, f: FolderInfo, x: number, y: number) {
    if (accountId === DEMO_ID || !api) return;
    const special = !!f.specialUse;
    const trashLike = f.specialUse === '\\Trash' || f.specialUse === '\\Junk' || /trash|junk|spam|kôš|kos|nevy/i.test(f.path);
    const done = (msg: string) => {
      setStatus(msg);
      setRefreshKey((k) => k + 1);
    };
    showCtx(x, y, [
      {
        label: 'Nový podpriečinok…',
        onClick: () =>
          setAsk({
            title: `Nový priečinok v „${f.name}“`,
            value: '',
            ok: (v) => void api.folderOp(accountId, 'create', f.path, v).then((r) => done(r.ok ? `Priečinok „${v}“ vytvorený.` : r.error)),
          }),
      },
      {
        label: 'Premenovať…',
        disabled: special,
        onClick: () =>
          setAsk({
            title: `Premenovať priečinok „${f.name}“`,
            value: f.name,
            ok: (v) =>
              void api.folderOp(accountId, 'rename', f.path, v).then((r) => {
                done(r.ok ? `Priečinok premenovaný na „${v}“.` : r.error);
                if (r.ok && selection.accountId === accountId && selection.folder === f.path) setSelection({ accountId, folder: r.data });
              }),
          }),
      },
      { divider: true },
      { label: 'Označiť všetko ako prečítané', onClick: () => void api.folderOp(accountId, 'markRead', f.path).then((r) => (r.ok ? (setMailTick((t) => t + 1), done('Všetky správy v priečinku sú prečítané.')) : setStatus(r.error))) },
      {
        label: 'Vyprázdniť priečinok…',
        disabled: !trashLike,
        onClick: () => {
          if (!window.confirm(`Natrvalo odstrániť všetky správy v priečinku „${f.name}“? Toto sa nedá vrátiť.`)) return;
          void api.folderOp(accountId, 'empty', f.path).then((r) => (r.ok ? (setMailTick((t) => t + 1), done(`Priečinok „${f.name}“ vyprázdnený.`)) : setStatus(r.error)));
        },
      },
      {
        label: 'Odstrániť priečinok…',
        disabled: special || f.path.toUpperCase() === 'INBOX',
        onClick: () => {
          if (!window.confirm(`Odstrániť priečinok „${f.name}“ aj so všetkými správami na serveri?`)) return;
          void api.folderOp(accountId, 'delete', f.path).then((r) => {
            done(r.ok ? `Priečinok „${f.name}“ odstránený.` : r.error);
            if (r.ok && selection.folder === f.path) setSelection({ accountId, folder: 'INBOX' });
          });
        },
      },
      { divider: true },
      { label: 'Obnoviť', onClick: () => setRefreshKey((k) => k + 1) },
    ]);
  }

  const guardDemo = (m: MessageSummary) => {
    if (m.accountId === DEMO_ID) {
      setStatus('V ukážkovom režime sa správy nemenia. Pridaj konto (+ v páse kont).');
      return true;
    }
    return false;
  };
  async function del() {
    if (multiSel.length > 1) return bulk(multiSel, 'delete');
    const m = current;
    if (!m || !api || guardDemo(m)) return;
    const r = await api.deleteMessages(m.accountId, m.folder, [m.uid]);
    if (!r.ok) return setStatus(r.error);
    removeFromList(m);
    setStatus(r.data.movedTo ? 'Správa presunutá do Odstránených.' : 'Správa natrvalo odstránená.');
  }
  async function archive() {
    if (multiSel.length > 1) return bulk(multiSel, 'archive');
    const m = current;
    if (!m || !api || guardDemo(m)) return;
    const r = await api.archiveMessages(m.accountId, m.folder, [m.uid]);
    if (!r.ok) return setStatus(r.error);
    removeFromList(m);
    setStatus(`Správa archivovaná (${r.data.movedTo}).`);
  }
  async function move(dest: string) {
    if (multiSel.length > 1) return bulk(multiSel, { move: dest });
    const m = current;
    if (!m || !api || guardDemo(m)) return;
    const r = await api.moveMessages(m.accountId, m.folder, [m.uid], dest);
    if (!r.ok) return setStatus(r.error);
    removeFromList(m);
    setStatus('Správa presunutá.');
  }
  async function toggleSeen() {
    // viac správ: ak je niektorá neprečítaná, všetky označiť ako prečítané, inak ako neprečítané
    if (multiSel.length > 1) return bulk(multiSel, multiSel.some((x) => !x.seen) ? 'seen' : 'unseen');
    const m = current;
    if (!m || !api || guardDemo(m)) return;
    const seen = !m.seen;
    const r = await api.setFlags(m.accountId, m.folder, [m.uid], { seen });
    if (!r.ok) return setStatus(r.error);
    patchMsg(m, { seen });
    setUnread((u) => ({ ...u, [m.accountId]: Math.max(0, (u[m.accountId] ?? 0) + (seen ? -1 : 1)) }));
  }
  async function toggleFlag() {
    if (multiSel.length > 1) return bulk(multiSel, multiSel.some((x) => !x.flagged) ? 'flag' : 'unflag');
    const m = current;
    if (!m || !api || guardDemo(m)) return;
    const flagged = !m.flagged;
    const r = await api.setFlags(m.accountId, m.folder, [m.uid], { flagged });
    if (!r.ok) return setStatus(r.error);
    patchMsg(m, { flagged });
  }

  async function setCategory(id: string, on: boolean) {
    if (multiSel.length > 1) return bulk(multiSel, { cat: id, on });
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
    if (multiSel.length > 1) return bulk(multiSel, multiSel.some((x) => !x.keywords?.includes(PIN_KEYWORD)) ? 'pin' : 'unpin');
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
    setStatus(d.uids.length > 1 ? `Presunuté: ${d.uids.length} ${d.uids.length < 5 ? 'správy' : 'správ'}.` : 'Správa presunutá.');
  }
  function saveCurrentSearch() {
    const q = search.trim();
    if (!q) return;
    const fl = SEARCH_FIELDS.find((x) => x.id === searchField);
    const s2: SavedSearch = { id: Math.random().toString(36).slice(2, 10), name: searchField === 'all' ? q : `${q} (${fl?.label.toLowerCase()})`, query: q, accountId: selection.accountId === ALL ? undefined : selection.accountId, field: searchField };
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

  // pravý panel: spoločný alebo vlastný pre konto (vlastný môže zdieľať viac kont)
  const rPanels = settings.rightPanels ?? [];
  const rPanelId = settings.rightPanelOf?.[activeKey] ?? 'shared';
  const rPanel = rPanels.find((p) => p.id === rPanelId);
  const panelView: AppSettings = rPanel ? { ...settings, webs: rPanel.webs, launchers: rPanel.launchers, icsFeeds: rPanel.icsFeeds } : settings;
  const savePanelView = (v: AppSettings) =>
    saveSettings(
      rPanel
        ? { ...settings, rightPanels: rPanels.map((p) => (p.id === rPanel.id ? { ...p, webs: v.webs, launchers: v.launchers, icsFeeds: v.icsFeeds } : p)) }
        : { ...settings, webs: v.webs, launchers: v.launchers, icsFeeds: v.icsFeeds },
    );
  const allPanelWebs = [...settings.webs, ...rPanels.flatMap((p) => p.webs)];
  const panelKeyName = (k: string) => (k === ALL ? 'Všetky kontá' : railAccounts.find((a) => a.id === k)?.name ?? '');
  const panelUsers = (id: string) =>
    [ALL, ...railAccounts.map((a) => a.id)]
      .filter((k) => (k !== ALL || imapAccounts.length > 1) && (settings.rightPanelOf?.[k] ?? 'shared') === id)
      .map(panelKeyName)
      .filter(Boolean);
  const panelChoice: PanelChoice = {
    accountName: panelKeyName(activeKey) || 'toto konto',
    current: rPanel ? rPanel.id : 'shared',
    panels: rPanels.map((p) => ({ id: p.id, name: p.name, users: panelUsers(p.id) })),
    sharedUsers: panelUsers('shared'),
    onChoose: (id) => {
      const of = { ...(settings.rightPanelOf ?? {}) };
      if (id === 'shared') delete of[activeKey];
      else of[activeKey] = id;
      saveSettings({ ...settings, rightPanelOf: of });
    },
    onNew: (copy) => {
      const nid = () => Math.random().toString(36).slice(2, 10);
      const id = `p-${nid()}`;
      const base = panelKeyName(activeKey) || 'Vlastný panel';
      const name = rPanels.some((p) => p.name === base) ? `${base} ${rPanels.length + 1}` : base;
      // kópia dostane nové id webov = vlastné prihlásenie, oddelené od pôvodného panela
      const webs = copy ? panelView.webs.map((w) => ({ ...w, id: nid() })) : [];
      const launchers = copy ? panelView.launchers.map((l) => ({ ...l, id: nid() })) : [];
      saveSettings({ ...settings, rightPanels: [...rPanels, { id, name, webs, launchers, icsFeeds: copy ? panelView.icsFeeds?.map((f) => ({ ...f })) : undefined }], rightPanelOf: { ...(settings.rightPanelOf ?? {}), [activeKey]: id } });
      if (webs[0]) setRightActive(webs[0].id);
    },
    onRename: (id, name) => saveSettings({ ...settings, rightPanels: rPanels.map((p) => (p.id === id ? { ...p, name } : p)) }),
    onDelete: (id) => {
      const of = Object.fromEntries(Object.entries(settings.rightPanelOf ?? {}).filter(([, v]) => v !== id));
      saveSettings({ ...settings, rightPanels: rPanels.filter((p) => p.id !== id), rightPanelOf: of });
    },
  };

  // náhľad kalendára a úloh vedľa čítacieho panela (len pošta IMAP/Zimbra)
  // online schôdze v udalostiach: Teams cez Outlook kontá, Jitsi
  const meetProps = useMemo(
    () => ({
      settings: settings.meetings ?? {},
      outlooks: accounts.filter(isM365).filter((a) => a.host !== 'web').map((a) => ({ id: a.id, name: a.name || a.email, personal: a.host === 'live' })),
      onSettings: (m: NonNullable<AppSettings['meetings']>) => saveSettings({ ...settings, meetings: m }),
    }),
    [settings, accounts], // eslint-disable-line react-hooks/exhaustive-deps
  );
  const myDay = settings.myDay ?? { calendar: false, tasks: false };
  const myDayOn = myDay.calendar || myDay.tasks;
  const [myMonth, setMyMonth] = useState(() => new Date(new Date().getFullYear(), new Date().getMonth(), 1));
  const [myEvents, setMyEvents] = useState<{ loading: boolean; error?: string; list: MyDayEvent[] }>({ loading: false, list: [] });
  const myDayAccs = (activeImap ? [activeImap] : accounts.filter(isImap)).filter((a) => a.id !== DEMO_ID);
  const myDayKey = myDayAccs.map((a) => a.id).join(',');
  useEffect(() => {
    if (!api || demo || !myDay.calendar || !myDayAccs.length) return;
    let alive = true;
    const run = async () => {
      const from = new Date(myMonth.getFullYear(), myMonth.getMonth(), 1 - 7);
      const to = new Date(myMonth.getFullYear(), myMonth.getMonth() + 1, 8);
      setMyEvents((x) => ({ ...x, loading: true }));
      const rs = await Promise.all(myDayAccs.map((a) => api.calEvents(a.id, [], from.toISOString(), to.toISOString())));
      if (!alive) return;
      const list = rs.flatMap((r, i) => (r.ok ? r.data.map((e) => ({ ...e, accountColor: accountColor(myDayAccs[i]!.color) })) : []));
      const errs = rs.filter((r) => !r.ok).length;
      setMyEvents({ loading: false, list, error: errs && !list.length ? 'Kalendár sa nepodarilo načítať.' : undefined });
    };
    void run();
    const t = window.setInterval(() => void run(), 10 * 60_000);
    return () => {
      alive = false;
      window.clearInterval(t);
    };
  }, [api, demo, myDay.calendar, myDayKey, myMonth, agendaTick]); // eslint-disable-line react-hooks/exhaustive-deps
  const myDayModule = (mod: ImapModule) => {
    const a = activeImap ?? myDayAccs[0];
    if (!a) return;
    if (!activeImap) setSelection({ accountId: a.id, folder: 'INBOX' });
    setImapModule((x) => ({ ...x, [a.id]: mod }));
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
    if (!api || POP_COMPOSE || POP_VIEW) return;
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
    if (typing || dialog.kind !== 'none' || activeTab) return;
    if (e.key === 'Enter' && current && !ctrl) return e.preventDefault(), openFull(current, e.shiftKey);
    if (ctrl && e.shiftKey && k === 'n') return e.preventDefault(), void api?.newWindow();
    if (ctrl && k === 'n') return e.preventDefault(), compose('new');
    if (ctrl && e.shiftKey && k === 'r') return e.preventDefault(), full.msg && compose('replyAll');
    if (ctrl && k === 'r') return e.preventDefault(), full.msg && compose('reply');
    if (ctrl && k === 'f') return e.preventDefault(), full.msg && compose('forward');
    if (ctrl && k === 'q') return e.preventDefault(), current && !current.seen && void toggleSeen();
    if (ctrl && k === 'u') return e.preventDefault(), current && current.seen && void toggleSeen();
    if (ctrl && k === 'p') return e.preventDefault(), current && void api?.printMessage(current.accountId, current.folder, current.uid, false);
    if (e.key === 'Insert') return e.preventDefault(), current && void toggleFlag();
    if (e.key === 'Delete' && multiSel.length > 1) return e.preventDefault(), void bulk(multiSel, 'delete');
    if (e.key === 'Delete' && current) return e.preventDefault(), void del();
    if (ctrl && e.shiftKey && k === 'v') return e.preventDefault(), setStatus('Presun: použite Domov → Presunúť alebo pretiahnite správu na priečinok.');
    if (e.key === 'F9') return e.preventDefault(), void sendReceive();
    if (ctrl && ['1', '2', '3', '4'].includes(e.key) && !activeM365) {
      e.preventDefault();
      const mod = (['mail', 'calendar', 'contacts', 'tasks'] as ImapModule[])[Number(e.key) - 1]!;
      if (mod === 'calendar' && !activeImap) return;
      setImapModule((x) => ({ ...x, [activeImap?.id ?? ALL_MODULES]: mod }));
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
    if (k === 'backup' || k === 'restore') return setTransfer(k === 'backup' ? 'export' : 'import');
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
  /** Poslať ďalej: bežne (citácia + prílohy), ako prílohu (.eml) alebo ako originál (presmerovanie) */
  function forwardAs(mode: 'inline' | 'attachment' | 'redirect' = 'inline') {
    const m = full.msg;
    if (mode === 'inline' || !m || m.accountId === DEMO_ID) return compose('forward');
    if (mode === 'redirect') return setRedirect(m);
    const me = imapAccounts.find((a) => a.id === m.accountId)?.email ?? '';
    if (activeImap) setImapModule((x) => ({ ...x, [activeImap.id]: 'mail' }));
    setRibbonTab('message');
    setCompose({ init: { ...replyInit(m, 'forward', m.accountId, me), forwardAsAttachment: true }, key: Date.now() });
  }
  function compose(kind: ComposeKind) {
    // už je otvorená rozpísaná správa: presunúť ju do samostatného okna (nič sa nestratí), potom otvoriť novú
    // predchádzajúca rozpísaná správa ostane ako karta dole
    composeNow(kind);
  }
  function composeNow(kind: ComposeKind) {
    const m = full.msg;
    const accountId = m && m.accountId !== DEMO_ID ? m.accountId : activeImap && activeImap.id !== DEMO_ID ? activeImap.id : imapAccounts.find((a) => a.id !== DEMO_ID)?.id;
    if (activeImap) setImapModule((x) => ({ ...x, [activeImap.id]: 'mail' }));
    setRibbonTab('message');
    if (kind === 'new' || !m) {
      setCompose({ init: { kind: 'new', accountId }, key: Date.now() });
      return;
    }
    const me = imapAccounts.find((a) => a.id === m.accountId)?.email ?? '';
    setCompose({ init: replyInit(m, kind, accountId, me), key: Date.now() });
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

  // lokálny adresár (v programe) a naposledy použité adresy – na dopĺňanie aj do modulu Kontakty
  const [bookTick, setBookTick] = useState(0);
  const [localBook, setLocalBook] = useState<Contact[]>([]);
  useEffect(() => {
    if (!api) return;
    void api.localContacts().then((r) => r.ok && setLocalBook(r.data.flatMap((c) => c.emails.map((email) => ({ name: c.name, email })))));
    if (bookTick && !demo) for (const a of accounts.filter(isImap)) void api.cardList(a.id).then((r) => r.ok && setDavContacts((x) => ({ ...x, [a.id]: r.data.flatMap((c) => c.emails.map((email) => ({ name: c.name, email }))) })));
  }, [api, bookTick]); // eslint-disable-line react-hooks/exhaustive-deps
  const bookAccounts = useMemo(() => imapAccounts.filter((a) => a.id !== DEMO_ID), [imapAccounts]);
  // naposledy použité adresy príjemcov – v hlavnom procese (spoločné pre všetky okná); starý zoznam z localStorage sa raz prevezme
  const [recentContacts, setRecentContacts] = useState<Contact[]>([]);
  const [recentTick, setRecentTick] = useState(0);
  useEffect(() => {
    if (!api) return;
    void (async () => {
      try {
        const old = localStorage.getItem('contacts');
        if (old) {
          const r = await api.recentAdd(JSON.parse(old) as Contact[], true);
          if (r.ok) localStorage.removeItem('contacts');
        }
      } catch {
        /* poškodený starý zoznam – ignorovať */
      }
      const r = await api.recentList();
      if (r.ok) setRecentContacts(r.data.map((x) => ({ name: x.name, email: x.email })));
    })();
  }, [api, recentTick]);

  const contacts = useMemo<Contact[]>(() => {
    const known: Contact[] = [...recentContacts, ...localBook, ...Object.values(davContacts).flat()];
    const all = [...known, ...messages.map((m) => ({ name: m.fromName, email: m.fromAddress }))].filter((c) => c.email);
    return all.filter((c, i) => all.findIndex((x) => x.email.toLowerCase() === c.email.toLowerCase()) === i);
  }, [messages, davContacts, localBook, recentContacts]);

  // hľadanie: fulltext v lokálnej cache (všetky uložené priečinky konta), na požiadanie aj na serveri
  const [found, setFound] = useState<{ q: string; list: MessageSummary[]; server?: 'loading' | 'done' } | null>(null);
  const searchScope = selection.accountId === ALL || searchAll ? undefined : selection.accountId;
  const [searchField, setSearchField] = useState<SearchField>('all');
  useEffect(() => setSearchField(settings.searchField ?? 'all'), [settings.searchField]);
  const serverAuto = settings.searchServerAuto !== false;
  useEffect(() => {
    const q = search.trim();
    if (!api || demo || q.length < 2) {
      setFound(null);
      return;
    }
    const t = window.setTimeout(() => {
      void api.searchLocal(q, searchScope, undefined, searchField).then((r) => {
        if (!r.ok) return;
        setFound({ q, list: r.data });
        if (serverAuto) void searchOnServer({ q, list: r.data });
      });
    }, serverAuto ? 600 : 250);
    return () => window.clearTimeout(t);
  }, [search, searchScope, api, demo, mailTick, searchField, serverAuto]); // eslint-disable-line react-hooks/exhaustive-deps

  // nedávne hľadania: uložiť, keď sa hľadanie chvíľu nemení
  useEffect(() => {
    const q = search.trim();
    if (q.length < 2) return;
    const t = window.setTimeout(() => {
      const h = settings.searchHistory ?? [];
      if (h[0] === q) return;
      saveSettings({ ...settings, searchHistory: [q, ...h.filter((x) => x !== q)].slice(0, 10) });
    }, 2500);
    return () => window.clearTimeout(t);
  }, [search]); // eslint-disable-line react-hooks/exhaustive-deps

  async function searchOnServer(base?: { q: string; list: MessageSummary[] }) {
    const cur = base ?? found;
    if (!api || !cur) return;
    const q = cur.q;
    setFound({ ...cur, server: 'loading' });
    const targets = searchScope ? [{ id: searchScope, folder: selection.folder }] : accounts.filter(isImap).map((a) => ({ id: a.id, folder: 'INBOX' }));
    const rs = await Promise.all(targets.map((t) => api.searchServer(t.id, t.folder, q, searchField)));
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
  const m365Base = (acc: M365Account) => (acc.host === 'web' ? WEBMAIL_MODULES : acc.host === 'live' ? M365_PERSONAL_MODULES : M365_WORK_MODULES);
  const [railEdit, setRailEdit] = useState(false);
  const railOff = !!activeM365?.railOff;
  function selectM365Module(acc: M365Account, mod: M365Module) {
    const st = tabs[acc.id] ?? initialTabs(acc);
    if (mod === 'mail') return setTabs((t) => ({ ...t, [acc.id]: { ...st, active: baseKey(acc.id) } }));
    const key = `${baseKey(acc.id)}:${mod}`;
    const exists = st.tabs.some((x) => x.key === key);
    // vlastný odkaz v páse aplikácií
    const cust = (mod as string).startsWith('custom:') ? acc.railCustom?.find((c) => c.key === mod) : undefined;
    if (cust) {
      setTabs((t) => ({ ...t, [acc.id]: { tabs: exists ? st.tabs : [...st.tabs, { key, title: cust.label, url: cust.url }], active: key } }));
      return;
    }
    setTabs((t) => ({
      ...t,
      [acc.id]: {
        tabs: exists ? st.tabs : [...st.tabs, { key, title: acc.host === 'web' ? 'Webmail' : MODULE_TITLES[mod], url: moduleUrl(mod, acc.host, acc.webUrl), module: mod }],
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

  const curImapModule: ImapModule = imapModule[activeImap?.id ?? ALL_MODULES] ?? 'mail';
  // samostatné okno písania: načítať koncept a zobraziť len editor
  const [popInit, setPopInit] = useState<ComposeInit | null>(null);
  const [popDone, setPopDone] = useState(false);
  const popStarted = useRef(false);
  useEffect(() => {
    if (POP_COMPOSE && popDone && !undo && !composeView) window.close();
  }, [popDone, undo, composeView]);
  useEffect(() => {
    if (!POP_COMPOSE || !api || popStarted.current) return;
    if (POP_COMPOSE.mode && !accounts.length) return;
    popStarted.current = true;
    if (POP_COMPOSE.mode) {
      void api.getMessage(POP_COMPOSE.acc, POP_COMPOSE.folder, POP_COMPOSE.uid).then((r) => {
        if (!r.ok) return setStatus(r.error);
        const me = accounts.find((a) => a.id === POP_COMPOSE.acc);
        const att = POP_COMPOSE.mode === 'forwardAttach';
        const init = replyInit(r.data, (att ? 'forward' : POP_COMPOSE.mode) as 'reply', POP_COMPOSE.acc, me && isImap(me) ? me.email : '');
        setPopInit(att ? { ...init, forwardAsAttachment: true } : init);
      });
      return;
    }
    void api.openDraft(POP_COMPOSE.acc, POP_COMPOSE.folder, POP_COMPOSE.uid).then((r) => {
      if (!r.ok) return setStatus(r.error);
      const d = r.data.msg;
      setPopInit({ kind: 'new', accountId: POP_COMPOSE.acc, to: d.to, cc: d.cc, subject: d.subject === '(bez predmetu)' ? '' : d.subject, html: d.html ?? `<p>${d.text}</p>`, draft: { folder: POP_COMPOSE.folder, uid: POP_COMPOSE.uid }, files: r.data.files });
    });
  }, [api, accounts.length]); // eslint-disable-line react-hooks/exhaustive-deps
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
              onQueued={() => undefined /* stav odosielania a „Späť“ ukáže hlavné okno */}
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

  if (POP_VIEW) {
    return (
      <FluentProvider theme={getTheme(themeDef)}>
        <div className={`app theme-${theme} ${themeDef.dark ? 'dark' : ''} pop-view`} style={themeVars(themeDef)}>
          <MessageView start={POP_VIEW} standalone color={colorOf(POP_VIEW.accountId)} categories={categoriesOf(settings)} />
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
        {lockDlg && <LockSettings onClose={() => setLockDlg(false)} />}
        {redirect && <RedirectDialog msg={redirect} onClose={() => setRedirect(null)} onDone={(t) => setStatus(t)} />}
        {transfer && <SettingsTransfer mode={transfer} onClose={() => setTransfer(null)} onStatus={setStatus} />}
      <div className="shell">
        <AccountRail
          accounts={railAccounts}
          active={summaryOn ? SUMMARY : activeKey}
          unread={unread}
          showAll={imapAccounts.length > 1 && settings.showAllInboxes !== false}
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
          showSummary={summaryEnabled}
          onToggleSummary={toggleSummary}
          onCrypto={() => openDialog({ kind: 'crypto' })}
          onAbout={() => openDialog({ kind: 'about' })}
          onTransfer={(m) => setTransfer(m)}
          onLockSettings={() => setLockDlg(true)}
          kindOf={(a) => (a.id === DEMO_ID ? '' : isM365(a) ? (a.host === 'web' ? 'WEB' : a.host === 'live' ? 'OS' : '365') : isGmail(a as ImapAccountPublic) ? 'GMAIL' : zimbraOk[a.id] ? 'Z' : 'IMAP')}
          onTestNotification={() => void api?.testNotification().then((r) => setStatus(r.ok ? r.data : r.error))}
          onToggleNotifications={() => saveSettings({ ...settings, notifications: settings.notifications === false })}
          onToggleGpu={() => void api?.setGpuDisabled(!gpuOff)}
          gpuForced={gpuForced}
          onToggleGpuForced={() => void api?.setGpuForced(!gpuForced)}
          onOpenLog={() => void api?.openLog()}
          pinned={!!settings.accountListPinned}
          onPin={(v) => saveSettings({ ...settings, accountListPinned: v })}
          onMenuOpen={setMenuOverlay}
        />
        <div className={`workspace ${summaryOn ? 'with-summary' : ''}`}>
          {summaryOn && (
            <SummaryPage
              accounts={railAccounts.filter((a) => a.id !== DEMO_ID)}
              unread={unread}
              colorOf={colorOf}
              calendarAccounts={imapAccounts.filter((a) => a.id !== DEMO_ID).map((a) => ({ id: a.id, name: a.name }))}
              tasks={settings.tasks ?? []}
              quota={quota}
              settings={settings.summary ?? {}}
              onSettings={(sm) => saveSettings({ ...settings, summary: sm })}
              onOpenAccount={(id) => setActiveId(id)}
              onOpenMessage={(m) => openTaskMessage({ accountId: m.accountId, folder: m.folder, uid: m.uid, subject: m.subject, from: m.fromName || m.fromAddress })}
              onOpenTasks={() => {
                setActiveId(ALL);
                setImapModule((x) => ({ ...x, [ALL_MODULES]: 'tasks' }));
              }}
              onToggleTask={(t) => saveSettings({ ...settings, tasks: (settings.tasks ?? []).map((x) => (x.id === t.id ? { ...x, done: !x.done, dirty: true } : x)) })}
              onOpenCalendar={(id) => {
                setActiveId(id);
                setImapModule((x) => ({ ...x, [id]: 'calendar' }));
              }}
              onDisable={() => {
                saveSettings({ ...settings, summary: { ...(settings.summary ?? {}), enabled: false } });
                setActiveId(localStorage.getItem('lastAccount') ?? '');
                setStatus('Súhrn je vypnutý. Zapneš ho v karte Zobrazenie alebo v ozubenom kolieske dole vľavo.');
              }}
            />
          )}
          <TopBar
            kind={activeM365 ? (activeM365.host === 'web' ? 'WEBMAIL' : activeM365.host === 'live' ? 'OUTLOOK · OSOBNÉ' : 'OUTLOOK · FIREMNÉ') : activeImap && activeImap.id !== DEMO_ID ? (isGmail(activeImap) ? 'GMAIL' : zimbraOk[activeImap.id] ? 'ZIMBRA' : 'IMAP') : undefined}
            title={topTitle}
            subtitle={topSub}
            color={topColor}
            search={search}
            onSearch={activeM365 ? undefined : setSearch}
            searchOptions={{
              field: searchField,
              onField: (f) => saveSettings({ ...settings, searchField: f }),
              serverAuto,
              onServerAuto: (v) => saveSettings({ ...settings, searchServerAuto: v }),
              history: settings.searchHistory ?? [],
              onHistory: (q) => setSearch(q),
              onClearHistory: () => saveSettings({ ...settings, searchHistory: [] }),
              onSaveFolder: search.trim() ? saveCurrentSearch : undefined,
            }}
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
              onForward={(mode) => forwardAs(mode)}
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
              openMode={settings.openMode ?? 'window'}
              onOpenMode={(v) => saveSettings({ ...settings, openMode: v })}
              onSendReceive={() => void sendReceive()}
              myDay={myDay}
              onMyDay={(v) => saveSettings({ ...settings, myDay: v })}
              summaryEnabled={summaryEnabled}
              onToggleSummary={toggleSummary}
              allInboxes={settings.showAllInboxes !== false}
              onToggleAllInboxes={imapAccounts.length > 1 ? toggleAllInboxes : undefined}
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
              Je pripravená nová verzia programu Universum eMail.{' '}
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
          <main
            className={`body ${rightOpen ? 'with-right' : ''} ${railOff ? 'no-rail' : ''}`}
            style={rightOpen ? { gridTemplateColumns: `${railOff ? '0px' : '68px'} minmax(0, 1fr) ${widths.right}px` } : railOff ? { gridTemplateColumns: '0px minmax(0, 1fr)' } : undefined}
          >
            {rightOpen && (
              <Splitter
                side="right"
                at={widths.right}
                value={widths.right}
                min={260}
                max={900}
                label="Šírka pravého panela"
                def={PW.right}
                onChange={(v) => setWidths((w) => ({ ...w, right: v }))}
                onDone={(v) => saveWidth('right', v)}
              />
            )}
            {activeM365 ? (
              activeM365.railOff ? null : (
                <ModuleRail
                  modules={customizeRail(m365Base(activeM365) as ModuleDef<string>[], activeM365)}
                  active={m365ActiveModule}
                  color={topColor}
                  onSelect={(k) => selectM365Module(activeM365, k as M365Module)}
                  footer={
                    <Tooltip content="Upraviť pás aplikácií: poradie, skryť/pridať, vypnúť" relationship="label" positioning="after">
                      <button className="rail-btn rail-edit-btn" onClick={() => setRailEdit(true)} aria-label="Upraviť pás aplikácií">
                        <EditSmallIcon />
                      </button>
                    </Tooltip>
                  }
                />
              )
            ) : (
              <ModuleRail
                modules={activeImap ? IMAP_MODULES : IMAP_MODULES.filter((m) => m.key !== 'calendar')}
                active={curImapModule}
                color={topColor}
                onSelect={(k) => setImapModule((m) => ({ ...m, [activeImap?.id ?? ALL_MODULES]: k }))}
                footer={
                  activeImap && activeImap.id !== DEMO_ID ? (
                    <AccountRailFooter
                      accountName={activeImap.name || activeImap.email}
                      onZimbra={zimbraOk[activeImap.id] ? (tab) => openDialog({ kind: 'zimbra', tab }) : undefined}
                      onAccount={() => editAccount(activeImap)}
                    />
                  ) : undefined
                }
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
                  onReorder={(order) => setTabs((t) => ({ ...t, [a.id]: { ...t[a.id]!, order } }))}
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
                <div className="mail-wrap">
                {msgTabs.length > 0 && (
                  <div className="msgtabs" role="tablist" aria-label="Otvorené správy">
                    <button role="tab" aria-selected={!activeTab} className={`msgtab ${!activeTab ? 'on' : ''}`} onClick={() => setActiveTab(null)}>
                      <MailIcon fontSize={16} /> <span className="t">Pošta</span>
                    </button>
                    {msgTabs.map((t) => (
                      <div
                        key={t.key}
                        role="tab"
                        tabIndex={0}
                        aria-selected={activeTab === t.key}
                        className={`msgtab ${activeTab === t.key ? 'on' : ''}`}
                        title={t.title}
                        onClick={() => setActiveTab(t.key)}
                        onAuxClick={(e) => e.button === 1 && closeMsgTab(t.key)}
                      >
                        <span className="t">{t.title}</span>
                        <button className="x" aria-label={`Zavrieť ${t.title}`} onClick={(e) => (e.stopPropagation(), closeMsgTab(t.key))}>×</button>
                      </div>
                    ))}
                  </div>
                )}
                {msgTabs.map((t) => (
                  <div key={t.key} className="msgtab-pane" style={{ display: activeTab === t.key ? undefined : 'none' }}>
                    <MessageView
                      start={t.ref}
                      color={colorOf(t.ref.accountId)}
                      categories={categoriesOf(settings)}
                      onTitle={(title) => setMsgTabs((all) => all.map((x) => (x.key === t.key ? { ...x, title } : x)))}
                      onPopOut={(r) => {
                        void api?.messageWindow(r.accountId, r.folder, r.uid);
                        closeMsgTab(t.key);
                      }}
                      onClose={() => closeMsgTab(t.key)}
                    />
                  </div>
                ))}
                <div className="mail" style={{ display: activeTab ? 'none' : undefined, gridTemplateColumns: `${widths.folders}px ${widths.list}px minmax(0, 1fr)${myDayOn ? ` ${widths.myday}px` : ''}` }}>
                  <Splitter
                    at={widths.folders}
                    value={widths.folders}
                    min={150}
                    max={480}
                    label="Šírka panela priečinkov"
                    def={PW.folders}
                    onChange={(v) => setWidths((w) => ({ ...w, folders: v }))}
                    onDone={(v) => saveWidth('folders', v)}
                  />
                  {myDayOn && (
                    <Splitter
                      side="right"
                      at={widths.myday}
                      value={widths.myday}
                      min={220}
                      max={520}
                      label="Šírka panela kalendára a úloh"
                      def={PW.myday}
                      onChange={(v) => setWidths((w) => ({ ...w, myday: v }))}
                      onDone={(v) => saveWidth('myday', v)}
                    />
                  )}
                  <Splitter
                    at={widths.folders + widths.list}
                    value={widths.list}
                    min={220}
                    max={900}
                    label="Šírka zoznamu správ"
                    def={PW.list}
                    onChange={(v) => setWidths((w) => ({ ...w, list: v }))}
                    onDone={(v) => saveWidth('list', v)}
                  />
                  <FolderPane
                    account={activeImap}
                    accounts={imapAccounts}
                    folders={folders}
                    unread={unread}
                    selection={selection}
                    onSelect={(s) => {
                      // klik na priečinok (aj ten istý) počas písania: správa ostane ako karta dole, pás → Domov
                      if (composeView) {
                        setCompose(null);
                        setRibbonTab('home');
                      }
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
                      setSearchField(x.field ?? 'all');
                      setSearch(x.query);
                    }}
                    onFolderContext={folderMenu}
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
                    onOpenFull={(m, alt) => openFull(m, alt)}
                    onContext={(msgs, x, y) => messageMenu(msgs, x, y)}
                    onDeleteMany={(msgs) => void bulk(msgs, 'delete')}
                    onMulti={setMultiSel}
                    listKey={`${selection.accountId}/${selection.folder}/${mailTick}`}
                    folderKey={`${selection.accountId}/${selection.folder}`}
                    onFlaggedServer={async () => {
                      if (!api || demo) return [];
                      const targets = selection.accountId === ALL ? imapAccounts.filter((a) => a.id !== DEMO_ID).map((a) => ({ id: a.id, folder: 'INBOX' })) : [{ id: selection.accountId, folder: selection.folder }];
                      const rs = await Promise.all(targets.map((t) => api.searchServer(t.id, t.folder, '', 'flagged')));
                      return rs.flatMap((r) => (r.ok ? r.data : [])).sort((a, b) => (b.date ?? '').localeCompare(a.date ?? ''));
                    }}
                    onQuick={(m, a) =>
                      // ikona na riadku, ktorý je súčasťou výberu viacerých správ → platí pre celý výber
                      void bulk(multiSel.length > 1 && multiSel.some((x) => x.uid === m.uid && x.accountId === m.accountId && x.folder === m.folder) ? multiSel : [m], a === 'delete' ? 'delete' : a === 'seen' ? (m.seen ? 'unseen' : 'seen') : a === 'flag' ? (m.flagged ? 'unflag' : 'flag') : m.keywords?.includes(PIN_KEYWORD) ? 'unpin' : 'pin')
                    }
                    sort={settings.listSort}
                    onSort={(s) => saveSettings({ ...settings, listSort: s })}
                    onLoadMore={!found && hasMore && !demo ? loadMore : undefined}
                    loadingMore={listState.loading}
                    onDeleteKey={() => void del()}
                  />
                  <div className="reading-slot">
                  {drafts.map((d) => (
                    <div key={d.key} className="draft-pane" style={{ display: d.key === activeDraft ? undefined : 'none' }}>
                    <Composer
                      ref={(h) => {
                        composerRefs.current[d.key] = h;
                      }}
                      init={d.init}
                      accounts={imapAccounts.filter((a) => a.id !== DEMO_ID).length ? imapAccounts.filter((a) => a.id !== DEMO_ID) : imapAccounts}
                      settings={settings}
                      contacts={contacts}
                      zimbra={zimbraOk}
                      onState={(st) => setComposerStates((m) => ({ ...m, [d.key]: st }))}
                      onQueued={(it) => it.kind === 'undo' && setUndo(it)}
                      onSaveTemplate={saveTemplate}
                      onClose={(st) => {
                        removeDraft(d.key);
                        if (d.key === activeDraft) setRibbonTab('home');
                        if (st) setStatus(st);
                      }}
                    />
                    </div>
                  ))}
                  {!composeView && (
                  <ReadingPane
                    message={full.msg}
                    loading={full.loading}
                    error={full.error}
                    color={colorOf(full.msg?.accountId ?? '')}
                    onReply={() => compose('reply')}
                    onReplyAll={() => compose('replyAll')}
                    onForward={(mode) => forwardAs(mode)}
                    thread={thread}
                    onOpenThread={(t) => void open(t)}
                    categories={categoriesOf(settings)}
                    onRemoveCategory={(id) => void setCategory(id, false)}
                    gmail={(() => {
                      const m = full.msg;
                      const acc = m && accounts.find((a) => a.id === m.accountId);
                      if (!m || !acc || !isImap(acc) || !isGmail(acc)) return undefined;
                      const fl = folders[acc.id];
                      return {
                        folders: Array.isArray(fl) ? fl : [],
                        onChanged: (removedCurrent: boolean) => {
                          if (removedCurrent) setRefreshKey((k) => k + 1);
                          void api?.notifyChanged(m.accountId, m.folder);
                        },
                      };
                    })()}
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
                    onAttachmentContext={(i, x, y) => {
                      const m = full.msg;
                      if (!m || !api) return;
                      const a = m.attachments.find((x, pos) => (x.index ?? pos) === i);
                      const real = m.attachments.map((x, pos) => ({ x, i: x.index ?? pos })).filter((y) => !y.x.inline);
                      showCtx(x, y, [
                        {
                          label: attachmentRisk(a?.filename ?? '').level === 'blocked' ? 'Otvoriť (zablokované – nebezpečný typ)' : 'Otvoriť…',
                          disabled: attachmentRisk(a?.filename ?? '').level === 'blocked',
                          onClick: () => attGuard(a?.filename ?? '', () => void api.openAttachment(m.accountId, m.folder, m.uid, i), () => void api.saveAttachment(m.accountId, m.folder, m.uid, i)),
                        },
                        { label: 'Uložiť ako…', onClick: () => void api.saveAttachment(m.accountId, m.folder, m.uid, i) },
                        { label: 'Uložiť všetky prílohy…', disabled: real.length < 2, onClick: async () => {
                          for (const { i: k } of real) {
                            const r = await api.saveAttachment(m.accountId, m.folder, m.uid, k);
                            if (r.ok && r.data === null) break;
                          }
                        } },
                        { divider: true },
                        { label: 'Poslať ďalej s prílohami', onClick: () => forwardAs('inline') },
                        { label: 'Poslať ďalej ako prílohu…', onClick: () => forwardAs('attachment') },
                        { label: 'Poslať ďalej ako originál…', onClick: () => forwardAs('redirect') },
                        { label: `Kopírovať názov „${a?.filename ?? ''}“`, onClick: () => void navigator.clipboard.writeText(a?.filename ?? '') },
                      ]);
                    }}
                    onDelete={() => void del()}
                    onArchive={() => void archive()}
                    onToggleSeen={() => void toggleSeen()}
                    onToggleFlag={() => void toggleFlag()}
                    onTogglePin={() => void togglePin()}
                    onMore={current ? (x, y) => messageMenu([current], x, y) : undefined}
                    onPopOut={full.msg && full.msg.accountId !== DEMO_ID ? () => void api?.messageWindow(full.msg!.accountId, full.msg!.folder, full.msg!.uid) : undefined}
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
                      return r.ok ? { ok: true, text: r.data } : { ok: false, text: r.error };
                    }}
                  />
                  )}
                  </div>
                  {myDayOn && (
                    <MyDayPane
                      calendar={myDay.calendar}
                      tasks={myDay.tasks}
                      events={myEvents.list}
                      loading={myEvents.loading}
                      error={myEvents.error}
                      month={myMonth}
                      onMonth={setMyMonth}
                      taskList={settings.tasks ?? []}
                      onToggleTask={(t) => saveSettings({ ...settings, tasks: (settings.tasks ?? []).map((x) => (x.id === t.id ? { ...x, done: !x.done, dirty: true } : x)) })}
                      onOpenTask={(t) => (t.message ? openTaskMessage(t.message) : myDayModule('tasks'))}
                      onAddTask={(title, due) =>
                        saveSettings({ ...settings, tasks: [...(settings.tasks ?? []), { id: Math.random().toString(36).slice(2, 10), title, due, done: false, created: new Date().toISOString() }] })
                      }
                      onOpenCalendar={() => myDayModule('calendar')}
                      onOpenTasks={() => myDayModule('tasks')}
                      onHide={(part) => saveSettings({ ...settings, myDay: { ...myDay, [part]: false } })}
                    />
                  )}
                </div>
                {drafts.length > 0 && (
                  <div className="draft-tabs" role="tablist" aria-label="Rozpísané správy">
                    {drafts.map((d) => {
                      const st = composerStates[d.key];
                      const t = st?.subject?.trim() || (d.init.subject ?? '').trim() || '(Bez predmetu)';
                      const on = d.key === activeDraft;
                      return (
                        <div
                          key={d.key}
                          role="tab"
                          tabIndex={0}
                          aria-selected={on}
                          className={`draft-tab ${on ? 'on' : ''}`}
                          title={t}
                          onClick={() => {
                            setActiveTab(null);
                            setActiveDraft(d.key);
                            setRibbonTab('message');
                          }}
                          onAuxClick={(e) => e.button === 1 && composerRefs.current[d.key]?.close()}
                        >
                          <span className="pen" aria-hidden="true">✎</span>
                          <span className="t">{t}</span>
                          <button
                            className="x"
                            aria-label={`Zavrieť ${t} (uloží sa do Konceptov)`}
                            title="Zavrieť – rozpísaná správa sa uloží do Konceptov"
                            onClick={(e) => {
                              e.stopPropagation();
                              composerRefs.current[d.key]?.close();
                            }}
                          >
                            ×
                          </button>
                        </div>
                      );
                    })}
                  </div>
                )}
                </div>
              ) : (
                curImapModule === 'tasks' ? (
                  <TasksView settings={settings} onSettings={saveSettings} onOpenMessage={openTaskMessage} />
                ) : activeImap && activeImap.id !== DEMO_ID && curImapModule === 'calendar' ? (
                  <CalendarView key={activeImap.id} account={activeImap} zimbra={!!zimbraOk[activeImap.id]} onStatus={setStatus} onSettings={() => editAccount(activeImap)} meet={meetProps} />
                ) : curImapModule === 'contacts' ? (
                  <AddressBook
                    accounts={bookAccounts}
                    zimbraOk={zimbraOk}
                    recent={recentContacts}
                    defaultAccountId={activeImap?.id ?? imapAccounts.find((a) => a.id !== DEMO_ID)?.id}
                    onStatus={setStatus}
                    onChanged={() => setBookTick((t) => t + 1)}
                    onCompose={(to, accId) => {
                      const acc = accId ?? activeImap?.id ?? imapAccounts[0]?.id;
                      setImapModule((m) => ({ ...m, [activeImap?.id ?? ALL_MODULES]: 'mail' }));
                      setCompose({ init: { kind: 'new', accountId: acc, to }, key: Date.now() });
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
                settings={panelView}
                onSettings={savePanelView}
                allWebs={allPanelWebs}
                choice={panelChoice}
                active={rightActive}
                onActive={setRightActive}
                overlay={overlay}
                onOverlay={setPanelOverlay}
                onStatus={setStatus}
                agenda={api && !demo ? agenda : undefined}
              />
            )}
          </main>
          {railEdit && activeM365 && (
            <RailEditor
              base={m365Base(activeM365) as ModuleDef<string>[]}
              acc={activeM365}
              onClose={() => setRailEdit(false)}
              onSave={(v) => {
                setRailEdit(false);
                const { hasPassword: _h, webHasPassword: _w, ...plain } = activeM365 as M365Account & { hasPassword?: boolean; webHasPassword?: boolean };
                void api?.saveAccount({ ...plain, ...v }).then(() => void loadAccounts());
                if (v.railOff) setStatus('Pás aplikácií je vypnutý. Zapneš ho v nastaveniach konta (pravý klik na konto → Nastavenia).');
              }}
            />
          )}
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
      <ContextMenu state={ctx} onClose={closeCtx} />
      {attGuardDialog}
      {ask && (
        <Dialog open onOpenChange={(_, d) => !d.open && setAsk(null)}>
          <DialogSurface style={{ maxWidth: 420 }}>
            <DialogBody>
              <DialogTitle>{ask.title}</DialogTitle>
              <DialogContent>
                <Input
                  autoFocus
                  style={{ width: '100%' }}
                  value={ask.value}
                  onChange={(_, d) => setAsk({ ...ask, value: d.value })}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' && ask.value.trim()) {
                      ask.ok(ask.value.trim());
                      setAsk(null);
                    }
                  }}
                />
              </DialogContent>
              <DialogActions>
                <Button appearance="primary" disabled={!ask.value.trim()} onClick={() => (ask.ok(ask.value.trim()), setAsk(null))}>OK</Button>
                <Button onClick={() => setAsk(null)}>Zrušiť</Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      )}
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
          onGmail={() => openDialog({ kind: 'imap', preset: 'gmail' })}
          onM365={(host) => openDialog({ kind: 'm365', host })}
        />
      )}
      {dialog.kind === 'imap' && (
        <AccountDialog
          key={dialogKey}
          open
          edit={dialog.edit}
          preset={dialog.preset}
          google={settings.google}
          onSaveGoogle={(g) => saveSettings({ ...settings, google: g })}
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
          host={dialog.host}
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
