import { FluentProvider } from '@fluentui/react-components';
import { useCallback, useEffect, useMemo, useState } from 'react';
import type { AccountPublic } from '../core/account';
import type { FolderInfo, MailApi, MessageFull, MessageSummary } from '../core/api';
import { AccountDialog } from './components/AccountDialog';
import { ComposeDialog, type ComposeInit } from './components/ComposeDialog';
import { AccountPane, MessageList, ReadingPane, WebColumn, type Selection, type WebKey } from './components/MailPanes';
import { ModuleRail, Ribbon, TopBar, type ModuleKey, type RibbonTab } from './components/Shell';
import { DEMO_ID, demoAccount, demoFolders, demoMessages } from './mock';
import { getTheme, type ThemeKey } from './theme';

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

export function App() {
  const api = window.mailApi;
  const [theme, setTheme] = usePersisted<ThemeKey>('theme', 'blue');
  const [webOpen, setWebOpen] = usePersisted('webOpen', false);
  const [webService, setWebService] = usePersisted<WebKey>('webService', 'copilot');
  const [compact, setCompact] = usePersisted('ribbonCompact', false);
  const [ribbonTab, setRibbonTab] = useState<RibbonTab>('home');
  const [module, setModule] = useState<ModuleKey>('mail');

  const [accounts, setAccounts] = useState<AccountPublic[]>([]);
  const [secretsWarning, setSecretsWarning] = useState(false);
  const [folders, setFolders] = useState<Record<string, FolderState>>({});
  const [selection, setSelection] = useState<Selection>({ accountId: 'all', folder: 'INBOX' });
  const [messages, setMessages] = useState<MessageSummary[]>([]);
  const [listState, setListState] = useState<{ loading: boolean; error?: string }>({ loading: false });
  const [current, setCurrent] = useState<MessageSummary | undefined>();
  const [full, setFull] = useState<{ msg?: MessageFull; loading: boolean; error?: string }>({ loading: false });
  const [search, setSearch] = useState('');
  const [refreshKey, setRefreshKey] = useState(0);
  const [status, setStatus] = useState('');
  const [accDialog, setAccDialog] = useState<{ open: boolean; edit?: AccountPublic; key: number }>({ open: false, key: 0 });
  const [compose, setCompose] = useState<{ open: boolean; init: ComposeInit; key: number }>({ open: false, init: {}, key: 0 });

  const demo = accounts.length === 0;
  const shown = demo ? [demoAccount] : accounts;
  const colorOf = useCallback((id: string) => shown.find((a) => a.id === id)?.color ?? '#888', [shown]);

  // kontá
  const loadAccounts = useCallback(async () => {
    if (!api) return;
    const r = await api.listAccounts();
    if (r.ok) {
      setAccounts(r.data.accounts);
      setSecretsWarning(!r.data.canStoreSecrets);
    } else setStatus(r.error);
  }, [api]);
  useEffect(() => {
    void loadAccounts();
  }, [loadAccounts]);

  // priečinky
  useEffect(() => {
    if (demo) {
      setFolders({ [DEMO_ID]: demoFolders });
      return;
    }
    for (const a of accounts) {
      setFolders((f) => ({ ...f, [a.id]: 'loading' }));
      void api!.listFolders(a.id).then((r) =>
        setFolders((f) => ({ ...f, [a.id]: r.ok ? r.data : { error: r.error } })),
      );
    }
  }, [accounts, demo, api, refreshKey]);

  // správy
  useEffect(() => {
    let cancelled = false;
    (async () => {
      if (demo) {
        setMessages(selection.folder === 'INBOX' ? demoMessages : []);
        return;
      }
      setListState({ loading: true });
      const targets = selection.accountId === 'all' ? accounts.map((a) => ({ id: a.id, folder: 'INBOX' })) : [{ id: selection.accountId, folder: selection.folder }];
      const results = await Promise.all(targets.map((t) => api!.listMessages(t.id, t.folder, 50)));
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
  }, [selection, accounts, demo, api, refreshKey]);

  async function open(m: MessageSummary) {
    setCurrent(m);
    if (m.accountId === DEMO_ID) {
      setFull({ loading: false, msg: demoMessages.find((x) => x.uid === m.uid) });
      return;
    }
    setFull({ loading: true });
    const r = await api!.getMessage(m.accountId, m.folder, m.uid);
    setFull(r.ok ? { loading: false, msg: r.data } : { loading: false, error: `${r.error}${r.hint ? ` ${r.hint}` : ''}` });
    if (r.ok) setMessages((list) => list.map((x) => (x === m ? { ...x, seen: true } : x)));
  }

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return messages;
    return messages.filter((m) => `${m.subject} ${m.fromName} ${m.fromAddress}`.toLowerCase().includes(q));
  }, [messages, search]);

  const title = useMemo(() => {
    if (selection.accountId === 'all') return 'Všetky doručené';
    const fs = folders[selection.accountId];
    const f = Array.isArray(fs) ? fs.find((x) => x.path === selection.folder) : undefined;
    return f?.name ?? selection.folder;
  }, [selection, folders]);

  function reply() {
    const m = full.msg;
    if (!m) return;
    const quoted = m.text
      .split('\n')
      .map((l) => `> ${l}`)
      .join('\n');
    setCompose({
      open: true,
      key: compose.key + 1,
      init: {
        accountId: m.accountId === DEMO_ID ? undefined : m.accountId,
        to: m.fromAddress,
        subject: /^re:/i.test(m.subject) ? m.subject : `Re: ${m.subject}`,
        text: `\n\n${m.date ? new Date(m.date).toLocaleString('sk-SK') : ''}, ${m.fromName} napísal(a):\n${quoted}`,
      },
    });
  }

  const selectModule = (k: ModuleKey) => {
    if (k === 'whatsapp' || k === 'teams' || k === 'copilot') {
      setWebService(k);
      setWebOpen(true);
      return;
    }
    setModule(k);
  };

  const editable = selection.accountId !== 'all' && !demo ? accounts.find((a) => a.id === selection.accountId) : undefined;

  return (
    <FluentProvider theme={getTheme(theme)} className={`app theme-${theme}`}>
      <TopBar search={search} onSearch={setSearch} />
      <Ribbon
        tab={ribbonTab}
        onTab={setRibbonTab}
        canReply={!!full.msg}
        onCompose={() => setCompose({ open: true, init: {}, key: compose.key + 1 })}
        onReply={reply}
        onRefresh={() => setRefreshKey((k) => k + 1)}
        onAddAccount={() => setAccDialog({ open: true, key: accDialog.key + 1 })}
        onEditAccount={editable ? () => setAccDialog({ open: true, edit: editable, key: accDialog.key + 1 }) : undefined}
        theme={theme}
        onTheme={setTheme}
        webOpen={webOpen}
        onToggleWeb={() => setWebOpen(!webOpen)}
        compact={compact}
        onCompact={setCompact}
      />
      {demo && (
        <div className="banner">
          Zobrazujú sa ukážkové dáta. <button className="link" onClick={() => setAccDialog({ open: true, key: accDialog.key + 1 })}>Pridať IMAP konto</button>
        </div>
      )}
      {secretsWarning && (
        <div className="banner warn">
          Systémové úložisko hesiel (keyring) nie je dostupné: heslá sa držia len do zatvorenia aplikácie.
        </div>
      )}
      <main className={`body ${webOpen ? 'with-web' : ''}`}>
        <ModuleRail active={module} onSelect={selectModule} />
        {module === 'mail' ? (
          <>
            <AccountPane accounts={shown} folders={folders} selection={selection} onSelect={(s) => { setSelection(s); setCurrent(undefined); setFull({ loading: false }); }} />
            <MessageList title={title} messages={filtered} loading={listState.loading} error={listState.error} selected={current} colorOf={colorOf} onOpen={open} />
            <ReadingPane message={full.msg} loading={full.loading} error={full.error} color={colorOf(full.msg?.accountId ?? '')} onReply={reply} />
          </>
        ) : (
          <section className="placeholder">
            <h2>{module === 'calendar' ? 'Kalendár' : module === 'tasks' ? 'Úlohy' : 'Spúšťače'}</h2>
            <p className="muted">{module === 'launchers' ? 'Pripravuje sa v etape 3.' : 'Pripravuje sa v etape 2 (CalDAV, Zimbra, MS365).'}</p>
          </section>
        )}
        {webOpen && <WebColumn service={webService} onService={setWebService} onClose={() => setWebOpen(false)} />}
      </main>
      <footer className="statusbar">{status || (demo ? 'Ukážkový režim' : `${accounts.length} ${accounts.length === 1 ? 'konto' : accounts.length < 5 ? 'kontá' : 'kont'}`)}</footer>

      {accDialog.open && (
        <AccountDialog
          key={accDialog.key}
          open
          edit={accDialog.edit}
          usedColors={accounts.map((a) => a.color)}
          onClose={() => setAccDialog({ ...accDialog, open: false })}
          onSaved={(a) => {
            setAccDialog({ ...accDialog, open: false });
            setStatus(`Konto ${a.name} uložené.`);
            void loadAccounts();
          }}
          onDeleted={() => {
            setAccDialog({ ...accDialog, open: false });
            setSelection({ accountId: 'all', folder: 'INBOX' });
            void loadAccounts();
          }}
        />
      )}
      {compose.open && (
        <ComposeDialog
          key={compose.key}
          open
          init={compose.init}
          accounts={accounts}
          onClose={() => setCompose({ ...compose, open: false })}
          onSent={(info) => {
            setCompose({ ...compose, open: false });
            setStatus(info);
          }}
        />
      )}
    </FluentProvider>
  );
}
