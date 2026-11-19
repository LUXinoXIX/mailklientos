// Panel kont, zoznam správ, čítací panel a 4. stĺpec (variant D).
import { useAttachmentGuard } from './AttachmentGuard';
import { Button, Menu, MenuDivider, MenuGroup, MenuGroupHeader, MenuItemRadio, MenuList, MenuPopover, MenuTrigger, Spinner, Tab, TabList } from '@fluentui/react-components';
import DOMPurify from 'dompurify';
import { useEffect, useMemo, useRef, useState } from 'react';
import type { ImapAccountPublic } from '../../core/account';
import type { FolderInfo, MessageFull, MessageSummary, OutboxItem } from '../../core/api';
import { PIN_KEYWORD, type Category, type SavedSearch } from '../../core/settings';
import { PreviewSmallIcon, SaveSmallIcon, AttachTinyIcon, ArchiveIcon, DeleteIcon, FlagIcon, MarkReadIcon, MarkUnreadIcon, MoreIcon, PdfIcon, PinIcon, PrintIcon, DeleteSmallIcon, FlagSmallIcon, PinOutlineSmallIcon, ReadSmallIcon, SortIcon, UnreadSmallIcon, NewWindowSmallIcon, LockSmallIcon, VerifiedIcon, WarnIcon, PinSmallIcon, SearchSmallIcon, AttachIcon, FlagFilledIcon, FolderIcon, ForwardIcon, InboxIcon, ReplyAllIcon, ReplyIcon, SentIcon, TrashIcon } from '../icons';
import { accountColor } from '../theme';

export type Selection = { accountId: string | 'all'; folder: string };

const initials = (name: string) =>
  name
    .split(/[\s.@]+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((p) => p[0]!.toUpperCase())
    .join('') || '?';

export function Avatar({ name, color, size = 32 }: { name: string; color: string; size?: number }) {
  return (
    <span className="avatar" style={{ background: color, width: size, height: size, fontSize: size * 0.4 }}>
      {initials(name)}
    </span>
  );
}

const folderIcon = (f: FolderInfo) =>
  f.specialUse === '\\Inbox' ? <InboxIcon /> : f.specialUse === '\\Sent' ? <SentIcon /> : f.specialUse === '\\Trash' ? <TrashIcon /> : <FolderIcon />;

/** Priečinky vybraného konta; v režime „Všetky“ zoznam doručených pôšt všetkých IMAP kont. */
export function FolderPane(props: {
  account: ImapAccountPublic | null;
  accounts: ImapAccountPublic[];
  folders: Record<string, FolderInfo[] | 'loading' | { error: string }>;
  unread: Record<string, number>;
  selection: Selection;
  onSelect: (s: Selection) => void;
  onDropMessages?: (destFolder: string, d: DragMessages) => void;
  savedSearches?: SavedSearch[];
  activeSearch?: string;
  onSavedSearch?: (s: SavedSearch) => void;
  onRemoveSearch?: (s: SavedSearch) => void;
  scheduled?: OutboxItem[];
  onScheduled?: (it: OutboxItem, edit: boolean) => void;
  quota?: { used: number; limit: number } | null;
  onFolderContext?: (accountId: string, f: FolderInfo, x: number, y: number) => void;
}) {
  const a = props.account;
  const [over, setOver] = useState<string | null>(null);
  const searches = (props.savedSearches ?? []).filter((x) => !x.accountId || !a || x.accountId === a.id);
  const sched = props.scheduled ?? [];
  const searchList = (searches.length > 0 || sched.length > 0) && (
    <>
      {sched.length > 0 && (
        <>
          <div className="folder-section">Naplánované na odoslanie</div>
          {sched.map((x) => (
            <div key={x.id} className="saved-search">
              <button className="folder" onClick={() => props.onScheduled?.(x, true)} title={`Upraviť – odošle sa ${new Date(x.sendAt).toLocaleString('sk-SK')}`}>
                <span className="grow">{x.subject || '(bez predmetu)'}</span>
                <span className="muted small">{new Date(x.sendAt).toLocaleString('sk-SK', { day: 'numeric', month: 'numeric', hour: '2-digit', minute: '2-digit' })}</span>
              </button>
              <button className="saved-search-x" aria-label="Zrušiť odoslanie" title="Zrušiť odoslanie" onClick={() => props.onScheduled?.(x, false)}>×</button>
            </div>
          ))}
        </>
      )}
      {searches.length > 0 && <div className="folder-section">Uložené hľadania</div>}
      {searches.map((x) => (
        <div key={x.id} className="saved-search">
          <button className={`folder ${props.activeSearch === x.query ? 'active' : ''}`} onClick={() => props.onSavedSearch?.(x)} title={x.query}>
            <SearchSmallIcon /> <span className="grow">{x.name}</span>
          </button>
          <button className="saved-search-x" aria-label={`Odstrániť hľadanie ${x.name}`} title="Odstrániť" onClick={() => props.onRemoveSearch?.(x)}>×</button>
        </div>
      ))}
    </>
  );
  const dropProps = (accountId: string, path: string) => ({
    onDragOver: (e: React.DragEvent) => {
      if (!e.dataTransfer.types.includes('application/x-mailklient')) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      setOver(`${accountId}|${path}`);
    },
    onDragLeave: () => setOver((o) => (o === `${accountId}|${path}` ? null : o)),
    onDrop: (e: React.DragEvent) => {
      setOver(null);
      try {
        const d = JSON.parse(e.dataTransfer.getData('application/x-mailklient')) as DragMessages;
        e.preventDefault();
        if (d.accountId === accountId && d.folder !== path) props.onDropMessages?.(path, d);
      } catch {
        /* nič */
      }
    },
  });
  if (!a)
    return (
      <aside className="accounts">
        <div className="pane-head">
          <span className="dot" style={{ background: 'var(--colorBrandBackground)' }} />
          <span className="pane-head-text"><b>Všetky kontá</b><span className="muted small">{props.accounts.length} IMAP kontá</span></span>
        </div>
        <button className={`folder ${props.selection.accountId === 'all' ? 'active' : ''}`} onClick={() => props.onSelect({ accountId: 'all', folder: 'INBOX' })}>
          <InboxIcon /> <span className="grow">Všetky doručené</span>
        </button>
        {props.accounts.map((x) => (
          <button key={x.id} className={`folder ${props.selection.accountId === x.id ? 'active' : ''}`} onClick={() => props.onSelect({ accountId: x.id, folder: 'INBOX' })}>
            <span className="dot" style={{ background: accountColor(x.color) }} /> <span className="grow">{x.name}</span>
            {(props.unread[x.id] ?? 0) > 0 && <span className="count">{props.unread[x.id]}</span>}
          </button>
        ))}
        {searchList}
      </aside>
    );
  const f = props.folders[a.id];
  return (
    <aside className="accounts">
      <div className="pane-head">
        <span className="dot" style={{ background: accountColor(a.color) }} />
        <span className="pane-head-text"><b>{a.name}</b><span className="muted small">{a.email}</span></span>
      </div>
      {f === 'loading' || f === undefined ? (
        <Spinner size="extra-tiny" className="pad" />
      ) : 'error' in f ? (
        <div className="pane-error" title={f.error}>⚠ {f.error}</div>
      ) : (
        f.map((fo) => (
          <button
            key={fo.path}
            className={`folder ${props.selection.accountId === a.id && props.selection.folder === fo.path ? 'active' : ''} ${over === `${a.id}|${fo.path}` ? 'drop-over' : ''}`}
            style={{ paddingLeft: 14 + fo.depth * 14 }}
            {...dropProps(a.id, fo.path)}
            onClick={() => props.onSelect({ accountId: a.id, folder: fo.path })}
            onContextMenu={(e) => {
              if (!props.onFolderContext) return;
              e.preventDefault();
              props.onFolderContext(a.id, fo, e.clientX, e.clientY);
            }}
          >
            {folderIcon(fo)} <span className="grow">{fo.name}</span>
            {fo.specialUse === '\\Inbox' && (props.unread[a.id] ?? 0) > 0 && <span className="count">{props.unread[a.id]}</span>}
          </button>
        ))
      )}
      {searchList}
      {props.quota && props.quota.limit > 0 && (() => {
        const p = props.quota.used / props.quota.limit;
        const gb = (b: number) => (b >= 1e9 ? `${(b / 1e9).toFixed(1)} GB` : `${Math.round(b / 1e6)} MB`);
        return (
          <div className={`quota ${p > 0.9 ? 'warn' : ''}`} title="Obsadenosť schránky na serveri">
            Schránka: {gb(props.quota.used)} z {gb(props.quota.limit)} ({Math.round(p * 100)} %){p > 0.9 ? ' – takmer plná, odstráňte staré správy alebo prílohy' : ''}
            <div className="quota-bar"><span style={{ width: `${Math.min(100, p * 100)}%` }} /></div>
          </div>
        );
      })()}
    </aside>
  );
}

function fmtDate(iso: string | null) {
  if (!iso) return '';
  const d = new Date(iso);
  const now = new Date();
  if (d.toDateString() === now.toDateString()) return d.toLocaleTimeString('sk-SK', { hour: '2-digit', minute: '2-digit' });
  return d.toLocaleDateString('sk-SK', { day: 'numeric', month: 'numeric', year: d.getFullYear() === now.getFullYear() ? undefined : '2-digit' });
}

const normSubject = (s: string) => s.replace(/^((re|fw|fwd|pd|odp|aw|wg)\s*(\[\d+\])?\s*:\s*)+/i, '').trim().toLowerCase();

/** Zoskupenie do konverzácií: Message-ID / In-Reply-To, inak rovnaký predmet bez RE:/PD: v rámci konta. */
export function groupThreads(list: MessageSummary[]): MessageSummary[][] {
  const parent = new Map<string, string>();
  const find = (x: string): string => {
    let r = x;
    while (parent.get(r) && parent.get(r) !== r) r = parent.get(r)!;
    parent.set(x, r);
    return r;
  };
  const union = (a: string, b: string) => {
    const ra = find(a);
    const rb = find(b);
    if (ra !== rb) parent.set(ra, rb);
  };
  const key = (m: MessageSummary) => `${m.accountId}|${m.folder}|${m.uid}`;
  for (const m of list) {
    const k = key(m);
    if (!parent.has(k)) parent.set(k, k);
    if (m.messageId) union(k, `id:${m.accountId}:${m.messageId}`);
    if (m.inReplyTo) union(k, `id:${m.accountId}:${m.inReplyTo}`);
    const ns = normSubject(m.subject);
    if (ns.length > 3) union(k, `subj:${m.accountId}:${ns}`);
  }
  const groups = new Map<string, MessageSummary[]>();
  for (const m of list) {
    const r = find(key(m));
    groups.set(r, [...(groups.get(r) ?? []), m]);
  }
  return [...groups.values()];
}

export interface DragMessages {
  accountId: string;
  folder: string;
  uids: number[];
}

export type SortBy = 'date' | 'from' | 'subject' | 'flag' | 'unread' | 'category';
const SORTS: Array<{ id: SortBy; label: string }> = [
  { id: 'date', label: 'Dátum' },
  { id: 'from', label: 'Od' },
  { id: 'subject', label: 'Predmet' },
  { id: 'category', label: 'Kategória' },
  { id: 'flag', label: 'Vlajka (dôležité)' },
  { id: 'unread', label: 'Neprečítané' },
];

export function MessageList(props: {
  title: string;
  messages: MessageSummary[];
  loading: boolean;
  error?: string;
  selected?: MessageSummary;
  colorOf: (accountId: string) => string;
  onOpen: (m: MessageSummary) => void;
  onDeleteKey?: () => void;
  /** pravý klik: označené správy (alebo tá pod kurzorom) */
  onContext?: (msgs: MessageSummary[], x: number, y: number) => void;
  /** Delete pri viacerých označených */
  onDeleteMany?: (msgs: MessageSummary[]) => void;
  /** zmena výberu viacerých správ (počet) */
  onMulti?: (msgs: MessageSummary[]) => void;
  /** rýchle akcie priamo v riadku (pri prejdení myšou) */
  onQuick?: (m: MessageSummary, a: 'flag' | 'pin' | 'delete' | 'seen') => void;
  sort?: { by: SortBy; dir: 'desc' | 'asc' };
  onSort?: (s: { by: SortBy; dir: 'desc' | 'asc' }) => void;
  headerExtra?: React.ReactNode;
  categories?: Category[];
  conversations?: boolean;
  density?: 'comfortable' | 'compact';
  previewLine?: boolean;
  /** dvojklik (shift = opačný spôsob než v nastavení) */
  onOpenFull?: (m: MessageSummary, alt: boolean) => void;
  /** v priečinku sú staršie správy, ktoré ešte nie sú načítané */
  onLoadMore?: () => void;
  loadingMore?: boolean;
}) {
  const [tab, setTab] = useState<'all' | 'unread'>('all');
  const [open, setOpen] = useState<string | null>(null);
  const base = tab === 'unread' ? props.messages.filter((m) => !m.seen) : props.messages;
  const isPinned = (m: MessageSummary) => !!m.keywords?.includes(PIN_KEYWORD);
  // pripnuté hore, potom podľa dátumu
  const groups = useMemo(() => {
    const g = props.conversations ? groupThreads(base) : base.map((m) => [m]);
    for (const x of g) x.sort((a, b) => (b.date ?? '').localeCompare(a.date ?? ''));
    const by = props.sort?.by ?? 'date';
    const dir = props.sort?.dir === 'asc' ? -1 : 1;
    const cats = props.categories ?? [];
    const catName = (m: MessageSummary) => cats.find((c) => m.keywords?.includes(c.id))?.name ?? '\uffff';
    const cmp = (a: MessageSummary, b: MessageSummary): number => {
      const byDate = (b.date ?? '').localeCompare(a.date ?? '');
      switch (by) {
        case 'from': return (a.fromName || a.fromAddress).localeCompare(b.fromName || b.fromAddress, 'sk') * dir || byDate;
        case 'subject': return a.subject.replace(/^((re|fw|fwd|pd|odp):\s*)+/i, '').localeCompare(b.subject.replace(/^((re|fw|fwd|pd|odp):\s*)+/i, ''), 'sk') * dir || byDate;
        case 'flag': return (Number(b.flagged) - Number(a.flagged)) * dir || byDate;
        case 'unread': return (Number(!b.seen) - Number(!a.seen)) * dir || byDate;
        case 'category': return catName(a).localeCompare(catName(b), 'sk') * dir || byDate;
        default: return byDate * dir;
      }
    };
    return g.sort((a, b) => Number(b.some(isPinned)) - Number(a.some(isPinned)) || cmp(a[0]!, b[0]!));
  }, [base, props.conversations, props.sort, props.categories]); // eslint-disable-line react-hooks/exhaustive-deps
  const cats = props.categories ?? [];
  // výber viacerých správ: Ctrl+klik pridá/odoberie, Shift+klik vyberie rozsah
  const keyOf = (m: MessageSummary) => `${m.accountId}/${m.folder}/${m.uid}`;
  const [multi, setMulti] = useState<string[]>([]);
  const anchor = useRef<string | null>(null);
  const order = useMemo(() => groups.flatMap((g) => g), [groups]);
  const multiMsgs = useMemo(() => order.filter((m) => multi.includes(keyOf(m))), [order, multi]); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => props.onMulti?.(multiMsgs), [multiMsgs]); // eslint-disable-line react-hooks/exhaustive-deps
  // klávesnica: ↑/↓ = predchádzajúca/ďalšia správa, Shift+↑/↓ = rozšíriť výber (ako v Outlooku), Home/End
  const cursor = useRef<string | null>(null);
  const keyNav = (e: React.KeyboardEvent<HTMLDivElement>) => {
    if (!['ArrowUp', 'ArrowDown', 'Home', 'End'].includes(e.key) || e.altKey) return false;
    const rows = [...e.currentTarget.querySelectorAll<HTMLElement>('.msg[data-k]')];
    if (!rows.length) return false;
    const keys = rows.map((r) => r.dataset.k!);
    const curKey = (e.shiftKey && cursor.current) || (props.selected ? keyOf(props.selected) : null);
    const i = curKey ? keys.indexOf(curKey) : -1;
    const j = e.key === 'Home' ? 0 : e.key === 'End' ? keys.length - 1 : Math.max(0, Math.min(keys.length - 1, i < 0 ? 0 : i + (e.key === 'ArrowDown' ? 1 : -1)));
    e.preventDefault();
    const target = order.find((x) => keyOf(x) === keys[j]);
    if (!target) return true;
    rows[j]!.focus({ preventScroll: true });
    rows[j]!.scrollIntoView({ block: 'nearest' });
    if (e.shiftKey) {
      const from = anchor.current && keys.includes(anchor.current) ? anchor.current : curKey ?? keys[j]!;
      if (!anchor.current || !keys.includes(anchor.current)) anchor.current = from;
      const a = keys.indexOf(from);
      setMulti(keys.slice(Math.min(a, j), Math.max(a, j) + 1));
      cursor.current = keys[j]!;
    } else {
      setMulti([]);
      anchor.current = keys[j]!;
      cursor.current = null;
      props.onOpen(target);
    }
    return true;
  };
  const clickRow = (m: MessageSummary, e: React.MouseEvent) => {
    const k = keyOf(m);
    if (e.ctrlKey || e.metaKey) {
      const base = multi.length ? multi : props.selected ? [keyOf(props.selected)] : [];
      setMulti(base.includes(k) ? base.filter((x) => x !== k) : [...base, k]);
      anchor.current = k;
      return true;
    }
    if (e.shiftKey) {
      const from = anchor.current ?? (props.selected ? keyOf(props.selected) : k);
      const i = order.findIndex((x) => keyOf(x) === from);
      const j = order.findIndex((x) => keyOf(x) === k);
      if (i >= 0 && j >= 0) {
        setMulti(order.slice(Math.min(i, j), Math.max(i, j) + 1).map(keyOf));
        cursor.current = k;
        return true;
      }
    }
    setMulti([]);
    anchor.current = k;
    cursor.current = null;
    return false;
  };
  const row = (m: MessageSummary, opts: { count?: number; unreadInThread?: boolean; child?: boolean; gk?: string } = {}) => {
    const sel = multi.length > 1 ? multi.includes(keyOf(m)) : props.selected && props.selected.uid === m.uid && props.selected.accountId === m.accountId && props.selected.folder === m.folder;
    const mc = cats.filter((c) => m.keywords?.includes(c.id));
    const unread = opts.unreadInThread ?? !m.seen;
    return (
      <button
        key={`${m.accountId}/${m.folder}/${m.uid}`}
        data-k={keyOf(m)}
        className={`msg ${sel ? 'selected' : ''} ${unread ? 'unread' : ''} ${m.flagged ? 'important' : ''} ${opts.child ? 'msg-child' : ''}`}
        draggable
        onDragStart={(e) => {
          const d: DragMessages = { accountId: m.accountId, folder: m.folder, uids: [m.uid] };
          e.dataTransfer.setData('application/x-mailklient', JSON.stringify(d));
          e.dataTransfer.effectAllowed = 'move';
        }}
        onClick={(e) => {
          if (clickRow(m, e)) return;
          if (opts.count && opts.count > 1 && opts.gk) setOpen((o) => (o === opts.gk ? o : opts.gk!));
          props.onOpen(m);
        }}
        onContextMenu={(e) => {
          if (!props.onContext) return;
          e.preventDefault();
          const k = keyOf(m);
          if (multi.length > 1 && multi.includes(k)) return props.onContext(multiMsgs, e.clientX, e.clientY);
          setMulti([]);
          props.onContext([m], e.clientX, e.clientY);
        }}
        onDoubleClick={(e) => props.onOpenFull?.(m, e.shiftKey)}
        onAuxClick={(e) => e.button === 1 && props.onOpenFull?.(m, true)}
      >
        {!opts.child && <Avatar name={m.fromName || m.fromAddress} color={props.colorOf(m.accountId)} />}
        <div className="msg-main">
          <div className="msg-row">
            <span className="msg-from">
              {opts.count && opts.count > 1 && (
                <span
                  className="thread-toggle"
                  role="button"
                  aria-label={open === opts.gk ? 'Zbaliť konverzáciu' : 'Rozbaliť konverzáciu'}
                  onClick={(e) => {
                    e.stopPropagation();
                    setOpen((o) => (o === opts.gk ? null : opts.gk!));
                  }}
                >
                  {open === opts.gk ? '▾' : '▸'}
                </span>
              )}
              {m.fromName || m.fromAddress}
              {opts.count && opts.count > 1 ? <span className="thread-count">{opts.count}</span> : null}
            </span>
            {m.hasAttachments && <AttachTinyIcon className="msg-att" aria-label="Má prílohu" title="Správa má prílohu" />}
            <span className="msg-date">{fmtDate(m.date)}</span>
          </div>
          <div className="msg-row">
            <span className="msg-subject">{m.subject}</span>
            <span className="msg-icons">
              {props.onQuick && !opts.child ? (
                <>
                  <span role="button" tabIndex={-1} className="qa hover" title={m.seen ? 'Označiť ako neprečítané' : 'Označiť ako prečítané'} aria-label={m.seen ? 'Označiť ako neprečítané' : 'Označiť ako prečítané'} onClick={(e) => (e.stopPropagation(), props.onQuick!(m, 'seen'))}>
                    {m.seen ? <UnreadSmallIcon /> : <ReadSmallIcon />}
                  </span>
                  <span role="button" tabIndex={-1} className="qa hover" title="Odstrániť" aria-label="Odstrániť" onClick={(e) => (e.stopPropagation(), props.onQuick!(m, 'delete'))}>
                    <DeleteSmallIcon />
                  </span>
                  <span role="button" tabIndex={-1} className={`qa ${isPinned(m) ? 'on pin' : 'hover'}`} title={isPinned(m) ? 'Odopnúť' : 'Pripnúť navrch'} aria-label={isPinned(m) ? 'Odopnúť' : 'Pripnúť navrch'} aria-pressed={isPinned(m)} onClick={(e) => (e.stopPropagation(), props.onQuick!(m, 'pin'))}>
                    {isPinned(m) ? <PinSmallIcon /> : <PinOutlineSmallIcon />}
                  </span>
                  <span role="button" tabIndex={-1} className={`qa ${m.flagged ? 'on flag' : 'hover'}`} title={m.flagged ? 'Zrušiť vlajku' : 'Označiť vlajkou'} aria-label={m.flagged ? 'Zrušiť vlajku' : 'Označiť vlajkou'} aria-pressed={m.flagged} onClick={(e) => (e.stopPropagation(), props.onQuick!(m, 'flag'))}>
                    {m.flagged ? <FlagFilledIcon /> : <FlagSmallIcon />}
                  </span>
                </>
              ) : (
                <>
                  {isPinned(m) && <PinSmallIcon className="pin" aria-label="Pripnuté" />}
                  {m.flagged && <FlagFilledIcon className="flag" />}
                </>
              )}
            </span>
          </div>
          {props.previewLine && m.preview && !opts.child && <div className="msg-preview">{m.preview}</div>}
          {mc.length > 0 && (
            <div className="msg-cats">
              {mc.map((c) => (
                <span key={c.id} className="cat-chip" style={{ background: `${c.color}22`, color: c.color, borderColor: `${c.color}55` }}>
                  {c.name}
                </span>
              ))}
            </div>
          )}
        </div>
      </button>
    );
  };
  return (
    <section className={`list ${props.density === 'compact' ? 'compact' : ''}`}>
      <div className="list-head">
        <div className="list-title">{props.title}</div>
        {props.headerExtra}
        <div className="list-tabsrow">
        <TabList size="small" selectedValue={tab} onTabSelect={(_, d) => setTab(d.value as 'all' | 'unread')}>
          <Tab value="all">Všetky</Tab>
          <Tab value="unread">Neprečítané</Tab>
        </TabList>
        {props.onSort && (
          <Menu
            checkedValues={{ by: [props.sort?.by ?? 'date'], dir: [props.sort?.dir ?? 'desc'] }}
            onCheckedValueChange={(_, d) => {
              const cur = props.sort ?? { by: 'date' as SortBy, dir: 'desc' as const };
              if (d.name === 'by' && d.checkedItems[0]) props.onSort!({ ...cur, by: d.checkedItems[0] as SortBy });
              if (d.name === 'dir' && d.checkedItems[0]) props.onSort!({ ...cur, dir: d.checkedItems[0] as 'desc' | 'asc' });
            }}
          >
            <MenuTrigger disableButtonEnhancement>
              <button className="sort-btn" title="Zoradiť">
                <SortIcon /> Podľa: {SORTS.find((x) => x.id === (props.sort?.by ?? 'date'))?.label} {props.sort?.dir === 'asc' ? '↑' : '↓'}
              </button>
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                <MenuGroup>
                  <MenuGroupHeader>Zoradiť podľa</MenuGroupHeader>
                  {SORTS.map((x) => (
                    <MenuItemRadio key={x.id} name="by" value={x.id}>{x.label}</MenuItemRadio>
                  ))}
                </MenuGroup>
                <MenuDivider />
                <MenuGroup>
                  <MenuGroupHeader>Spôsob zoradenia</MenuGroupHeader>
                  <MenuItemRadio name="dir" value="asc">{(props.sort?.by ?? 'date') === 'date' ? 'Najstaršie na začiatku' : 'Vzostupne (A → Z)'}</MenuItemRadio>
                  <MenuItemRadio name="dir" value="desc">{(props.sort?.by ?? 'date') === 'date' ? 'Najnovšie na začiatku' : 'Zostupne (Z → A)'}</MenuItemRadio>
                </MenuGroup>
              </MenuList>
            </MenuPopover>
          </Menu>
        )}
        </div>
      </div>
      <div
        className="list-body"
        tabIndex={-1}
        onKeyDown={(e) => {
          if (keyNav(e)) return;
          if (e.key === 'Delete') {
            // Delete: vybrané správy (aj viac naraz) do Odstránených a hneď vybrať ďalšiu, ako v Outlooku
            e.preventDefault();
            e.stopPropagation(); // nespúšťať aj globálnu skratku (inak by sa mazalo dvakrát)
            const del = multiMsgs.length > 1 ? multiMsgs : props.selected ? [props.selected] : [];
            if (!del.length) return;
            const list = e.currentTarget;
            const keys = [...list.querySelectorAll<HTMLElement>('.msg[data-k]')].map((r) => r.dataset.k!);
            const dk = new Set(del.map(keyOf));
            const idx = keys.flatMap((k, i) => (dk.has(k) ? [i] : []));
            const nextKey = idx.length ? keys.slice(Math.min(...idx) + 1).find((k) => !dk.has(k)) ?? keys.slice(0, Math.min(...idx)).reverse().find((k) => !dk.has(k)) : undefined;
            const next = nextKey ? order.find((x) => keyOf(x) === nextKey) : undefined;
            if (del.length > 1) props.onDeleteMany?.(del);
            else props.onDeleteKey?.();
            setMulti([]);
            cursor.current = null;
            if (next && nextKey) {
              anchor.current = nextKey;
              props.onOpen(next);
              window.setTimeout(() => list.querySelector<HTMLElement>(`.msg[data-k="${CSS.escape(nextKey)}"]`)?.focus({ preventScroll: true }), 50);
            }
          }
          if (e.key === 'Escape') setMulti([]);
          if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'a') {
            e.preventDefault();
            setMulti(order.map(keyOf));
          }
        }}
        onScroll={(e) => {
          // na konci zoznamu automaticky načítať staršie správy
          const el = e.currentTarget;
          if (props.onLoadMore && !props.loadingMore && el.scrollHeight - el.scrollTop - el.clientHeight < 120) props.onLoadMore();
        }}
      >
        {props.loading && <div className="progress" role="progressbar" aria-label="Načítavam" />}
        {props.error && <div className="pane-error pad">⚠ {props.error}</div>}
        {props.loading && base.length === 0 && <div className="empty">Načítavam…</div>}
        {!props.loading && !props.error && base.length === 0 && <div className="empty">Žiadne správy</div>}
        {groups.map((g, i) => {
          const head = g[0]!;
          const gk = `${head.accountId}/${head.folder}/${head.uid}`;
          const pinned = g.some(isPinned);
          // nadpisy sekcií: Pripnuté hore, potom ostatné
          const section =
            pinned && i === 0 ? <div key="sec-pin" className="list-section">📌 Pripnuté</div>
            : !pinned && i > 0 && groups[i - 1]!.some(isPinned) ? <div key="sec-rest" className="list-section">Ostatné</div>
            : null;
          const body =
            g.length === 1 ? row(head) : (
              <div key={`g-${gk}`} className="thread">
                {row(head, { count: g.length, unreadInThread: g.some((m) => !m.seen), gk })}
                {open === gk && g.slice(1).map((m) => row(m, { child: true }))}
              </div>
            );
          return section ? [section, body] : body;
        })}
        {props.onLoadMore && base.length > 0 && (
          <div className="load-more">
            <Button size="small" appearance="subtle" disabled={props.loadingMore} onClick={props.onLoadMore}>
              {props.loadingMore ? 'Načítavam staršie správy…' : 'Načítať staršie správy'}
            </Button>
          </div>
        )}
      </div>
    </section>
  );
}

function InviteCard(props: { invite: NonNullable<MessageFull['invite']>; onRespond?: (p: 'ACCEPTED' | 'TENTATIVE' | 'DECLINED') => void }) {
  const { method, event: e } = props.invite;
  const [busy, setBusy] = useState(false);
  const when = e.allDay
    ? new Date(e.start).toLocaleDateString('sk-SK', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })
    : `${new Date(e.start).toLocaleString('sk-SK', { weekday: 'long', day: 'numeric', month: 'long', hour: '2-digit', minute: '2-digit' })} – ${new Date(e.end).toLocaleTimeString('sk-SK', { hour: '2-digit', minute: '2-digit' })}`;
  const head =
    method === 'CANCEL' ? 'Udalosť bola zrušená' : method === 'REPLY' ? 'Odpoveď na pozvánku' : method === 'REQUEST' ? 'Pozvánka na udalosť' : 'Udalosť';
  const reply = method === 'REPLY' ? e.attendees?.[0] : undefined;
  return (
    <div className={`invite ${method === 'CANCEL' ? 'cancel' : ''}`}>
      <div className="invite-head">{head}</div>
      <div className="invite-title">{e.summary}</div>
      <div className="muted">{when}{e.rrule ? ' · opakuje sa' : ''}</div>
      {e.location && <div className="muted">{e.location}</div>}
      {e.organizer && <div className="muted small">Organizátor: {e.organizer.name || e.organizer.email}</div>}
      {reply && (
        <div className="small">
          {reply.name || reply.email}:{' '}
          <b>{reply.partstat === 'ACCEPTED' ? 'prijal(a)' : reply.partstat === 'DECLINED' ? 'odmietol(a)' : reply.partstat === 'TENTATIVE' ? 'predbežne prijal(a)' : reply.partstat}</b>
        </div>
      )}
      {method === 'REQUEST' && props.onRespond && (
        <div className="row gap" style={{ marginTop: 8 }}>
          {(['ACCEPTED', 'TENTATIVE', 'DECLINED'] as const).map((p) => (
            <Button
              key={p}
              size="small"
              appearance={p === 'ACCEPTED' ? 'primary' : 'secondary'}
              disabled={busy}
              onClick={async () => {
                setBusy(true);
                await props.onRespond!(p);
                setBusy(false);
              }}
            >
              {p === 'ACCEPTED' ? 'Prijať' : p === 'TENTATIVE' ? 'Predbežne' : 'Odmietnuť'}
            </Button>
          ))}
        </div>
      )}
    </div>
  );
}

/** HTML e-mailu: sanitizované (DOMPurify) a v izolovanom iframe bez skriptov. Vzdialené obrázky blokované. */
function HtmlBody({ html, showImages }: { html: string; showImages: boolean }) {
  const doc = useMemo(() => {
    // WHOLE_DOCUMENT zachová <style> v hlavičke (rozloženie newsletterov)
    const clean = DOMPurify.sanitize(html, { WHOLE_DOCUMENT: true });
    const csp = showImages
      ? "default-src 'none'; img-src data: cid: https: http:; style-src 'unsafe-inline'; font-src data: https:"
      : "default-src 'none'; img-src data: cid:; style-src 'unsafe-inline'";
    const head = `<meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="${csp}"><base target="_blank"><style>body{font-family:Segoe UI,system-ui,sans-serif;font-size:14px;color:#242424;margin:0;padding:0 4px;word-wrap:break-word}img{max-width:100%;height:auto}</style>`;
    return /<head[^>]*>/i.test(clean)
      ? `<!doctype html>${clean.replace(/<head[^>]*>/i, (m) => m + head)}`
      : `<!doctype html><html><head>${head}</head><body>${clean}</body></html>`;
  }, [html, showImages]);
  return <iframe className="mail-frame" sandbox="allow-popups allow-popups-to-escape-sandbox" srcDoc={doc} title="Obsah správy" />;
}

export function ReadingPane(props: {
  message?: MessageFull;
  loading: boolean;
  error?: string;
  color: string;
  onReply: () => void;
  onReplyAll: () => void;
  onForward: () => void;
  thread?: MessageSummary[];
  onOpenThread?: (m: MessageSummary) => void;
  categories?: Category[];
  onRemoveCategory?: (id: string) => void;
  onRespond?: (partstat: 'ACCEPTED' | 'TENTATIVE' | 'DECLINED') => void;
  onBlockSender?: () => void;
  onPrint?: (pdf: boolean) => void;
  onUnsubscribe?: () => void;
  onPopOut?: () => void;
  onAttachmentContext?: (index: number, x: number, y: number) => void;
  onDelete?: () => void;
  onArchive?: () => void;
  onToggleSeen?: () => void;
  onToggleFlag?: () => void;
  onTogglePin?: () => void;
  /** ďalšie akcie (⋯) – menu na súradniciach */
  onMore?: (x: number, y: number) => void;
}) {
  const [preview, setPreview] = useState<{ src: string; name: string } | null>(null);
  const [showImages, setShowImages] = useState(false);
  const [showInline, setShowInline] = useState(false);
  const [attErr, setAttErr] = useState<string | null>(null);
  const { guard, dialog: attGuard } = useAttachmentGuard();
  const m = props.message;
  const api = window.mailApi;
  // Pri načítaní ostáva viditeľná predchádzajúca správa (žiadne biele bliknutie), hore len pásik priebehu.
  if (props.loading && !props.message) return <section className="reading"><div className="progress" role="progressbar" aria-label="Otváram správu" /><div className="empty">Otváram správu…</div></section>;
  if (props.error) return <section className="reading"><div className="pane-error pad">⚠ {props.error}</div></section>;
  if (!m) return <section className="reading"><div className="empty">Vyber správu na čítanie</div></section>;
  const demo = m.accountId === 'demo';
  return (
    <section className="reading">
      {props.loading && <div className="progress" role="progressbar" aria-label="Otváram správu" />}
      <div className="read-actions" role="toolbar" aria-label="Akcie so správou">
        <Button size="small" icon={<ReplyIcon />} onClick={props.onReply}>Odpovedať</Button>
        <Button size="small" icon={<ReplyAllIcon />} onClick={props.onReplyAll} aria-label="Odpovedať všetkým" title="Odpovedať všetkým" />
        <Button size="small" icon={<ForwardIcon />} onClick={props.onForward} aria-label="Preposlať" title="Preposlať" />
        <span className="ra-sep" />
        {props.onDelete && <Button size="small" appearance="subtle" icon={<DeleteIcon />} onClick={props.onDelete} title="Odstrániť (Delete)" aria-label="Odstrániť" />}
        {props.onArchive && <Button size="small" appearance="subtle" icon={<ArchiveIcon />} onClick={props.onArchive} title="Archivovať" aria-label="Archivovať" />}
        {props.onToggleSeen && (
          <Button size="small" appearance="subtle" icon={m.seen ? <MarkUnreadIcon /> : <MarkReadIcon />} onClick={props.onToggleSeen} title={m.seen ? 'Označiť ako neprečítané (Ctrl+U)' : 'Označiť ako prečítané (Ctrl+Q)'} aria-label={m.seen ? 'Označiť ako neprečítané' : 'Označiť ako prečítané'} />
        )}
        {props.onToggleFlag && (
          <Button size="small" appearance="subtle" className={m.flagged ? 'ra-on flag' : ''} icon={m.flagged ? <FlagFilledIcon /> : <FlagIcon />} onClick={props.onToggleFlag} title={m.flagged ? 'Zrušiť vlajku' : 'Označiť vlajkou (Insert)'} aria-label={m.flagged ? 'Zrušiť vlajku' : 'Označiť vlajkou'} aria-pressed={m.flagged} />
        )}
        {props.onTogglePin && (
          <Button size="small" appearance="subtle" className={!!m.keywords?.includes(PIN_KEYWORD) ? 'ra-on pin' : ''} icon={<PinIcon />} onClick={props.onTogglePin} title={!!m.keywords?.includes(PIN_KEYWORD) ? 'Odopnúť' : 'Pripnúť navrch'} aria-label={!!m.keywords?.includes(PIN_KEYWORD) ? 'Odopnúť' : 'Pripnúť navrch'} aria-pressed={!!m.keywords?.includes(PIN_KEYWORD)} />
        )}
        <span className="ra-sep" />
        {props.onPrint && <Button size="small" appearance="subtle" icon={<PrintIcon />} onClick={() => props.onPrint!(false)} title="Tlačiť (Ctrl+P)" aria-label="Tlačiť" />}
        {props.onPrint && <Button size="small" appearance="subtle" icon={<PdfIcon />} onClick={() => props.onPrint!(true)} title="Uložiť ako PDF" aria-label="Uložiť ako PDF" />}
        {props.onPopOut && <Button size="small" appearance="subtle" icon={<NewWindowSmallIcon />} onClick={props.onPopOut} title="Otvoriť v samostatnom okne" aria-label="Otvoriť v samostatnom okne" />}
        {props.onMore && <Button size="small" appearance="subtle" icon={<MoreIcon />} onClick={(e) => { const r = (e.currentTarget as HTMLElement).getBoundingClientRect(); props.onMore!(r.left, r.bottom); }} title="Ďalšie akcie" aria-label="Ďalšie akcie" />}
        {!props.onMore && props.onBlockSender && <Button size="small" appearance="subtle" onClick={props.onBlockSender} title="Pridať odosielateľa medzi blokovaných v Zimbre">Blokovať</Button>}
      </div>
      <h2 className="read-subject">{m.subject}</h2>
      <div className="read-head">
        <Avatar name={m.fromName || m.fromAddress} color={props.color} size={40} />
        <div className="read-meta">
          <div>
            <b>{m.fromName}</b> <span className="muted">&lt;{m.fromAddress}&gt;</span>
            {m.auth?.verified && !m.auth.warnings.length && (
              <span className="verified" title={`Server overil odosielateľa (SPF: ${m.auth.spf ?? '–'}, DKIM: ${m.auth.dkim ?? '–'}, DMARC: ${m.auth.dmarc ?? '–'})`}>
                <VerifiedIcon /> overený
              </span>
            )}
          </div>
          <Recipients label="Komu" value={m.to} />
          {m.cc && <Recipients label="Kópia" value={m.cc} />}
          <div className="muted read-date">{m.date ? new Date(m.date).toLocaleString('sk-SK') : ''}</div>
        </div>
      </div>
      {(() => {
        const mc = (props.categories ?? []).filter((c) => m.keywords?.includes(c.id));
        return mc.length ? (
          <div className="read-cats">
            {mc.map((c) => (
              <span key={c.id} className="cat-chip" style={{ background: `${c.color}22`, color: c.color, borderColor: `${c.color}55` }}>
                {c.name}
                {props.onRemoveCategory && (
                  <button className="cat-x" aria-label={`Odobrať kategóriu ${c.name}`} onClick={() => props.onRemoveCategory!(c.id)}>×</button>
                )}
              </span>
            ))}
          </div>
        ) : null;
      })()}
      {props.onUnsubscribe && (
        <div className="unsub">
          Hromadná správa (newsletter).
          <Button size="small" appearance="transparent" onClick={props.onUnsubscribe}>Odhlásiť odber</Button>
        </div>
      )}
      {m.security && (
        <div className={`secbar ${m.security.decryptError || m.security.signatureValid === false ? 'bad' : 'good'}`}>
          {m.security.encrypted && (
            <span><LockSmallIcon /> {m.security.decryptError ? m.security.decryptError : `Zašifrovaná správa (${m.security.encrypted === 'pgp' ? 'PGP' : 'S/MIME'})`}</span>
          )}
          {m.security.signed && (
            <span>
              {m.security.signatureValid
                ? `✓ Podpísal(a) ${m.security.signer ?? 'odosielateľ'}${m.security.trusted === false ? ' – certifikát nie je od dôveryhodnej autority' : ''}`
                : `⚠ Podpis neplatný${m.security.signatureNote ? `: ${m.security.signatureNote}` : ''}`}
            </span>
          )}
          {m.security.signatureValid && m.security.signatureNote && <span className="muted">{m.security.signatureNote}</span>}
        </div>
      )}
      {m.auth && m.auth.warnings.length > 0 && (
        <div className="phish">
          <div className="phish-head"><WarnIcon /> Pozor, správa môže byť podvodná</div>
          <ul>{m.auth.warnings.map((w, i) => <li key={i}>{w}</li>)}</ul>
          <div className="small">Neklikajte na odkazy a neposielajte heslá ani platobné údaje, kým si odosielateľa neoveríte inou cestou.</div>
        </div>
      )}
      {m.invite && <InviteCard invite={m.invite} onRespond={props.onRespond} />}
      {props.thread && props.thread.length > 1 && (
        <details className="thread-box">
          <summary>Konverzácia: {props.thread.length} správ</summary>
          {props.thread.map((t) => {
            const cur = t.uid === m.uid && t.folder === m.folder;
            return (
              <button key={`${t.folder}/${t.uid}`} className={`thread-item ${cur ? 'current' : ''}`} disabled={cur} onClick={() => props.onOpenThread?.(t)}>
                <span className="grow">{t.fromName || t.fromAddress}</span>
                <span className="muted small">{t.folder === 'INBOX' ? 'Doručená pošta' : t.folder}</span>
                <span className="muted small">{fmtDate(t.date)}</span>
              </button>
            );
          })}
        </details>
      )}
      {(() => {
        const list = m.attachments.map((a, pos) => ({ a, pos, i: a.index ?? pos }));
        const real = list.filter((x) => !x.a.inline);
        const inl = list.filter((x) => x.a.inline);
        const chip = ({ a, pos, i }: (typeof list)[number]) => (
            <span
              key={pos}
              className={`attachment ${a.inline ? 'inline-img' : ''}`}
              title={a.inline ? 'Obrázok z textu správy (napr. logo v podpise)' : undefined}
              onContextMenu={(e) => {
                if (!props.onAttachmentContext || demo) return;
                e.preventDefault();
                props.onAttachmentContext(i, e.clientX, e.clientY);
              }}
            >
              <span className="attachment-name" title={`${a.filename} – otvorí sa tlačidlom Náhľad`}>
                <AttachIcon /> <span className="attachment-fn">{a.filename}</span> <span className="muted">({Math.ceil(a.size / 1024)} kB)</span>
              </span>
              <Button
                size="small"
                appearance="transparent"
                icon={<PreviewSmallIcon />}
                disabled={demo}
                title="Náhľad: obrázky priamo tu, PDF v okne náhľadu, Office v nainštalovanom programe"
                onClick={() =>
                  guard(
                    a.filename,
                    async () => {
                      const r = await api?.previewAttachment(m.accountId, m.folder, m.uid, i);
                      if (r?.ok && r.data.kind === 'image' && r.data.src) setPreview({ src: r.data.src, name: r.data.filename });
                      else if (r && !r.ok) setAttErr(r.error);
                    },
                    () => void api?.saveAttachment(m.accountId, m.folder, m.uid, i),
                  )
                }
              >
                Náhľad
              </Button>
              <Button size="small" appearance="transparent" icon={<SaveSmallIcon />} disabled={demo} title="Uložiť na disk (nič sa nespustí)" onClick={() => void api?.saveAttachment(m.accountId, m.folder, m.uid, i)}>
                Uložiť
              </Button>
            </span>
        );
        const word = inl.length === 1 ? 'obrázok' : inl.length < 5 ? 'obrázky' : 'obrázkov';
        return (
          <>
            {real.length > 0 && <div className={`attachments ${real.length > 4 ? 'many' : ''}`} aria-label="Prílohy">{real.map(chip)}</div>}
            {inl.length > 0 && (
              <div className={`inline-imgs ${showInline ? 'open' : ''}`}>
                <button className="link inline-toggle" onClick={() => setShowInline(!showInline)} aria-expanded={showInline}>
                  {showInline ? '▾' : '▸'} {inl.length} {word} z textu správy (logo v podpise a pod.) – nie sú to prílohy
                </button>
                {showInline && <div className="inline-list">{inl.map(chip)}</div>}
              </div>
            )}
          </>
        );
      })()}
      {m.html && !showImages && (
        <div className="images-bar">
          Vzdialené obrázky sú zablokované kvôli súkromiu. <Button size="small" appearance="transparent" onClick={() => setShowImages(true)}>Zobraziť obrázky</Button>
        </div>
      )}
      {attGuard}
      {attErr && (
        <div className="att-err" role="alert">
          ⛔ {attErr} <Button size="small" appearance="transparent" onClick={() => setAttErr(null)}>Zavrieť</Button>
        </div>
      )}
      {preview && (
        <div className="att-preview" onClick={() => setPreview(null)} role="dialog" aria-label={`Náhľad ${preview.name}`}>
          <div className="att-preview-bar">{preview.name} <span className="muted">– kliknutím zavrieť</span></div>
          <img src={preview.src} alt={preview.name} />
        </div>
      )}
      <div className="read-body">{m.html ? <HtmlBody html={m.html} showImages={showImages} /> : <pre className="plain">{m.text}</pre>}</div>
    </section>
  );
}

/** „Meno" <adresa>, adresa, … → zoznam (čiarky v úvodzovkách sa nedelia) */
export function splitAddresses(s: string): { name: string; email: string }[] {
  const out: { name: string; email: string }[] = [];
  let cur = '';
  let q = false;
  let angle = false;
  for (const ch of s + ',') {
    if (ch === '"') q = !q;
    else if (ch === '<' && !q) angle = true;
    else if (ch === '>' && !q) angle = false;
    if ((ch === ',' || ch === ';') && !q && !angle) {
      const t = cur.trim();
      cur = '';
      if (!t) continue;
      const m = /^(.*?)\s*<([^>]+)>$/.exec(t);
      const name = (m ? m[1]! : '').replace(/^"|"$/g, '').trim();
      const email = (m ? m[2]! : t).trim();
      out.push({ name: name && name !== email ? name : '', email });
    } else cur += ch;
  }
  return out;
}

/** Príjemcovia v hlavičke otvorenej správy: pri veľa adresátoch (hromadné pozvánky na Teams a pod.)
    len prví traja menom + „a ďalších N“, po rozbalení všetci v obmedzenej výške s posuvníkom. */
function Recipients(props: { label: string; value: string }) {
  const [open, setOpen] = useState(false);
  const list = useMemo(() => splitAddresses(props.value), [props.value]);
  const LIMIT = 3;
  if (!list.length) return null;
  const many = list.length > LIMIT + 2; // pri 4–5 adresátoch sa ukážu všetci
  const shown = open || !many ? list : list.slice(0, LIMIT);
  const rest = list.length - LIMIT;
  return (
    <div className={`muted read-rcpt ${open && many ? 'open' : ''}`}>
      <span className="read-rcpt-label">{props.label}:</span>{' '}
      {open && many && (
        <button className="link read-rcpt-more" onClick={() => setOpen(false)}>
          skryť ({list.length})
        </button>
      )}
      <span className="read-rcpt-list">
        {shown.map((a, i) => (
          <span key={i} className="read-rcpt-item" title={a.name ? `${a.name} <${a.email}>` : a.email}>
            {open && a.name ? `${a.name} <${a.email}>` : a.name || a.email}
            {i < shown.length - 1 ? '; ' : ''}
          </span>
        ))}
        {many && !open && (
          <button className="link read-rcpt-more" onClick={() => setOpen(true)} title="Zobraziť všetkých adresátov">
            {` a ${rest < 5 ? 'ďalší' : 'ďalších'} ${rest}`}
          </button>
        )}
      </span>
    </div>
  );
}
