// Panel kont, zoznam správ, čítací panel a 4. stĺpec (variant D).
import { Button, Spinner, Tab, TabList } from '@fluentui/react-components';
import DOMPurify from 'dompurify';
import { useMemo, useState } from 'react';
import type { ImapAccountPublic } from '../../core/account';
import type { FolderInfo, MessageFull, MessageSummary, OutboxItem } from '../../core/api';
import { PIN_KEYWORD, type Category, type SavedSearch } from '../../core/settings';
import { LockSmallIcon, VerifiedIcon, WarnIcon, PinSmallIcon, SearchSmallIcon, AttachIcon, FlagFilledIcon, FolderIcon, ForwardIcon, InboxIcon, ReplyAllIcon, ReplyIcon, SentIcon, TrashIcon } from '../icons';
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

export function MessageList(props: {
  title: string;
  messages: MessageSummary[];
  loading: boolean;
  error?: string;
  selected?: MessageSummary;
  colorOf: (accountId: string) => string;
  onOpen: (m: MessageSummary) => void;
  onDeleteKey?: () => void;
  headerExtra?: React.ReactNode;
  categories?: Category[];
  conversations?: boolean;
  density?: 'comfortable' | 'compact';
  previewLine?: boolean;
}) {
  const [tab, setTab] = useState<'all' | 'unread'>('all');
  const [open, setOpen] = useState<string | null>(null);
  const base = tab === 'unread' ? props.messages.filter((m) => !m.seen) : props.messages;
  const isPinned = (m: MessageSummary) => !!m.keywords?.includes(PIN_KEYWORD);
  // pripnuté hore, potom podľa dátumu
  const groups = useMemo(() => {
    const g = props.conversations ? groupThreads(base) : base.map((m) => [m]);
    for (const x of g) x.sort((a, b) => (b.date ?? '').localeCompare(a.date ?? ''));
    return g.sort((a, b) => Number(b.some(isPinned)) - Number(a.some(isPinned)) || (b[0]!.date ?? '').localeCompare(a[0]!.date ?? ''));
  }, [base, props.conversations]); // eslint-disable-line react-hooks/exhaustive-deps
  const cats = props.categories ?? [];
  const row = (m: MessageSummary, opts: { count?: number; unreadInThread?: boolean; child?: boolean; gk?: string } = {}) => {
    const sel = props.selected && props.selected.uid === m.uid && props.selected.accountId === m.accountId && props.selected.folder === m.folder;
    const mc = cats.filter((c) => m.keywords?.includes(c.id));
    const unread = opts.unreadInThread ?? !m.seen;
    return (
      <button
        key={`${m.accountId}/${m.folder}/${m.uid}`}
        className={`msg ${sel ? 'selected' : ''} ${unread ? 'unread' : ''} ${m.flagged ? 'important' : ''} ${opts.child ? 'msg-child' : ''}`}
        draggable
        onDragStart={(e) => {
          const d: DragMessages = { accountId: m.accountId, folder: m.folder, uids: [m.uid] };
          e.dataTransfer.setData('application/x-mailklient', JSON.stringify(d));
          e.dataTransfer.effectAllowed = 'move';
        }}
        onClick={() => {
          if (opts.count && opts.count > 1 && opts.gk) setOpen((o) => (o === opts.gk ? o : opts.gk!));
          props.onOpen(m);
        }}
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
            <span className="msg-date">{fmtDate(m.date)}</span>
          </div>
          <div className="msg-row">
            <span className="msg-subject">{m.subject}</span>
            <span className="msg-icons">
              {isPinned(m) && <PinSmallIcon className="pin" aria-label="Pripnuté" />}
              {m.flagged && <FlagFilledIcon className="flag" />}
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
        <TabList size="small" selectedValue={tab} onTabSelect={(_, d) => setTab(d.value as 'all' | 'unread')}>
          <Tab value="all">Všetky</Tab>
          <Tab value="unread">Neprečítané</Tab>
        </TabList>
      </div>
      <div
        className="list-body"
        tabIndex={-1}
        onKeyDown={(e) => {
          if (e.key === 'Delete' && props.onDeleteKey) props.onDeleteKey();
        }}
      >
        {props.loading && <div className="progress" role="progressbar" aria-label="Načítavam" />}
        {props.error && <div className="pane-error pad">⚠ {props.error}</div>}
        {props.loading && base.length === 0 && <div className="empty">Načítavam…</div>}
        {!props.loading && !props.error && base.length === 0 && <div className="empty">Žiadne správy</div>}
        {groups.map((g) => {
          const head = g[0]!;
          const gk = `${head.accountId}/${head.folder}/${head.uid}`;
          if (g.length === 1) return row(head);
          return (
            <div key={`g-${gk}`} className="thread">
              {row(head, { count: g.length, unreadInThread: g.some((m) => !m.seen), gk })}
              {open === gk && g.slice(1).map((m) => row(m, { child: true }))}
            </div>
          );
        })}
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
}) {
  const [preview, setPreview] = useState<{ src: string; name: string } | null>(null);
  const [showImages, setShowImages] = useState(false);
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
          <div className="muted">Komu: {m.to}{m.cc ? ` · Kópia: ${m.cc}` : ''}</div>
          <div className="muted">{m.date ? new Date(m.date).toLocaleString('sk-SK') : ''}</div>
        </div>
        <div className="read-actions">
          <Button size="small" icon={<ReplyIcon />} onClick={props.onReply}>Odpovedať</Button>
          <Button size="small" icon={<ReplyAllIcon />} onClick={props.onReplyAll} aria-label="Odpovedať všetkým" title="Odpovedať všetkým" />
          <Button size="small" icon={<ForwardIcon />} onClick={props.onForward} aria-label="Preposlať" title="Preposlať" />
          {props.onPrint && <Button size="small" appearance="subtle" onClick={() => props.onPrint!(false)} title="Tlačiť (Ctrl+P)">Tlačiť</Button>}
          {props.onPrint && <Button size="small" appearance="subtle" onClick={() => props.onPrint!(true)} title="Uložiť ako PDF">PDF</Button>}
          {props.onBlockSender && <Button size="small" appearance="subtle" onClick={props.onBlockSender} title="Pridať odosielateľa medzi blokovaných v Zimbre">Blokovať</Button>}
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
      {m.attachments.length > 0 && (
        <div className="attachments">
          {m.attachments.map((a, i) => (
            <span key={i} className="attachment">
              <button
                className="attachment-open"
                disabled={demo}
                title="Otvoriť"
                onClick={() => void api?.openAttachment(m.accountId, m.folder, m.uid, i)}
              >
                <AttachIcon /> {a.filename} <span className="muted">({Math.ceil(a.size / 1024)} kB)</span>
              </button>
              <Button
                size="small"
                appearance="transparent"
                disabled={demo}
                title="Náhľad: obrázky priamo tu, PDF v okne náhľadu, Office v nainštalovanom programe"
                onClick={async () => {
                  const r = await api?.previewAttachment(m.accountId, m.folder, m.uid, i);
                  if (r?.ok && r.data.kind === 'image' && r.data.src) setPreview({ src: r.data.src, name: r.data.filename });
                }}
              >
                Náhľad
              </Button>
              <Button size="small" appearance="transparent" disabled={demo} onClick={() => void api?.saveAttachment(m.accountId, m.folder, m.uid, i)}>
                Uložiť
              </Button>
            </span>
          ))}
        </div>
      )}
      {m.html && !showImages && (
        <div className="images-bar">
          Vzdialené obrázky sú zablokované kvôli súkromiu. <Button size="small" appearance="transparent" onClick={() => setShowImages(true)}>Zobraziť obrázky</Button>
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
