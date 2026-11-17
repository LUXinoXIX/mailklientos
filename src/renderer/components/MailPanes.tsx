// Panel kont, zoznam správ, čítací panel a 4. stĺpec (variant D).
import { Button, Spinner, Tab, TabList } from '@fluentui/react-components';
import DOMPurify from 'dompurify';
import { useMemo, useState } from 'react';
import type { ImapAccountPublic } from '../../core/account';
import type { FolderInfo, MessageFull, MessageSummary } from '../../core/api';
import { AttachIcon, FlagFilledIcon, FolderIcon, ForwardIcon, InboxIcon, ReplyAllIcon, ReplyIcon, SentIcon, TrashIcon } from '../icons';
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
}) {
  const a = props.account;
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
            className={`folder ${props.selection.accountId === a.id && props.selection.folder === fo.path ? 'active' : ''}`}
            style={{ paddingLeft: 14 + fo.depth * 14 }}
            onClick={() => props.onSelect({ accountId: a.id, folder: fo.path })}
          >
            {folderIcon(fo)} <span className="grow">{fo.name}</span>
            {fo.specialUse === '\\Inbox' && (props.unread[a.id] ?? 0) > 0 && <span className="count">{props.unread[a.id]}</span>}
          </button>
        ))
      )}
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
}) {
  const [tab, setTab] = useState<'all' | 'unread'>('all');
  const list = tab === 'unread' ? props.messages.filter((m) => !m.seen) : props.messages;
  return (
    <section className="list">
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
        {props.loading && list.length === 0 && <div className="empty">Načítavam…</div>}
        {!props.loading && !props.error && list.length === 0 && <div className="empty">Žiadne správy</div>}
        {list.map((m) => {
          const sel = props.selected && props.selected.uid === m.uid && props.selected.accountId === m.accountId && props.selected.folder === m.folder;
          return (
            <button
              key={`${m.accountId}/${m.folder}/${m.uid}`}
              className={`msg ${sel ? 'selected' : ''} ${m.seen ? '' : 'unread'}`}
              onClick={() => props.onOpen(m)}
            >
              <Avatar name={m.fromName || m.fromAddress} color={props.colorOf(m.accountId)} />
              <div className="msg-main">
                <div className="msg-row">
                  <span className="msg-from">{m.fromName || m.fromAddress}</span>
                  <span className="msg-date">{fmtDate(m.date)}</span>
                </div>
                <div className="msg-row">
                  <span className="msg-subject">{m.subject}</span>
                  {m.flagged && <FlagFilledIcon className="flag" />}
                </div>
              </div>
            </button>
          );
        })}
      </div>
    </section>
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
}) {
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
          </div>
          <div className="muted">Komu: {m.to}{m.cc ? ` · Kópia: ${m.cc}` : ''}</div>
          <div className="muted">{m.date ? new Date(m.date).toLocaleString('sk-SK') : ''}</div>
        </div>
        <div className="read-actions">
          <Button size="small" icon={<ReplyIcon />} onClick={props.onReply}>Odpovedať</Button>
          <Button size="small" icon={<ReplyAllIcon />} onClick={props.onReplyAll} aria-label="Odpovedať všetkým" title="Odpovedať všetkým" />
          <Button size="small" icon={<ForwardIcon />} onClick={props.onForward} aria-label="Preposlať" title="Preposlať" />
        </div>
      </div>
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
      <div className="read-body">{m.html ? <HtmlBody html={m.html} showImages={showImages} /> : <pre className="plain">{m.text}</pre>}</div>
    </section>
  );
}
