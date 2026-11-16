// Panel kont, zoznam správ, čítací panel a 4. stĺpec (variant D).
import { Button, Spinner, Tab, TabList } from '@fluentui/react-components';
import DOMPurify from 'dompurify';
import { useMemo, useState } from 'react';
import type { AccountPublic } from '../../core/account';
import type { FolderInfo, MessageFull, MessageSummary } from '../../core/api';
import { AttachIcon, FolderIcon, InboxIcon, SentIcon, TrashIcon } from '../icons';

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

export function AccountPane(props: {
  accounts: AccountPublic[];
  folders: Record<string, FolderInfo[] | 'loading' | { error: string }>;
  selection: Selection;
  onSelect: (s: Selection) => void;
}) {
  const [collapsed, setCollapsed] = useState<Record<string, boolean>>({});
  return (
    <aside className="accounts">
      <button
        className={`folder all ${props.selection.accountId === 'all' ? 'active' : ''}`}
        onClick={() => props.onSelect({ accountId: 'all', folder: 'INBOX' })}
      >
        <InboxIcon /> <span>Všetky doručené</span>
      </button>
      {props.accounts.map((a) => {
        const f = props.folders[a.id];
        return (
          <div key={a.id} className="account">
            <button className="account-head" onClick={() => setCollapsed((c) => ({ ...c, [a.id]: !c[a.id] }))}>
              <span className="dot" style={{ background: a.color }} />
              <span className="account-name" title={a.email}>
                {a.name}
              </span>
              <span className="chev">{collapsed[a.id] ? '▸' : '▾'}</span>
            </button>
            {!collapsed[a.id] &&
              (f === 'loading' || f === undefined ? (
                <Spinner size="extra-tiny" className="pad" />
              ) : 'error' in f ? (
                <div className="pane-error" title={f.error}>
                  ⚠ {f.error}
                </div>
              ) : (
                f.map((fo) => (
                  <button
                    key={fo.path}
                    className={`folder ${props.selection.accountId === a.id && props.selection.folder === fo.path ? 'active' : ''}`}
                    style={{ paddingLeft: 20 + fo.depth * 14 }}
                    onClick={() => props.onSelect({ accountId: a.id, folder: fo.path })}
                  >
                    {folderIcon(fo)} <span>{fo.name}</span>
                  </button>
                ))
              ))}
          </div>
        );
      })}
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
}) {
  const [tab, setTab] = useState<'all' | 'unread'>('all');
  const list = tab === 'unread' ? props.messages.filter((m) => !m.seen) : props.messages;
  return (
    <section className="list">
      <div className="list-head">
        <div className="list-title">{props.title}</div>
        <TabList size="small" selectedValue={tab} onTabSelect={(_, d) => setTab(d.value as 'all' | 'unread')}>
          <Tab value="all">Všetky</Tab>
          <Tab value="unread">Neprečítané</Tab>
        </TabList>
      </div>
      <div className="list-body">
        {props.loading && <Spinner size="small" label="Načítavam…" className="pad" />}
        {props.error && <div className="pane-error pad">⚠ {props.error}</div>}
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
                <div className="msg-subject">{m.subject}</div>
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

export function ReadingPane(props: { message?: MessageFull; loading: boolean; error?: string; color: string; onReply: () => void }) {
  const [showImages, setShowImages] = useState(false);
  const m = props.message;
  if (props.loading) return <section className="reading"><Spinner label="Otváram správu…" className="pad" /></section>;
  if (props.error) return <section className="reading"><div className="pane-error pad">⚠ {props.error}</div></section>;
  if (!m) return <section className="reading"><div className="empty">Vyber správu na čítanie</div></section>;
  return (
    <section className="reading">
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
        <Button size="small" onClick={props.onReply}>Odpovedať</Button>
      </div>
      {m.attachments.length > 0 && (
        <div className="attachments">
          {m.attachments.map((a, i) => (
            <span key={i} className="attachment"><AttachIcon /> {a.filename} <span className="muted">({Math.ceil(a.size / 1024)} kB)</span></span>
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

export const WEB_SERVICES = {
  copilot: { label: 'Copilot', url: 'https://copilot.microsoft.com' },
  teams: { label: 'Teams', url: 'https://teams.microsoft.com' },
  whatsapp: { label: 'WhatsApp Web', url: 'https://web.whatsapp.com' },
} as const;
export type WebKey = keyof typeof WEB_SERVICES;

export function WebColumn(props: { service: WebKey; onService: (k: WebKey) => void; onClose: () => void }) {
  const s = WEB_SERVICES[props.service];
  return (
    <aside className="web">
      <div className="web-head">
        <TabList size="small" selectedValue={props.service} onTabSelect={(_, d) => props.onService(d.value as WebKey)}>
          {Object.entries(WEB_SERVICES).map(([k, v]) => (
            <Tab key={k} value={k}>{v.label}</Tab>
          ))}
        </TabList>
        <Button size="small" appearance="subtle" onClick={props.onClose} aria-label="Zavrieť">✕</Button>
      </div>
      <div className="web-placeholder">
        <p><b>{s.label}</b> sa tu bude zobrazovať priamo v okne (izolované prihlásenie cez WebContentsView).</p>
        <p className="muted">Pripravené v etape 3. Zatiaľ sa dá otvoriť v prehliadači:</p>
        <Button onClick={() => window.mailApi?.openExternal(s.url)}>Otvoriť {s.label}</Button>
      </div>
    </aside>
  );
}
