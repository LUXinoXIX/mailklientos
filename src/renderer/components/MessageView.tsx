// Správa otvorená celá: v samostatnom okne alebo na karte hlavného okna (ako Thunderbird).
// Má vlastný pás tlačidiel; zmeny oznámi ostatným oknám (zoznam sa obnoví).
import { Toolbar, ToolbarButton, ToolbarDivider } from '@fluentui/react-components';
import { useCallback, useEffect, useState } from 'react';
import type { MessageFull, MessageSummary } from '../../core/api';
import { PIN_KEYWORD, type Category } from '../../core/settings';
import {
  DeleteIcon, FlagIcon, ForwardIcon, MarkReadIcon, MarkUnreadIcon, NewWindowIcon, NextIcon, PinIcon, PrevIcon, PrintIcon, ReplyAllIcon, ReplyIcon,
} from '../icons';
import { ReadingPane } from './MailPanes';

export interface MessageRef {
  accountId: string;
  folder: string;
  uid: number;
}

export function MessageView(props: {
  start: MessageRef;
  color: string;
  categories?: Category[];
  /** samostatné okno (nastaví názov okna, po odstránení posledného sa zatvorí) */
  standalone?: boolean;
  onTitle?: (t: string) => void;
  onPopOut?: (r: MessageRef) => void;
  onClose?: () => void;
}) {
  const api = window.mailApi;
  const [ref, setRef] = useState<MessageRef>(props.start);
  const [msg, setMsg] = useState<MessageFull | undefined>();
  const [state, setState] = useState<{ loading: boolean; error?: string }>({ loading: true });
  const [list, setList] = useState<MessageSummary[]>([]);
  const [status, setStatus] = useState('');

  const changed = useCallback(() => void api?.notifyChanged(ref.accountId, ref.folder), [api, ref]);

  useEffect(() => {
    if (!api) return;
    let off = false;
    setState((s) => ({ ...s, loading: true }));
    void api.getMessage(ref.accountId, ref.folder, ref.uid).then(async (r) => {
      if (off) return;
      if (!r.ok) return setState({ loading: false, error: r.error });
      setMsg(r.data);
      setState({ loading: false });
      props.onTitle?.(r.data.subject);
      if (props.standalone) document.title = `${r.data.subject} – Universum eMail`;
      if (!r.data.seen) {
        const f = await api.setFlags(ref.accountId, ref.folder, [ref.uid], { seen: true });
        if (f.ok) {
          setMsg((m) => (m ? { ...m, seen: true } : m));
          changed();
        }
      }
    });
    return () => {
      off = true;
    };
  }, [api, ref]); // eslint-disable-line react-hooks/exhaustive-deps

  // poradie správ v priečinku pre šípky Predchádzajúca / Nasledujúca
  useEffect(() => {
    if (!api) return;
    void api.listCached(ref.accountId, ref.folder, 2000).then((r) => r.ok && setList([...r.data].sort((a, b) => (b.date ?? '').localeCompare(a.date ?? ''))));
  }, [api, ref.accountId, ref.folder]);

  const idx = list.findIndex((x) => x.uid === ref.uid);
  const go = (d: number) => {
    const n = list[idx + d];
    if (n) setRef({ accountId: n.accountId, folder: n.folder, uid: n.uid });
  };

  async function flags(f: { seen?: boolean; flagged?: boolean }) {
    if (!api || !msg) return;
    const r = await api.setFlags(ref.accountId, ref.folder, [ref.uid], f);
    if (!r.ok) return setStatus(r.error);
    setMsg({ ...msg, ...f });
    changed();
  }
  async function pin() {
    if (!api || !msg) return;
    const on = !msg.keywords?.includes(PIN_KEYWORD);
    const r = await api.setKeywords(ref.accountId, ref.folder, [ref.uid], on ? [PIN_KEYWORD] : [], on ? [] : [PIN_KEYWORD]);
    if (!r.ok) return setStatus(r.error);
    const kw = new Set(msg.keywords ?? []);
    if (on) kw.add(PIN_KEYWORD);
    else kw.delete(PIN_KEYWORD);
    setMsg({ ...msg, keywords: [...kw] });
    changed();
  }
  async function del() {
    if (!api) return;
    const r = await api.deleteMessages(ref.accountId, ref.folder, [ref.uid]);
    if (!r.ok) return setStatus(r.error);
    changed();
    const next = list[idx + 1] ?? list[idx - 1];
    setList((l) => l.filter((x) => x.uid !== ref.uid));
    if (next) setRef({ accountId: next.accountId, folder: next.folder, uid: next.uid });
    else if (props.standalone) window.close();
    else props.onClose?.();
  }
  const reply = (mode: 'reply' | 'replyAll' | 'forward') => void api?.replyWindow(ref.accountId, ref.folder, ref.uid, mode);

  // klávesy v okne správy
  useEffect(() => {
    const h = (e: KeyboardEvent) => {
      const ctrl = e.ctrlKey || e.metaKey;
      const k = e.key.toLowerCase();
      if (ctrl && e.shiftKey && k === 'r') return e.preventDefault(), reply('replyAll');
      if (ctrl && k === 'r') return e.preventDefault(), reply('reply');
      if (ctrl && k === 'f') return e.preventDefault(), reply('forward');
      if (ctrl && k === 'p') return e.preventDefault(), void api?.printMessage(ref.accountId, ref.folder, ref.uid, false);
      if (e.key === 'Delete') return e.preventDefault(), void del();
      if (e.key === 'Insert') return e.preventDefault(), void flags({ flagged: !msg?.flagged });
      if (e.key === 'Escape' && props.standalone) return window.close();
      if (e.key === 'ArrowUp' && e.altKey) return go(-1);
      if (e.key === 'ArrowDown' && e.altKey) return go(1);
    };
    if (!props.standalone) return;
    window.addEventListener('keydown', h);
    return () => window.removeEventListener('keydown', h);
  });

  const pinned = !!msg?.keywords?.includes(PIN_KEYWORD);
  return (
    <div className="msgview">
      <Toolbar className="msgview-bar" size="small" aria-label="Akcie správy">
        <ToolbarButton icon={<ReplyIcon />} onClick={() => reply('reply')}>Odpovedať</ToolbarButton>
        <ToolbarButton icon={<ReplyAllIcon />} onClick={() => reply('replyAll')}>Všetkým</ToolbarButton>
        <ToolbarButton icon={<ForwardIcon />} onClick={() => reply('forward')}>Preposlať</ToolbarButton>
        <ToolbarDivider />
        <ToolbarButton icon={<FlagIcon />} aria-pressed={!!msg?.flagged} onClick={() => void flags({ flagged: !msg?.flagged })}>{msg?.flagged ? 'Zrušiť vlajku' : 'Vlajka'}</ToolbarButton>
        <ToolbarButton icon={<PinIcon />} aria-pressed={pinned} onClick={() => void pin()}>{pinned ? 'Odopnúť' : 'Pripnúť'}</ToolbarButton>
        <ToolbarButton icon={msg?.seen ? <MarkUnreadIcon /> : <MarkReadIcon />} onClick={() => void flags({ seen: !msg?.seen })}>{msg?.seen ? 'Neprečítané' : 'Prečítané'}</ToolbarButton>
        <ToolbarButton icon={<DeleteIcon />} onClick={() => void del()}>Odstrániť</ToolbarButton>
        <ToolbarDivider />
        <ToolbarButton icon={<PrintIcon />} onClick={() => void api?.printMessage(ref.accountId, ref.folder, ref.uid, false)} aria-label="Tlačiť" title="Tlačiť (Ctrl+P)" />
        <ToolbarButton icon={<PrevIcon />} disabled={idx <= 0} onClick={() => go(-1)} aria-label="Predchádzajúca (novšia) správa" title="Predchádzajúca správa (Alt+↑)" />
        <ToolbarButton icon={<NextIcon />} disabled={idx < 0 || idx >= list.length - 1} onClick={() => go(1)} aria-label="Nasledujúca (staršia) správa" title="Nasledujúca správa (Alt+↓)" />
        {props.onPopOut && <ToolbarButton icon={<NewWindowIcon />} onClick={() => props.onPopOut!(ref)}>Do okna</ToolbarButton>}
        {status && <span className="msgview-status">{status}</span>}
      </Toolbar>
      <ReadingPane
        message={msg}
        loading={state.loading}
        error={state.error}
        color={props.color}
        categories={props.categories}
        onReply={() => reply('reply')}
        onReplyAll={() => reply('replyAll')}
        onForward={() => reply('forward')}
        onPrint={(pdf) => void api?.printMessage(ref.accountId, ref.folder, ref.uid, pdf)}
        onUnsubscribe={
          msg?.unsubscribe
            ? async () => {
                const r = await api!.unsubscribe(ref.accountId, msg.unsubscribe!);
                setStatus(r.ok ? r.data : r.error);
              }
            : undefined
        }
      />
    </div>
  );
}
