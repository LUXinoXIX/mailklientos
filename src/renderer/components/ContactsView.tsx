// Kontakty konta (CardDAV, napr. adresár Zimbry): hľadanie, detail, pridanie, úprava, odstránenie, napísať e-mail.
import { Button, Field, Input, Spinner } from '@fluentui/react-components';
import { useEffect, useMemo, useState } from 'react';
import type { ImapAccountPublic } from '../../core/account';
import type { DavCollection } from '../../core/api';
import type { Contact } from '../../core/ical';
import { accountColor } from '../theme';
import { Avatar } from './MailPanes';

export function ContactsView(props: {
  account: ImapAccountPublic;
  onCompose: (to: string) => void;
  onStatus: (s: string) => void;
  onLoaded?: (list: Contact[]) => void;
  zimbra?: boolean;
}) {
  const api = window.mailApi;
  const [list, setList] = useState<Contact[] | null>(null);
  const [books, setBooks] = useState<DavCollection[]>([]);
  const [q, setQ] = useState('');
  const [sel, setSel] = useState<Contact | null>(null);
  const [edit, setEdit] = useState<Contact | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [tick, setTick] = useState(0);
  const [source, setSource] = useState<'mine' | 'gal'>('mine');
  const [gal, setGal] = useState<Contact[] | null>(null);
  useEffect(() => {
    if (source !== 'gal' || !api) return;
    let alive = true;
    const t = window.setTimeout(() => {
      void api.zimbraGal(props.account.id, q.trim() || '.').then((r) => {
        if (!alive) return;
        if (r.ok)
          setGal(r.data.map((g) => ({ uid: g.email, name: g.name, emails: [g.email], phones: [g.phone, g.mobile].filter(Boolean) as string[], org: [g.company, g.department].filter(Boolean).join(' · ') || undefined, title: g.title })));
        else setError(r.error);
      });
    }, 250);
    return () => {
      alive = false;
      window.clearTimeout(t);
    };
  }, [source, q, api, props.account.id]);

  useEffect(() => {
    if (!api) return;
    let alive = true;
    setError(null);
    void Promise.all([api.cardList(props.account.id), api.davCollections(props.account.id)]).then(([r, b]) => {
      if (!alive) return;
      if (b.ok) setBooks(b.data.filter((x) => x.kind === 'addressbook'));
      if (r.ok) {
        setList(r.data);
        props.onLoaded?.(r.data);
      } else {
        setList([]);
        setError(`${r.error}${r.hint ? ` ${r.hint}` : ''}`);
      }
    });
    return () => {
      alive = false;
    };
  }, [api, props.account.id, tick]); // eslint-disable-line react-hooks/exhaustive-deps

  const shownMine = useMemo(() => {
    const n = (x: string) => x.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
    const qq = n(q.trim());
    return (list ?? []).filter((c) => !qq || n(`${c.name} ${c.emails.join(' ')} ${c.org ?? ''} ${c.phones.join(' ')}`).includes(qq));
  }, [list, q]);
  const shown = source === 'gal' ? gal ?? [] : shownMine;

  async function save() {
    if (!edit || !api) return;
    if (!edit.name.trim()) return setError('Zadaj meno.');
    const r = await api.cardSave(props.account.id, edit.bookId ?? books[0]?.id ?? '', { ...edit, emails: edit.emails.filter(Boolean), phones: edit.phones.filter(Boolean) });
    if (!r.ok) return setError(r.error);
    setEdit(null);
    setSel(r.data);
    setTick((t) => t + 1);
    props.onStatus('Kontakt uložený.');
  }
  async function remove(c: Contact) {
    if (!api || !c.href) return;
    const r = await api.cardDelete(props.account.id, c.href);
    if (!r.ok) return setError(r.error);
    setSel(null);
    setTick((t) => t + 1);
    props.onStatus('Kontakt odstránený.');
  }

  const color = accountColor(props.account.color);
  return (
    <section className="contacts">
      <div className="contacts-list">
        <div className="contacts-head">
          <Input className="grow" placeholder="Hľadať kontakt" value={q} onChange={(_, d) => setQ(d.value)} />
          <Button appearance="primary" disabled={!books.length} onClick={() => setEdit({ uid: '', name: '', emails: [''], phones: [''], bookId: books[0]?.id })}>Nový</Button>
        </div>
        {props.zimbra && (
          <div className="contacts-src">
            <button className={source === 'mine' ? 'on' : ''} onClick={() => setSource('mine')}>Moje kontakty</button>
            <button className={source === 'gal' ? 'on' : ''} onClick={() => setSource('gal')}>Adresár firmy</button>
          </div>
        )}
        {list === null && <Spinner size="tiny" className="pad" />}
        {error && <div className="pane-error pad">⚠ {error}</div>}
        {shown.map((c) => (
          <button key={c.href ?? c.uid ?? c.name} className={`contact-row ${sel === c ? 'active' : ''}`} onClick={() => { setSel(c); setEdit(null); }}>
            <Avatar name={c.name || c.emails[0] || '?'} color={color} size={30} />
            <span className="contact-text">
              <b>{c.name || c.emails[0]}</b>
              <span className="muted small">{c.emails[0] ?? c.phones[0] ?? ''}</span>
            </span>
          </button>
        ))}
        {list && !shown.length && !error && <div className="empty">Žiadne kontakty</div>}
      </div>
      <div className="contact-detail">
        {edit ? (
          <div className="contact-edit">
            <Field label="Meno a priezvisko" required><Input value={edit.name} onChange={(_, d) => setEdit({ ...edit, name: d.value })} /></Field>
            {edit.emails.map((e, i) => (
              <Field key={`e${i}`} label={i ? 'Ďalší e-mail' : 'E-mail'}>
                <Input value={e} onChange={(_, d) => setEdit({ ...edit, emails: edit.emails.map((x, j) => (j === i ? d.value : x)) })} />
              </Field>
            ))}
            <Button size="small" onClick={() => setEdit({ ...edit, emails: [...edit.emails, ''] })}>+ e-mail</Button>
            {edit.phones.map((p, i) => (
              <Field key={`p${i}`} label={i ? 'Ďalší telefón' : 'Telefón'}>
                <Input value={p} onChange={(_, d) => setEdit({ ...edit, phones: edit.phones.map((x, j) => (j === i ? d.value : x)) })} />
              </Field>
            ))}
            <Button size="small" onClick={() => setEdit({ ...edit, phones: [...edit.phones, ''] })}>+ telefón</Button>
            <div className="grid2">
              <Field label="Firma"><Input value={edit.org ?? ''} onChange={(_, d) => setEdit({ ...edit, org: d.value || undefined })} /></Field>
              <Field label="Funkcia"><Input value={edit.title ?? ''} onChange={(_, d) => setEdit({ ...edit, title: d.value || undefined })} /></Field>
            </div>
            <div className="row gap">
              <Button appearance="primary" onClick={() => void save()}>Uložiť</Button>
              <Button onClick={() => setEdit(null)}>Zrušiť</Button>
            </div>
          </div>
        ) : sel ? (
          <div className="contact-card">
            <Avatar name={sel.name || sel.emails[0] || '?'} color={color} size={64} />
            <h2>{sel.name}</h2>
            {(sel.title || sel.org) && <div className="muted">{[sel.title, sel.org].filter(Boolean).join(' · ')}</div>}
            {sel.emails.map((e) => (
              <div key={e} className="row gap">
                <button className="link" onClick={() => props.onCompose(sel.name ? `"${sel.name}" <${e}>` : e)}>{e}</button>
              </div>
            ))}
            {sel.phones.map((p) => <div key={p}><a href={`tel:${p.replace(/\s/g, '')}`}>{p}</a></div>)}
            <div className="row gap" style={{ marginTop: 12 }}>
              {sel.emails[0] && <Button appearance="primary" onClick={() => props.onCompose(sel.name ? `"${sel.name}" <${sel.emails[0]}>` : sel.emails[0]!)}>Napísať e-mail</Button>}
              {source === 'gal' && <Button onClick={() => { setEdit({ ...sel, uid: '', href: undefined, bookId: books[0]?.id }); }}>Uložiť do mojich kontaktov</Button>}
              {source === 'mine' && <Button onClick={() => setEdit({ ...sel, emails: sel.emails.length ? sel.emails : [''], phones: sel.phones.length ? sel.phones : [''] })}>Upraviť</Button>}
              {source === 'mine' && <Button appearance="subtle" onClick={() => void remove(sel)}>Odstrániť</Button>}
            </div>
          </div>
        ) : (
          <div className="empty">Vyberte kontakt</div>
        )}
      </div>
    </section>
  );
}
