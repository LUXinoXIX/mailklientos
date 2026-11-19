// Kontakty spoločné pre všetky kontá (ako v eM Cliente): vľavo zdroje (lokálny adresár v programe,
// adresáre kont na serveri – CardDAV/Zimbra, adresár firmy, naposledy použité), v strede zoznam,
// vpravo detail. Ten istý človek v dvoch zdrojoch (rovnaký e-mail) je jeden riadok s bodkami zdrojov.
import { Button, Field, Input, Menu, MenuItem, MenuList, MenuPopover, MenuTrigger, Select, Spinner, Textarea } from '@fluentui/react-components';
import { useCallback, useEffect, useMemo, useState } from 'react';
import type { ImapAccountPublic } from '../../core/account';
import type { DavCollection } from '../../core/api';
import { contactKeys, mergeContact, toGoogleCsv, toOutlookCsv, toVcf } from '../../core/contactsio';
import type { Contact } from '../../core/ical';
import { ContactsIcon, ServerIcon, SnoozeIcon } from '../icons';
import { accountColor } from '../theme';
import { ContactImport, type ImportTarget } from './ContactImport';
import { Avatar } from './MailPanes';

const LOCAL = 'local';
const LOCAL_COLOR = '#6b7280';

interface Source {
  id: string;
  label: string;
  color: string;
  account?: ImapAccountPublic;
  books: DavCollection[];
  list: Contact[] | null;
  error?: string;
}
interface Item {
  src: string;
  c: Contact;
}
interface Entry {
  key: string;
  c: Contact;
  items: Item[];
}

type View = 'all' | 'recent' | string; // string = id zdroja alebo gal:<id konta>

const norm = (x: string) => x.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();

export function AddressBook(props: {
  accounts: ImapAccountPublic[];
  zimbraOk: Record<string, boolean>;
  /** naposledy použité adresy (z odoslaných správ) */
  recent: { name: string; email: string }[];
  defaultAccountId?: string;
  onCompose: (to: string, accountId?: string) => void;
  onStatus: (s: string) => void;
  /** zmena kontaktov → obnoviť dopĺňanie adries */
  onChanged: () => void;
}) {
  const api = window.mailApi;
  const [sources, setSources] = useState<Source[]>([]);
  const [view, setView] = useState<View>('all');
  const [q, setQ] = useState('');
  const [selKey, setSelKey] = useState<string | null>(null);
  const [edit, setEdit] = useState<{ c: Contact; src: string; bookId?: string } | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [gal, setGal] = useState<Contact[] | null>(null);
  const [importing, setImporting] = useState(false);
  const [tick, setTick] = useState(0);

  // načítanie: lokálny adresár + adresáre všetkých IMAP kont (CardDAV)
  useEffect(() => {
    if (!api) return;
    let alive = true;
    const base: Source[] = [
      { id: LOCAL, label: 'Lokálne (v programe)', color: LOCAL_COLOR, books: [], list: null },
      ...props.accounts.map((a) => ({ id: a.id, label: a.name || a.email, color: accountColor(a.color), account: a, books: [], list: null })),
    ];
    setSources((prev) => base.map((s) => ({ ...s, list: prev.find((p) => p.id === s.id)?.list ?? null, books: prev.find((p) => p.id === s.id)?.books ?? [] })));
    const upd = (id: string, patch: Partial<Source>) => alive && setSources((xs) => xs.map((s) => (s.id === id ? { ...s, ...patch } : s)));
    void api.localContacts().then((r) => upd(LOCAL, r.ok ? { list: r.data } : { list: [], error: r.error }));
    for (const a of props.accounts)
      void Promise.all([api.davCollections(a.id), api.cardList(a.id)]).then(([b, r]) => {
        const books = b.ok ? b.data.filter((x) => x.kind === 'addressbook') : [];
        upd(a.id, { books, list: r.ok ? r.data : [], error: r.ok ? undefined : r.error });
      });
    return () => {
      alive = false;
    };
  }, [api, props.accounts, tick]);

  // zdroje, ktoré naozaj majú adresár (konto bez CardDAV sa nezobrazí)
  const shownSources = sources.filter((s) => s.id === LOCAL || s.books.length > 0 || (s.list?.length ?? 0) > 0);
  const loading = sources.some((s) => s.list === null);
  const galAccounts = props.accounts.filter((a) => props.zimbraOk[a.id]);

  // adresár firmy (Zimbra GAL) – hľadá sa na serveri podľa poľa hľadania
  const galId = view.startsWith('gal:') ? view.slice(4) : null;
  useEffect(() => {
    if (!galId || !api) return;
    let alive = true;
    setGal(null);
    const t = window.setTimeout(() => {
      void api.zimbraGal(galId, q.trim() || '.').then((r) => {
        if (!alive) return;
        if (r.ok)
          setGal(r.data.map((g) => ({ uid: g.email, name: g.name, emails: [g.email], phones: [g.phone, g.mobile].filter(Boolean) as string[], org: [g.company, g.department].filter(Boolean).join(' · ') || undefined, title: g.title })));
        else {
          setGal([]);
          setError(r.error);
        }
      });
    }, 250);
    return () => {
      alive = false;
      window.clearTimeout(t);
    };
  }, [galId, q, api]);

  // zlúčenie rovnakých ľudí naprieč zdrojmi
  const entries = useMemo<Entry[]>(() => {
    const items: Item[] =
      view === 'all' ? shownSources.flatMap((s) => (s.list ?? []).map((c) => ({ src: s.id, c })))
      : view === 'recent' ? props.recent.map((r) => ({ src: 'recent', c: { uid: r.email, name: r.name, emails: [r.email], phones: [] } }))
      : galId ? (gal ?? []).map((c) => ({ src: view, c }))
      : (shownSources.find((s) => s.id === view)?.list ?? []).map((c) => ({ src: view, c }));
    const byKey = new Map<string, Entry>();
    const out: Entry[] = [];
    for (const it of items) {
      const keys = contactKeys(it.c);
      const hit = keys.map((k) => byKey.get(k)).find(Boolean);
      if (hit) {
        hit.c = mergeContact(hit.c, it.c);
        hit.items.push(it);
        for (const k of contactKeys(hit.c)) byKey.set(k, hit);
      } else {
        const e: Entry = { key: keys[0] ?? `${it.src}:${it.c.uid}`, c: it.c, items: [it] };
        out.push(e);
        for (const k of keys) byKey.set(k, e);
      }
    }
    return out.sort((a, b) => (a.c.name || a.c.emails[0] || '').localeCompare(b.c.name || b.c.emails[0] || '', 'sk'));
  }, [view, shownSources, props.recent, gal, galId]);

  const shown = useMemo(() => {
    if (galId) return entries;
    const qq = norm(q.trim());
    return entries.filter((e) => !qq || norm(`${e.c.name} ${e.c.emails.join(' ')} ${e.c.org ?? ''} ${e.c.phones.join(' ')}`).includes(qq));
  }, [entries, q, galId]);
  const sel = shown.find((e) => e.key === selKey) ?? entries.find((e) => e.key === selKey) ?? null;

  const srcOf = useCallback((id: string) => sources.find((s) => s.id === id), [sources]);
  const writable = shownSources.filter((s) => s.id === LOCAL || s.books.length > 0);

  function reload(msg?: string) {
    setTick((t) => t + 1);
    props.onChanged();
    if (msg) props.onStatus(msg);
  }

  async function saveTo(srcId: string, c: Contact, bookId?: string): Promise<Contact> {
    if (!api) throw new Error('Nedostupné');
    const clean = { ...c, name: c.name.trim(), emails: c.emails.map((x) => x.trim()).filter(Boolean), phones: c.phones.map((x) => x.trim()).filter(Boolean) };
    if (srcId === LOCAL) {
      const r = await api.localPut([clean]);
      if (!r.ok) throw new Error(r.error);
      return r.data[0]!;
    }
    const s = srcOf(srcId);
    const r = await api.cardSave(srcId, bookId ?? c.bookId ?? s?.books[0]?.id ?? '', clean);
    if (!r.ok) throw new Error(r.error);
    return r.data;
  }

  async function save() {
    if (!edit) return;
    if (!edit.c.name.trim() && !edit.c.emails.some((x) => x.trim())) return setError('Zadaj meno alebo e-mail.');
    try {
      const saved = await saveTo(edit.src, edit.c, edit.bookId);
      setEdit(null);
      setSelKey(contactKeys(saved)[0] ?? null);
      reload('Kontakt uložený.');
    } catch (e) {
      setError((e as Error).message);
    }
  }

  async function remove(it: Item) {
    if (!api) return;
    const r = it.src === LOCAL ? await api.localRemove([it.c.uid]) : it.c.href ? await api.cardDelete(it.src, it.c.href) : null;
    if (r && !r.ok) return setError(r.error);
    setSelKey(null);
    reload('Kontakt odstránený.');
  }

  async function exportAs(kind: 'vcf' | 'outlook' | 'google') {
    if (!api) return;
    const list = shown.map((e) => e.c);
    if (!list.length) return props.onStatus('Nie je čo exportovať.');
    const d = new Date().toISOString().slice(0, 10);
    const [name, text] =
      kind === 'vcf' ? [`kontakty-${d}.vcf`, toVcf(list)]
      : kind === 'outlook' ? [`kontakty-outlook-${d}.csv`, toOutlookCsv(list)]
      : [`kontakty-google-${d}.csv`, toGoogleCsv(list)];
    const r = await api.contactsSaveExport(name, text);
    if (!r.ok) setError(r.error);
    else if (r.data) props.onStatus(`Exportovaných ${list.length} kontaktov do ${r.data}`);
  }

  const importTargets: ImportTarget[] = writable.map((s) => ({
    id: s.id,
    label: s.label,
    color: s.color,
    existing: s.list ?? [],
    save: async (list, progress) => {
      if (s.id === LOCAL) {
        const r = await api!.localPut(list);
        progress(list.length);
        return r.ok ? { ok: list.length, failed: [] } : { ok: 0, failed: [r.error] };
      }
      const failed: string[] = [];
      let i = 0;
      for (const c of list) {
        const r = await api!.cardSave(s.id, c.bookId ?? s.books[0]?.id ?? '', c);
        if (!r.ok) failed.push(c.name || c.emails[0] || '?');
        progress(++i);
      }
      return { ok: list.length - failed.length, failed };
    },
  }));

  const newContact = (src = view !== 'all' && writable.some((s) => s.id === view) ? view : LOCAL, base?: Contact) =>
    setEdit({ src, bookId: srcOf(src)?.books[0]?.id, c: { uid: '', name: base?.name ?? '', emails: base?.emails.length ? [...base.emails] : [''], phones: base?.phones.length ? [...base.phones] : [''], org: base?.org, title: base?.title, note: base?.note } });

  const composeTo = (e: Entry, email: string) => {
    const acc = e.items.find((i) => i.src !== LOCAL && i.src !== 'recent' && !i.src.startsWith('gal:'))?.src ?? (galId || props.defaultAccountId);
    props.onCompose(e.c.name ? `"${e.c.name}" <${email}>` : email, acc);
  };
  const editable = sel ? sel.items.filter((i) => i.src === LOCAL || (srcOf(i.src)?.books.length ?? 0) > 0) : [];
  const missingIn = sel ? writable.filter((s) => !sel.items.some((i) => i.src === s.id)) : [];
  const viewLabel = view === 'all' ? 'Všetky kontakty' : view === 'recent' ? 'Naposledy použité' : galId ? 'Adresár firmy' : srcOf(view)?.label ?? '';

  return (
    <section className="contacts ab">
      <nav className="ab-sources" aria-label="Zdroje kontaktov">
        <button className={`ab-src ${view === 'all' ? 'on' : ''}`} onClick={() => setView('all')}>
          <span className="ab-ico"><ContactsIcon /></span><span className="grow">Všetky kontakty</span><span className="muted small">{loading ? '…' : entries.length && view === 'all' ? entries.length : ''}</span>
        </button>
        <div className="ab-h">Adresáre</div>
        {shownSources.map((s) => (
          <button key={s.id} className={`ab-src ${view === s.id ? 'on' : ''}`} onClick={() => setView(s.id)} title={s.error ?? s.label}>
            <span className="dot" style={{ background: s.color }} />
            <span className="grow">{s.label}</span>
            <span className="muted small">{s.list === null ? '…' : s.list.length}</span>
          </button>
        ))}
        {galAccounts.length > 0 && <div className="ab-h">Firma</div>}
        {galAccounts.map((a) => (
          <button key={a.id} className={`ab-src ${view === `gal:${a.id}` ? 'on' : ''}`} onClick={() => setView(`gal:${a.id}`)}>
            <span className="ab-ico"><ServerIcon /></span><span className="grow">Adresár firmy{galAccounts.length > 1 ? ` · ${a.name}` : ''}</span>
          </button>
        ))}
        <div className="ab-h">Ďalšie</div>
        <button className={`ab-src ${view === 'recent' ? 'on' : ''}`} onClick={() => setView('recent')} title="Každá adresa, na ktorú pošleš správu (aj ručne napísaná), sa sem uloží automaticky a ponúka sa pri písaní">
          <span className="ab-ico"><SnoozeIcon /></span><span className="grow">Naposledy použité</span><span className="muted small">{props.recent.length}</span>
        </button>
        <div className="ab-note muted small">Dopĺňanie adries pri písaní používa všetky zdroje naraz, z ktoréhokoľvek konta píšeš.</div>
      </nav>

      <div className="contacts-list">
        <div className="contacts-head">
          <Input className="grow" placeholder={galId ? 'Hľadať v adresári firmy' : `Hľadať v: ${viewLabel}`} value={q} onChange={(_, d) => setQ(d.value)} />
        </div>
        <div className="ab-tools">
          <Button size="small" appearance="primary" onClick={() => newContact()}>Nový</Button>
          <Button size="small" onClick={() => setImporting(true)}>Import…</Button>
          <Menu>
            <MenuTrigger disableButtonEnhancement>
              <Button size="small" disabled={!shown.length}>Export ▾</Button>
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                <MenuItem onClick={() => void exportAs('vcf')}>vCard .vcf (Apple, Android, Google)</MenuItem>
                <MenuItem onClick={() => void exportAs('outlook')}>CSV pre Outlook</MenuItem>
                <MenuItem onClick={() => void exportAs('google')}>CSV pre Google Kontakty</MenuItem>
              </MenuList>
            </MenuPopover>
          </Menu>
          <span className="muted small grow ab-count">{shown.length} {shown.length === 1 ? 'kontakt' : shown.length > 1 && shown.length < 5 ? 'kontakty' : 'kontaktov'}</span>
        </div>
        {view === 'recent' && <div className="ab-hint muted small">Každá adresa, na ktorú pošleš správu – aj ručne napísaná – sa sem uloží automaticky a ponúka sa pri písaní zo všetkých kont. Natrvalo do adresára ju dáš cez „Uložiť do kontaktov“.</div>}
        {error && <div className="pane-error pad" onClick={() => setError(null)}>⚠ {error}</div>}
        {((galId && gal === null) || (!galId && loading && !entries.length)) && <Spinner size="tiny" className="pad" />}
        {shown.map((e) => (
          <button key={e.key} className={`contact-row ${sel?.key === e.key ? 'active' : ''}`} onClick={() => { setSelKey(e.key); setEdit(null); }}>
            <Avatar name={e.c.name || e.c.emails[0] || '?'} color={srcOf(e.items[0]!.src)?.color ?? LOCAL_COLOR} size={30} />
            <span className="contact-text grow">
              <b>{e.c.name || e.c.emails[0]}</b>
              <span className="muted small">{e.c.emails[0] ?? e.c.phones[0] ?? ''}</span>
            </span>
            {view === 'all' && (
              <span className="ab-dots" title={e.items.map((i) => srcOf(i.src)?.label).join(', ')}>
                {[...new Set(e.items.map((i) => i.src))].map((s) => <span key={s} className="dot" style={{ background: srcOf(s)?.color ?? LOCAL_COLOR }} />)}
              </span>
            )}
          </button>
        ))}
        {!loading && !shown.length && !(galId && gal === null) && (
          <div className="empty ab-empty">
            {q ? 'Nič sa nenašlo.' : view === LOCAL || view === 'all' ? <>Zatiaľ žiadne kontakty.<br /><Button size="small" onClick={() => setImporting(true)}>Importovať z Outlooku, Googlu alebo Apple</Button></> : 'Žiadne kontakty'}
          </div>
        )}
      </div>

      <div className="contact-detail">
        {edit ? (
          <div className="contact-edit">
            <h3 style={{ margin: 0 }}>{edit.c.uid || edit.c.href ? 'Upraviť kontakt' : 'Nový kontakt'}</h3>
            {!edit.c.uid && !edit.c.href && (
              <Field label="Uložiť do">
                <Select value={edit.src} onChange={(_, d) => setEdit({ ...edit, src: d.value, bookId: srcOf(d.value)?.books[0]?.id })}>
                  {writable.map((s) => <option key={s.id} value={s.id}>{s.label}{s.id === LOCAL ? '' : ' (na serveri)'}</option>)}
                </Select>
              </Field>
            )}
            <Field label="Meno a priezvisko"><Input value={edit.c.name} onChange={(_, d) => setEdit({ ...edit, c: { ...edit.c, name: d.value } })} /></Field>
            {edit.c.emails.map((e, i) => (
              <Field key={`e${i}`} label={i ? 'Ďalší e-mail' : 'E-mail'}>
                <Input type="email" value={e} onChange={(_, d) => setEdit({ ...edit, c: { ...edit.c, emails: edit.c.emails.map((x, j) => (j === i ? d.value : x)) } })} />
              </Field>
            ))}
            <Button size="small" onClick={() => setEdit({ ...edit, c: { ...edit.c, emails: [...edit.c.emails, ''] } })}>+ e-mail</Button>
            {edit.c.phones.map((p, i) => (
              <Field key={`p${i}`} label={i ? 'Ďalší telefón' : 'Telefón'}>
                <Input value={p} onChange={(_, d) => setEdit({ ...edit, c: { ...edit.c, phones: edit.c.phones.map((x, j) => (j === i ? d.value : x)) } })} />
              </Field>
            ))}
            <Button size="small" onClick={() => setEdit({ ...edit, c: { ...edit.c, phones: [...edit.c.phones, ''] } })}>+ telefón</Button>
            <div className="grid2">
              <Field label="Firma"><Input value={edit.c.org ?? ''} onChange={(_, d) => setEdit({ ...edit, c: { ...edit.c, org: d.value || undefined } })} /></Field>
              <Field label="Funkcia"><Input value={edit.c.title ?? ''} onChange={(_, d) => setEdit({ ...edit, c: { ...edit.c, title: d.value || undefined } })} /></Field>
            </div>
            <Field label="Poznámka"><Textarea value={edit.c.note ?? ''} onChange={(_, d) => setEdit({ ...edit, c: { ...edit.c, note: d.value || undefined } })} /></Field>
            <div className="row gap">
              <Button appearance="primary" onClick={() => void save()}>Uložiť</Button>
              <Button onClick={() => setEdit(null)}>Zrušiť</Button>
            </div>
          </div>
        ) : sel ? (
          <div className="contact-card">
            <Avatar name={sel.c.name || sel.c.emails[0] || '?'} color={srcOf(sel.items[0]!.src)?.color ?? LOCAL_COLOR} size={64} />
            <h2>{sel.c.name || sel.c.emails[0]}</h2>
            {(sel.c.title || sel.c.org) && <div className="muted">{[sel.c.title, sel.c.org].filter(Boolean).join(' · ')}</div>}
            {sel.c.emails.map((e) => (
              <button key={e} className="link" onClick={() => composeTo(sel, e)}>{e}</button>
            ))}
            {sel.c.phones.map((p) => <div key={p}><a href={`tel:${p.replace(/\s/g, '')}`}>{p}</a></div>)}
            {sel.c.note && <div className="ab-notetext">{sel.c.note}</div>}
            <div className="ab-stored muted small">
              Uložené v:{' '}
              {sel.items.map((i, n) => (
                <span key={n} className="ab-chip"><span className="dot" style={{ background: srcOf(i.src)?.color ?? LOCAL_COLOR }} />{i.src === 'recent' ? 'naposledy použité' : i.src.startsWith('gal:') ? 'adresár firmy' : srcOf(i.src)?.label}</span>
              ))}
            </div>
            <div className="row gap" style={{ marginTop: 12, flexWrap: 'wrap' }}>
              {sel.c.emails[0] && <Button appearance="primary" onClick={() => composeTo(sel, sel.c.emails[0]!)}>Napísať e-mail</Button>}
              {editable.length === 1 && <Button onClick={() => setEdit({ src: editable[0]!.src, c: { ...editable[0]!.c, emails: editable[0]!.c.emails.length ? editable[0]!.c.emails : [''], phones: editable[0]!.c.phones.length ? editable[0]!.c.phones : [''] } })}>Upraviť</Button>}
              {editable.length > 1 && (
                <Menu>
                  <MenuTrigger disableButtonEnhancement><Button>Upraviť ▾</Button></MenuTrigger>
                  <MenuPopover><MenuList>
                    {editable.map((i, n) => <MenuItem key={n} onClick={() => setEdit({ src: i.src, c: { ...i.c, emails: i.c.emails.length ? i.c.emails : [''], phones: i.c.phones.length ? i.c.phones : [''] } })}>v: {srcOf(i.src)?.label}</MenuItem>)}
                  </MenuList></MenuPopover>
                </Menu>
              )}
              {missingIn.length > 0 && (
                <Menu>
                  <MenuTrigger disableButtonEnhancement><Button>{editable.length ? 'Kopírovať do ▾' : 'Uložiť do kontaktov ▾'}</Button></MenuTrigger>
                  <MenuPopover><MenuList>
                    {missingIn.map((s) => <MenuItem key={s.id} onClick={() => newContact(s.id, sel.c)}>{s.label}</MenuItem>)}
                  </MenuList></MenuPopover>
                </Menu>
              )}
              {editable.length === 1 && <Button appearance="subtle" onClick={() => void remove(editable[0]!)}>Odstrániť</Button>}
              {view === 'recent' && (
                <Button appearance="subtle" title="Adresa sa už nebude ponúkať pri písaní (kým na ňu znova nenapíšeš)" onClick={() => void api?.recentRemove(sel.c.emails).then(() => { setSelKey(null); props.onChanged(); })}>
                  Zabudnúť adresu
                </Button>
              )}
              {editable.length > 1 && (
                <Menu>
                  <MenuTrigger disableButtonEnhancement><Button appearance="subtle">Odstrániť ▾</Button></MenuTrigger>
                  <MenuPopover><MenuList>
                    {editable.map((i, n) => <MenuItem key={n} onClick={() => void remove(i)}>z: {srcOf(i.src)?.label}</MenuItem>)}
                  </MenuList></MenuPopover>
                </Menu>
              )}
            </div>
          </div>
        ) : (
          <div className="empty">Vyberte kontakt</div>
        )}
      </div>

      {importing && (
        <ContactImport
          targets={importTargets}
          defaultTarget={writable.some((s) => s.id === view) ? view : LOCAL}
          onClose={() => setImporting(false)}
          onDone={(msg) => {
            setImporting(false);
            reload(msg);
          }}
        />
      )}
    </section>
  );
}
