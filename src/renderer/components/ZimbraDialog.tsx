// Nastavenia na serveri Zimbra: neprítomnosť, filtre, zdieľania od kolegov, dôveryhodní/blokovaní odosielatelia,
// zmena hesla a prevzatie podpisov, identít a značiek.
import {
  Button, Checkbox, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input, MessageBar, MessageBarBody,
  Select, Spinner, Tab, TabList, Textarea,
} from '@fluentui/react-components';
import { useEffect, useState } from 'react';
import type { ImapAccountPublic } from '../../core/account';
import type { FolderInfo, ServerOoo, ServerRule, ZShare } from '../../core/api';
import type { AppSettings, Category, Signature } from '../../core/settings';

export type ZimbraTab = 'ooo' | 'filters' | 'shares' | 'spam' | 'sync' | 'password';

const FIELD_LABEL: Record<ServerRule['conds'][number]['field'], string> = {
  from: 'Odosielateľ obsahuje',
  to: 'Príjemca obsahuje',
  subject: 'Predmet obsahuje',
  any: 'Odosielateľ, príjemca alebo predmet obsahuje',
};
const ACTION_LABEL: Record<ServerRule['actions'][number]['type'], string> = {
  move: 'Presunúť do priečinka',
  read: 'Označiť ako prečítané',
  flag: 'Označiť vlajkou',
  tag: 'Priradiť značku',
  delete: 'Presunúť do koša',
  stop: 'Ďalšie filtre nespúšťať',
};
const VIEW: Record<string, string> = { message: 'Pošta', appointment: 'Kalendár', contact: 'Kontakty', task: 'Úlohy', document: 'Aktovka' };

export function ZimbraDialog(props: {
  account: ImapAccountPublic;
  folders: FolderInfo[];
  settings: AppSettings;
  tab?: ZimbraTab;
  onSettings: (s: AppSettings) => void;
  onAccountIdentities: (ids: Array<{ name: string; email: string }>) => void;
  onStatus: (s: string) => void;
  onClose: () => void;
}) {
  const api = window.mailApi!;
  const id = props.account.id;
  const [tab, setTab] = useState<ZimbraTab>(props.tab ?? 'ooo');
  const [msg, setMsg] = useState<{ ok: boolean; text: string } | null>(null);
  const [busy, setBusy] = useState(false);
  const [ooo, setOoo] = useState<ServerOoo | null>(null);
  const [rules, setRules] = useState<ServerRule[] | null>(null);
  const [shares, setShares] = useState<ZShare[] | null>(null);
  const [wb, setWb] = useState<{ white: string; black: string } | null>(null);
  const [pw, setPw] = useState({ old: '', n1: '', n2: '' });
  const result = (ok: boolean, text: string) => {
    setBusy(false);
    setMsg({ ok, text });
  };

  useEffect(() => {
    setMsg(null);
    if (tab === 'ooo' && !ooo) void api.zimbraGetOoo(id).then((r) => (r.ok ? setOoo(r.data) : result(false, r.error)));
    if (tab === 'filters' && !rules) void api.zimbraGetFilters(id).then((r) => (r.ok ? setRules(r.data) : result(false, r.error)));
    if (tab === 'shares' && !shares) void api.zimbraShares(id).then((r) => (r.ok ? setShares(r.data) : result(false, r.error)));
    if (tab === 'spam' && !wb) void api.zimbraGetWB(id).then((r) => (r.ok ? setWb({ white: r.data.white.join('\n'), black: r.data.black.join('\n') }) : result(false, r.error)));
  }, [tab]); // eslint-disable-line react-hooks/exhaustive-deps

  async function syncFromServer() {
    setBusy(true);
    const [sigs, ids, tags] = await Promise.all([api.zimbraSignatures(id), api.zimbraIdentities(id), api.zimbraTags(id)]);
    let s = props.settings;
    const notes: string[] = [];
    if (sigs.ok) {
      const list: Signature[] = [...s.signatures];
      let n = 0;
      for (const z of sigs.data) {
        const html = z.html ?? (z.text ? `<p>${z.text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/\n/g, '<br>')}</p>` : '');
        const i = list.findIndex((x) => x.zimbraId === z.id || x.name === z.name);
        if (i >= 0) list[i] = { ...list[i]!, html, text: z.text ?? list[i]!.text, zimbraId: z.id };
        else list.push({ id: Math.random().toString(36).slice(2, 10), name: z.name, html, text: z.text ?? '', zimbraId: z.id });
        n++;
      }
      // vlastné podpisy, ktoré na serveri nie sú, nahrať
      let up = 0;
      for (const sig of list) {
        if (sig.zimbraId) continue;
        const r = await api.zimbraSaveSignature(id, { name: sig.name, html: sig.html ?? '', text: sig.text });
        if (r.ok) {
          sig.zimbraId = r.data;
          up++;
        }
      }
      s = { ...s, signatures: list };
      notes.push(`podpisy: ${n} zo servera, ${up} nahraté`);
    }
    if (tags.ok && tags.data.length) {
      const cats: Category[] = [...(s.categories ?? [])];
      for (const t of tags.data) {
        const i = cats.findIndex((c) => c.id === t.name || c.name === t.name);
        if (i >= 0) cats[i] = { ...cats[i]!, id: t.name, color: t.color };
        else cats.push({ id: t.name, name: t.name, color: t.color });
      }
      s = { ...s, categories: cats };
      notes.push(`značky: ${tags.data.length}`);
    }
    props.onSettings(s);
    if (ids.ok) {
      const others = ids.data.filter((x) => x.email.toLowerCase() !== props.account.email.toLowerCase());
      if (others.length) props.onAccountIdentities(others);
      notes.push(`identity: ${others.length}`);
    }
    result(sigs.ok || tags.ok || ids.ok, `Prevzaté zo Zimbry – ${notes.join(', ')}.`);
  }

  return (
    <Dialog open onOpenChange={(_, d) => !d.open && props.onClose()}>
      <DialogSurface style={{ maxWidth: 900, width: '92vw' }}>
        <DialogBody style={{ display: 'flex', flexDirection: 'column', maxHeight: '88vh' }}>
          <DialogTitle>Zimbra – {props.account.name}</DialogTitle>
          <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: 12, overflow: 'auto' }}>
            <TabList selectedValue={tab} onTabSelect={(_, d) => setTab(d.value as ZimbraTab)}>
              <Tab value="ooo">Neprítomnosť</Tab>
              <Tab value="filters">Filtre na serveri</Tab>
              <Tab value="shares">Zdieľané od kolegov</Tab>
              <Tab value="spam">Spam</Tab>
              <Tab value="sync">Podpisy a značky</Tab>
              <Tab value="password">Heslo</Tab>
            </TabList>
            {msg && <MessageBar intent={msg.ok ? 'success' : 'error'}><MessageBarBody>{msg.text}</MessageBarBody></MessageBar>}

            {tab === 'ooo' && (ooo ? (
              <div className="org-list">
                <div className="muted small">Automatickú odpoveď posiela priamo server Zimbra – funguje aj s vypnutým počítačom.</div>
                <Checkbox label="Posielať automatickú odpoveď" checked={ooo.enabled} onChange={(_, d) => setOoo({ ...ooo, enabled: !!d.checked })} />
                <div className="org-row">
                  <label className="sig-label" htmlFor="z-from">Od</label>
                  <input id="z-from" type="date" className="date-input" value={ooo.from ?? ''} onChange={(e) => setOoo({ ...ooo, from: e.target.value || undefined })} />
                  <label className="sig-label" htmlFor="z-to">Do</label>
                  <input id="z-to" type="date" className="date-input" value={ooo.until ?? ''} onChange={(e) => setOoo({ ...ooo, until: e.target.value || undefined })} />
                </div>
                <Field label="Text odpovede">
                  <Textarea resize="vertical" rows={6} value={ooo.reply} onChange={(_, d) => setOoo({ ...ooo, reply: d.value })} />
                </Field>
                <div><Button appearance="primary" disabled={busy} onClick={async () => { setBusy(true); const r = await api.zimbraSetOoo(id, ooo); result(r.ok, r.ok ? 'Neprítomnosť uložená na serveri.' : r.error); }}>Uložiť na server</Button></div>
              </div>
            ) : <Spinner size="tiny" />)}

            {tab === 'filters' && (rules ? (
              <div className="org-list">
                <div className="muted small">Filtre vykonáva server pri doručení – fungujú aj na telefóne a vo webovej Zimbre. Filtre s podmienkami, ktoré Universum eMail nevie upraviť, ostávajú nezmenené.</div>
                {rules.map((r, i) => {
                  const set = (p: Partial<ServerRule>) => setRules(rules.map((x, j) => (j === i ? { ...x, ...p } : x)));
                  return (
                    <div key={i} className="org-card">
                      <div className="org-row">
                        <Checkbox checked={r.active} onChange={(_, d) => set({ active: !!d.checked })} aria-label="Zapnutý" />
                        <Input className="grow" value={r.name} onChange={(_, d) => set({ name: d.value })} />
                        <Button size="small" appearance="subtle" disabled={i === 0} onClick={() => setRules([...rules.slice(0, i - 1), r, rules[i - 1]!, ...rules.slice(i + 1)])}>↑</Button>
                        <Button size="small" appearance="subtle" onClick={() => setRules(rules.filter((_, j) => j !== i))}>Odstrániť</Button>
                      </div>
                      {r.readOnly ? (
                        <div className="muted small">Tento filter obsahuje pokročilé podmienky – upravte ho vo webovej Zimbre. Tu ho môžete zapnúť, vypnúť, premenovať alebo odstrániť.</div>
                      ) : (
                        <>
                          <div className="org-sub">
                            <b>Ak</b>
                            <Select size="small" value={r.match} onChange={(_, d) => set({ match: d.value as 'all' | 'any' })}>
                              <option value="all">platia všetky podmienky</option>
                              <option value="any">platí ktorákoľvek</option>
                            </Select>
                          </div>
                          {r.conds.map((c, j) => (
                            <div key={j} className="org-row org-indent">
                              <Select size="small" value={c.field} onChange={(_, d) => set({ conds: r.conds.map((x, k) => (k === j ? { ...x, field: d.value as typeof c.field } : x)) })}>
                                {Object.entries(FIELD_LABEL).map(([v, l]) => <option key={v} value={v}>{l}</option>)}
                              </Select>
                              <Input size="small" className="grow" value={c.contains} onChange={(_, d) => set({ conds: r.conds.map((x, k) => (k === j ? { ...x, contains: d.value } : x)) })} />
                              <Button size="small" appearance="subtle" onClick={() => set({ conds: r.conds.filter((_, k) => k !== j) })}>×</Button>
                            </div>
                          ))}
                          <div className="org-indent"><Button size="small" onClick={() => set({ conds: [...r.conds, { field: 'from', contains: '' }] })}>+ Podmienka</Button></div>
                          <div className="org-sub"><b>Potom</b></div>
                          {r.actions.map((a, j) => (
                            <div key={j} className="org-row org-indent">
                              <Select size="small" value={a.type} onChange={(_, d) => set({ actions: r.actions.map((x, k) => (k === j ? { ...x, type: d.value as typeof a.type } : x)) })}>
                                {Object.entries(ACTION_LABEL).map(([v, l]) => <option key={v} value={v}>{l}</option>)}
                              </Select>
                              {a.type === 'move' && (
                                <Select size="small" value={a.folder ?? ''} onChange={(_, d) => set({ actions: r.actions.map((x, k) => (k === j ? { ...x, folder: d.value } : x)) })}>
                                  <option value="">(priečinok)</option>
                                  {props.folders.map((f) => <option key={f.path} value={f.path}>{f.path}</option>)}
                                </Select>
                              )}
                              {a.type === 'tag' && (
                                <Select size="small" value={a.tag ?? ''} onChange={(_, d) => set({ actions: r.actions.map((x, k) => (k === j ? { ...x, tag: d.value } : x)) })}>
                                  <option value="">(značka)</option>
                                  {(props.settings.categories ?? []).map((c) => <option key={c.id} value={c.name}>{c.name}</option>)}
                                </Select>
                              )}
                              <Button size="small" appearance="subtle" onClick={() => set({ actions: r.actions.filter((_, k) => k !== j) })}>×</Button>
                            </div>
                          ))}
                          <div className="org-indent"><Button size="small" onClick={() => set({ actions: [...r.actions, { type: 'move' }] })}>+ Akcia</Button></div>
                        </>
                      )}
                    </div>
                  );
                })}
                <div className="org-row">
                  <Button onClick={() => setRules([...rules, { name: 'Nový filter', active: true, match: 'all', conds: [{ field: 'from', contains: '' }], actions: [{ type: 'move' }] }])}>+ Nový filter</Button>
                  <Button appearance="primary" disabled={busy} onClick={async () => { setBusy(true); const r = await api.zimbraSetFilters(id, rules); result(r.ok, r.ok ? 'Filtre uložené na serveri.' : r.error); }}>Uložiť na server</Button>
                </div>
              </div>
            ) : <Spinner size="tiny" />)}

            {tab === 'shares' && (shares ? (
              <div className="org-list">
                <div className="muted small">Priečinky, kalendáre a adresáre, ktoré vám kolegovia nazdieľali. Po pripojení sa zobrazia medzi priečinkami, kalendármi a kontaktmi.</div>
                {shares.map((s, i) => (
                  <div key={i} className="org-row crypto-row">
                    <span className="grow"><b>{s.ownerName || s.ownerEmail}</b> – {s.folderPath} <span className="muted small">({VIEW[s.view] ?? s.view}, {s.rights.includes('w') ? 'úpravy' : 'čítanie'})</span></span>
                    {s.mounted ? (
                      <span className="muted small">pripojené</span>
                    ) : (
                      <Button size="small" onClick={async () => {
                        const name = `${s.ownerName || s.ownerEmail.split('@')[0]} – ${s.folderPath.split('/').pop()}`;
                        const r = await api.zimbraMount(id, s, name);
                        if (r.ok) setShares(shares.map((x, j) => (j === i ? { ...x, mounted: true } : x)));
                        result(r.ok, r.ok ? `Pripojené ako „${name}“. Obnovte priečinky alebo kalendár.` : r.error);
                      }}>Pripojiť</Button>
                    )}
                  </div>
                ))}
                {!shares.length && <div className="muted small">Nikto vám zatiaľ nič nenazdieľal.</div>}
              </div>
            ) : <Spinner size="tiny" />)}

            {tab === 'spam' && (wb ? (
              <div className="org-list">
                <div className="grid2">
                  <Field label="Dôveryhodní odosielatelia (nikdy spam)" hint="Jedna adresa alebo doména na riadok">
                    <Textarea resize="vertical" rows={10} value={wb.white} onChange={(_, d) => setWb({ ...wb, white: d.value })} />
                  </Field>
                  <Field label="Blokovaní odosielatelia" hint="Pošta od nich pôjde rovno do spamu">
                    <Textarea resize="vertical" rows={10} value={wb.black} onChange={(_, d) => setWb({ ...wb, black: d.value })} />
                  </Field>
                </div>
                <div>
                  <Button appearance="primary" disabled={busy} onClick={async () => {
                    setBusy(true);
                    const clean = (t: string) => [...new Set(t.split(/[\s,;]+/).map((x) => x.trim().toLowerCase()).filter(Boolean))];
                    const r = await api.zimbraSetWB(id, clean(wb.white), clean(wb.black));
                    result(r.ok, r.ok ? 'Zoznamy uložené na serveri.' : r.error);
                  }}>Uložiť na server</Button>
                </div>
              </div>
            ) : <Spinner size="tiny" />)}

            {tab === 'sync' && (
              <div className="org-list">
                <div className="muted small">
                  Prevezme podpisy, identity (ďalšie adresy v poli Od) a značky (ako kategórie s farbami) zo Zimbry. Vaše podpisy z programu Universum eMail, ktoré na serveri chýbajú, sa nahrajú do Zimbry.
                </div>
                <div><Button appearance="primary" disabled={busy} onClick={() => void syncFromServer()}>{busy ? 'Synchronizujem…' : 'Synchronizovať so Zimbrou'}</Button></div>
              </div>
            )}

            {tab === 'password' && (
              <div className="org-list" style={{ maxWidth: 420 }}>
                <Field label="Súčasné heslo"><Input type="password" value={pw.old} onChange={(_, d) => setPw({ ...pw, old: d.value })} /></Field>
                <Field label="Nové heslo"><Input type="password" value={pw.n1} onChange={(_, d) => setPw({ ...pw, n1: d.value })} /></Field>
                <Field label="Nové heslo znova"><Input type="password" value={pw.n2} onChange={(_, d) => setPw({ ...pw, n2: d.value })} /></Field>
                <div>
                  <Button appearance="primary" disabled={busy || !pw.old || !pw.n1} onClick={async () => {
                    if (pw.n1 !== pw.n2) return result(false, 'Nové heslá sa nezhodujú.');
                    setBusy(true);
                    const r = await api.zimbraChangePassword(id, pw.old, pw.n1);
                    result(r.ok, r.ok ? 'Heslo zmenené na serveri aj v programe Universum eMail.' : `${r.error}${r.hint ? ` ${r.hint}` : ''}`);
                    if (r.ok) setPw({ old: '', n1: '', n2: '' });
                  }}>Zmeniť heslo</Button>
                </div>
                <div className="muted small">Server môže vyžadovať minimálnu dĺžku a zložitosť hesla.</div>
              </div>
            )}
          </DialogContent>
          <DialogActions style={{ justifyContent: 'flex-end', alignSelf: 'flex-end' }}>
            <Button onClick={props.onClose}>Zavrieť</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
