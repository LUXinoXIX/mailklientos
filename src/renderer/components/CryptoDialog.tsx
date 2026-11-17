// Šifrovanie: vlastné kľúče PGP (vytvoriť, importovať, exportovať verejný), kľúče kontaktov,
// vlastný certifikát S/MIME (.p12/.pfx) a certifikáty kontaktov.
import { Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input, MessageBar, MessageBarBody, Tab, TabList, Textarea } from '@fluentui/react-components';
import { useEffect, useState } from 'react';
import type { ImapAccountPublic } from '../../core/account';
import type { CryptoSummary } from '../../core/api';

const fp = (f: string) => f.replace(/(.{4})/g, '$1 ').trim();
const date = (iso: string) => new Date(iso).toLocaleDateString('sk-SK');

export function CryptoDialog(props: { accounts: ImapAccountPublic[]; onClose: () => void }) {
  const api = window.mailApi;
  const [tab, setTab] = useState<'pgp' | 'smime'>('pgp');
  const [sum, setSum] = useState<CryptoSummary | null>(null);
  const [msg, setMsg] = useState<{ ok: boolean; text: string } | null>(null);
  const [busy, setBusy] = useState(false);
  const [gen, setGen] = useState({ name: props.accounts[0]?.name ?? '', email: props.accounts[0]?.email ?? '', pass: '' });
  const [pendingFile, setPendingFile] = useState<{ kind: 'pgp' | 'p12'; file: string } | null>(null);
  const [pass, setPass] = useState('');
  const [paste, setPaste] = useState('');

  const reload = () => void api?.cryptoSummary().then((r) => r.ok && setSum(r.data));
  useEffect(reload, [api]);
  const done = (ok: boolean, text: string) => {
    setMsg({ ok, text });
    setBusy(false);
    reload();
  };

  async function generate() {
    if (!api) return;
    if (!gen.email.includes('@')) return setMsg({ ok: false, text: 'Zadajte e-mailovú adresu.' });
    if (gen.pass.length < 6) return setMsg({ ok: false, text: 'Heslo ku kľúču musí mať aspoň 6 znakov.' });
    setBusy(true);
    const r = await api.pgpGenerate(gen.name, gen.email, gen.pass);
    done(r.ok, r.ok ? 'Kľúč vytvorený. Verejný kľúč pošlite tým, ktorí vám chcú písať šifrovane (tlačidlo Exportovať).' : r.error);
  }
  async function pick(kind: 'pgp' | 'p12' | 'cert') {
    if (!api) return;
    const r = await api.cryptoPick(kind);
    if (!r.ok || !r.data) return;
    if (kind === 'cert') {
      const c = await api.certImportFile(r.data);
      return done(c.ok, c.ok ? `Certifikát uložený: ${c.data.subject}` : c.error);
    }
    if (kind === 'p12') {
      setPendingFile({ kind, file: r.data });
      setPass('');
      return;
    }
    const x = await api.pgpImportFile(r.data);
    if (!x.ok && /heslo/i.test(x.error)) {
      setPendingFile({ kind, file: r.data });
      setPass('');
      return setMsg({ ok: false, text: x.error });
    }
    done(x.ok, x.ok ? `Importované: ${x.data.own} vlastných a ${x.data.pub} verejných kľúčov.` : x.error);
  }
  async function finishPending() {
    if (!api || !pendingFile) return;
    setBusy(true);
    if (pendingFile.kind === 'p12') {
      const r = await api.smimeImportFile(pendingFile.file, pass);
      done(r.ok, r.ok ? `Certifikát importovaný: ${r.data.subject} (${r.data.emails.join(', ')}), platný do ${date(r.data.validTo)}.` : r.error);
      if (r.ok) setPendingFile(null);
    } else {
      const r = await api.pgpImportFile(pendingFile.file, pass);
      done(r.ok, r.ok ? `Importované: ${r.data.own} vlastných a ${r.data.pub} verejných kľúčov.` : r.error);
      if (r.ok) setPendingFile(null);
    }
  }

  return (
    <Dialog open onOpenChange={(_, d) => !d.open && props.onClose()}>
      <DialogSurface style={{ maxWidth: 860, width: '92vw' }}>
        <DialogBody style={{ display: 'flex', flexDirection: 'column', maxHeight: '88vh' }}>
          <DialogTitle>Šifrovanie a podpis</DialogTitle>
          <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: 12, overflow: 'auto' }}>
            <TabList selectedValue={tab} onTabSelect={(_, d) => setTab(d.value as 'pgp' | 'smime')}>
              <Tab value="pgp">PGP (OpenPGP)</Tab>
              <Tab value="smime">S/MIME (certifikát)</Tab>
            </TabList>
            {msg && <MessageBar intent={msg.ok ? 'success' : 'error'}><MessageBarBody>{msg.text}</MessageBarBody></MessageBar>}
            {pendingFile && (
              <div className="org-card">
                <b>{pendingFile.kind === 'p12' ? 'Heslo k súboru .p12 / .pfx' : 'Heslo k súkromnému kľúču'}</b>
                <div className="org-row">
                  <Input type="password" className="grow" value={pass} onChange={(_, d) => setPass(d.value)} onKeyDown={(e) => e.key === 'Enter' && void finishPending()} autoFocus />
                  <Button appearance="primary" disabled={busy} onClick={() => void finishPending()}>Importovať</Button>
                  <Button onClick={() => setPendingFile(null)}>Zrušiť</Button>
                </div>
                <div className="muted small">Heslo sa uloží v zabezpečenom úložisku Windows, aby ste ho nemuseli zadávať pri každej správe.</div>
              </div>
            )}

            {tab === 'pgp' && sum && (
              <>
                <div className="folder-section">Moje kľúče</div>
                {sum.pgpOwn.map((k) => (
                  <div key={k.fingerprint} className="org-row crypto-row">
                    <span className="grow"><b>{k.userIds.join(', ')}</b><br /><span className="muted small mono">{fp(k.fingerprint)} · {date(k.created)}</span></span>
                    <Button size="small" onClick={async () => { const r = await api!.pgpExport(k.fingerprint); if (r.ok && r.data) setMsg({ ok: true, text: `Verejný kľúč uložený: ${r.data}` }); }}>Exportovať verejný</Button>
                    <Button size="small" appearance="subtle" onClick={async () => { await api!.cryptoRemove('pgpOwn', k.fingerprint); reload(); }}>Odstrániť</Button>
                  </div>
                ))}
                {!sum.pgpOwn.length && <div className="muted small">Zatiaľ nemáte vlastný kľúč. Vytvorte nový alebo importujte existujúci (napr. z Thunderbirdu alebo Kleopatry).</div>}
                <div className="org-card">
                  <b>Vytvoriť nový kľúč</b>
                  <div className="grid2">
                    <Field label="Meno"><Input value={gen.name} onChange={(_, d) => setGen({ ...gen, name: d.value })} /></Field>
                    <Field label="E-mail">
                      <Input value={gen.email} onChange={(_, d) => setGen({ ...gen, email: d.value.trim() })} list="crypto-emails" />
                    </Field>
                  </div>
                  <datalist id="crypto-emails">{props.accounts.map((a) => <option key={a.id} value={a.email} />)}</datalist>
                  <Field label="Heslo ku kľúču" hint="Chráni kľúč, keby sa dostal do nesprávnych rúk. Uloží sa v zabezpečenom úložisku Windows.">
                    <Input type="password" value={gen.pass} onChange={(_, d) => setGen({ ...gen, pass: d.value })} />
                  </Field>
                  <div className="org-row">
                    <Button appearance="primary" disabled={busy} onClick={() => void generate()}>{busy ? 'Vytváram…' : 'Vytvoriť kľúč'}</Button>
                    <Button onClick={() => void pick('pgp')}>Importovať kľúč zo súboru…</Button>
                  </div>
                </div>
                <div className="folder-section">Kľúče kontaktov</div>
                {sum.pgpContacts.map((k) => (
                  <div key={k.fingerprint} className="org-row crypto-row">
                    <span className="grow">{k.userIds.join(', ')}<br /><span className="muted small mono">{fp(k.fingerprint)}</span></span>
                    <Button size="small" appearance="subtle" onClick={async () => { await api!.cryptoRemove('pgpContacts', k.fingerprint); reload(); }}>Odstrániť</Button>
                  </div>
                ))}
                {!sum.pgpContacts.length && <div className="muted small">Verejný kľúč kontaktu importujte zo súboru .asc alebo ho vložte nižšie.</div>}
                <Field label="Vložiť verejný kľúč (text začínajúci -----BEGIN PGP PUBLIC KEY BLOCK-----)">
                  <Textarea resize="vertical" value={paste} onChange={(_, d) => setPaste(d.value)} />
                </Field>
                <div className="org-row">
                  <Button disabled={!paste.includes('BEGIN PGP')} onClick={async () => { const r = await api!.pgpImportText(paste); done(r.ok, r.ok ? `Importované kľúče: ${r.data.pub + r.data.own}.` : r.error); if (r.ok) setPaste(''); }}>Uložiť vložený kľúč</Button>
                  <Button onClick={() => void pick('pgp')}>Importovať zo súboru…</Button>
                </div>
              </>
            )}

            {tab === 'smime' && sum && (
              <>
                <div className="muted small">
                  S/MIME používa certifikát vydaný certifikačnou autoritou (napr. firemný). Funguje s Outlookom bez doplnkov.
                  Certifikáty kontaktov sa uložia automaticky, keď vám pošlú podpísanú správu.
                </div>
                <div className="folder-section">Môj certifikát</div>
                {sum.smimeOwn.map((c) => (
                  <div key={c.subject + c.validTo} className="org-row crypto-row">
                    <span className="grow"><b>{c.subject}</b> · {c.emails.join(', ')}<br /><span className="muted small">Vydal: {c.issuer} · platný do {date(c.validTo)}</span></span>
                    <Button size="small" appearance="subtle" onClick={async () => { await api!.cryptoRemove('smimeOwn', c.subject + c.validTo); reload(); }}>Odstrániť</Button>
                  </div>
                ))}
                <div className="org-row">
                  <Button appearance="primary" onClick={() => void pick('p12')}>Importovať certifikát (.p12 / .pfx)…</Button>
                </div>
                <div className="folder-section">Certifikáty kontaktov</div>
                {sum.smimeContacts.map((c) => (
                  <div key={c.subject + c.validTo} className="org-row crypto-row">
                    <span className="grow">{c.subject} · {c.emails.join(', ')}<br /><span className="muted small">Vydal: {c.issuer} · platný do {date(c.validTo)}</span></span>
                    <Button size="small" appearance="subtle" onClick={async () => { await api!.cryptoRemove('smimeContacts', c.subject + c.validTo); reload(); }}>Odstrániť</Button>
                  </div>
                ))}
                {!sum.smimeContacts.length && <div className="muted small">Zatiaľ žiadne.</div>}
                <div className="org-row">
                  <Button onClick={() => void pick('cert')}>Importovať certifikát kontaktu (.cer)…</Button>
                </div>
              </>
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
