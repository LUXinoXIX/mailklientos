// Pridanie / úprava IMAP konta: šifrovanie, certifikáty (výnimka s odtlačkom),
// režim pre staré servery (Zimbra) a test pripojenia.
import {
  Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input,
  MessageBar, MessageBarBody, MessageBarTitle, Select, Spinner, Switch, Textarea,
} from '@fluentui/react-components';
import { useState } from 'react';
import {
  DEFAULT_PORTS, strictPolicy, type AccountConfig, type AccountPublic, type ImapAccountPublic, type CertMode, type Security, type ServerConfig,
} from '../../core/account';
import type { AccountTestResult } from '../../core/api';
import type { ProbeResult, Protocol } from '../../core/tls-types';
import { ACCOUNT_COLORS } from '../theme';

const emptyServer = (proto: Protocol, host = ''): ServerConfig => ({
  host,
  port: DEFAULT_PORTS[proto].ssl,
  security: 'ssl',
  tls: strictPolicy(),
});

const fmtFp = (fp: string) => fp.toUpperCase();

function ProbeView(props: {
  r: ProbeResult;
  onTrust: () => void;
  onIgnore: () => void;
  onLegacy: () => void;
}) {
  const { r } = props;
  if (r.ok)
    return (
      <MessageBar intent={r.hint ? 'warning' : 'success'}>
        <MessageBarBody>
          {r.message}
          {r.hint && <div>{r.hint}</div>}
        </MessageBarBody>
      </MessageBar>
    );
  if (r.needsTrust && r.cert)
    return (
      <MessageBar intent="warning" layout="multiline">
        <MessageBarBody>
          <MessageBarTitle>Dôverovať tomuto certifikátu?</MessageBarTitle>
          {r.message}
          <div className="cert">
            <div><span className="muted">Vydané pre:</span> {r.cert.subject}</div>
            <div><span className="muted">Vydal:</span> {r.cert.selfSigned ? 'sám sebe (self-signed)' : r.cert.issuer}</div>
            <div><span className="muted">Platnosť:</span> {r.cert.validFrom} – {r.cert.validTo}</div>
            {r.cert.altNames && <div><span className="muted">Názvy:</span> {r.cert.altNames}</div>}
            <div><span className="muted">SHA-256:</span> <code className="fp">{fmtFp(r.cert.fingerprint256)}</code></div>
            {r.protocol && <div><span className="muted">Protokol:</span> {r.protocol}{r.cipher ? `, ${r.cipher}` : ''}</div>}
          </div>
          <div className="muted small">Odtlačok si môžeš overiť u správcu servera. Výnimka platí len pre tento server v tomto konte; pri zmene certifikátu sa aplikácia opýta znova.</div>
          <div className="row gap">
            <Button appearance="primary" size="small" onClick={props.onTrust}>Dôverovať tomuto certifikátu</Button>
            <Button size="small" onClick={props.onIgnore}>Ignorovať všetky chyby certifikátu</Button>
          </div>
        </MessageBarBody>
      </MessageBar>
    );
  return (
    <MessageBar intent="error" layout="multiline">
      <MessageBarBody>
        <MessageBarTitle>{r.message}</MessageBarTitle>
        {r.hint}
        {r.suggestLegacy && (
          <div className="row gap">
            <Button size="small" appearance="primary" onClick={props.onLegacy}>Zapnúť kompatibilitu so starými servermi</Button>
          </div>
        )}
      </MessageBarBody>
    </MessageBar>
  );
}

function ServerSection(props: {
  proto: Protocol;
  value: ServerConfig;
  onChange: (v: ServerConfig) => void;
  probe?: ProbeResult;
  onRetest: (v: ServerConfig) => void;
}) {
  const v = props.value;
  const set = (patch: Partial<ServerConfig>) => props.onChange({ ...v, ...patch });
  const setTls = (patch: Partial<ServerConfig['tls']>) => {
    const next = { ...v, tls: { ...v.tls, ...patch } };
    props.onChange(next);
    return next;
  };
  const setSecurity = (security: Security) => {
    const wasDefault = Object.values(DEFAULT_PORTS[props.proto]).includes(v.port);
    set({ security, port: wasDefault ? DEFAULT_PORTS[props.proto][security] : v.port });
  };
  const label = props.proto === 'imap' ? 'Prichádzajúca pošta (IMAP)' : 'Odchádzajúca pošta (SMTP)';
  return (
    <fieldset className="server">
      <legend>{label}</legend>
      <div className="grid3">
        <Field label="Server">
          <Input value={v.host} onChange={(_, d) => set({ host: d.value.trim() })} placeholder="mail.firma.sk" />
        </Field>
        <Field label="Šifrovanie">
          <Select value={v.security} onChange={(_, d) => setSecurity(d.value as Security)}>
            <option value="ssl">SSL/TLS</option>
            <option value="starttls">STARTTLS</option>
            <option value="none">Žiadne (nebezpečné)</option>
          </Select>
        </Field>
        <Field label="Port">
          <Input type="number" value={String(v.port)} onChange={(_, d) => set({ port: Number(d.value) || 0 })} />
        </Field>
      </div>
      {v.security !== 'none' && (
        <div className="grid2">
          <Field label="Certifikát servera">
            <Select value={v.tls.certMode} onChange={(_, d) => setTls({ certMode: d.value as CertMode })}>
              <option value="strict">Overovať (odporúčané)</option>
              <option value="pinned" disabled={!v.tls.pinnedSha256}>Dôverovať uloženému certifikátu</option>
              <option value="ignore">Ignorovať chyby certifikátu (nebezpečné)</option>
            </Select>
          </Field>
          <Field label="Staré servery" hint="Napr. staré verzie Zimbry: TLS 1.0/1.1, slabšie šifry a kľúče.">
            <Switch
              checked={v.tls.legacy}
              onChange={(_, d) => setTls({ legacy: d.checked })}
              label="Kompatibilita so starými servermi"
            />
          </Field>
        </div>
      )}
      {v.tls.certMode === 'pinned' && v.tls.pinnedSha256 && (
        <div className="muted small">
          Uložená výnimka, SHA-256: <code className="fp">{fmtFp(v.tls.pinnedSha256)}</code>{' '}
          <Button size="small" appearance="transparent" onClick={() => setTls({ certMode: 'strict', pinnedSha256: undefined })}>
            Zrušiť výnimku
          </Button>
        </div>
      )}
      {v.tls.certMode === 'ignore' && (
        <MessageBar intent="warning">
          <MessageBarBody>Certifikát sa neoveruje: spojenie je šifrované, ale nie je chránené pred podvrhnutím servera.</MessageBarBody>
        </MessageBar>
      )}
      {props.probe && (
        <ProbeView
          r={props.probe}
          onTrust={() => props.onRetest(setTls({ certMode: 'pinned', pinnedSha256: props.probe!.cert!.fingerprint256 }))}
          onIgnore={() => props.onRetest(setTls({ certMode: 'ignore' }))}
          onLegacy={() => props.onRetest(setTls({ legacy: true }))}
        />
      )}
    </fieldset>
  );
}

export function AccountDialog(props: {
  open: boolean;
  edit?: ImapAccountPublic;
  usedColors: string[];
  onClose: () => void;
  onSaved: (a: AccountPublic) => void;
  onDeleted: (id: string) => void;
}) {
  const api = window.mailApi;
  const init = (): AccountConfig =>
    props.edit
      ? (({ hasPassword: _h, ...rest }) => rest)(props.edit)
      : {
          id: crypto.randomUUID(),
          kind: 'imap',
          name: '',
          email: '',
          color: ACCOUNT_COLORS.find((c) => !props.usedColors.includes(c)) ?? ACCOUNT_COLORS[0]!,
          username: '',
          imap: emptyServer('imap'),
          smtp: emptyServer('smtp'),
        };
  const [acc, setAcc] = useState<AccountConfig>(init);
  const [aliases, setAliases] = useState(() => (acc.identities ?? []).map((x) => (x.name ? `${x.name} <${x.email}>` : x.email)).join('\n'));
  const [password, setPassword] = useState('');
  const [busy, setBusy] = useState<string | null>(null);
  const [test, setTest] = useState<AccountTestResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [log, setLog] = useState<string[]>([]);

  const set = (patch: Partial<AccountConfig>) => setAcc((a) => ({ ...a, ...patch }));
  const onEmail = (email: string) => {
    const domain = email.split('@')[1] ?? '';
    setAcc((a) => ({
      ...a,
      email,
      username: !a.username || a.username === a.email ? email : a.username,
      name: !a.name || a.name === a.email.split('@')[0] ? email.split('@')[0] ?? '' : a.name,
      imap: !a.imap.host || a.imap.host === `mail.${a.email.split('@')[1] ?? ''}` ? { ...a.imap, host: domain ? `mail.${domain}` : '' } : a.imap,
      smtp: !a.smtp.host || a.smtp.host === `mail.${a.email.split('@')[1] ?? ''}` ? { ...a.smtp, host: domain ? `mail.${domain}` : '' } : a.smtp,
    }));
  };

  async function runTest(next: AccountConfig = acc) {
    if (!api) return;
    setBusy('Testujem pripojenie…');
    setError(null);
    const r = await api.testAccount(next, password || undefined);
    setBusy(null);
    if (r.ok) setTest(r.data);
    else setError(r.error);
  }

  async function runAutodetect() {
    if (!api || (!acc.imap.host && !acc.email.includes('@'))) return;
    setBusy('Zisťujem nastavenia servera…');
    setError(null);
    // 1) Autoconfig / databáza Thunderbirdu (ISPDB) podľa domény e-mailu
    const ac = acc.email.includes('@') ? await api.autoconfig(acc.email) : null;
    if (ac?.ok && ac.data && (ac.data.imap || ac.data.smtp)) {
      const d = ac.data;
      const local = acc.email.split('@')[0] ?? acc.email;
      const user = d.imap?.username === 'local' ? local : acc.email;
      const next: AccountConfig = {
        ...acc,
        username: !acc.username || acc.username === acc.email || acc.username === local ? user : acc.username,
        imap: d.imap ? { ...acc.imap, host: d.imap.host, port: d.imap.port, security: d.imap.security } : acc.imap,
        smtp: d.smtp ? { ...acc.smtp, host: d.smtp.host, port: d.smtp.port, security: d.smtp.security } : acc.smtp,
      };
      setAcc(next);
      setLog([
        `Nastavenie nájdené: ${d.source}${d.name ? ` (${d.name})` : ''}`,
        d.imap ? `IMAP: ${d.imap.host}:${d.imap.port} ${d.imap.security.toUpperCase()}` : 'IMAP: nenájdené',
        d.smtp ? `SMTP: ${d.smtp.host}:${d.smtp.port} ${d.smtp.security.toUpperCase()}` : 'SMTP: nenájdené',
      ]);
      setBusy(null);
      await runTest(next);
      return;
    }
    if (!acc.imap.host) {
      setBusy(null);
      setError('Nastavenie sa nenašlo. Doplň server IMAP (napr. mail.firma.sk) a skús znova.');
      return;
    }
    // 2) skúšanie portov a šifrovania na zadanom serveri
    const [im, sm] = await Promise.all([api.autodetect(acc.imap.host, 'imap'), api.autodetect(acc.smtp.host || acc.imap.host, 'smtp')]);
    const lines: string[] = [];
    const next = { ...acc };
    if (im.ok) {
      lines.push(`IMAP: ${im.data.tried.join(' · ')}`);
      if (im.data.config) next.imap = im.data.config;
    } else lines.push(`IMAP: ${im.error}`);
    if (sm.ok) {
      lines.push(`SMTP: ${sm.data.tried.join(' · ')}`);
      if (sm.data.config) next.smtp = sm.data.config;
    } else lines.push(`SMTP: ${sm.error}`);
    setAcc(next);
    setLog(lines);
    setBusy(null);
    await runTest(next);
  }

  async function save() {
    if (!api) return;
    if (!acc.email.includes('@') || !acc.imap.host || !acc.smtp.host) {
      setError('Vyplň e-mail a servery IMAP a SMTP.');
      return;
    }
    if (!props.edit && !password) {
      setError('Zadaj heslo.');
      return;
    }
    setBusy('Ukladám…');
    const r = await api.saveAccount({ ...acc, name: acc.name || acc.email }, password || undefined);
    setBusy(null);
    if (r.ok) props.onSaved(r.data);
    else setError(r.error);
  }

  async function remove() {
    if (!api || !props.edit) return;
    if (!window.confirm(`Odstrániť konto ${props.edit.name}? Správy na serveri zostanú.`)) return;
    const r = await api.deleteAccount(props.edit.id);
    if (r.ok) props.onDeleted(props.edit.id);
    else setError(r.error);
  }

  return (
    <Dialog open={props.open} onOpenChange={(_, d) => !d.open && props.onClose()}>
      <DialogSurface className="account-dialog">
        <DialogBody>
          <DialogTitle>{props.edit ? `Nastavenia konta ${props.edit.name}` : 'Pridať konto IMAP'}</DialogTitle>
          <DialogContent className="dialog-content">
            {!api && (
              <MessageBar intent="error"><MessageBarBody>Aplikácia nebeží v Electrone (chýba most do hlavného procesu).</MessageBarBody></MessageBar>
            )}
            <div className="grid2">
              <Field label="E-mail" required>
                <Input value={acc.email} onChange={(_, d) => onEmail(d.value.trim())} placeholder="meno@firma.sk" />
              </Field>
              <Field label="Zobrazované meno">
                <Input value={acc.name} onChange={(_, d) => set({ name: d.value })} />
              </Field>
              <Field label="Používateľské meno">
                <Input value={acc.username} onChange={(_, d) => set({ username: d.value.trim() })} />
              </Field>
              <Field label="Heslo" required={!props.edit} hint={props.edit?.hasPassword ? 'Nechaj prázdne, ak sa nemení.' : undefined}>
                <Input type="password" value={password} onChange={(_, d) => setPassword(d.value)} />
              </Field>
            </div>
            <div className="row gap colors">
              <span className="muted">Farba konta</span>
              {ACCOUNT_COLORS.map((c) => (
                <button key={c} className={`swatch ${acc.color === c ? 'on' : ''}`} style={{ background: c }} onClick={() => set({ color: c })} aria-label={c} />
              ))}
            </div>
            <Field label="Ďalšie adresy (aliasy) pre pole Od" hint="Jedna na riadok, napr. Podpora <podpora@firma.sk>. Server musí odosielanie z nich povoliť.">
              <Textarea
                resize="vertical"
                value={aliases}
                onChange={(_, d) => {
                  setAliases(d.value);
                  const ids = d.value
                    .split('\n')
                    .map((l) => l.trim())
                    .filter(Boolean)
                    .map((l) => {
                      const m = /^(.*?)\s*<([^>]+)>$/.exec(l);
                      return m ? { name: m[1]!.replace(/^"|"$/g, '').trim(), email: m[2]!.trim() } : { name: '', email: l };
                    })
                    .filter((x) => x.email.includes('@'));
                  set({ identities: ids });
                }}
              />
            </Field>
            <ServerSection proto="imap" value={acc.imap} onChange={(imap) => setAcc((a) => ({ ...a, imap }))} probe={test?.imap} onRetest={(imap) => runTest({ ...acc, imap })} />
            <ServerSection proto="smtp" value={acc.smtp} onChange={(smtp) => setAcc((a) => ({ ...a, smtp }))} probe={test?.smtp} onRetest={(smtp) => runTest({ ...acc, smtp })} />
            {test?.login && (
              <MessageBar intent={test.login.ok ? 'success' : 'error'}>
                <MessageBarBody>
                  {test.login.message} {test.login.hint}
                </MessageBarBody>
              </MessageBar>
            )}
            {log.length > 0 && <div className="muted small log">{log.map((l, i) => <div key={i}>{l}</div>)}</div>}
            {error && (
              <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar>
            )}
          </DialogContent>
          <DialogActions fluid className="dialog-actions">
            {props.edit && <Button appearance="subtle" onClick={remove}>Odstrániť konto</Button>}
            <div className="grow" />
            {busy && <Spinner size="tiny" label={busy} />}
            <Button disabled={!!busy || (!acc.imap.host && !acc.email.includes('@'))} onClick={runAutodetect}>Zistiť automaticky</Button>
            <Button disabled={!!busy} onClick={() => runTest()}>Otestovať</Button>
            <Button appearance="primary" disabled={!!busy} onClick={save}>Uložiť</Button>
            <Button disabled={!!busy} onClick={props.onClose}>Zrušiť</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
