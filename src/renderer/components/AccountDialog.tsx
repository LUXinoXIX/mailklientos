// Pridanie / úprava IMAP konta: šifrovanie, certifikáty (výnimka s odtlačkom),
// režim pre staré servery (Zimbra) a test pripojenia.
import { AccountIconPicker, type IconCfg } from './AccountIcon';
import { initials } from './AccountRail';
import {
  Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input,
  MessageBar, MessageBarBody, MessageBarTitle, Radio, RadioGroup, Select, Spinner, Switch, Textarea,
} from '@fluentui/react-components';
import { useEffect, useState } from 'react';
import {
  DEFAULT_PORTS, GMAIL_SERVERS, isGmail, strictPolicy, type AccountConfig, type AccountPublic, type ImapAccountPublic, type CertMode, type Security, type ServerConfig,
} from '../../core/account';
import type { AccountTestResult, GoogleGrant } from '../../core/api';
import type { AppSettings } from '../../core/settings';
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

/** Gmail: prihlásenie cez Google (OAuth) alebo heslom pre aplikácie + nastavenie klienta Google. */
function GmailAuth(props: {
  acc: AccountConfig;
  edit?: ImapAccountPublic;
  set: (p: Partial<AccountConfig>) => void;
  password: string;
  setPassword: (v: string) => void;
  grant: GoogleGrant | null;
  setGrant: (g: GoogleGrant | null) => void;
  google?: AppSettings['google'];
  onSaveGoogle?: (g: AppSettings['google']) => void;
  setError: (e: string | null) => void;
}) {
  const api = window.mailApi;
  const { acc } = props;
  const [signing, setSigning] = useState(false);
  const [showClient, setShowClient] = useState(!props.google?.clientId);
  const [cid, setCid] = useState(props.google?.clientId ?? '');
  const [csec, setCsec] = useState(props.google?.clientSecret ?? '');
  const open = (u: string) => void api?.openExternal(u);

  async function signIn() {
    if (!api) return;
    props.setError(null);
    setSigning(true);
    const r = await api.googleSignIn(acc.id, acc.email || undefined);
    setSigning(false);
    if (!r.ok) {
      props.setError(r.hint ? `${r.error} ${r.hint}` : r.error);
      return;
    }
    props.setGrant(r.data);
    const email = r.data.email || acc.email;
    props.set({ email, username: email, name: acc.name || r.data.name || email.split('@')[0] || email });
  }
  async function signOut() {
    if (!api) return;
    await api.googleSignOut(acc.id);
    props.setGrant(null);
  }
  const yes = (b: boolean) => (b ? '✓' : '✗');

  return (
    <fieldset className="server gmail-auth">
      <legend>Prihlásenie do Gmailu</legend>
      <RadioGroup value={acc.auth ?? 'password'} onChange={(_, d) => props.set({ auth: d.value as 'oauth' | 'password' })}>
        <Radio value="oauth" label="Prihlásiť sa cez Google (odporúčané) – pošta, kalendár aj kontakty" />
        <Radio value="password" label="Heslo pre aplikácie – len pošta (kalendár a kontakty Google takto nefungujú)" />
      </RadioGroup>

      {(acc.auth ?? 'password') === 'oauth' ? (
        <div className="col gap">
          {props.grant ? (
            <MessageBar intent="success">
              <MessageBarBody>
                Prihlásené ako <b>{props.grant.email || acc.email}</b> · pošta {yes(props.grant.mail)} · kalendár {yes(props.grant.calendar)} · kontakty {yes(props.grant.contacts)}
                {(!props.grant.calendar || !props.grant.contacts) && ' – chýbajúce povolíš opätovným prihlásením.'}
              </MessageBarBody>
            </MessageBar>
          ) : (
            <div className="muted small">
              Otvorí sa stránka Google v prehliadači. Prihlás sa a povoľ prístup – program heslo Google nikdy nevidí.
            </div>
          )}
          <div className="row gap">
            <Button appearance="primary" disabled={signing || !props.google?.clientId} onClick={signIn}>
              {props.grant ? 'Prihlásiť sa znova' : 'Prihlásiť sa cez Google'}
            </Button>
            {signing && (
              <>
                <Spinner size="tiny" label="Dokonči prihlásenie v prehliadači…" />
                <Button size="small" onClick={() => void api?.googleCancel()}>Zrušiť</Button>
              </>
            )}
            {props.grant && !signing && <Button appearance="subtle" onClick={signOut}>Odhlásiť z Google</Button>}
            <div className="grow" />
            <Button appearance="subtle" size="small" onClick={() => setShowClient((v) => !v)}>
              {showClient ? 'Skryť nastavenie klienta Google' : 'Nastavenie klienta Google…'}
            </Button>
          </div>
          {!props.google?.clientId && (
            <MessageBar intent="warning">
              <MessageBarBody>Najprv treba raz vložiť ID klienta Google (postup nižšie). Platí pre všetky kontá Gmail.</MessageBarBody>
            </MessageBar>
          )}
          {showClient && (
            <div className="gmail-client col gap">
              <ol className="small steps">
                <li>
                  Otvor <a href="#" onClick={(e) => { e.preventDefault(); open('https://console.cloud.google.com/projectcreate'); }}>Google Cloud</a> a vytvor projekt (napr. „Universum eMail“).
                </li>
                <li>
                  V <a href="#" onClick={(e) => { e.preventDefault(); open('https://console.cloud.google.com/apis/library'); }}>knižnici rozhraní</a> zapni <b>Gmail API</b>, <b>CalDAV API</b> a <b>CardDAV API</b>.
                </li>
                <li>
                  <a href="#" onClick={(e) => { e.preventDefault(); open('https://console.cloud.google.com/auth/branding'); }}>Google Auth Platform</a>: vyplň názov a e-mail, typ používateľov <b>Externí</b> (pri Google Workspace môže byť Interní). V časti Publikum pridaj svoj Gmail medzi testovacích používateľov,
                  alebo daj <b>Zverejniť aplikáciu</b> – inak prihlásenie vyprší po 7 dňoch.
                </li>
                <li>
                  <a href="#" onClick={(e) => { e.preventDefault(); open('https://console.cloud.google.com/auth/clients'); }}>Klienti</a> → Vytvoriť klienta → typ <b>Počítačová aplikácia</b> (Desktop app). Skopíruj ID klienta a tajný kľúč sem.
                </li>
              </ol>
              <div className="muted small">
                Pri prihlásení Google ukáže „Google túto aplikáciu neoveril“ – je to tvoja vlastná aplikácia, pokračuj cez Rozšírené → Prejsť do aplikácie.
              </div>
              <div className="grid2">
                <Field label="ID klienta">
                  <Input value={cid} onChange={(_, d) => setCid(d.value.trim())} placeholder="…apps.googleusercontent.com" />
                </Field>
                <Field label="Tajný kľúč klienta (client secret)">
                  <Input type="password" value={csec} onChange={(_, d) => setCsec(d.value.trim())} placeholder="GOCSPX-…" />
                </Field>
              </div>
              <div className="row gap">
                <Button
                  disabled={!cid.endsWith('.apps.googleusercontent.com')}
                  onClick={() => {
                    props.onSaveGoogle?.({ clientId: cid, clientSecret: csec || undefined });
                    setShowClient(false);
                  }}
                >
                  Uložiť klienta Google
                </Button>
                {cid && !cid.endsWith('.apps.googleusercontent.com') && <span className="muted small">ID klienta končí na .apps.googleusercontent.com</span>}
              </div>
            </div>
          )}
        </div>
      ) : (
        <div className="col gap">
          <ol className="small steps">
            <li>
              V Google konte zapni <a href="#" onClick={(e) => { e.preventDefault(); open('https://myaccount.google.com/signinoptions/twosv'); }}>dvojstupňové overenie</a> (ak ešte nie je).
            </li>
            <li>
              Otvor <a href="#" onClick={(e) => { e.preventDefault(); open('https://myaccount.google.com/apppasswords'); }}>Heslá aplikácií</a>, zadaj názov „Universum eMail“ a daj Vytvoriť.
            </li>
            <li>Skopíruj 16 písmen sem (medzery nevadia). Pri firemnom Google Workspace to musí povoliť správca.</li>
          </ol>
          <Field label="Heslo pre aplikácie" required={!props.edit?.hasPassword || props.edit.auth === 'oauth'} hint={props.edit?.hasPassword && props.edit.auth !== 'oauth' ? 'Nechaj prázdne, ak sa nemení.' : undefined}>
            <Input type="password" value={props.password} onChange={(_, d) => props.setPassword(d.value.replace(/\s+/g, ''))} placeholder="abcd efgh ijkl mnop" />
          </Field>
        </div>
      )}
    </fieldset>
  );
}

export function AccountDialog(props: {
  open: boolean;
  edit?: ImapAccountPublic;
  /** nové konto: predvolené nastavenia (Gmail) */
  preset?: 'gmail';
  google?: AppSettings['google'];
  onSaveGoogle?: (g: AppSettings['google']) => void;
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
          ...(props.preset === 'gmail'
            ? { provider: 'gmail' as const, auth: props.google?.clientId ? ('oauth' as const) : ('password' as const), imap: { ...GMAIL_SERVERS.imap }, smtp: { ...GMAIL_SERVERS.smtp } }
            : {}),
        };
  const [acc, setAcc] = useState<AccountConfig>(init);
  const [aliases, setAliases] = useState(() => (acc.identities ?? []).map((x) => (x.name ? `${x.name} <${x.email}>` : x.email)).join('\n'));
  const [password, setPassword] = useState('');
  const [busy, setBusy] = useState<string | null>(null);
  const [test, setTest] = useState<AccountTestResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [log, setLog] = useState<string[]>([]);
  const gmail = isGmail(acc);
  const [grant, setGrant] = useState<GoogleGrant | null>(null);
  useEffect(() => {
    if (gmail && props.edit && api) void api.googleGrant(props.edit.id).then((r) => r.ok && setGrant(r.data));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const set = (patch: Partial<AccountConfig>) => setAcc((a) => ({ ...a, ...patch }));
  const onEmail = (email: string) => {
    if (gmail) {
      setAcc((a) => ({ ...a, email, username: email, name: !a.name || a.name === a.email.split('@')[0] ? email.split('@')[0] ?? '' : a.name }));
      return;
    }
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
      setError(gmail ? 'Vyplň adresu Gmailu.' : 'Vyplň e-mail a servery IMAP a SMTP.');
      return;
    }
    if (gmail && acc.auth === 'oauth') {
      if (!grant) {
        setError('Prihlás sa cez Google (tlačidlo „Prihlásiť sa cez Google“).');
        return;
      }
    } else if ((!props.edit || (gmail && props.edit.auth === 'oauth')) && !password) {
      setError(gmail ? 'Zadaj heslo pre aplikácie.' : 'Zadaj heslo.');
      return;
    }
    // prechod z prihlásenia cez Google na heslo: token zrušiť
    if (gmail && acc.auth !== 'oauth' && grant) {
      await api.googleSignOut(acc.id);
      setGrant(null);
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
      <DialogSurface className="account-dialog" style={{ maxWidth: 1000, width: "94vw" }}>
        <DialogBody>
          <DialogTitle>{props.edit ? `Nastavenia konta ${props.edit.name}` : gmail ? 'Pridať konto Gmail' : 'Pridať konto IMAP'}</DialogTitle>
          <DialogContent className="dialog-content">
            {!api && (
              <MessageBar intent="error"><MessageBarBody>Aplikácia nebeží v Electrone (chýba most do hlavného procesu).</MessageBarBody></MessageBar>
            )}
            <div className="grid2">
              <Field label="E-mail" required>
                <Input value={acc.email} onChange={(_, d) => onEmail(d.value.trim())} placeholder={gmail ? 'meno@gmail.com' : 'meno@firma.sk'} />
              </Field>
              <Field label="Zobrazované meno">
                <Input value={acc.name} onChange={(_, d) => set({ name: d.value })} />
              </Field>
              <Field label="Ikonka v páse kont">
                <AccountIconPicker
                  value={acc}
                  color={acc.color}
                  fallback={initials(acc.name || acc.email || '?')}
                  onChange={(v) => set(v)}
                />
              </Field>
              <Field label="Štítok v zbalenom páse kont" hint={`Max. 8 znakov, napr. FIRMA, PROJEKT. Prázdne = ${gmail ? 'GMAIL' : 'ZIMBRA alebo IMAP'}.`}>
                <Input value={acc.short ?? ''} maxLength={8} placeholder={gmail ? 'GMAIL' : 'ZIMBRA'} onChange={(_, d) => set({ short: d.value.toUpperCase() || undefined })} />
              </Field>
              {!gmail && <><Field label="Používateľské meno">
                <Input value={acc.username} onChange={(_, d) => set({ username: d.value.trim() })} />
              </Field>
              <Field label="Heslo" required={!props.edit} hint={props.edit?.hasPassword ? 'Nechaj prázdne, ak sa nemení.' : undefined}>
                <Input type="password" value={password} onChange={(_, d) => setPassword(d.value)} />
              </Field></>}
            </div>
            {gmail && (
              <GmailAuth
                acc={acc}
                edit={props.edit}
                set={set}
                password={password}
                setPassword={setPassword}
                grant={grant}
                setGrant={setGrant}
                google={props.google}
                onSaveGoogle={props.onSaveGoogle}
                setError={setError}
              />
            )}
            <div className="row gap colors">
              <span className="muted">Farba konta</span>
              {ACCOUNT_COLORS.map((c) => (
                <button key={c} className={`swatch ${acc.color === c ? 'on' : ''}`} style={{ background: c }} onClick={() => set({ color: c })} aria-label={c} />
              ))}
              <label className={`swatch custom ${ACCOUNT_COLORS.includes(acc.color) ? '' : 'on'}`} title="Vlastná farba" style={{ background: ACCOUNT_COLORS.includes(acc.color) ? undefined : acc.color }}>
                <input type="color" value={acc.color} onChange={(e) => set({ color: e.target.value })} aria-label="Vlastná farba konta" />
              </label>
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
            {!gmail && (
              <>
                <Field label="Adresa kalendára a kontaktov (CalDAV/CardDAV)" hint="Nechajte prázdne – zistí sa automaticky. Zimbra: https://server/dav/meno@firma.sk/">
                  <Input value={acc.davUrl ?? ''} onChange={(_, d) => set({ davUrl: d.value.trim() || undefined })} placeholder="automaticky" />
                </Field>
                <ServerSection proto="imap" value={acc.imap} onChange={(imap) => setAcc((a) => ({ ...a, imap }))} probe={test?.imap} onRetest={(imap) => runTest({ ...acc, imap })} />
                <ServerSection proto="smtp" value={acc.smtp} onChange={(smtp) => setAcc((a) => ({ ...a, smtp }))} probe={test?.smtp} onRetest={(smtp) => runTest({ ...acc, smtp })} />
              </>
            )}
            {gmail && <div className="muted small">Servery Gmailu: imap.gmail.com:993 a smtp.gmail.com:465 (SSL/TLS) – nastavené automaticky.</div>}
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
            {!gmail && <Button disabled={!!busy || (!acc.imap.host && !acc.email.includes('@'))} onClick={runAutodetect}>Zistiť automaticky</Button>}
            <Button disabled={!!busy} onClick={() => runTest()}>Otestovať</Button>
            <Button appearance="primary" disabled={!!busy} onClick={save}>Uložiť</Button>
            <Button disabled={!!busy} onClick={props.onClose}>Zrušiť</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
