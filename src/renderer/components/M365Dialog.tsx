// Výber typu nového konta a dialóg pre konto Microsoft 365 (vložený Outlook na webe).
import { AccountIconPicker, type IconCfg } from './AccountIcon';
import { initials } from './AccountRail';
import {
  Button, Checkbox, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input, MessageBar,
  MessageBarBody, Radio, RadioGroup,
} from '@fluentui/react-components';
import { useState } from 'react';
import { accountPartition, type M365Account } from '../../core/account';
import { ACCOUNT_COLORS, accountColor } from '../theme';
import { MailIcon, TeamsIcon, WebmailIcon } from '../icons';

export function AddAccountChooser(props: { onClose: () => void; onImap: () => void; onM365: (host: 'office' | 'live' | 'web') => void }) {
  return (
    <Dialog open onOpenChange={(_, d) => !d.open && props.onClose()}>
      <DialogSurface className="chooser-dialog">
        <DialogBody>
          <DialogTitle>Pridať konto</DialogTitle>
          <DialogContent className="chooser">
            <button className="chooser-card" onClick={props.onImap}>
              <MailIcon />
              <span>
                <b>E-mail (IMAP / SMTP)</b>
                <span className="muted">Zimbra, Gmail, Seznam, firemný server… Plnohodnotný poštový klient v aplikácii.</span>
              </span>
            </button>
            <button className="chooser-card" onClick={() => props.onM365('office')}>
              <TeamsIcon />
              <span>
                <b>Outlook – firemný (Microsoft 365)</b>
                <span className="muted">Pracovné alebo školské konto. Vložený Outlook na webe a firemné aplikácie: Teams, Planner, SharePoint, OneDrive pre prácu, Word, Excel, Copilot…</span>
              </span>
            </button>
            <button className="chooser-card" onClick={() => props.onM365('live')}>
              <MailIcon />
              <span>
                <b>Outlook – osobný (Outlook.com, Hotmail, Live)</b>
                <span className="muted">Osobné konto Microsoft. Vložený Outlook.com a osobné aplikácie: OneDrive, Word, Excel, OneNote, Copilot…</span>
              </span>
            </button>
            <button className="chooser-card" onClick={() => props.onM365('web')}>
              <WebmailIcon />
              <span>
                <b>Webmail (vložená stránka)</b>
                <span className="muted">Ľubovoľná webová pošta – napr. mail.firma.sk, Zimbra, Roundcube, SOGo. Zobrazí sa ako stránka v okne s vlastným prihlásením, ako Outlook na webe.</span>
              </span>
            </button>
          </DialogContent>
          <DialogActions>
            <Button onClick={props.onClose}>Zrušiť</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

export function M365Dialog(props: {
  edit?: M365Account;
  /** typ nového konta zvolený pri pridávaní */
  host?: 'office' | 'live' | 'web';
  usedColors: string[];
  onClose: () => void;
  onSaved: (a: M365Account) => void;
  onDeleted: () => void;
}) {
  const e = props.edit;
  const [name, setName] = useState(e?.name ?? (props.host === 'web' ? 'Webmail' : props.host === 'live' ? 'Outlook osobný' : 'Outlook firemný'));
  const [webUrl, setWebUrl] = useState(e?.webUrl ?? '');
  const [ignoreCert, setIgnoreCert] = useState(!!e?.webIgnoreCert);
  const [railOff, setRailOff] = useState(!!e?.railOff);
  const [webUser, setWebUser] = useState(e?.webUser ?? '');
  const [webPass, setWebPass] = useState('');
  const [autoLogin, setAutoLogin] = useState(!!e?.webAutoLogin);
  const hasStored = !!(e as { webHasPassword?: boolean } | undefined)?.webHasPassword;
  const [email, setEmail] = useState(e?.email ?? '');
  const [short, setShort] = useState(e?.short ?? '');
  const [icon, setIcon] = useState<IconCfg>({ iconKind: e?.iconKind, iconText: e?.iconText, iconPreset: e?.iconPreset, iconImage: e?.iconImage });
  const [host, setHost] = useState<'office' | 'live' | 'web'>(e?.host ?? props.host ?? 'office');
  const web = host === 'web';
  const [color, setColor] = useState(accountColor(e?.color ?? ACCOUNT_COLORS.find((c) => !props.usedColors.includes(c)) ?? ACCOUNT_COLORS[1]!));
  const [error, setError] = useState<string | null>(null);
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [busy, setBusy] = useState(false);

  async function save() {
    if (!window.mailApi) return;
    if (!name.trim()) return setError('Zadaj názov konta.');
    let url = webUrl.trim();
    if (web) {
      if (url && !/^https?:\/\//i.test(url)) url = `https://${url}`;
      try {
        if (!url || !/^https?:$/.test(new URL(url).protocol)) throw new Error();
      } catch {
        return setError('Zadaj adresu webmailu, napr. https://mail.firma.sk');
      }
    }
    const acc: M365Account = {
      id: e?.id ?? `m365-${Date.now().toString(36)}`,
      kind: 'm365',
      name: name.trim(),
      email: email.trim(),
      color,
      host,
      webUrl: web ? url : undefined,
      webIgnoreCert: web && ignoreCert ? true : undefined,
      webUser: web && webUser.trim() ? webUser.trim() : undefined,
      webAutoLogin: web && webUser.trim() && autoLogin ? true : undefined,
      short: short.trim() || undefined,
      ...icon,
      // pás aplikácií (upravuje sa ceruzkou dole v páse); tu len zapnúť/vypnúť
      railOrder: e?.railOrder,
      railHidden: e?.railHidden,
      railCustom: e?.railCustom,
      railOff: railOff || undefined,
    };
    setBusy(true);
    if (web && webUser.trim() && !webPass && !hasStored) {
      setBusy(false);
      return setError('Zadaj heslo, alebo vymaž prihlasovacie meno, ak ho nechceš ukladať.');
    }
    const r = await window.mailApi.saveAccount(acc, web && webPass ? webPass : undefined);
    setBusy(false);
    if (r.ok) props.onSaved(acc);
    else setError(r.error);
  }

  async function remove() {
    if (!e || !window.mailApi) return;
    setBusy(true);
    const r = await window.mailApi.deleteAccount(e.id);
    setBusy(false);
    if (r && 'ok' in r && !r.ok) {
      setConfirmDelete(false);
      return setError(r.error);
    }
    props.onDeleted();
  }

  async function signOut() {
    if (!e || !window.mailApi) return;
    await window.mailApi.webClear(accountPartition(e.id));
    setError(null);
    props.onSaved(e);
  }

  return (
    <Dialog open onOpenChange={(_, d) => !d.open && props.onClose()}>
      <DialogSurface className="compose-dialog">
        <DialogBody>
          <DialogTitle>{web ? (e ? 'Nastavenia webmailu' : 'Nový webmail') : e ? 'Nastavenia konta Microsoft 365' : 'Nové konto Microsoft 365'}</DialogTitle>
          <DialogContent className="dialog-content">
            <MessageBar intent="info">
              <MessageBarBody>
                {web
                  ? 'Webová pošta sa zobrazí ako stránka v okne s vlastným prihlásením. Heslo môžeš uložiť nižšie – program ho vyplní na prihlasovacej stránke.'
                  : 'Konto sa zobrazí ako vložený Outlook na webe. Prihlásiš sa priamo v ňom (aj s overením MFA); aplikácia heslo neukladá.'}
              </MessageBarBody>
            </MessageBar>
            {web && (
              <Field label="Adresa webmailu" required>
                <Input value={webUrl} placeholder="https://mail.firma.sk" onChange={(_, d) => setWebUrl(d.value)} />
              </Field>
            )}
            {!web && (
              <Checkbox
                label="Zobraziť pás aplikácií vľavo (Pošta, Kalendár, Teams, Word…). Poradie a výber upravíš ceruzkou dole v páse."
                checked={!railOff}
                onChange={(_, d) => setRailOff(!d.checked)}
              />
            )}
            {web && (
              <div className="webmail-cred">
                <div className="grid2">
                  <Field label="Prihlasovacie meno" hint="Nepovinné. Program ho s heslom vyplní na prihlasovacej stránke.">
                    <Input value={webUser} autoComplete="off" placeholder="meno alebo meno@firma.sk" onChange={(_, d) => setWebUser(d.value)} />
                  </Field>
                  <Field label="Heslo" hint={hasStored ? 'Heslo je uložené. Nové zadaj len pri zmene.' : 'Uloží sa zašifrované v trezore systému (ako heslá IMAP kont).'}>
                    <Input type="password" value={webPass} autoComplete="new-password" placeholder={hasStored ? '••••••••' : ''} disabled={!webUser.trim()} onChange={(_, d) => setWebPass(d.value)} />
                  </Field>
                </div>
                <Checkbox label="Po vyplnení sa rovno prihlásiť" checked={autoLogin} disabled={!webUser.trim()} onChange={(_, d) => setAutoLogin(!!d.checked)} />
                {hasStored && !webUser.trim() && <span className="muted small">Uložené heslo sa po uložení zabudne.</span>}
              </div>
            )}
            {web && (
              <Checkbox
                label="Dôverovať certifikátu servera, aj keď je neplatný (interný server, vlastný certifikát)"
                checked={ignoreCert}
                onChange={(_, d) => setIgnoreCert(!!d.checked)}
              />
            )}
            <Field label="Názov v páse kont">
              <Input value={name} onChange={(_, d) => setName(d.value)} />
            </Field>
            <Field label="Ikonka v páse kont">
              <AccountIconPicker value={icon} color={color} fallback={initials(name || '?')} onChange={setIcon} />
            </Field>
            <Field label="Štítok v zbalenom páse kont" hint={web ? 'Max. 8 znakov. Prázdne = WEBMAIL.' : 'Max. 8 znakov. Prázdne = MS365 (firemné) alebo MSO (osobné).'}>
              <Input value={short} maxLength={8} placeholder={web ? 'WEBMAIL' : host === 'live' ? 'MSO' : 'MS365'} onChange={(_, d) => setShort(d.value.toUpperCase())} />
            </Field>
            <Field label="E-mail (len na zobrazenie)">
              <Input value={email} placeholder="meno@firma.sk" onChange={(_, d) => setEmail(d.value)} />
            </Field>
            {!web && <Field label="Typ konta">
              <RadioGroup value={host} onChange={(_, d) => setHost(d.value as 'office' | 'live')}>
                <Radio value="office" label="Firemné alebo školské (Microsoft 365, outlook.office.com)" />
                <Radio value="live" label="Osobné (Outlook.com, Hotmail, Live – outlook.live.com)" />
              </RadioGroup>
            </Field>}
            <Field label="Farba">
              <div className="row colors">
                {ACCOUNT_COLORS.map((c) => (
                  <button key={c} className={`swatch ${c === color ? 'on' : ''}`} style={{ background: c }} onClick={() => setColor(c)} aria-label={`Farba ${c}`} />
                ))}
                <label className={`swatch custom ${ACCOUNT_COLORS.includes(color) ? '' : 'on'}`} title="Vlastná farba" style={{ background: ACCOUNT_COLORS.includes(color) ? undefined : color }}>
                  <input type="color" value={color} onChange={(e) => setColor(e.target.value)} aria-label="Vlastná farba konta" />
                </label>
              </div>
            </Field>
            {error && <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar>}
          </DialogContent>
          <DialogActions className="dialog-actions">
            {confirmDelete ? (
              // potvrdenie priamo v lište tlačidiel – vždy viditeľné (v obsahu dialógu bolo pod posúvaním a nebolo ho vidieť)
              <>
                <span className="confirm-del">Odstrániť konto z aplikácie a vymazať jeho prihlásenie? Pošta {web ? 'na serveri' : 'v Microsoft 365'} zostane nedotknutá.</span>
                <div className="grow" />
                <Button appearance="primary" disabled={busy} onClick={() => void remove()}>{busy ? 'Odstraňujem…' : 'Odstrániť'}</Button>
                <Button disabled={busy} onClick={() => setConfirmDelete(false)}>Nie</Button>
              </>
            ) : (
              <>
                {e && <Button appearance="subtle" onClick={() => setConfirmDelete(true)}>Odstrániť konto</Button>}
                {e && <Button appearance="subtle" onClick={() => void signOut()}>Odhlásiť</Button>}
                <div className="grow" />
                <Button appearance="primary" disabled={busy} onClick={() => void save()}>Uložiť</Button>
                <Button onClick={props.onClose}>Zrušiť</Button>
              </>
            )}
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
