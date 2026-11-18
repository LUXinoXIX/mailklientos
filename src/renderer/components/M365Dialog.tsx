// Výber typu nového konta a dialóg pre konto Microsoft 365 (vložený Outlook na webe).
import {
  Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input, MessageBar,
  MessageBarBody, Radio, RadioGroup,
} from '@fluentui/react-components';
import { useState } from 'react';
import { accountPartition, type M365Account } from '../../core/account';
import { ACCOUNT_COLORS, accountColor } from '../theme';
import { MailIcon, TeamsIcon } from '../icons';

export function AddAccountChooser(props: { onClose: () => void; onImap: () => void; onM365: (host: 'office' | 'live') => void }) {
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
  host?: 'office' | 'live';
  usedColors: string[];
  onClose: () => void;
  onSaved: (a: M365Account) => void;
  onDeleted: () => void;
}) {
  const e = props.edit;
  const [name, setName] = useState(e?.name ?? (props.host === 'live' ? 'Outlook osobný' : 'Outlook firemný'));
  const [email, setEmail] = useState(e?.email ?? '');
  const [short, setShort] = useState(e?.short ?? '');
  const [host, setHost] = useState<'office' | 'live'>(e?.host ?? props.host ?? 'office');
  const [color, setColor] = useState(accountColor(e?.color ?? ACCOUNT_COLORS.find((c) => !props.usedColors.includes(c)) ?? ACCOUNT_COLORS[1]!));
  const [error, setError] = useState<string | null>(null);
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [busy, setBusy] = useState(false);

  async function save() {
    if (!window.mailApi) return;
    if (!name.trim()) return setError('Zadaj názov konta.');
    const acc: M365Account = {
      id: e?.id ?? `m365-${Date.now().toString(36)}`,
      kind: 'm365',
      name: name.trim(),
      email: email.trim(),
      color,
      host,
      short: short.trim() || undefined,
    };
    setBusy(true);
    const r = await window.mailApi.saveAccount(acc);
    setBusy(false);
    if (r.ok) props.onSaved(acc);
    else setError(r.error);
  }

  async function remove() {
    if (!e || !window.mailApi) return;
    setBusy(true);
    await window.mailApi.deleteAccount(e.id);
    setBusy(false);
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
          <DialogTitle>{e ? 'Nastavenia konta Microsoft 365' : 'Nové konto Microsoft 365'}</DialogTitle>
          <DialogContent className="dialog-content">
            <MessageBar intent="info">
              <MessageBarBody>
                Konto sa zobrazí ako vložený Outlook na webe. Prihlásiš sa priamo v ňom (aj s overením MFA); aplikácia heslo neukladá.
              </MessageBarBody>
            </MessageBar>
            <Field label="Názov v páse kont">
              <Input value={name} onChange={(_, d) => setName(d.value)} />
            </Field>
            <Field label="Štítok v zbalenom páse kont" hint="Max. 8 znakov. Prázdne = MS365 (firemné) alebo MSO (osobné).">
              <Input value={short} maxLength={8} placeholder={host === 'live' ? 'MSO' : 'MS365'} onChange={(_, d) => setShort(d.value.toUpperCase())} />
            </Field>
            <Field label="E-mail (len na zobrazenie)">
              <Input value={email} placeholder="meno@firma.sk" onChange={(_, d) => setEmail(d.value)} />
            </Field>
            <Field label="Typ konta">
              <RadioGroup value={host} onChange={(_, d) => setHost(d.value as 'office' | 'live')}>
                <Radio value="office" label="Firemné alebo školské (Microsoft 365, outlook.office.com)" />
                <Radio value="live" label="Osobné (Outlook.com, Hotmail, Live – outlook.live.com)" />
              </RadioGroup>
            </Field>
            <Field label="Farba">
              <div className="row colors">
                {ACCOUNT_COLORS.map((c) => (
                  <button key={c} className={`swatch ${c === color ? 'on' : ''}`} style={{ background: c }} onClick={() => setColor(c)} aria-label={`Farba ${c}`} />
                ))}
              </div>
            </Field>
            {error && <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar>}
            {confirmDelete && (
              <MessageBar intent="warning">
                <MessageBarBody>
                  Odstrániť konto z aplikácie a vymazať jeho prihlásenie? Pošta v Microsoft 365 zostane nedotknutá.
                  <div className="row gap">
                    <Button size="small" appearance="primary" onClick={() => void remove()}>Odstrániť</Button>
                    <Button size="small" onClick={() => setConfirmDelete(false)}>Nie</Button>
                  </div>
                </MessageBarBody>
              </MessageBar>
            )}
          </DialogContent>
          <DialogActions className="dialog-actions">
            {e && <Button appearance="subtle" onClick={() => setConfirmDelete(true)}>Odstrániť konto</Button>}
            {e && <Button appearance="subtle" onClick={() => void signOut()}>Odhlásiť</Button>}
            <div className="grow" />
            <Button appearance="primary" disabled={busy} onClick={() => void save()}>Uložiť</Button>
            <Button onClick={props.onClose}>Zrušiť</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
