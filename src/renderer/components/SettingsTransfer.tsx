// Export / import nastavení (prenos na iný počítač, aj Windows ↔ Mac). Bez hesiel.
import { Button, Checkbox, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, MessageBar, MessageBarBody, Spinner } from '@fluentui/react-components';
import { useEffect, useState } from 'react';
import type { BackupInfo } from '../../core/api';

const OS: Record<string, string> = { win32: 'Windows', darwin: 'macOS', linux: 'Linux' };

export function SettingsTransfer(props: { mode: 'export' | 'import'; onClose: () => void; onStatus: (s: string) => void }) {
  const api = window.mailApi;
  const [parts, setParts] = useState({ settings: true, accounts: true, contacts: true, calendars: true });
  const [info, setInfo] = useState<BackupInfo | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');

  const pick = async () => {
    if (!api) return;
    setError('');
    const r = await api.settingsPick();
    if (!r.ok) return setError(r.error);
    if (!r.data) return info ? undefined : props.onClose();
    setInfo(r.data);
    setParts({ settings: true, accounts: r.data.has.accounts > 0, contacts: r.data.has.contacts > 0, calendars: r.data.has.calendars > 0 });
  };
  useEffect(() => {
    if (props.mode === 'import') void pick();
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const doExport = async () => {
    if (!api) return;
    setBusy(true);
    const r = await api.settingsExport({ accounts: parts.accounts, contacts: parts.contacts, calendars: parts.calendars });
    setBusy(false);
    if (!r.ok) return setError(r.error);
    if (!r.data) return;
    props.onStatus(`Nastavenia uložené: ${r.data} (bez hesiel).`);
    props.onClose();
  };
  const doImport = async () => {
    if (!api || !info) return;
    setBusy(true);
    const r = await api.settingsApply(info.token, parts);
    if (!r.ok) {
      setBusy(false);
      setError(r.error);
    }
    // pri úspechu sa program reštartuje
  };

  const box = (k: keyof typeof parts, label: string, disabled = false, fixed = false) => (
    <Checkbox label={label} checked={fixed || (parts[k] && !disabled)} disabled={disabled || busy} onChange={(_, d) => setParts({ ...parts, [k]: !!d.checked })} />
  );
  const missing = info?.launchers.filter((l) => !l.ok) ?? [];
  const okL = info?.launchers.filter((l) => l.ok) ?? [];

  return (
    <Dialog open onOpenChange={(_, d) => !d.open && !busy && props.onClose()}>
      <DialogSurface style={{ maxWidth: 580 }}>
        <DialogBody>
          <DialogTitle>{props.mode === 'export' ? 'Exportovať nastavenia' : 'Importovať nastavenia'}</DialogTitle>
          <DialogContent className="dialog-content">
            {props.mode === 'export' ? (
              <>
                <span>Uloží nastavenia do jedného súboru – na prenos na iný počítač (aj z Windows na Mac a naopak) alebo ako záloha.</span>
                {box('settings', 'Nastavenia programu (vzhľad, Súhrn, pravý panel, podpisy, pravidlá, kategórie, šablóny…)', true, true)}
                {box('accounts', 'Kontá (servery a prihlasovacie mená)')}
                {box('contacts', 'Lokálne kontakty a nedávne adresy')}
                {box('calendars', 'Lokálne kalendáre a odbery kalendárov')}
                <MessageBar intent="info"><MessageBarBody>Heslá sa neukladajú – na novom počítači ich zadáš pri prvom pripojení. Súbor obsahuje adresy serverov, preto ho neposielaj cudzím.</MessageBarBody></MessageBar>
              </>
            ) : !info ? (
              <Spinner size="small" label="Vyber súbor s nastaveniami…" />
            ) : (
              <>
                <span>
                  <b>{info.file}</b> – z {OS[info.platform] ?? info.platform}, verzia {info.app}, {new Date(info.exported).toLocaleString('sk-SK')}
                </span>
                {box('settings', 'Nastavenia programu (nahradia súčasné)')}
                {box('accounts', `Kontá: ${info.has.accounts}${info.accountsExisting ? ` (${info.accountsExisting} už tu je – preskočia sa)` : ''}`, !info.has.accounts)}
                {box('contacts', `Lokálne kontakty: ${info.has.contacts} (doplnia sa)`, !info.has.contacts)}
                {box('calendars', `Lokálne kalendáre: ${info.has.calendars} (doplnia sa)`, !info.has.calendars)}
                {info.launchers.length > 0 && parts.settings && (
                  <MessageBar intent={missing.length ? 'warning' : 'info'}>
                    <MessageBarBody>
                      Programy v pravom paneli:{' '}
                      {okL.length ? `prevezmú sa ${okL.map((l) => l.name).join(', ')}. ` : ''}
                      {missing.length
                        ? `Na tomto počítači sa nenašli (vynechajú sa): ${missing.map((l) => l.name).join(', ')}. Pridáš ich znova cez + v pravom paneli.`
                        : ''}
                      {' '}Programy, ktoré tu už máš, ostanú.
                    </MessageBarBody>
                  </MessageBar>
                )}
                <span className="muted small">Po importe sa program reštartuje. K pridaným kontám zadáš heslá pri prvom pripojení.</span>
              </>
            )}
            {error && <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar>}
          </DialogContent>
          <DialogActions>
            {props.mode === 'import' && info && <Button style={{ whiteSpace: "nowrap" }} onClick={() => void pick()} disabled={busy}>Iný súbor…</Button>}
            <Button onClick={props.onClose} disabled={busy}>Zrušiť</Button>
            {props.mode === 'export' ? (
              <Button appearance="primary" onClick={() => void doExport()} disabled={busy}>Uložiť súbor…</Button>
            ) : (
              <Button appearance="primary" onClick={() => void doImport()} disabled={busy || !info || !Object.values(parts).some(Boolean)}>
                {busy ? 'Importujem…' : 'Importovať'}
              </Button>
            )}
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
