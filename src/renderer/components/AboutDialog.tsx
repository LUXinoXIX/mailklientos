// O programe: verzia, autori, použité technológie a licencie.
import { Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface } from '@fluentui/react-components';
import { useEffect, useState } from 'react';
import icon from '../app-icon.svg';

const TECH: Array<[string, string]> = [
  ['Electron', 'okno aplikácie (Chromium + Node.js)'],
  ['React, Fluent UI', 'rozhranie vo vzhľade Microsoft 365'],
  ['TipTap', 'editor správ, podpisov a šablón'],
  ['imapflow, nodemailer, mailparser', 'IMAP, SMTP a čítanie správ'],
  ['SQLite (vstavané v Node.js)', 'lokálna cache a fulltext'],
  ['OpenPGP.js, node-forge', 'šifrovanie a podpis PGP a S/MIME'],
  ['@xmldom/xmldom', 'CalDAV a CardDAV (kalendár, kontakty)'],
  ['DOMPurify', 'bezpečné zobrazenie HTML správ'],
  ['Zimbra SOAP', 'adresár firmy, filtre, neprítomnosť na serveri'],
];

export function AboutDialog(props: { onClose: () => void }) {
  const [info, setInfo] = useState<{ version: string; electron: string; chrome: string; node: string; platform: string } | null>(null);
  useEffect(() => {
    void window.mailApi?.appInfo().then((r) => r.ok && setInfo(r.data));
  }, []);
  return (
    <Dialog open onOpenChange={(_, d) => !d.open && props.onClose()}>
      <DialogSurface className="about-surface" style={{ maxWidth: 640 }}>
        <DialogBody>
          <DialogContent>
            <div className="about-head">
              <img src={icon} width={84} height={84} alt="" />
              <div>
                <div className="about-name">Universum eMail</div>
                <div className="about-ver">
                  verzia {info?.version ?? '–'}
                  {info && ` · Electron ${info.electron} · Chromium ${info.chrome.split('.')[0]} · ${info.platform}`}
                </div>
              </div>
            </div>
            <div className="about-sec">Autori</div>
            <div>Martin – zadanie, návrh a testovanie</div>
            <div>Claude (Anthropic) – návrh architektúry a vývoj</div>
            <div className="about-sec">Použité technológie (open source)</div>
            <div className="about-tech">
              {TECH.map(([n, d]) => (
                <div key={n}><b>{n}</b> <span className="muted">– {d}</span></div>
              ))}
            </div>
            <div className="about-sec">Zdrojový kód</div>
            <div className="muted small">github.com/LUXinoXIX/mailklientos · knižnice tretích strán podliehajú svojim licenciám (MIT, Apache 2.0, LGPL, BSD).</div>
          </DialogContent>
          <DialogActions>
            <Button appearance="primary" onClick={() => void window.mailApi?.updateCheck()}>Skontrolovať aktualizácie</Button>
            <Button onClick={props.onClose}>Zavrieť</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

const SHORTCUTS: Array<[string, string]> = [
  ['Ctrl+N', 'Nová správa'],
  ['Ctrl+R', 'Odpovedať'],
  ['Ctrl+Shift+R', 'Odpovedať všetkým'],
  ['Ctrl+F', 'Preposlať'],
  ['Ctrl+Enter', 'Odoslať (pri písaní)'],
  ['Ctrl+Q / Ctrl+U', 'Označiť ako prečítané / neprečítané'],
  ['Insert', 'Vlajka'],
  ['Delete', 'Odstrániť'],
  ['Ctrl+P', 'Tlačiť'],
  ['Ctrl+E', 'Hľadať'],
  ['F9', 'Odoslať a prijať'],
  ['Enter / dvojklik', 'Otvoriť správu v okne (Shift = na karte)'],
  ['Alt+↑ / Alt+↓', 'Predchádzajúca / nasledujúca správa (v okne správy)'],
  ['Ctrl+1 … 4', 'Pošta, Kalendár, Kontakty, Úlohy'],
];

export function ShortcutsDialog(props: { onClose: () => void }) {
  return (
    <Dialog open onOpenChange={(_, d) => !d.open && props.onClose()}>
      <DialogSurface style={{ maxWidth: 480 }}>
        <DialogBody>
          <DialogContent>
            <div className="about-name" style={{ fontSize: 18, marginBottom: 10 }}>Klávesové skratky</div>
            <table className="keys">
              <tbody>
                {SHORTCUTS.map(([k, d]) => (
                  <tr key={k}><td><kbd>{k}</kbd></td><td>{d}</td></tr>
                ))}
              </tbody>
            </table>
          </DialogContent>
          <DialogActions><Button onClick={props.onClose}>Zavrieť</Button></DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
