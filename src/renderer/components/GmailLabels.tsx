// Gmail: štítky otvorenej správy (X-GM-LABELS) – zobrazenie, odobratie (×) a pridanie (+ Štítok).
// Štítky Gmailu sú v IMAP priečinky: odobratie štítku aktuálneho priečinka správu z neho presunie.
import { Menu, MenuItem, MenuList, MenuPopover, MenuTrigger } from '@fluentui/react-components';
import { useEffect, useState } from 'react';
import type { FolderInfo, MessageFull } from '../../core/api';

/** Systémové štítky, ktoré má zmysel ukázať (ostatné – Odoslané, Koncepty… – sú zrejmé z priečinka). */
const SYSTEM: Record<string, string> = { '\\Inbox': 'Doručená pošta', '\\Important': 'Dôležité' };

export function GmailLabels(props: { message: MessageFull; folders: FolderInfo[]; onChanged: (removedCurrent: boolean) => void }) {
  const api = window.mailApi;
  const m = props.message;
  const [labels, setLabels] = useState<string[] | null>(null);
  const [err, setErr] = useState<string | null>(null);

  useEffect(() => {
    let live = true;
    setLabels(null);
    setErr(null);
    void api?.gmailLabels(m.accountId, m.folder, m.uid).then((r) => {
      if (!live) return;
      if (r.ok) setLabels(r.data);
      else setErr(r.error);
    });
    return () => {
      live = false;
    };
  }, [api, m.accountId, m.folder, m.uid]);

  const userFolders = props.folders.filter((f) => !f.specialUse && f.path.toUpperCase() !== 'INBOX');
  const nameOf = (l: string) => SYSTEM[l] ?? props.folders.find((f) => f.path === l)?.path ?? l;
  const isCurrent = (l: string) => l === m.folder || (l === '\\Inbox' && m.folder.toUpperCase() === 'INBOX');

  async function change(add: string[], remove: string[]) {
    if (!api) return;
    const r = await api.gmailSetLabels(m.accountId, m.folder, [m.uid], add, remove);
    if (!r.ok) {
      setErr(r.error);
      return;
    }
    setLabels((cur) => [...new Set([...(cur ?? []).filter((x) => !remove.includes(x)), ...add])]);
    props.onChanged(remove.some(isCurrent));
  }

  if (err) return <div className="read-cats muted small">Štítky Gmailu: {err}</div>;
  if (!labels) return null;
  const shown = labels.filter((l) => SYSTEM[l] || !l.startsWith('\\'));
  const addable = [
    ...(labels.includes('\\Inbox') ? [] : [{ id: '\\Inbox', name: 'Doručená pošta' }]),
    ...userFolders.filter((f) => !labels.includes(f.path)).map((f) => ({ id: f.path, name: f.path })),
  ];

  return (
    <div className="read-cats gmail-labels" aria-label="Štítky Gmailu">
      {shown.map((l) => (
        <span key={l} className="cat-chip gmail-label" title={isCurrent(l) ? 'Odobratím štítku správa zmizne z tohto priečinka' : undefined}>
          {nameOf(l)}
          {l !== '\\Important' && (
            <button className="cat-x" aria-label={`Odobrať štítok ${nameOf(l)}`} onClick={() => void change([], [l])}>×</button>
          )}
        </span>
      ))}
      {addable.length > 0 && (
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <button className="cat-chip gmail-label add" aria-label="Pridať štítok">+ Štítok</button>
          </MenuTrigger>
          <MenuPopover>
            <MenuList style={{ maxHeight: 320, overflowY: 'auto' }}>
              {addable.map((a) => (
                <MenuItem key={a.id} onClick={() => void change([a.id], [])}>{a.name}</MenuItem>
              ))}
            </MenuList>
          </MenuPopover>
        </Menu>
      )}
    </div>
  );
}
