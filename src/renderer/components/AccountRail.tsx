// Pás kont úplne vľavo: „Všetky doručené“, kontá (ikona s iniciálami, neprečítané), pridať konto.
// Poradie sa mení ťahaním, pravé tlačidlo = nastavenia konta.
import { Menu, MenuDivider, MenuItem, MenuList, MenuPopover, MenuTrigger, Tooltip } from '@fluentui/react-components';
import { useState } from 'react';
import type { AccountPublic } from '../../core/account';
import { AddIcon, AllInboxIcon, SettingsIcon } from '../icons';
import { accountColor } from '../theme';

export const ALL = 'all';

export const initials = (name: string) =>
  name
    .split(/[\s.@_-]+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((p) => p[0]!.toUpperCase())
    .join('') || '?';

const shortName = (a: AccountPublic) => (a.name.length > 10 ? `${a.name.slice(0, 9)}…` : a.name);

export function AccountRail(props: {
  accounts: AccountPublic[];
  active: string;
  unread: Record<string, number>;
  showAll: boolean;
  onSelect: (id: string) => void;
  onAdd: () => void;
  onEdit: (a: AccountPublic) => void;
  onReorder: (ids: string[]) => void;
  onSettings: () => void;
  onSignatures: () => void;
  gpuOff: boolean;
  onToggleGpu: () => void;
  notifications: boolean;
  onToggleNotifications: () => void;
  gpuForced: boolean;
  onToggleGpuForced: () => void;
  onOpenLog: () => void;
  /** Ponuka je otvorená: vložené weby sa musia skryť, inak by ju prekryli. */
  onMenuOpen: (open: boolean) => void;
}) {
  const [drag, setDrag] = useState<string | null>(null);
  const totalUnread = Object.values(props.unread).reduce((a, b) => a + b, 0);

  const drop = (targetId: string) => {
    if (!drag || drag === targetId) return;
    const ids = props.accounts.map((a) => a.id).filter((id) => id !== drag);
    ids.splice(ids.indexOf(targetId), 0, drag);
    props.onReorder(ids);
    setDrag(null);
  };

  return (
    <nav className="acc-rail" aria-label="Kontá">
      {props.showAll && (
        <Tooltip content="Všetky doručené (IMAP kontá)" relationship="label" positioning="after">
          <button className={`acc-btn ${props.active === ALL ? 'on' : ''}`} onClick={() => props.onSelect(ALL)} aria-pressed={props.active === ALL}>
            <span className="acc-bar" style={{ background: props.active === ALL ? 'var(--colorBrandBackground)' : 'transparent' }} />
            <span className="acc-icon all">
              <AllInboxIcon />
            </span>
            {totalUnread > 0 && <span className="acc-badge">{totalUnread > 999 ? '999+' : totalUnread}</span>}
            <span className="acc-label">Všetky</span>
          </button>
        </Tooltip>
      )}
      {props.showAll && <div className="acc-sep" />}
      {props.accounts.map((a) => {
        const on = props.active === a.id;
        const color = accountColor(a.color);
        const n = props.unread[a.id] ?? 0;
        return (
          <Tooltip key={a.id} content={`${a.name} – ${a.email}${a.kind === 'm365' ? ' (Outlook na webe)' : ''}`} relationship="label" positioning="after">
            <button
              className={`acc-btn ${on ? 'on' : ''} ${drag === a.id ? 'dragging' : ''}`}
              onClick={() => props.onSelect(a.id)}
              onContextMenu={(e) => {
                e.preventDefault();
                props.onEdit(a);
              }}
              draggable
              onDragStart={() => setDrag(a.id)}
              onDragEnd={() => setDrag(null)}
              onDragOver={(e) => e.preventDefault()}
              onDrop={() => drop(a.id)}
              aria-pressed={on}
            >
              <span className="acc-bar" style={{ background: on ? color : 'transparent' }} />
              <span
                className="acc-icon"
                style={{
                  background: color,
                  borderRadius: on ? 14 : '50%',
                  boxShadow: on ? `0 0 0 2px var(--accRailBg), 0 0 0 4px ${color}` : 'none',
                }}
              >
                {initials(a.name)}
              </span>
              {n > 0 && <span className="acc-badge">{n > 999 ? '999+' : n}</span>}
              <span className="acc-label">{shortName(a)}</span>
            </button>
          </Tooltip>
        );
      })}
      <Tooltip content="Pridať konto" relationship="label" positioning="after">
        <button className="acc-add" onClick={props.onAdd} aria-label="Pridať konto">
          <AddIcon />
        </button>
      </Tooltip>
      <div className="grow" />
      <Menu positioning="after-bottom" onOpenChange={(_, d) => props.onMenuOpen(d.open)}>
        <MenuTrigger disableButtonEnhancement>
          <Tooltip content="Nastavenia" relationship="label" positioning="after">
            <button className="acc-tool" aria-label="Nastavenia">
              <SettingsIcon />
            </button>
          </Tooltip>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuItem onClick={props.onSettings}>Nastavenia konta…</MenuItem>
            <MenuItem onClick={props.onAdd}>Pridať konto…</MenuItem>
            <MenuItem onClick={props.onSignatures}>Podpisy…</MenuItem>
            <MenuItem onClick={props.onToggleNotifications}>
              {props.notifications ? 'Vypnúť upozornenia na novú poštu' : 'Zapnúť upozornenia na novú poštu'}
            </MenuItem>
            <MenuDivider />
            <MenuItem onClick={props.onToggleGpu}>
              {props.gpuOff ? 'Zapnúť hardvérovú akceleráciu (reštart)' : 'Vypnúť hardvérovú akceleráciu, ak obraz bliká (reštart)'}
            </MenuItem>
            <MenuItem onClick={props.onToggleGpuForced}>
              {props.gpuForced ? 'Zrušiť vynútenie grafickej karty (reštart)' : 'Vynútiť grafickú kartu bez sandboxu – test proti blikaniu (reštart)'}
            </MenuItem>
            <MenuItem onClick={props.onOpenLog}>Otvoriť záznam chýb</MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>
    </nav>
  );
}
