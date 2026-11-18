// Pás kont úplne vľavo: „Všetky doručené“, kontá (ikona s iniciálami, neprečítané), pridať konto.
// Poradie sa mení ťahaním, pravé tlačidlo = nastavenia konta.
import { Menu, MenuDivider, MenuItem, MenuList, MenuPopover, MenuTrigger, Tooltip } from '@fluentui/react-components';
import { useState } from 'react';
import type { AccountPublic } from '../../core/account';
import { AddIcon, AllInboxIcon, AppMailIcon, InfoIcon, PinIcon, SettingsIcon } from '../icons';
import { accountColor } from '../theme';
import { AccountIconContent } from './AccountIcon';

export const ALL = 'all';

export const initials = (name: string) =>
  name
    .split(/[\s.@_\-–—]+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((p) => p[0]!.toUpperCase())
    .join('') || '?';


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
  onCrypto: () => void;
  onAbout: () => void;
  /** krátke označenie typu konta na ikone (Z, IMAP, 365) */
  kindOf?: (a: AccountPublic) => string;
  onTestNotification: () => void;
  onZimbra?: () => void;
  onToggleNotifications: () => void;
  gpuForced: boolean;
  onToggleGpuForced: () => void;
  onOpenLog: () => void;
  /** Ponuka je otvorená: vložené weby sa musia skryť, inak by ju prekryli. */
  onMenuOpen: (open: boolean) => void;
  /** zoznam kont s menami natrvalo rozbalený (pripnutý) */
  pinned: boolean;
  onPin: (v: boolean) => void;
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

  const settingsMenu = (
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
            <MenuItem onClick={props.onCrypto}>Šifrovanie (PGP, S/MIME)…</MenuItem>
            {props.onZimbra && <MenuItem onClick={props.onZimbra}>Zimbra: neprítomnosť, filtre, zdieľania, heslo…</MenuItem>}
            <MenuItem onClick={props.onTestNotification}>Otestovať upozornenie</MenuItem>
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
  );
  const [open, setOpen] = useState(false);
  const expanded = props.pinned || open;
  const kindWord = (a: AccountPublic) => {
    if (a.short) return a.short;
    const k = props.kindOf?.(a) ?? '';
    return k === 'Z' ? 'ZIMBRA' : k === 'OS' ? 'MSO' : k === '365' ? 'MS365' : k;
  };
  const setFly = (v: boolean) => {
    setOpen(v);
    if (!props.pinned) props.onMenuOpen(v);
  };
  const pick = (id: string) => {
    props.onSelect(id);
    if (!props.pinned) setFly(false);
  };
  const fmt = (n: number) => (n > 9999 ? '9999+' : n.toLocaleString('sk-SK'));
  const dragProps = (a: AccountPublic) => ({
    draggable: true,
    onDragStart: () => setDrag(a.id),
    onDragEnd: () => setDrag(null),
    onDragOver: (e: React.DragEvent) => e.preventDefault(),
    onDrop: () => drop(a.id),
    onContextMenu: (e: React.MouseEvent) => {
      e.preventDefault();
      props.onEdit(a);
    },
  });

  // rozbalený zoznam kont s menami (B): pripnutý v rozložení alebo vysúvaný nad obsahom
  const list = (
    <nav className={`acc-list ${props.pinned ? 'docked' : 'flyout'}`} aria-label="Kontá">
      <div className="acc-list-head">
        <button className="acc-tool sm" onClick={() => (props.pinned ? props.onPin(false) : setFly(false))} aria-label="Zbaliť zoznam kont" title="Zbaliť">
          «
        </button>
        <AppMailIcon className="acc-list-logo" aria-hidden="true" />
        <span className="acc-list-title">Kontá</span>
        <button
          className={`acc-tool sm pin ${props.pinned ? 'on' : ''}`}
          onClick={() => {
            props.onPin(!props.pinned);
            setOpen(false);
            props.onMenuOpen(false);
          }}
          aria-pressed={props.pinned}
          title={props.pinned ? 'Odopnúť (zoznam sa bude vysúvať)' : 'Pripnúť zoznam natrvalo'}
          aria-label={props.pinned ? 'Odopnúť zoznam kont' : 'Pripnúť zoznam kont'}
        >
          <PinIcon />
        </button>
      </div>
      {props.showAll && (
        <button className={`acc-row ${props.active === ALL ? 'on' : ''}`} onClick={() => pick(ALL)} aria-pressed={props.active === ALL}>
          <AllInboxIcon className="acc-row-ico" />
          <span className="n">Všetky doručené</span>
          {totalUnread > 0 && <span className="c">{fmt(totalUnread)}</span>}
        </button>
      )}
      {props.showAll && <div className="acc-list-sep" />}
      {props.accounts.map((a) => {
        const on = props.active === a.id;
        const color = accountColor(a.color);
        const n = props.unread[a.id] ?? 0;
        return (
          <button
            key={a.id}
            className={`acc-row ${on ? 'on' : ''} ${drag === a.id ? 'dragging' : ''}`}
            style={on ? { boxShadow: `inset 3px 0 0 ${color}` } : undefined}
            title={`${a.name} – ${a.email}`}
            onClick={() => pick(a.id)}
            aria-pressed={on}
            {...dragProps(a)}
          >
            <span className="acc-dot" style={{ background: color }} />
            <span className="n">{a.name}</span>
            <span className="k">{kindWord(a)}</span>
            {n > 0 && <span className="c">{fmt(n)}</span>}
          </button>
        );
      })}
      <button className="acc-row add" onClick={props.onAdd}>
        <AddIcon className="acc-row-ico" />
        <span className="n">Pridať konto</span>
      </button>
      <div className="grow" />
      <div className="acc-list-tools">
        <button className="acc-tool" aria-label="O programe" title="O programe" onClick={props.onAbout}>
          <InfoIcon />
        </button>
        {settingsMenu}
      </div>
    </nav>
  );

  if (props.pinned) return list;

  return (
    <div className="acc-wrap">
    <nav className="acc-rail" aria-label="Kontá">
      <div className="acc-logo" title="Universum eMail" aria-hidden="true">
        <AppMailIcon />
      </div>
      <Tooltip content="Rozbaliť zoznam kont" relationship="label" positioning="after">
        <button className="acc-tool sm" aria-label="Rozbaliť zoznam kont" aria-expanded={open} onClick={() => setFly(!open)}>
          »
        </button>
      </Tooltip>
      {props.showAll && (
        <Tooltip content={`Všetky doručené${totalUnread ? ` · ${fmt(totalUnread)} neprečítaných` : ''}`} relationship="label" positioning="after">
          <button className={`acc-btn ${props.active === ALL ? 'on' : ''}`} onClick={() => props.onSelect(ALL)} aria-pressed={props.active === ALL}>
            <span className="acc-sq all">
              <AllInboxIcon />
              {totalUnread > 0 && <span className="acc-new" />}
            </span>
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
          <Tooltip key={a.id} content={`${a.name} – ${a.email}${n ? ` · ${fmt(n)} neprečítaných` : ''}`} relationship="label" positioning="after">
            <button className={`acc-btn ${on ? 'on' : ''} ${drag === a.id ? 'dragging' : ''}`} onClick={() => props.onSelect(a.id)} aria-pressed={on} {...dragProps(a)}>
              <span className="acc-sq" style={on ? { background: color, color: '#fff', borderColor: color } : { borderLeftColor: color }}>
                <AccountIconContent cfg={a} fallback={initials(a.name)} />
                {n > 0 && <span className="acc-new" style={{ background: on ? '#fff' : color }} />}
              </span>
              <span className="acc-label">{kindWord(a) || a.name}</span>
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
      <Tooltip content="O programe" relationship="label" positioning="after">
        <button className="acc-tool" aria-label="O programe" onClick={props.onAbout}>
          <InfoIcon />
        </button>
      </Tooltip>
      {settingsMenu}
    </nav>
    {open && (
      <>
        <div className="acc-backdrop" onClick={() => setFly(false)} />
        {list}
      </>
    )}
    </div>
  );
}
