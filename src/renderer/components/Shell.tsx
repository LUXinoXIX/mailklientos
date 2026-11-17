// Horný pruh (vo farbe konta), ribbon pre IMAP kontá a pás modulov podľa typu konta.
import {
  Menu, MenuItem, MenuList, MenuPopover, MenuTrigger, SearchBox, Select, Tab, TabList, Toolbar, ToolbarButton,
  ToolbarDivider, Tooltip,
} from '@fluentui/react-components';
import type { ReactNode } from 'react';
import type { FolderInfo } from '../../core/api';
import {
  ArchiveIcon, CalendarIcon, CloudIcon, ComposeIcon, ContactsIcon, DeleteIcon, ExcelIcon, FlagIcon, ForwardIcon, MailIcon,
  MarkReadIcon, MarkUnreadIcon, MoveIcon, PanelIcon, PowerPointIcon, ReplyAllIcon, ReplyIcon, SettingsIcon, SyncIcon,
  TasksIcon, TeamsIcon, WordIcon,
} from '../icons';
import { THEMES, type ThemeKey } from '../theme';

export function TopBar(props: {
  title: string;
  subtitle: string;
  color: string;
  search?: string;
  onSearch?: (v: string) => void;
  rightOpen: boolean;
  onToggleRight: () => void;
}) {
  return (
    <header className="topbar" style={{ background: props.color }}>
      <div className="topbar-title">
        <div className="topbar-name">{props.title}</div>
        <div className="topbar-sub">{props.subtitle}</div>
      </div>
      {props.onSearch ? (
        <SearchBox
          className="topbar-search"
          placeholder="Hľadať v načítaných správach"
          value={props.search ?? ''}
          onChange={(_, d) => props.onSearch!(d.value)}
          appearance="filled-lighter"
        />
      ) : (
        <div />
      )}
      <div className="topbar-tools">
        <Tooltip content={props.rightOpen ? 'Skryť pravý panel' : 'Pravý panel: weby a programy'} relationship="label">
          <button className={`topbar-btn ${props.rightOpen ? 'on' : ''}`} onClick={props.onToggleRight} aria-pressed={props.rightOpen}>
            <PanelIcon /> <span>Pravý panel</span>
          </button>
        </Tooltip>
      </div>
    </header>
  );
}

export type RibbonTab = 'home' | 'view';

export function Ribbon(props: {
  tab: RibbonTab;
  onTab: (t: RibbonTab) => void;
  hasMessage: boolean;
  seen: boolean;
  flagged: boolean;
  canMove: boolean;
  folders: FolderInfo[];
  onCompose: () => void;
  onReply: () => void;
  onReplyAll: () => void;
  onForward: () => void;
  onDelete: () => void;
  onArchive: () => void;
  onMove: (dest: string) => void;
  onToggleSeen: () => void;
  onToggleFlag: () => void;
  onRefresh: () => void;
  onEditAccount?: () => void;
  theme: ThemeKey;
  onTheme: (t: ThemeKey) => void;
  rightOpen: boolean;
  onToggleRight: () => void;
  compact: boolean;
  onCompact: (v: boolean) => void;
}) {
  const p = props;
  const sz = p.compact ? 'small' : 'medium';
  return (
    <div className={`ribbon ${p.compact ? 'ribbon-compact' : ''}`}>
      <TabList size="small" selectedValue={p.tab} onTabSelect={(_, d) => p.onTab(d.value as RibbonTab)}>
        <Tab value="home">Domov</Tab>
        <Tab value="view">Zobrazenie</Tab>
      </TabList>
      <Toolbar size={sz} className="ribbon-bar">
        {p.tab === 'home' ? (
          <>
            <ToolbarButton appearance="primary" icon={<ComposeIcon />} onClick={p.onCompose}>Nová správa</ToolbarButton>
            <ToolbarDivider />
            <ToolbarButton icon={<ReplyIcon />} disabled={!p.hasMessage} onClick={p.onReply}>Odpovedať</ToolbarButton>
            <ToolbarButton icon={<ReplyAllIcon />} disabled={!p.hasMessage} onClick={p.onReplyAll}>Všetkým</ToolbarButton>
            <ToolbarButton icon={<ForwardIcon />} disabled={!p.hasMessage} onClick={p.onForward}>Preposlať</ToolbarButton>
            <ToolbarDivider />
            <ToolbarButton icon={<DeleteIcon />} disabled={!p.hasMessage} onClick={p.onDelete}>Odstrániť</ToolbarButton>
            <ToolbarButton icon={<ArchiveIcon />} disabled={!p.hasMessage} onClick={p.onArchive}>Archív</ToolbarButton>
            <Menu>
              <MenuTrigger disableButtonEnhancement>
                <ToolbarButton icon={<MoveIcon />} disabled={!p.hasMessage || !p.canMove}>Presunúť</ToolbarButton>
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  {p.folders.map((f) => (
                    <MenuItem key={f.path} onClick={() => p.onMove(f.path)} style={{ paddingLeft: 12 + f.depth * 14 }}>
                      {f.name}
                    </MenuItem>
                  ))}
                </MenuList>
              </MenuPopover>
            </Menu>
            <ToolbarDivider />
            <ToolbarButton icon={p.seen ? <MarkUnreadIcon /> : <MarkReadIcon />} disabled={!p.hasMessage} onClick={p.onToggleSeen}>
              {p.seen ? 'Neprečítané' : 'Prečítané'}
            </ToolbarButton>
            <ToolbarButton icon={<FlagIcon />} disabled={!p.hasMessage} onClick={p.onToggleFlag}>
              {p.flagged ? 'Zrušiť vlajku' : 'Vlajka'}
            </ToolbarButton>
            <ToolbarDivider />
            <ToolbarButton icon={<SyncIcon />} onClick={p.onRefresh}>Obnoviť</ToolbarButton>
            {p.onEditAccount && (
              <ToolbarButton icon={<SettingsIcon />} onClick={p.onEditAccount}>Nastavenia konta</ToolbarButton>
            )}
          </>
        ) : (
          <>
            <label className="ribbon-label">Téma</label>
            <Select size="small" value={p.theme} onChange={(_, d) => p.onTheme(d.value as ThemeKey)}>
              {Object.entries(THEMES).map(([k, t]) => (
                <option key={k} value={k}>{t.label}</option>
              ))}
            </Select>
            <ToolbarDivider />
            <ToolbarButton icon={<PanelIcon />} onClick={p.onToggleRight}>
              {p.rightOpen ? 'Skryť pravý panel' : 'Zobraziť pravý panel'}
            </ToolbarButton>
            <ToolbarButton onClick={() => p.onCompact(!p.compact)}>{p.compact ? 'Plný ribbon' : 'Kompaktný ribbon'}</ToolbarButton>
          </>
        )}
      </Toolbar>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Pás modulov: mení sa podľa typu konta.

export type ImapModule = 'mail' | 'calendar' | 'tasks' | 'contacts';
export type M365Module = 'mail' | 'calendar' | 'todo' | 'people' | 'teams' | 'onedrive' | 'word' | 'excel' | 'powerpoint';

export interface ModuleDef<K extends string> {
  key: K;
  label: string;
  icon: ReactNode;
  sep?: boolean;
}

export const IMAP_MODULES: ModuleDef<ImapModule>[] = [
  { key: 'mail', label: 'Pošta', icon: <MailIcon /> },
  { key: 'calendar', label: 'Kalendár', icon: <CalendarIcon /> },
  { key: 'tasks', label: 'Úlohy', icon: <TasksIcon /> },
  { key: 'contacts', label: 'Kontakty', icon: <ContactsIcon /> },
];

export const M365_MODULES: ModuleDef<M365Module>[] = [
  { key: 'mail', label: 'Pošta', icon: <MailIcon /> },
  { key: 'calendar', label: 'Kalendár', icon: <CalendarIcon /> },
  { key: 'todo', label: 'To Do', icon: <TasksIcon /> },
  { key: 'people', label: 'Ľudia', icon: <ContactsIcon /> },
  { key: 'teams', label: 'Teams', icon: <TeamsIcon />, sep: true },
  { key: 'onedrive', label: 'OneDrive', icon: <CloudIcon /> },
  { key: 'word', label: 'Word', icon: <WordIcon /> },
  { key: 'excel', label: 'Excel', icon: <ExcelIcon /> },
  { key: 'powerpoint', label: 'PowerPoint', icon: <PowerPointIcon /> },
];

export function ModuleRail<K extends string>(props: { modules: ModuleDef<K>[]; active: K | null; color: string; onSelect: (k: K) => void }) {
  return (
    <nav className="rail" aria-label="Moduly konta">
      {props.modules.map((m) => {
        const on = props.active === m.key;
        return (
          <div key={m.key}>
            {m.sep && <div className="rail-sep" />}
            <Tooltip content={m.label} relationship="label" positioning="after">
              <button
                className={`rail-btn ${on ? 'active' : ''}`}
                style={on ? { color: props.color, boxShadow: `inset 3px 0 0 ${props.color}` } : undefined}
                onClick={() => props.onSelect(m.key)}
                aria-pressed={on}
              >
                {m.icon}
                <span className="rail-label">{m.label}</span>
              </button>
            </Tooltip>
          </div>
        );
      })}
    </nav>
  );
}
