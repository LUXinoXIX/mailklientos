// Horný pruh, ribbon a pás modulov (variant D).
import {
  SearchBox, Select, Tab, TabList, Toolbar, ToolbarButton, ToolbarDivider, Tooltip,
} from '@fluentui/react-components';
import type { ReactNode } from 'react';
import {
  AddAccountIcon, CalendarIcon, ComposeIcon, CopilotIcon, DeleteIcon, LaunchersIcon, MailIcon, PanelIcon,
  ReplyIcon, SettingsIcon, SyncIcon, TasksIcon, TeamsIcon, WhatsAppIcon,
} from '../icons';
import { THEMES, type ThemeKey } from '../theme';

export function TopBar(props: { search: string; onSearch: (v: string) => void }) {
  return (
    <header className="topbar">
      <div className="topbar-title">Mail klient</div>
      <SearchBox
        className="topbar-search"
        placeholder="Hľadať v načítaných správach"
        value={props.search}
        onChange={(_, d) => props.onSearch(d.value)}
        appearance="filled-lighter"
      />
      <div />
    </header>
  );
}

export type RibbonTab = 'home' | 'view';

export function Ribbon(props: {
  tab: RibbonTab;
  onTab: (t: RibbonTab) => void;
  canReply: boolean;
  onCompose: () => void;
  onReply: () => void;
  onRefresh: () => void;
  onAddAccount: () => void;
  onEditAccount?: () => void;
  theme: ThemeKey;
  onTheme: (t: ThemeKey) => void;
  webOpen: boolean;
  onToggleWeb: () => void;
  compact: boolean;
  onCompact: (v: boolean) => void;
}) {
  return (
    <div className={`ribbon ${props.compact ? 'ribbon-compact' : ''}`}>
      <TabList size="small" selectedValue={props.tab} onTabSelect={(_, d) => props.onTab(d.value as RibbonTab)}>
        <Tab value="home">Domov</Tab>
        <Tab value="view">Zobrazenie</Tab>
      </TabList>
      <Toolbar size={props.compact ? 'small' : 'medium'} className="ribbon-bar">
        {props.tab === 'home' ? (
          <>
            <ToolbarButton appearance="primary" icon={<ComposeIcon />} onClick={props.onCompose}>
              Nová správa
            </ToolbarButton>
            <ToolbarDivider />
            <ToolbarButton icon={<ReplyIcon />} disabled={!props.canReply} onClick={props.onReply}>
              Odpovedať
            </ToolbarButton>
            <ToolbarButton icon={<DeleteIcon />} disabled title="Príde v ďalšom kroku">
              Odstrániť
            </ToolbarButton>
            <ToolbarDivider />
            <ToolbarButton icon={<SyncIcon />} onClick={props.onRefresh}>
              Obnoviť
            </ToolbarButton>
            <ToolbarDivider />
            <ToolbarButton icon={<AddAccountIcon />} onClick={props.onAddAccount}>
              Pridať konto
            </ToolbarButton>
            {props.onEditAccount && (
              <ToolbarButton icon={<SettingsIcon />} onClick={props.onEditAccount}>
                Nastavenia konta
              </ToolbarButton>
            )}
          </>
        ) : (
          <>
            <label className="ribbon-label">Téma</label>
            <Select size="small" value={props.theme} onChange={(_, d) => props.onTheme(d.value as ThemeKey)}>
              {Object.entries(THEMES).map(([k, t]) => (
                <option key={k} value={k}>
                  {t.label}
                </option>
              ))}
            </Select>
            <ToolbarDivider />
            <ToolbarButton icon={<PanelIcon />} onClick={props.onToggleWeb}>
              {props.webOpen ? 'Skryť 4. stĺpec' : 'Zobraziť 4. stĺpec'}
            </ToolbarButton>
            <ToolbarButton onClick={() => props.onCompact(!props.compact)}>
              {props.compact ? 'Plný ribbon' : 'Kompaktný ribbon'}
            </ToolbarButton>
          </>
        )}
      </Toolbar>
    </div>
  );
}

export type ModuleKey = 'mail' | 'calendar' | 'tasks' | 'whatsapp' | 'teams' | 'copilot' | 'launchers';

const MODULES: Array<{ key: ModuleKey; label: string; icon: ReactNode; sep?: boolean }> = [
  { key: 'mail', label: 'Pošta', icon: <MailIcon /> },
  { key: 'calendar', label: 'Kalendár', icon: <CalendarIcon /> },
  { key: 'tasks', label: 'Úlohy', icon: <TasksIcon /> },
  { key: 'whatsapp', label: 'WhatsApp', icon: <WhatsAppIcon />, sep: true },
  { key: 'teams', label: 'Teams', icon: <TeamsIcon /> },
  { key: 'copilot', label: 'Copilot', icon: <CopilotIcon /> },
  { key: 'launchers', label: 'Spúšťače', icon: <LaunchersIcon />, sep: true },
];

export function ModuleRail(props: { active: ModuleKey; onSelect: (m: ModuleKey) => void }) {
  return (
    <nav className="rail">
      {MODULES.map((m) => (
        <div key={m.key}>
          {m.sep && <div className="rail-sep" />}
          <Tooltip content={m.label} relationship="label" positioning="after">
            <button
              className={`rail-btn ${props.active === m.key ? 'active' : ''}`}
              onClick={() => props.onSelect(m.key)}
              aria-label={m.label}
            >
              {m.icon}
            </button>
          </Tooltip>
        </div>
      ))}
    </nav>
  );
}
