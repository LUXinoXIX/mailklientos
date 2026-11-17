// Konto Microsoft 365: vložený Outlook na webe ako celé okno, s kartami.
// Office aplikácie otvorené z Outlooku (prepínač aplikácií) alebo z pása modulov idú do kariet, nie do nových okien.
import { Button, Tooltip } from '@fluentui/react-components';
import { accountPartition, type M365Account } from '../../core/account';
import { BackIcon, CloseSmallIcon, OpenOutsideIcon, ReloadIcon, SettingsIcon } from '../icons';
import type { M365Module } from './Shell';
import { useRef } from 'react';
import { WebSlot, type WebSlotHandle } from './WebSlot';

export interface WebTab {
  key: string;
  title: string;
  url: string;
  module?: M365Module;
}

export interface TabsState {
  tabs: WebTab[];
  active: string;
}

export const baseKey = (accountId: string) => `acc-${accountId}`;

/** Adresy webových aplikácií pre firemné (office) a osobné (live) kontá. */
export function moduleUrl(m: M365Module, host: 'office' | 'live' = 'office'): string {
  const o = host === 'office';
  switch (m) {
    case 'mail': return o ? 'https://outlook.office.com/mail/' : 'https://outlook.live.com/mail/';
    case 'calendar': return o ? 'https://outlook.office.com/calendar/' : 'https://outlook.live.com/calendar/';
    case 'people': return o ? 'https://outlook.office.com/people/' : 'https://outlook.live.com/people/';
    case 'todo': return o ? 'https://to-do.office.com/tasks/' : 'https://to-do.live.com/tasks/';
    case 'teams': return o ? 'https://teams.microsoft.com/' : 'https://teams.live.com/';
    case 'onedrive': return o ? 'https://www.office.com/launch/onedrive' : 'https://onedrive.live.com/';
    case 'word': return 'https://www.office.com/launch/word';
    case 'excel': return 'https://www.office.com/launch/excel';
    case 'powerpoint': return 'https://www.office.com/launch/powerpoint';
    case 'copilot': return o ? 'https://m365.cloud.microsoft/chat' : 'https://copilot.microsoft.com/';
    case 'planner': return o ? 'https://planner.cloud.microsoft/' : 'https://to-do.live.com/tasks/';
    case 'sharepoint': return 'https://www.office.com/launch/sharepoint';
    case 'onenote': return o ? 'https://www.office.com/launch/onenote' : 'https://www.onenote.com/notebooks';
    case 'forms': return 'https://forms.office.com/';
    case 'loop': return 'https://loop.cloud.microsoft/';
  }
}

export const MODULE_TITLES: Record<M365Module, string> = {
  loop: 'Loop',
  mail: 'Outlook', calendar: 'Kalendár', people: 'Ľudia', todo: 'To Do', teams: 'Teams', onedrive: 'OneDrive',
  word: 'Word', excel: 'Excel', powerpoint: 'PowerPoint', copilot: 'Copilot', planner: 'Planner', sharepoint: 'SharePoint',
  onenote: 'OneNote', forms: 'Forms',
};

export function initialTabs(acc: M365Account): TabsState {
  const key = baseKey(acc.id);
  return { tabs: [{ key, title: 'Outlook', url: moduleUrl('mail', acc.host), module: 'mail' }], active: key };
}

/** Pekný názov karty z titulku stránky („Pošta – Ján Novák – Outlook“ → „Pošta“). */
export function tabTitle(raw: string, fallback: string) {
  const t = raw.split(/\s[-–|]\s/)[0]?.trim();
  if (!t || /^(about:|https?:)/.test(t)) return fallback;
  return t.length > 26 ? `${t.slice(0, 25)}…` : t;
}

export function M365View(props: {
  account: M365Account;
  state: TabsState;
  color: string;
  /** konto je práve vybraté (inak ostáva načítané, len skryté) */
  visible: boolean;
  onActivate: (key: string) => void;
  onClose: (key: string) => void;
  onTitle: (key: string, title: string) => void;
  onEditAccount: () => void;
}) {
  const api = window.mailApi;
  const { account, state } = props;
  const partition = accountPartition(account.id);
  const active = state.tabs.find((t) => t.key === state.active) ?? state.tabs[0]!;
  const slots = useRef(new Map<string, WebSlotHandle>());
  const cur = () => slots.current.get(active.key);

  return (
    <section className={`m365 ${props.visible ? '' : 'm365-hidden'}`} aria-hidden={!props.visible}>
      <div className="m365-bar">
        <div className="m365-tabs" role="tablist" aria-label="Karty konta">
          {state.tabs.map((t, i) => {
            const on = t.key === active.key;
            return (
              <div key={t.key} className={`m365-tab ${on ? 'on' : ''}`} style={on ? { boxShadow: `inset 0 -2px 0 ${props.color}` } : undefined}>
                <button role="tab" aria-selected={on} className="m365-tab-btn" onClick={() => props.onActivate(t.key)} title={t.url}>
                  {t.title}
                </button>
                {i > 0 && (
                  <button className="m365-tab-x" aria-label={`Zavrieť kartu ${t.title}`} onClick={() => props.onClose(t.key)}>
                    <CloseSmallIcon />
                  </button>
                )}
              </div>
            );
          })}
        </div>
        <div className="m365-tools">
          <Tooltip content="Späť" relationship="label">
            <Button size="small" appearance="subtle" icon={<BackIcon />} onClick={() => cur()?.back()} aria-label="Späť" />
          </Tooltip>
          <Tooltip content="Obnoviť stránku" relationship="label">
            <Button size="small" appearance="subtle" icon={<ReloadIcon />} onClick={() => cur()?.reload()} aria-label="Obnoviť" />
          </Tooltip>
          <Tooltip content="Otvoriť túto stránku v Edge" relationship="label">
            <Button size="small" appearance="subtle" icon={<OpenOutsideIcon />} onClick={() => api?.webOpenOutside(cur()?.url() ?? active.url)}>
              Edge
            </Button>
          </Tooltip>
          <Tooltip content="Nastavenia konta" relationship="label">
            <Button size="small" appearance="subtle" icon={<SettingsIcon />} onClick={props.onEditAccount} aria-label="Nastavenia konta" />
          </Tooltip>
        </div>
      </div>
      <div className="m365-webs">
        {state.tabs.map((t) => (
          <WebSlot
            key={t.key}
            ref={(h) => {
              if (h) slots.current.set(t.key, h);
              else slots.current.delete(t.key);
            }}
            url={t.url}
            partition={partition}
            visible={t.key === active.key}
            onTitle={(title) => props.onTitle(t.key, title)}
            className="m365-web"
          />
        ))}
      </div>
    </section>
  );
}
