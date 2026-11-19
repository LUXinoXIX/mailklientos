// Konto Microsoft 365: vložený Outlook na webe ako celé okno, s kartami.
// Office aplikácie otvorené z Outlooku (prepínač aplikácií) alebo z pása modulov idú do kariet, nie do nových okien.
import { Button, Tooltip } from '@fluentui/react-components';
import { accountPartition, type M365Account } from '../../core/account';
import { BackIcon, CloseSmallIcon, OpenOutsideIcon, ReloadIcon, SettingsIcon } from '../icons';
import type { M365Module } from './Shell';
import { useEffect, useRef, useState } from 'react';
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
export function moduleUrl(m: M365Module, host: 'office' | 'live' | 'web' = 'office', webUrl?: string): string {
  if (host === 'web') return webUrl || 'about:blank';
  const o = host === 'office';
  switch (m) {
    case 'mail': return o ? 'https://outlook.office.com/mail/' : 'https://outlook.live.com/mail/';
    case 'calendar': return o ? 'https://outlook.office.com/calendar/' : 'https://outlook.live.com/calendar/';
    case 'people': return o ? 'https://outlook.office.com/people/' : 'https://outlook.live.com/people/';
    case 'todo': return o ? 'https://to-do.office.com/tasks/' : 'https://to-do.live.com/tasks/';
    case 'teams': return o ? 'https://teams.microsoft.com/' : 'https://teams.live.com/';
    case 'onedrive': return o ? 'https://www.office.com/launch/onedrive?auth=2' : 'https://onedrive.live.com/';
    // auth=2 = pracovné/školské konto, auth=1 = osobné konto Microsoft
    case 'word': return `https://www.office.com/launch/word?auth=${o ? 2 : 1}`;
    case 'excel': return `https://www.office.com/launch/excel?auth=${o ? 2 : 1}`;
    case 'powerpoint': return `https://www.office.com/launch/powerpoint?auth=${o ? 2 : 1}`;
    case 'copilot': return o ? 'https://m365.cloud.microsoft/chat' : 'https://copilot.microsoft.com/';
    case 'planner': return 'https://tasks.office.com/'; // klasický Planner, len firemné kontá
    case 'sharepoint': return 'https://www.office.com/launch/sharepoint?auth=2';
    case 'onenote': return o ? 'https://www.office.com/launch/onenote?auth=2' : 'https://www.onenote.com/notebooks';
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
  return { tabs: [{ key, title: acc.host === 'web' ? 'Webmail' : 'Outlook', url: moduleUrl('mail', acc.host, acc.webUrl), module: 'mail' }], active: key };
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
  /** počet neprečítaných v Doručenej pošte vloženého Outlooku */
  onUnread?: (n: number) => void;
}) {
  const api = window.mailApi;
  const { account, state } = props;
  const partition = accountPartition(account.id);
  const active = state.tabs.find((t) => t.key === state.active) ?? state.tabs[0]!;
  const slots = useRef(new Map<string, WebSlotHandle>());
  const cur = () => slots.current.get(active.key);
  const [curUrl, setCurUrl] = useState(active.url);
  useEffect(() => {
    const t = window.setInterval(() => setCurUrl(cur()?.url() ?? active.url), 3000);
    return () => window.clearInterval(t);
  }); // eslint-disable-line react-hooks/exhaustive-deps
  const isDoc = /sharepoint\.com|onedrive|:[wxp]:\/|Doc\.aspx|\.(docx?|xlsx?|pptx?)(\?|$)/i.test(curUrl) && !/outlook\./i.test(curUrl);

  // počet neprečítaných z Outlooku na webe (priečinok Doručená pošta v strome priečinkov)
  const onUnread = useRef(props.onUnread);
  onUnread.current = props.onUnread;
  useEffect(() => {
    const first = state.tabs[0];
    if (!first) return;
    const js = `(() => {
      const items = [...document.querySelectorAll('[role="treeitem"]')];
      const inbox = items.find((e) => /^(Doručená pošta|Doručené|Inbox|Posteingang|Doručená)/i.test((e.getAttribute('aria-label') || e.textContent || '').trim()));
      if (!inbox) { const m = /\((\d+)\)/.exec(document.title); return m ? Number(m[1]) : -1; }
      const label = inbox.getAttribute('aria-label') || '';
      const m = /(\d+)\s*(neprečítan|unread|ungelesen)/i.exec(label) || /(\d+)/.exec(inbox.textContent || '');
      return m ? Number(m[1]) : 0;
    })()`;
    const run = async () => {
      const n = await slots.current.get(first.key)?.exec<number>(js);
      if (typeof n === 'number' && n >= 0) onUnread.current?.(n);
    };
    const t = window.setInterval(() => void run(), 30_000);
    const t0 = window.setTimeout(() => void run(), 8000);
    return () => {
      window.clearInterval(t);
      window.clearTimeout(t0);
    };
  }, [state.tabs[0]?.key]); // eslint-disable-line react-hooks/exhaustive-deps

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
          {isDoc && (
            <Tooltip content="Otvoriť tento dokument v nainštalovanom Office (Word, Excel, PowerPoint)" relationship="label">
              <Button size="small" appearance="subtle" onClick={() => void api?.openInOffice(curUrl)}>Otvoriť v Office</Button>
            </Tooltip>
          )}
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
