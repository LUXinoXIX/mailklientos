// Pravý panel: vložené weby (karty) a spúšťače programov.
import {
  Menu, MenuDivider, MenuGroupHeader, MenuItem, MenuItemRadio, MenuList, MenuPopover, MenuTrigger,
  Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input, MessageBar,
  MessageBarBody, Tooltip,
} from '@fluentui/react-components';
import { useEffect, useRef, useState } from 'react';
import { normalizeUrl, webPartition, type AppSettings, type LauncherItem, type Task, type WebItem } from '../../core/settings';
import type { CalEvent } from '../../core/ical';
import { IcsPane } from './IcsPane';
import { AddIcon, AppsIcon, ChevronDownIcon, PanelIcon, CloseSmallIcon, EditSmallIcon, GlobeIcon, OpenOutsideIcon, ReloadIcon } from '../icons';
import { WebSlot, type WebSlotHandle } from './WebSlot';

const newId = () => Math.random().toString(36).slice(2, 10);

function LauncherButton(props: { item: LauncherItem; editing: boolean; onRun: () => void; onRemove: () => void }) {
  const [icon, setIcon] = useState<string | null>(null);
  useEffect(() => {
    let alive = true;
    void window.mailApi?.programIcon(props.item.target).then((r) => alive && r.ok && setIcon(r.data));
    return () => {
      alive = false;
    };
  }, [props.item.target]);
  return (
    <div className="launcher">
      <Tooltip content={props.item.name} relationship="label">
        <button className="launcher-btn" onClick={props.onRun} aria-label={props.item.name}>
          {icon ? <img src={icon} alt="" width={22} height={22} /> : <AppsIcon />}
        </button>
      </Tooltip>
      {props.editing && (
        <button className="launcher-x" onClick={props.onRemove} aria-label={`Odstrániť ${props.item.name}`}>
          <CloseSmallIcon />
        </button>
      )}
    </div>
  );
}

export interface AgendaData {
  loading: boolean;
  error?: string;
  events: Array<CalEvent & { accountName: string; accountColor: string }>;
  tasks: Task[];
  onOpenTask: (t: Task) => void;
  onOpenCalendar: () => void;
  onToggleTask: (t: Task) => void;
}

/** Dnešný deň: stretnutia zo všetkých kalendárov a úlohy na dnes / po termíne. */
function Agenda({ data }: { data: AgendaData }) {
  const now = Date.now();
  const time = (iso: string) => new Date(iso).toLocaleTimeString('sk-SK', { hour: '2-digit', minute: '2-digit' });
  const today = new Date().toLocaleDateString('sk-SK', { weekday: 'long', day: 'numeric', month: 'long' });
  return (
    <div className="agenda">
      <div className="agenda-day">{today}</div>
      <div className="folder-section">Stretnutia</div>
      {data.loading && <div className="muted small pad">Načítavam…</div>}
      {data.error && <div className="muted small pad">⚠ {data.error}</div>}
      {!data.loading && !data.events.length && <div className="muted small pad">Dnes žiadne stretnutia.</div>}
      {data.events.map((e) => {
        const past = Date.parse(e.end) < now;
        const live = Date.parse(e.start) <= now && Date.parse(e.end) > now;
        return (
          <button key={`${e.uid}-${e.start}`} className={`agenda-item ${past ? 'past' : ''} ${live ? 'live' : ''}`} style={{ borderLeftColor: e.color ?? e.accountColor }} onClick={data.onOpenCalendar}>
            <span className="agenda-time">{e.allDay ? 'celý deň' : `${time(e.start)}–${time(e.end)}`}</span>
            <b>{e.summary}</b>
            {e.location && <span className="muted small">{e.location}</span>}
          </button>
        );
      })}
      <div className="folder-section">Úlohy na dnes</div>
      {!data.tasks.length && <div className="muted small pad">Nič na dnes.</div>}
      {data.tasks.map((t) => (
        <div key={t.id} className="agenda-task">
          <input type="checkbox" checked={t.done} onChange={() => data.onToggleTask(t)} aria-label="Hotové" />
          <button className="link" onClick={() => data.onOpenTask(t)}>{t.title}</button>
          {t.due && t.due < new Date().toISOString().slice(0, 10) && <span className="overdue small">po termíne</span>}
        </div>
      ))}
    </div>
  );
}

export interface PanelChoice {
  /** meno aktuálneho konta (alebo „Všetky kontá“) */
  accountName: string;
  current: string;
  panels: Array<{ id: string; name: string; users: string[] }>;
  sharedUsers: string[];
  onChoose: (id: string) => void;
  onNew: (copy: boolean) => void;
  onRename: (id: string, name: string) => void;
  onDelete: (id: string) => void;
}

function PanelChooser({ c, onOverlay }: { c: PanelChoice; onOverlay: (open: boolean) => void }) {
  const [rename, setRename] = useState<{ id: string; name: string } | null>(null);
  const cur = c.panels.find((p) => p.id === c.current);
  const label = cur ? cur.name : 'Spoločný';
  const who = (u: string[]) => (u.length ? u.join(', ') : 'žiadne konto');
  const openRename = (r: { id: string; name: string } | null) => {
    setRename(r);
    onOverlay(!!r);
  };
  return (
    <div className="panel-chooser">
      <PanelIcon className="panel-chooser-ico" />
      <Menu
        checkedValues={{ panel: [c.current] }}
        onCheckedValueChange={(_, d) => d.checkedItems[0] && c.onChoose(d.checkedItems[0])}
        onOpenChange={(_, d) => onOverlay(d.open)}
      >
        <MenuTrigger disableButtonEnhancement>
          <button className="panel-chooser-btn" title={`Pravý panel pre: ${c.accountName}`}>
            <span className="muted small">Panel:</span> <b>{label}</b> <ChevronDownIcon />
          </button>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuGroupHeader>Panel pre {c.accountName}</MenuGroupHeader>
            <MenuItemRadio name="panel" value="shared" secondaryContent={who(c.sharedUsers)}>
              Spoločný
            </MenuItemRadio>
            {c.panels.map((p) => (
              <MenuItemRadio key={p.id} name="panel" value={p.id} secondaryContent={who(p.users)}>
                {p.name}
              </MenuItemRadio>
            ))}
            <MenuDivider />
            <MenuItem icon={<AddIcon />} onClick={() => c.onNew(false)}>Nový vlastný panel (prázdny)</MenuItem>
            <MenuItem onClick={() => c.onNew(true)}>Nový vlastný panel ako kópia tohto</MenuItem>
            {cur && (
              <>
                <MenuDivider />
                <MenuItem onClick={() => openRename({ id: cur.id, name: cur.name })}>Premenovať „{cur.name}“…</MenuItem>
                <MenuItem onClick={() => c.onDelete(cur.id)}>Odstrániť „{cur.name}“ (kontá prejdú na spoločný)</MenuItem>
              </>
            )}
          </MenuList>
        </MenuPopover>
      </Menu>
      {rename && (
        <Dialog open onOpenChange={(_, d) => !d.open && openRename(null)}>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>Premenovať panel</DialogTitle>
              <DialogContent className="dialog-content">
                <Field label="Názov">
                  <Input autoFocus value={rename.name} onChange={(_, d) => setRename({ ...rename, name: d.value })} />
                </Field>
              </DialogContent>
              <DialogActions>
                <Button
                  appearance="primary"
                  onClick={() => {
                    if (rename.name.trim()) c.onRename(rename.id, rename.name.trim());
                    openRename(null);
                  }}
                >
                  Uložiť
                </Button>
                <Button onClick={() => openRename(null)}>Zrušiť</Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      )}
    </div>
  );
}

export function RightPanel(props: {
  settings: AppSettings;
  onSettings: (s: AppSettings) => void;
  active: string;
  onActive: (id: string) => void;
  overlay: boolean;
  onOverlay: (open: boolean) => void;
  onStatus: (s: string) => void;
  /** dnešné udalosti a úlohy (karta Agenda) */
  agenda?: AgendaData;
  /** weby všetkých panelov – ostávajú načítané aj pri prepnutí konta */
  allWebs?: WebItem[];
  /** výber panela pre aktuálne konto (spoločný alebo vlastný) */
  choice?: PanelChoice;
}) {
  const api = window.mailApi;
  const { settings } = props;
  const [editing, setEditing] = useState(false);
  const [addWeb, setAddWeb] = useState<{ name: string; url: string; error?: string; /** úprava existujúcej karty */ editId?: string } | null>(null);
  // karta Kalendáre (ICS) – po pridaní cez + sa zobrazí aj kým ešte nemá žiadny kalendár
  const [icsAdded, setIcsAdded] = useState(false);
  const showIcs = !!settings.icsFeeds || icsAdded;
  const icsOn = props.active === 'ics' && (!!settings.icsFeeds || icsAdded);
  const agendaOnly = props.active === 'agenda' && !!props.agenda && !icsOn;
  // agendaOn = nezobrazuje sa web (Agenda alebo Kalendáre)
  const agendaOn = agendaOnly || icsOn;
  const web: WebItem | undefined = agendaOn ? undefined : settings.webs.find((w) => w.id === props.active) ?? settings.webs[0];
  const slots = useRef(new Map<string, WebSlotHandle>());
  const cur = () => (web ? slots.current.get(web.id) : undefined);

  const openAdd = (open: boolean) => {
    setAddWeb(open ? { name: '', url: '' } : null);
    props.onOverlay(open);
  };

  const editWeb = (w: { id: string; name: string; url: string }) => {
    setAddWeb({ name: w.name, url: w.url, editId: w.id });
    props.onOverlay(true);
  };
  const moveWeb = (i: number, d: number) => {
    const l = [...settings.webs];
    const [x] = l.splice(i, 1);
    l.splice(i + d, 0, x!);
    props.onSettings({ ...settings, webs: l });
  };

  function saveWeb() {
    if (!addWeb) return;
    const url = normalizeUrl(addWeb.url);
    if (!url) return setAddWeb({ ...addWeb, error: 'Zadaj platnú webovú adresu (napr. chatgpt.com).' });
    const name = addWeb.name.trim() || new URL(url).hostname.replace(/^www\./, '');
    if (addWeb.editId) {
      props.onSettings({ ...settings, webs: settings.webs.map((x) => (x.id === addWeb.editId ? { ...x, name, url } : x)) });
      return openAdd(false);
    }
    const item = { id: newId(), name, url };
    props.onSettings({ ...settings, webs: [...settings.webs, item] });
    props.onActive(item.id);
    openAdd(false);
  }

  async function addProgram() {
    if (!api) return;
    const r = await api.pickProgram();
    if (!r.ok) return props.onStatus(r.error);
    if (!r.data) return;
    props.onSettings({ ...settings, launchers: [...settings.launchers, { id: newId(), name: r.data.name, target: r.data.target }] });
  }

  async function run(l: LauncherItem) {
    const r = await api?.runLauncher(l);
    if (r && (!r.ok || r.data)) props.onStatus(r.ok ? r.data : r.error);
    else props.onStatus(`Spúšťam ${l.name}…`);
  }

  return (
    <aside className="right" aria-label="Pravý panel">
      {props.choice && <PanelChooser c={props.choice} onOverlay={props.onOverlay} />}
      <div className="right-tabs">
        {props.agenda && (
          <div className={`right-tab ${agendaOnly ? 'on' : ''}`}>
            <button onClick={() => props.onActive('agenda')}>Agenda</button>
          </div>
        )}
        {showIcs && (
          <div className={`right-tab ${icsOn ? 'on' : ''}`}>
            <button onClick={() => props.onActive('ics')} title="Kalendáre z odkazu (ICS)">Kalendáre</button>
            {editing && (
              <button className="right-tab-x" aria-label="Odstrániť kartu Kalendáre" onClick={() => { setIcsAdded(false); props.onSettings({ ...settings, icsFeeds: undefined }); if (icsOn) props.onActive('agenda'); }}>
                <CloseSmallIcon />
              </button>
            )}
          </div>
        )}
        {settings.webs.map((w, wi) => (
          <div key={w.id} className={`right-tab ${web?.id === w.id ? 'on' : ''}`}>
            {/* pravý klik: upraviť (názov, adresa), presunúť, odstrániť */}
            <Menu openOnContext>
              <MenuTrigger disableButtonEnhancement>
                <button onClick={() => props.onActive(w.id)} onDoubleClick={() => editWeb(w)} title={`${w.url}\nPravý klik: upraviť, presunúť, odstrániť`}>{w.name}</button>
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  <MenuItem onClick={() => editWeb(w)}>Upraviť názov a adresu…</MenuItem>
                  <MenuItem disabled={wi === 0} onClick={() => moveWeb(wi, -1)}>Presunúť vľavo</MenuItem>
                  <MenuItem disabled={wi === settings.webs.length - 1} onClick={() => moveWeb(wi, 1)}>Presunúť vpravo</MenuItem>
                  <MenuItem onClick={() => props.onSettings({ ...settings, webs: settings.webs.filter((x) => x.id !== w.id) })}>Odstrániť kartu</MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
            {editing && (
              <button
                className="right-tab-x"
                aria-label={`Odstrániť ${w.name}`}
                onClick={() => {
                  props.onSettings({ ...settings, webs: settings.webs.filter((x) => x.id !== w.id) });
                }}
              >
                <CloseSmallIcon />
              </button>
            )}
          </div>
        ))}
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <Tooltip content="Pridať kartu" relationship="label">
              <Button size="small" appearance="subtle" icon={<AddIcon />} aria-label="Pridať kartu" />
            </Tooltip>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem icon={<GlobeIcon />} onClick={() => openAdd(true)}>Web podľa adresy…</MenuItem>
              <MenuItem disabled={showIcs} onClick={() => { setIcsAdded(true); props.onActive('ics'); }}>Kalendáre z odkazu (ICS){showIcs ? ' – už je pridaná' : '…'}</MenuItem>
              <MenuItem icon={<AppsIcon />} onClick={() => void addProgram()}>Program (.exe)…</MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>
        <div className="grow" />
        <Tooltip content={editing ? 'Hotovo' : 'Upraviť weby a programy'} relationship="label">
          <Button size="small" appearance={editing ? 'primary' : 'subtle'} icon={<EditSmallIcon />} onClick={() => setEditing(!editing)} aria-label="Upraviť" />
        </Tooltip>
      </div>
      {icsOn ? (
        <IcsPane feeds={settings.icsFeeds ?? []} onFeeds={(f) => props.onSettings({ ...settings, icsFeeds: f })} onStatus={props.onStatus} />
      ) : agendaOn && props.agenda ? (
        <Agenda data={props.agenda} />
      ) : null}
      {/* weby ostávajú načítané aj pri zobrazenej Agende a pri prepnutí na iný panel */}
      {web && !agendaOn && (
        <div className="right-url">
          <GlobeIcon />
          <span className="grow ellipsis" title={web.url}>{web.url.replace(/^https?:\/\//, '')}</span>
          <Button size="small" appearance="subtle" icon={<ReloadIcon />} onClick={() => cur()?.reload()} aria-label="Obnoviť" />
          <Button size="small" appearance="subtle" icon={<OpenOutsideIcon />} onClick={() => api?.webOpenOutside(cur()?.url() ?? web.url)} aria-label="Otvoriť v prehliadači" />
        </div>
      )}
      {!web && !agendaOn && <div className="empty grow">Pridaj web tlačidlom +</div>}
      <div className="right-webs" style={agendaOn || !web ? { display: 'none' } : undefined}>
        {(props.allWebs ?? settings.webs).map((w) => (
          <WebSlot
            key={w.id}
            ref={(h) => {
              if (h) slots.current.set(w.id, h);
              else slots.current.delete(w.id);
            }}
            url={w.url}
            partition={webPartition(w.id)}
            visible={!agendaOn && w.id === web?.id}
            className="right-web"
          />
        ))}
      </div>
      <div className="launchers">
        <div className="launchers-title">Programy</div>
        <div className="launchers-grid">
          {settings.launchers.map((l) => (
            <LauncherButton
              key={l.id}
              item={l}
              editing={editing}
              onRun={() => void run(l)}
              onRemove={() => props.onSettings({ ...settings, launchers: settings.launchers.filter((x) => x.id !== l.id) })}
            />
          ))}
          <Tooltip content="Pridať program (otvára sa vo vlastnom okne)" relationship="label">
            <button className="launcher-add" onClick={() => void addProgram()} aria-label="Pridať program">
              <AddIcon />
            </button>
          </Tooltip>
        </div>
      </div>

      {addWeb && (
        <Dialog open onOpenChange={(_, d) => !d.open && openAdd(false)}>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>{addWeb.editId ? 'Upraviť kartu' : 'Pridať web do pravého panela'}</DialogTitle>
              <DialogContent className="dialog-content">
                <Field label="Adresa">
                  <Input autoFocus placeholder="napr. chatgpt.com" value={addWeb.url} onChange={(_, d) => setAddWeb({ ...addWeb, url: d.value, error: undefined })} />
                </Field>
                <Field label="Názov (nepovinné)">
                  <Input value={addWeb.name} onChange={(_, d) => setAddWeb({ ...addWeb, name: d.value })} />
                </Field>
                <div className="muted small">Každý web má vlastné prihlásenie, oddelené od ostatných.</div>
                {addWeb.error && <MessageBar intent="error"><MessageBarBody>{addWeb.error}</MessageBarBody></MessageBar>}
              </DialogContent>
              <DialogActions>
                <Button appearance="primary" onClick={saveWeb}>{addWeb.editId ? 'Uložiť' : 'Pridať'}</Button>
                <Button onClick={() => openAdd(false)}>Zrušiť</Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      )}
    </aside>
  );
}
