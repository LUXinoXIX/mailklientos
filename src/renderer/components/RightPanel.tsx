// Pravý panel: vložené weby (karty) a spúšťače programov.
import {
  Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input, MessageBar,
  MessageBarBody, Tooltip,
} from '@fluentui/react-components';
import { useEffect, useRef, useState } from 'react';
import { normalizeUrl, webPartition, type AppSettings, type LauncherItem, type WebItem } from '../../core/settings';
import { AddIcon, AppsIcon, CloseSmallIcon, EditSmallIcon, GlobeIcon, OpenOutsideIcon, ReloadIcon } from '../icons';
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
      <button className="launcher-btn" onClick={props.onRun} title={props.item.target} disabled={props.editing}>
        {icon ? <img src={icon} alt="" width={28} height={28} /> : <AppsIcon />}
        <span>{props.item.name}</span>
      </button>
      {props.editing && (
        <button className="launcher-x" onClick={props.onRemove} aria-label={`Odstrániť ${props.item.name}`}>
          <CloseSmallIcon />
        </button>
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
}) {
  const api = window.mailApi;
  const { settings } = props;
  const [editing, setEditing] = useState(false);
  const [addWeb, setAddWeb] = useState<{ name: string; url: string; error?: string } | null>(null);
  const web: WebItem | undefined = settings.webs.find((w) => w.id === props.active) ?? settings.webs[0];
  const slots = useRef(new Map<string, WebSlotHandle>());
  const cur = () => (web ? slots.current.get(web.id) : undefined);

  const openAdd = (open: boolean) => {
    setAddWeb(open ? { name: '', url: '' } : null);
    props.onOverlay(open);
  };

  function saveWeb() {
    if (!addWeb) return;
    const url = normalizeUrl(addWeb.url);
    if (!url) return setAddWeb({ ...addWeb, error: 'Zadaj platnú webovú adresu (napr. chatgpt.com).' });
    const name = addWeb.name.trim() || new URL(url).hostname.replace(/^www\./, '');
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
      <div className="right-tabs">
        {settings.webs.map((w) => (
          <div key={w.id} className={`right-tab ${web?.id === w.id ? 'on' : ''}`}>
            <button onClick={() => props.onActive(w.id)} title={w.url}>{w.name}</button>
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
        <Tooltip content="Pridať web" relationship="label">
          <Button size="small" appearance="subtle" icon={<AddIcon />} onClick={() => openAdd(true)} aria-label="Pridať web" />
        </Tooltip>
        <div className="grow" />
        <Tooltip content={editing ? 'Hotovo' : 'Upraviť weby a programy'} relationship="label">
          <Button size="small" appearance={editing ? 'primary' : 'subtle'} icon={<EditSmallIcon />} onClick={() => setEditing(!editing)} aria-label="Upraviť" />
        </Tooltip>
      </div>
      {web ? (
        <>
          <div className="right-url">
            <GlobeIcon />
            <span className="grow ellipsis" title={web.url}>{web.url.replace(/^https?:\/\//, '')}</span>
            <Button size="small" appearance="subtle" icon={<ReloadIcon />} onClick={() => cur()?.reload()} aria-label="Obnoviť" />
            <Button size="small" appearance="subtle" icon={<OpenOutsideIcon />} onClick={() => api?.webOpenOutside(cur()?.url() ?? web.url)} aria-label="Otvoriť v prehliadači" />
          </div>
          <div className="right-webs">
            {settings.webs.map((w) => (
              <WebSlot
                key={w.id}
                ref={(h) => {
                  if (h) slots.current.set(w.id, h);
                  else slots.current.delete(w.id);
                }}
                url={w.url}
                partition={webPartition(w.id)}
                visible={w.id === web.id}
                className="right-web"
              />
            ))}
          </div>
        </>
      ) : (
        <div className="empty grow">Pridaj web tlačidlom +</div>
      )}
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
          <button className="launcher-add" onClick={() => void addProgram()}>
            <AddIcon />
            <span>Program</span>
          </button>
        </div>
        <div className="muted small">Programy sa otvárajú vo vlastnom okne.</div>
      </div>

      {addWeb && (
        <Dialog open onOpenChange={(_, d) => !d.open && openAdd(false)}>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>Pridať web do pravého panela</DialogTitle>
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
                <Button appearance="primary" onClick={saveWeb}>Pridať</Button>
                <Button onClick={() => openAdd(false)}>Zrušiť</Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      )}
    </aside>
  );
}
