// Online schôdza v udalosti: bez / Teams (cez zvolené Outlook konto) / Jitsi / vlastný odkaz.
// Teams: Microsoft neumožňuje vytvoriť schôdzu cez API pre osobné kontá a firemné vyžaduje registráciu aplikácie,
// preto program otvorí Teams s prihlásením konta, používateľ klikne „Vytvoriť odkaz na schôdzu → Kopírovať“
// a program odkaz zo schránky zachytí sám.
import { Button, Field, Input, Select, Spinner } from '@fluentui/react-components';
import { useEffect, useRef, useState } from 'react';
import { DEFAULT_JITSI, MEETING_LABEL, meetingKindOf, newJitsiUrl, teamsLinkIn, type MeetingSettings } from '../../core/meeting';

export interface MeetOutlook {
  id: string;
  name: string;
  personal: boolean;
}
type Choice = 'none' | 'teams' | 'jitsi' | 'link';

export function MeetingPicker(props: {
  url: string | null;
  title: string;
  isNew: boolean;
  settings: MeetingSettings;
  outlooks: MeetOutlook[];
  onUrl: (url: string | null) => void;
  onSettings: (s: MeetingSettings) => void;
}) {
  const api = window.mailApi;
  const server = props.settings.jitsiServer || DEFAULT_JITSI;
  const kindOf = (u: string | null): Choice => {
    if (!u) return 'none';
    const k = meetingKindOf(u, server);
    return k === 'teams' || k === 'teams-personal' ? 'teams' : k === 'jitsi' ? 'jitsi' : 'link';
  };
  const [choice, setChoice] = useState<Choice>(() => (props.url ? kindOf(props.url) : props.isNew ? props.settings.default ?? 'none' : 'none'));
  const [waiting, setWaiting] = useState(false);
  const [msg, setMsg] = useState('');
  const teamsAcc = props.outlooks.find((o) => o.id === props.settings.teamsAccountId) ?? props.outlooks[0];
  const poll = useRef<number | null>(null);

  // nová udalosť s predvoleným Jitsi → miestnosť hneď
  useEffect(() => {
    if (props.isNew && !props.url && choice === 'jitsi') props.onUrl(newJitsiUrl(server, props.title));
    return () => stopPoll();
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  function stopPoll() {
    if (poll.current) window.clearInterval(poll.current);
    poll.current = null;
    setWaiting(false);
  }

  function pick(c: Choice) {
    setChoice(c);
    setMsg('');
    stopPoll();
    if (c === 'none') props.onUrl(null);
    if (c === 'jitsi') props.onUrl(newJitsiUrl(server, props.title));
    if (c === 'teams' || c === 'link') props.onUrl(props.url && kindOf(props.url) === c ? props.url : null);
  }

  async function createTeams() {
    if (!api || !teamsAcc) return;
    setMsg('');
    const before = await api.clipboardRead();
    const initial = before.ok ? before.data : '';
    const r = await api.meetTeams(teamsAcc.id);
    if (!r.ok) return setMsg(r.error);
    setWaiting(true);
    const started = Date.now();
    poll.current = window.setInterval(async () => {
      if (Date.now() - started > 5 * 60_000) {
        stopPoll();
        return setMsg('Odkaz sa nezachytil. Skopíruj ho v Teams a vlož ho tlačidlom „Vložiť zo schránky“.');
      }
      const c = await api.clipboardRead();
      const link = c.ok && c.data !== initial ? teamsLinkIn(c.data) : null;
      if (link) {
        stopPoll();
        props.onUrl(link);
        setMsg('Odkaz na Teams zachytený.');
        void api.meetTeamsClose();
      }
    }, 1000);
  }

  async function paste() {
    const c = await api?.clipboardRead();
    const t = c?.ok ? c.data.trim() : '';
    const link = choice === 'teams' ? teamsLinkIn(t) : /^https:\/\/\S+$/.test(t) ? t : null;
    if (link) {
      props.onUrl(link);
      setMsg('');
    } else setMsg(choice === 'teams' ? 'V schránke nie je odkaz na Teams schôdzu.' : 'V schránke nie je odkaz (https://…).');
  }

  return (
    <div className="meet">
      <div className="meet-row">
        <Field label="Online schôdza" className="grow">
          <Select value={choice} onChange={(_, d) => pick(d.value as Choice)}>
            <option value="none">Bez online schôdze</option>
            <option value="teams" disabled={!props.outlooks.length}>
              Microsoft Teams{teamsAcc ? ` – cez konto ${teamsAcc.name}${teamsAcc.personal ? ' (osobné)' : ''}` : ' (pridaj konto Outlook)'}
            </option>
            <option value="jitsi">Jitsi Meet – {server.replace(/^https:\/\//, '')}</option>
            <option value="link">Iný odkaz (Zoom, Google Meet…)</option>
          </Select>
        </Field>
        {choice === 'teams' && props.outlooks.length > 1 && (
          <Field label="Teams cez konto">
            <Select value={teamsAcc?.id} onChange={(_, d) => props.onSettings({ ...props.settings, teamsAccountId: d.value })}>
              {props.outlooks.map((o) => <option key={o.id} value={o.id}>{o.name}{o.personal ? ' (osobné)' : ' (firemné)'}</option>)}
            </Select>
          </Field>
        )}
      </div>

      {choice === 'teams' && !props.url && (
        <div className="meet-help">
          <div className="row gap" style={{ flexWrap: 'wrap' }}>
            <Button appearance="primary" size="small" onClick={() => void createTeams()} disabled={waiting || !teamsAcc}>Vytvoriť odkaz v Teams…</Button>
            <Button size="small" onClick={() => void paste()}>Vložiť zo schránky</Button>
            {waiting && <span className="row gap muted small"><Spinner size="extra-tiny" /> čakám na skopírovaný odkaz…</span>}
          </div>
          <div className="muted small">
            Otvorí sa Teams prihlásený ako {teamsAcc?.name ?? 'konto Outlook'}. Tam klikni na <b>Kalendár → Schôdza teraz / Vytvoriť odkaz na schôdzu</b> (v osobných Teams <b>Schôdza → Vytvoriť odkaz</b>) a potom <b>Kopírovať odkaz</b>. Program ho zachytí sám a okno Teams zavrie.
          </div>
        </div>
      )}
      {choice === 'link' && !props.url && (
        <div className="row gap">
          <Input className="grow" placeholder="https://zoom.us/j/… alebo https://meet.google.com/…" onChange={(_, d) => /^https:\/\/\S+$/.test(d.value.trim()) && props.onUrl(d.value.trim())} />
          <Button size="small" onClick={() => void paste()}>Vložiť zo schránky</Button>
        </div>
      )}
      {props.url && (
        <div className="meet-link">
          <span className="meet-kind">{MEETING_LABEL[meetingKindOf(props.url, server)]}</span>
          <a className="meet-url" href={props.url} title={props.url} onClick={(ev) => {
            ev.preventDefault();
            if (kindOf(props.url) === 'jitsi') void api?.meetJitsi(props.url!);
            else void api?.openExternal?.(props.url!);
          }}>{props.url}</a>
          {choice === 'jitsi' && <Button size="small" appearance="subtle" onClick={() => props.onUrl(newJitsiUrl(server, props.title))} title="Nová miestnosť">Nová</Button>}
          {choice === 'jitsi' && <Button size="small" onClick={() => void api?.meetJitsi(props.url!)}>Spustiť</Button>}
          <Button size="small" appearance="subtle" onClick={() => { props.onUrl(null); setMsg(''); }} title="Odstrániť odkaz">×</Button>
        </div>
      )}
      {choice === 'jitsi' && (
        <details className="meet-adv">
          <summary className="muted small">Jitsi server</summary>
          <div className="row gap" style={{ alignItems: 'center', marginTop: 6 }}>
            <Input size="small" className="grow" defaultValue={props.settings.jitsiServer ?? ''} placeholder={DEFAULT_JITSI} onBlur={(ev) => {
              const v = ev.target.value.trim().replace(/\/+$/, '');
              if (!v || /^https:\/\/[\w.-]+(:\d+)?(\/[\w.-]*)*$/.test(v)) props.onSettings({ ...props.settings, jitsiServer: v || undefined });
            }} />
          </div>
          <div className="muted small">Na meet.jit.si sa organizátor pri začatí schôdze prihlási cez Google, GitHub alebo Facebook; ostatní účastníci nie. Firemný alebo vlastný Jitsi server to nevyžaduje.</div>
        </details>
      )}
      {choice !== 'none' && (
        <label className="meet-def muted small">
          <input type="checkbox" checked={(props.settings.default ?? 'none') === choice} disabled={choice === 'link'} onChange={(ev) => props.onSettings({ ...props.settings, default: ev.target.checked ? (choice as 'teams' | 'jitsi') : 'none' })} />{' '}
          Predvolene pri nových udalostiach
        </label>
      )}
      {msg && <div className={msg.startsWith('Odkaz na Teams zachytený') ? 'muted small' : 'pane-error'}>{msg}</div>}
    </div>
  );
}
