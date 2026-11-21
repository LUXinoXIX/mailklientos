// Nastavenia zámku programu (ozubené koliesko → Zámok programu…).
import { Button, Checkbox, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Dropdown, Field, Input, MessageBar, MessageBarBody, Option, Radio, RadioGroup } from '@fluentui/react-components';
import { useEffect, useState } from 'react';
import type { LockStatus } from '../../core/api';

const IDLE = [
  [0, 'Nikdy'],
  [5, '5 minút'],
  [15, '15 minút'],
  [30, '30 minút'],
  [60, '1 hodina'],
] as const;

export function LockSettings(props: { onClose: () => void }) {
  const api = window.mailApi!;
  const [st, setSt] = useState<LockStatus | null>(null);
  const [mode, setMode] = useState<'view' | 'set' | 'disable' | 'code'>('view');
  const [kind, setKind] = useState<'pin' | 'password'>('pin');
  const [cur, setCur] = useState('');
  const [n1, setN1] = useState('');
  const [n2, setN2] = useState('');
  const [code, setCode] = useState('');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  const refresh = () => void api.lockStatus().then((r) => r.ok && setSt(r.data));
  useEffect(refresh, []); // eslint-disable-line react-hooks/exhaustive-deps
  if (!st) return null;

  const digits = (v: string) => (kind === 'pin' ? v.replace(/\D/g, '') : v);
  const opt = (o: Parameters<typeof api.lockOptions>[0]) => void api.lockOptions(o).then(refresh);
  const save = async () => {
    if (n1 !== n2) return setError(kind === 'pin' ? 'PIN-y sa nezhodujú.' : 'Heslá sa nezhodujú.');
    setBusy(true);
    const r = await api.lockSet(kind, n1, st.enabled ? cur : undefined);
    setBusy(false);
    if (!r.ok) return setError(r.error);
    setCode(r.data);
    setMode('code');
    refresh();
  };
  const disable = async () => {
    setBusy(true);
    const r = await api.lockDisable(cur);
    setBusy(false);
    if (!r.ok) return setError(r.error);
    setMode('view');
    setCur('');
    refresh();
  };
  const start = (m: 'set' | 'disable') => {
    setError('');
    setCur('');
    setN1('');
    setN2('');
    setKind(st.enabled ? st.kind : 'pin');
    setMode(m);
  };
  const curField = st.enabled && (
    <Field label={st.kind === 'pin' ? 'Súčasný PIN' : 'Súčasné heslo'}>
      <Input type="password" value={cur} onChange={(_, d) => setCur(d.value)} autoFocus />
    </Field>
  );

  return (
    <Dialog open onOpenChange={(_, d) => !d.open && props.onClose()}>
      <DialogSurface style={{ maxWidth: 520 }}>
        <DialogBody>
          <DialogTitle>Zámok programu</DialogTitle>
          <DialogContent className="dialog-content">
            {mode === 'view' && (
              <>
                <Checkbox
                  label={st.enabled ? `Chrániť spustenie ${st.kind === 'pin' ? 'PIN-om' : 'heslom'} (zapnuté)` : 'Chrániť spustenie PIN-om alebo heslom'}
                  checked={st.enabled}
                  onChange={(_, d) => start(d.checked ? 'set' : 'disable')}
                />
                {st.enabled && (
                  <>
                    <Field label="Zamknúť po nečinnosti">
                      <Dropdown
                        value={IDLE.find(([m]) => m === st.idleMin)?.[1] ?? `${st.idleMin} min`}
                        selectedOptions={[String(st.idleMin)]}
                        onOptionSelect={(_, d) => opt({ idleMin: Number(d.optionValue) })}
                      >
                        {IDLE.map(([m, l]) => (
                          <Option key={m} value={String(m)}>
                            {l}
                          </Option>
                        ))}
                      </Dropdown>
                    </Field>
                    <Checkbox label="Zamknúť pri zamknutí obrazovky a uspaní počítača" checked={st.onScreenLock} onChange={(_, d) => opt({ onScreenLock: !!d.checked })} />
                    {st.touchIdAvailable && <Checkbox label="Odomykať aj Touch ID" checked={st.touchId} onChange={(_, d) => opt({ touchId: !!d.checked })} />}
                    <Checkbox label="Pri zamknutí skryť obsah upozornení (len „Nová pošta“)" checked={st.hideNotif} onChange={(_, d) => opt({ hideNotif: !!d.checked })} />
                    <div>
                      <Button onClick={() => start('set')}>Zmeniť PIN / heslo…</Button>
                    </div>
                    <span className="muted small">Zamknúť hneď: Ctrl+L alebo ozubené koliesko → Zamknúť teraz{navigator.platform.startsWith('Win') ? ', aj pravým klikom na ikonu pri hodinách' : ''}.</span>
                  </>
                )}
                <MessageBar intent="info">
                  <MessageBarBody>
                    Pri zabudnutom PIN-e alebo hesle sa dnu dostaneš heslom k niektorému e-mailovému kontu alebo záchranným kódom – nič sa nezmaže. Zámok chráni okno programu; pošta sa na pozadí sťahuje ďalej.
                  </MessageBarBody>
                </MessageBar>
              </>
            )}
            {mode === 'set' && (
              <>
                {curField}
                <RadioGroup value={kind} onChange={(_, d) => (setKind(d.value as 'pin' | 'password'), setN1(''), setN2(''))} layout="horizontal">
                  <Radio value="pin" label="PIN (4–12 číslic)" />
                  <Radio value="password" label="Heslo (aspoň 6 znakov)" />
                </RadioGroup>
                <Field label={kind === 'pin' ? 'Nový PIN' : 'Nové heslo'}>
                  <Input type="password" value={n1} onChange={(_, d) => setN1(digits(d.value))} input={{ inputMode: kind === 'pin' ? 'numeric' : undefined }} autoFocus={!st.enabled} />
                </Field>
                <Field label="Zopakuj">
                  <Input type="password" value={n2} onChange={(_, d) => setN2(digits(d.value))} input={{ inputMode: kind === 'pin' ? 'numeric' : undefined }} />
                </Field>
              </>
            )}
            {mode === 'disable' && (
              <>
                <span>Vypnúť zámok programu?</span>
                {curField}
              </>
            )}
            {mode === 'code' && (
              <>
                <span>
                  <b>Záchranný kód</b> – zapíš si ho alebo vytlač a odlož. Použiješ ho, ak zabudneš PIN/heslo a nebudeš vedieť heslo k žiadnemu e-mailovému kontu. Program ho už znova neukáže (pri každej zmene PIN-u/hesla dostaneš nový).
                </span>
                <div className="lock-code-light">{code}</div>
                <div>
                  <Button onClick={() => void navigator.clipboard.writeText(code)}>Kopírovať</Button>
                </div>
              </>
            )}
            {error && (
              <MessageBar intent="error">
                <MessageBarBody>{error}</MessageBarBody>
              </MessageBar>
            )}
          </DialogContent>
          <DialogActions>
            {mode === 'view' && <Button onClick={props.onClose}>Zavrieť</Button>}
            {mode === 'set' && (
              <>
                <Button onClick={() => setMode('view')}>Zrušiť</Button>
                <Button appearance="primary" disabled={busy || !n1 || !n2 || (st.enabled && !cur)} onClick={() => void save()}>
                  Uložiť
                </Button>
              </>
            )}
            {mode === 'disable' && (
              <>
                <Button onClick={() => setMode('view')}>Zrušiť</Button>
                <Button appearance="primary" disabled={busy || !cur} onClick={() => void disable()}>
                  Vypnúť zámok
                </Button>
              </>
            )}
            {mode === 'code' && (
              <Button appearance="primary" onClick={() => setMode('view')}>
                Mám zapísané
              </Button>
            )}
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
