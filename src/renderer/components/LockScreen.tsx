// Zamknutá obrazovka – prekryje celé okno (hlavné aj samostatné okná písania/správ, aj vložený Outlook).
import { useEffect, useRef, useState } from 'react';
import type { LockStatus } from '../../core/api';

type Step = 'unlock' | 'recover' | 'reset' | 'code';

export function LockGate() {
  const api = window.mailApi;
  const [st, setSt] = useState<LockStatus | null>(null);
  const refresh = () => void api?.lockStatus().then((r) => r.ok && setSt(r.data));
  useEffect(() => {
    refresh();
    const off = api?.onAppEvent((ev) => ev.type === 'lock-changed' && refresh());
    // Ctrl+L / Cmd+L = zamknúť hneď
    const key = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && !e.shiftKey && !e.altKey && e.key.toLowerCase() === 'l') {
        e.preventDefault();
        void api?.lockNow();
      }
    };
    window.addEventListener('keydown', key, true);
    return () => {
      off?.();
      window.removeEventListener('keydown', key, true);
    };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  useEffect(() => {
    document.body.classList.toggle('app-locked', !!st?.locked);
  }, [st?.locked]);
  if (!st?.locked) return null;
  return <LockScreen st={st} onChange={refresh} />;
}

function LockScreen(props: { st: LockStatus; onChange: () => void }) {
  const api = window.mailApi!;
  const { st } = props;
  const [step, setStep] = useState<Step>('unlock');
  const [secret, setSecret] = useState('');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  // obnova
  const [method, setMethod] = useState<'account' | 'code'>(st.recoveryAccounts.length ? 'account' : 'code');
  const [accId, setAccId] = useState(st.recoveryAccounts[0]?.id ?? '');
  const [accPw, setAccPw] = useState('');
  const [code, setCode] = useState('');
  const [token, setToken] = useState('');
  const [kind, setKind] = useState<'pin' | 'password'>(st.kind);
  const [new1, setNew1] = useState('');
  const [new2, setNew2] = useState('');
  const [newCode, setNewCode] = useState('');
  const input = useRef<HTMLInputElement>(null);

  useEffect(() => {
    setError('');
    setTimeout(() => input.current?.focus(), 50);
  }, [step]);
  useEffect(() => {
    if (st.touchId && step === 'unlock' && document.hasFocus()) void api.lockTouchId().then((r) => r.ok && props.onChange());
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const run = async (f: () => Promise<{ ok: true; data?: unknown } | { ok: false; error: string }>, done: (d: unknown) => void) => {
    setBusy(true);
    setError('');
    const r = await f();
    setBusy(false);
    if (!r.ok) return setError(r.error);
    done(r.data);
  };

  const unlock = () => run(() => api.lockUnlock(secret), () => props.onChange());
  const recover = () =>
    run(
      () => (method === 'account' ? api.lockRecoverAccount(accId, accPw) : api.lockRecoverCode(code)),
      (t) => {
        setToken(t as string);
        setStep('reset');
      },
    );
  const reset = () => {
    if (new1 !== new2) return setError(kind === 'pin' ? 'PIN-y sa nezhodujú.' : 'Heslá sa nezhodujú.');
    void run(
      () => api.lockReset(token, kind, new1),
      (c) => {
        setNewCode(c as string);
        setStep('code');
      },
    );
  };

  const pinProps = (k: 'pin' | 'password') => (k === 'pin' ? { inputMode: 'numeric' as const, pattern: '[0-9]*', maxLength: 12 } : {});

  return (
    <div className="lock-screen" role="dialog" aria-modal="true" aria-label="Program je zamknutý">
      <form
        className="lock-box"
        onSubmit={(e) => {
          e.preventDefault();
          if (busy) return;
          if (step === 'unlock') void unlock();
          else if (step === 'recover') void recover();
          else if (step === 'reset') reset();
          else void run(() => api.lockUnlock(new1), () => props.onChange());
        }}
      >
        <div className="lock-ico" aria-hidden="true">
          <svg viewBox="0 0 24 24" width="24" height="24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
            <rect x="5" y="11" width="14" height="10" rx="2" />
            <path d="M8 11V7a4 4 0 0 1 8 0v4" />
          </svg>
        </div>
        {step === 'unlock' && (
          <>
            <h1>Universum eMail je zamknutý</h1>
            <input
              ref={input}
              type="password"
              autoComplete="off"
              placeholder={st.kind === 'pin' ? 'PIN' : 'Heslo'}
              value={secret}
              onChange={(e) => setSecret(st.kind === 'pin' ? e.target.value.replace(/\D/g, '') : e.target.value)}
              aria-label={st.kind === 'pin' ? 'PIN' : 'Heslo'}
              {...pinProps(st.kind)}
            />
            <div className="lock-row">
              <button type="submit" className="lock-primary" disabled={busy || !secret}>
                Odomknúť
              </button>
              {st.touchId && (
                <button type="button" onClick={() => void run(() => api.lockTouchId(), () => props.onChange())}>
                  Touch ID
                </button>
              )}
            </div>
            <button type="button" className="lock-link" onClick={() => setStep('recover')}>
              Zabudnutý {st.kind === 'pin' ? 'PIN' : 'heslo'}?
            </button>
          </>
        )}
        {step === 'recover' && (
          <>
            <h1>Obnovenie prístupu</h1>
            <p className="lock-note">Nič sa nezmaže – po overení si nastavíš nový PIN alebo heslo.</p>
            <div className="lock-tabs" role="radiogroup">
              {st.recoveryAccounts.length > 0 && (
                <label>
                  <input type="radio" checked={method === 'account'} onChange={() => setMethod('account')} /> Heslom k e-mailovému kontu
                </label>
              )}
              {st.hasRecoveryCode && (
                <label>
                  <input type="radio" checked={method === 'code'} onChange={() => setMethod('code')} /> Záchranným kódom
                </label>
              )}
            </div>
            {method === 'account' && st.recoveryAccounts.length > 0 ? (
              <>
                <select value={accId} onChange={(e) => setAccId(e.target.value)} aria-label="Konto">
                  {st.recoveryAccounts.map((a) => (
                    <option key={a.id} value={a.id}>
                      {a.name} – {a.email}
                    </option>
                  ))}
                </select>
                <input ref={input} type="password" autoComplete="off" placeholder="Heslo k tomuto kontu" value={accPw} onChange={(e) => setAccPw(e.target.value)} aria-label="Heslo ku kontu" />
                <p className="lock-note">Overí sa prihlásením na poštový server (treba pripojenie).</p>
              </>
            ) : (
              <input ref={input} type="text" autoComplete="off" spellCheck={false} placeholder="XXXX-XXXX-XXXX-XXXX" value={code} onChange={(e) => setCode(e.target.value.toUpperCase())} aria-label="Záchranný kód" />
            )}
            <div className="lock-row">
              <button type="submit" className="lock-primary" disabled={busy || (method === 'account' ? !accPw : !code)}>
                {busy ? 'Overujem…' : 'Overiť'}
              </button>
              <button type="button" onClick={() => setStep('unlock')}>
                Späť
              </button>
            </div>
          </>
        )}
        {step === 'reset' && (
          <>
            <h1>Nový PIN alebo heslo</h1>
            <div className="lock-tabs" role="radiogroup">
              <label>
                <input type="radio" checked={kind === 'pin'} onChange={() => setKind('pin')} /> PIN (4–12 číslic)
              </label>
              <label>
                <input type="radio" checked={kind === 'password'} onChange={() => setKind('password')} /> Heslo (aspoň 6 znakov)
              </label>
            </div>
            <input ref={input} type="password" autoComplete="new-password" placeholder={kind === 'pin' ? 'Nový PIN' : 'Nové heslo'} value={new1} onChange={(e) => setNew1(kind === 'pin' ? e.target.value.replace(/\D/g, '') : e.target.value)} {...pinProps(kind)} />
            <input type="password" autoComplete="new-password" placeholder="Zopakuj" value={new2} onChange={(e) => setNew2(kind === 'pin' ? e.target.value.replace(/\D/g, '') : e.target.value)} {...pinProps(kind)} />
            <button type="submit" className="lock-primary" disabled={busy || !new1 || !new2}>
              Uložiť a odomknúť
            </button>
          </>
        )}
        {step === 'code' && (
          <>
            <h1>Nový záchranný kód</h1>
            <p className="lock-note">Starý kód už neplatí. Zapíš si tento:</p>
            <div className="lock-code">{newCode}</div>
            <button type="submit" className="lock-primary">
              Mám zapísané
            </button>
          </>
        )}
        {error && (
          <div className="lock-error" role="alert">
            {error}
          </div>
        )}
        {step === 'unlock' && st.blockedFor > 0 && !error && <div className="lock-error">Príliš veľa pokusov – počkaj {st.blockedFor} s.</div>}
      </form>
    </div>
  );
}
