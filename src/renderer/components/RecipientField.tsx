// Pole príjemcov s „čipmi“ a dopĺňaním adries (ako v Outlooku).
import { useEffect, useMemo, useRef, useState } from 'react';
import { CloseSmallIcon } from '../icons';

export interface Contact {
  name: string;
  email: string;
}

const EMAIL = /[^\s<>,;"]+@[^\s<>,;"]+\.[^\s<>,;"]+/;

/** "Meno <a@b.sk>" alebo "a@b.sk" → kontakt. */
export function parseAddress(s: string): Contact | null {
  const t = s.trim().replace(/[;,]$/, '');
  if (!t) return null;
  const m = t.match(/^"?([^"<]*)"?\s*<([^>]+)>$/);
  if (m) return { name: m[1]!.trim(), email: m[2]!.trim() };
  const e = t.match(EMAIL);
  return e ? { name: '', email: e[0] } : null;
}

export const formatAddress = (c: Contact) => (c.name ? `"${c.name.replace(/"/g, '')}" <${c.email}>` : c.email);

/** Rozdelí reťazec adries ("a@b, Meno <c@d>") na kontakty. */
export function splitAddresses(s: string): Contact[] {
  return s
    .split(/[,;](?=(?:[^"]*"[^"]*")*[^"]*$)/)
    .map(parseAddress)
    .filter((x): x is Contact => !!x);
}

export function RecipientField(props: {
  label: string;
  value: Contact[];
  onChange: (v: Contact[]) => void;
  contacts: Contact[];
  autoFocus?: boolean;
  extra?: React.ReactNode;
  /** vyhľadanie v adresári firmy (Zimbra GAL) */
  lookup?: (q: string) => Promise<Contact[]>;
}) {
  const [text, setText] = useState('');
  const [remote, setRemote] = useState<Contact[]>([]);
  useEffect(() => {
    const q = text.trim();
    if (!props.lookup || q.length < 2) return setRemote([]);
    let alive = true;
    const t = window.setTimeout(() => void props.lookup!(q).then((r) => alive && setRemote(r)), 200);
    return () => {
      alive = false;
      window.clearTimeout(t);
    };
  }, [text, props.lookup]);
  const [hi, setHi] = useState(0);
  const [focus, setFocus] = useState(false);
  const input = useRef<HTMLInputElement>(null);

  const suggestions = useMemo(() => {
    const q = text.trim().toLowerCase();
    if (q.length < 1) return [];
    const taken = new Set(props.value.map((v) => v.email.toLowerCase()));
    const n = (x: string) => x.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
    const nq = n(q);
    const local = props.contacts.filter((c) => !taken.has(c.email.toLowerCase()) && n(`${c.name} ${c.email}`).includes(nq));
    const all = [...local, ...remote.filter((c) => !taken.has(c.email.toLowerCase()))];
    return all.filter((c, i) => all.findIndex((x) => x.email.toLowerCase() === c.email.toLowerCase()) === i).slice(0, 10);
  }, [text, props.contacts, props.value, remote]);

  const commit = (c?: Contact | null) => {
    const v = c ?? parseAddress(text);
    if (v && !props.value.some((x) => x.email.toLowerCase() === v.email.toLowerCase())) props.onChange([...props.value, v]);
    if (v || !text.trim()) setText('');
    setHi(0);
  };

  return (
    <div className="rcpt-row">
      <span className="rcpt-label">{props.label}</span>
      <div className="rcpt-box" onClick={() => input.current?.focus()}>
        {props.value.map((c, i) => (
          <span key={`${c.email}-${i}`} className="rcpt-chip" title={c.email}>
            {c.name || c.email}
            <button
              className="icon-x"
              aria-label={`Odobrať ${c.name || c.email}`}
              onClick={(e) => {
                e.stopPropagation();
                props.onChange(props.value.filter((_, j) => j !== i));
              }}
            >
              <CloseSmallIcon />
            </button>
          </span>
        ))}
        <input
          ref={input}
          className="rcpt-input"
          aria-label={props.label}
          autoFocus={props.autoFocus}
          value={text}
          onFocus={() => setFocus(true)}
          onBlur={() => {
            setFocus(false);
            commit();
          }}
          onChange={(e) => {
            const v = e.target.value;
            if (/[,;]$/.test(v)) {
              const c = parseAddress(v);
              if (c) {
                commit(c);
                return;
              }
            }
            setText(v);
            setHi(0);
          }}
          onPaste={(e) => {
            const t = e.clipboardData.getData('text');
            const many = splitAddresses(t);
            if (many.length > 1) {
              e.preventDefault();
              const seen = new Set(props.value.map((v) => v.email.toLowerCase()));
              props.onChange([...props.value, ...many.filter((m) => !seen.has(m.email.toLowerCase()))]);
            }
          }}
          onKeyDown={(e) => {
            if (suggestions.length && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
              e.preventDefault();
              setHi((h) => (h + (e.key === 'ArrowDown' ? 1 : suggestions.length - 1)) % suggestions.length);
            } else if (e.key === 'Enter' || e.key === 'Tab') {
              if (suggestions.length) {
                e.preventDefault();
                commit(suggestions[hi]);
              } else if (text.trim()) {
                if (e.key === 'Enter') e.preventDefault();
                commit();
              }
            } else if (e.key === 'Backspace' && !text && props.value.length) {
              props.onChange(props.value.slice(0, -1));
            }
          }}
        />
        {focus && suggestions.length > 0 && (
          <div className="rcpt-suggest" role="listbox">
            {suggestions.map((c, i) => (
              <button
                key={c.email}
                role="option"
                aria-selected={i === hi}
                className={i === hi ? 'on' : ''}
                onMouseDown={(e) => {
                  e.preventDefault();
                  commit(c);
                }}
              >
                <b>{c.name || c.email}</b>
                {c.name && <span className="muted"> {c.email}</span>}
              </button>
            ))}
          </div>
        )}
      </div>
      {props.extra}
    </div>
  );
}
