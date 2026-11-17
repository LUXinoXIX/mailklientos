import {
  Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input,
  MessageBar, MessageBarBody, Select, Spinner, Textarea,
} from '@fluentui/react-components';
import { useState } from 'react';
import type { ImapAccountPublic } from '../../core/account';
import type { PickedFile, SendInput } from '../../core/api';
import { AttachIcon, CloseSmallIcon } from '../icons';

export interface ComposeInit {
  accountId?: string;
  to?: string;
  cc?: string;
  subject?: string;
  text?: string;
  replyMessageId?: string;
  title?: string;
}

export function ComposeDialog(props: {
  open: boolean;
  init: ComposeInit;
  accounts: ImapAccountPublic[];
  onClose: () => void;
  onSent: (info: string) => void;
}) {
  const [accountId, setAccountId] = useState(props.init.accountId ?? props.accounts[0]?.id ?? '');
  const [to, setTo] = useState(props.init.to ?? '');
  const [cc, setCc] = useState(props.init.cc ?? '');
  const [files, setFiles] = useState<PickedFile[]>([]);
  const [subject, setSubject] = useState(props.init.subject ?? '');
  const [text, setText] = useState(props.init.text ?? '');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  async function send() {
    if (!window.mailApi || !accountId) return;
    if (!to.includes('@')) return setError('Zadaj príjemcu.');
    setBusy(true);
    setError(null);
    const msg: SendInput = {
      to,
      cc,
      subject,
      text,
      replyMessageId: props.init.replyMessageId,
      attachments: files.map((f) => ({ path: f.path, filename: f.filename })),
    };
    const r = await window.mailApi.send(accountId, msg);
    setBusy(false);
    if (r.ok) props.onSent(r.data.savedToSent ? 'Správa odoslaná a uložená do Odoslaných.' : 'Správa odoslaná.');
    else setError(`${r.error}${r.hint ? ` ${r.hint}` : ''}`);
  }

  return (
    <Dialog open={props.open} onOpenChange={(_, d) => !d.open && props.onClose()}>
      <DialogSurface className="compose-dialog">
        <DialogBody>
          <DialogTitle>{props.init.title ?? 'Nová správa'}</DialogTitle>
          <DialogContent className="dialog-content">
            {props.accounts.length === 0 ? (
              <MessageBar intent="warning"><MessageBarBody>Najprv pridaj konto (Domov → Pridať konto).</MessageBarBody></MessageBar>
            ) : (
              <>
                <Field label="Od">
                  <Select value={accountId} onChange={(_, d) => setAccountId(d.value)}>
                    {props.accounts.map((a) => (
                      <option key={a.id} value={a.id}>{a.name} &lt;{a.email}&gt;</option>
                    ))}
                  </Select>
                </Field>
                <Field label="Komu"><Input value={to} onChange={(_, d) => setTo(d.value)} /></Field>
                <Field label="Kópia"><Input value={cc} onChange={(_, d) => setCc(d.value)} /></Field>
                <Field label="Predmet"><Input value={subject} onChange={(_, d) => setSubject(d.value)} /></Field>
                <Textarea className="compose-text" resize="vertical" value={text} onChange={(_, d) => setText(d.value)} />
                <div className="row gap">
                  <Button
                    size="small"
                    icon={<AttachIcon />}
                    onClick={async () => {
                      const r = await window.mailApi?.pickFiles();
                      if (r?.ok) setFiles((f) => [...f, ...r.data]);
                    }}
                  >
                    Priložiť súbor
                  </Button>
                  {files.map((f, i) => (
                    <span key={`${f.path}-${i}`} className="attachment">
                      {f.filename} <span className="muted">({Math.ceil(f.size / 1024)} kB)</span>
                      <button className="icon-x" aria-label={`Odobrať ${f.filename}`} onClick={() => setFiles((x) => x.filter((_, j) => j !== i))}>
                        <CloseSmallIcon />
                      </button>
                    </span>
                  ))}
                </div>
              </>
            )}
            {error && <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar>}
          </DialogContent>
          <DialogActions>
            {busy && <Spinner size="tiny" label="Odosielam…" />}
            <Button appearance="primary" disabled={busy || props.accounts.length === 0} onClick={send}>Odoslať</Button>
            <Button disabled={busy} onClick={props.onClose}>Zrušiť</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
