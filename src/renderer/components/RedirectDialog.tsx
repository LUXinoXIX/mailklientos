// Poslať ďalej ako originál (presmerovanie): pôvodná správa odíde bez zmeny, ako ju poslal pôvodný odosielateľ.
import { Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Field, Input, MessageBar, MessageBarBody } from '@fluentui/react-components';
import { useState } from 'react';
import type { MessageFull } from '../../core/api';

export function RedirectDialog(props: { msg: MessageFull; onClose: () => void; onDone: (text: string) => void }) {
  const m = props.msg;
  const [to, setTo] = useState('');
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const send = async () => {
    setBusy(true);
    setError('');
    const r = await window.mailApi!.redirectMessage(m.accountId, m.folder, m.uid, to);
    setBusy(false);
    if (!r.ok) return setError(r.error);
    props.onDone(`Správa „${m.subject}“ poslaná ďalej ako originál: ${to}`);
    props.onClose();
  };
  return (
    <Dialog open onOpenChange={(_, d) => !d.open && !busy && props.onClose()}>
      <DialogSurface style={{ maxWidth: 520 }}>
        <DialogBody>
          <DialogTitle>Poslať ďalej ako originál</DialogTitle>
          <DialogContent className="dialog-content">
            <span>
              <b>{m.subject || '(bez predmetu)'}</b> – od {m.fromName || m.fromAddress}
            </span>
            <Field label="Komu" hint="Viac adries oddeľte čiarkou.">
              <Input
                value={to}
                onChange={(_, d) => setTo(d.value)}
                placeholder="meno@firma.sk"
                autoFocus
                onKeyDown={(e) => e.key === 'Enter' && to.includes('@') && !busy && void send()}
              />
            </Field>
            <MessageBar intent="info">
              <MessageBarBody>
                Správa odíde bez zmeny – príjemca ju uvidí od pôvodného odosielateľa, s pôvodnými prílohami a s informáciou, že ju preposielate vy. Niektoré servery takto
                poslané správy môžu označiť ako podozrivé; ak to vadí, použite „Poslať ďalej ako prílohu“.
              </MessageBarBody>
            </MessageBar>
            {error && (
              <MessageBar intent="error">
                <MessageBarBody>{error}</MessageBarBody>
              </MessageBar>
            )}
          </DialogContent>
          <DialogActions>
            <Button onClick={props.onClose} disabled={busy}>
              Zrušiť
            </Button>
            <Button appearance="primary" onClick={() => void send()} disabled={busy || !to.includes('@')}>
              {busy ? 'Posielam…' : 'Poslať'}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
