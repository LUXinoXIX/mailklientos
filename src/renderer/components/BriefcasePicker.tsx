// Aktovka Zimbry: výber súborov na priloženie k správe.
import { Button, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Input, Spinner } from '@fluentui/react-components';
import { useEffect, useState } from 'react';
import type { BriefcaseDoc, PickedFile } from '../../core/api';

export function BriefcasePicker(props: { accountId: string; onPick: (f: PickedFile) => void; onClose: () => void }) {
  const api = window.mailApi!;
  const [q, setQ] = useState('');
  const [docs, setDocs] = useState<BriefcaseDoc[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  useEffect(() => {
    let alive = true;
    const t = window.setTimeout(() => {
      void api.zimbraBriefcase(props.accountId, q.trim()).then((r) => {
        if (!alive) return;
        if (r.ok) setDocs(r.data);
        else setError(r.error);
      });
    }, q ? 300 : 0);
    return () => {
      alive = false;
      window.clearTimeout(t);
    };
  }, [q, props.accountId]); // eslint-disable-line react-hooks/exhaustive-deps

  return (
    <Dialog open onOpenChange={(_, d) => !d.open && props.onClose()}>
      <DialogSurface style={{ maxWidth: 640 }}>
        <DialogBody>
          <DialogTitle>Priložiť z Aktovky</DialogTitle>
          <DialogContent className="dialog-content">
            <Input placeholder="Hľadať súbor" value={q} onChange={(_, d) => setQ(d.value)} />
            {error && <div className="pane-error">⚠ {error}</div>}
            {docs === null && !error && <Spinner size="tiny" />}
            <div className="brief-list">
              {docs?.map((d) => (
                <div key={d.id} className="org-row crypto-row">
                  <span className="grow">{d.name}<br /><span className="muted small">{Math.ceil(d.size / 1024)} kB · {new Date(d.modified).toLocaleDateString('sk-SK')}</span></span>
                  <Button size="small" disabled={busy === d.id} onClick={async () => {
                    setBusy(d.id);
                    const r = await api.zimbraDownload(props.accountId, d);
                    setBusy(null);
                    if (r.ok) props.onPick(r.data);
                    else setError(r.error);
                  }}>{busy === d.id ? 'Sťahujem…' : 'Priložiť'}</Button>
                </div>
              ))}
              {docs && !docs.length && <div className="muted small">Aktovka je prázdna.</div>}
            </div>
          </DialogContent>
          <DialogActions><Button onClick={props.onClose}>Hotovo</Button></DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
