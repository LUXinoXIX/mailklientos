// Varovanie pred otvorením prílohy (náhľad / otvoriť). Dá sa vypnúť pre konkrétny typ súboru (napr. .pdf),
// okrem súborov s makrami; spustiteľné a maskované súbory sa neotvárajú vôbec (len uložiť).
import { Button, Checkbox, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle } from '@fluentui/react-components';
import { useState, type ReactNode } from 'react';
import { attachmentRisk, SAFE_IMAGE, type AttachmentRisk } from '../../core/attrisk';

const KEY = 'attTrust';
function trusted(): string[] {
  try {
    return JSON.parse(localStorage.getItem(KEY) ?? '[]') as string[];
  } catch {
    return [];
  }
}
function trust(ext: string) {
  try {
    localStorage.setItem(KEY, JSON.stringify([...new Set([...trusted(), ext])]));
  } catch {
    /* úložisko nedostupné – len sa bude pýtať znova */
  }
}
/** zoznam typov, pri ktorých sa už nepýta (na zobrazenie / zrušenie v nastaveniach) */
export function trustedTypes() {
  return trusted();
}
export function resetTrustedTypes() {
  try {
    localStorage.removeItem(KEY);
  } catch {
    /* nič */
  }
}

type Ask = { filename: string; risk: AttachmentRisk; run: () => void; onSave?: () => void };

export function useAttachmentGuard(): { guard: (filename: string, run: () => void, onSave?: () => void) => void; dialog: ReactNode } {
  const [ask, setAsk] = useState<Ask | null>(null);
  const [dontAsk, setDontAsk] = useState(false);
  const guard = (filename: string, run: () => void, onSave?: () => void) => {
    const risk = attachmentRisk(filename);
    // obrázky sa ukazujú priamo v programe – nič sa nespúšťa
    if (risk.level === 'normal' && (SAFE_IMAGE.test(filename) || trusted().includes(risk.ext))) return run();
    setDontAsk(false);
    setAsk({ filename, risk, run, onSave });
  };
  const close = () => setAsk(null);
  const r = ask?.risk;
  const dialog = ask && r && (
    <Dialog open onOpenChange={(_, d) => !d.open && close()}>
      <DialogSurface className={`att-guard ${r.level}`}>
        <DialogBody>
          <DialogTitle>{r.level === 'blocked' ? '⛔ Túto prílohu program neotvorí' : '⚠ Otvoriť prílohu?'}</DialogTitle>
          <DialogContent className="dialog-content">
            <div className="att-guard-file">
              <b>{ask.filename.replace(/[‪-‮⁦-⁩]/g, '�')}</b>
              <span className={`att-guard-ext ${r.level}`}>skutočný typ: .{r.ext || '?'}</span>
            </div>
            {r.reasons.map((x) => (
              <p key={x} className="att-guard-reason">{x}</p>
            ))}
            {r.level === 'blocked' ? (
              <p>Ak súbor naozaj potrebuješ a odosielateľa poznáš, ulož ho a pred otvorením ho nechaj skontrolovať antivírusom.</p>
            ) : (
              <p>
                Prílohy môžu byť nebezpečné – aj súbor, ktorý vyzerá ako Word alebo PDF, môže spustiť škodlivý program. Otvor ju, len ak poznáš
                odosielateľa a prílohu si čakal.
              </p>
            )}
            {r.level === 'normal' && r.ext && (
              <Checkbox checked={dontAsk} onChange={(_, d) => setDontAsk(!!d.checked)} label={`Pri súboroch .${r.ext} sa už nepýtať a otvoriť hneď`} />
            )}
          </DialogContent>
          <DialogActions>
            {r.level !== 'blocked' && (
              <Button
                appearance="primary"
                onClick={() => {
                  if (dontAsk && r.ext) trust(r.ext);
                  close();
                  ask.run();
                }}
              >
                Otvoriť
              </Button>
            )}
            {ask.onSave && (
              <Button
                appearance={r.level === 'blocked' ? 'primary' : 'secondary'}
                onClick={() => {
                  close();
                  ask.onSave?.();
                }}
              >
                Uložiť…
              </Button>
            )}
            <Button onClick={close}>Zrušiť</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
  return { guard, dialog: dialog || null };
}
