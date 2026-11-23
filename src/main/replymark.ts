// Po odoslaní odpovede/preposlania označí pôvodnú správu (\\Answered / $Forwarded) a obnoví zoznamy.
import { BrowserWindow } from 'electron';
import type { SendInput } from '../core/api';
import * as store from './accounts';
import { log } from './diagnostics';
import * as mail from './mail';

export async function markOriginAfterSend(msg: SendInput) {
  const o = msg.origin;
  if (!o) return;
  try {
    await mail.markOrigin(store.getImap(o.accountId), o);
    for (const w of BrowserWindow.getAllWindows()) if (!w.isDestroyed()) w.webContents.send('app:event', { type: 'mail-changed', accountId: o.accountId, folder: o.folder, newCount: 0 });
  } catch (e) {
    // chyba označenia nesmie ovplyvniť odoslanie
    log('označenie odpovedanej/preposlanej správy zlyhalo', String(e));
  }
}
