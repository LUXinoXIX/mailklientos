// Prístupové tokeny OAuth pre IMAP/SMTP/DAV podľa poskytovateľa konta (Google alebo Microsoft).
import { isOutlookCom, type AccountConfig } from '../core/account';
import * as google from './google';
import * as microsoft from './microsoft';

export const accessToken = (acc: AccountConfig) => (isOutlookCom(acc) ? microsoft.accessToken(acc) : google.accessToken(acc));

export function invalidate(acc: AccountConfig) {
  if (isOutlookCom(acc)) microsoft.invalidate(acc.id);
  else google.invalidate(acc.id);
}

export const signOut = (acc: AccountConfig) => (isOutlookCom(acc) ? microsoft.signOut(acc.id) : google.signOut(acc.id));
