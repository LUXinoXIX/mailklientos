// Nastavenia pravého panela: vložené weby a spúšťače programov. Čisté typy a predvoľby.

export interface WebItem {
  id: string;
  name: string;
  url: string;
}

export interface LauncherItem {
  id: string;
  name: string;
  /** Cesta k programu (.exe, .app, skript) alebo URL schéma (ms-word:, teams:, https://…). */
  target: string;
  args?: string[];
}

export interface Signature {
  id: string;
  name: string;
  /** čistý text (záloha, staršie podpisy) */
  text: string;
  /** formátovaný podpis (HTML, obrázky ako data: URL) */
  html?: string;
}

const escHtml = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Podpis ako HTML: formátovaný, alebo text s riadkami ako <br>. */
export function signatureToHtml(sig: Signature | undefined): string {
  if (!sig) return '';
  if (sig.html && sig.html.replace(/<p><\/p>/g, '').trim()) return sig.html;
  return sig.text ? `<p>${escHtml(sig.text).replace(/\n/g, '<br>')}</p>` : '';
}

/** Predvolené podpisy konta: pre nové správy a pre odpovede/preposlania. */
export interface SignatureDefaults {
  newMsg?: string;
  reply?: string;
}

export interface AppSettings {
  webs: WebItem[];
  launchers: LauncherItem[];
  signatures: Signature[];
  signatureDefaults: Record<string, SignatureDefaults>;
  /** upozornenia Windows na novú poštu (predvolene zapnuté) */
  notifications?: boolean;
}

/** Oddeľovač podpisu (štandard „-- “ s medzerou). */
export const SIG_SEPARATOR = '\n-- \n';

/** Text podpisu pre konto a typ správy (alebo ''). */
export function signatureText(s: AppSettings, accountId: string | undefined, kind: 'new' | 'reply'): string {
  if (!accountId) return '';
  const d = s.signatureDefaults[accountId];
  const id = kind === 'new' ? d?.newMsg : d?.reply;
  return s.signatures.find((x) => x.id === id)?.text ?? '';
}

export const DEFAULT_SETTINGS: AppSettings = {
  webs: [
    { id: 'copilot', name: 'Copilot', url: 'https://copilot.microsoft.com' },
    { id: 'whatsapp', name: 'WhatsApp', url: 'https://web.whatsapp.com' },
    { id: 'teams', name: 'Teams', url: 'https://teams.microsoft.com' },
  ],
  launchers: [],
  signatures: [],
  signatureDefaults: {},
};

/** Webová relácia (cookies, prihlásenie) webu v pravom paneli. */
export const webPartition = (id: string) => `persist:web-${id}`;

/** Doplní https:// a overí, že ide o webovú adresu. */
export function normalizeUrl(input: string): string | null {
  const t = input.trim();
  if (!t) return null;
  const withScheme = /^[a-z][a-z0-9+.-]*:/i.test(t) ? t : `https://${t}`;
  try {
    const u = new URL(withScheme);
    return u.protocol === 'https:' || u.protocol === 'http:' ? u.toString() : null;
  } catch {
    return null;
  }
}
