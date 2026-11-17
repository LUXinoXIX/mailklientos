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

export interface AppSettings {
  webs: WebItem[];
  launchers: LauncherItem[];
}

export const DEFAULT_SETTINGS: AppSettings = {
  webs: [
    { id: 'copilot', name: 'Copilot', url: 'https://copilot.microsoft.com' },
    { id: 'whatsapp', name: 'WhatsApp', url: 'https://web.whatsapp.com' },
    { id: 'teams', name: 'Teams', url: 'https://teams.microsoft.com' },
  ],
  launchers: [],
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
