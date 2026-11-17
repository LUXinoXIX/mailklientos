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
  /** id podpisu na serveri Zimbra (synchronizácia) */
  zimbraId?: string;
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
  /** kategórie = IMAP kľúčové slová ($label1–5 sú zhodné s Thunderbirdom) */
  categories?: Category[];
  rules?: Rule[];
  quickSteps?: QuickStep[];
  savedSearches?: SavedSearch[];
  /** zoznam správ zoskupený do konverzácií */
  conversations?: boolean;
  templates?: Template[];
  vacation?: Vacation;
  tasks?: Task[];
  customThemes?: ThemeDef[];
}

export interface ThemeDef {
  id: string;
  name: string;
  dark: boolean;
  /** akcent: tlačidlá, výber, odkazy */
  accent: string;
  /** horný pruh; null = farba vybraného konta */
  top: string | null;
  rail: string;
  pane: string;
  read: string;
  text: string;
  /** zaoblenie rohov v px */
  radius: number;
}

export interface Task {
  id: string;
  title: string;
  /** termín (YYYY-MM-DD) */
  due?: string;
  done: boolean;
  note?: string;
  created: string;
  /** úloha vytvorená zo správy */
  message?: { accountId: string; folder: string; uid: number; messageId?: string; subject: string; from: string };
  /** úloha na serveri (CalDAV VTODO, napr. Úlohy v Zimbre) */
  remote?: { accountId: string; uid: string; href?: string; etag?: string };
  /** zmenené lokálne a ešte neodoslané na server */
  dirty?: boolean;
}

export interface Template {
  id: string;
  name: string;
  subject?: string;
  html: string;
}

export interface Vacation {
  enabled: boolean;
  /** prázdne = všetky IMAP kontá */
  accountIds?: string[];
  /** ISO dátum od / do (vrátane) */
  from?: string;
  until?: string;
  subject: string;
  html: string;
  text: string;
}

export interface Category {
  /** IMAP kľúčové slovo */
  id: string;
  name: string;
  color: string;
}

export type RuleField = 'from' | 'to' | 'subject' | 'any';
export interface RuleCond {
  field: RuleField;
  contains: string;
}
export type ActionType = 'move' | 'read' | 'unread' | 'flag' | 'category' | 'pin' | 'archive' | 'delete';
export interface RuleAction {
  type: ActionType;
  /** cieľový priečinok (move) */
  folder?: string;
  /** kľúčové slovo kategórie (category) */
  category?: string;
}
export interface Rule {
  id: string;
  name: string;
  enabled: boolean;
  /** prázdne = všetky IMAP kontá */
  accountId?: string;
  match: 'all' | 'any';
  conds: RuleCond[];
  actions: RuleAction[];
  /** po splnení ďalšie pravidlá nespúšťať */
  stop?: boolean;
}
export interface QuickStep {
  id: string;
  name: string;
  actions: RuleAction[];
}
export interface SavedSearch {
  id: string;
  name: string;
  query: string;
  accountId?: string;
}

export const PIN_KEYWORD = '$Pinned';

export const DEFAULT_CATEGORIES: Category[] = [
  { id: '$label1', name: 'Dôležité', color: '#c50f1f' },
  { id: '$label2', name: 'Práca', color: '#d83b01' },
  { id: '$label3', name: 'Osobné', color: '#107c10' },
  { id: '$label4', name: 'Úlohy', color: '#0f6cbd' },
  { id: '$label5', name: 'Neskôr', color: '#6d28d9' },
];

export const categoriesOf = (s: AppSettings) => s.categories ?? DEFAULT_CATEGORIES;

/** Platí podmienka pravidla pre správu? (texty bez ohľadu na veľkosť písmen a diakritiku) */
export function ruleMatches(rule: Rule, m: { from: string; to: string; subject: string }): boolean {
  const norm = (x: string) => x.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  const conds = rule.conds.filter((c) => c.contains.trim());
  if (!conds.length) return false;
  const test = (c: RuleCond) => {
    const hay = c.field === 'from' ? m.from : c.field === 'to' ? m.to : c.field === 'subject' ? m.subject : `${m.from} ${m.to} ${m.subject}`;
    return norm(hay).includes(norm(c.contains.trim()));
  };
  return rule.match === 'any' ? conds.some(test) : conds.every(test);
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
  // príklad vlastnej témy, ktorý si používateľ vybral v návrhu tém
  customThemes: [
    { id: 'custom-moja', name: 'Moja téma', dark: false, accent: '#2F6F5E', top: '#2F6F5E', rail: '#e3e6ea', pane: '#f5f6f7', read: '#ffffff', text: '#242424', radius: 6 },
  ],
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
