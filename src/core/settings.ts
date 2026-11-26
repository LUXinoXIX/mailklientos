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

/** vlastný pravý panel (weby + programy), ktorý môže používať jedno alebo viac kont */
/** Kalendár z odkazu (ICS/webcal) v karte pravého panela – poradie = poradie v zozname. */
export interface IcsFeed {
  id: string;
  name: string;
  url: string;
  /** počet zobrazených najbližších udalostí */
  count: number;
  color?: string;
}

export interface RightPanelDef {
  id: string;
  name: string;
  webs: WebItem[];
  launchers: LauncherItem[];
  /** karta „Kalendáre (ICS)“ v tomto paneli */
  icsFeeds?: IcsFeed[];
  /** obnovovanie kalendárov z odkazu v minútach (predvolene 5) */
  icsRefreshMin?: number;
}

export interface AppSettings {
  /** klient OAuth Google (Google Cloud, typ „Počítačová aplikácia“) pre Prihlásiť sa cez Google */
  google?: { clientId: string; clientSecret?: string };
  /** spoločný pravý panel (pre kontá bez vlastného) */
  webs: WebItem[];
  launchers: LauncherItem[];
  /** vlastné pravé panely */
  rightPanels?: RightPanelDef[];
  /** ktorý panel používa konto (id konta alebo „Všetky“ → id panela; chýba = spoločný) */
  rightPanelOf?: Record<string, string>;
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
  /** po zatvorení okna bežať ďalej pri hodinách (predvolene áno) */
  runInBackground?: boolean;
  /** hustota zoznamu správ */
  density?: 'comfortable' | 'compact';
  /** riadok s náhľadom textu v zozname */
  previewLine?: boolean;
  /** dvojklik na správu: samostatné okno (predvolené) alebo karta v hlavnom okne */
  openMode?: 'window' | 'tab';
  /** kde hľadať */
  searchField?: SearchField;
  /** po hľadaní v uložených správach hneď hľadať aj na serveri (predvolene áno) */
  searchServerAuto?: boolean;
  /** nedávne hľadania (najnovšie prvé) */
  searchHistory?: string[];
  /** zoznam kont s menami natrvalo rozbalený (inak úzky pás a vysúvanie po kliknutí) */
  accountListPinned?: boolean;
  /** náhľad kalendára / úloh vedľa čítacieho panela v pošte */
  myDay?: { calendar: boolean; tasks: boolean };
  /** online schôdze v udalostiach (Teams cez zvolené Outlook konto, Jitsi server) */
  meetings?: import('./meeting').MeetingSettings;
  /** karta „Kalendáre (ICS)“ v spoločnom pravom paneli */
  icsFeeds?: IcsFeed[];
  /** obnovovanie kalendárov z odkazu (ICS) v minútach, predvolene 5 */
  icsRefreshMin?: number;
  /** úvodná stránka Súhrn */
  summary?: SummarySettings;
  /** ikona „Všetky doručené“ v páse kont (predvolene zapnutá) */
  showAllInboxes?: boolean;
  /** zoradenie zoznamu správ */
  listSort?: { by: 'date' | 'from' | 'subject' | 'flag' | 'unread' | 'category'; dir: 'desc' | 'asc' };
  /** šírky stĺpcov v px (priečinky, zoznam správ, pravý panel) */
  paneWidths?: { folders?: number; list?: number; right?: number };
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
  /** kde hľadať (predvolene všade) */
  field?: SearchField;
}

/** Kde hľadať: ako v eM Cliente / Thunderbirde. */
export type SearchField = 'all' | 'people' | 'subject' | 'body' | 'from' | 'to';
export const SEARCH_FIELDS: Array<{ id: SearchField; label: string }> = [
  { id: 'all', label: 'Predmet, odosielateľ, príjemcovia a text' },
  { id: 'people', label: 'Predmet, odosielateľ a príjemcovia' },
  { id: 'subject', label: 'Predmet' },
  { id: 'body', label: 'Text správy' },
  { id: 'from', label: 'Odosielateľ' },
  { id: 'to', label: 'Príjemcovia' },
];

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
/** je adresa lokálny súbor (file://) */
export const isFileUrl = (url: string) => /^file:/i.test(url);
/** krátky popis adresy na zobrazenie: pri súbore jeho názov, inak adresa bez https:// */
export function urlLabel(url: string): string {
  if (!isFileUrl(url)) return url.replace(/^https?:\/\//, '');
  try {
    return decodeURIComponent(new URL(url).pathname.split('/').pop() || url);
  } catch {
    return url;
  }
}

export function normalizeUrl(input: string): string | null {
  let t = input.trim().replace(/^"(.*)"$/, '$1');
  if (!t) return null;
  // lokálny súbor: cesta Windows (C:\…\x.html, \\server\…) alebo macOS/Linux (/Users/…/x.html) → file:// adresa
  if (/^[a-z]:[\\/]/i.test(t)) t = 'file:///' + encodeURI(t.replace(/\\/g, '/')).replace(/#/g, '%23').replace(/\?/g, '%3F');
  else if (/^\\\\[^\\]/.test(t)) t = 'file:' + encodeURI(t.replace(/\\/g, '/')).replace(/#/g, '%23').replace(/\?/g, '%3F');
  else if (t.startsWith('/')) t = 'file://' + encodeURI(t).replace(/#/g, '%23').replace(/\?/g, '%3F');
  if (/^file:/i.test(t)) {
    try {
      const u = new URL(t);
      return /\.(html?|xhtml|svg|pdf|txt)$/i.test(decodeURIComponent(u.pathname)) ? u.toString() : null;
    } catch {
      return null;
    }
  }
  const withScheme = /^[a-z][a-z0-9+.-]*:/i.test(t) ? t : `https://${t}`;
  try {
    const u = new URL(withScheme);
    return u.protocol === 'https:' || u.protocol === 'http:' ? u.toString() : null;
  } catch {
    return null;
  }
}

export type SummaryTile = 'accounts' | 'today' | 'tasks' | 'flagged' | 'notice' | 'feeds' | 'news' | 'weather';
export interface SummarySettings {
  /** false = Súhrn úplne vypnutý (nie je v páse kont ani po spustení) */
  enabled?: boolean;
  /** čo otvoriť po spustení (predvolene Súhrn) */
  startup?: 'summary' | 'all' | 'last';
  /** poradie a viditeľnosť dlaždíc */
  tiles?: { id: SummaryTile; on: boolean; /** vlastný názov dlaždice */ title?: string }[];
  /** vstavaná tapeta (id) alebo vlastný obrázok (data URL, zmenšený) */
  wallpaper?: string;
  customImage?: string;
  /** vlastný obrázok je tmavý (zistené pri nahratí) */
  customDark?: boolean;
  /** kalendáre z odkazu (iCal/webcal) – najbližšie udalosti */
  feeds?: { id: string; name: string; url: string; count: number }[];
  /** správy z RSS/Atom kanálov (dlaždica Správy) */
  news?: { id: string; name: string; url: string; count: number }[];
  /** počasie – miesto (dlaždica Počasie) */
  weather?: { name: string; lat: number; lon: number; detail?: string };
  /** kalendáre z odkazu na celú šírku (stĺpec pre každý), pod ostatnými dlaždicami */
  feedsWide?: boolean;
  /** nepriehľadnosť blokov 0.3–1 */
  opacity?: number;
  /** rozmazanie pozadia pod blokmi */
  blur?: boolean;
  /** tmavé bloky (undefined = podľa tapety) */
  dark?: boolean;
}
