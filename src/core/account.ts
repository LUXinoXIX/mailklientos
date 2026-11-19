// Model konta. Čisté typy, bez závislosti na Electrone (použiteľné aj v rendereri).

/** Šifrovanie spojenia. */
export type Security = 'ssl' | 'starttls' | 'none';

/**
 * Ako sa správať k certifikátu servera.
 * - strict: plná kontrola (predvolené)
 * - pinned: dôverovať len certifikátu s uloženým odtlačkom SHA-256 (self-signed, interné servery)
 * - ignore: neoverovať vôbec (len explicitne, s varovaním)
 */
export type CertMode = 'strict' | 'pinned' | 'ignore';

export interface TlsPolicy {
  certMode: CertMode;
  /** Odtlačok SHA-256 v tvare "AB:CD:..." (pri certMode = 'pinned'). */
  pinnedSha256?: string;
  /**
   * Režim pre staré servery (napr. staré verzie Zimbry): povolí TLS 1.0/1.1,
   * slabšie šifry a krátke kľúče / SHA-1 podpisy.
   */
  legacy: boolean;
}

export interface ServerConfig {
  host: string;
  port: number;
  security: Security;
  tls: TlsPolicy;
}

export interface AccountConfig {
  id: string;
  kind: 'imap';
  name: string;
  email: string;
  /** Farba konta v UI (bodka, avatar). */
  color: string;
  username: string;
  imap: ServerConfig;
  smtp: ServerConfig;
  /** ďalšie adresy (aliasy) toho istého konta pre pole Od */
  identities?: Identity[];
  /** adresa CalDAV/CardDAV (kalendár, kontakty); prázdne = zistiť automaticky */
  davUrl?: string;
  /** adresa Zimbry (https://server); prázdne = server IMAP */
  zimbraUrl?: string;
  /** krátky štítok v zbalenom páse kont (prázdne = typ konta: ZIMBRA, IMAP) */
  short?: string;
  /** ikonka v páse kont */
  iconKind?: 'auto' | 'text' | 'preset' | 'image';
  iconText?: string;
  iconPreset?: string;
  iconImage?: string;
}

export interface Identity {
  name: string;
  email: string;
}

/**
 * Konto Microsoft 365 / Outlook.com: vložený Outlook na webe ako celé okno.
 * Aplikácia nedrží heslo ani token; prihlásenie si pamätá webová relácia konta.
 */
export interface M365Account {
  id: string;
  kind: 'm365';
  name: string;
  email: string;
  color: string;
  /** Adresa Outlooku: outlook.office.com (firemné M365) alebo outlook.live.com (osobné);
      'web' = ľubovoľný webmail (Zimbra, Roundcube, SOGo…) vložený ako stránka. */
  host?: 'office' | 'live' | 'web';
  /** adresa webmailu pri host 'web' (napr. https://mail.firma.sk) */
  webUrl?: string;
  /** webmail s neplatným certifikátom (interný server) – dôverovať mu */
  webIgnoreCert?: boolean;
  /** webmail: prihlasovacie meno na vyplnenie (heslo je v trezore systému ako pri IMAP) */
  webUser?: string;
  /** webmail: po vyplnení aj odoslať prihlásenie */
  webAutoLogin?: boolean;
  /** krátky štítok v zbalenom páse kont (prázdne = MS365 / MSO) */
  short?: string;
  /** ikonka v páse kont */
  iconKind?: 'auto' | 'text' | 'preset' | 'image';
  iconText?: string;
  iconPreset?: string;
  iconImage?: string;
}

export type AnyAccount = AccountConfig | M365Account;

/** IMAP konto tak, ako ide do rendereru (bez hesla). */
export type ImapAccountPublic = AccountConfig & { hasPassword: boolean };
/** Ľubovoľné konto pre renderer. */
export type AccountPublic = ImapAccountPublic | (M365Account & { hasPassword: false; webHasPassword?: boolean });

export const isImap = (a: { kind: string }): a is ImapAccountPublic => a.kind === 'imap';
export const isM365 = (a: { kind: string }): a is M365Account & { hasPassword: false } => a.kind === 'm365';
/** vložený webmail (nie Outlook) */
export const isWebmail = (a: { kind: string; host?: string }) => a.kind === 'm365' && a.host === 'web';

/** Webová relácia (cookies, prihlásenie) konta M365. */
export const accountPartition = (id: string) => `persist:acc-${id}`;

export const DEFAULT_PORTS: Record<'imap' | 'smtp', Record<Security, number>> = {
  imap: { ssl: 993, starttls: 143, none: 143 },
  smtp: { ssl: 465, starttls: 587, none: 25 },
};

export const strictPolicy = (): TlsPolicy => ({ certMode: 'strict', legacy: false });
