// Online schôdze v udalostiach kalendára: Teams (firemné / osobné), Jitsi Meet, prípadne iné odkazy.
// Odkaz sa ukladá do udalosti (CalEvent.onlineUrl → iCalendar URL + X-MICROSOFT-SKYPETEAMSMEETINGURL pri Teams)
// a zároveň do miesta a popisu, aby ho videli aj programy, ktoré tieto vlastnosti nepoznajú (staršia Zimbra, telefóny).
import type { CalEvent } from './ical';

export type MeetingKind = 'teams' | 'teams-personal' | 'jitsi' | 'google' | 'zoom' | 'other';

export interface MeetingSettings {
  /** predvolená voľba pre nové udalosti */
  default?: 'none' | 'jitsi' | 'teams';
  /** adresa Jitsi servera, predvolene https://meet.jit.si */
  jitsiServer?: string;
  /** Outlook konto, cez ktoré sa vytvárajú odkazy na Teams (aj keď pozvánku posiela Zimbra) */
  teamsAccountId?: string;
}

export const DEFAULT_JITSI = 'https://meet.jit.si';

const PATTERNS: Array<[MeetingKind, RegExp]> = [
  ['teams', /https:\/\/teams\.microsoft\.com\/(?:l\/meetup-join|meet)\/[^\s<>"')\]]+/i],
  ['teams-personal', /https:\/\/teams\.live\.com\/meet\/[^\s<>"')\]]+/i],
  ['google', /https:\/\/meet\.google\.com\/[a-z]{3}-[a-z]{4}-[a-z]{3}\S*/i],
  ['zoom', /https:\/\/[\w.-]*zoom\.us\/(?:j|my|w)\/[^\s<>"')\]]+/i],
  ['jitsi', /https:\/\/(?:meet\.jit\.si|8x8\.vc|[\w.-]*jitsi[\w.-]*)\/[^\s<>"')\]]+/i],
];

export const MEETING_LABEL: Record<MeetingKind, string> = {
  teams: 'Microsoft Teams',
  'teams-personal': 'Microsoft Teams',
  jitsi: 'Jitsi Meet',
  google: 'Google Meet',
  zoom: 'Zoom',
  other: 'Online schôdza',
};

/** Druh schôdze podľa adresy (vlastný Jitsi server sa spozná podľa nastavenia). */
export function meetingKindOf(url: string, jitsiServer?: string): MeetingKind {
  for (const [k, re] of PATTERNS) if (re.test(url)) return k;
  if (jitsiServer) {
    try {
      if (new URL(url).host === new URL(jitsiServer).host) return 'jitsi';
    } catch {
      /* nič */
    }
  }
  return 'other';
}

/** Nájde odkaz na online schôdzu v udalosti (vlastnosť, miesto, popis). */
export function findMeeting(e: Pick<CalEvent, 'onlineUrl' | 'location' | 'description'>, jitsiServer?: string): { kind: MeetingKind; url: string } | null {
  if (e.onlineUrl) return { kind: meetingKindOf(e.onlineUrl, jitsiServer), url: e.onlineUrl };
  const text = `${e.location ?? ''}\n${e.description ?? ''}`;
  for (const [k, re] of PATTERNS) {
    const m = re.exec(text);
    if (m) return { kind: k, url: m[0].replace(/[.,;]+$/, '') };
  }
  if (jitsiServer) {
    const host = (() => {
      try {
        return new URL(jitsiServer).host;
      } catch {
        return '';
      }
    })();
    const m = host ? new RegExp(`https://${host.replace(/\./g, '\\.')}/[^\\s<>"')\\]]+`, 'i').exec(text) : null;
    if (m) return { kind: 'jitsi', url: m[0] };
  }
  return null;
}

/** Je text odkaz na Teams schôdzu? (zachytávanie zo schránky) */
export function teamsLinkIn(text: string): string | null {
  const m = PATTERNS[0]![1].exec(text) ?? PATTERNS[1]![1].exec(text);
  return m ? m[0] : null;
}

const SLUG_CHARS = 'abcdefghjkmnpqrstuvwxyz23456789';

/** Nová Jitsi miestnosť: čitateľný názov z predmetu + náhodná časť (aby sa nedala uhádnuť). */
export function newJitsiUrl(server: string | undefined, title: string): string {
  const base = (server || DEFAULT_JITSI).replace(/\/+$/, '');
  const words = title
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Za-z0-9]+/g, ' ')
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 4)
    .map((w) => w[0]!.toUpperCase() + w.slice(1).toLowerCase());
  const bytes = new Uint8Array(10);
  globalThis.crypto.getRandomValues(bytes);
  const rnd = [...bytes].map((b) => SLUG_CHARS[b % SLUG_CHARS.length]).join('');
  return `${base}/${words.join('') || 'Schodza'}-${rnd}`;
}

const BLOCK_START = '────────── Online schôdza ──────────';

/** Pripojí k udalosti odkaz: vlastnosť, miesto (ak je prázdne) a blok v popise; starý blok nahradí. */
export function applyMeeting(e: CalEvent, url: string | null, jitsiServer?: string): CalEvent {
  const desc = stripMeetingBlock(e.description ?? '');
  const old = e.onlineUrl;
  const oldLabel = old ? MEETING_LABEL[meetingKindOf(old, jitsiServer)] : '';
  const loc = e.location && old && (e.location === oldLabel || e.location.includes(old)) ? undefined : e.location;
  if (!url) return { ...e, onlineUrl: undefined, location: loc, description: desc || undefined };
  const kind = meetingKindOf(url, jitsiServer);
  const label = MEETING_LABEL[kind];
  const block = [BLOCK_START, `${label} – pripojiť sa:`, url, kind === 'jitsi' ? 'Funguje v prehliadači, bez inštalácie.' : ''].filter(Boolean).join('\n');
  // pozvánka z Outlooku už odkaz v texte má → nepridávať druhýkrát
  if (desc.includes(url)) return { ...e, onlineUrl: url, location: loc || label, description: desc };
  return { ...e, onlineUrl: url, location: loc || label, description: (desc ? `${desc}\n\n` : '') + block };
}

export function stripMeetingBlock(desc: string): string {
  const i = desc.indexOf(BLOCK_START);
  return (i >= 0 ? desc.slice(0, i) : desc).trimEnd();
}
