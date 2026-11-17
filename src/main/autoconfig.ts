// Automatické nastavenie servera podľa domény e-mailu (formát Mozilla Autoconfig / ISPDB z Thunderbirdu).
// Poradie ako v Thunderbirde: autoconfig.<doména> → <doména>/.well-known → ISPDB (autoconfig.thunderbird.net) → MX doména v ISPDB.
import dns from 'node:dns/promises';
import { net } from 'electron';
import type { Security } from '../core/account';
import type { AutoConfigResult, AutoServer } from '../core/api';

async function fetchText(url: string, ms = 6000): Promise<string | null> {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), ms);
  try {
    const r = await net.fetch(url, { signal: ctrl.signal, redirect: 'follow' });
    if (!r.ok) return null;
    const text = await r.text();
    return text.includes('<clientConfig') ? text : null;
  } catch {
    return null;
  } finally {
    clearTimeout(t);
  }
}

const tag = (xml: string, name: string) => xml.match(new RegExp(`<${name}>\\s*([^<]*?)\\s*</${name}>`, 'i'))?.[1];

function parseServer(block: string): AutoServer | undefined {
  const host = tag(block, 'hostname');
  const port = Number(tag(block, 'port'));
  if (!host || !port) return undefined;
  const st = (tag(block, 'socketType') ?? '').toUpperCase();
  const security: Security = st === 'SSL' ? 'ssl' : st === 'STARTTLS' ? 'starttls' : 'none';
  const u = tag(block, 'username') ?? '';
  return { host, port, security, username: u.includes('%EMAILADDRESS%') ? 'email' : u.includes('%EMAILLOCALPART%') ? 'local' : 'other' };
}

/** Vyberie najlepší server daného typu (preferuje SSL, potom STARTTLS). */
function pick(xml: string, kind: 'incomingServer' | 'outgoingServer', type: string): AutoServer | undefined {
  const re = new RegExp(`<${kind}[^>]*type="${type}"[^>]*>([\\s\\S]*?)</${kind}>`, 'gi');
  const all = [...xml.matchAll(re)].map((m) => parseServer(m[1]!)).filter((x): x is AutoServer => !!x);
  const rank = (s: AutoServer) => (s.security === 'ssl' ? 0 : s.security === 'starttls' ? 1 : 2);
  return all.sort((a, b) => rank(a) - rank(b))[0];
}

function parse(xml: string, source: string, domain: string): AutoConfigResult | null {
  const sub = (s?: AutoServer) => (s ? { ...s, host: s.host.replace(/%EMAILDOMAIN%/gi, domain) } : undefined);
  const imap = sub(pick(xml, 'incomingServer', 'imap'));
  const smtp = sub(pick(xml, 'outgoingServer', 'smtp'));
  if (!imap && !smtp) return null;
  return { source, name: tag(xml, 'displayName'), imap, smtp };
}

export async function autoconfig(email: string): Promise<AutoConfigResult | null> {
  const domain = email.split('@')[1]?.trim().toLowerCase();
  if (!domain || !/^[a-z0-9.-]+\.[a-z]{2,}$/.test(domain)) return null;
  const q = encodeURIComponent(email);
  const tries: Array<[string, string]> = [
    [`https://autoconfig.${domain}/mail/config-v1.1.xml?emailaddress=${q}`, `autoconfig.${domain}`],
    [`https://${domain}/.well-known/autoconfig/mail/config-v1.1.xml?emailaddress=${q}`, domain],
    [`https://autoconfig.thunderbird.net/v1.1/${domain}`, 'databáza Thunderbirdu (ISPDB)'],
  ];
  for (const [url, source] of tries) {
    const xml = await fetchText(url);
    const r = xml && parse(xml, source, domain);
    if (r) return r;
  }
  // Doména hostovaná u iného poskytovateľa: skúsiť ISPDB podľa MX záznamu (napr. Google Workspace).
  try {
    const mx = (await dns.resolveMx(domain)).sort((a, b) => a.priority - b.priority)[0]?.exchange;
    const parts = mx?.toLowerCase().split('.') ?? [];
    for (let i = Math.max(0, parts.length - 3); i < parts.length - 1; i++) {
      const base = parts.slice(i).join('.');
      const xml = await fetchText(`https://autoconfig.thunderbird.net/v1.1/${base}`);
      const r = xml && parse(xml, `databáza Thunderbirdu (MX ${base})`, domain);
      if (r) return r;
    }
  } catch {
    /* bez MX */
  }
  return null;
}
