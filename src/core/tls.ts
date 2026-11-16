// TLS jadro: mapovanie nastavení konta na voľby Node `tls`, režim pre staré servery
// (staré Zimbry: TLS 1.0/1.1, slabé šifry, 1024-bit kľúče, SHA-1), sonda certifikátu
// (implicitné SSL aj STARTTLS pre IMAP a SMTP) a pinning odtlačku.
//
// Beží v hlavnom procese (Node). Electron používa BoringSSL namiesto OpenSSL,
// preto sa reťazec šifier vyberá skúšobne (legacyCiphers), nie natvrdo.

import net from 'node:net';
import tls from 'node:tls';
import type { Security, ServerConfig, TlsPolicy } from './account';

import type { CertInfo, ProbeResult, Protocol } from './tls-types';
export type { CertInfo, ProbeResult, Protocol } from './tls-types';

// ---------------------------------------------------------------------------
// Šifry pre staré servery

const LEGACY_CIPHER_CANDIDATES = [
  'DEFAULT:@SECLEVEL=0', // OpenSSL 3 (čistý Node)
  'ALL:@SECLEVEL=0',
  'ALL', // BoringSSL (Electron) nepozná @SECLEVEL
  'HIGH:MEDIUM:!aNULL',
];
let legacyCipherCache: string | null | undefined;

/** Vráti reťazec šifier pre staré servery, ktorý akceptuje aktuálna TLS knižnica. */
export function legacyCiphers(): string | undefined {
  if (legacyCipherCache !== undefined) return legacyCipherCache ?? undefined;
  for (const c of LEGACY_CIPHER_CANDIDATES) {
    try {
      tls.createSecureContext({ ciphers: c, minVersion: 'TLSv1' });
      legacyCipherCache = c;
      return c;
    } catch {
      /* skúsiť ďalší */
    }
  }
  legacyCipherCache = null;
  return undefined;
}

// ---------------------------------------------------------------------------
// Mapovanie politiky na voľby tls.connect (použije sa aj pre imapflow a nodemailer)

export function buildTlsOptions(host: string, policy: TlsPolicy): tls.ConnectionOptions {
  const o: tls.ConnectionOptions = {
    minVersion: policy.legacy ? 'TLSv1' : 'TLSv1.2',
  };
  if (!net.isIP(host)) o.servername = host;
  if (policy.legacy) {
    const c = legacyCiphers();
    if (c) o.ciphers = c;
  }
  // pinned: overenie reťazca vypneme a odtlačok kontrolujeme sami (sonda + verifyPinnedSocket).
  if (policy.certMode !== 'strict') o.rejectUnauthorized = false;
  return o;
}

/** Voľby pre imapflow (ImapFlow constructor). */
export function imapflowConnection(cfg: ServerConfig) {
  return {
    host: cfg.host,
    port: cfg.port,
    secure: cfg.security === 'ssl',
    // true = STARTTLS povinné, false = nikdy
    doSTARTTLS: cfg.security === 'starttls' ? true : cfg.security === 'none' ? false : undefined,
    tls: buildTlsOptions(cfg.host, cfg.tls),
  };
}

/** Voľby pre nodemailer.createTransport. */
export function nodemailerConnection(cfg: ServerConfig) {
  return {
    host: cfg.host,
    port: cfg.port,
    secure: cfg.security === 'ssl',
    requireTLS: cfg.security === 'starttls',
    ignoreTLS: cfg.security === 'none',
    tls: buildTlsOptions(cfg.host, cfg.tls),
  };
}

// ---------------------------------------------------------------------------
// Odtlačky a dôvera

export const normFp = (fp: string) => fp.replace(/[^0-9a-f]/gi, '').toUpperCase();

export function certInfo(c: tls.PeerCertificate): CertInfo {
  const name = (x: Record<string, unknown> | undefined) =>
    x ? Object.entries(x).map(([k, v]) => `${k}=${String(v)}`).join(', ') : '';
  return {
    subject: name(c.subject as unknown as Record<string, unknown>),
    issuer: name(c.issuer as unknown as Record<string, unknown>),
    validFrom: c.valid_from,
    validTo: c.valid_to,
    fingerprint256: c.fingerprint256,
    selfSigned: name(c.subject as unknown as Record<string, unknown>) === name(c.issuer as unknown as Record<string, unknown>),
    altNames: c.subjectaltname,
  };
}

/** Rozhodne, či je spojenie dôveryhodné podľa politiky konta. Vráti text chyby alebo null. */
export function trustError(
  policy: TlsPolicy,
  r: { authorized: boolean; hostnameMismatch: boolean; fingerprint256?: string },
): string | null {
  if (policy.certMode === 'ignore') return null;
  if (policy.certMode === 'pinned') {
    if (!policy.pinnedSha256) return 'Chýba uložený odtlačok certifikátu.';
    if (!r.fingerprint256 || normFp(r.fingerprint256) !== normFp(policy.pinnedSha256)) {
      return 'Certifikát servera sa zmenil (odtlačok nesedí s uloženou výnimkou).';
    }
    return null;
  }
  if (!r.authorized) return 'Certifikát servera nie je dôveryhodný.';
  if (r.hostnameMismatch) return 'Certifikát nepatrí k zadanému názvu servera.';
  return null;
}

/** Kontrola už otvoreného TLS socketu (napr. socket imapflow), ak je dostupný. */
export function verifySocket(sock: unknown, host: string, policy: TlsPolicy): string | null {
  if (!(sock instanceof tls.TLSSocket)) return null; // nedá sa overiť, spoliehame sa na sondu
  const cert = sock.getPeerCertificate();
  const mismatch = !!cert && !!tls.checkServerIdentity(host, cert);
  return trustError(policy, {
    authorized: sock.authorized,
    hostnameMismatch: mismatch,
    fingerprint256: cert?.fingerprint256,
  });
}

// ---------------------------------------------------------------------------
// Vysvetlenie chýb po slovensky

const VERIFY_TEXT: Record<string, string> = {
  DEPTH_ZERO_SELF_SIGNED_CERT: 'Certifikát je self-signed (podpísaný sám sebou).',
  SELF_SIGNED_CERT_IN_CHAIN: 'V reťazci je self-signed (interná) certifikačná autorita.',
  UNABLE_TO_VERIFY_LEAF_SIGNATURE: 'Certifikát vydala neznáma autorita alebo chýba medzicertifikát.',
  UNABLE_TO_GET_ISSUER_CERT_LOCALLY: 'Certifikát vydala neznáma (interná) autorita.',
  CERT_HAS_EXPIRED: 'Platnosť certifikátu vypršala.',
  CERT_NOT_YET_VALID: 'Certifikát ešte nie je platný (skontroluj aj čas v počítači).',
  ERR_TLS_CERT_ALTNAME_INVALID: 'Certifikát nepatrí k zadanému názvu servera.',
};

export function explainVerify(code?: string): string {
  if (!code) return '';
  return VERIFY_TEXT[code] ?? `Chyba certifikátu: ${code}`;
}

export function explainError(err: unknown, security: Security): Omit<ProbeResult, 'ok'> {
  const e = err as NodeJS.ErrnoException & { reason?: string };
  const code = e?.code ?? '';
  const msg = String(e?.message ?? err);
  const low = msg.toLowerCase();
  if (code === 'ENOTFOUND' || code === 'EAI_AGAIN')
    return { message: 'Server sa nenašiel (DNS).', hint: 'Skontroluj názov servera.' };
  if (code === 'ECONNREFUSED')
    return { message: 'Server odmietol spojenie.', hint: 'Skontroluj port a či je služba na serveri zapnutá.' };
  if (code === 'ETIMEDOUT' || code === 'TIMEOUT')
    return { message: 'Server neodpovedá (časový limit).', hint: 'Port môže blokovať firewall, alebo je zlý port/šifrovanie.' };
  if (code === 'ECONNRESET')
    return {
      message: 'Server ukončil spojenie počas nadväzovania.',
      hint: 'Pri starších serveroch (Zimbra) zapni „Kompatibilita so starými servermi“.',
      suggestLegacy: true,
    };
  if (low.includes('wrong version number') || low.includes('packet length too long') || low.includes('http request'))
    return {
      message: 'Na tomto porte server nehovorí priamo cez SSL/TLS.',
      hint: security === 'ssl' ? 'Skús šifrovanie STARTTLS (IMAP 143, SMTP 587).' : 'Skontroluj port a šifrovanie.',
    };
  if (
    low.includes('unsupported protocol') ||
    low.includes('protocol version') ||
    low.includes('no protocols available') ||
    low.includes('dh key too small') ||
    low.includes('ee key too small') ||
    low.includes('ca md too weak') ||
    low.includes('no ciphers available') ||
    low.includes('handshake failure') ||
    low.includes('no shared cipher') ||
    low.includes('unsupported_protocol') ||
    low.includes('no_supported_versions') ||
    low.includes('no_ciphers')
  )
    return {
      message: 'Server používa staré šifrovanie (TLS 1.0/1.1 alebo slabé šifry/kľúče).',
      hint: 'Zapni „Kompatibilita so starými servermi (Zimbra)“ v nastaveniach konta.',
      suggestLegacy: true,
    };
  if (low.includes('starttls'))
    return { message: 'Server nepodporuje STARTTLS.', hint: 'Skús šifrovanie SSL/TLS (port 993 / 465).' };
  return { message: `Spojenie zlyhalo: ${msg}` };
}

// ---------------------------------------------------------------------------
// Sonda: pripojí sa, zistí certifikát a protokol, nič neprihlasuje

function readUntil(sock: net.Socket, done: (text: string) => boolean, timeoutMs: number): Promise<string> {
  return new Promise((resolve, reject) => {
    let buf = '';
    const timer = setTimeout(() => finish(Object.assign(new Error('timeout'), { code: 'TIMEOUT' })), timeoutMs);
    const onData = (d: Buffer) => {
      buf += d.toString('latin1');
      if (done(buf)) finish();
    };
    const onErr = (e: Error) => finish(e);
    const onEnd = () => finish(Object.assign(new Error('Server ukončil spojenie'), { code: 'ECONNRESET' }));
    function finish(e?: Error) {
      clearTimeout(timer);
      sock.off('data', onData);
      sock.off('error', onErr);
      sock.off('end', onEnd);
      if (e) reject(e);
      else resolve(buf);
    }
    sock.on('data', onData);
    sock.once('error', onErr);
    sock.once('end', onEnd);
  });
}

const lastLine = (t: string) => {
  const lines = t.split('\r\n').filter(Boolean);
  return lines[lines.length - 1] ?? '';
};
const smtpComplete = (t: string) => /^\d{3} /m.test(lastLine(t)) && t.endsWith('\r\n');

async function startTls(sock: net.Socket, proto: Protocol, timeoutMs: number): Promise<void> {
  if (proto === 'imap') {
    const greet = await readUntil(sock, (t) => t.includes('\r\n'), timeoutMs);
    if (!/^\* (OK|PREAUTH)/i.test(greet)) throw new Error(`Neočakávaný pozdrav IMAP: ${greet.trim()}`);
    sock.write('A1 STARTTLS\r\n');
    const res = await readUntil(sock, (t) => /^A1 /m.test(t) && t.endsWith('\r\n'), timeoutMs);
    if (!/^A1 OK/im.test(res)) throw new Error(`Server odmietol STARTTLS: ${lastLine(res)}`);
  } else {
    const greet = await readUntil(sock, smtpComplete, timeoutMs);
    if (!greet.startsWith('220')) throw new Error(`Neočakávaný pozdrav SMTP: ${greet.trim()}`);
    sock.write('EHLO mailklient\r\n');
    const ehlo = await readUntil(sock, smtpComplete, timeoutMs);
    if (!/STARTTLS/i.test(ehlo)) throw new Error('Server nepodporuje STARTTLS.');
    sock.write('STARTTLS\r\n');
    const res = await readUntil(sock, smtpComplete, timeoutMs);
    if (!res.startsWith('220')) throw new Error(`Server odmietol STARTTLS: ${lastLine(res)}`);
  }
}

function tcpConnect(host: string, port: number, timeoutMs: number): Promise<net.Socket> {
  return new Promise((resolve, reject) => {
    const s = net.connect({ host, port });
    const t = setTimeout(() => {
      s.destroy();
      reject(Object.assign(new Error('timeout'), { code: 'TIMEOUT' }));
    }, timeoutMs);
    s.once('connect', () => {
      clearTimeout(t);
      resolve(s);
    });
    s.once('error', (e) => {
      clearTimeout(t);
      reject(e);
    });
  });
}

function tlsHandshake(opts: tls.ConnectionOptions, timeoutMs: number): Promise<tls.TLSSocket> {
  return new Promise((resolve, reject) => {
    const s = tls.connect(opts);
    const t = setTimeout(() => {
      s.destroy();
      reject(Object.assign(new Error('timeout'), { code: 'TIMEOUT' }));
    }, timeoutMs);
    s.once('secureConnect', () => {
      clearTimeout(t);
      resolve(s);
    });
    s.once('error', (e) => {
      clearTimeout(t);
      reject(e);
    });
  });
}

/**
 * Otestuje server: spojenie, šifrovanie a certifikát. Overenie reťazca sa pri sonde
 * nevynucuje, aby sme vedeli používateľovi ukázať odtlačok; o dôvere rozhoduje politika konta.
 */
export async function probeServer(cfg: ServerConfig, proto: Protocol, timeoutMs = 15000): Promise<ProbeResult> {
  let plain: net.Socket | undefined;
  let secure: tls.TLSSocket | undefined;
  try {
    if (cfg.security === 'none') {
      plain = await tcpConnect(cfg.host, cfg.port, timeoutMs);
      const g = await readUntil(plain, (t) => t.includes('\r\n'), timeoutMs);
      return {
        ok: true,
        protocol: null,
        message: `Spojenie bez šifrovania funguje. Server: ${g.trim().slice(0, 80)}`,
        hint: 'Heslo pôjde po sieti nešifrované. Používaj len v dôveryhodnej sieti.',
      };
    }
    const base = { ...buildTlsOptions(cfg.host, cfg.tls), rejectUnauthorized: false };
    if (cfg.security === 'ssl') {
      secure = await tlsHandshake({ ...base, host: cfg.host, port: cfg.port }, timeoutMs);
    } else {
      plain = await tcpConnect(cfg.host, cfg.port, timeoutMs);
      await startTls(plain, proto, timeoutMs);
      secure = await tlsHandshake({ ...base, socket: plain }, timeoutMs);
    }
    const peer = secure.getPeerCertificate();
    const cert = peer && peer.fingerprint256 ? certInfo(peer) : undefined;
    const hostnameMismatch = !!peer && !!tls.checkServerIdentity(cfg.host, peer);
    const verifyError = secure.authorized ? undefined : String(secure.authorizationError ?? '');
    const protocol = secure.getProtocol();
    const cipher = secure.getCipher()?.name;
    const err = trustError(cfg.tls, {
      authorized: secure.authorized,
      hostnameMismatch,
      fingerprint256: cert?.fingerprint256,
    });
    const oldProto = protocol === 'TLSv1' || protocol === 'TLSv1.1';
    const proto_ = `${protocol ?? '?'}${cipher ? `, ${cipher}` : ''}`;
    if (err) {
      const why =
        cfg.tls.certMode === 'strict'
          ? [explainVerify(verifyError), hostnameMismatch ? explainVerify('ERR_TLS_CERT_ALTNAME_INVALID') : '']
              .filter(Boolean)
              .join(' ')
          : '';
      return {
        ok: false,
        needsTrust: true,
        protocol,
        cipher,
        cert,
        verifyError,
        hostnameMismatch,
        message: `${err} ${why}`.trim(),
        hint: 'Ak serveru dôveruješ (napr. firemná Zimbra), over odtlačok a ulož výnimku pre toto konto.',
      };
    }
    return {
      ok: true,
      protocol,
      cipher,
      cert,
      verifyError,
      hostnameMismatch,
      message: `Zabezpečené spojenie funguje (${proto_}).`,
      hint: oldProto ? 'Server používa zastaraný protokol TLS. Odporúča sa aktualizovať server.' : undefined,
    };
  } catch (e) {
    return { ok: false, ...explainError(e, cfg.security) };
  } finally {
    secure?.destroy();
    plain?.destroy();
  }
}

/**
 * Automatické zistenie: skúsi SSL a potom STARTTLS na štandardných portoch,
 * najprv moderne, a ak handshake zlyhá na starom protokole, aj v režime pre staré servery.
 */
export async function autodetect(
  host: string,
  proto: Protocol,
  timeoutMs = 8000,
): Promise<{ config?: ServerConfig; result: ProbeResult; tried: string[] }> {
  const order: Array<[Security, number]> =
    proto === 'imap' ? [['ssl', 993], ['starttls', 143]] : [['ssl', 465], ['starttls', 587]];
  const tried: string[] = [];
  let last: ProbeResult = { ok: false, message: 'Nič sa nepodarilo.' };
  for (const [security, port] of order) {
    for (const legacy of [false, true]) {
      const cfg: ServerConfig = { host, port, security, tls: { certMode: 'strict', legacy } };
      const r = await probeServer(cfg, proto, timeoutMs);
      tried.push(`${security}:${port}${legacy ? ' (staré TLS)' : ''} → ${r.ok ? 'OK' : r.needsTrust ? 'certifikát' : r.message}`);
      // spojenie funguje (aj keď treba potvrdiť certifikát) → máme výsledok
      if (r.ok || r.needsTrust) return { config: cfg, result: r, tried };
      last = r;
      if (!r.suggestLegacy) break; // legacy má zmysel skúšať len pri chybe protokolu/šifier
    }
  }
  return { result: last, tried };
}
