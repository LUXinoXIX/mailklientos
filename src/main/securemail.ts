// Podpisovanie a šifrovanie pošty: PGP/MIME (RFC 3156) a S/MIME (RFC 8551), čítanie šifrovaných a overenie podpísaných správ.
import crypto from 'node:crypto';
import tls from 'node:tls';
import MailComposer from 'nodemailer/lib/mail-composer';
import * as openpgp from 'openpgp';
import forge from 'node-forge';
import * as store from './cryptostore';
import { log } from './diagnostics';

export type SecureMode = 'pgp' | 'smime';
export interface SecureOptions {
  sign?: SecureMode;
  encrypt?: SecureMode;
}

/** Výsledok čítania zabezpečenej správy (zobrazí sa v čítacom paneli). */
export interface SecureInfo {
  encrypted?: SecureMode;
  decryptError?: string;
  signed?: SecureMode;
  signatureValid?: boolean;
  signer?: string;
  /** certifikát od dôveryhodnej autority (S/MIME) alebo známy kľúč (PGP) */
  trusted?: boolean;
  signatureNote?: string;
}

const CRLF = '\r\n';
const boundary = () => `----=_MailKlient_${crypto.randomBytes(12).toString('hex')}`;

// ---------------------------------------------------------------------------
// pomocné funkcie MIME (pracujeme s „binary“ reťazcami, aby sa nezmenil ani jeden bajt)

function splitHead(raw: string): { head: string; body: string } {
  const i = raw.indexOf('\r\n\r\n');
  const j = raw.indexOf('\n\n');
  if (i >= 0 && (j < 0 || i <= j)) return { head: raw.slice(0, i), body: raw.slice(i + 4) };
  if (j >= 0) return { head: raw.slice(0, j), body: raw.slice(j + 2) };
  return { head: raw, body: '' };
}

function headerLines(head: string): string[] {
  return head.replace(/\r?\n[ \t]+/g, ' ').split(/\r?\n/).filter(Boolean);
}

function headerValue(head: string, name: string): string {
  const l = headerLines(head).find((x) => x.toLowerCase().startsWith(`${name.toLowerCase()}:`));
  return l ? l.slice(name.length + 1).trim() : '';
}

function param(value: string, name: string): string {
  const m = new RegExp(`;\\s*${name}\\s*=\\s*("([^"]*)"|[^;\\s]+)`, 'i').exec(value);
  return m ? (m[2] ?? m[1]!).trim() : '';
}

/** Časti multipart presne tak, ako sú v zdroji (pre overenie podpisu). */
function parts(body: string, b: string): string[] {
  const out: string[] = [];
  const delim = `--${b}`;
  let pos = body.indexOf(delim);
  if (pos < 0) return out;
  for (;;) {
    let start = body.indexOf('\n', pos);
    if (start < 0) break;
    start += 1;
    let next = body.indexOf(`\n${delim}`, start);
    if (next < 0) break;
    let end = next;
    if (body[end - 1] === '\r') end -= 1;
    out.push(body.slice(start, end));
    pos = next + 1;
    if (body.startsWith(`${delim}--`, pos)) break;
  }
  return out;
}

const toCrlf = (s: string) => s.replace(/\r?\n/g, CRLF);

/** Obsah správy (text, HTML, prílohy) ako samostatná entita MIME – to sa podpisuje/šifruje. */
async function contentEntity(mail: Record<string, unknown>): Promise<string> {
  const { text, html, attachments, attachDataUrls } = mail;
  const raw: Buffer = await new MailComposer({ text, html, attachments, attachDataUrls } as never).compile().build();
  const { head, body } = splitHead(raw.toString('binary'));
  const keep = headerLines(head).filter((l) => /^content-/i.test(l));
  return toCrlf(`${keep.join(CRLF)}${CRLF}${CRLF}${body}`);
}

/** Hlavičky správy (Od, Komu, Predmet, …) bez Content-*; nový obsah sa pridá za ne. */
async function outerHeaders(mail: Record<string, unknown>): Promise<string> {
  const { text: _t, html: _h, attachments: _a, attachDataUrls: _d, icalEvent: _i, ...rest } = mail;
  const raw: Buffer = await new MailComposer({ ...rest, text: '' } as never).compile().build();
  const { head } = splitHead(raw.toString('binary'));
  return headerLines(head)
    .filter((l) => !/^(content-|mime-version)/i.test(l))
    .join(CRLF);
}

const b64 = (bin: string) => (Buffer.from(bin, 'binary').toString('base64').match(/.{1,76}/g) ?? []).join(CRLF);

function addressList(v: unknown): string[] {
  const s = Array.isArray(v) ? v.join(',') : String(v ?? '');
  return [...s.matchAll(/[\w.+'-]+@[\w-]+(\.[\w-]+)+/g)].map((m) => m[0].toLowerCase());
}

// ---------------------------------------------------------------------------
// PGP

async function pgpSignEntity(entity: string, signer: openpgp.PrivateKey): Promise<string> {
  const b = boundary();
  const sig = await openpgp.sign({
    message: await openpgp.createMessage({ binary: Buffer.from(entity, 'binary') }),
    signingKeys: signer,
    detached: true,
    format: 'armored',
    config: { preferredHashAlgorithm: openpgp.enums.hash.sha256 },
  });
  return [
    `Content-Type: multipart/signed; micalg=pgp-sha256; protocol="application/pgp-signature"; boundary="${b}"`,
    '',
    `--${b}`,
    entity,
    `--${b}`,
    'Content-Type: application/pgp-signature; name="signature.asc"',
    'Content-Description: OpenPGP digital signature',
    'Content-Disposition: attachment; filename="signature.asc"',
    '',
    toCrlf(String(sig)),
    `--${b}--`,
    '',
  ].join(CRLF);
}

async function pgpEncryptEntity(entity: string, keys: openpgp.Key[], signer?: openpgp.PrivateKey): Promise<string> {
  const b = boundary();
  const enc = await openpgp.encrypt({
    message: await openpgp.createMessage({ binary: Buffer.from(entity, 'binary') }),
    encryptionKeys: keys,
    signingKeys: signer,
    format: 'armored',
  });
  return [
    `Content-Type: multipart/encrypted; protocol="application/pgp-encrypted"; boundary="${b}"`,
    '',
    'This is an OpenPGP/MIME encrypted message (RFC 4880 and 3156)',
    `--${b}`,
    'Content-Type: application/pgp-encrypted',
    'Content-Description: PGP/MIME version identification',
    '',
    'Version: 1',
    '',
    `--${b}`,
    'Content-Type: application/octet-stream; name="encrypted.asc"',
    'Content-Description: OpenPGP encrypted message',
    'Content-Disposition: inline; filename="encrypted.asc"',
    '',
    toCrlf(String(enc)),
    `--${b}--`,
    '',
  ].join(CRLF);
}

// ---------------------------------------------------------------------------
// S/MIME

function smimeSignEntity(entity: string, own: NonNullable<ReturnType<typeof store.smimeOwnFor>>): string {
  const p7 = forge.pkcs7.createSignedData();
  p7.content = forge.util.createBuffer(entity, 'raw');
  p7.addCertificate(own.cert);
  for (const c of own.chain) p7.addCertificate(c);
  p7.addSigner({
    key: own.key,
    certificate: own.cert,
    digestAlgorithm: forge.pki.oids.sha256!,
    authenticatedAttributes: [
      { type: forge.pki.oids.contentType!, value: forge.pki.oids.data },
      { type: forge.pki.oids.messageDigest! },
      { type: forge.pki.oids.signingTime!, value: new Date() as unknown as string },
    ],
  });
  p7.sign({ detached: true });
  const der = forge.asn1.toDer(p7.toAsn1()).getBytes();
  const b = boundary();
  return [
    `Content-Type: multipart/signed; protocol="application/pkcs7-signature"; micalg=sha-256; boundary="${b}"`,
    '',
    'This is a cryptographically signed message in MIME format.',
    '',
    `--${b}`,
    entity,
    `--${b}`,
    'Content-Type: application/pkcs7-signature; name="smime.p7s"',
    'Content-Transfer-Encoding: base64',
    'Content-Disposition: attachment; filename="smime.p7s"',
    'Content-Description: S/MIME Cryptographic Signature',
    '',
    b64(der),
    `--${b}--`,
    '',
  ].join(CRLF);
}

function smimeEncryptEntity(entity: string, certs: forge.pki.Certificate[]): string {
  const p7 = forge.pkcs7.createEnvelopedData();
  for (const c of certs) p7.addRecipient(c);
  p7.content = forge.util.createBuffer(entity, 'raw');
  p7.encrypt(undefined, forge.pki.oids['aes256-CBC']);
  const der = forge.asn1.toDer(p7.toAsn1()).getBytes();
  return [
    'Content-Type: application/pkcs7-mime; smime-type=enveloped-data; name="smime.p7m"',
    'Content-Transfer-Encoding: base64',
    'Content-Disposition: attachment; filename="smime.p7m"',
    'Content-Description: S/MIME Encrypted Message',
    '',
    b64(der),
    '',
  ].join(CRLF);
}

// ---------------------------------------------------------------------------
// odoslanie

export class SecureError extends Error {}

/**
 * Zostaví podpísanú a/alebo zašifrovanú správu. Šifruje sa aj pre odosielateľa (kópia v Odoslaných je čitateľná).
 * Vráti surový text správy a obálku pre SMTP.
 */
export async function buildSecure(mail: Record<string, unknown>, opts: SecureOptions, fromEmail: string) {
  let entity = await contentEntity(mail);
  const recipients = [...new Set([...addressList(mail.to), ...addressList(mail.cc), ...addressList(mail.bcc)])];
  const mode = opts.encrypt ?? opts.sign;
  if (mode === 'pgp') {
    const signer = opts.sign === 'pgp' ? await store.pgpPrivateFor(fromEmail) : undefined;
    if (opts.sign === 'pgp' && !signer) throw new SecureError('Na podpis chýba váš kľúč PGP. Vytvorte ho v Nastaveniach → Šifrovanie.');
    if (opts.encrypt === 'pgp') {
      const { keys, missing } = await store.pgpPublicFor(recipients);
      if (missing.length) throw new SecureError(`Chýba verejný kľúč PGP pre: ${missing.join(', ')}. Importujte ho (Nastavenia → Šifrovanie) alebo pošlite správu nešifrovanú.`);
      const own = await store.pgpPublicFor([fromEmail]);
      entity = await pgpEncryptEntity(entity, [...keys, ...own.keys], signer ?? undefined);
    } else if (signer) entity = await pgpSignEntity(entity, signer);
  } else if (mode === 'smime') {
    const own = store.smimeOwnFor(fromEmail);
    if (opts.sign === 'smime') {
      if (!own) throw new SecureError('Na podpis chýba váš certifikát S/MIME. Importujte súbor .p12/.pfx v Nastaveniach → Šifrovanie.');
      entity = smimeSignEntity(entity, own);
    }
    if (opts.encrypt === 'smime') {
      const { certs, missing } = store.smimeCertsFor(recipients);
      if (missing.length) throw new SecureError(`Chýba certifikát S/MIME pre: ${missing.join(', ')}. Certifikát sa uloží automaticky, keď vám dotyčný pošle podpísanú správu.`);
      entity = smimeEncryptEntity(entity, own ? [...certs, own.cert] : certs);
    }
  }
  const head = await outerHeaders(mail);
  const raw = `${head}${CRLF}MIME-Version: 1.0${CRLF}${entity}`;
  return { raw: Buffer.from(raw, 'binary'), recipients };
}

// ---------------------------------------------------------------------------
// čítanie: dešifrovanie a overenie podpisu

let roots: forge.pki.CAStore | null = null;
function caStore() {
  if (roots) return roots;
  roots = forge.pki.createCaStore();
  for (const pem of tls.rootCertificates) {
    try {
      roots.addCertificate(forge.pki.certificateFromPem(pem));
    } catch {
      /* forge nepozná niektoré typy kľúčov (EC) – takú autoritu preskočíme */
    }
  }
  return roots;
}

const DIGESTS: Record<string, 'sha1' | 'sha256' | 'sha384' | 'sha512'> = {
  [forge.pki.oids.sha1!]: 'sha1',
  [forge.pki.oids.sha256!]: 'sha256',
  [forge.pki.oids.sha384!]: 'sha384',
  [forge.pki.oids.sha512!]: 'sha512',
};

/** Overenie podpisu S/MIME (forge to sám nevie): odtlačok obsahu a podpis podpísaných atribútov. */
function smimeVerify(content: string, p7der: string): SecureInfo {
  const info: SecureInfo = { signed: 'smime', signatureValid: false };
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const msg = forge.pkcs7.messageFromAsn1(forge.asn1.fromDer(p7der)) as any;
    const certs: forge.pki.Certificate[] = msg.certificates ?? [];
    const si = msg.rawCapture?.signerInfos?.[0] as forge.asn1.Asn1 | undefined;
    if (!si || !Array.isArray(si.value)) throw new Error('chýba podpis');
    const kids = si.value as forge.asn1.Asn1[];
    const sid = kids[1]!;
    const serial = forge.util.bytesToHex(((sid.value as forge.asn1.Asn1[])[1]?.value as string) ?? '');
    const cert = certs.find((c) => c.serialNumber.replace(/^0+/, '') === serial.replace(/^0+/, '')) ?? certs[0];
    if (!cert) throw new Error('chýba certifikát odosielateľa');
    const digOid = forge.asn1.derToOid(((kids[2]!.value as forge.asn1.Asn1[])[0]!.value as string));
    const alg = DIGESTS[digOid] ?? 'sha256';
    const attrs = kids.find((k) => k.tagClass === forge.asn1.Class.CONTEXT_SPECIFIC && k.type === 0);
    const sigOctets = kids.find((k, i) => i > 3 && k.tagClass === forge.asn1.Class.UNIVERSAL && k.type === forge.asn1.Type.OCTETSTRING);
    if (!sigOctets) throw new Error('chýba hodnota podpisu');
    const contentHash = crypto.createHash(alg).update(Buffer.from(content, 'binary')).digest('binary');
    const md = forge.md[alg].create();
    if (attrs) {
      const digestAttr = (attrs.value as forge.asn1.Asn1[]).find(
        (a) => forge.asn1.derToOid(((a.value as forge.asn1.Asn1[])[0]!.value as string)) === forge.pki.oids.messageDigest,
      );
      const stated = digestAttr ? ((((digestAttr.value as forge.asn1.Asn1[])[1]!.value as forge.asn1.Asn1[])[0]!.value as string)) : '';
      if (stated !== contentHash) {
        info.signatureNote = 'Obsah správy bol po podpísaní zmenený.';
        return info;
      }
      const set = forge.asn1.create(forge.asn1.Class.UNIVERSAL, forge.asn1.Type.SET, true, attrs.value as forge.asn1.Asn1[]);
      md.update(forge.asn1.toDer(set).getBytes());
    } else md.update(content);
    const pub = cert.publicKey as forge.pki.rsa.PublicKey;
    if (typeof pub.verify !== 'function') {
      info.signatureNote = 'Typ kľúča odosielateľa (EC) zatiaľ nevieme overiť.';
      return info;
    }
    info.signatureValid = pub.verify(md.digest().bytes(), sigOctets.value as string);
    const cn = (cert.subject.getField('CN')?.value as string) ?? '';
    const email = (cert.subject.getField('E')?.value as string) ?? '';
    info.signer = email ? `${cn} <${email}>` : cn;
    const now = new Date();
    if (now < cert.validity.notBefore || now > cert.validity.notAfter) info.signatureNote = 'Certifikát odosielateľa nie je platný (vypršal alebo ešte neplatí).';
    try {
      info.trusted = forge.pki.verifyCertificateChain(caStore(), [cert, ...certs.filter((c) => c !== cert)]);
    } catch {
      info.trusted = false;
    }
    if (info.signatureValid) store.smimeRemember(cert);
  } catch (e) {
    info.signatureNote = `Podpis sa nepodarilo overiť: ${(e as Error).message}`;
  }
  return info;
}

async function pgpVerify(content: string, armoredSig: string): Promise<SecureInfo> {
  const info: SecureInfo = { signed: 'pgp', signatureValid: false };
  try {
    const keys = await store.pgpAllPublic();
    const signature = await openpgp.readSignature({ armoredSignature: armoredSig });
    const r = await openpgp.verify({ message: await openpgp.createMessage({ binary: Buffer.from(content, 'binary') }), signature, verificationKeys: keys });
    const s = r.signatures[0];
    if (!s) throw new Error('bez podpisu');
    const id = s.keyID.toHex().toUpperCase();
    const key = keys.find((k) => k.getKeys(s.keyID).length);
    try {
      await s.verified;
      info.signatureValid = true;
      info.trusted = !!key;
      info.signer = key?.getUserIDs()[0] ?? `kľúč ${id}`;
    } catch (e) {
      if (!key) info.signatureNote = `Podpis kľúčom ${id}, ktorý nemáte uložený – importujte verejný kľúč odosielateľa.`;
      else info.signatureNote = `Podpis nesedí: ${(e as Error).message}`;
    }
  } catch (e) {
    info.signatureNote = `Podpis sa nepodarilo overiť: ${(e as Error).message}`;
  }
  return info;
}

/**
 * Ak je správa šifrovaná alebo podpísaná, vráti zdroj na zobrazenie (dešifrovaný) a informácie o zabezpečení.
 * Hlavičky (Od, Predmet…) zostávajú z pôvodnej správy.
 */
export async function unwrap(source: Buffer): Promise<{ source: Buffer; info?: SecureInfo }> {
  let raw = source.toString('binary');
  const info: SecureInfo = {};
  for (let depth = 0; depth < 3; depth++) {
    const { head, body } = splitHead(raw);
    const ct = headerValue(head, 'content-type');
    const type = ct.split(';')[0]!.trim().toLowerCase();
    const outer = headerLines(head).filter((l) => !/^(content-|mime-version)/i.test(l)).join(CRLF);
    if (type === 'multipart/encrypted' && /pgp-encrypted/i.test(param(ct, 'protocol'))) {
      info.encrypted = 'pgp';
      const ps = parts(body, param(ct, 'boundary'));
      const armored = ps.map((p) => splitHead(p).body).find((b) => b.includes('BEGIN PGP MESSAGE'));
      if (!armored) return { source, info: { ...info, decryptError: 'Šifrovaná časť chýba.' } };
      try {
        const keys = await store.pgpAllPrivate();
        if (!keys.length) return { source, info: { ...info, decryptError: 'Správa je zašifrovaná PGP, ale nemáte uložený súkromný kľúč.' } };
        const r = await openpgp.decrypt({
          message: await openpgp.readMessage({ armoredMessage: armored }),
          decryptionKeys: keys,
          verificationKeys: await store.pgpAllPublic(),
          format: 'binary',
        });
        const inner = Buffer.from(r.data as Uint8Array).toString('binary');
        if (r.signatures.length) {
          info.signed = 'pgp';
          try {
            await r.signatures[0]!.verified;
            info.signatureValid = true;
            info.trusted = true;
          } catch {
            info.signatureValid = false;
            info.signatureNote = 'Podpis sa nepodarilo overiť (chýba verejný kľúč odosielateľa alebo nesedí).';
          }
        }
        raw = `${outer}${CRLF}${inner}`;
        continue;
      } catch (e) {
        log('PGP: dešifrovanie zlyhalo', String(e));
        return { source, info: { ...info, decryptError: `Správu sa nepodarilo dešifrovať: ${(e as Error).message}` } };
      }
    }
    if ((type === 'application/pkcs7-mime' || type === 'application/x-pkcs7-mime') && !/signed-data/i.test(param(ct, 'smime-type'))) {
      info.encrypted = 'smime';
      try {
        const der = Buffer.from(body.replace(/[^A-Za-z0-9+/=]/g, ''), 'base64').toString('binary');
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const p7 = forge.pkcs7.messageFromAsn1(forge.asn1.fromDer(der)) as any;
        const owns = store.smimeAllOwn();
        if (!owns.length) return { source, info: { ...info, decryptError: 'Správa je zašifrovaná S/MIME, ale nemáte importovaný certifikát (.p12).' } };
        let done = false;
        for (const o of owns) {
          const rcpt = p7.findRecipient(o.cert);
          if (!rcpt) continue;
          p7.decrypt(rcpt, o.key);
          done = true;
          break;
        }
        if (!done) return { source, info: { ...info, decryptError: 'Správa nie je zašifrovaná pre žiadny z vašich certifikátov.' } };
        const inner = p7.content.getBytes() as string;
        raw = `${outer}${CRLF}${inner}`;
        continue;
      } catch (e) {
        return { source, info: { ...info, decryptError: `Správu sa nepodarilo dešifrovať: ${(e as Error).message}` } };
      }
    }
    if ((type === 'application/pkcs7-mime' || type === 'application/x-pkcs7-mime') && /signed-data/i.test(param(ct, 'smime-type'))) {
      // nepriehľadne podpísaná správa: obsah je vnútri podpisu
      try {
        const der = Buffer.from(body.replace(/[^A-Za-z0-9+/=]/g, ''), 'base64').toString('binary');
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const p7 = forge.pkcs7.messageFromAsn1(forge.asn1.fromDer(der)) as any;
        const inner = (p7.rawCapture?.content?.value?.[0]?.value as string) ?? '';
        Object.assign(info, smimeVerify(inner, der));
        raw = `${outer}${CRLF}${inner}`;
        continue;
      } catch (e) {
        Object.assign(info, { signed: 'smime', signatureValid: false, signatureNote: String((e as Error).message) });
        break;
      }
    }
    if (type === 'multipart/signed') {
      const proto = param(ct, 'protocol').toLowerCase();
      const ps = parts(body, param(ct, 'boundary'));
      if (ps.length >= 2) {
        const content = ps[0]!;
        const sigPart = splitHead(ps[1]!);
        if (proto.includes('pgp-signature')) Object.assign(info, await pgpVerify(content, sigPart.body));
        else if (proto.includes('pkcs7-signature')) {
          const der = Buffer.from(sigPart.body.replace(/[^A-Za-z0-9+/=]/g, ''), 'base64').toString('binary');
          Object.assign(info, smimeVerify(content, der));
        }
        // zobrazí sa len obsah bez podpisovej prílohy
        raw = `${outer}${CRLF}${content}`;
      }
      break;
    }
    break;
  }
  if (!info.encrypted && !info.signed) return { source };
  return { source: Buffer.from(raw, 'binary'), info };
}

/** Vložená (inline) správa PGP v texte. */
export async function decryptInline(text: string): Promise<string | null> {
  const m = /-----BEGIN PGP MESSAGE-----[\s\S]+?-----END PGP MESSAGE-----/.exec(text);
  if (!m) return null;
  try {
    const keys = await store.pgpAllPrivate();
    if (!keys.length) return null;
    const r = await openpgp.decrypt({ message: await openpgp.readMessage({ armoredMessage: m[0] }), decryptionKeys: keys });
    return text.replace(m[0], String(r.data));
  } catch {
    return null;
  }
}
