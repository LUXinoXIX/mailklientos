// Kľúče a certifikáty na šifrovanie: vlastné kľúče PGP (s heslom uloženým cez systémové úložisko),
// verejné kľúče PGP korešpondentov, vlastný certifikát S/MIME (.p12) a certifikáty korešpondentov.
import fs from 'node:fs';
import path from 'node:path';
import { app, safeStorage } from 'electron';
import * as openpgp from 'openpgp';
import forge from 'node-forge';
import { log } from './diagnostics';

export interface PgpOwnKey {
  fingerprint: string;
  userIds: string[];
  emails: string[];
  armoredPrivate: string;
  armoredPublic: string;
  /** heslo ku kľúču, zašifrované systémovým úložiskom (base64) */
  passphraseEnc?: string;
  created: string;
}
export interface PgpContactKey {
  fingerprint: string;
  emails: string[];
  userIds: string[];
  armoredPublic: string;
}
export interface SmimeOwn {
  emails: string[];
  subject: string;
  issuer: string;
  validTo: string;
  certPem: string;
  /** súkromný kľúč PEM, zašifrovaný systémovým úložiskom (base64) */
  keyEnc: string;
  chainPem: string[];
}
export interface SmimeContact {
  emails: string[];
  subject: string;
  issuer: string;
  validTo: string;
  certPem: string;
}
interface Store {
  pgpOwn: PgpOwnKey[];
  pgpContacts: PgpContactKey[];
  smimeOwn: SmimeOwn[];
  smimeContacts: SmimeContact[];
}

const file = () => path.join(app.getPath('userData'), 'crypto.json');
let cache: Store | null = null;

function load(): Store {
  if (cache) return cache;
  try {
    cache = { pgpOwn: [], pgpContacts: [], smimeOwn: [], smimeContacts: [], ...(JSON.parse(fs.readFileSync(file(), 'utf8')) as Partial<Store>) };
  } catch {
    cache = { pgpOwn: [], pgpContacts: [], smimeOwn: [], smimeContacts: [] };
  }
  return cache;
}
function save() {
  fs.writeFileSync(file(), JSON.stringify(load(), null, 2), 'utf8');
}

const seal = (s: string) => (safeStorage.isEncryptionAvailable() ? safeStorage.encryptString(s).toString('base64') : `plain:${Buffer.from(s).toString('base64')}`);
const unseal = (s: string) =>
  s.startsWith('plain:') ? Buffer.from(s.slice(6), 'base64').toString('utf8') : safeStorage.decryptString(Buffer.from(s, 'base64'));

const emailsOf = (userIds: string[]) => userIds.map((u) => /<([^>]+)>/.exec(u)?.[1] ?? (u.includes('@') ? u : '')).filter(Boolean).map((e) => e.toLowerCase());

// ---------------------------------------------------------------------------
// PGP

export async function pgpGenerate(name: string, email: string, passphrase: string) {
  const { privateKey, publicKey } = await openpgp.generateKey({ type: 'ecc', curve: 'curve25519Legacy', userIDs: [{ name, email }], passphrase, format: 'armored' });
  const key = await openpgp.readKey({ armoredKey: publicKey });
  const own: PgpOwnKey = {
    fingerprint: key.getFingerprint().toUpperCase(),
    userIds: key.getUserIDs(),
    emails: [email.toLowerCase()],
    armoredPrivate: privateKey,
    armoredPublic: publicKey,
    passphraseEnc: seal(passphrase),
    created: new Date().toISOString(),
  };
  load().pgpOwn.push(own);
  save();
  return own.fingerprint;
}

/** Import kľúča (.asc): súkromný (s heslom) ide medzi vlastné, verejný ku kontaktom. */
export async function pgpImport(armored: string, passphrase?: string) {
  const keys = await openpgp.readKeys({ armoredKeys: armored });
  let own = 0;
  let pub = 0;
  for (const k of keys) {
    const fp = k.getFingerprint().toUpperCase();
    const userIds = k.getUserIDs();
    if (k.isPrivate()) {
      if (k.isDecrypted() === false) {
        if (!passphrase) throw new Error('Súkromný kľúč je chránený heslom – zadajte ho.');
        await openpgp.decryptKey({ privateKey: k as openpgp.PrivateKey, passphrase });
      }
      const s = load();
      s.pgpOwn = s.pgpOwn.filter((x) => x.fingerprint !== fp);
      s.pgpOwn.push({
        fingerprint: fp,
        userIds,
        emails: emailsOf(userIds),
        armoredPrivate: k.armor(),
        armoredPublic: k.toPublic().armor(),
        passphraseEnc: passphrase ? seal(passphrase) : undefined,
        created: k.getCreationTime().toISOString(),
      });
      own++;
    } else {
      const s = load();
      s.pgpContacts = s.pgpContacts.filter((x) => x.fingerprint !== fp);
      s.pgpContacts.push({ fingerprint: fp, userIds, emails: emailsOf(userIds), armoredPublic: k.armor() });
      pub++;
    }
  }
  save();
  return { own, pub };
}

export async function pgpPrivateFor(email: string): Promise<openpgp.PrivateKey | null> {
  const k = load().pgpOwn.find((x) => x.emails.includes(email.toLowerCase())) ?? load().pgpOwn[0];
  if (!k) return null;
  const key = await openpgp.readPrivateKey({ armoredKey: k.armoredPrivate });
  if (key.isDecrypted()) return key;
  if (!k.passphraseEnc) throw new Error('Chýba heslo ku kľúču PGP.');
  return openpgp.decryptKey({ privateKey: key, passphrase: unseal(k.passphraseEnc) });
}

export async function pgpAllPrivate(): Promise<openpgp.PrivateKey[]> {
  const out: openpgp.PrivateKey[] = [];
  for (const k of load().pgpOwn) {
    try {
      const key = await openpgp.readPrivateKey({ armoredKey: k.armoredPrivate });
      out.push(key.isDecrypted() ? key : await openpgp.decryptKey({ privateKey: key, passphrase: unseal(k.passphraseEnc ?? '') }));
    } catch (e) {
      log('PGP: kľúč sa nepodarilo odomknúť', k.fingerprint, String(e));
    }
  }
  return out;
}

/** Verejné kľúče príjemcov; vráti aj zoznam adries bez kľúča. */
export async function pgpPublicFor(emails: string[]) {
  const s = load();
  const keys: openpgp.Key[] = [];
  const missing: string[] = [];
  for (const e of emails.map((x) => x.toLowerCase())) {
    const hit = s.pgpContacts.find((c) => c.emails.includes(e)) ?? s.pgpOwn.find((c) => c.emails.includes(e));
    if (hit) keys.push(await openpgp.readKey({ armoredKey: hit.armoredPublic }));
    else missing.push(e);
  }
  return { keys, missing };
}

export async function pgpAllPublic(): Promise<openpgp.Key[]> {
  const s = load();
  return Promise.all([...s.pgpContacts, ...s.pgpOwn].map((k) => openpgp.readKey({ armoredKey: k.armoredPublic })));
}

/** Kľúč pripojený k prijatej správe (Autocrypt / príloha .asc) si zapamätať. */
export async function pgpRemember(armored: string) {
  try {
    const r = await pgpImport(armored);
    return r.pub > 0;
  } catch {
    return false;
  }
}

// ---------------------------------------------------------------------------
// S/MIME

function certInfo(cert: forge.pki.Certificate) {
  const emails = new Set<string>();
  const e = cert.subject.getField('E')?.value as string | undefined;
  if (e) emails.add(e.toLowerCase());
  const san = cert.getExtension('subjectAltName') as { altNames?: Array<{ type: number; value: string }> } | null;
  for (const a of san?.altNames ?? []) if (a.type === 1) emails.add(a.value.toLowerCase());
  const cn = (cert.subject.getField('CN')?.value as string | undefined) ?? '';
  const icn = (cert.issuer.getField('CN')?.value as string | undefined) ?? (cert.issuer.getField('O')?.value as string | undefined) ?? '';
  return { emails: [...emails], subject: cn, issuer: icn, validTo: cert.validity.notAfter.toISOString() };
}

/** Import vlastného certifikátu S/MIME zo súboru .p12 / .pfx. */
export function smimeImportP12(p12Der: Buffer, password: string) {
  const p12 = forge.pkcs12.pkcs12FromAsn1(forge.asn1.fromDer(p12Der.toString('binary')), password);
  const keyBag = (p12.getBags({ bagType: forge.pki.oids.pkcs8ShroudedKeyBag })[forge.pki.oids.pkcs8ShroudedKeyBag] ?? [])[0]
    ?? (p12.getBags({ bagType: forge.pki.oids.keyBag })[forge.pki.oids.keyBag] ?? [])[0];
  const certs = (p12.getBags({ bagType: forge.pki.oids.certBag })[forge.pki.oids.certBag] ?? []).map((b) => b.cert!).filter(Boolean);
  if (!keyBag?.key || !certs.length) throw new Error('Súbor neobsahuje súkromný kľúč a certifikát.');
  const key = keyBag.key as forge.pki.rsa.PrivateKey;
  // certifikát patriaci ku kľúču
  const mine = certs.find((c) => (c.publicKey as forge.pki.rsa.PublicKey).n.equals(key.n)) ?? certs[0]!;
  const info = certInfo(mine);
  const s = load();
  s.smimeOwn = s.smimeOwn.filter((x) => x.certPem !== forge.pki.certificateToPem(mine));
  s.smimeOwn.push({
    ...info,
    certPem: forge.pki.certificateToPem(mine),
    keyEnc: seal(forge.pki.privateKeyToPem(key)),
    chainPem: certs.filter((c) => c !== mine).map((c) => forge.pki.certificateToPem(c)),
  });
  save();
  return info;
}

export function smimeOwnFor(email: string) {
  const s = load();
  const o = s.smimeOwn.find((x) => x.emails.includes(email.toLowerCase())) ?? s.smimeOwn[0];
  if (!o) return null;
  return {
    cert: forge.pki.certificateFromPem(o.certPem),
    key: forge.pki.privateKeyFromPem(unseal(o.keyEnc)) as forge.pki.rsa.PrivateKey,
    chain: o.chainPem.map((p) => forge.pki.certificateFromPem(p)),
  };
}

export function smimeAllOwn() {
  return load().smimeOwn.map((o) => ({ cert: forge.pki.certificateFromPem(o.certPem), key: forge.pki.privateKeyFromPem(unseal(o.keyEnc)) as forge.pki.rsa.PrivateKey }));
}

export function smimeCertsFor(emails: string[]) {
  const s = load();
  const certs: forge.pki.Certificate[] = [];
  const missing: string[] = [];
  for (const e of emails.map((x) => x.toLowerCase())) {
    const hit = s.smimeContacts.find((c) => c.emails.includes(e)) ?? s.smimeOwn.find((c) => c.emails.includes(e));
    if (hit) certs.push(forge.pki.certificateFromPem(hit.certPem));
    else missing.push(e);
  }
  return { certs, missing };
}

/** Certifikát odosielateľa z podpísanej správy si zapamätať (na neskoršie šifrovanie). */
export function smimeRemember(cert: forge.pki.Certificate) {
  const info = certInfo(cert);
  if (!info.emails.length) return;
  const pem = forge.pki.certificateToPem(cert);
  const s = load();
  if (s.smimeContacts.some((c) => c.certPem === pem) || s.smimeOwn.some((c) => c.certPem === pem)) return;
  s.smimeContacts = s.smimeContacts.filter((c) => !c.emails.some((e) => info.emails.includes(e)) || Date.parse(c.validTo) > Date.parse(info.validTo));
  s.smimeContacts.push({ ...info, certPem: pem });
  save();
}

export function smimeImportCert(pemOrDer: Buffer) {
  const text = pemOrDer.toString('utf8');
  const cert = text.includes('BEGIN CERTIFICATE') ? forge.pki.certificateFromPem(text) : forge.pki.certificateFromAsn1(forge.asn1.fromDer(pemOrDer.toString('binary')));
  smimeRemember(cert);
  return certInfo(cert);
}

// ---------------------------------------------------------------------------
// prehľad pre nastavenia

export function summary() {
  const s = load();
  return {
    pgpOwn: s.pgpOwn.map((k) => ({ fingerprint: k.fingerprint, userIds: k.userIds, created: k.created })),
    pgpContacts: s.pgpContacts.map((k) => ({ fingerprint: k.fingerprint, userIds: k.userIds })),
    smimeOwn: s.smimeOwn.map(({ emails, subject, issuer, validTo }) => ({ emails, subject, issuer, validTo })),
    smimeContacts: s.smimeContacts.map(({ emails, subject, issuer, validTo }) => ({ emails, subject, issuer, validTo })),
  };
}

export function pgpExportPublic(fingerprint: string) {
  const s = load();
  return (s.pgpOwn.find((k) => k.fingerprint === fingerprint) ?? s.pgpContacts.find((k) => k.fingerprint === fingerprint))?.armoredPublic ?? '';
}

export function remove(kind: 'pgpOwn' | 'pgpContacts' | 'smimeOwn' | 'smimeContacts', id: string) {
  const s = load();
  if (kind === 'pgpOwn') s.pgpOwn = s.pgpOwn.filter((k) => k.fingerprint !== id);
  if (kind === 'pgpContacts') s.pgpContacts = s.pgpContacts.filter((k) => k.fingerprint !== id);
  if (kind === 'smimeOwn') s.smimeOwn = s.smimeOwn.filter((k) => k.subject + k.validTo !== id);
  if (kind === 'smimeContacts') s.smimeContacts = s.smimeContacts.filter((k) => k.subject + k.validTo !== id);
  save();
}
