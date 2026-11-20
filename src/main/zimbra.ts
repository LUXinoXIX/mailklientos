// Zimbra: serverové rozhranie (SOAP vo formáte JSON, /service/soap) s rovnakým prihlásením ako IMAP.
// Funguje aj so staršími verziami a neplatným certifikátom (rovnaká politika ako pri IMAP).
import fs from 'node:fs';
import https from 'node:https';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import tls from 'node:tls';
import { isGmail, isOutlookCom, type AccountConfig } from '../core/account';
import { buildTlsOptions, verifySocket } from '../core/tls';
import { getPassword } from './accounts';
import { log } from './diagnostics';
import { UserError } from './mail';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type J = any;

interface Session {
  base: string;
  token: string;
  until: number;
}
const sessions = new Map<string, Session>();
const notZimbra = new Map<string, number>();

function rawRequest(acc: AccountConfig, url: string, body?: string, headers: Record<string, string> = {}): Promise<{ status: number; body: Buffer; headers: http.IncomingHttpHeaders }> {
  const u = new URL(url);
  const sameHost = u.hostname.toLowerCase() === acc.imap.host.toLowerCase();
  const policy = sameHost ? acc.imap.tls : { certMode: 'strict' as const, legacy: false };
  const lib = u.protocol === 'http:' ? http : https;
  return new Promise((resolve, reject) => {
    const req = lib.request(
      u,
      {
        method: body === undefined ? 'GET' : 'POST',
        ...(u.protocol === 'https:' ? buildTlsOptions(u.hostname, policy) : {}),
        headers: { 'User-Agent': 'MailKlient', ...(body !== undefined ? { 'Content-Type': 'application/json; charset=utf-8', 'Content-Length': String(Buffer.byteLength(body)) } : {}), ...headers },
        timeout: 30000,
      },
      (res) => {
        if (res.socket instanceof tls.TLSSocket && policy.certMode === 'pinned') {
          const bad = verifySocket(res.socket, u.hostname, policy);
          if (bad) {
            res.destroy();
            return reject(new UserError(bad));
          }
        }
        const chunks: Buffer[] = [];
        res.on('data', (c: Buffer) => chunks.push(c));
        res.on('end', () => resolve({ status: res.statusCode ?? 0, body: Buffer.concat(chunks), headers: res.headers }));
      },
    );
    req.on('timeout', () => req.destroy(new Error('Server Zimbra neodpovedá.')));
    req.on('error', reject);
    if (body !== undefined) req.write(body);
    req.end();
  });
}

const baseUrl = (acc: AccountConfig) => (acc.zimbraUrl?.trim() || `https://${acc.imap.host}`).replace(/\/+$/, '');

async function soapRaw(acc: AccountConfig, name: string, ns: string, body: J, token?: string): Promise<J> {
  const env = {
    Header: { context: { _jsns: 'urn:zimbra', ...(token ? { authToken: { _content: token } } : {}), format: { type: 'js' }, userAgent: { name: 'MailKlient' } } },
    Body: { [`${name}Request`]: { _jsns: ns, ...body } },
  };
  const r = await rawRequest(acc, `${baseUrl(acc)}/service/soap/${name}Request`, JSON.stringify(env));
  let json: J;
  try {
    json = JSON.parse(r.body.toString('utf8'));
  } catch {
    throw new UserError(r.status === 404 ? 'Server nie je Zimbra (chýba rozhranie /service/soap).' : `Zimbra vrátila neplatnú odpoveď (${r.status}).`);
  }
  const fault = json?.Body?.Fault;
  if (fault) {
    const code = fault.Detail?.Error?.Code ?? '';
    const text = fault.Reason?.Text ?? code;
    const err = new UserError(
      /AUTH_FAILED|auth credentials/i.test(`${code} ${text}`) ? 'Zimbra odmietla prihlásenie.' : `Zimbra: ${text}`,
      /AUTH_FAILED/i.test(code) ? 'Skontroluj heslo konta.' : undefined,
    );
    (err as J).code = code;
    throw err;
  }
  return json?.Body?.[`${name}Response`] ?? {};
}

async function session(acc: AccountConfig): Promise<Session> {
  if (isGmail(acc) || isOutlookCom(acc)) throw new UserError('Konto Gmail ani Outlook.com nie je Zimbra.');
  const s = sessions.get(acc.id);
  if (s && s.until > Date.now() + 60_000) return s;
  const pass = getPassword(acc.id);
  if (!pass) throw new UserError('Pre konto chýba heslo.');
  const r = await soapRaw(acc, 'Auth', 'urn:zimbraAccount', { account: { by: 'name', _content: acc.username.includes('@') ? acc.username : acc.email }, password: { _content: pass } });
  const token = r.authToken?.[0]?._content ?? r.authToken?._content ?? r.authToken;
  if (!token) throw new UserError('Zimbra nevrátila prihlasovací token.');
  const life = Number(r.lifetime ?? 12 * 3600_000);
  const ses = { base: baseUrl(acc), token: String(token), until: Date.now() + Math.min(life, 12 * 3600_000) };
  sessions.set(acc.id, ses);
  return ses;
}

/** Volanie Zimbry s prihlásením (pri vypršanom tokene raz znova). */
export async function soap(acc: AccountConfig, name: string, ns: 'urn:zimbraAccount' | 'urn:zimbraMail', body: J = {}): Promise<J> {
  const s = await session(acc);
  try {
    return await soapRaw(acc, name, ns, body, s.token);
  } catch (e) {
    if (/AUTH_EXPIRED|AUTH_REQUIRED/.test(String((e as J).code ?? ''))) {
      sessions.delete(acc.id);
      return soapRaw(acc, name, ns, body, (await session(acc)).token);
    }
    throw e;
  }
}

/** Je server konta Zimbra? (výsledok sa pamätá; neúspech sa skúsi znova o hodinu) */
export async function isZimbra(acc: AccountConfig): Promise<boolean> {
  if (isGmail(acc) || isOutlookCom(acc)) return false; // Gmail/Outlook.com: heslo nesmie odísť na iný server
  if (sessions.has(acc.id)) return true;
  const n = notZimbra.get(acc.id);
  if (n && Date.now() - n < 3600_000) return false;
  try {
    await session(acc);
    return true;
  } catch (e) {
    log(`Zimbra: ${acc.name} – rozhranie nedostupné`, String((e as Error).message));
    notZimbra.set(acc.id, Date.now());
    return false;
  }
}

export function forget(id: string) {
  sessions.delete(id);
  notZimbra.delete(id);
}

const arr = <T>(x: T | T[] | undefined): T[] => (x === undefined ? [] : Array.isArray(x) ? x : [x]);
/** _attrs vo formáte JSON: objekt alebo pole {n, _content} */
function attrs(x: J): Record<string, string> {
  if (!x) return {};
  if (Array.isArray(x)) return Object.fromEntries(x.map((a: J) => [a.n, a._content]));
  if (Array.isArray(x.a)) return Object.fromEntries(x.a.map((a: J) => [a.n, a._content]));
  return x;
}

// ---------------------------------------------------------------------------
// 1. Globálny adresár (GAL)

export interface GalEntry {
  name: string;
  email: string;
  company?: string;
  title?: string;
  department?: string;
  phone?: string;
  mobile?: string;
  type?: 'account' | 'resource' | 'group';
}

const galEntry = (cn: J): GalEntry => {
  const a = attrs(cn._attrs);
  return {
    name: a.fullName || [a.firstName, a.lastName].filter(Boolean).join(' ') || a.email || '',
    email: a.email ?? a.email2 ?? '',
    company: a.company,
    title: a.jobTitle ?? a.title,
    department: a.department,
    phone: a.workPhone,
    mobile: a.mobilePhone,
    type: a.zimbraCalResType ? 'resource' : a.type === 'group' ? 'group' : 'account',
  };
};

export async function galSearch(acc: AccountConfig, q: string, limit = 50): Promise<GalEntry[]> {
  const r = await soap(acc, 'SearchGal', 'urn:zimbraAccount', { name: q || '.', type: 'all', limit, offset: 0 });
  return arr(r.cn).map(galEntry).filter((x) => x.email);
}

export async function galAutocomplete(acc: AccountConfig, q: string): Promise<GalEntry[]> {
  if (q.trim().length < 2) return [];
  try {
    const r = await soap(acc, 'AutoCompleteGal', 'urn:zimbraAccount', { name: q, type: 'all', limit: 20 });
    return arr(r.cn).map(galEntry).filter((x) => x.email);
  } catch {
    return galSearch(acc, q, 20);
  }
}

// ---------------------------------------------------------------------------
// 2. Voľno / obsadené

export interface FreeBusy {
  email: string;
  /** b = obsadený, t = predbežne, u = mimo kancelárie, n = nedostupné údaje */
  slots: Array<{ s: number; e: number; kind: 'b' | 't' | 'u' | 'n' }>;
}

export async function freeBusy(acc: AccountConfig, emails: string[], from: number, to: number): Promise<FreeBusy[]> {
  const r = await soap(acc, 'GetFreeBusy', 'urn:zimbraMail', { s: from, e: to, name: emails.join(',') });
  return arr(r.usr).map((u: J) => ({
    email: u.id,
    slots: (['b', 't', 'u', 'n'] as const).flatMap((k) => arr(u[k]).map((x: J) => ({ s: Number(x.s), e: Number(x.e), kind: k }))),
  }));
}

// ---------------------------------------------------------------------------
// 3. Zasadačky a zdroje

export interface Resource {
  name: string;
  email: string;
  kind: 'Location' | 'Equipment';
  capacity?: number;
  location?: string;
  description?: string;
}

export async function resources(acc: AccountConfig): Promise<Resource[]> {
  const r = await soap(acc, 'SearchCalendarResources', 'urn:zimbraAccount', {
    attrs: 'email,fullName,zimbraCalResType,zimbraCalResCapacity,zimbraCalResLocationDisplayName,description',
    searchFilter: { conds: { or: 1, cond: [{ attr: 'zimbraCalResType', op: 'eq', value: 'Location' }, { attr: 'zimbraCalResType', op: 'eq', value: 'Equipment' }] } },
  });
  return arr(r.calresource).map((c: J) => {
    const a = attrs(c._attrs);
    return {
      name: a.fullName || c.name,
      email: a.email || c.name,
      kind: a.zimbraCalResType === 'Equipment' ? 'Equipment' : 'Location',
      capacity: a.zimbraCalResCapacity ? Number(a.zimbraCalResCapacity) : undefined,
      location: a.zimbraCalResLocationDisplayName,
      description: a.description,
    } as Resource;
  });
}

// ---------------------------------------------------------------------------
// 4. Zdieľania od kolegov

export interface Share {
  ownerEmail: string;
  ownerName?: string;
  ownerId: string;
  folderId: string;
  folderPath: string;
  view: string;
  rights: string;
  mounted: boolean;
}

export async function shares(acc: AccountConfig): Promise<Share[]> {
  const r = await soap(acc, 'GetShareInfo', 'urn:zimbraAccount', { includeSelf: 0 });
  return arr(r.share).map((s: J) => ({
    ownerEmail: s.ownerEmail,
    ownerName: s.ownerName,
    ownerId: s.ownerId,
    folderId: String(s.folderId),
    folderPath: s.folderPath,
    view: s.view ?? 'message',
    rights: s.rights ?? 'r',
    mounted: !!s.mid,
  }));
}

export async function mountShare(acc: AccountConfig, s: Share, name: string) {
  await soap(acc, 'CreateMountpoint', 'urn:zimbraMail', { link: { l: '1', name, zid: s.ownerId, rid: s.folderId, view: s.view } });
  return true as const;
}

// ---------------------------------------------------------------------------
// 5. Neprítomnosť na serveri

export interface ServerOoo {
  enabled: boolean;
  reply: string;
  from?: string;
  until?: string;
  external?: string;
  externalEnabled?: boolean;
}

const genTime = (iso?: string, end = false) => {
  if (!iso) return '';
  const d = new Date(`${iso}T${end ? '23:59:59' : '00:00:00'}`);
  // GeneralizedTime: yyyyMMddHHmmssZ (bez „T“)
  return d.toISOString().replace(/[-:T]/g, '').replace(/\.\d{3}/, '');
};
const fromGenTime = (g?: string) => {
  const m = /^(\d{4})(\d{2})(\d{2})T?(\d{2})?(\d{2})?(\d{2})?(Z)?/.exec(g ?? '');
  if (!m) return undefined;
  const d = m[7] ? new Date(Date.UTC(+m[1]!, +m[2]! - 1, +m[3]!, +(m[4] ?? 0), +(m[5] ?? 0), +(m[6] ?? 0))) : new Date(+m[1]!, +m[2]! - 1, +m[3]!);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
};

export async function getOoo(acc: AccountConfig): Promise<ServerOoo> {
  const r = await soap(acc, 'GetPrefs', 'urn:zimbraAccount', {
    pref: ['zimbraPrefOutOfOfficeReplyEnabled', 'zimbraPrefOutOfOfficeReply', 'zimbraPrefOutOfOfficeFromDate', 'zimbraPrefOutOfOfficeUntilDate', 'zimbraPrefOutOfOfficeExternalReply', 'zimbraPrefOutOfOfficeExternalReplyEnabled'].map((n) => ({ name: n })),
  });
  const p = Object.fromEntries(arr(r.pref).map((x: J) => [x.name, x._content])) as Record<string, string>;
  const a = r._attrs ? attrs(r._attrs) : p;
  const g = (k: string) => p[k] ?? a[k];
  return {
    enabled: g('zimbraPrefOutOfOfficeReplyEnabled') === 'TRUE',
    reply: g('zimbraPrefOutOfOfficeReply') ?? '',
    from: fromGenTime(g('zimbraPrefOutOfOfficeFromDate')),
    until: fromGenTime(g('zimbraPrefOutOfOfficeUntilDate')),
    external: g('zimbraPrefOutOfOfficeExternalReply'),
    externalEnabled: g('zimbraPrefOutOfOfficeExternalReplyEnabled') === 'TRUE',
  };
}

export async function setOoo(acc: AccountConfig, o: ServerOoo) {
  const pref = [
    { name: 'zimbraPrefOutOfOfficeReplyEnabled', _content: o.enabled ? 'TRUE' : 'FALSE' },
    { name: 'zimbraPrefOutOfOfficeReply', _content: o.reply },
    { name: 'zimbraPrefOutOfOfficeFromDate', _content: genTime(o.from) },
    { name: 'zimbraPrefOutOfOfficeUntilDate', _content: genTime(o.until, true) },
  ];
  await soap(acc, 'ModifyPrefs', 'urn:zimbraAccount', { pref });
  return true as const;
}

// ---------------------------------------------------------------------------
// 6. Filtre na serveri

export interface ServerRule {
  name: string;
  active: boolean;
  match: 'all' | 'any';
  conds: Array<{ field: 'from' | 'to' | 'subject' | 'any'; contains: string }>;
  actions: Array<{ type: 'move' | 'read' | 'flag' | 'tag' | 'delete' | 'stop'; folder?: string; tag?: string }>;
  /** pravidlo obsahuje podmienky/akcie, ktoré MailKlient nevie upraviť – zachová sa bezo zmeny */
  readOnly?: boolean;
  raw?: J;
}

function ruleFromZimbra(f: J): ServerRule {
  const t = f.filterTests?.[0] ?? f.filterTests ?? {};
  const conds: ServerRule['conds'] = [];
  let unsupported = false;
  for (const [k, v] of Object.entries(t)) {
    if (k === 'condition') continue;
    for (const x of arr(v as J)) {
      if (k === 'headerTest' && x.stringComparison === 'contains' && !x.negative) {
        const h = String(x.header ?? '').toLowerCase();
        const field = h === 'from' ? 'from' : h === 'to' || h === 'to,cc' || h === 'cc' ? 'to' : h === 'subject' ? 'subject' : null;
        if (field) conds.push({ field, contains: x.value });
        else unsupported = true;
      } else if (k === 'addressTest' && x.stringComparison === 'contains' && !x.negative) {
        conds.push({ field: String(x.header).toLowerCase() === 'from' ? 'from' : 'to', contains: x.value });
      } else unsupported = true;
    }
  }
  const a = f.filterActions?.[0] ?? f.filterActions ?? {};
  const actions: ServerRule['actions'] = [];
  for (const [k, v] of Object.entries(a)) {
    for (const x of arr(v as J)) {
      if (k === 'actionFileInto') actions.push(/^\/?trash$/i.test(x.folderPath) ? { type: 'delete' } : { type: 'move', folder: String(x.folderPath).replace(/^\//, '') });
      else if (k === 'actionFlag' && x.flagName === 'read') actions.push({ type: 'read' });
      else if (k === 'actionFlag' && x.flagName === 'flagged') actions.push({ type: 'flag' });
      else if (k === 'actionTag') actions.push({ type: 'tag', tag: x.tagName });
      else if (k === 'actionStop') actions.push({ type: 'stop' });
      else if (k === 'actionKeep') {
        /* predvolené */
      } else unsupported = true;
    }
  }
  return { name: f.name, active: f.active !== false && f.active !== 0, match: t.condition === 'anyof' ? 'any' : 'all', conds, actions, readOnly: unsupported, raw: f };
}

function ruleToZimbra(r: ServerRule): J {
  if (r.readOnly && r.raw) return { ...r.raw, name: r.name, active: r.active };
  const headerTest = r.conds
    .filter((c) => c.contains.trim())
    .map((c, i) => ({ index: i, header: c.field === 'from' ? 'from' : c.field === 'to' ? 'to,cc' : c.field === 'subject' ? 'subject' : 'from,to,cc,subject', stringComparison: 'contains', value: c.contains.trim() }));
  const fa: J = {};
  let i = 0;
  for (const a of r.actions) {
    if (a.type === 'move' && a.folder) (fa.actionFileInto ??= []).push({ index: i++, folderPath: a.folder.startsWith('/') ? a.folder : `/${a.folder}` });
    if (a.type === 'delete') (fa.actionFileInto ??= []).push({ index: i++, folderPath: '/Trash' });
    if (a.type === 'read') (fa.actionFlag ??= []).push({ index: i++, flagName: 'read' });
    if (a.type === 'flag') (fa.actionFlag ??= []).push({ index: i++, flagName: 'flagged' });
    if (a.type === 'tag' && a.tag) (fa.actionTag ??= []).push({ index: i++, tagName: a.tag });
    if (a.type === 'stop') (fa.actionStop ??= []).push({ index: i++ });
  }
  if (!Object.keys(fa).length) fa.actionKeep = [{ index: 0 }];
  return { name: r.name, active: r.active, filterTests: [{ condition: r.match === 'any' ? 'anyof' : 'allof', headerTest }], filterActions: [fa] };
}

export async function getFilters(acc: AccountConfig): Promise<ServerRule[]> {
  const r = await soap(acc, 'GetFilterRules', 'urn:zimbraMail');
  const rules = r.filterRules?.[0]?.filterRule ?? r.filterRules?.filterRule ?? [];
  return arr(rules).map(ruleFromZimbra);
}

export async function setFilters(acc: AccountConfig, rules: ServerRule[]) {
  await soap(acc, 'ModifyFilterRules', 'urn:zimbraMail', { filterRules: [{ filterRule: rules.map(ruleToZimbra) }] });
  return true as const;
}

// ---------------------------------------------------------------------------
// 7. Podpisy a identity

export interface ServerSignature {
  id: string;
  name: string;
  html?: string;
  text?: string;
}

export async function getSignatures(acc: AccountConfig): Promise<ServerSignature[]> {
  const r = await soap(acc, 'GetSignatures', 'urn:zimbraAccount');
  return arr(r.signature).map((s: J) => {
    const c = arr(s.content);
    return { id: s.id, name: s.name, html: c.find((x: J) => /html/i.test(x.type))?._content, text: c.find((x: J) => /plain/i.test(x.type))?._content };
  });
}

export async function saveSignature(acc: AccountConfig, s: { id?: string; name: string; html: string; text: string }) {
  const signature = {
    ...(s.id ? { id: s.id } : {}),
    name: s.name,
    content: [
      { type: 'text/html', _content: s.html },
      { type: 'text/plain', _content: s.text },
    ],
  };
  const r = await soap(acc, s.id ? 'ModifySignature' : 'CreateSignature', 'urn:zimbraAccount', { signature });
  return String(r.signature?.[0]?.id ?? r.signature?.id ?? s.id ?? '');
}

export async function getIdentities(acc: AccountConfig): Promise<Array<{ name: string; email: string; signatureId?: string }>> {
  const r = await soap(acc, 'GetIdentities', 'urn:zimbraAccount');
  return arr(r.identity).map((i: J) => {
    const a = attrs(i._attrs);
    return { name: a.zimbraPrefFromDisplay ?? '', email: a.zimbraPrefFromAddress ?? '', signatureId: a.zimbraPrefDefaultSignatureId };
  }).filter((x) => x.email);
}

// ---------------------------------------------------------------------------
// 8. Značky (tags)

const ZCOLORS = ['#616161', '#0f6cbd', '#0e7490', '#107c10', '#6d28d9', '#c50f1f', '#b45309', '#be123c', '#4b5563', '#d83b01'];

export async function tags(acc: AccountConfig): Promise<Array<{ id: string; name: string; color: string }>> {
  const r = await soap(acc, 'GetTag', 'urn:zimbraMail');
  return arr(r.tag).map((t: J) => ({ id: String(t.id), name: t.name, color: t.rgb ?? ZCOLORS[Number(t.color ?? 0)] ?? ZCOLORS[0]! }));
}

// ---------------------------------------------------------------------------
// 10. Aktovka (Briefcase)

export interface BriefcaseDoc {
  id: string;
  name: string;
  size: number;
  contentType: string;
  modified: string;
  folder: string;
}

export async function briefcase(acc: AccountConfig, q = ''): Promise<BriefcaseDoc[]> {
  const r = await soap(acc, 'Search', 'urn:zimbraMail', { types: 'document', query: q ? `${q}` : 'in:briefcase OR is:anywhere', limit: 200, sortBy: 'dateDesc' });
  return arr(r.doc).map((d: J) => ({ id: String(d.id), name: d.name, size: Number(d.s ?? 0), contentType: d.ct ?? '', modified: new Date(Number(d.md ?? d.d ?? 0) * (Number(d.md ?? 0) < 1e11 ? 1000 : 1)).toISOString(), folder: String(d.l ?? '') }));
}

/** Stiahne dokument z Aktovky do dočasného súboru (na priloženie k správe). */
export async function downloadDoc(acc: AccountConfig, d: BriefcaseDoc) {
  const s = await session(acc);
  const r = await rawRequest(acc, `${s.base}/service/home/~/?id=${encodeURIComponent(d.id)}&auth=co`, undefined, { Cookie: `ZM_AUTH_TOKEN=${s.token}` });
  if (r.status >= 400) throw new UserError(`Súbor sa nepodarilo stiahnuť (chyba ${r.status}).`);
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'mailklient-aktovka-'));
  const file = path.join(dir, d.name.replace(/[\\/:*?"<>|]/g, '_'));
  fs.writeFileSync(file, r.body);
  return { path: file, filename: d.name, size: r.body.length };
}

// ---------------------------------------------------------------------------
// 11. Kvóta, 12. heslo, 13. dôveryhodní / blokovaní

export async function quota(acc: AccountConfig): Promise<{ used: number; limit: number }> {
  const r = await soap(acc, 'GetInfo', 'urn:zimbraAccount', { sections: 'mbox,attrs' });
  const a = attrs(r.attrs?._attrs ?? r.attrs);
  return { used: Number(r.used ?? 0), limit: Number(a.zimbraMailQuota ?? 0) };
}

export async function changePassword(acc: AccountConfig, oldPass: string, newPass: string) {
  await soapRaw(acc, 'ChangePassword', 'urn:zimbraAccount', {
    account: { by: 'name', _content: acc.username.includes('@') ? acc.username : acc.email },
    oldPassword: { _content: oldPass },
    password: { _content: newPass },
  });
  sessions.delete(acc.id);
  return true as const;
}

export async function getWhiteBlack(acc: AccountConfig): Promise<{ white: string[]; black: string[] }> {
  const r = await soap(acc, 'GetWhiteBlackList', 'urn:zimbraAccount');
  const list = (x: J) => arr(x?.[0]?.addr ?? x?.addr).map((a: J) => String(a._content ?? a)).filter(Boolean);
  return { white: list(r.whiteList), black: list(r.blackList) };
}

export async function setWhiteBlack(acc: AccountConfig, white: string[], black: string[]) {
  const cur = await getWhiteBlack(acc);
  const diff = (oldL: string[], newL: string[]) => [
    ...oldL.filter((x) => !newL.includes(x)).map((x) => ({ op: '-', _content: x })),
    ...newL.filter((x) => !oldL.includes(x)).map((x) => ({ op: '+', _content: x })),
  ];
  await soap(acc, 'ModifyWhiteBlackList', 'urn:zimbraAccount', {
    whiteList: [{ addr: diff(cur.white, white) }],
    blackList: [{ addr: diff(cur.black, black) }],
  });
  return true as const;
}

export async function blockSender(acc: AccountConfig, email: string) {
  const cur = await getWhiteBlack(acc);
  if (cur.black.includes(email.toLowerCase())) return true as const;
  return setWhiteBlack(acc, cur.white.filter((x) => x !== email.toLowerCase()), [...cur.black, email.toLowerCase()]);
}
