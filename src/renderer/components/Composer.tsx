// Písanie správy priamo v čítacom paneli (ako Outlook): formátovaný text (TipTap), príjemcovia s dopĺňaním,
// podpis podľa konta, citácia pôvodnej správy v štýle Outlooku, prílohy (aj pretiahnutím), priebežné koncepty.
import { kb } from '../platform';
import { Button, Menu, MenuItem, MenuList, MenuPopover, MenuTrigger, Popover, PopoverSurface, PopoverTrigger, Input, Select, Tooltip } from '@fluentui/react-components';
import { EditorContent, useEditor, type Editor } from '@tiptap/react';
import StarterKit from '@tiptap/starter-kit';
import { Color, FontFamily, FontSize, TextStyle } from '@tiptap/extension-text-style';
import { Highlight } from '@tiptap/extension-highlight';
import { TextAlign } from '@tiptap/extension-text-align';
import { Image } from '@tiptap/extension-image';
import { TableKit } from '@tiptap/extension-table';
import { Placeholder } from '@tiptap/extensions';
import { forwardRef, useCallback, useEffect, useImperativeHandle, useRef, useState } from 'react';
import type { ImapAccountPublic } from '../../core/account';
import type { DraftRef, MessageFull, OutboxItem, PickedFile, SendInput } from '../../core/api';
import { signatureToHtml, type AppSettings } from '../../core/settings';
import { AlignCenterIcon, AlignLeftIcon, AlignRightIcon, AttachIcon, PreviewSmallIcon, BulletsIcon, ImportantIcon, LowPriorityIcon, CloseSmallIcon, DeleteIcon, NewWindowIcon, IndentIcon, NumberingIcon, OutdentIcon } from '../icons';
import { accountColor } from '../theme';
import { RecipientField, formatAddress, splitAddresses, type Contact } from './RecipientField';
import { BriefcasePicker } from './BriefcasePicker';
import { useAttachmentGuard } from './AttachmentGuard';

export type ComposeKind = 'new' | 'reply' | 'replyAll' | 'forward';

export interface ComposeInit {
  kind: ComposeKind;
  accountId?: string;
  /** pôvodná správa pri odpovedi / preposlaní */
  original?: MessageFull;
  to?: string;
  cc?: string;
  subject?: string;
  /** hotový obsah (koncept, vrátené odoslanie) – bez podpisu a citácie */
  html?: string;
  bcc?: string;
  draft?: DraftRef;
  files?: PickedFile[];
  importance?: boolean;
  receipt?: boolean;
  /** adresa odosielateľa (alias) */
  fromEmail?: string;
  security?: { sign?: 'pgp' | 'smime'; encrypt?: 'pgp' | 'smime' };
}

export interface ComposerHandle {
  send(): void;
  discard(): void;
  attach(): void;
  insertImage(): void;
  insertTable(): void;
  insertHr(): void;
  insertDate(): void;
  setSignature(id: string | null): void;
  toggleImportance(): void;
  toggleReceipt(): void;
  sendLater(at: Date): void;
  popOut(): void;
  /** odložiť rozpísanú správu do samostatného okna (prázdnu len zavrieť); true = editor sa zavrel */
  park(): Promise<boolean>;
  /** zavrieť a nechať koncept v Konceptoch */
  close(): void;
  insertTemplate(id: string): void;
  saveAsTemplate(): void;
  setSecurity(s: { sign?: 'pgp' | 'smime'; encrypt?: 'pgp' | 'smime' }): void;
  setFollowUp(days: number): void;
  setPriority(p: 'high' | 'normal' | 'low'): void;
}

export interface ComposerState {
  followUp?: number;
  security?: { sign?: 'pgp' | 'smime'; encrypt?: 'pgp' | 'smime' };
  importance: boolean;
  priority?: 'high' | 'normal' | 'low';
  receipt: boolean;
  savedAt: string | null;
  sending: boolean;
  /** predmet (názov karty rozpísanej správy) */
  subject?: string;
}

const FONTS = ['Aptos', 'Calibri', 'Arial', 'Segoe UI', 'Verdana', 'Georgia', 'Times New Roman', 'Courier New'];
const SIZES = ['8', '9', '10', '11', '12', '14', '16', '18', '20', '24', '28', '36'];
const COLORS = ['#242424', '#c50f1f', '#d83b01', '#b45309', '#3f6212', '#2F6F5E', '#0e7490', '#1d4ed8', '#6d28d9', '#be123c', '#616161', '#ffffff'];
const HIGHLIGHTS = ['#fff3a3', '#c7f0d8', '#cde4ff', '#ffd6e7', '#ffe0b2'];

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const textToHtml = (t: string) =>
  t
    .split(/\n{2,}/)
    .map((p) => `<p>${esc(p).replace(/\n/g, '<br>')}</p>`)
    .join('');

/** Predvolený podpis konta ako HTML. */
function signatureHtml(s: AppSettings, accountId: string | undefined, kind: ComposeKind) {
  if (!accountId) return '';
  const d = s.signatureDefaults[accountId];
  const id = kind === 'new' ? d?.newMsg : d?.reply;
  return signatureToHtml(s.signatures.find((x) => x.id === id));
}

/** prípona a farba štítka prílohy (ako v Outlooku: W, X, P, PDF…) */
const extOf = (n: string) => (/\.([a-z0-9]{1,5})$/i.exec(n)?.[1] ?? '').toLowerCase();
function extLabel(n: string) {
  const e = extOf(n);
  if (['doc', 'docx', 'odt', 'rtf'].includes(e)) return 'W';
  if (['xls', 'xlsx', 'ods', 'csv'].includes(e)) return 'X';
  if (['ppt', 'pptx', 'odp'].includes(e)) return 'P';
  return (e || '?').slice(0, 3).toUpperCase();
}
function extClass(n: string) {
  const e = extOf(n);
  if (['doc', 'docx', 'odt', 'rtf'].includes(e)) return 'word';
  if (['xls', 'xlsx', 'ods', 'csv'].includes(e)) return 'excel';
  if (['ppt', 'pptx', 'odp'].includes(e)) return 'ppt';
  if (e === 'pdf') return 'pdf';
  if (['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'svg', 'heic'].includes(e)) return 'img';
  if (['zip', '7z', 'rar', 'gz'].includes(e)) return 'zip';
  return '';
}
const fmtSize = (b: number) => (b >= 1024 * 1024 ? `${(b / 1024 / 1024).toFixed(1).replace('.', ',')} MB` : `${Math.max(1, Math.ceil(b / 1024))} kB`);

/** Hlavička pôvodnej správy ako v Outlooku (Od / Odoslané / Komu / Kópia / Predmet) + jej obsah. */
function quoteHtml(m: MessageFull) {
  const date = m.date ? new Date(m.date).toLocaleString('sk-SK', { dateStyle: 'full', timeStyle: 'short' }) : '';
  const head = [
    `<strong>Od:</strong> ${esc(m.fromName ? `${m.fromName} <${m.fromAddress}>` : m.fromAddress)}`,
    `<strong>Odoslané:</strong> ${esc(date)}`,
    `<strong>Komu:</strong> ${esc(m.to)}`,
    m.cc ? `<strong>Kópia:</strong> ${esc(m.cc)}` : '',
    `<strong>Predmet:</strong> ${esc(m.subject)}`,
  ]
    .filter(Boolean)
    .join('<br>');
  const body = m.html ? m.html.replace(/^[\s\S]*<body[^>]*>/i, '').replace(/<\/body>[\s\S]*$/i, '') : textToHtml(m.text);
  return `<hr><p>${head}</p><p></p>${body}`;
}

function initialHtml(init: ComposeInit, sig: string) {
  if (init.html !== undefined) return init.html.replace(/^[\s\S]*<body[^>]*>/i, '').replace(/<\/body>[\s\S]*$/i, '');
  const top = `<p></p><p></p>${sig}`;
  return init.original ? `${top}<p></p>${quoteHtml(init.original)}` : top;
}

/** HTML na odoslanie: predvolené písmo, prázdne odseky viditeľné aj v Outlooku. */
function emailHtml(inner: string) {
  const body = inner.replace(/<p><\/p>/g, '<p><br></p>');
  return `<!doctype html><html><head><meta charset="utf-8"><style>p{margin:0}table{border-collapse:collapse}td,th{border:1px solid #c8c8c8;padding:4px 8px}</style></head><body><div style="font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt;color:#242424">${body}</div></body></html>`;
}

export const readAsDataUrl = (f: File) =>
  new Promise<string>((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(String(r.result));
    r.onerror = () => rej(r.error);
    r.readAsDataURL(f);
  });

function Tb(props: { label: string; active?: boolean; onClick: () => void; children: React.ReactNode; style?: React.CSSProperties }) {
  return (
    <Tooltip content={props.label} relationship="label">
      <button
        type="button"
        className={`tb ${props.active ? 'on' : ''}`}
        aria-label={props.label}
        aria-pressed={props.active}
        onMouseDown={(e) => e.preventDefault()}
        onClick={props.onClick}
        style={props.style}
      >
        {props.children}
      </button>
    </Tooltip>
  );
}

export function FormatBar({ editor, onImage, table = true }: { editor: Editor; onImage: () => void; table?: boolean }) {
  const [link, setLink] = useState('');
  const [linkOpen, setLinkOpen] = useState(false);
  const ts = editor.getAttributes('textStyle') as { fontFamily?: string; fontSize?: string };
  const font = (ts.fontFamily ?? '').replace(/["']/g, '').split(',')[0] || 'Aptos';
  const size = (ts.fontSize ?? '').replace(/pt$/, '') || '11';
  const chain = () => editor.chain().focus();
  return (
    <div className="fmt" role="toolbar" aria-label="Formátovanie textu">
      <Select size="small" aria-label="Písmo" value={FONTS.includes(font) ? font : 'Aptos'} onChange={(_, d) => chain().setFontFamily(d.value).run()} className="fmt-font">
        {FONTS.map((f) => <option key={f} value={f}>{f}</option>)}
      </Select>
      <Select size="small" aria-label="Veľkosť písma" value={SIZES.includes(size) ? size : '11'} onChange={(_, d) => chain().setFontSize(`${d.value}pt`).run()} className="fmt-size">
        {SIZES.map((f) => <option key={f} value={f}>{f}</option>)}
      </Select>
      <span className="fmt-sep" />
      <Tb label={kb("Tučné (Ctrl+B)")} active={editor.isActive('bold')} onClick={() => chain().toggleBold().run()} style={{ fontWeight: 700 }}>B</Tb>
      <Tb label={kb("Kurzíva (Ctrl+I)")} active={editor.isActive('italic')} onClick={() => chain().toggleItalic().run()} style={{ fontStyle: 'italic', fontFamily: 'Georgia, serif' }}>I</Tb>
      <Tb label={kb("Podčiarknuté (Ctrl+U)")} active={editor.isActive('underline')} onClick={() => chain().toggleUnderline().run()} style={{ textDecoration: 'underline' }}>U</Tb>
      <Tb label="Prečiarknuté" active={editor.isActive('strike')} onClick={() => chain().toggleStrike().run()} style={{ textDecoration: 'line-through' }}>S</Tb>
      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <button type="button" className="tb" aria-label="Farba písma" onMouseDown={(e) => e.preventDefault()} style={{ fontWeight: 700, color: '#c50f1f', boxShadow: 'inset 0 -3px 0 #c50f1f' }}>A</button>
        </MenuTrigger>
        <MenuPopover>
          <div className="swatches">
            {COLORS.map((c) => (
              <button key={c} className="swatch" style={{ background: c }} aria-label={`Farba ${c}`} onClick={() => chain().setColor(c).run()} />
            ))}
            <button className="swatch-reset" onClick={() => chain().unsetColor().run()}>Automatická</button>
          </div>
        </MenuPopover>
      </Menu>
      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <button type="button" className="tb" aria-label="Zvýraznenie" onMouseDown={(e) => e.preventDefault()} style={{ boxShadow: 'inset 0 -3px 0 #ffd400' }}>ab</button>
        </MenuTrigger>
        <MenuPopover>
          <div className="swatches">
            {HIGHLIGHTS.map((c) => (
              <button key={c} className="swatch" style={{ background: c }} aria-label={`Zvýraznenie ${c}`} onClick={() => chain().toggleHighlight({ color: c }).run()} />
            ))}
            <button className="swatch-reset" onClick={() => chain().unsetHighlight().run()}>Bez zvýraznenia</button>
          </div>
        </MenuPopover>
      </Menu>
      <span className="fmt-sep" />
      <Tb label="Odrážky" active={editor.isActive('bulletList')} onClick={() => chain().toggleBulletList().run()}><BulletsIcon /></Tb>
      <Tb label="Číslovanie" active={editor.isActive('orderedList')} onClick={() => chain().toggleOrderedList().run()}><NumberingIcon /></Tb>
      <Tb label="Zmenšiť odsadenie" onClick={() => chain().liftListItem('listItem').run()}><OutdentIcon /></Tb>
      <Tb label="Zväčšiť odsadenie" onClick={() => chain().sinkListItem('listItem').run()}><IndentIcon /></Tb>
      <Tb label="Zarovnať doľava" active={editor.isActive({ textAlign: 'left' })} onClick={() => chain().setTextAlign('left').run()}><AlignLeftIcon /></Tb>
      <Tb label="Na stred" active={editor.isActive({ textAlign: 'center' })} onClick={() => chain().setTextAlign('center').run()}><AlignCenterIcon /></Tb>
      <Tb label="Zarovnať doprava" active={editor.isActive({ textAlign: 'right' })} onClick={() => chain().setTextAlign('right').run()}><AlignRightIcon /></Tb>
      <span className="fmt-sep" />
      <Popover open={linkOpen} onOpenChange={(_, d) => { setLinkOpen(d.open); if (d.open) setLink(editor.getAttributes('link').href ?? ''); }} trapFocus>
        <PopoverTrigger disableButtonEnhancement>
          <button type="button" className={`tb wide ${editor.isActive('link') ? 'on' : ''}`} onMouseDown={(e) => e.preventDefault()}>Odkaz</button>
        </PopoverTrigger>
        <PopoverSurface className="link-pop">
          <Input
            autoFocus
            placeholder="https://"
            value={link}
            onChange={(_, d) => setLink(d.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                const href = link.trim() ? (/^[a-z]+:/i.test(link.trim()) ? link.trim() : `https://${link.trim()}`) : '';
                if (href) chain().extendMarkRange('link').setLink({ href }).run();
                else chain().extendMarkRange('link').unsetLink().run();
                setLinkOpen(false);
              }
            }}
          />
          <span className="muted small">Enter = vložiť, prázdne = odstrániť odkaz</span>
        </PopoverSurface>
      </Popover>
      <button type="button" className="tb wide" onMouseDown={(e) => e.preventDefault()} onClick={onImage}>Obrázok</button>
      {table && <button type="button" className="tb wide" onMouseDown={(e) => e.preventDefault()} onClick={() => chain().insertTable({ rows: 3, cols: 3, withHeaderRow: false }).run()}>Tabuľka</button>}
      <button type="button" className="tb wide" onMouseDown={(e) => e.preventDefault()} onClick={() => chain().unsetAllMarks().clearNodes().run()}>Vymazať formát</button>
    </div>
  );
}

export const Composer = forwardRef<
  ComposerHandle,
  {
    init: ComposeInit;
    accounts: ImapAccountPublic[];
    settings: AppSettings;
    contacts: Contact[];
    onState: (s: ComposerState) => void;
    onClose: (status?: string) => void;
    /** odoslanie čaká 10 s na späťvzatie / je naplánované */
    onQueued?: (item: OutboxItem) => void;
    onSaveTemplate?: (t: { name: string; subject: string; html: string }) => void;
    /** samostatné okno: po odoslaní/zahodení zavrieť okno */
    standalone?: boolean;
    /** kontá na Zimbre: adresár firmy a Aktovka */
    zimbra?: Record<string, boolean>;
  }
>(function Composer(props, ref) {
  const api = window.mailApi;
  const { guard: attGuard, dialog: attGuardDialog } = useAttachmentGuard();
  const [attPreview, setAttPreview] = useState<{ src: string; name: string } | null>(null);
  const [attErr, setAttErr] = useState('');
  type PrevRes = { ok: true; data: { kind: string; src?: string; filename: string } } | { ok: false; error: string } | undefined;
  const runPreview = async (run: () => Promise<PrevRes>) => {
    setAttErr('');
    const r = await run();
    if (r?.ok && r.data.kind === 'image' && r.data.src) setAttPreview({ src: r.data.src, name: r.data.filename });
    else if (r && !r.ok) setAttErr(r.error);
  };
  /** vlastný súbor si používateľ práve vybral sám → bez varovania (spustiteľné súbory hlavný proces aj tak neotvorí);
      príloha preposlanej cudzej správy → rovnaké varovanie ako pri čítaní */
  const showPreview = (filename: string, run: () => Promise<PrevRes>, onSave?: () => void) =>
    onSave ? attGuard(filename, () => runPreview(run), onSave) : void runPreview(run);
  const { init } = props;
  const [accountId, setAccountId] = useState(init.accountId ?? props.accounts[0]?.id ?? '');
  const [to, setTo] = useState<Contact[]>(() => splitAddresses(init.to ?? ''));
  const [cc, setCc] = useState<Contact[]>(() => splitAddresses(init.cc ?? ''));
  const [bcc, setBcc] = useState<Contact[]>(() => splitAddresses(init.bcc ?? ''));
  const [fromEmail, setFromEmail] = useState(init.fromEmail ?? '');
  const [showCc, setShowCc] = useState(!!init.cc);
  const [showBcc, setShowBcc] = useState(!!init.bcc);
  const [subject, setSubject] = useState(init.subject ?? '');
  const [files, setFiles] = useState<PickedFile[]>(init.files ?? []);
  const [priority, setPriority] = useState<'high' | 'normal' | 'low'>(init.importance ? 'high' : 'normal');
  const importance = priority === 'high';
  const [receipt, setReceipt] = useState(!!init.receipt);
  const [attachWarn, setAttachWarn] = useState(false);
  const [followUp, setFollowUp] = useState(0);
  const [security, setSecurity] = useState<{ sign?: 'pgp' | 'smime'; encrypt?: 'pgp' | 'smime' }>(init.security ?? {});
  const [savedAt, setSavedAt] = useState<string | null>(null);
  const [sending, setSending] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const draft = useRef<DraftRef | null>(init.draft ?? null);
  const dirty = useRef(false);
  const saving = useRef<Promise<unknown>>(Promise.resolve());
  const sig = useRef(init.html !== undefined ? '' : signatureHtml(props.settings, accountId, init.kind));
  const imageInput = useRef<HTMLInputElement>(null);

  const editor = useEditor({
    extensions: [
      StarterKit.configure({ link: { openOnClick: false, autolink: true, defaultProtocol: 'https' } }),
      TextStyle,
      Color,
      FontFamily,
      FontSize,
      Highlight.configure({ multicolor: true }),
      TextAlign.configure({ types: ['heading', 'paragraph'] }),
      Image.configure({ inline: true, allowBase64: true, resize: { enabled: true, alwaysPreserveAspectRatio: true, minWidth: 16, minHeight: 16 } }),
      TableKit.configure({ table: { resizable: false } }),
      Placeholder.configure({ placeholder: 'Napíšte správu…' }),
    ],
    content: initialHtml(init, sig.current),
    shouldRerenderOnTransaction: true,
    autofocus: init.kind === 'new' ? false : 'start',
    editorProps: {
      attributes: { class: 'composer-body', spellcheck: 'true' },
      handlePaste: (_view, event) => {
        const imgs = [...(event.clipboardData?.files ?? [])].filter((f) => f.type.startsWith('image/'));
        if (!imgs.length) return false;
        void Promise.all(imgs.map(readAsDataUrl)).then((srcs) => srcs.forEach((src) => editorRef.current?.chain().focus().setImage({ src }).run()));
        return true;
      },
      handleDrop: (_view, event) => {
        const fs = [...((event as DragEvent).dataTransfer?.files ?? [])];
        if (!fs.length) return false;
        addDropped(fs);
        return true;
      },
    },
    onUpdate: () => {
      dirty.current = true;
    },
  });
  const editorRef = useRef<Editor | null>(null);
  editorRef.current = editor;


  function addDropped(fs: File[]) {
    const imgs = fs.filter((f) => f.type.startsWith('image/') && f.size < 3_000_000);
    const others = fs.filter((f) => !imgs.includes(f));
    void Promise.all(imgs.map(readAsDataUrl)).then((srcs) => srcs.forEach((src) => editorRef.current?.chain().focus().setImage({ src }).run()));
    const picked = others
      .map((f) => ({ path: api?.pathForFile(f) ?? '', filename: f.name, size: f.size }))
      .filter((p) => p.path);
    if (picked.length) setFiles((x) => [...x, ...picked]);
  }

  const buildInput = useCallback((): SendInput | null => {
    const ed = editorRef.current;
    if (!ed) return null;
    return {
      to: to.map(formatAddress).join(', '),
      cc: cc.map(formatAddress).join(', '),
      bcc: bcc.map(formatAddress).join(', '),
      subject,
      text: ed.getText({ blockSeparator: '\n' }),
      html: emailHtml(ed.getHTML()),
      replyMessageId: init.original && init.kind !== 'forward' ? init.original.messageId : undefined,
      attachments: files.map((f) => ({ path: f.path, filename: f.filename })),
      priority,
      readReceipt: receipt,
      draft: draft.current ?? undefined,
      forward: init.kind === 'forward' && init.original ? { folder: init.original.folder, uid: init.original.uid } : undefined,
      from: fromIdentity(),
      security: security.sign || security.encrypt ? security : undefined,
      followUpDays: followUp || undefined,
    };
  }, [to, cc, bcc, subject, files, priority, receipt, init, fromEmail, accountId, security, followUp]); // eslint-disable-line react-hooks/exhaustive-deps

  // priebežné ukladanie konceptu (každých 5 s, ak sa niečo zmenilo)
  useEffect(() => {
    dirty.current = true;
  }, [to, cc, bcc, subject, files, fromEmail]);
  useEffect(() => {
    if (!api || !accountId || accountId === 'demo') return;
    const t = window.setInterval(() => {
      if (!dirty.current || sending) return;
      const input = buildInput();
      if (!input || (!input.text.trim() && !input.to && !input.subject)) return;
      dirty.current = false;
      saving.current = saving.current.then(async () => {
        const r = await api.saveDraft(accountId, { ...input, draft: undefined }, draft.current ?? undefined);
        if (r.ok && r.data) {
          draft.current = r.data;
          setSavedAt(new Date().toLocaleTimeString('sk-SK', { hour: '2-digit', minute: '2-digit' }));
        }
      });
    }, 5000);
    return () => window.clearInterval(t);
  }, [api, accountId, buildInput, sending]);

  useEffect(() => {
    props.onState({ importance, priority, receipt, savedAt, sending, security, followUp, subject });
  }, [priority, receipt, savedAt, sending, security, followUp, subject]); // eslint-disable-line react-hooks/exhaustive-deps

  /** Zmena odosielateľa: vymeniť podpis za podpis nového konta. */
  function changeAccount(id: string) {
    const next = signatureHtml(props.settings, id, init.kind);
    const ed = editorRef.current;
    if (ed && sig.current && ed.getHTML().includes(sig.current)) ed.commands.setContent(ed.getHTML().replace(sig.current, next), { emitUpdate: true });
    else if (ed && next) ed.chain().focus('start').insertContent(next).run();
    sig.current = next;
    setAccountId(id);
  }

  const account = props.accounts.find((a) => a.id === accountId);
  function fromIdentity() {
    const acc = props.accounts.find((a) => a.id === accountId);
    const id = acc?.identities?.find((x) => x.email === fromEmail);
    return id ? { name: id.name, email: id.email } : undefined;
  }

  const zimbraOn = !!props.zimbra?.[accountId];
  const lookup = useCallback(
    async (q: string): Promise<Contact[]> => {
      if (!api || !zimbraOn) return [];
      const r = await api.zimbraGalAuto(accountId, q);
      return r.ok ? r.data.map((g) => ({ name: g.name, email: g.email })) : [];
    },
    [api, accountId, zimbraOn],
  );
  const [briefcase, setBriefcase] = useState(false);

  function rememberContacts() {
    try {
      // príjemcov si zapamätať na dopĺňanie
      const known = JSON.parse(localStorage.getItem('contacts') ?? '[]') as Contact[];
      const all = [...[...to, ...cc, ...bcc], ...known].filter((c, i, arr) => arr.findIndex((x) => x.email.toLowerCase() === c.email.toLowerCase()) === i);
      localStorage.setItem('contacts', JSON.stringify(all.slice(0, 500)));
    } catch {
      /* nevadí */
    }
  }

  /** Text sľubuje prílohu, ale žiadna nie je (ako Outlook/Thunderbird). */
  function missingAttachment(input: SendInput) {
    if (files.length || (init.kind === 'forward' && (init.original?.attachments.filter((a) => !a.inline).length ?? 0) > 0)) return false;
    // len vlastný text, nie citovaná pôvodná správa
    const own = (input.text.split(/\n-{2,}|\nOd: |\nFrom: /)[0] ?? '') + ' ' + input.subject;
    return /(v\s+prílohe|v\s+prilohe|prikladám|prikladam|priložen|prilozen|príloh[ua]|prilohu|attach)/i.test(own);
  }

  async function send(at?: Date, force = false) {
    if (!api) return;
    const input = buildInput();
    if (!input) return;
    if (!to.length && !cc.length && !bcc.length) return setError('Zadaj aspoň jedného príjemcu.');
    if (!accountId || accountId === 'demo') return setError('V ukážkovom režime sa nedá odoslať. Pridaj konto.');
    if (!force && missingAttachment(input)) {
      setAttachWarn(true);
      return;
    }
    setAttachWarn(false);
    setError(null);
    setSending(true);
    await saving.current;
    const final = { ...input, draft: draft.current ?? undefined };
    // odoslanie ide cez frontu: 10 s na späťvzatie, alebo v naplánovaný čas
    const r = await api.outboxQueue(accountId, final, (at ?? new Date(Date.now() + 10_000)).toISOString(), at ? 'later' : 'undo');
    setSending(false);
    if (!r.ok) return setError(`${r.error}${r.hint ? ` ${r.hint}` : ''}`);
    rememberContacts();
    props.onQueued?.(r.data);
    props.onClose(
      at
        ? `Správa sa odošle ${at.toLocaleString('sk-SK', { weekday: 'short', day: 'numeric', month: 'numeric', hour: '2-digit', minute: '2-digit' })}.`
        : undefined,
    );
  }

  async function popOut(): Promise<boolean> {
    if (!api || !accountId || accountId === 'demo') return false;
    const input = buildInput();
    if (!input) return false;
    await saving.current;
    const r = await api.saveDraft(accountId, { ...input, draft: undefined }, draft.current ?? undefined);
    if (!r.ok || !r.data) {
      setError(r.ok ? 'Koncept sa nepodarilo uložiť (chýba priečinok Koncepty).' : r.error);
      return false;
    }
    draft.current = r.data;
    await api.composeWindow(accountId, r.data.folder, r.data.uid);
    props.onClose('Rozpísaná správa sa presunula do samostatného okna.');
    return true;
  }

  function insertTemplate(id: string) {
    const t = props.settings.templates?.find((x) => x.id === id);
    const ed = editorRef.current;
    if (!t || !ed) return;
    ed.chain().focus().insertContent(t.html).run();
    if (t.subject && !subject.trim()) setSubject(t.subject);
  }

  const isEmpty = () => {
    const ed = editorRef.current;
    const body = (ed?.getText() ?? '').replace(sig.current.replace(/<[^>]+>/g, ''), '').trim();
    return !to.length && !cc.length && !bcc.length && !subject.trim() && !files.length && (init.kind === 'new' ? !body : true) && !draft.current;
  };

  /** X: zavrieť, rozpísaný text ostane ako koncept v priečinku Koncepty. */
  async function closeKeep() {
    if (!api || !accountId || accountId === 'demo' || isEmpty()) return props.onClose();
    const input = buildInput();
    await saving.current;
    if (input) {
      const r = await api.saveDraft(accountId, { ...input, draft: undefined }, draft.current ?? undefined);
      if (!r.ok) return setError(`Koncept sa nepodarilo uložiť: ${r.error}`);
    }
    props.onClose('Rozpísaná správa je uložená v priečinku Koncepty.');
  }

  async function park(): Promise<boolean> {
    if (isEmpty() || !accountId || accountId === 'demo') {
      props.onClose();
      return true;
    }
    return popOut();
  }

  async function discard() {
    await saving.current;
    if (draft.current && api) void api.deleteDraft(accountId, draft.current);
    props.onClose('Koncept zahodený.');
  }

  async function attach() {
    const r = await api?.pickFiles();
    if (r?.ok) setFiles((f) => [...f, ...r.data]);
  }

  useImperativeHandle(ref, () => ({
    send: () => void send(),
    discard: () => void discard(),
    attach: () => void attach(),
    insertImage: () => imageInput.current?.click(),
    insertTable: () => editorRef.current?.chain().focus().insertTable({ rows: 3, cols: 3, withHeaderRow: false }).run(),
    insertHr: () => editorRef.current?.chain().focus().setHorizontalRule().run(),
    insertDate: () => editorRef.current?.chain().focus().insertContent(esc(new Date().toLocaleString('sk-SK'))).run(),
    setSignature: (id) => {
      const s = props.settings.signatures.find((x) => x.id === id);
      const next = signatureToHtml(s);
      const ed = editorRef.current;
      if (!ed) return;
      if (sig.current && ed.getHTML().includes(sig.current)) ed.commands.setContent(ed.getHTML().replace(sig.current, next), { emitUpdate: true });
      else if (next) ed.chain().focus().insertContent(next).run();
      sig.current = next;
    },
    sendLater: (at) => void send(at),
    setSecurity: (x) => setSecurity(x),
    setFollowUp: (d) => setFollowUp(d),
    popOut: () => void popOut(),
    park: () => park(),
    close: () => void closeKeep(),
    insertTemplate: (id) => insertTemplate(id),
    saveAsTemplate: () => {
      const ed = editorRef.current;
      if (ed) props.onSaveTemplate?.({ name: subject.trim() || 'Nová šablóna', subject, html: ed.getHTML() });
    },
    toggleImportance: () => setPriority((v) => (v === 'high' ? 'normal' : 'high')),
    setPriority,
    toggleReceipt: () => setReceipt((v) => !v),
  }));

  const title = init.kind === 'new' ? 'Nová správa' : init.kind === 'forward' ? 'Preposlať' : init.kind === 'replyAll' ? 'Odpovedať všetkým' : 'Odpovedať';

  return (
    <section
      className={`composer ${dragOver ? 'drag' : ''}`}
      onKeyDown={(e) => {
        if (e.key === 'Escape' && !e.defaultPrevented && !document.querySelector('.fui-MenuPopover, .fui-DialogSurface')) {
          e.preventDefault();
          void closeKeep();
        }
      }}
      onDragOver={(e) => {
        if (e.dataTransfer.types.includes('Files')) {
          e.preventDefault();
          setDragOver(true);
        }
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        setDragOver(false);
        if (!e.dataTransfer.files.length) return;
        e.preventDefault();
        addDropped([...e.dataTransfer.files]);
      }}
    >
      {init.original && <h2 className="composer-subject">{init.original.subject}</h2>}
      <div className="composer-card">
        <div className="composer-top">
          <Button appearance="primary" disabled={sending} onClick={() => void send()} style={{ background: accountColor(account?.color) }}>
            {sending ? 'Odosielam…' : 'Odoslať'}
          </Button>

          <span className="composer-title">{title}</span>
          <span className="grow" />
          <Menu>
            <MenuTrigger disableButtonEnhancement>
              <Button
                appearance="subtle"
                size="small"
                className={`prio-btn prio-${priority}`}
                icon={priority === 'low' ? <LowPriorityIcon /> : <ImportantIcon />}
                title="Dôležitosť správy"
              >
                {priority === 'high' ? 'Vysoká dôležitosť' : priority === 'low' ? 'Nízka dôležitosť' : 'Dôležitosť'}
              </Button>
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                <MenuItem icon={<ImportantIcon style={{ color: '#c50f1f' }} />} onClick={() => setPriority('high')}>{priority === 'high' ? '✓ ' : ''}Vysoká dôležitosť</MenuItem>
                <MenuItem onClick={() => setPriority('normal')}>{priority === 'normal' ? '✓ ' : ''}Normálna</MenuItem>
                <MenuItem icon={<LowPriorityIcon />} onClick={() => setPriority('low')}>{priority === 'low' ? '✓ ' : ''}Nízka dôležitosť</MenuItem>
              </MenuList>
            </MenuPopover>
          </Menu>
          {receipt && <span className="pill">Potvrdenie o prečítaní</span>}
          {security.encrypt && <span className="pill sec">🔒 Šifrovať ({security.encrypt === 'pgp' ? 'PGP' : 'S/MIME'})</span>}
          {security.sign && <span className="pill sec">✓ Podpísať ({security.sign === 'pgp' ? 'PGP' : 'S/MIME'})</span>}
          <Button appearance="subtle" icon={<DeleteIcon />} onClick={() => void discard()} title="Zahodiť (zmaže aj koncept)" aria-label="Zahodiť" />
          {!props.standalone && (
            <Button appearance="subtle" icon={<NewWindowIcon />} onClick={() => void popOut()} title="Otvoriť v samostatnom okne" aria-label="Otvoriť v samostatnom okne" />
          )}
          <button className="composer-x" onClick={() => void closeKeep()} title="Zavrieť (rozpísaná správa ostane v Konceptoch) – Esc" aria-label="Zavrieť správu">
            ×
          </button>
          <span className="grow" />
          <span className="muted small">{savedAt ? `Koncept uložený ${savedAt}` : ''}</span>
        </div>
        <div className="composer-head">
          <div className="rcpt-row">
            <span className="rcpt-label">Od</span>
            <Select
              size="small"
              value={`${accountId}|${fromEmail}`}
              onChange={(_, d) => {
                const [id, email] = d.value.split('|');
                if (id && id !== accountId) changeAccount(id);
                setFromEmail(email ?? '');
              }}
              className="from-select"
              aria-label="Od"
            >
              {props.accounts.flatMap((a) => [
                <option key={a.id} value={`${a.id}|`}>
                  {a.name} &lt;{a.email}&gt;
                </option>,
                ...(a.identities ?? []).map((x) => (
                  <option key={`${a.id}|${x.email}`} value={`${a.id}|${x.email}`}>
                    {x.name || a.name} &lt;{x.email}&gt;
                  </option>
                )),
              ])}
            </Select>
          </div>
          <RecipientField
            label="Komu"
            value={to}
            onChange={setTo}
            contacts={props.contacts}
            lookup={zimbraOn ? lookup : undefined}
            autoFocus={init.kind === 'new' || init.kind === 'forward'}
            extra={
              <span className="rcpt-extra">
                {!showCc && <button className="linkbtn" onClick={() => setShowCc(true)}>Kópia</button>}
                {!showBcc && <button className="linkbtn" onClick={() => setShowBcc(true)}>Skrytá</button>}
              </span>
            }
          />
          {showCc && <RecipientField label="Kópia" value={cc} onChange={setCc} contacts={props.contacts} lookup={zimbraOn ? lookup : undefined} />}
          {showBcc && <RecipientField label="Skrytá" value={bcc} onChange={setBcc} contacts={props.contacts} lookup={zimbraOn ? lookup : undefined} />}
          <div className="rcpt-row">
            <span className="rcpt-label">Predmet</span>
            <input className="subject-input" aria-label="Predmet" placeholder="Pridajte predmet" value={subject} onChange={(e) => setSubject(e.target.value)} />
          </div>
        </div>
        {(files.length > 0 || (init.kind === 'forward' && (init.original?.attachments.filter((a) => !a.inline).length ?? 0) > 0)) && (
          <div className={`composer-files ${files.length + (init.kind === 'forward' ? init.original?.attachments.filter((a) => !a.inline).length ?? 0 : 0) > 4 ? 'many' : ''}`} aria-label="Prílohy">
            {init.kind === 'forward' &&
              init.original?.attachments.filter((a) => !a.inline).map((a, i) => (
                <span key={`fw-${i}`} className="att-card" title={`${a.filename} – príloha preposlanej správy`}>
                  <span className={`att-ext ${extClass(a.filename)}`}>{extLabel(a.filename)}</span>
                  <span className="att-meta"><span className="att-name">{a.filename}</span><span className="att-size">{fmtSize(a.size)}</span></span>
                  <button className="icon-x att-prev" aria-label={`Náhľad ${a.filename}`} title="Náhľad" onClick={() => { const o = init.original!; showPreview(a.filename, async () => api?.previewAttachment(o.accountId, o.folder, o.uid, a.index ?? i), () => void api?.saveAttachment(o.accountId, o.folder, o.uid, a.index ?? i)); }}>
                    <PreviewSmallIcon />
                  </button>
                </span>
              ))}
            {files.map((f, i) => (
              <span key={`${f.path}-${i}`} className="att-card" title={f.filename}>
                <span className={`att-ext ${extClass(f.filename)}`}>{extLabel(f.filename)}</span>
                <span className="att-meta"><span className="att-name">{f.filename}</span><span className="att-size">{fmtSize(f.size)}</span></span>
                <button className="icon-x att-prev" aria-label={`Náhľad ${f.filename}`} title="Náhľad" onClick={() => showPreview(f.filename, async () => api?.previewFile(f.path, f.filename))}>
                  <PreviewSmallIcon />
                </button>
                <button className="icon-x" aria-label={`Odobrať ${f.filename}`} title="Odobrať prílohu" onClick={() => setFiles((x) => x.filter((_, j) => j !== i))}>
                  <CloseSmallIcon />
                </button>
              </span>
            ))}
          </div>
        )}
        {attErr && <div className="att-err" role="alert">{attErr}</div>}
        {attPreview && (
          <div className="att-preview" onClick={() => setAttPreview(null)} role="dialog" aria-label={`Náhľad ${attPreview.name}`}>
            <div className="att-preview-bar">{attPreview.name} <span className="muted">– kliknutím zavrieť</span></div>
            <img src={attPreview.src} alt={attPreview.name} />
          </div>
        )}
        {attGuardDialog}
        {editor && <FormatBar editor={editor} onImage={() => imageInput.current?.click()} />}
        <div className="composer-editor">
          <EditorContent editor={editor} />
        </div>
        <div className="composer-foot">
          <Button size="small" icon={<AttachIcon />} onClick={() => void attach()}>Priložiť súbor</Button>
          {zimbraOn && <Button size="small" onClick={() => setBriefcase(true)}>Z Aktovky</Button>}
          <span className="muted small">Súbory aj obrázky môžeš pretiahnuť sem alebo vložiť cez {kb("Ctrl+V")}.</span>
        </div>
        {briefcase && (
          <BriefcasePicker
            accountId={accountId}
            onPick={(f) => setFiles((x) => [...x, f])}
            onClose={() => setBriefcase(false)}
          />
        )}
        {attachWarn && (
          <div className="attach-warn">
            V texte spomínate prílohu, ale žiadna nie je priložená.
            <Button size="small" onClick={() => void attach()}>Priložiť súbor</Button>
            <Button size="small" appearance="subtle" onClick={() => void send(undefined, true)}>Odoslať bez prílohy</Button>
          </div>
        )}
        {error && <div className="pane-error pad">⚠ {error}</div>}
      </div>
      <input
        ref={imageInput}
        type="file"
        accept="image/*"
        hidden
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) void readAsDataUrl(f).then((src) => editorRef.current?.chain().focus().setImage({ src }).run());
          e.target.value = '';
        }}
      />
    </section>
  );
});
