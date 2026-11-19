// Horný pruh (vo farbe konta), ribbon pre IMAP kontá a pás modulov podľa typu konta.
import { kb } from '../platform';
import {
  Button, Menu, MenuDivider, MenuGroupHeader, MenuItem, MenuItemCheckbox, MenuItemRadio, MenuList, MenuPopover, MenuTrigger, SearchBox, Select, Tab, TabList, Toolbar, ToolbarButton,
  ToolbarDivider, Tooltip,
} from '@fluentui/react-components';
import { useState, type ReactNode } from 'react';
import type { FolderInfo } from '../../core/api';
import {
  ArchiveIcon, AttachIcon, CalendarIcon, SignatureIcon, CloudIcon, ComposeIcon, ContactsIcon, DeleteIcon, ExcelIcon, FlagIcon, ForwardIcon, MailIcon,
  MarkReadIcon, MarkUnreadIcon, MoveIcon, PanelIcon, PowerPointIcon, ReplyAllIcon, ReplyIcon, SettingsIcon, SyncIcon,
  SendReceiveIcon, TasksIcon, TeamsIcon, WordIcon, TagIcon, LoopIcon, TemplateIcon, FileIcon, LockIcon, ServerIcon, CopilotIcon, PlannerIcon, OneNoteIcon, SharePointIcon, FormsIcon, PinIcon, SnoozeIcon, QuickStepIcon, RulesIcon, ConversationIcon, SearchSmallIcon,
  CalendarSmallIcon, TasksSmallIcon,
} from '../icons';
import { BUILTIN_THEMES, onColor, type ThemeDef, type ThemeKey } from '../theme';
import { SEARCH_FIELDS, type Category, type QuickStep, type SearchField, type Signature, type Template } from '../../core/settings';
import type { OrganizeTab } from './OrganizeDialog';
import type { ComposerHandle, ComposerState } from './Composer';

export function TopBar(props: {
  title: string;
  subtitle: string;
  color: string;
  search?: string;
  onSearch?: (v: string) => void;
  rightOpen: boolean;
  onToggleRight: () => void;
  /** typ konta: ZIMBRA / IMAP / OUTLOOK WEB */
  kind?: string;
  searchOptions?: {
    field: SearchField;
    onField: (f: SearchField) => void;
    serverAuto: boolean;
    onServerAuto: (v: boolean) => void;
    history: string[];
    onHistory: (q: string) => void;
    onClearHistory: () => void;
    onSaveFolder?: () => void;
  };
}) {
  const so = props.searchOptions;
  return (
    <header className="topbar" style={{ background: props.color, color: onColor(props.color) }}>
      <div className="topbar-title">
        <div className="topbar-name">
          {props.title}
          {props.kind && <span className="kind-badge" title="Typ konta">{props.kind}</span>}
        </div>
        <div className="topbar-sub">{props.subtitle}</div>
      </div>
      {props.onSearch ? (
        <div className="topbar-searchwrap">
          <SearchBox
            className="topbar-search"
            placeholder={so && so.field !== 'all' ? `Hľadať: ${SEARCH_FIELDS.find((x) => x.id === so.field)?.label.toLowerCase()}` : 'Hľadať v pošte (predmet, odosielateľ, text)'}
            value={props.search ?? ''}
            onChange={(_, d) => props.onSearch!(d.value)}
            appearance="filled-lighter"
          />
          {so && (
            <Menu
              positioning="below-end"
              checkedValues={{ field: [so.field], server: so.serverAuto ? ['on'] : [] }}
              onCheckedValueChange={(_, d) => {
                if (d.name === 'field' && d.checkedItems[0]) so.onField(d.checkedItems[0] as SearchField);
                if (d.name === 'server') so.onServerAuto(d.checkedItems.includes('on'));
              }}
            >
              <MenuTrigger disableButtonEnhancement>
                <button className="topbar-search-opt" aria-label="Možnosti hľadania" title="Možnosti hľadania">▾</button>
              </MenuTrigger>
              <MenuPopover style={{ minWidth: 380, maxWidth: 'none' }}>
                <MenuList>
                  <Menu>
                    <MenuTrigger disableButtonEnhancement>
                      <MenuItem disabled={!so.history.length}>Nedávne hľadania</MenuItem>
                    </MenuTrigger>
                    <MenuPopover>
                      <MenuList>
                        {so.history.map((q) => (
                          <MenuItem key={q} onClick={() => so.onHistory(q)}>{q}</MenuItem>
                        ))}
                        <MenuDivider />
                        <MenuItem onClick={so.onClearHistory}>Vymazať históriu</MenuItem>
                      </MenuList>
                    </MenuPopover>
                  </Menu>
                  <MenuDivider />
                  {SEARCH_FIELDS.map((f) => (
                    <MenuItemRadio key={f.id} name="field" value={f.id}>{f.label}</MenuItemRadio>
                  ))}
                  <MenuDivider />
                  <MenuItemCheckbox name="server" value="on">Hľadať aj na serveri (celý priečinok)</MenuItemCheckbox>
                  <MenuDivider />
                  <MenuItem disabled={!so.onSaveFolder} onClick={() => so.onSaveFolder?.()}>Vytvoriť priečinok hľadania…</MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
          )}
        </div>
      ) : (
        <div />
      )}
      <div className="topbar-tools">
        <Tooltip content={props.rightOpen ? 'Skryť pravý panel' : 'Pravý panel: weby a programy'} relationship="label">
          <button className={`topbar-btn ${props.rightOpen ? 'on' : ''}`} onClick={props.onToggleRight} aria-pressed={props.rightOpen}>
            <PanelIcon /> <span>Pravý panel</span>
          </button>
        </Tooltip>
      </div>
    </header>
  );
}

/** Odoslať neskôr: zajtra ráno, pondelok ráno, vlastný čas. */
function SendLaterMenu(props: { onPick: (d: Date) => void }) {
  const [custom, setCustom] = useState('');
  const at = (days: number, hour: number) => {
    const d = new Date();
    d.setDate(d.getDate() + days);
    d.setHours(hour, 0, 0, 0);
    return d;
  };
  const nextMonday = () => {
    const d = at(1, 8);
    while (d.getDay() !== 1) d.setDate(d.getDate() + 1);
    return d;
  };
  return (
    <Menu>
      <MenuTrigger disableButtonEnhancement>
        <ToolbarButton icon={<SnoozeIcon />}>Odoslať neskôr</ToolbarButton>
      </MenuTrigger>
      <MenuPopover>
        <MenuList>
          <MenuItem onClick={() => props.onPick(at(1, 8))}>Zajtra o 8:00</MenuItem>
          <MenuItem onClick={() => props.onPick(at(1, 13))}>Zajtra o 13:00</MenuItem>
          <MenuItem onClick={() => props.onPick(nextMonday())}>Budúci pondelok o 8:00</MenuItem>
          <div className="snooze-custom" onClick={(e) => e.stopPropagation()} onKeyDown={(e) => e.stopPropagation()}>
            <input type="datetime-local" value={custom} onChange={(e) => setCustom(e.target.value)} aria-label="Dátum a čas odoslania" />
            <Button size="small" disabled={!custom} onClick={() => custom && props.onPick(new Date(custom))}>Naplánovať</Button>
          </div>
          <div className="muted small pad">Universum eMail musí byť v tom čase spustený.</div>
        </MenuList>
      </MenuPopover>
    </Menu>
  );
}

/** Úloha zo správy s termínom (ako vlajky s termínom v Outlooku). */
function TaskMenu(props: { disabled: boolean; onTask: (due?: string) => void }) {
  const iso = (d: Date) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  const plus = (n: number) => iso(new Date(Date.now() + n * 86400000));
  const weekday = (target: number, next: boolean) => {
    const d = new Date();
    let add = (target - d.getDay() + 7) % 7;
    if (next) add = add === 0 ? 7 : add + (add < 7 && target === 1 ? 0 : 0);
    if (next && add <= 0) add += 7;
    return iso(new Date(d.getFullYear(), d.getMonth(), d.getDate() + add));
  };
  return (
    <Menu>
      <MenuTrigger disableButtonEnhancement>
        <ToolbarButton icon={<TasksIcon />} disabled={props.disabled}>Úloha</ToolbarButton>
      </MenuTrigger>
      <MenuPopover>
        <MenuList>
          <MenuItem onClick={() => props.onTask(plus(0))}>Dnes</MenuItem>
          <MenuItem onClick={() => props.onTask(plus(1))}>Zajtra</MenuItem>
          <MenuItem onClick={() => props.onTask(weekday(5, false))}>Tento týždeň (piatok)</MenuItem>
          <MenuItem onClick={() => props.onTask(weekday(1, true))}>Budúci týždeň (pondelok)</MenuItem>
          <MenuItem onClick={() => props.onTask(undefined)}>Bez termínu</MenuItem>
        </MenuList>
      </MenuPopover>
    </Menu>
  );
}

/** Odložiť: rýchle voľby ako v Outlooku + vlastný dátum a čas. */
function SnoozeMenu(props: { disabled: boolean; onSnooze: (d: Date) => void }) {
  const [custom, setCustom] = useState('');
  const at = (days: number, hour: number) => {
    const d = new Date();
    d.setDate(d.getDate() + days);
    d.setHours(hour, 0, 0, 0);
    return d;
  };
  const laterToday = () => new Date(Date.now() + 3 * 3600_000);
  const nextMonday = () => {
    const d = at(1, 8);
    while (d.getDay() !== 1) d.setDate(d.getDate() + 1);
    return d;
  };
  return (
    <Menu>
      <MenuTrigger disableButtonEnhancement>
        <ToolbarButton icon={<SnoozeIcon />} disabled={props.disabled}>Odložiť</ToolbarButton>
      </MenuTrigger>
      <MenuPopover>
        <MenuList>
          <MenuItem onClick={() => props.onSnooze(laterToday())}>O 3 hodiny</MenuItem>
          <MenuItem onClick={() => props.onSnooze(at(1, 8))}>Zajtra o 8:00</MenuItem>
          <MenuItem onClick={() => props.onSnooze(nextMonday())}>Budúci pondelok o 8:00</MenuItem>
          <div className="snooze-custom" onClick={(e) => e.stopPropagation()} onKeyDown={(e) => e.stopPropagation()}>
            <input type="datetime-local" value={custom} onChange={(e) => setCustom(e.target.value)} aria-label="Vlastný dátum a čas" />
            <Button size="small" disabled={!custom} onClick={() => custom && props.onSnooze(new Date(custom))}>Odložiť</Button>
          </div>
        </MenuList>
      </MenuPopover>
    </Menu>
  );
}

export type RibbonTab = 'home' | 'message' | 'insert' | 'view';

export interface RibbonComposer {
  state: ComposerState;
  signatures: Signature[];
  templates: Template[];
  onManageTemplates: () => void;
  onManageCrypto: () => void;
  run: (fn: (h: ComposerHandle) => void) => void;
}

export function Ribbon(props: {
  tab: RibbonTab;
  onTab: (t: RibbonTab) => void;
  hasMessage: boolean;
  seen: boolean;
  flagged: boolean;
  canMove: boolean;
  folders: FolderInfo[];
  onCompose: () => void;
  onReply: () => void;
  onReplyAll: () => void;
  onForward: () => void;
  onDelete: () => void;
  onArchive: () => void;
  onMove: (dest: string) => void;
  onToggleSeen: () => void;
  onToggleFlag: () => void;
  categories: Category[];
  keywords: string[];
  onCategory: (id: string, on: boolean) => void;
  pinned: boolean;
  onTogglePin: () => void;
  onTask: (due?: string) => void;
  onIo: (k: 'eml' | 'mbox' | 'import' | 'thunderbird' | 'update' | 'print' | 'pdf' | 'backup' | 'restore' | 'shortcuts') => void;
  onZimbra?: (tab?: 'ooo' | 'filters' | 'shares' | 'spam' | 'sync' | 'password') => void;
  onSnooze: (until: Date) => void;
  quickSteps: QuickStep[];
  onQuickStep: (q: QuickStep) => void;
  onOrganize: (tab: OrganizeTab) => void;
  conversations: boolean;
  onConversations: (v: boolean) => void;
  density: 'comfortable' | 'compact';
  onDensity: (d: 'comfortable' | 'compact') => void;
  previewLine: boolean;
  onPreviewLine: (v: boolean) => void;
  runInBackground: boolean;
  onRunInBackground: (v: boolean) => void;
  openMode: 'window' | 'tab';
  onOpenMode: (v: 'window' | 'tab') => void;
  onSendReceive: () => void;
  /** náhľad kalendára / úloh vedľa čítacieho panela (len pošta IMAP/Zimbra) */
  myDay?: { calendar: boolean; tasks: boolean };
  onMyDay?: (v: { calendar: boolean; tasks: boolean }) => void;
  onRefresh: () => void;
  onEditAccount?: () => void;
  onSignatures: () => void;
  theme: ThemeKey;
  onTheme: (t: ThemeKey) => void;
  customThemes: ThemeDef[];
  onEditThemes: () => void;
  rightOpen: boolean;
  onToggleRight: () => void;
  compact: boolean;
  onCompact: (v: boolean) => void;
  gpuOff: boolean;
  onToggleGpu: () => void;
  onOpenLog: () => void;
  composer?: RibbonComposer;
}) {
  const p = props;
  const c = p.composer;
  const sz = p.compact ? 'small' : 'medium';
  return (
    <div className={`ribbon ${p.compact ? 'ribbon-compact' : ''}`}>
      <TabList size="small" selectedValue={p.tab} onTabSelect={(_, d) => p.onTab(d.value as RibbonTab)}>
        <Tab value="home">Domov</Tab>
        {c && <Tab value="message">Správa</Tab>}
        {c && <Tab value="insert">Vložiť</Tab>}
        <Tab value="view">Zobrazenie</Tab>
      </TabList>
      <Toolbar size={sz} className="ribbon-bar">
        {c && p.tab === 'message' ? (
          <>
            <ToolbarButton appearance="primary" disabled={c.state.sending} onClick={() => c.run((h) => h.send())}>Odoslať</ToolbarButton>
            <SendLaterMenu onPick={(d) => c.run((h) => h.sendLater(d))} />
            <ToolbarButton onClick={() => c.run((h) => h.discard())}>Zahodiť</ToolbarButton>
            <ToolbarButton onClick={() => c.run((h) => h.popOut())}>Samostatné okno</ToolbarButton>
            <ToolbarDivider />
            <ToolbarButton icon={<AttachIcon />} onClick={() => c.run((h) => h.attach())}>Priložiť súbor</ToolbarButton>
            <Menu>
              <MenuTrigger disableButtonEnhancement>
                <ToolbarButton icon={<SignatureIcon />}>Podpis</ToolbarButton>
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  {c.signatures.map((s) => (
                    <MenuItem key={s.id} onClick={() => c.run((h) => h.setSignature(s.id))}>{s.name}</MenuItem>
                  ))}
                  <MenuItem onClick={() => c.run((h) => h.setSignature(null))}>(bez podpisu)</MenuItem>
                  <MenuItem onClick={p.onSignatures}>Upraviť podpisy…</MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
            <Menu>
              <MenuTrigger disableButtonEnhancement>
                <ToolbarButton icon={<TemplateIcon />}>Šablóny</ToolbarButton>
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  {c.templates.map((t) => (
                    <MenuItem key={t.id} onClick={() => c.run((h) => h.insertTemplate(t.id))}>{t.name}</MenuItem>
                  ))}
                  {!c.templates.length && <MenuItem disabled>Zatiaľ žiadne</MenuItem>}
                  <MenuItem onClick={() => c.run((h) => h.saveAsTemplate())}>Uložiť túto správu ako šablónu</MenuItem>
                  <MenuItem onClick={c.onManageTemplates}>Spravovať šablóny…</MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
            <ToolbarDivider />
            <ToolbarButton onClick={() => c.run((h) => h.setPriority(c.state.priority === 'high' ? 'normal' : 'high'))} aria-pressed={c.state.priority === 'high'}>
              {c.state.priority === 'high' ? '✓ ' : ''}Vysoká dôležitosť
            </ToolbarButton>
            <ToolbarButton onClick={() => c.run((h) => h.setPriority(c.state.priority === 'low' ? 'normal' : 'low'))} aria-pressed={c.state.priority === 'low'}>
              {c.state.priority === 'low' ? '✓ ' : ''}Nízka dôležitosť
            </ToolbarButton>
            <ToolbarButton onClick={() => c.run((h) => h.toggleReceipt())} aria-pressed={c.state.receipt}>
              {c.state.receipt ? '✓ ' : ''}Potvrdenie o prečítaní
            </ToolbarButton>
            <Menu>
              <MenuTrigger disableButtonEnhancement>
                <ToolbarButton icon={<SnoozeIcon />}>{c.state.followUp ? `Pripomenúť o ${c.state.followUp} d` : 'Pripomenúť, ak neodpovie'}</ToolbarButton>
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  {[0, 1, 2, 3, 5, 7, 14].map((d) => (
                    <MenuItem key={d} onClick={() => c.run((h) => h.setFollowUp(d))}>
                      {(c.state.followUp ?? 0) === d ? '✓ ' : ''}{d ? `ak neodpovie do ${d} ${d === 1 ? 'dňa' : 'dní'}` : 'nepripomínať'}
                    </MenuItem>
                  ))}
                </MenuList>
              </MenuPopover>
            </Menu>
            <ToolbarDivider />
            <Menu>
              <MenuTrigger disableButtonEnhancement>
                <ToolbarButton icon={<LockIcon />}>
                  {c.state.security?.encrypt ? 'Šifrované' : c.state.security?.sign ? 'Podpísané' : 'Zabezpečenie'}
                </ToolbarButton>
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  {(() => {
                    const sec = c.state.security ?? {};
                    const set = (x: typeof sec) => c.run((h) => h.setSecurity(x));
                    return (
                      <>
                        <MenuItem onClick={() => set({})}>{!sec.sign && !sec.encrypt ? '✓ ' : ''}Bez zabezpečenia</MenuItem>
                        <MenuItem onClick={() => set({ sign: 'smime' })}>{sec.sign === 'smime' && !sec.encrypt ? '✓ ' : ''}Podpísať (S/MIME)</MenuItem>
                        <MenuItem onClick={() => set({ sign: 'smime', encrypt: 'smime' })}>{sec.encrypt === 'smime' ? '✓ ' : ''}Podpísať a šifrovať (S/MIME)</MenuItem>
                        <MenuItem onClick={() => set({ sign: 'pgp' })}>{sec.sign === 'pgp' && !sec.encrypt ? '✓ ' : ''}Podpísať (PGP)</MenuItem>
                        <MenuItem onClick={() => set({ sign: 'pgp', encrypt: 'pgp' })}>{sec.encrypt === 'pgp' ? '✓ ' : ''}Podpísať a šifrovať (PGP)</MenuItem>
                        <MenuItem onClick={c.onManageCrypto}>Kľúče a certifikáty…</MenuItem>
                      </>
                    );
                  })()}
                </MenuList>
              </MenuPopover>
            </Menu>
          </>
        ) : c && p.tab === 'insert' ? (
          <>
            <ToolbarButton icon={<AttachIcon />} onClick={() => c.run((h) => h.attach())}>Súbor</ToolbarButton>
            <ToolbarButton onClick={() => c.run((h) => h.insertImage())}>Obrázok</ToolbarButton>
            <ToolbarButton onClick={() => c.run((h) => h.insertTable())}>Tabuľka</ToolbarButton>
            <ToolbarButton onClick={() => c.run((h) => h.insertHr())}>Vodorovná čiara</ToolbarButton>
            <ToolbarButton onClick={() => c.run((h) => h.insertDate())}>Dátum a čas</ToolbarButton>
          </>
        ) : p.tab === 'home' || (!c && (p.tab === 'message' || p.tab === 'insert')) ? (
          <>
            <ToolbarButton appearance="primary" icon={<ComposeIcon />} onClick={p.onCompose}>Nová správa</ToolbarButton>
            <ToolbarButton className="sr-btn" icon={<SendReceiveIcon />} onClick={p.onSendReceive} title="Skontrolovať poštu na všetkých kontách a odoslať čakajúce správy (F9)">Odoslať a prijať</ToolbarButton>
            <ToolbarDivider />
            <ToolbarButton icon={<ReplyIcon />} disabled={!p.hasMessage} onClick={p.onReply}>Odpovedať</ToolbarButton>
            <ToolbarButton icon={<ReplyAllIcon />} disabled={!p.hasMessage} onClick={p.onReplyAll}>Všetkým</ToolbarButton>
            <ToolbarButton icon={<ForwardIcon />} disabled={!p.hasMessage} onClick={p.onForward}>Preposlať</ToolbarButton>
            <ToolbarDivider />
            <ToolbarButton icon={<DeleteIcon />} disabled={!p.hasMessage} onClick={p.onDelete}>Odstrániť</ToolbarButton>
            <ToolbarButton icon={<ArchiveIcon />} disabled={!p.hasMessage} onClick={p.onArchive}>Archív</ToolbarButton>
            <Menu>
              <MenuTrigger disableButtonEnhancement>
                <ToolbarButton icon={<MoveIcon />} disabled={!p.hasMessage || !p.canMove}>Presunúť</ToolbarButton>
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  {p.folders.map((f) => (
                    <MenuItem key={f.path} onClick={() => p.onMove(f.path)} style={{ paddingLeft: 12 + f.depth * 14 }}>
                      {f.name}
                    </MenuItem>
                  ))}
                </MenuList>
              </MenuPopover>
            </Menu>
            <ToolbarDivider />
            <ToolbarButton icon={p.seen ? <MarkUnreadIcon /> : <MarkReadIcon />} disabled={!p.hasMessage} onClick={p.onToggleSeen}>
              {p.seen ? 'Neprečítané' : 'Prečítané'}
            </ToolbarButton>
            <ToolbarButton icon={<FlagIcon />} disabled={!p.hasMessage} onClick={p.onToggleFlag}>
              {p.flagged ? 'Zrušiť vlajku' : 'Vlajka'}
            </ToolbarButton>
            <Menu>
              <MenuTrigger disableButtonEnhancement>
                <ToolbarButton icon={<TagIcon />} disabled={!p.hasMessage}>Kategorizovať</ToolbarButton>
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  {p.categories.map((c) => {
                    const on = p.keywords.includes(c.id);
                    return (
                      <MenuItem key={c.id} onClick={() => p.onCategory(c.id, !on)} icon={<span className="cat-dot" style={{ background: c.color }} />}>
                        {on ? '✓ ' : ''}{c.name}
                      </MenuItem>
                    );
                  })}
                  <MenuItem onClick={() => p.onOrganize('categories')}>Spravovať kategórie…</MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
            <ToolbarButton icon={<PinIcon />} disabled={!p.hasMessage} onClick={p.onTogglePin}>
              {p.pinned ? 'Odopnúť' : 'Pripnúť'}
            </ToolbarButton>
            <SnoozeMenu disabled={!p.hasMessage} onSnooze={p.onSnooze} />
            <TaskMenu disabled={!p.hasMessage} onTask={p.onTask} />
            <ToolbarDivider />
            <Menu>
              <MenuTrigger disableButtonEnhancement>
                <ToolbarButton icon={<QuickStepIcon />}>Rýchle kroky</ToolbarButton>
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  {p.quickSteps.map((q) => (
                    <MenuItem key={q.id} disabled={!p.hasMessage} onClick={() => p.onQuickStep(q)}>{q.name}</MenuItem>
                  ))}
                  {!p.quickSteps.length && <MenuItem disabled>Zatiaľ žiadne</MenuItem>}
                  <MenuItem onClick={() => p.onOrganize('quick')}>Spravovať rýchle kroky…</MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
            <ToolbarButton icon={<RulesIcon />} onClick={() => p.onOrganize('rules')}>Pravidlá</ToolbarButton>
            <ToolbarDivider />
            <ToolbarButton icon={<SyncIcon />} onClick={p.onRefresh}>Obnoviť</ToolbarButton>
            <ToolbarButton icon={<SignatureIcon />} onClick={p.onSignatures}>Podpisy</ToolbarButton>
            <Menu>
              <MenuTrigger disableButtonEnhancement>
                <ToolbarButton icon={<FileIcon />}>Súbor</ToolbarButton>
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  <MenuItem disabled={!p.hasMessage} onClick={() => p.onIo('print')}>{kb('Tlačiť správu (Ctrl+P)')}</MenuItem>
                  <MenuItem disabled={!p.hasMessage} onClick={() => p.onIo('pdf')}>Uložiť správu ako PDF</MenuItem>
                  <MenuItem disabled={!p.hasMessage} onClick={() => p.onIo('eml')}>Uložiť správu ako .eml</MenuItem>
                  <MenuItem onClick={() => p.onIo('mbox')}>Exportovať priečinok (.mbox)</MenuItem>
                  <MenuItem onClick={() => p.onIo('import')}>Importovať .eml / .mbox…</MenuItem>
                  <MenuItem onClick={() => p.onIo('thunderbird')}>Importovať z Thunderbirdu…</MenuItem>
                  <MenuItem onClick={() => p.onIo('backup')}>Zálohovať nastavenia…</MenuItem>
                  <MenuItem onClick={() => p.onIo('restore')}>Obnoviť nastavenia zo zálohy…</MenuItem>
                  <MenuItem onClick={() => p.onIo('shortcuts')}>Klávesové skratky</MenuItem>
                  <MenuItem onClick={() => p.onIo('update')}>Skontrolovať aktualizácie</MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
            {p.onEditAccount && (
              <ToolbarButton icon={<SettingsIcon />} onClick={p.onEditAccount}>Nastavenia konta</ToolbarButton>
            )}
          </>
        ) : (
          <>
            <label className="ribbon-label">Téma</label>
            <Select size="small" value={p.theme} onChange={(_, d) => p.onTheme(d.value as ThemeKey)}>
              {[...BUILTIN_THEMES, ...p.customThemes].map((t) => (
                <option key={t.id} value={t.id}>{t.name}</option>
              ))}
            </Select>
            <ToolbarButton onClick={p.onEditThemes}>Vlastné témy…</ToolbarButton>
            <ToolbarDivider />
            <ToolbarButton icon={<PanelIcon />} onClick={p.onToggleRight}>
              {p.rightOpen ? 'Skryť pravý panel' : 'Zobraziť pravý panel'}
            </ToolbarButton>
            <ToolbarButton onClick={() => p.onCompact(!p.compact)}>{p.compact ? 'Plný ribbon' : 'Kompaktný ribbon'}</ToolbarButton>
            {p.onMyDay && (
              <>
                <ToolbarButton icon={<CalendarSmallIcon />} onClick={() => p.onMyDay!({ ...p.myDay!, calendar: !p.myDay!.calendar })} aria-pressed={p.myDay!.calendar} title="Náhľad kalendára vedľa čítacieho panela">
                  {p.myDay!.calendar ? '✓ ' : ''}Kalendár pri pošte
                </ToolbarButton>
                <ToolbarButton icon={<TasksSmallIcon />} onClick={() => p.onMyDay!({ ...p.myDay!, tasks: !p.myDay!.tasks })} aria-pressed={p.myDay!.tasks} title="Náhľad úloh vedľa čítacieho panela">
                  {p.myDay!.tasks ? '✓ ' : ''}Úlohy pri pošte
                </ToolbarButton>
              </>
            )}
            <ToolbarButton icon={<ConversationIcon />} onClick={() => p.onConversations(!p.conversations)} aria-pressed={p.conversations}>
              {p.conversations ? '✓ ' : ''}Konverzácie
            </ToolbarButton>
            <ToolbarButton icon={<SearchSmallIcon />} onClick={() => p.onOrganize('searches')}>Uložené hľadania</ToolbarButton>
            <ToolbarDivider />
            <ToolbarButton onClick={() => p.onDensity(p.density === 'compact' ? 'comfortable' : 'compact')}>
              {p.density === 'compact' ? 'Pohodlný zoznam' : 'Kompaktný zoznam'}
            </ToolbarButton>
            <ToolbarButton onClick={() => p.onPreviewLine(!p.previewLine)} aria-pressed={p.previewLine}>
              {p.previewLine ? '✓ ' : ''}Náhľad textu v zozname
            </ToolbarButton>
            <ToolbarButton onClick={() => p.onOpenMode(p.openMode === 'window' ? 'tab' : 'window')} title="Dvojklik alebo Enter na správu; so Shiftom opačne">
              Dvojklik otvorí: {p.openMode === 'window' ? 'nové okno' : 'kartu'}
            </ToolbarButton>
            <ToolbarButton onClick={() => p.onRunInBackground(!p.runInBackground)} aria-pressed={p.runInBackground}>
              {p.runInBackground ? '✓ ' : ''}Po zatvorení bežať pri hodinách
            </ToolbarButton>
            <ToolbarDivider />
            <ToolbarButton onClick={p.onToggleGpu}>{p.gpuOff ? 'Zapnúť hardvérovú akceleráciu' : 'Vypnúť hardvérovú akceleráciu'}</ToolbarButton>
            <ToolbarButton onClick={p.onOpenLog}>Záznam chýb</ToolbarButton>
          </>
        )}
      </Toolbar>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Pás modulov: mení sa podľa typu konta.

export type ImapModule = 'mail' | 'calendar' | 'tasks' | 'contacts';
export type M365Module = 'mail' | 'calendar' | 'todo' | 'people' | 'copilot' | 'planner' | 'teams' | 'onedrive' | 'sharepoint' | 'word' | 'excel' | 'powerpoint' | 'onenote' | 'forms' | 'loop';

export interface ModuleDef<K extends string> {
  key: K;
  label: string;
  icon: ReactNode;
  sep?: boolean;
  /** nadpis sekcie nad ikonou (napr. Firemné aplikácie) */
  section?: string;
}

export const IMAP_MODULES: ModuleDef<ImapModule>[] = [
  { key: 'mail', label: 'Pošta', icon: <MailIcon /> },
  { key: 'calendar', label: 'Kalendár', icon: <CalendarIcon /> },
  { key: 'tasks', label: 'Úlohy', icon: <TasksIcon /> },
  { key: 'contacts', label: 'Kontakty', icon: <ContactsIcon /> },
];

/** Firemný Outlook (Microsoft 365 pre prácu): pošta + firemné aplikácie. */
export const M365_WORK_MODULES: ModuleDef<M365Module>[] = [
  { key: 'mail', label: 'Pošta', icon: <MailIcon />, section: 'Outlook' },
  { key: 'calendar', label: 'Kalendár', icon: <CalendarIcon /> },
  { key: 'todo', label: 'To Do', icon: <TasksIcon /> },
  { key: 'people', label: 'Ľudia', icon: <ContactsIcon /> },
  { key: 'copilot', label: 'Copilot', icon: <CopilotIcon />, sep: true, section: 'Firemné' },
  { key: 'teams', label: 'Teams (firemné)', icon: <TeamsIcon /> },
  { key: 'planner', label: 'Planner', icon: <PlannerIcon /> },
  { key: 'onedrive', label: 'OneDrive pre prácu', icon: <CloudIcon /> },
  { key: 'sharepoint', label: 'SharePoint', icon: <SharePointIcon /> },
  { key: 'word', label: 'Word', icon: <WordIcon /> },
  { key: 'excel', label: 'Excel', icon: <ExcelIcon /> },
  { key: 'powerpoint', label: 'PowerPoint', icon: <PowerPointIcon /> },
  { key: 'onenote', label: 'OneNote', icon: <OneNoteIcon /> },
  { key: 'forms', label: 'Forms', icon: <FormsIcon /> },
  { key: 'loop', label: 'Loop', icon: <LoopIcon /> },
];

/** Osobný Outlook (Outlook.com, Hotmail, Live): pošta + osobné aplikácie. */
/** Vložený webmail (Zimbra, Roundcube, SOGo, Horde…): len samotná stránka. */
export const WEBMAIL_MODULES: ModuleDef<M365Module>[] = [{ key: 'mail', label: 'Webmail', icon: <MailIcon /> }];

export const M365_PERSONAL_MODULES: ModuleDef<M365Module>[] = [
  { key: 'mail', label: 'Pošta', icon: <MailIcon />, section: 'Outlook' },
  { key: 'calendar', label: 'Kalendár', icon: <CalendarIcon /> },
  { key: 'todo', label: 'To Do', icon: <TasksIcon /> },
  { key: 'people', label: 'Ľudia', icon: <ContactsIcon /> },
  { key: 'copilot', label: 'Copilot', icon: <CopilotIcon />, sep: true, section: 'Osobné' },
  { key: 'teams', label: 'Teams (osobné)', icon: <TeamsIcon /> },
  { key: 'onedrive', label: 'OneDrive', icon: <CloudIcon /> },
  { key: 'word', label: 'Word', icon: <WordIcon /> },
  { key: 'excel', label: 'Excel', icon: <ExcelIcon /> },
  { key: 'powerpoint', label: 'PowerPoint', icon: <PowerPointIcon /> },
  { key: 'onenote', label: 'OneNote', icon: <OneNoteIcon /> },
  { key: 'forms', label: 'Forms', icon: <FormsIcon /> },
];

export function ModuleRail<K extends string>(props: { modules: ModuleDef<K>[]; active: K | null; color: string; onSelect: (k: K) => void; footer?: React.ReactNode }) {
  return (
    <nav className="rail" aria-label="Moduly konta">
      {props.modules.map((m) => {
        const on = props.active === m.key;
        return (
          <div key={m.key}>
            {m.sep && <div className="rail-sep" />}
            {m.section && <div className="rail-section">{m.section}</div>}
            <Tooltip content={m.label} relationship="label" positioning="after">
              <button
                className={`rail-btn ${on ? 'active' : ''}`}
                style={on ? { color: props.color, boxShadow: `inset 3px 0 0 ${props.color}` } : undefined}
                onClick={() => props.onSelect(m.key)}
                aria-pressed={on}
              >
                {m.icon}
                <span className="rail-label">{m.label}</span>
              </button>
            </Tooltip>
          </div>
        );
      })}
      {props.footer && <div className="rail-footer">{props.footer}</div>}
    </nav>
  );
}

/** Spodok pásu modulov IMAP konta: nastavenia Zimbry a nastavenia tohto konta (nie globálne v ozubenom kolieske). */
export function AccountRailFooter(props: { accountName: string; onZimbra?: (tab?: 'ooo' | 'filters' | 'shares' | 'spam' | 'sync' | 'password') => void; onAccount: () => void }) {
  return (
    <>
      {props.onZimbra && (
        <Menu positioning="after-bottom">
          <MenuTrigger disableButtonEnhancement>
            <Tooltip content={`Zimbra – ${props.accountName}: neprítomnosť, filtre, zdieľania, heslo…`} relationship="label" positioning="after">
              <button className="rail-btn">
                <ServerIcon />
                <span className="rail-label">Zimbra</span>
              </button>
            </Tooltip>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuGroupHeader>Zimbra – {props.accountName}</MenuGroupHeader>
              <MenuItem onClick={() => props.onZimbra!('ooo')}>Neprítomnosť (automatická odpoveď)</MenuItem>
              <MenuItem onClick={() => props.onZimbra!('filters')}>Filtre na serveri</MenuItem>
              <MenuItem onClick={() => props.onZimbra!('shares')}>Zdieľané od kolegov</MenuItem>
              <MenuItem onClick={() => props.onZimbra!('spam')}>Spam: dôveryhodní a blokovaní</MenuItem>
              <MenuItem onClick={() => props.onZimbra!('sync')}>Podpisy, identity a značky</MenuItem>
              <MenuItem onClick={() => props.onZimbra!('password')}>Zmeniť heslo</MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>
      )}
      <Tooltip content={`Nastavenia konta ${props.accountName}`} relationship="label" positioning="after">
        <button className="rail-btn" onClick={props.onAccount}>
          <SettingsIcon />
          <span className="rail-label">Konto</span>
        </button>
      </Tooltip>
    </>
  );
}
