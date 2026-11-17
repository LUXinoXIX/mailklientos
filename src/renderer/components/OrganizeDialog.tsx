// Organizácia pošty: pravidlá (automaticky pri novej pošte), rýchle kroky (jedno kliknutie = viac akcií),
// kategórie (farebné štítky = IMAP kľúčové slová) a uložené hľadania.
import {
  Button, Checkbox, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, Input, Select, Tab, TabList,
} from '@fluentui/react-components';
import { useState } from 'react';
import type { ImapAccountPublic } from '../../core/account';
import type { FolderInfo } from '../../core/api';
import {
  DEFAULT_CATEGORIES, categoriesOf, type ActionType, type AppSettings, type Category, type QuickStep, type Rule, type RuleAction,
  type RuleField, type SavedSearch, type Vacation,
} from '../../core/settings';
import { SignatureEditor } from './SignatureEditor';

export type OrganizeTab = 'rules' | 'quick' | 'categories' | 'searches' | 'vacation';

const newId = () => Math.random().toString(36).slice(2, 10);

const ACTIONS: Array<{ type: ActionType; label: string }> = [
  { type: 'move', label: 'Presunúť do priečinka' },
  { type: 'read', label: 'Označiť ako prečítané' },
  { type: 'unread', label: 'Označiť ako neprečítané' },
  { type: 'flag', label: 'Označiť vlajkou' },
  { type: 'category', label: 'Priradiť kategóriu' },
  { type: 'pin', label: 'Pripnúť' },
  { type: 'archive', label: 'Archivovať' },
  { type: 'delete', label: 'Odstrániť' },
];

const FIELDS: Array<{ v: RuleField; label: string }> = [
  { v: 'from', label: 'Odosielateľ obsahuje' },
  { v: 'to', label: 'Príjemca obsahuje' },
  { v: 'subject', label: 'Predmet obsahuje' },
  { v: 'any', label: 'Odosielateľ, príjemca alebo predmet obsahuje' },
];

const COLORS = ['#c50f1f', '#d83b01', '#b45309', '#107c10', '#2F6F5E', '#0e7490', '#0f6cbd', '#6d28d9', '#be123c', '#616161'];

function ActionsEditor(props: { actions: RuleAction[]; onChange: (a: RuleAction[]) => void; folders: FolderInfo[]; categories: Category[] }) {
  const set = (i: number, p: Partial<RuleAction>) => props.onChange(props.actions.map((a, j) => (j === i ? { ...a, ...p } : a)));
  return (
    <div className="org-actions">
      {props.actions.map((a, i) => (
        <div key={i} className="org-row">
          <Select size="small" value={a.type} onChange={(_, d) => set(i, { type: d.value as ActionType })}>
            {ACTIONS.map((x) => <option key={x.type} value={x.type}>{x.label}</option>)}
          </Select>
          {a.type === 'move' &&
            (props.folders.length ? (
              <Select size="small" value={a.folder ?? ''} onChange={(_, d) => set(i, { folder: d.value })}>
                <option value="">(vyber priečinok)</option>
                {props.folders.map((f) => <option key={f.path} value={f.path}>{'  '.repeat(f.depth)}{f.name}</option>)}
              </Select>
            ) : (
              <Input size="small" placeholder="Názov priečinka" value={a.folder ?? ''} onChange={(_, d) => set(i, { folder: d.value })} />
            ))}
          {a.type === 'category' && (
            <Select size="small" value={a.category ?? ''} onChange={(_, d) => set(i, { category: d.value })}>
              <option value="">(vyber kategóriu)</option>
              {props.categories.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </Select>
          )}
          <Button size="small" appearance="subtle" aria-label="Odobrať akciu" onClick={() => props.onChange(props.actions.filter((_, j) => j !== i))}>×</Button>
        </div>
      ))}
      <Button size="small" onClick={() => props.onChange([...props.actions, { type: 'read' }])}>+ Pridať akciu</Button>
    </div>
  );
}

export function OrganizeDialog(props: {
  settings: AppSettings;
  accounts: ImapAccountPublic[];
  /** priečinky podľa konta (pre výber cieľa presunu) */
  folders: Record<string, FolderInfo[]>;
  defaultAccountId?: string;
  tab?: OrganizeTab;
  onRunRules?: () => Promise<string>;
  onSave: (s: AppSettings) => void;
  onClose: () => void;
}) {
  const [tab, setTab] = useState<OrganizeTab>(props.tab ?? 'rules');
  const [rules, setRules] = useState<Rule[]>(props.settings.rules ?? []);
  const [quick, setQuick] = useState<QuickStep[]>(props.settings.quickSteps ?? []);
  const [cats, setCats] = useState<Category[]>(categoriesOf(props.settings));
  const [searches, setSearches] = useState<SavedSearch[]>(props.settings.savedSearches ?? []);
  const [runInfo, setRunInfo] = useState('');
  const [vac, setVac] = useState<Vacation>(
    props.settings.vacation ?? { enabled: false, subject: 'Automatická odpoveď', html: '<p>Dobrý deň,</p><p>som mimo kancelárie a na správy odpoviem po návrate.</p>', text: '' },
  );
  const foldersFor = (accountId?: string) => props.folders[accountId ?? props.defaultAccountId ?? props.accounts[0]?.id ?? ''] ?? [];

  const save = () => props.onSave({ ...props.settings, rules, quickSteps: quick, categories: cats, savedSearches: searches, vacation: vac });

  return (
    <Dialog open onOpenChange={(_, d) => !d.open && props.onClose()}>
      <DialogSurface style={{ maxWidth: 980, width: '92vw' }}>
        <DialogBody style={{ display: 'flex', flexDirection: 'column', maxHeight: '86vh' }}>
          <DialogTitle>Organizácia pošty</DialogTitle>
          <DialogContent style={{ overflow: 'auto', display: 'flex', flexDirection: 'column', gap: 12 }}>
            <TabList selectedValue={tab} onTabSelect={(_, d) => setTab(d.value as OrganizeTab)}>
              <Tab value="rules">Pravidlá</Tab>
              <Tab value="quick">Rýchle kroky</Tab>
              <Tab value="categories">Kategórie</Tab>
              <Tab value="searches">Uložené hľadania</Tab>
              <Tab value="vacation">Neprítomnosť</Tab>
            </TabList>

            {tab === 'rules' && (
              <div className="org-list">
                <div className="muted small">
                  Pravidlá sa spúšťajú automaticky na novú poštu v Doručenej pošte, keď je MailKlient spustený. Poradie rozhoduje: pravidlo s presunom ďalšie pravidlá zastaví.
                </div>
                {rules.map((r, i) => {
                  const set = (p: Partial<Rule>) => setRules(rules.map((x) => (x.id === r.id ? { ...x, ...p } : x)));
                  return (
                    <div key={r.id} className="org-card">
                      <div className="org-row">
                        <Checkbox checked={r.enabled} onChange={(_, d) => set({ enabled: !!d.checked })} aria-label="Zapnuté" />
                        <Input className="grow" value={r.name} onChange={(_, d) => set({ name: d.value })} aria-label="Názov pravidla" />
                        <Select size="small" value={r.accountId ?? ''} onChange={(_, d) => set({ accountId: d.value || undefined })} aria-label="Konto">
                          <option value="">Všetky kontá</option>
                          {props.accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
                        </Select>
                        <Button size="small" appearance="subtle" disabled={i === 0} aria-label="Vyššie" onClick={() => setRules([...rules.slice(0, i - 1), r, rules[i - 1]!, ...rules.slice(i + 1)])}>↑</Button>
                        <Button size="small" appearance="subtle" aria-label="Odstrániť pravidlo" onClick={() => setRules(rules.filter((x) => x.id !== r.id))}>Odstrániť</Button>
                      </div>
                      <div className="org-sub">
                        <b>Ak</b>
                        <Select size="small" value={r.match} onChange={(_, d) => set({ match: d.value as 'all' | 'any' })}>
                          <option value="all">platia všetky podmienky</option>
                          <option value="any">platí ktorákoľvek podmienka</option>
                        </Select>
                      </div>
                      {r.conds.map((c, j) => (
                        <div key={j} className="org-row org-indent">
                          <Select size="small" value={c.field} onChange={(_, d) => set({ conds: r.conds.map((x, k) => (k === j ? { ...x, field: d.value as RuleField } : x)) })}>
                            {FIELDS.map((f) => <option key={f.v} value={f.v}>{f.label}</option>)}
                          </Select>
                          <Input size="small" className="grow" placeholder="napr. faktura, @dodavatel.sk" value={c.contains} onChange={(_, d) => set({ conds: r.conds.map((x, k) => (k === j ? { ...x, contains: d.value } : x)) })} />
                          <Button size="small" appearance="subtle" aria-label="Odobrať podmienku" onClick={() => set({ conds: r.conds.filter((_, k) => k !== j) })}>×</Button>
                        </div>
                      ))}
                      <div className="org-indent">
                        <Button size="small" onClick={() => set({ conds: [...r.conds, { field: 'from', contains: '' }] })}>+ Pridať podmienku</Button>
                      </div>
                      <div className="org-sub"><b>Potom</b></div>
                      <div className="org-indent">
                        <ActionsEditor actions={r.actions} onChange={(a) => set({ actions: a })} folders={foldersFor(r.accountId)} categories={cats} />
                      </div>
                      <div className="org-indent">
                        <Checkbox label="Ďalšie pravidlá už nespúšťať" checked={!!r.stop} onChange={(_, d) => set({ stop: !!d.checked })} />
                      </div>
                    </div>
                  );
                })}
                <div className="org-row">
                  <Button
                    appearance="primary"
                    onClick={() =>
                      setRules([...rules, { id: newId(), name: 'Nové pravidlo', enabled: true, match: 'all', conds: [{ field: 'from', contains: '' }], actions: [{ type: 'move' }] }])
                    }
                  >
                    + Nové pravidlo
                  </Button>
                  {props.onRunRules && (
                    <Button
                      onClick={async () => {
                        props.onSave({ ...props.settings, rules, quickSteps: quick, categories: cats, savedSearches: searches, vacation: vac });
                        setRunInfo('Spúšťam…');
                        setRunInfo(await props.onRunRules!());
                      }}
                    >
                      Uložiť a spustiť teraz na otvorený priečinok
                    </Button>
                  )}
                  <span className="muted small">{runInfo}</span>
                </div>
              </div>
            )}

            {tab === 'quick' && (
              <div className="org-list">
                <div className="muted small">Rýchly krok vykoná na vybranej správe viac akcií naraz (napr. prečítané + kategória + presun). Nájdete ich na karte Domov.</div>
                {quick.map((q) => {
                  const set = (p: Partial<QuickStep>) => setQuick(quick.map((x) => (x.id === q.id ? { ...x, ...p } : x)));
                  return (
                    <div key={q.id} className="org-card">
                      <div className="org-row">
                        <Input className="grow" value={q.name} onChange={(_, d) => set({ name: d.value })} aria-label="Názov rýchleho kroku" />
                        <Button size="small" appearance="subtle" onClick={() => setQuick(quick.filter((x) => x.id !== q.id))}>Odstrániť</Button>
                      </div>
                      <ActionsEditor actions={q.actions} onChange={(a) => set({ actions: a })} folders={foldersFor()} categories={cats} />
                    </div>
                  );
                })}
                <div className="org-row">
                  <Button appearance="primary" onClick={() => setQuick([...quick, { id: newId(), name: 'Nový rýchly krok', actions: [{ type: 'read' }, { type: 'archive' }] }])}>+ Nový rýchly krok</Button>
                </div>
              </div>
            )}

            {tab === 'categories' && (
              <div className="org-list">
                <div className="muted small">
                  Kategórie sa ukladajú na server ako IMAP štítky, takže ich vidno aj na inom počítači. Prvých päť je zhodných s Thunderbirdom.
                </div>
                {cats.map((c) => (
                  <div key={c.id} className="org-row">
                    <span className="cat-chip" style={{ background: `${c.color}22`, color: c.color, borderColor: `${c.color}55`, minWidth: 90 }}>{c.name || '…'}</span>
                    <Input className="grow" value={c.name} onChange={(_, d) => setCats(cats.map((x) => (x.id === c.id ? { ...x, name: d.value } : x)))} aria-label="Názov kategórie" />
                    <div className="swatch-row">
                      {COLORS.map((col) => (
                        <button
                          key={col}
                          className={`swatch ${c.color === col ? 'on' : ''}`}
                          style={{ background: col }}
                          aria-label={`Farba ${col}`}
                          onClick={() => setCats(cats.map((x) => (x.id === c.id ? { ...x, color: col } : x)))}
                        />
                      ))}
                    </div>
                    <Button size="small" appearance="subtle" onClick={() => setCats(cats.filter((x) => x.id !== c.id))}>Odstrániť</Button>
                  </div>
                ))}
                <div className="org-row">
                  <Button appearance="primary" onClick={() => setCats([...cats, { id: `mk_${newId()}`, name: 'Nová kategória', color: COLORS[cats.length % COLORS.length]! }])}>+ Nová kategória</Button>
                  <Button appearance="subtle" onClick={() => setCats(DEFAULT_CATEGORIES)}>Obnoviť predvolené</Button>
                </div>
              </div>
            )}

            {tab === 'vacation' && (
              <div className="org-list">
                <div className="muted small">
                  Automatická odpoveď sa posiela, keď je MailKlient spustený: každému odosielateľovi najviac raz za 4 dni, nie na hromadnú poštu ani automatické správy.
                  Ak máte poštu na Zimbre, použite radšej neprítomnosť priamo na serveri (ozubené koliesko → Zimbra) – funguje aj s vypnutým počítačom.
                </div>
                <Checkbox label="Posielať automatickú odpoveď" checked={vac.enabled} onChange={(_, d) => setVac({ ...vac, enabled: !!d.checked })} />
                <div className="org-row">
                  <label className="sig-label" htmlFor="vac-from">Od</label>
                  <input id="vac-from" type="date" className="date-input" value={vac.from ?? ''} onChange={(e) => setVac({ ...vac, from: e.target.value || undefined })} />
                  <label className="sig-label" htmlFor="vac-to">Do</label>
                  <input id="vac-to" type="date" className="date-input" value={vac.until ?? ''} onChange={(e) => setVac({ ...vac, until: e.target.value || undefined })} />
                  <span className="muted small">prázdne = bez obmedzenia</span>
                </div>
                <div className="org-row">
                  <span className="sig-label">Kontá</span>
                  {props.accounts.map((a) => (
                    <Checkbox
                      key={a.id}
                      label={a.name}
                      checked={!vac.accountIds?.length || vac.accountIds.includes(a.id)}
                      onChange={(_, d) => {
                        const all = vac.accountIds?.length ? vac.accountIds : props.accounts.map((x) => x.id);
                        const next = d.checked ? [...new Set([...all, a.id])] : all.filter((x) => x !== a.id);
                        setVac({ ...vac, accountIds: next.length === props.accounts.length ? undefined : next });
                      }}
                    />
                  ))}
                </div>
                <div className="org-row">
                  <label className="sig-label" htmlFor="vac-subj">Predmet</label>
                  <Input id="vac-subj" className="grow" value={vac.subject} onChange={(_, d) => setVac({ ...vac, subject: d.value })} />
                </div>
                <SignatureEditor id="vacation" html={vac.html} onChange={(html, text) => setVac((v) => ({ ...v, html, text }))} />
              </div>
            )}

            {tab === 'searches' && (
              <div className="org-list">
                <div className="muted small">Uložené hľadania sú v paneli priečinkov ako virtuálne priečinky. Nové uložíte aj tlačidlom pri výsledkoch hľadania.</div>
                {searches.map((x) => {
                  const set = (p: Partial<SavedSearch>) => setSearches(searches.map((y) => (y.id === x.id ? { ...y, ...p } : y)));
                  return (
                    <div key={x.id} className="org-row">
                      <Input value={x.name} onChange={(_, d) => set({ name: d.value })} aria-label="Názov" />
                      <Input className="grow" value={x.query} onChange={(_, d) => set({ query: d.value })} aria-label="Hľadaný text" />
                      <Select size="small" value={x.accountId ?? ''} onChange={(_, d) => set({ accountId: d.value || undefined })}>
                        <option value="">Všetky kontá</option>
                        {props.accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
                      </Select>
                      <Button size="small" appearance="subtle" onClick={() => setSearches(searches.filter((y) => y.id !== x.id))}>Odstrániť</Button>
                    </div>
                  );
                })}
                <div className="org-row">
                  <Button appearance="primary" onClick={() => setSearches([...searches, { id: newId(), name: 'Faktúry', query: 'faktura' }])}>+ Nové hľadanie</Button>
                </div>
              </div>
            )}
          </DialogContent>
          <DialogActions style={{ justifyContent: 'flex-end', alignSelf: 'flex-end' }}>
            <Button appearance="primary" onClick={save}>Uložiť</Button>
            <Button onClick={props.onClose}>Zrušiť</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
