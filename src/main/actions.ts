// Akcie nad správami pre pravidlá a rýchle kroky: presun, prečítané, vlajka, kategória, pripnutie, archív, odstránenie.
import type { AccountConfig } from '../core/account';
import { PIN_KEYWORD, ruleMatches, type Rule, type RuleAction } from '../core/settings';
import * as cache from './cache';
import { log } from './diagnostics';
import * as mail from './mail';

/** Vykoná akcie postupne; presun/archív/odstránenie ukončia reťaz (správa už v priečinku nie je). */
export async function applyActions(acc: AccountConfig, folder: string, uids: number[], actions: RuleAction[]) {
  for (const a of actions) {
    switch (a.type) {
      case 'read':
        await mail.setFlags(acc, folder, uids, { seen: true });
        break;
      case 'unread':
        await mail.setFlags(acc, folder, uids, { seen: false });
        break;
      case 'flag':
        await mail.setFlags(acc, folder, uids, { flagged: true });
        break;
      case 'category':
        if (a.category) await mail.setKeywords(acc, folder, uids, [a.category], []);
        break;
      case 'pin':
        await mail.setKeywords(acc, folder, uids, [PIN_KEYWORD], []);
        break;
      case 'move':
        if (a.folder && a.folder !== folder) {
          await mail.moveMessages(acc, folder, uids, a.folder);
          return;
        }
        break;
      case 'archive':
        await mail.archiveMessages(acc, folder, uids);
        return;
      case 'delete':
        await mail.deleteMessages(acc, folder, uids);
        return;
    }
  }
}

/** Spustí pravidlá na zadané správy (nové v Doručenej pošte alebo „Spustiť teraz“). Vráti počet správ, na ktoré sa použili. */
export async function runRules(
  acc: AccountConfig,
  folder: string,
  rules: Rule[],
  msgs: Array<{ uid: number; from: string; to: string; subject: string }>,
): Promise<number> {
  const active = rules.filter((r) => r.enabled && (!r.accountId || r.accountId === acc.id) && r.actions.length);
  if (!active.length) return 0;
  let hits = 0;
  for (const m of msgs) {
    for (const r of active) {
      if (!ruleMatches(r, m)) continue;
      hits++;
      try {
        await applyActions(acc, folder, [m.uid], r.actions);
      } catch (e) {
        log(`pravidlo „${r.name}“ zlyhalo`, String((e as Error)?.message ?? e));
      }
      if (r.stop || r.actions.some((x) => ['move', 'archive', 'delete'].includes(x.type))) break;
    }
  }
  return hits;
}

/** „Spustiť pravidlá teraz“ na uložených správach priečinka. */
export async function runRulesOnFolder(acc: AccountConfig, folder: string, rules: Rule[]) {
  const list = cache.cachedList(acc.id, folder, 500);
  const msgs = list.map((m) => ({ uid: m.uid, from: `${m.fromName} ${m.fromAddress}`, to: cache.recipientsOf(acc.id, folder, m.uid), subject: m.subject }));
  return runRules(acc, folder, rules, msgs);
}
