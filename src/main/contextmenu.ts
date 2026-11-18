// Systémové menu na pravý klik pre text, polia na písanie, odkazy a obrázky
// (aj v tele správy a vo vložených weboch). Vlastné menu (správy, priečinky) robí rozhranie.
import { clipboard, Menu, shell, type MenuItemConstructorOptions, type WebContents } from 'electron';

export function attachContextMenu(wc: WebContents, opts: { web?: boolean } = {}) {
  wc.on('context-menu', (_e, p) => {
    const t: MenuItemConstructorOptions[] = [];
    if (p.misspelledWord) {
      for (const s of p.dictionarySuggestions.slice(0, 6)) t.push({ label: s, click: () => wc.replaceMisspelling(s) });
      if (!p.dictionarySuggestions.length) t.push({ label: 'Žiadne návrhy', enabled: false });
      t.push({ label: 'Pridať do slovníka', click: () => wc.session.addWordToSpellCheckerDictionary(p.misspelledWord) });
      t.push({ type: 'separator' });
    }
    if (p.linkURL) {
      const mail = p.linkURL.startsWith('mailto:');
      if (!mail) t.push({ label: 'Otvoriť odkaz v prehliadači', click: () => void shell.openExternal(p.linkURL) });
      t.push({ label: mail ? 'Kopírovať e-mailovú adresu' : 'Kopírovať odkaz', click: () => clipboard.writeText(mail ? p.linkURL.slice(7).split('?')[0]! : p.linkURL) });
      t.push({ type: 'separator' });
    }
    if (p.mediaType === 'image' && p.srcURL) {
      t.push({ label: 'Kopírovať obrázok', click: () => wc.copyImageAt(p.x, p.y) });
      if (/^https?:/.test(p.srcURL)) t.push({ label: 'Uložiť obrázok…', click: () => wc.downloadURL(p.srcURL) });
      t.push({ type: 'separator' });
    }
    if (p.isEditable) {
      t.push(
        { label: 'Späť', role: 'undo', enabled: p.editFlags.canUndo },
        { type: 'separator' },
        { label: 'Vystrihnúť', role: 'cut', enabled: p.editFlags.canCut },
        { label: 'Kopírovať', role: 'copy', enabled: p.editFlags.canCopy },
        { label: 'Prilepiť', role: 'paste', enabled: p.editFlags.canPaste },
        { label: 'Prilepiť ako čistý text', role: 'pasteAndMatchStyle', enabled: p.editFlags.canPaste },
        { type: 'separator' },
        { label: 'Vybrať všetko', role: 'selectAll' },
      );
    } else if (p.selectionText.trim()) {
      t.push({ label: 'Kopírovať', role: 'copy' }, { label: 'Vybrať všetko', role: 'selectAll' });
      const q = p.selectionText.trim().slice(0, 100);
      t.push({ label: `Hľadať na webe „${q.length > 30 ? `${q.slice(0, 30)}…` : q}“`, click: () => void shell.openExternal(`https://www.bing.com/search?q=${encodeURIComponent(q)}`) });
    } else if (opts.web) {
      t.push(
        { label: 'Späť', enabled: wc.navigationHistory.canGoBack(), click: () => wc.navigationHistory.goBack() },
        { label: 'Dopredu', enabled: wc.navigationHistory.canGoForward(), click: () => wc.navigationHistory.goForward() },
        { label: 'Obnoviť', click: () => wc.reload() },
        { type: 'separator' },
        { label: 'Kopírovať adresu stránky', click: () => clipboard.writeText(wc.getURL()) },
        { label: 'Otvoriť v prehliadači', click: () => void shell.openExternal(wc.getURL()) },
      );
    } else {
      t.push({ label: 'Vybrať všetko', role: 'selectAll' });
    }
    while (t.length && t[t.length - 1]!.type === 'separator') t.pop();
    if (t.length) Menu.buildFromTemplate(t).popup();
  });
}
