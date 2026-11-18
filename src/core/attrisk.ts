// Posúdenie rizika prílohy pred otvorením (bez Electronu – používa hlavný proces aj rozhranie).

/** Súbory, ktoré sa po otvorení môžu priamo spustiť – z programu sa neotvárajú, dajú sa len uložiť. */
const BLOCKED = new Set([
  'exe', 'com', 'scr', 'pif', 'bat', 'cmd', 'msi', 'msp', 'mst', 'js', 'jse', 'vbs', 'vbe', 'wsf', 'wsh', 'wsc',
  'ps1', 'psm1', 'psd1', 'ps1xml', 'hta', 'cpl', 'jar', 'lnk', 'reg', 'scf', 'inf', 'msc', 'gadget', 'application',
  'appx', 'appxbundle', 'msix', 'msixbundle', 'chm', 'dll', 'sys', 'drv', 'ocx', 'iso', 'img', 'vhd', 'vhdx',
  'url', 'website', 'settingcontent-ms', 'library-ms', 'search-ms', 'xll', 'xla', 'xlam', 'ppa', 'ppam', 'sh', 'command',
  'app', 'pkg', 'dmg', 'desktop', 'appimage', 'deb', 'rpm', 'py', 'pyw', 'pl', 'rb',
]);

/** Office s makrami a iné súbory, ktoré môžu obsahovať spustiteľný kód – otvoriť len s výslovným potvrdením. */
const MACRO = new Set(['docm', 'dotm', 'xlsm', 'xltm', 'xlsb', 'pptm', 'potm', 'ppsm', 'sldm', 'one', 'rtf', 'svg', 'html', 'htm', 'xhtml', 'mht', 'mhtml', 'xml']);

/** Prípony, za ktoré sa spustiteľný súbor často maskuje (faktura.pdf.exe). */
const DOCS = /\.(pdf|docx?|xlsx?|pptx?|txt|jpe?g|png|gif|zip|rar|7z|csv|odt|ods)$/i;

/** Obrázky, ktoré sa zobrazujú priamo v programe (bez spustenia iného programu). */
export const SAFE_IMAGE = /\.(png|jpe?g|gif|webp|bmp)$/i;

export type AttachmentRisk = {
  level: 'blocked' | 'macro' | 'normal';
  /** skutočná prípona (posledná), malými písmenami */
  ext: string;
  /** dôvody na zobrazenie používateľovi */
  reasons: string[];
};

export function attachmentRisk(filename: string): AttachmentRisk {
  const reasons: string[] = [];
  // znaky, ktoré otáčajú alebo skrývajú text (faktura‮fdp.exe sa zobrazí ako „fakturaexe.pdf“)
  const hidden = /[‪-‮⁦-⁩​-‏­]/.test(filename);
  const clean = filename.replace(/[‪-‮⁦-⁩​-‏­]/g, '').trim().replace(/[.\s]+$/, '');
  const ext = (/\.([a-z0-9-]{1,20})$/i.exec(clean)?.[1] ?? '').toLowerCase();
  if (hidden) reasons.push('Názov obsahuje skryté znaky, ktoré menia, ako sa názov zobrazuje – typický znak podvodu.');
  const base = clean.slice(0, clean.length - ext.length - 1);
  if (ext && DOCS.test(base) && BLOCKED.has(ext)) reasons.push(`Súbor sa tvári ako dokument (${base.split('.').pop()}), ale v skutočnosti je to .${ext}.`);
  if (/\s{4,}\.[a-z0-9]+$/i.test(filename)) reasons.push('Pred príponou je veľa medzier – prípona sa tak schová.');
  if (BLOCKED.has(ext)) {
    reasons.unshift(`Súbor .${ext} sa po otvorení môže sám spustiť a nainštalovať škodlivý program.`);
    return { level: 'blocked', ext, reasons };
  }
  if (hidden) return { level: 'blocked', ext, reasons };
  if (MACRO.has(ext)) {
    reasons.unshift(`Súbor .${ext} môže obsahovať makrá alebo skripty, ktoré sa spustia po otvorení.`);
    return { level: 'macro', ext, reasons };
  }
  return { level: 'normal', ext, reasons };
}
