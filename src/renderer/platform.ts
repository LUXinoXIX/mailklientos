// Rozdiely medzi Windows/Linux a macOS v rozhraní (popisy klávesových skratiek).
export const isMac = typeof window !== 'undefined' && (window as unknown as { mailApi?: { platform?: string } }).mailApi?.platform === 'darwin';

/**
 * Popis skratky podľa systému: na Macu Ctrl → ⌘ (Cmd). Výnimka Ctrl+Q: Cmd+Q na Macu ukončí program,
 * preto „označiť ako prečítané“ ostáva na klávese Control (⌃Q).
 */
export function kb(s: string): string {
  if (!isMac) return s;
  return s
    .replace(/Ctrl\+Q/g, '⌃Q')
    .replace(/Ctrl\+Shift\+/g, '⇧⌘')
    .replace(/Ctrl\+/g, '⌘')
    .replace(/Ctrl\b/g, '⌘');
}
