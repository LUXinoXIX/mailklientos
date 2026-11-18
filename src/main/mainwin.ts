// Hlavné okno aplikácie (okná správ a písania sú vedľajšie).
import { BrowserWindow } from 'electron';

let main: BrowserWindow | null = null;
export const setMainWindow = (w: BrowserWindow) => {
  main = w;
  w.on('closed', () => {
    if (main === w) main = null;
  });
};
export const getMainWindow = (): BrowserWindow | null =>
  main && !main.isDestroyed() ? main : BrowserWindow.getAllWindows().find((w) => !w.isDestroyed()) ?? null;
