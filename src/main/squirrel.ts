// Udalosti inštalátora Squirrel (Windows): pri inštalácii, aktualizácii a odinštalovaní
// sa aplikácia spustí s prepínačom --squirrel-*. Vtedy len vytvoríme alebo odstránime
// zástupcov a skončíme (náhrada za balík electron-squirrel-startup, bez závislosti).
import { spawn } from 'node:child_process';
import path from 'node:path';
import { app } from 'electron';

/** Vráti true, ak ide o udalosť inštalátora a aplikácia sa má hneď ukončiť. */
export function handleSquirrelEvent(): boolean {
  if (process.platform !== 'win32') return false;
  const cmd = process.argv[1];
  if (!cmd || !cmd.startsWith('--squirrel-')) return false;

  const updateExe = path.resolve(path.dirname(process.execPath), '..', 'Update.exe');
  const exeName = path.basename(process.execPath);
  const runUpdate = (args: string[]) => {
    try {
      spawn(updateExe, args, { detached: true }).on('close', () => app.quit());
    } catch {
      app.quit();
    }
    setTimeout(() => app.quit(), 5000);
  };

  switch (cmd) {
    case '--squirrel-install':
    case '--squirrel-updated':
      runUpdate(['--createShortcut', exeName]);
      return true;
    case '--squirrel-uninstall':
      runUpdate(['--removeShortcut', exeName]);
      return true;
    case '--squirrel-obsolete':
      app.quit();
      return true;
    default:
      return false;
  }
}
