// Konfigurácia Electron Forge (zostavenie a inštalátory).
// Zámerne v čistom JavaScripte, aby Forge nepotreboval ts-node.
const { FusesPlugin } = require('@electron-forge/plugin-fuses');
const { FuseV1Options, FuseVersion } = require('@electron/fuses');

module.exports = {
  packagerConfig: {
    asar: true,
    appBundleId: 'io.github.luxinoxix.mailklient',
    // ikona aplikácie (Windows .ico, Linux/macOS .png); zobrazovaný názov „Universum eMail“.
    // productName ostáva „MailKlient“ – podľa neho sa volá priečinok s nastaveniami a .exe (aktualizácie).
    icon: 'assets/icon',
    // ikona pre oznamovaciu oblasť (tray) – mimo asar, v priečinku resources
    extraResource: ['assets/tray.png'],
    win32metadata: {
      ProductName: 'Universum eMail',
      FileDescription: 'Universum eMail',
      CompanyName: 'Universum',
    },
  },
  rebuildConfig: {},
  makers: [
    // Windows: inštalátor bez práv správcu (inštaluje do %LocalAppData%).
    {
      name: '@electron-forge/maker-squirrel',
      platforms: ['win32'],
      config: {
        name: 'mailklient',
        title: 'Universum eMail',
        setupExe: 'MailKlient-Setup.exe',
        setupIcon: 'assets/icon.ico',
        iconUrl: 'https://raw.githubusercontent.com/LUXinoXIX/mailklientos/main/assets/icon.ico',
      },
    },
    // macOS a Linux: zatiaľ prenosný ZIP (podpis pre macOS príde neskôr).
    { name: '@electron-forge/maker-zip', platforms: ['darwin', 'linux'] },
    {
      name: '@electron-forge/maker-deb',
      platforms: ['linux'],
      config: { options: { name: 'mailklient', bin: 'MailKlient', productName: 'Universum eMail', icon: 'assets/icon.png' } },
    },
  ],
  plugins: [
    {
      name: '@electron-forge/plugin-vite',
      config: {
        build: [
          { entry: 'src/main.ts', config: 'vite.main.config.ts', target: 'main' },
          { entry: 'src/preload.ts', config: 'vite.preload.config.ts', target: 'preload' },
        ],
        renderer: [{ name: 'main_window', config: 'vite.renderer.config.ts' }],
      },
    },
    new FusesPlugin({
      version: FuseVersion.V1,
      [FuseV1Options.RunAsNode]: false,
      [FuseV1Options.EnableCookieEncryption]: true,
      [FuseV1Options.EnableNodeOptionsEnvironmentVariable]: false,
      [FuseV1Options.EnableNodeCliInspectArguments]: false,
      [FuseV1Options.EnableEmbeddedAsarIntegrityValidation]: true,
      [FuseV1Options.OnlyLoadAppFromAsar]: true,
    }),
  ],
};
