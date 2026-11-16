# MailKlient

Rich e-mailový klient pre Windows, macOS a Linux v štýle Microsoft 365 / Outlook.
IMAP/SMTP je predvolený typ konta (vrátane starých serverov Zimbra s neplatným certifikátom),
konektor MS365 (Entra ID / Graph) príde ako voliteľný modul.

## Stiahnutie

Najnovší inštalátor: [Releases → Najnovšie zostavenie](../../releases/tag/latest-build)

- **Windows:** `MailKlient-Setup.exe` (inštaluje sa bez práv správcu). Ak sa zobrazí
  „Windows ochránil váš počítač“, klikni na **Ďalšie informácie → Spustiť aj tak**
  (inštalátor zatiaľ nie je podpísaný).
- **macOS / Linux:** ZIP, resp. `.deb` (zatiaľ nepodpísané, testovacie).

Inštalátory zostavuje GitHub Actions pri každej zmene vo vetve `main`.

## Stav

- Hotové: rozloženie D (pás modulov, kontá, zoznam, čítací panel, 4. stĺpec), ribbon plný aj kompaktný,
  témy modrá / sivozelená / sivá, IMAP kontá (autodetekcia, SSL/TLS, STARTTLS, dôvera certifikátu,
  režim pre staré servery), čítanie, odoslanie a odpoveď.
- Ďalej: lokálna cache a fulltext, IDLE, mazanie a presun, prílohy, kalendár, úlohy, spúšťače, vložené weby.

## Vývoj

```
npm install
npm start
```

Štruktúra: `src/core` (TLS a typy, bez Electronu), `src/main` (kontá, IMAP/SMTP, IPC),
`src/preload.ts` (`window.mailApi`), `src/renderer` (React + Fluent UI).
