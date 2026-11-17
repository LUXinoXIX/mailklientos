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

## Stav (verzia 0.4)

- **Rozloženie:** pás kont úplne vľavo (poradie ťahaním, pravé tlačidlo = nastavenia), pás modulov podľa konta,
  horný pruh vo farbe konta, pravý panel (tlačidlo vpravo hore).
- **IMAP kontá:** plnohodnotný klient: priečinky, zoznam, čítanie, nová správa, odpovedať / všetkým, preposlať,
  prílohy (otvoriť, uložiť, priložiť), odstrániť, archivovať, presunúť, prečítané / neprečítané, vlajka,
  počty neprečítaných, „Všetky doručené“, podpisy pre každé konto, automatické nastavenie servera
  (Autoconfig a databáza Thunderbirdu), SSL/TLS, STARTTLS, výnimky certifikátov, režim pre staré Zimbry.
- **Písanie správ ako v Outlooku:** odpoveď a nová správa priamo v čítacom paneli, formátovaný text
  (písmo, farby, zoznamy, zarovnanie, odkazy, obrázky aj cez Ctrl+V, tabuľky), karty Správa a Vložiť v ribbone,
  príjemcovia s dopĺňaním, citácia ako v Outlooku, koncepty sa ukladajú priebežne do Konceptov.
- **Rýchlosť:** jedno trvalé spojenie na konto, otvorené správy sa pamätajú, pri načítaní žiadne biele plochy.
- **Kontá Microsoft 365 / Outlook.com:** vložený Outlook na webe ako celé okno, vlastné prihlásenie pre každé konto;
  Kalendár, To Do, Teams, OneDrive, Word, Excel, PowerPoint v kartách (aj keď ich otvoríš cez prepínač aplikácií
  v Outlooku); tlačidlo Edge.
- **Pravý panel:** vložené weby (Copilot, WhatsApp, Teams, vlastná adresa) a spúšťače programov (.exe) s ikonami.
- Ďalej: kalendár, úlohy a kontakty pre IMAP (CalDAV/CardDAV), lokálna cache a fulltext, okamžité doručovanie (IDLE).

## Vývoj

```
npm install
npm start
```

Štruktúra: `src/core` (TLS a typy, bez Electronu), `src/main` (kontá, IMAP/SMTP, IPC),
`src/preload.ts` (`window.mailApi`), `src/renderer` (React + Fluent UI).
