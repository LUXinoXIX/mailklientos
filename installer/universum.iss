; Klasický inštalátor Windows (Inno Setup 6) pre Universum eMail – druhý inštalátor popri Squirrel (MailKlient-Setup.exe).
; Zostavuje sa v GitHub Actions: iscc /DAppVersion=0.9.9 /DSrcDir=..\out\MailKlient-win32-x64 /DOutDir=..\out\make\inno installer\universum.iss
; Pri inštalácii sa používateľ rozhodne: pre všetkých používateľov (Program Files, práva správcu) alebo len pre seba (bez správcu).

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#ifndef SrcDir
  #define SrcDir "..\out\MailKlient-win32-x64"
#endif
#ifndef OutDir
  #define OutDir "..\out\make\inno"
#endif
#define AppName "Universum eMail"
#define AppExe "MailKlient.exe"

[Setup]
; AppId sa nesmie meniť – podľa neho inštalátor spozná predchádzajúcu inštaláciu (aktualizácia na mieste)
AppId={{A687D649-544D-4238-9602-3C466699310D}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher=Universum eMail
AppPublisherURL=https://github.com/LUXinoXIX/mailklientos
AppSupportURL=https://github.com/LUXinoXIX/mailklientos
VersionInfoVersion={#AppVersion}
VersionInfoDescription={#AppName} – inštalátor
; {autopf} = Program Files pri inštalácii pre všetkých, %LOCALAPPDATA%\Programs pri inštalácii len pre seba
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
; predvolene bez správcu; dialóg ponúkne aj inštaláciu pre všetkých (Program Files)
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog commandline
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir={#OutDir}
OutputBaseFilename=UniversumEmail-Setup
SetupIconFile=..\assets\icon.ico
UninstallDisplayIcon={app}\{#AppExe}
UninstallDisplayName={#AppName}
WizardStyle=modern
Compression=lzma2/max
SolidCompression=yes
; bežiaci program (aj na pozadí pri hodinách) sa pred inštaláciou zatvorí a po nej znova spustí
CloseApplications=force
CloseApplicationsFilter=*.exe,*.dll,*.pak
RestartApplications=no
SetupMutex=UniversumEmailSetup
MinVersion=10.0

[Languages]
#ifexist "compiler:Languages\Slovak.isl"
Name: "sk"; MessagesFile: "compiler:Languages\Slovak.isl"
#endif
Name: "en"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Ikona na pracovnej ploche"; GroupDescription: "Ďalšie možnosti:"

[Files]
Source: "{#SrcDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; AppUserModelID rovnaké ako v programe – inak by Windows nezobrazoval upozornenia na novú poštu
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"; AppUserModelID: "com.squirrel.mailklient.MailKlient"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon; AppUserModelID: "com.squirrel.mailklient.MailKlient"

[Run]
Filename: "{app}\{#AppExe}"; Description: "Spustiť {#AppName}"; Flags: nowait postinstall skipifsilent
; po tichej aktualizácii (z programu) sa program spustí sám
Filename: "{app}\{#AppExe}"; Flags: nowait runasoriginaluser; Check: WizardSilent

[UninstallRun]
Filename: "{cmd}"; Parameters: "/C taskkill /IM {#AppExe} /F"; Flags: runhidden; RunOnceId: "KillApp"

; Nastavenia, kontá a heslá (%APPDATA%\MailKlient) sa pri odinštalovaní nemažú – zostanú pre novú inštaláciu.
