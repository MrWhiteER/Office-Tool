; Office Tool — Inno Setup installer script.
;
; AppId is a fixed GUID (never change it) — it's how Windows/Inno Setup
; recognise a newer Setup.exe as an UPDATE to an existing install (same
; entry in "Apps & Features", "Modify/Repair/Uninstall" offered instead of
; a duplicate side-by-side install) rather than a completely different
; program. To ship an update: bump the VERSION file at the project root,
; rerun build.bat, hand out the new Setup.exe — running it over an
; existing install upgrades in place.
;
; MyAppVersion below is only a fallback for running ISCC directly on this
; file. build.bat instead passes /DMyAppVersion=<contents of VERSION> on
; the command line so the installer's version always matches the running
; app's own version.py (which reads that same VERSION file) — one number,
; never edited in two places.
;
; Installs per-user (DefaultDirName under LocalAppData, PrivilegesRequired
; lowest) — no admin/UAC prompt, and critically: the app WRITES its own
; data (config.json, drafts/, submissions/, clients/, finance/, generated
; caches) directly into its own install folder at runtime (see engine.py's
; DATA_BASE), which a normal user cannot do inside Program Files without
; elevation. Per-user install sidesteps that entirely.
;
; Only [Files] entries below are ever installed OR removed by the
; installer/uninstaller — config.json/drafts/submissions/clients/finance
; are created by the APP itself at first run, not listed here, so they are
; untouched by an uninstall and preserved across every future update.

#define MyAppName "Office Tool"
#ifndef MyAppVersion
  #define MyAppVersion "1.0.0"
#endif
#define MyAppPublisher "Artemis Electricals Est"
#define MyAppExeName "OfficeTool.exe"

[Setup]
; Kept the SAME GUID from before this rename — renaming the product does
; not need a new AppId, and changing it would break update-in-place for
; anyone who already installed under the old name.
AppId={{4C6E9A5E-6B2B-4B9E-9C7A-8E7B1E7F2A11}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\OfficeTool
DefaultGroupName={#MyAppName}
PrivilegesRequired=lowest
DisableProgramGroupPage=yes
OutputDir=installer_output
OutputBaseFilename=OfficeTool-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}
; Update installs run just as smoothly as a first install — same wizard,
; same "Next, Next, Install" — Inno Setup handles "already installed at
; this version/newer" detection on its own via AppId+AppVersion above.
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &Desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Files]
; Everything PyInstaller produced (the .exe + its _internal runtime/
; bundled-resources folder) — the ONLY files this installer ever manages.
Source: "dist\OfficeTool\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\OfficeTool\_internal\*"; DestDir: "{app}\_internal"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName} now"; Flags: nowait postinstall skipifsilent
