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
SetupIconFile=icon.ico
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
; Does NOT include bundled_browser\ (the ~400MB Chromium payload) as of
; RUNTIME_VERSION/runtime_manager.py — see build.bat's own comment for
; why that's now a separate, rarely-changing download instead of part of
; every single app update.
Source: "dist\OfficeTool\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\OfficeTool\_internal\*"; DestDir: "{app}\_internal"; Flags: ignoreversion recursesubdirs createallsubdirs
; Anyone updating from a version that still shipped Chromium inside
; _internal\ (every version through v1.1.26) has a real, already-
; downloaded ~400MB sitting at {app}\_internal\bundled_browser\ — this is
; deliberately NOT deleted here. runtime_manager.py's own startup check
; MOVES it into the new {app}\runtime\ location instead of re-downloading
; it from scratch (and cleans up the old empty folder itself once
; that's done) — see its own module docstring. Forcibly deleting it here,
; before the app ever gets a chance to migrate it, would just force a
; redundant ~400MB re-download for every single existing install, which
; is exactly the problem this whole feature exists to avoid.

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; No skipifsilent — the in-app auto-updater (update_checker.py) runs this
; installer with /VERYSILENT so an update never shows the wizard at all,
; and expects the app to relaunch itself afterward automatically rather
; than leaving the user to go find and reopen it themselves. A normal
; interactive install still shows this as a "Launch now" checkbox on the
; wizard's finish page either way (postinstall's own purpose).
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName} now"; Flags: nowait postinstall
