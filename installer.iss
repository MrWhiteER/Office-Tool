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
; ---- Thin installer, per explicit request: "make it that the installer
; will be very light... make it so the installer will hook all the data
; from github" ----
; Setup.exe itself now carries NONE of the app's own files (no more
; ~90MB [Files] section) — it's just the small Inno Setup wizard shell.
; The [Code] section below downloads OfficeTool-Payload.zip (everything
; PyInstaller produced under dist\OfficeTool\ — the .exe + its whole
; _internal\ folder, zipped by build.bat) straight from THIS SAME
; release's GitHub assets during the "Installing" step, and extracts it
; into {app} via PowerShell's own Expand-Archive (built into every
; Windows 10/11 install — nothing extra to bundle for that either).
; Exactly the same "download the big rarely-bundled thing from a GitHub
; release instead of shipping it in the installer" idea runtime_manager.py
; already uses for the ~400MB Chromium payload (see that file's own
; docstring) — this just applies it to the app's OWN payload too, which
; is why Setup.exe can now stay small even though every release still
; changes app.py/templates/static.
;
; Publishing a release: build.bat produces BOTH installer_output\
; OfficeTool-Setup.exe (small, this script) AND installer_output\
; OfficeTool-Payload.zip (the actual app) — `gh release create` must
; attach BOTH to the same tag, or first installs/updates have nothing to
; download. update_checker.py's own asset picker only ever looks for a
; "...setup.exe"-named asset (see its own code) so the payload zip sitting
; alongside it on the same release never confuses that lookup.
;
; Only [Icons]/shortcuts are set up here now — nothing under [Files] is
; ever installed OR removed by the installer/uninstaller directly (the
; downloaded payload is copied by the [Code] section instead, but IS still
; removed on uninstall — see [UninstallDelete] below, since Inno doesn't
; automatically know about files it didn't itself place via [Files]).
; config.json/drafts/submissions/clients/finance are created by the APP
; itself at first run, not touched here, so they're untouched by an
; uninstall and preserved across every future update.

#define MyAppName "Office Tool"
#ifndef MyAppVersion
  #define MyAppVersion "1.0.0"
#endif
#define MyAppPublisher "Artemis Electricals Est"
#define MyAppExeName "OfficeTool.exe"
; Public repo — releases/download URLs work directly, no GitHub API call
; (and no rate limit/auth concern) needed just to fetch this one asset,
; unlike runtime_manager.py's Chromium lookup (that one's tag isn't the
; app's own version tag, so IT has to ask the API which release that is).
#define MyAppRepo "MrWhiteER/Office-Tool"
#define MyPayloadUrl "https://github.com/" + MyAppRepo + "/releases/download/v" + MyAppVersion + "/OfficeTool-Payload.zip"

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

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[UninstallDelete]
; The [Code] section below places these via a download+extract, not a
; [Files] entry — Inno only auto-removes what [Files] itself copied, so
; without this the app's own .exe/_internal\ (everything the payload zip
; unpacks) would survive an uninstall as orphaned dead weight. Does NOT
; touch {app}\runtime\ (the separate Chromium download) or any of the
; app's own data folders (config.json, drafts\, etc.) — same "only ever
; remove what was actually installed here" contract the old [Files]-based
; version had, just expressed as an explicit delete list instead of an
; implicit one.
Type: files; Name: "{app}\{#MyAppExeName}"
Type: filesandordirs; Name: "{app}\_internal"

[Run]
; No skipifsilent — the in-app auto-updater (update_checker.py) runs this
; installer with /VERYSILENT so an update never shows the wizard at all,
; and expects the app to relaunch itself afterward automatically rather
; than leaving the user to go find and reopen it themselves. A normal
; interactive install still shows this as a "Launch now" checkbox on the
; wizard's finish page either way (postinstall's own purpose).
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName} now"; Flags: nowait postinstall

[Code]
// Downloads OfficeTool-Payload.zip from this exact release (the URL is
// fully known at compile time — MyAppVersion is baked in via build.bat's
// /DMyAppVersion, so there's never a version mismatch between this
// installer and the payload it fetches) and extracts it straight into
// {app}. Runs once, during the normal "Installing" step, so both a fresh
// install and an in-place update look identical to the user — just a
// slightly longer "Installing..." page than before (network-bound now,
// not disk-bound) instead of a separate visible phase.
//
// PowerShell (Invoke-WebRequest + Expand-Archive), not a bundled unzip
// tool/plugin — both cmdlets ship with every Windows 10/11 install, so
// this adds zero bytes to Setup.exe itself. $ProgressPreference is
// silenced first: PowerShell's default download progress bar renders to
// the (hidden, since Setup runs it with SW_HIDE) console on every single
// buffer write, which is a well-known, large, pure-overhead slowdown on
// a file this size when nothing is even there to see it.
function DownloadAndExtractPayload(): Boolean;
var
  ResultCode: Integer;
  PSCommand: String;
  ZipPath: String;
begin
  ZipPath := ExpandConstant('{tmp}\OfficeTool-Payload.zip');
  PSCommand :=
    '$ProgressPreference=''SilentlyContinue''; ' +
    '[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; ' +
    'try { ' +
    '  Invoke-WebRequest -Uri ''{#MyPayloadUrl}'' -OutFile ''' + ZipPath + ''' -UseBasicParsing; ' +
    '  Expand-Archive -Path ''' + ZipPath + ''' -DestinationPath ''' + ExpandConstant('{app}') + ''' -Force; ' +
    '  exit 0 ' +
    '} catch { ' +
    '  Write-Error $_.Exception.Message; exit 1 ' +
    '}';
  Result := Exec('powershell.exe',
    '-NoProfile -ExecutionPolicy Bypass -Command "' + PSCommand + '"',
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode) and (ResultCode = 0);
  DeleteFile(ZipPath);
  // Belt-and-braces: even a ResultCode=0 PowerShell exit is only trusted
  // once the one file every later step actually needs is confirmed to
  // really be sitting there — a partial/corrupt extract (a truncated
  // download that still unzipped some files before failing partway, say)
  // should fail loudly here, not surface later as a baffling "the app
  // won't open" with no obvious cause.
  if Result then
    Result := FileExists(ExpandConstant('{app}\{#MyAppExeName}'));
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  Page: TOutputProgressWizardPage;
begin
  if CurStep = ssInstall then
  begin
    Page := CreateOutputProgressPage('Downloading Office Tool', 'Please wait while the application files are downloaded…');
    Page.Show;
    try
      Page.SetProgress(0, 1);
      if not DownloadAndExtractPayload() then
      begin
        MsgBox('Could not download the application files from GitHub.' + #13#10 + #13#10 +
          'Check your internet connection and try running Setup again. If this keeps happening, ' +
          'the release may be missing its OfficeTool-Payload.zip asset.',
          mbCriticalError, MB_OK);
        Abort();
      end;
      Page.SetProgress(1, 1);
    finally
      Page.Hide;
    end;
  end;
end;
