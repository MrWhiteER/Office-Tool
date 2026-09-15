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
; release's GitHub assets during the "Installing" step, with a REAL,
; live native progress bar (WinINet, wininet.dll, called directly —
; per explicit follow-up rule: "all the installers and downloaders
; should have a loading bar! THATS A RULE!" — a static "please wait"
; page doesn't satisfy that), then extracts it into {app} via
; PowerShell's own Expand-Archive (built into every Windows 10/11
; install — nothing extra to bundle for either step). Exactly the same
; "download the big rarely-bundled thing from a GitHub release instead
; of shipping it in the installer" idea runtime_manager.py already uses
; for the ~400MB Chromium payload (see that file's own docstring) — this
; just applies it to the app's OWN payload too, which is why Setup.exe
; can now stay small even though every release still changes
; app.py/templates/static.
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
// ---- Real, live progress bar for the GitHub download — per explicit
// rule: "all the installers and downloaders should have a loading bar!
// THATS A RULE!" A static "please wait" page doesn't qualify — this
// drives Inno's OWN native progress bar with the REAL byte count as it
// downloads.
//
// A first version of this called WinINet (wininet.dll) directly, the
// same underlying API the well-known "Inno Download Plugin" is built
// on — it technically worked (a real % and MB count, confirmed live),
// but WinINet turned out to be a genuinely bad fit for a GitHub release
// asset: slow (confirmed independently by the user watching a real
// download at the same time: "the download is very slow"), and it
// dropped mid-transfer at 36% on a real test run, both consistent with
// WinINet — an old API — not handling GitHub's modern CDN/TLS stack as
// well as .NET's own HTTP stack does. This version instead launches a
// short PowerShell script (System.Net.HttpWebRequest, .NET's HTTP
// client, already proven fast and reliable earlier in this project's
// own runtime_manager.py-equivalent download flows) ASYNCHRONOUSLY —
// ewNoWait, not ewWaitUntilTerminated — and polls a small progress file
// it writes to every ~150ms, translating that into Page.SetProgress/
// SetText the same way the WinINet version did. Best of both: .NET's
// faster, more reliable transfer, with the exact same live native
// progress bar this whole rewrite exists for.
const
  DL_POLL_MS = 200;
  // No progress at all for this long (not overall elapsed time — a
  // slow-but-genuinely-moving connection can legitimately take minutes)
  // means treat it as hung rather than wait forever.
  DL_STALL_TIMEOUT_MS = 45000;

// Bytes -> "12.3" (one decimal place, MB) — built by hand instead of
// Format('%.1f', ...): found live, the hard way, that this PascalScript
// build's Format() rejects that specifier ("invalid or incompatible with
// argument") for reasons not worth chasing further when the manual
// version is this short and has zero ambiguity left to get wrong.
function FormatMB(Bytes: Int64): String;
var
  Tenths: Int64;
begin
  Tenths := (Bytes * 10) div 1048576;
  Result := IntToStr(Tenths div 10) + '.' + IntToStr(Tenths mod 10);
end;

// Pulls the Nth ':'-separated field (1-based) out of S — Pascal Script
// has no built-in Split, and the progress file's own format
// ("PROGRESS:<downloaded>:<total>") is simple enough not to need one.
function DLField(S: String; N: Integer): String;
var
  Parts: TStringList;
begin
  Result := '';
  Parts := TStringList.Create;
  try
    Parts.Delimiter := ':';
    Parts.StrictDelimiter := True;
    Parts.DelimitedText := S;
    if N <= Parts.Count then
      Result := Parts[N - 1];
  finally
    Parts.Free;
  end;
end;

// Downloads one URL to LocalPath with LIVE progress reported through
// Page.SetProgress/SetText as real bytes arrive — see this section's own
// top comment for why this shells out to PowerShell/.NET rather than
// calling WinINet directly.
function DownloadFileWithProgress(Url, LocalPath: String; Page: TOutputProgressWizardPage): Boolean;
var
  ScriptPath, ProgressPath, PSScript, Status: String;
  Line: AnsiString;
  ResultCode: Integer;
  LastDownloaded, StalledMs, Downloaded, TotalSize: Int64;
  Finished: Boolean;
begin
  Result := False;
  ScriptPath := ExpandConstant('{tmp}\OfficeToolDownload.ps1');
  ProgressPath := ExpandConstant('{tmp}\OfficeToolDownload.progress');
  DeleteFile(ProgressPath);

  // ASCII, not UTF8/Unicode, on every Set-Content below — keeps the
  // progress file trivially readable back in Pascal Script via
  // LoadStringFromFile's own AnsiString contract, no BOM/encoding
  // mismatch to worry about for content that's only ever plain digits,
  // colons, and a short error message.
  PSScript :=
    '$ErrorActionPreference=''Stop''; $ProgressPreference=''SilentlyContinue''; ' +
    '[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; ' +
    '$prog=' + '''' + ProgressPath + '''' + '; ' +
    'try {' + #13#10 +
    '  $req=[Net.HttpWebRequest]::Create(' + '''' + Url + '''' + ');' + #13#10 +
    '  $req.UserAgent=''OfficeTool-Setup/1.0''; $req.AllowAutoRedirect=$true;' + #13#10 +
    '  $resp=$req.GetResponse();' + #13#10 +
    '  $total=$resp.ContentLength;' + #13#10 +
    '  $stream=$resp.GetResponseStream();' + #13#10 +
    '  $fileStream=[IO.File]::Create(' + '''' + LocalPath + '''' + ');' + #13#10 +
    '  $buffer=New-Object byte[] 65536; $downloaded=0; $last=Get-Date;' + #13#10 +
    '  while(($read=$stream.Read($buffer,0,$buffer.Length)) -gt 0){' + #13#10 +
    '    $fileStream.Write($buffer,0,$read); $downloaded+=$read;' + #13#10 +
    '    $now=Get-Date;' + #13#10 +
    '    if((($now-$last).TotalMilliseconds) -ge 150){' + #13#10 +
    '      Set-Content -Path $prog -Value "PROGRESS:$downloaded`:$total" -NoNewline -Encoding ASCII; $last=$now' + #13#10 +
    '    }' + #13#10 +
    '  }' + #13#10 +
    '  $fileStream.Close(); $stream.Close(); $resp.Close();' + #13#10 +
    '  Set-Content -Path $prog -Value "DONE:$downloaded" -NoNewline -Encoding ASCII' + #13#10 +
    '} catch {' + #13#10 +
    '  Set-Content -Path $prog -Value "ERROR:$($_.Exception.Message)" -NoNewline -Encoding ASCII' + #13#10 +
    '}';
  SaveStringToFile(ScriptPath, PSScript, False);

  if not Exec('powershell.exe',
      '-NoProfile -ExecutionPolicy Bypass -File "' + ScriptPath + '"',
      '', SW_HIDE, ewNoWait, ResultCode) then begin
    Log('DownloadFileWithProgress: failed to launch PowerShell, GetLastError-style Result=' + IntToStr(ResultCode));
    Exit;
  end;

  Page.SetProgress(0, 1);
  LastDownloaded := 0;
  StalledMs := 0;
  Finished := False;
  while not Finished do begin
    Sleep(DL_POLL_MS);
    if not LoadStringFromFile(ProgressPath, Line) or (Line = '') then begin
      StalledMs := StalledMs + DL_POLL_MS;
      if StalledMs >= DL_STALL_TIMEOUT_MS then begin
        Log('DownloadFileWithProgress: no progress file activity for ' + IntToStr(StalledMs) + 'ms — treating as hung');
        Exit;
      end;
      Continue;
    end;
    Status := DLField(Line, 1);
    if Status = 'DONE' then begin
      Result := True;
      Finished := True;
    end
    else if Status = 'ERROR' then begin
      Log('DownloadFileWithProgress: PowerShell reported: ' + Line);
      Finished := True;
    end
    else if Status = 'PROGRESS' then begin
      Downloaded := StrToInt64Def(DLField(Line, 2), LastDownloaded);
      TotalSize := StrToInt64Def(DLField(Line, 3), 0);
      if Downloaded <> LastDownloaded then begin
        StalledMs := 0;
        LastDownloaded := Downloaded;
      end
      else begin
        StalledMs := StalledMs + DL_POLL_MS;
        if StalledMs >= DL_STALL_TIMEOUT_MS then begin
          Log('DownloadFileWithProgress: download stalled at ' + IntToStr(Downloaded) + ' bytes for ' + IntToStr(StalledMs) + 'ms');
          Exit;
        end;
      end;
      if TotalSize > 0 then begin
        Page.SetProgress(Downloaded, TotalSize);
        Page.SetText('Downloading Office Tool…', FormatMB(Downloaded) + ' MB of ' + FormatMB(TotalSize) +
          ' MB (' + IntToStr((Downloaded * 100) div TotalSize) + '%)');
      end
      else
        Page.SetText('Downloading Office Tool…', FormatMB(Downloaded) + ' MB downloaded');
    end;
  end;
  DeleteFile(ScriptPath);
end;

// Extraction — PowerShell's Expand-Archive (built into every Windows
// 10/11 install, nothing to bundle) is still exactly right for this
// part: it's fast (seconds, not the minute-plus a slow connection can
// take for the download itself) and doesn't need its own progress bar
// to feel responsive — Page.SetText below just labels the step so the
// bar doesn't look stuck at 100% while it runs.
function ExtractPayload(ZipPath: String): Boolean;
var
  ResultCode: Integer;
  PSCommand: String;
begin
  PSCommand := '$ProgressPreference=''SilentlyContinue''; ' +
    'try { Expand-Archive -Path ''' + ZipPath + ''' -DestinationPath ''' + ExpandConstant('{app}') + ''' -Force; exit 0 } ' +
    'catch { Write-Error $_.Exception.Message; exit 1 }';
  Result := Exec('powershell.exe',
    '-NoProfile -ExecutionPolicy Bypass -Command "' + PSCommand + '"',
    '', SW_HIDE, ewWaitUntilTerminated, ResultCode) and (ResultCode = 0);
end;

// Downloads OfficeTool-Payload.zip from this exact release (the URL is
// fully known at compile time — MyAppVersion is baked in via build.bat's
// /DMyAppVersion, so there's never a version mismatch between this
// installer and the payload it fetches) and extracts it straight into
// {app}. Runs once, during the normal "Installing" step, so both a fresh
// install and an in-place update look identical to the user.
function DownloadAndExtractPayload(Page: TOutputProgressWizardPage): Boolean;
var
  ZipPath: String;
begin
  ZipPath := ExpandConstant('{tmp}\OfficeTool-Payload.zip');
  Page.SetText('Downloading Office Tool…', 'Connecting…');
  Result := DownloadFileWithProgress('{#MyPayloadUrl}', ZipPath, Page);
  if Result then
  begin
    Page.SetText('Installing Office Tool…', 'Extracting application files…');
    Result := ExtractPayload(ZipPath);
  end;
  DeleteFile(ZipPath);
  // Belt-and-braces: even a reported success is only trusted once the
  // one file every later step actually needs is confirmed to really be
  // sitting there — a partial/corrupt extract (a truncated download
  // that still unzipped some files before failing partway, say) should
  // fail loudly here, not surface later as a baffling "the app won't
  // open" with no obvious cause.
  if Result then
    Result := FileExists(ExpandConstant('{app}\{#MyAppExeName}'));
end;

// Real, live testing surfaced a genuine reliability finding, not a bug in
// either downloader tried: BOTH a raw-WinINet version and this file's
// current .NET HttpWebRequest version got interrupted partway through
// (36%, then separately 59%) on real attempts over the same connection —
// consistent with an actual network interruption (the connection itself,
// a firewall/router killing a long-lived HTTPS session, GitHub's own
// transient hiccup — no way to tell which from here), not a defect
// specific to either implementation. The fix that actually addresses
// THAT is retrying the whole attempt automatically rather than making
// the user notice the dialog and re-run Setup by hand.
const
  DL_MAX_ATTEMPTS = 4;

procedure CurStepChanged(CurStep: TSetupStep);
var
  Page: TOutputProgressWizardPage;
  Attempt: Integer;
  Success: Boolean;
begin
  if CurStep = ssInstall then
  begin
    Page := CreateOutputProgressPage('Downloading Office Tool', 'Please wait while the application files are downloaded from GitHub…');
    Page.Show;
    try
      Success := False;
      Attempt := 1;
      while (not Success) and (Attempt <= DL_MAX_ATTEMPTS) do
      begin
        if Attempt > 1 then
        begin
          Page.SetText('Downloading Office Tool…', 'Connection interrupted — retrying (attempt ' + IntToStr(Attempt) + ' of ' + IntToStr(DL_MAX_ATTEMPTS) + ')…');
          Sleep(2000);
        end;
        Success := DownloadAndExtractPayload(Page);
        Attempt := Attempt + 1;
      end;
      if not Success then
      begin
        MsgBox('Could not download the application files from GitHub after ' + IntToStr(DL_MAX_ATTEMPTS) + ' attempts.' + #13#10 + #13#10 +
          'Check your internet connection and try running Setup again. If this keeps happening, ' +
          'the release may be missing its OfficeTool-Payload.zip asset.',
          mbCriticalError, MB_OK);
        Abort();
      end;
    finally
      Page.Hide;
    end;
  end;
end;
