@echo off
REM ---- Build the installable Office Tool .exe (Windows) ----
REM Produces dist\OfficeTool\OfficeTool.exe — a real native-window app (no
REM browser needed), built from the same app.py/engine.py/html_engine.py
REM this project runs in dev with `python app.py`.
REM
REM As of RUNTIME_VERSION/runtime_manager.py, this build does NOT bundle
REM bundled_browser\ (the ~400MB Chromium Playwright needs for live-
REM preview rendering) into the app installer at all anymore — per
REM explicit request: shipping a ~400MB payload that almost never changes
REM inside the SAME installer as the app's own code (which changes almost
REM every release) meant every single app update re-downloaded and
REM re-installed the whole browser for nothing. The app now downloads
REM that separately, ONCE, the first time it's needed (see
REM runtime_manager.py's own module docstring) — this build only needs
REM `playwright install chromium` locally so build_runtime.bat can package
REM THAT as its own separate, rarely-rebuilt release asset; the main app
REM build below only bundles RUNTIME_VERSION (a one-line text file naming
REM which Chromium build this app expects), not the Chromium files
REM themselves.
where python >nul 2>nul || (echo Python is not installed. Get it from https://python.org & pause & exit /b)
python -m pip install -r requirements.txt
python -m pip show pyinstaller >nul 2>nul || python -m pip install pyinstaller
REM --clean (plus deleting any leftover dist\build folders first) — found
REM the hard way: an incremental PyInstaller build can silently keep a
REM STALE cached copy of a module even after its source changed, with no
REM warning. That shipped a real regression once already (the live-
REM preview PDF fix in html_engine.py's _get_browser() sat in source for
REM most of a session while every incremental rebuild kept serving the
REM pre-fix cached version) — always building clean is the only way to be
REM sure what's actually in the .exe matches what's actually in source.
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
pyinstaller --name OfficeTool --noconfirm --clean --windowed ^
  --icon "icon.ico" ^
  --add-data "templates_html;templates_html" ^
  --add-data "templates;templates" ^
  --add-data "static;static" ^
  --add-data "VERSION;." ^
  --add-data "RUNTIME_VERSION;." ^
  --add-data "r2_readonly.json;." ^
  --add-data "google_oauth.json;." ^
  --add-data "branding;branding" ^
  --collect-all playwright ^
  --collect-all pystray ^
  app.py
echo.
echo Build complete: dist\OfficeTool\OfficeTool.exe
echo Portable use: copy that .exe (and its _internal folder) into THIS
echo project folder so it finds your existing config.json/drafts/submissions
echo here — no data migration needed, it just reads/writes right where it's put.
echo.
REM Real incident: a v1.1.79 release once shipped with v1.1.78 still
REM embedded inside dist\OfficeTool\_internal\VERSION — a stale dist\
REM folder from an EARLIER build sat there while the actual rebuild
REM silently no-op'd (a shell-invocation quirk, not this script's own
REM fault — bare "build.bat" failed to launch via one automation path
REM while ".\build.bat" worked), and every later step (payload zip,
REM installer, gh release upload) ran against that stale output without
REM any error, so the wrong app version shipped under the right tag. The
REM in-app Update Center kept re-showing "update available" after
REM install, which is how it surfaced. This check can't prevent every
REM way a build could go stale, but it catches the one symptom that
REM actually matters: the built package's own VERSION disagreeing with
REM the one everything else (git tag, release notes) was built from.
set /p SRC_VERSION=<VERSION
set /p BUILT_VERSION=<dist\OfficeTool\_internal\VERSION
if not "%SRC_VERSION%"=="%BUILT_VERSION%" (
  echo.
  echo *** BUILD ABORTED: version mismatch ***
  echo   Source VERSION file says: %SRC_VERSION%
  echo   Built package embeds:     %BUILT_VERSION%
  echo This means dist\OfficeTool\ is stale — the rebuild above didn't
  echo actually run against current source. Do NOT package or release
  echo this output. Delete build\ and dist\ and rerun this script.
  echo.
  pause
  exit /b 1
)
echo Version check OK — built package embeds v%BUILT_VERSION%, matches source.
echo.
echo ---- Packaging the payload (OfficeTool-Payload.zip) ----
REM Per explicit request ("make it that the installer will be very
REM light... hook all the data from github"): installer.iss no longer
REM bundles dist\OfficeTool\* into Setup.exe directly — it downloads THIS
REM zip from the matching GitHub release at install time instead (see
REM installer.iss's own top-of-file comment). Built here, every single
REM build, same as the .exe itself — it's the exact same
REM dist\OfficeTool\ contents, just zipped instead of embedded, so there's
REM nothing extra to keep in sync by hand.
REM
REM Compress-Archive (not 7-Zip/WinRAR) — already on every Windows dev
REM machine that can run this script at all, so no extra build-tool
REM install beyond what building the .exe itself already needed.
if not exist installer_output mkdir installer_output
del /q installer_output\OfficeTool-Payload.zip 2>nul
powershell -NoProfile -Command "Compress-Archive -Path 'dist\OfficeTool\*' -DestinationPath 'installer_output\OfficeTool-Payload.zip' -CompressionLevel Optimal"
if not exist installer_output\OfficeTool-Payload.zip (
  echo Payload zip failed to build — aborting before the installer step.
  pause
  exit /b 1
)
echo Payload: installer_output\OfficeTool-Payload.zip
echo IMPORTANT: `gh release create` must attach BOTH OfficeTool-Setup.exe
echo AND OfficeTool-Payload.zip to the same release tag — the installer has
echo nothing to install without the payload zip sitting on that same release.
echo.
echo ---- Building the installer (installer.iss) ----
REM Bump the VERSION file at the project root before shipping an update —
REM it's read here AND by the running app (version.py), so both always
REM agree; same AppId in installer.iss (never change that one) is what
REM makes re-running Setup.exe upgrade an existing install in place
REM instead of a side-by-side copy. After this, tag+release on GitHub —
REM see update_checker.py's top-of-file comment for the exact steps.
set /p APP_VERSION=<VERSION
set ISCC="%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if not exist %ISCC% set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if exist %ISCC% (
  %ISCC% /DMyAppVersion=%APP_VERSION% installer.iss
  echo Installer: installer_output\OfficeTool-Setup.exe ^(v%APP_VERSION%^)
) else (
  echo Inno Setup not found — install it first: winget install JRSoftware.InnoSetup
)
pause
