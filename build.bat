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
