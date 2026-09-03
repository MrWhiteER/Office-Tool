@echo off
REM ---- Build the installable Office Tool .exe (Windows) ----
REM Produces dist\OfficeTool\OfficeTool.exe — a real native-window app (no
REM browser needed), built from the same app.py/engine.py/html_engine.py
REM this project runs in dev with `python app.py`.
REM
REM Requires `playwright install chromium` to have been run at least once on
REM THIS machine already (the .exe uses that shared browser cache rather
REM than bundling its own copy of Chromium — see html_engine.py's own
REM comment on PLAYWRIGHT_BROWSERS_PATH for why).
where python >nul 2>nul || (echo Python is not installed. Get it from https://python.org & pause & exit /b)
python -m pip install -r requirements.txt
python -m pip show pyinstaller >nul 2>nul || python -m pip install pyinstaller
pyinstaller --name OfficeTool --noconfirm ^
  --add-data "templates_html;templates_html" ^
  --add-data "templates;templates" ^
  --add-data "static;static" ^
  --add-data "VERSION;." ^
  --add-data "r2_readonly.json;." ^
  --collect-all playwright ^
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
