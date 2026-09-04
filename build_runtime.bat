@echo off
REM ---- Package the separate Chromium runtime (rare — see runtime_manager.py) ----
REM Produces installer_output\OfficeTool-Runtime-<version>.zip — the ONE
REM thing this contains is bundled_browser\<version>\... (the real
REM chrome.exe Playwright needs for live-preview rendering), zipped so its
REM own top-level entry IS the "<version>" folder itself — that's the
REM exact layout runtime_manager.ensure_runtime() extracts straight into
REM place, matching what a build.bat-bundled copy already looked like.
REM
REM You only need to run this when Playwright's own Chromium build
REM actually changes (i.e. after `playwright install chromium` picks up a
REM newer one) — NOT on every ordinary app release. Most releases need
REM nothing here at all; see build.bat's own comment for why the two are
REM now separate.
where python >nul 2>nul || (echo Python is not installed. Get it from https://python.org & pause & exit /b)

if not exist bundled_browser\*.* (
  echo Populating bundled_browser\ from %LOCALAPPDATA%\ms-playwright ...
  for /d %%D in ("%LOCALAPPDATA%\ms-playwright\chromium-*") do (
    if exist "%%D\chrome-win64\chrome.exe" (
      mkdir "bundled_browser\%%~nxD" 2>nul
      xcopy /e /i /y /q "%%D" "bundled_browser\%%~nxD" >nul
    )
  )
)
set "FOUND_CHROME="
for /f "delims=" %%F in ('dir /s /b "bundled_browser\chrome.exe" 2^>nul') do set "FOUND_CHROME=%%F"
if not defined FOUND_CHROME (
  echo No local Chromium install found to package — run "playwright install chromium" first.
  pause & exit /b
)

set /p RUNTIME_VERSION=<RUNTIME_VERSION
if not exist "bundled_browser\%RUNTIME_VERSION%\chrome-win64\chrome.exe" (
  echo RUNTIME_VERSION says "%RUNTIME_VERSION%" but bundled_browser\%RUNTIME_VERSION%\ doesn't have a real chrome.exe.
  echo Either bump RUNTIME_VERSION to match what's actually in bundled_browser\, or delete bundled_browser\ and rerun this to repopulate it fresh.
  pause & exit /b
)

if not exist installer_output mkdir installer_output
set ZIP_PATH=installer_output\OfficeTool-Runtime-%RUNTIME_VERSION%.zip
del "%ZIP_PATH%" 2>nul
REM PowerShell's Compress-Archive, invoked from cmd — zips bundled_browser\
REM %RUNTIME_VERSION%\ itself as the archive's own top-level folder (not
REM its CONTENTS at the archive root), matching runtime_manager.py's
REM extractall() expectation exactly.
powershell -NoProfile -Command "Compress-Archive -Path 'bundled_browser\%RUNTIME_VERSION%' -DestinationPath '%ZIP_PATH%' -Force"

echo.
echo Runtime package: %ZIP_PATH%
echo.
echo ---- Publishing (only needed when this actually changed) ----
echo 1. On GitHub: Releases -^> Draft a new release, tag "runtime-%RUNTIME_VERSION%".
echo 2. Attach %ZIP_PATH% as the release asset.
echo 3. Publish. Every install whose local runtime doesn't already match
echo    picks this up automatically the next time it needs it.
pause
