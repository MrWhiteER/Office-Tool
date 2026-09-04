r"""
Manages the Chromium runtime Playwright's live-preview rendering needs
(html_engine.py) as a package SEPARATE from the app's own installer —
per explicit request: "the update should update only the thing what's
changing... so we won't update things which are already ok and not
changed." The app's own code (app.py, templates, static assets) changes
almost every release; this ~400MB Chromium payload almost never does —
bundling both into the SAME installer meant every single app update
re-downloaded and re-installed the whole browser for nothing, every time.

---- How it works ----
- RUNTIME_VERSION (bundled with the app, tiny — like VERSION) names which
  Chromium build THIS app build expects, e.g. "chromium-1228". Deliberately
  matches Playwright's own build-folder naming exactly (see
  build_runtime.bat) so there's no separate version scheme to invent or
  keep in sync by hand.
- The actual runtime files live at engine.DATA_BASE/runtime/ —
  DELIBERATELY outside {app}\_internal\, which is the ONLY thing
  installer.iss's [Files] section for the main app ever replaces on an
  update (see installer.iss's own [Files] entries). Inno Setup never
  touches, and never deletes, a folder it doesn't know about, so once
  this is populated it survives every future app update untouched,
  unless RUNTIME_VERSION itself changes (i.e. Playwright's own Chromium
  build actually moved on).
- is_runtime_ready() is a cheap, no-network check (just two local file
  checks) — safe to call on every single startup. ensure_runtime() only
  does real work (a GitHub API call + a real download) the first time a
  given machine needs a given runtime version at all.

---- Publishing a new runtime build (rare — only when Playwright's own
     Chromium version actually changes, e.g. after a `playwright install
     chromium` picks up a newer build) ----
1. Run build_runtime.bat — zips up whatever's in bundled_browser\ (see
   build.bat's own comment for how that gets populated locally from
   %LOCALAPPDATA%\ms-playwright) into installer_output\
   OfficeTool-Runtime-<version>.zip, e.g. OfficeTool-Runtime-chromium-1228.zip.
2. On GitHub: Releases -> Draft a new release, tag "runtime-<version>"
   (must exactly match the real folder name under bundled_browser\, e.g.
   "runtime-chromium-1228" — this is how ensure_runtime() finds it, NOT
   the app's own "latest" release, which changes every app version and
   is a completely separate release train from this one), attach that zip.
3. Bump the RUNTIME_VERSION file at the project root to that same
   "<version>" string, commit, and ship the next app release as normal
   (build.bat/GUPDATE) — every install whose local runtime doesn't
   already match picks up the new one automatically, once, the same way
   any other update works, without needing its own separate update flow.
"""
import json
import os
import shutil
import tempfile
import urllib.request
import zipfile

import engine

GITHUB_REPO = "MrWhiteER/Office-Tool"
_HEADERS = {"Accept": "application/vnd.github+json", "User-Agent": "OfficeTool-RuntimeManager"}

# Outside {app}\_internal\ on purpose — see module docstring.
RUNTIME_DIR = os.path.join(engine.DATA_BASE, "runtime")
BUNDLED_BROWSER_DIR = os.path.join(RUNTIME_DIR, "bundled_browser")
INSTALLED_VERSION_FILE = os.path.join(RUNTIME_DIR, "INSTALLED_VERSION.txt")


def _read_expected_version():
    try:
        with open(os.path.join(engine.BASE, "RUNTIME_VERSION"), "r", encoding="utf-8") as f:
            return f.read().strip()
    except Exception:
        return ""


def _read_installed_version():
    try:
        with open(INSTALLED_VERSION_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    except Exception:
        return ""


def _legacy_bundled_browser_dir():
    """The OLD location, pre-runtime-split (every version through
    v1.1.26) — Chromium shipped inside {app}\\_internal\\bundled_browser as
    part of the app's own installer/[Files]. installer.iss deliberately
    leaves this in place on an update instead of deleting it (see its own
    comment) specifically so _migrate_legacy_runtime() below gets a
    chance to reuse it before anything is ever re-downloaded."""
    return os.path.join(engine.BASE, "bundled_browser")


def _migrate_legacy_runtime():
    """One-time, local-only move: if an existing install already has a
    real, matching Chromium sitting at the OLD pre-split location, MOVE
    it (not copy — this is ~400MB) into the new location instead of
    re-downloading identical bytes from GitHub. Safe to call on every
    startup — it no-ops instantly once already migrated, and does
    nothing at all on a fresh install or a dev checkout (neither ever
    has the legacy folder in the first place)."""
    expected = _read_expected_version()
    if not expected:
        return
    legacy_root = _legacy_bundled_browser_dir()
    legacy_chrome = os.path.join(legacy_root, expected, "chrome-win64", "chrome.exe")
    if not os.path.isfile(legacy_chrome):
        return  # nothing to migrate — ensure_runtime()'s normal download path handles it
    new_chrome = os.path.join(BUNDLED_BROWSER_DIR, expected, "chrome-win64", "chrome.exe")
    if os.path.isfile(new_chrome):
        # Already migrated (or separately downloaded) on a prior launch —
        # just clean up the stale old copy so it's not permanent dead weight.
        shutil.rmtree(legacy_root, ignore_errors=True)
        return
    try:
        os.makedirs(BUNDLED_BROWSER_DIR, exist_ok=True)
        dest = os.path.join(BUNDLED_BROWSER_DIR, expected)
        if os.path.isdir(dest):
            shutil.rmtree(dest, ignore_errors=True)
        shutil.move(os.path.join(legacy_root, expected), dest)
        with open(INSTALLED_VERSION_FILE, "w", encoding="utf-8") as f:
            f.write(expected)
    except Exception:
        return  # move failed for any reason — ensure_runtime()'s real download
                 # path is still there as a fallback; never block startup on this
    shutil.rmtree(legacy_root, ignore_errors=True)


def is_runtime_ready():
    """True if the local runtime cache already matches what this app
    build expects AND a real chrome.exe is actually sitting there — a
    cheap, local-only check (no network), safe to call on every launch."""
    expected = _read_expected_version()
    if not expected:
        return True  # dev checkout / no RUNTIME_VERSION shipped — nothing to enforce
    _migrate_legacy_runtime()
    if _read_installed_version() != expected:
        return False
    chrome_path = os.path.join(BUNDLED_BROWSER_DIR, expected, "chrome-win64", "chrome.exe")
    return os.path.isfile(chrome_path)


def ensure_runtime(on_progress=None):
    """Downloads + extracts the matching runtime .zip if not already
    present; a no-op (no network call at all) if it already is. Meant to
    run on a background thread — the caller (app.py's startup sequence)
    is responsible for that and for showing on_progress(done, total) to
    the user somehow, the same download-progress shape update_checker.py
    already uses for app updates. Raises on failure — deliberately not
    swallowed here, since silently proceeding without a real browser
    would just turn into "live preview mysteriously doesn't work" later
    with no explanation; the caller decides how to surface a real error."""
    if is_runtime_ready():
        return
    expected = _read_expected_version()
    if not expected:
        return
    tag = "runtime-" + expected
    api_url = "https://api.github.com/repos/{}/releases/tags/{}".format(GITHUB_REPO, tag)
    req = urllib.request.Request(api_url, headers=_HEADERS)
    with urllib.request.urlopen(req, timeout=15) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    asset_url = None
    for a in data.get("assets", []):
        if (a.get("name") or "").lower().endswith(".zip"):
            asset_url = a.get("browser_download_url")
            break
    if not asset_url:
        raise RuntimeError("No runtime package found for release tag " + tag)

    os.makedirs(RUNTIME_DIR, exist_ok=True)
    tmp_dir = tempfile.mkdtemp(prefix="officetool_runtime_")
    try:
        zip_path = os.path.join(tmp_dir, "runtime.zip")
        dl_req = urllib.request.Request(asset_url, headers={"User-Agent": "OfficeTool-RuntimeManager"})
        with urllib.request.urlopen(dl_req, timeout=30) as resp:
            total = int(resp.headers.get("Content-Length", 0) or 0)
            done = 0
            with open(zip_path, "wb") as f:
                while True:
                    chunk = resp.read(65536)
                    if not chunk:
                        break
                    f.write(chunk)
                    done += len(chunk)
                    if on_progress and total:
                        on_progress(done, total)
        # The zip's own top-level entry is the chromium-<n> folder itself
        # (see build_runtime.bat) — extracting straight into
        # BUNDLED_BROWSER_DIR reproduces the exact bundled_browser/
        # layout _get_browser() already expects, whether it came from
        # this download or the old bundled-at-build-time path.
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(BUNDLED_BROWSER_DIR)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    with open(INSTALLED_VERSION_FILE, "w", encoding="utf-8") as f:
        f.write(expected)
