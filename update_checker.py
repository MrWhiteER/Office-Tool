r"""
Auto-update checker — polls GitHub's free public Releases API for a version
newer than the one currently running, and can fetch + launch that release's
installer on the user's say-so. No paid service, no account, no API key:
GitHub's REST API is free and unauthenticated for public repos.

---- One-time setup, once the project is pushed to GitHub ----
1. Set GITHUB_REPO below to "yourusername/Office-Tool".
2. Create your first Release on GitHub, tagged "v1.0.0" (must match VERSION
   at the repo root — see version.py), with installer_output\
   OfficeTool-Setup.exe (built by build.bat) attached as the release asset.

---- Shipping every future update ----
1. Bump the VERSION file (e.g. "1.0.1").
2. Rerun build.bat — it rebuilds the .exe AND the installer with that
   version baked in.
3. On GitHub: Releases -> Draft a new release, tag "v1.0.1", attach the new
   installer_output\OfficeTool-Setup.exe.
That's it — every copy of the app already running polls this API on its own
(see checkForAppUpdate() in app.py's page script) and will offer the update
automatically next time it's open, no manual download for the user.
"""
import json
import os
import re
import subprocess
import tempfile
import threading
import time
import urllib.request

import engine
from version import APP_VERSION

GITHUB_REPO = "MrWhiteER/Office-Tool"

_API_URL = "https://api.github.com/repos/{}/releases/latest".format(GITHUB_REPO)
_HEADERS = {"Accept": "application/vnd.github+json", "User-Agent": "OfficeTool-UpdateChecker"}


def _parse_version(v):
    """'v1.2.3' / '1.2' / etc -> (1, 2, 3) so tuples compare correctly."""
    nums = re.findall(r"\d+", v or "")
    nums = [int(n) for n in nums[:3]]
    while len(nums) < 3:
        nums.append(0)
    return tuple(nums)


def check_for_update(timeout=6):
    """
    Hits GitHub's latest-release endpoint once. Returns a JSON-safe dict:
      {available, current, latest, download_url, notes, page_url}
    or {available: False, error: "..."} if the repo isn't set up yet, the
    machine is offline, or GitHub is unreachable — callers should treat
    that as "no update", not as a hard failure.
    """
    if "YOUR_GITHUB_USERNAME" in GITHUB_REPO:
        return {"available": False, "current": APP_VERSION, "error": "GITHUB_REPO not configured yet"}
    try:
        req = urllib.request.Request(_API_URL, headers=_HEADERS)
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        latest_tag = data.get("tag_name", "") or ""
        latest = _parse_version(latest_tag)
        current = _parse_version(APP_VERSION)
        asset_url = None
        for a in data.get("assets", []):
            name = (a.get("name") or "").lower()
            if name.endswith("setup.exe"):
                asset_url = a.get("browser_download_url")
                break
        return {
            "available": latest > current and asset_url is not None,
            "current": APP_VERSION,
            "latest": latest_tag.lstrip("vV") or APP_VERSION,
            "download_url": asset_url,
            "notes": (data.get("body") or "").strip(),
            "page_url": data.get("html_url", ""),
        }
    except Exception as e:
        return {"available": False, "current": APP_VERSION, "error": str(e)}


def download_and_launch_installer(download_url, on_progress=None):
    """
    Downloads the release's Setup.exe to a fresh temp folder and runs it
    with /SILENT /SUPPRESSMSGBOXES /NORESTART — no wizard, no "Welcome to
    Setup" screen or license/directory pages, just Inno Setup's own small
    native progress bar while it works (see installer.iss's [Run] —
    skipifsilent was removed specifically so a silent run still reopens
    the app afterward on its own). Same AppId as the current install
    (installer.iss) means this is always an in-place upgrade, never a
    fresh install, so config.json/drafts/etc. survive exactly as before.

    /SILENT, not /VERYSILENT: confirmed directly that once Chromium got
    bundled into the installer (v1.1.6), the actual file-copy step alone
    can take several minutes (a real, timed run: ~290s) — and this app's
    own process has to exit before that starts (Windows won't let the
    installer overwrite files this process still has open), so with
    /VERYSILENT there'd be several minutes of NOTHING visible at all:
    the app just closes, and nothing reappears for a long, silent
    stretch that reads exactly like a failure (a real user hit this and
    reported it as one). /SILENT's small native progress window is a
    minor step back from "never looks like installing software," but a
    long silent gap with zero feedback is a worse experience than a
    small progress bar — this is the actual tradeoff, not a preference.

    /DIR= pins the target explicitly to THIS running instance's own
    folder (engine.DATA_BASE) rather than trusting Inno Setup's registry
    lookup alone — matters if this is ever a portable copy rather than a
    real tracked installation (AppId-based upgrade-detection only works
    for a real install; without /DIR a portable copy's "update" would
    silently install a SEPARATE fresh copy to the default location
    instead of updating the one actually running, which is a much worse
    version of "feels like new software").

    Caller is expected to exit this app shortly after this returns so the
    installer isn't blocked trying to close a running instance of it.
    """
    tmp_dir = tempfile.mkdtemp(prefix="officetool_update_")
    dest = os.path.join(tmp_dir, "OfficeTool-Setup.exe")
    req = urllib.request.Request(download_url, headers={"User-Agent": "OfficeTool-UpdateChecker"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        total = int(resp.headers.get("Content-Length", 0) or 0)
        done = 0
        with open(dest, "wb") as f:
            while True:
                chunk = resp.read(65536)
                if not chunk:
                    break
                f.write(chunk)
                done += len(chunk)
                if on_progress and total:
                    on_progress(done, total)
    subprocess.Popen(
        [dest, "/SILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/DIR=" + engine.DATA_BASE],
        close_fds=True,
    )
    return dest


# Real download progress for the UI — the confirm-then-click install flow
# previously showed only a static "Downloading…" for however long an
# 80MB+ installer took, with no feedback at all, which (combined with a
# separate double-click-confirm timing bug) got reported as "nothing
# happens" after confirming. _PROGRESS is a single shared dict (this app
# only ever runs one update at a time) polled by the frontend via
# /api/apply-update-progress while start_update_async() does the real
# work on a background thread.
_PROGRESS = {"status": "idle", "done": 0, "total": 0, "error": None}


def get_progress():
    return dict(_PROGRESS)


def start_update_async(download_url):
    """Kicks off download_and_launch_installer() on a background thread
    and returns immediately, so the calling HTTP request doesn't block for
    the whole download — the frontend polls get_progress() instead. The
    thread updates _PROGRESS as it goes; the caller (app.py) is
    responsible for exiting the process shortly after status hits
    "launched", same as before."""
    _PROGRESS.update(status="downloading", done=0, total=0, error=None)

    def _run():
        try:
            def _on_progress(done, total):
                _PROGRESS["done"] = done
                _PROGRESS["total"] = total
            download_and_launch_installer(download_url, on_progress=_on_progress)
            _PROGRESS["status"] = "launched"
            # Give the frontend's poll loop a real chance to see
            # status="launched" — and, now, actually READ the toast
            # explaining the install can take a few minutes (see the
            # p.status==='launched' branch in app.py's page script) —
            # before this window disappears. Was 1.0s (same grace the
            # old synchronous version gave the HTTP response); too short
            # to read a toast, not just show it, so the app closing right
            # after felt abrupt right when the user most needed the
            # context that a long silent wait afterward is normal, not a
            # failure.
            time.sleep(3.0)
            os._exit(0)
        except Exception as e:
            _PROGRESS["status"] = "error"
            _PROGRESS["error"] = str(e)

    threading.Thread(target=_run, daemon=True).start()
