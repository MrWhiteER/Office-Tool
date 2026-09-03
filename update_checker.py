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
import time
import urllib.request

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
    Downloads the release's Setup.exe to a fresh temp folder and launches
    it, non-silently — the same install wizard the user already ran once,
    so "Next, Next, Install" upgrades in place (same AppId as the current
    install — see installer.iss) without wiping config.json/drafts/etc.
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
    subprocess.Popen([dest], close_fds=True)
    return dest
