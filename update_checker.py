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
import runtime_manager
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

    Deliberately NOT called on a short interval (e.g. every few seconds)
    while the app is running — confirmed directly against the real API
    (a live rate_limit check before/after several calls) that even a
    conditional If-None-Match request that gets back a genuine 304 Not
    Modified STILL costs 1 unit of the unauthenticated 60/hour budget
    here; an ETag-based "free polling" cache was tried and measured, not
    just assumed, and removed once the numbers showed it doesn't help.
    That 60/hour is also PER IP, shared across every Office Tool install
    on the same office network, not per-install — so anything faster
    than roughly a couple of minutes risks the whole office collectively
    exhausting it and update checks silently failing for everyone. See
    initUpdateChecking() in the page script for the interval actually
    used while running, and its own comment for the real numbers this
    was based on.
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


# ---- Downloaded-installer cache -----------------------------------------
# Lives under DATA_BASE (this install's own persistent, writable folder —
# see engine.py's BASE/DATA_BASE split) so a failed install attempt never
# costs the user another 80MB+ download to try again. Per explicit
# request: "keep the installation file until the system was not confirmed
# to be updated to the latest version... he will not be required to
# download the update again". Named with the TARGET version so a leftover
# file from some earlier, already-superseded update attempt is never
# mistaken for the one currently being offered.
UPDATE_CACHE_DIR = os.path.join(engine.DATA_BASE, "update_cache")
_MIN_VALID_INSTALLER_BYTES = 10 * 1024 * 1024  # guards against a half-written file from a previous crash/interrupt


def _cached_installer_path(version):
    safe = re.sub(r"[^0-9A-Za-z.]", "", version or "")
    return os.path.join(UPDATE_CACHE_DIR, "OfficeTool-Setup-v{}.exe".format(safe))


def cleanup_update_cache():
    """
    Called once at startup (see app.py, right alongside the other
    best-effort startup housekeeping). If this install is already AT or
    PAST some cached installer's own version, that cached file already
    did its job — remove it. This is what actually closes the loop on
    "keep it until the system is confirmed updated": the NEXT successful
    launch after an update is exactly the confirmation, and this is where
    that gets noticed. Never raises; routine housekeeping, not a
    load-bearing correctness check.
    """
    try:
        if not os.path.isdir(UPDATE_CACHE_DIR):
            return
        current = _parse_version(APP_VERSION)
        for name in os.listdir(UPDATE_CACHE_DIR):
            m = re.match(r"OfficeTool-Setup-v(.+)\.exe$", name)
            if not m:
                continue
            if _parse_version(m.group(1)) <= current:
                try:
                    os.remove(os.path.join(UPDATE_CACHE_DIR, name))
                except Exception:
                    pass
    except Exception:
        pass


def _kill_bundled_browser_processes():
    """
    Terminates every leftover process this install's own Playwright
    integration can leave running — chrome.exe/chrome_proxy.exe (and its
    GPU/renderer/utility helper processes, which all share that same exe
    path) under bundled_browser\\, AND playwright\\driver\\node.exe, the
    SEPARATE Node.js process Playwright's own driver runs as (it talks to
    the browser over its own protocol; it is not part of Chromium at
    all) — right before the installer gets a chance to try overwriting
    any of those files.

    This is the real fix for a confirmed, reproduced bug: a real update
    attempt (v1.1.17 -> v1.1.20) downloaded fine but the installer
    "quit instantly" with zero explanation. installer.iss's own
    CloseApplications=yes should handle this on its own via Windows'
    Restart Manager, but evidently doesn't reliably close every one of
    these under /SILENT /SUPPRESSMSGBOXES — and SUPPRESSMSGBOXES hides
    whatever error/retry UI Inno Setup would otherwise show for a locked
    file, so a failure here is completely silent to the user, exactly
    matching what was reported. Directly confirmed the underlying cause
    TWICE, live, on the user's own real PC: first pass found six
    chrome.exe processes holding handles under bundled_browser\\... —
    fixed, shipped, then reproduced AGAIN live (v1.1.22 -> v1.1.25):
    Inno Setup's own install log named the exact remaining culprit —
    "RestartManager found an application using one of our files:
    Node.js JavaScript Runtime" — "Some applications could not be shut
    down" -> defaults to Abort under SUPPRESSMSGBOXES -> "User canceled
    the installation process. Rolling back changes." — the Playwright
    Node driver process was never being killed at all, only its Chromium
    child was. Confirmed the actual bundled path directly
    (_internal\\playwright\\driver\\node.exe) before writing this fix.

    Uses WMI (via the pywin32 dependency already bundled) to filter by
    ExecutablePath specifically, not just process NAME — "node.exe"/
    "chrome.exe" alone would also match some unrelated real Node.js app
    or the user's own real Google Chrome browser, neither of which must
    ever be touched. Best-effort: any failure here (WMI unavailable,
    permission issue, etc.) is swallowed — worst case, this just falls
    back to relying on installer.iss's own CloseApplications=yes exactly
    as before this fix existed.
    """
    try:
        import win32com.client
        # As of v1.1.27 (see runtime_manager.py), the real Chromium moved
        # OUT of engine.BASE\bundled_browser (inside _internal\, which
        # this installer replaces) and into
        # runtime_manager.BUNDLED_BROWSER_DIR (engine.DATA_BASE\runtime\
        # bundled_browser\, outside _internal\ entirely, never touched by
        # a normal app update at all). Checking BOTH here: the new
        # location for any install actually running the split runtime,
        # the old one as a harmless no-op fallback for anyone somehow
        # still on the pre-split layout mid-upgrade.
        bundled_browser_dirs = tuple(os.path.normcase(d) for d in (
            runtime_manager.BUNDLED_BROWSER_DIR,
            os.path.join(engine.BASE, "bundled_browser"),
        ))
        playwright_driver_dir = os.path.normcase(os.path.join(engine.BASE, "playwright", "driver"))
        wmi = win32com.client.GetObject("winmgmts:")
        procs = wmi.ExecQuery(
            "SELECT ProcessId, ExecutablePath FROM Win32_Process "
            "WHERE Name='chrome.exe' OR Name='chrome_proxy.exe' OR Name='node.exe'"
        )
        killed = 0
        for p in procs:
            exe_path = os.path.normcase(p.ExecutablePath or "")
            if exe_path.startswith(bundled_browser_dirs) or exe_path.startswith(playwright_driver_dir):
                try:
                    # p.Terminate() (WMI's own Win32_Process method) — NOT
                    # os.system("taskkill ...") like this used to be. On a
                    # --windowed PyInstaller build (no console of its own
                    # at all), os.system() spawns a real cmd.exe /c child
                    # for every single call, and each one briefly flashes
                    # its own console window — exactly what was reported
                    # live: "multiple times the console opens and closes,
                    # and there is no text there" (empty because the old
                    # command redirected its own output to nul, but the
                    # window itself still flashed). Calling Terminate()
                    # straight through the existing WMI COM object kills
                    # the process with zero new processes spawned at all,
                    # so there's nothing left to flash a window.
                    p.Terminate()
                    killed += 1
                except Exception:
                    pass
        if killed:
            # A moment for Windows to actually release the file handles —
            # TerminateProcess (what taskkill /F uses under the hood)
            # returns as soon as the kill is requested, not necessarily
            # once every mapped file handle involved is fully torn down.
            time.sleep(0.5)
    except Exception:
        pass


def _launch_installer(dest):
    """Kills any lock-holding bundled_browser processes, then schedules the
    installer to run — NOT immediately. Real /LOG= from a fresh reproduced
    failure (v1.1.28, "again the same issue... Rolling Back Changes"):

        RestartManager found an application using one of our files: OfficeTool.exe
        Some applications could not be shut down.
        User canceled the installation process. Rolling back changes.

    ...timestamped under half a second after Setup.exe started. The old
    code here launched Setup.exe via Popen FIRST, and only THEN had the
    caller (start_update_async) sleep 3s before os._exit(0) — so Setup's
    RestartManager check, which fires almost instantly, was GUARANTEED to
    still find this app's own process alive every single time, no matter
    how reliably chrome.exe/node.exe get killed above. RestartManager's
    own "ask it to close" mechanism (CloseApplications=yes in
    installer.iss) evidently isn't fast/reliable enough either — it's
    what actually reported "some applications could not be shut down."

    Fix: launch Setup.exe from a detached helper that waits for THIS
    process's own PID to actually disappear (capped at 10s as a safety
    net, in case this process somehow never exits on its own — better to
    install late than hang forever) before starting it, so Setup's very
    first RestartManager check always runs after this process has
    already fully exited — never racing it. PowerShell's Get-Process,
    not a hand-rolled tasklist/findstr/goto batch loop: far less fragile
    to get right as a single-line invocation, and Get-Process -Id
    <pid> -ErrorAction SilentlyContinue is a clean, direct "does this PID
    still exist" check rather than parsing tasklist's text table.
    -WindowStyle Hidden + CREATE_NO_WINDOW below: this helper (and
    whatever it briefly spawns) must never itself flash a visible window
    — see _kill_bundled_browser_processes()'s own comment on the exact
    live symptom that pattern caused elsewhere in this same fix.
    /LOG= writes a real install log even under fully silent mode —
    previously a failure here was invisible even to us; this gives a
    concrete file to check next time something goes wrong, without
    changing the silent user experience at all."""
    _kill_bundled_browser_processes()
    log_path = os.path.join(engine.DATA_BASE, "last_update_install.log")
    pid = os.getpid()

    def _pq(s):
        # Single-quoted PowerShell string literal — doubling any literal
        # single quote is PowerShell's own escape for that, not related
        # to cmd.exe's separate (and much hairier) quoting rules.
        return "'" + s.replace("'", "''") + "'"

    # Deliberately NO -WindowStyle Hidden on THIS Start-Process call — that
    # would hide Setup.exe's own small native progress window, undoing the
    # whole point of /SILENT (not /VERYSILENT) established back in v1.1.7:
    # a completely silent multi-minute install with zero visible feedback
    # reads exactly like a hang/failure. Hidden is only for the PowerShell
    # host wrapping this wait — see the Popen call below.
    ps_script = (
        "$n=0; while ((Get-Process -Id {pid} -ErrorAction SilentlyContinue) -and ($n -lt 40)) "
        "{{ Start-Sleep -Milliseconds 250; $n++ }}; "
        "Start-Process -FilePath {dest} -ArgumentList "
        "'/SILENT','/SUPPRESSMSGBOXES','/NORESTART',('/LOG=' + {log}),('/DIR=' + {dir})"
    ).format(pid=pid, dest=_pq(dest), log=_pq(log_path), dir=_pq(engine.DATA_BASE))
    subprocess.Popen(
        ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps_script],
        creationflags=subprocess.CREATE_NO_WINDOW,
        close_fds=True,
    )


def download_and_launch_installer(download_url, target_version=None, on_progress=None):
    """
    Downloads the release's Setup.exe and runs it with /SILENT
    /SUPPRESSMSGBOXES /NORESTART — no wizard, no "Welcome to Setup"
    screen or license/directory pages, just Inno Setup's own small
    native progress bar while it works (see installer.iss's [Run] —
    skipifsilent was removed specifically so a silent run still reopens
    the app afterward on its own). Same AppId as the current install
    (installer.iss) means this is always an in-place upgrade, never a
    fresh install, so config.json/drafts/etc. survive exactly as before.

    target_version, if given, names the download against UPDATE_CACHE_DIR
    (see above) instead of a throwaway temp folder — and if a complete,
    valid-looking copy from an earlier attempt is already sitting there,
    this skips the download entirely and launches straight away. This is
    what actually saves the user a repeat download after a failed
    install, per explicit request.

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
    if target_version:
        os.makedirs(UPDATE_CACHE_DIR, exist_ok=True)
        dest = _cached_installer_path(target_version)
        if os.path.isfile(dest) and os.path.getsize(dest) >= _MIN_VALID_INSTALLER_BYTES:
            size = os.path.getsize(dest)
            if on_progress:
                on_progress(size, size)
            _launch_installer(dest)
            return dest
    else:
        os.makedirs(tempfile.gettempdir(), exist_ok=True)
        dest = os.path.join(tempfile.mkdtemp(prefix="officetool_update_"), "OfficeTool-Setup.exe")

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
    _launch_installer(dest)
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


def start_update_async(download_url, target_version=None):
    """Kicks off download_and_launch_installer() on a background thread
    and returns immediately, so the calling HTTP request doesn't block for
    the whole download — the frontend polls get_progress() instead. The
    thread updates _PROGRESS as it goes; the caller (app.py) is
    responsible for exiting the process shortly after status hits
    "launched", same as before. target_version is threaded straight
    through to download_and_launch_installer() for the installer cache
    (see its own docstring)."""
    _PROGRESS.update(status="downloading", done=0, total=0, error=None)

    def _run():
        try:
            def _on_progress(done, total):
                _PROGRESS["done"] = done
                _PROGRESS["total"] = total
            download_and_launch_installer(download_url, target_version=target_version, on_progress=_on_progress)
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
