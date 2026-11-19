r"""
Auto-update checker — polls GitHub's free public Releases API for a version
newer than the one currently running, and can fetch + launch that release's
installer on the user's say-so. No paid service, no account, no API key:
GitHub's REST API is free and unauthenticated for public repos.

---- One-time setup, once the project is pushed to GitHub ----
1. Set GITHUB_REPO below to "yourusername/Office-Tool".
2. Create your first Release on GitHub, tagged "v1.0.0" (must match VERSION
   at the repo root — see version.py), with BOTH installer_output\
   OfficeTool-Setup.exe AND installer_output\OfficeTool-Payload.zip
   (both built by build.bat) attached as release assets.

---- Shipping every future update ----
1. Bump the VERSION file (e.g. "1.0.1").
2. Rerun build.bat — it rebuilds the .exe, the payload zip, and the
   installer, all with that version baked in.
3. On GitHub: Releases -> Draft a new release, tag "v1.0.1", attach BOTH
   installer_output\OfficeTool-Setup.exe AND
   installer_output\OfficeTool-Payload.zip.
That's it — every copy of the app already running polls this API on its own
(see checkForAppUpdate() in app.py's page script) and will offer the update
automatically next time it's open, no manual download for the user.

---- Why the installer needs the payload zip too (thin-installer split) ----
Per explicit request ("make it that the installer will be very light...
hook all the data from github"), Setup.exe no longer embeds the app's own
files at all — it's just the small Inno Setup wizard shell now, and
downloads OfficeTool-Payload.zip straight from this SAME release's GitHub
assets during install (see installer.iss's own top-of-file comment for the
mechanism). This module's own asset picker below only ever looks for a
"...setup.exe"-named asset, so the payload zip sitting alongside it never
confuses THIS lookup — but forgetting to attach the zip breaks every fresh
install and in-place update alike (Setup.exe would have nothing to fetch),
so the two assets are never optional independently of each other.
"""
import json
import os
import re
import subprocess
import threading
import time
import urllib.error
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
        # download_url now names the PAYLOAD zip, not Setup.exe — per
        # explicit request ("do the update inside the software, no
        # windows installation or nothing... download files with
        # progress bar, also the installation as well"), updating no
        # longer launches any external installer at all (see
        # start_inapp_update_async below); this app downloads+extracts+
        # swaps its own files in place. installer_output\OfficeTool-
        # Setup.exe still ships on every release (a first-time install
        # from the GitHub page still needs it), just never fetched by
        # THIS path anymore.
        asset_url = None
        for a in data.get("assets", []):
            name = (a.get("name") or "").lower()
            if name.endswith("payload.zip"):
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

# Where a downloaded payload zip gets extracted to before it's swapped into
# place — see download_and_stage_payload()/_swap_and_relaunch() below. Lives
# under DATA_BASE (this install's own writable folder), never inside the
# live app folder itself, so a half-finished extract can never leave the
# running app's own files in a half-swapped state.
UPDATE_STAGING_DIR = os.path.join(engine.DATA_BASE, "_update_staging")


def _cached_payload_path(version):
    safe = re.sub(r"[^0-9A-Za-z.]", "", version or "")
    return os.path.join(UPDATE_CACHE_DIR, "OfficeTool-Payload-v{}.zip".format(safe))


def cleanup_update_cache():
    """
    Called once at startup (see app.py, right alongside the other
    best-effort startup housekeeping). If this install is already AT or
    PAST some cached installer's/payload's own version, that cached file
    already did its job — remove it. This is what actually closes the loop
    on "keep it until the system is confirmed updated": the NEXT successful
    launch after an update is exactly the confirmation, and this is where
    that gets noticed. Never raises; routine housekeeping, not a
    load-bearing correctness check.
    """
    try:
        if not os.path.isdir(UPDATE_CACHE_DIR):
            return
        current = _parse_version(APP_VERSION)
        for name in os.listdir(UPDATE_CACHE_DIR):
            m = re.match(r"OfficeTool-(?:Setup|Payload)-v(.+)\.(?:exe|zip)$", name)
            if not m:
                continue
            if _parse_version(m.group(1)) <= current:
                try:
                    os.remove(os.path.join(UPDATE_CACHE_DIR, name))
                except Exception:
                    pass
    except Exception:
        pass
    # Any leftover staging dir from a crash mid-swap is stale the moment
    # this app successfully starts back up — either the swap already
    # finished (staging's copy is now live, nothing to gain by keeping the
    # extra copy around) or it never got that far in the first place, in
    # which case a fresh download+extract next time is simpler and safer
    # than trying to figure out how far a half-swap on disk actually got.
    try:
        if os.path.isdir(UPDATE_STAGING_DIR):
            import shutil
            shutil.rmtree(UPDATE_STAGING_DIR, ignore_errors=True)
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


def _download_resumable(url, dest, on_progress=None):
    """
    HTTP Range-based resumable download — the Python twin of the
    Pascal/PowerShell DownloadFileWithProgress() already built and live-
    tested in installer.iss, reused here for the in-app payload download.
    Same three response cases:

      - 206 Partial Content: server honored our Range request and is
        resuming — the TRUE total comes from the Content-Range response
        header ("bytes start-end/total"), never from this partial
        response's own Content-Length (that only covers the remaining
        bytes, not the whole file).
      - 200 OK: server ignored Range entirely (some CDNs/redirects do)
        and is sending the file from byte 0 — any local partial is stale
        and gets overwritten from scratch.
      - 416 Range Not Satisfiable: our Range start is already at/past the
        server's EOF, i.e. the local file is already complete — nothing
        left to do.

    Per explicit request ("In case of the crash of the software mid
    update or mid install... dont delete the update files which where
    downloaded to temp file and let the installation retry again"): dest
    is never touched on failure, so a retry naturally resumes instead of
    re-downloading from zero.
    """
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    resume_from = os.path.getsize(dest) if os.path.isfile(dest) else 0

    headers = {"User-Agent": "OfficeTool-UpdateChecker"}
    if resume_from > 0:
        headers["Range"] = "bytes={}-".format(resume_from)

    req = urllib.request.Request(url, headers=headers)
    try:
        resp = urllib.request.urlopen(req, timeout=30)
    except urllib.error.HTTPError as e:
        if e.code == 416:
            total = resume_from
            if on_progress:
                on_progress(total, total)
            return dest
        raise

    try:
        status = resp.status
    except AttributeError:
        status = resp.getcode()

    if status == 206:
        content_range = resp.headers.get("Content-Range", "") or ""
        m = re.search(r"/(\d+)\s*$", content_range)
        total = int(m.group(1)) if m else resume_from + int(resp.headers.get("Content-Length", 0) or 0)
        mode, done = "ab", resume_from
    else:
        # Fresh start — either no Range was sent, or the server ignored it.
        total, mode, done = int(resp.headers.get("Content-Length", 0) or 0), "wb", 0

    try:
        with open(dest, mode) as f:
            while True:
                chunk = resp.read(65536)
                if not chunk:
                    break
                f.write(chunk)
                done += len(chunk)
                if on_progress and total:
                    on_progress(done, total)
    finally:
        resp.close()
    return dest


# ---- Delta updates: download only the files that actually changed -------
# Why updates used to be slow: every release ships the WHOLE app as one
# ~125MB OfficeTool-Payload.zip (Python runtime, Playwright's driver,
# PyMuPDF, botocore, numpy, static fonts/images...), and the updater
# downloaded and unpacked all ~3,900 files every time, even though a normal
# release only changes a handful of them. Measured on v1.1.101 -> v1.1.102:
# 3 of 3,898 files differed (OfficeTool.exe, base_library.zip, VERSION),
# 14.6MB of the 125MB zip.
#
# No build or release changes are needed for this: a zip's own central
# directory (the index at the end of the file) already lists every file's
# size and CRC32. GitHub release downloads support HTTP Range requests (the
# resumable full download above already relies on that), so the updater
# reads just that index, compares it with the files already installed,
# fetches only the byte ranges of changed files, and builds the staging
# folder from hard links (or copies) of the unchanged installed files plus
# the freshly downloaded ones. The staging folder ends up identical to a
# full extract, so _swap_and_relaunch() and its rollback are unchanged.
# Anything unexpected (no Range support, a dev run, too much changed, any
# error) falls back to the full download, so this can only make an update
# faster, never break one.

_DELTA_MAX_SHARE = 0.6          # more than this share of the zip changed -> just get the whole zip
_DELTA_MERGE_GAP = 256 * 1024   # join nearby byte ranges into one request
_DELTA_TAIL_BYTES = 4 * 1024 * 1024


class _MissingRange(Exception):
    pass


class _SparseRemoteZip(object):
    """Read-only, seekable file object over the parts of a remote zip that
    have been downloaded so far (the central directory at the end, plus the
    byte ranges of changed files). Reading anywhere else raises
    _MissingRange, which can only mean a bug in the range maths; the caller
    then falls back to the full download."""

    def __init__(self, total):
        self.total = total
        self.chunks = []  # [(start, bytes)]
        self.pos = 0

    def add(self, start, data):
        self.chunks.append((start, data))

    def seekable(self):
        return True

    def readable(self):
        return True

    def tell(self):
        return self.pos

    def seek(self, offset, whence=0):
        if whence == 0:
            self.pos = offset
        elif whence == 1:
            self.pos += offset
        else:
            self.pos = self.total + offset
        return self.pos

    def read(self, n=-1):
        if n is None or n < 0:
            n = self.total - self.pos
        n = max(0, min(n, self.total - self.pos))
        out = []
        while n > 0:
            for start, data in self.chunks:
                if start <= self.pos < start + len(data):
                    piece = data[self.pos - start:self.pos - start + n]
                    out.append(piece)
                    self.pos += len(piece)
                    n -= len(piece)
                    break
            else:
                raise _MissingRange("byte {} of the update was not downloaded".format(self.pos))
        return b"".join(out)

    def close(self):
        pass


def _http_range(url, start, end, timeout=60):
    """Bytes start..end (inclusive) of url. Raises unless the server really
    answered with that range (206), so a server that ignores Range can never
    be mistaken for the requested bytes."""
    req = urllib.request.Request(url, headers={
        "User-Agent": "OfficeTool-UpdateChecker",
        "Range": "bytes={}-{}".format(start, end),
    })
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        status = getattr(resp, "status", None) or resp.getcode()
        if status != 206:
            raise RuntimeError("server ignored the Range request (HTTP {})".format(status))
        content_range = resp.headers.get("Content-Range", "") or ""
        data = resp.read()
    if len(data) != end - start + 1:
        raise RuntimeError("short range read ({} of {} bytes)".format(len(data), end - start + 1))
    return data, content_range


def _remote_size(url):
    _, content_range = _http_range(url, 0, 0)
    m = re.search(r"/(\d+)\s*$", content_range)
    if not m:
        raise RuntimeError("no Content-Range total from server")
    return int(m.group(1))


def _local_file_matches(path, size, crc):
    try:
        if os.path.getsize(path) != size:
            return False
        import zlib
        value = 0
        with open(path, "rb") as f:
            while True:
                block = f.read(1024 * 1024)
                if not block:
                    break
                value = zlib.crc32(block, value)
        return (value & 0xFFFFFFFF) == crc
    except OSError:
        return False


def _link_or_copy(src, dst):
    """Hard link when possible (instant, no extra disk space; safe because
    the swap only ever renames/replaces whole files, never writes into one),
    a normal copy otherwise."""
    import shutil
    try:
        os.link(src, dst)
    except Exception:
        shutil.copy2(src, dst)


def _try_delta_stage(download_url, install_dir, staging_dir, on_progress=None):
    """Builds staging_dir from the installed files plus only the changed
    files of the remote payload zip. Returns True when staging_dir is ready,
    or False when a delta isn't worth it / isn't possible here (the caller
    then does the normal full download). May also raise; the caller treats
    that as False too."""
    import sys
    import shutil
    import zipfile

    if not getattr(sys, "frozen", False):
        return False
    if not os.path.isfile(os.path.join(install_dir, "OfficeTool.exe")):
        return False

    if on_progress:
        on_progress("checking", 0, 0)

    total = _remote_size(download_url)
    src = _SparseRemoteZip(total)
    tail_start = max(0, total - _DELTA_TAIL_BYTES)
    tail, _ = _http_range(download_url, tail_start, total - 1)
    src.add(tail_start, tail)
    zf = zipfile.ZipFile(src)  # reads only the central directory, all inside `tail`
    members = zf.infolist()
    # Compress-Archive (build.bat) writes backslash paths; Python on Windows
    # already turns them into "/", this just makes sure of it.
    for m in members:
        m.filename = m.filename.replace("\\", "/")
    files = [m for m in members if not m.is_dir()]

    # Compare every file in the new release with what's installed now.
    changed = []
    for i, m in enumerate(files, 1):
        name = m.filename
        local = os.path.join(install_dir, *name.split("/"))
        if not _local_file_matches(local, m.file_size, m.CRC):
            changed.append(m)
        if on_progress and i % 50 == 0:
            on_progress("checking", i, len(files))

    # Each file's bytes in the zip run from its local header up to the next
    # file's header (or the central directory).
    starts = sorted(set([m.header_offset for m in members] + [zf.start_dir]))
    next_start = {a: b for a, b in zip(starts, starts[1:])}
    spans = sorted((m.header_offset, next_start.get(m.header_offset, zf.start_dir) - 1)
                   for m in changed)
    merged = []
    for a, b in spans:
        if merged and a <= merged[-1][1] + 1 + _DELTA_MERGE_GAP:
            merged[-1][1] = max(merged[-1][1], b)
        else:
            merged.append([a, b])
    need = [(a, b) for a, b in merged if not (a >= tail_start)]
    to_get = sum(b - a + 1 for a, b in need)
    if to_get > total * _DELTA_MAX_SHARE:
        return False

    done = 0
    if on_progress:
        on_progress("downloading", 0, to_get)
    for a, b in need:
        # Pieces of at most 8MB so the progress bar moves and a dropped
        # connection only costs one piece.
        pos = a
        while pos <= b:
            end = min(b, pos + 8 * 1024 * 1024 - 1)
            data, _ = _http_range(download_url, pos, end)
            src.add(pos, data)
            done += len(data)
            pos = end + 1
            if on_progress:
                on_progress("downloading", done, to_get)

    if os.path.isdir(staging_dir):
        shutil.rmtree(staging_dir, ignore_errors=True)
    os.makedirs(staging_dir, exist_ok=True)

    changed_names = set(m.filename for m in changed)
    for i, m in enumerate(members, 1):
        name = m.filename
        parts = [p for p in name.split("/") if p not in ("", ".", "..")]
        if not parts:
            continue
        dest = os.path.join(staging_dir, *parts)
        if m.is_dir():
            os.makedirs(dest, exist_ok=True)
        elif m.filename in changed_names:
            zf.extract(m, staging_dir)  # CRC is checked while reading
        else:
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            _link_or_copy(os.path.join(install_dir, *parts), dest)
        if on_progress and (i % 50 == 0 or i == len(members)):
            on_progress("extracting", i, len(members))

    if not os.path.isfile(os.path.join(staging_dir, "OfficeTool.exe")):
        raise RuntimeError("delta staging is missing OfficeTool.exe")
    return True


def _log_update(msg):
    try:
        with open(os.path.join(engine.DATA_BASE, "last_inapp_update.log"), "a", encoding="utf-8") as f:
            f.write("{}  {}\n".format(time.strftime("%Y-%m-%dT%H:%M:%S"), msg))
    except Exception:
        pass


_MIN_VALID_PAYLOAD_BYTES = 10 * 1024 * 1024  # a real payload zip is tens of MB+; anything smaller is a bad/partial download


def download_and_stage_payload(download_url, target_version=None, on_progress=None):
    """
    Downloads the release's OfficeTool-Payload.zip (resumable, see
    _download_resumable above) into UPDATE_CACHE_DIR, then extracts it
    into a clean UPDATE_STAGING_DIR, verifying OfficeTool.exe actually
    landed before calling it good. Replaces the old download-then-launch-
    Setup.exe flow entirely — per explicit request ("do the update inside
    the software no windows installation or nothing... The Update files
    download with progress bar, also the installation as well"), nothing
    here ever shells out to an external installer.

    on_progress(phase, done, total) is called for phase in
    ("downloading", "extracting") so the frontend can show a real bar for
    both, not just the download.

    Returns UPDATE_STAGING_DIR on success (ready for _swap_and_relaunch).
    Raises on failure — the caller (start_inapp_update_async) is
    responsible for surfacing that as _PROGRESS["status"]="error". The
    cached zip is deliberately NOT deleted until extraction is verified
    good, so a crash/failure mid-extract still leaves the download itself
    reusable on retry instead of costing the user another 80MB+ fetch.
    """
    os.makedirs(UPDATE_CACHE_DIR, exist_ok=True)

    # Fast path: only the files that changed (see _try_delta_stage above).
    try:
        if _try_delta_stage(download_url, engine.DATA_BASE, UPDATE_STAGING_DIR, on_progress=on_progress):
            _log_update("delta update staged (only changed files downloaded)")
            return UPDATE_STAGING_DIR
        _log_update("delta update not used, downloading the full payload")
    except Exception as e:
        _log_update("delta update failed ({}), downloading the full payload".format(e))

    zip_path = _cached_payload_path(target_version) if target_version else \
        os.path.join(UPDATE_CACHE_DIR, "OfficeTool-Payload-latest.zip")

    def _dl_progress(done, total):
        if on_progress:
            on_progress("downloading", done, total)

    _download_resumable(download_url, zip_path, on_progress=_dl_progress)

    if os.path.getsize(zip_path) < _MIN_VALID_PAYLOAD_BYTES:
        try:
            os.remove(zip_path)
        except Exception:
            pass
        raise RuntimeError("Downloaded update file is too small — likely a corrupt/interrupted download. Removed so a retry starts clean.")

    import shutil
    import zipfile
    if os.path.isdir(UPDATE_STAGING_DIR):
        shutil.rmtree(UPDATE_STAGING_DIR, ignore_errors=True)
    os.makedirs(UPDATE_STAGING_DIR, exist_ok=True)

    try:
        with zipfile.ZipFile(zip_path) as zf:
            members = zf.infolist()
            total = len(members)
            for i, member in enumerate(members, 1):
                zf.extract(member, UPDATE_STAGING_DIR)
                if on_progress:
                    on_progress("extracting", i, total)
    except zipfile.BadZipFile:
        # A corrupt zip is never salvageable by resuming (unlike a
        # truncated download, this is a content error) — remove it so a
        # retry re-downloads instead of re-extracting the same bad file.
        try:
            os.remove(zip_path)
        except Exception:
            pass
        raise

    if not os.path.isfile(os.path.join(UPDATE_STAGING_DIR, "OfficeTool.exe")):
        raise RuntimeError("Extracted update is missing OfficeTool.exe — refusing to install a broken payload.")

    try:
        os.remove(zip_path)
    except Exception:
        pass

    return UPDATE_STAGING_DIR


def _ps_escape(s):
    """Escapes a value for safe interpolation inside a PowerShell
    double-quoted string literal (backtick is PS's escape char; $ would
    otherwise trigger variable expansion)."""
    return s.replace("`", "``").replace('"', '`"').replace("$", "`$")


# The actual file-swap + relaunch happens in a DETACHED PowerShell script,
# not in this Python process — because this process's own OfficeTool.exe
# and _internal\*.dll/.pyd files are locked open for as long as it's
# running, exactly the same constraint _launch_installer() above worked
# around for the old Setup.exe flow. Written to a real .ps1 file (not
# passed inline via -Command) so the multi-step swap+rollback logic reads
# and debugs like a normal script instead of one giant escaped one-liner —
# installer.iss's own Pascal Format()/'[' parsing traps (see that file's
# comments) are exactly the class of bug this sidesteps.
_SWAP_PS_TEMPLATE = r'''$ErrorActionPreference = "Stop"
$log = "__LOG__"
$targetPid = __PID__
$dst = "__DST__"
$src = "__SRC__"
$exe = Join-Path $dst "OfficeTool.exe"

function Log($msg) {
    "$(Get-Date -Format o)  $msg" | Out-File -FilePath $log -Append -Encoding utf8
}

Log "in-app update swap starting, waiting for pid $targetPid to exit"
$n = 0
while ((Get-Process -Id $targetPid -ErrorAction SilentlyContinue) -and ($n -lt 40)) {
    Start-Sleep -Milliseconds 250
    $n++
}

$internalDst = Join-Path $dst "_internal"
$internalBak = Join-Path $dst "_internal.bak"
$exeBak = Join-Path $dst "OfficeTool.exe.bak"

try {
    if (Test-Path $internalBak) { Remove-Item -Recurse -Force $internalBak -ErrorAction SilentlyContinue }
    if (Test-Path $exeBak) { Remove-Item -Force $exeBak -ErrorAction SilentlyContinue }

    # Back up the CURRENT (working) files before touching anything, so a
    # failure partway through this block can restore a known-good app
    # instead of leaving a half-swapped, unlaunchable one.
    if (Test-Path $internalDst) { Rename-Item -Path $internalDst -NewName "_internal.bak" }
    if (Test-Path $exe) { Copy-Item -Path $exe -Destination $exeBak -Force }

    Move-Item -Path (Join-Path $src "_internal") -Destination $internalDst -Force
    Copy-Item -Path (Join-Path $src "OfficeTool.exe") -Destination $exe -Force
    Get-ChildItem -Path $src -File | Where-Object { $_.Name -ne "OfficeTool.exe" } | ForEach-Object {
        Copy-Item -Force $_.FullName (Join-Path $dst $_.Name)
    }
    Get-ChildItem -Path $src -Directory | Where-Object { $_.Name -ne "_internal" } | ForEach-Object {
        Copy-Item -Recurse -Force $_.FullName (Join-Path $dst $_.Name)
    }

    Log "files swapped, launching new exe"
    Start-Process -FilePath $exe
    Start-Sleep -Seconds 3

    $running = Get-Process -Name "OfficeTool" -ErrorAction SilentlyContinue | Where-Object {
        $_.Path -and ($_.Path -eq $exe)
    }
    if ($running) {
        Log "new version confirmed running, cleaning up backups + staging"
        Remove-Item -Recurse -Force $internalBak -ErrorAction SilentlyContinue
        Remove-Item -Force $exeBak -ErrorAction SilentlyContinue
        Remove-Item -Recurse -Force $src -ErrorAction SilentlyContinue
    } else {
        throw "new OfficeTool.exe did not start"
    }
} catch {
    Log ("swap FAILED: " + $_.Exception.Message + " -- rolling back to the previous version")
    try {
        if (Test-Path $internalDst) { Remove-Item -Recurse -Force $internalDst -ErrorAction SilentlyContinue }
        if (Test-Path $internalBak) { Rename-Item -Path $internalBak -NewName "_internal" }
        if (Test-Path $exeBak) { Copy-Item -Path $exeBak -Destination $exe -Force }
        Start-Process -FilePath $exe
        Log "rollback complete, previous version relaunched"
    } catch {
        Log ("ROLLBACK ALSO FAILED: " + $_.Exception.Message)
    }
}

# This helper script's own job is done either way (swapped, or rolled
# back) — per explicit request ("after the update was downloaded and
# installed successfully it should delete all the temp downloaded
# files... i dont want to have garbage inside the pc"), it deletes
# itself as the very last action instead of sitting in update_cache\
# forever. Safe while still running: PowerShell has already read the
# whole script into memory by this point, and NTFS allows deleting an
# open file (the actual unlink just waits for the last handle to close).
# Best-effort, own try/catch — a failure here is cosmetic clutter, never
# worth surfacing as a swap error this late.
try {
    Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue
} catch {}
'''


def _swap_and_relaunch(staging_dir):
    """
    Kills any lock-holding bundled_browser processes (same as the old
    installer path — see _kill_bundled_browser_processes()'s own
    docstring for the real RestartManager bug this guards against), then
    hands off to the detached PowerShell script above: wait for THIS
    process's own PID to actually exit, back up the current _internal\\ +
    OfficeTool.exe, swap in the staged ones, relaunch, and only clean up
    the backups once the new process is confirmed alive — rolling back
    to the previous working version otherwise. Fire-and-forget: by the
    time this script's real work starts, the Python process that launched
    it is already gone.
    """
    _kill_bundled_browser_processes()
    log_path = os.path.join(engine.DATA_BASE, "last_inapp_update.log")
    ps1_path = os.path.join(UPDATE_CACHE_DIR, "_swap_helper.ps1")
    os.makedirs(UPDATE_CACHE_DIR, exist_ok=True)

    script = (
        _SWAP_PS_TEMPLATE
        .replace("__LOG__", _ps_escape(log_path))
        .replace("__PID__", str(os.getpid()))
        .replace("__DST__", _ps_escape(engine.DATA_BASE))
        .replace("__SRC__", _ps_escape(staging_dir))
    )
    with open(ps1_path, "w", encoding="utf-8") as f:
        f.write(script)

    subprocess.Popen(
        ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-ExecutionPolicy", "Bypass", "-File", ps1_path],
        creationflags=subprocess.CREATE_NO_WINDOW,
        close_fds=True,
    )


# Real download/extract progress for the UI — see the dimmed in-app
# overlay in app.py's page script. _PROGRESS is a single shared dict
# (this app only ever runs one update at a time) polled via
# /api/apply-update-progress while start_inapp_update_async() does the
# real work on a background thread. status cycles through:
#   ["checking" ->] "downloading" -> "extracting" -> "installing" -> (process exits)
# ("checking" only for a delta update, see _try_delta_stage.)
# "installing" is the signal to the frontend that the window is about to
# vanish and reappear on its own — see actuallyInstallUpdate()'s poll
# loop for how it turns that into "hold on, restarting..." instead of
# just going blank.
_PROGRESS = {"status": "idle", "done": 0, "total": 0, "error": None}


def get_progress():
    return dict(_PROGRESS)


def start_inapp_update_async(download_url, target_version=None, on_before_exit=None):
    """
    Kicks off download_and_stage_payload() + _swap_and_relaunch() on a
    background thread and returns immediately, so the calling HTTP
    request doesn't block for the whole download — the frontend polls
    get_progress() instead. Per explicit final request ("you know if its
    possilbe you can do the update inside the software no windows
    installation or nothing, it will all be on the software only. The
    Update files download with progress bar, also the installation as
    well") this never launches Inno Setup at all — see
    download_and_stage_payload()/_swap_and_relaunch() above for the
    actual mechanism.

    on_before_exit: optional zero-arg callback run right before
    os._exit(0) — per explicit report ("sometime there are 2 softwares in
    this trail, some time even more"): this module has no reference to
    app.py's own tray icon object (a different module entirely), so
    os._exit(0) here used to skip pystray's own icon.stop() call
    entirely, meaning Windows never got the Shell_NotifyIcon(NIM_DELETE)
    that actually removes the icon from the notification area — it just
    sits there orphaned until something makes Explorer notice the owning
    process is gone. app.py passes its own _stop_tray_icon() here so this
    module can trigger that teardown without needing to import app.py
    itself (which would be circular — app.py is what imports THIS
    module). Best-effort: any failure here must never block the actual
    exit, same reasoning as _stop_tray_icon()'s own try/except.
    """
    _PROGRESS.update(status="downloading", done=0, total=0, error=None)

    def _run():
        try:
            def _on_progress(phase, done, total):
                _PROGRESS["status"] = phase
                _PROGRESS["done"] = done
                _PROGRESS["total"] = total
            staging_dir = download_and_stage_payload(download_url, target_version=target_version, on_progress=_on_progress)
            _PROGRESS.update(status="installing", done=0, total=0)
            _swap_and_relaunch(staging_dir)
            # Give the frontend's poll loop a real chance to see
            # status="installing" (and switch to the "restarting now"
            # message) before this process actually disappears — same
            # grace-period rationale as the old start_update_async().
            time.sleep(1.5)
            if on_before_exit:
                try:
                    on_before_exit()
                except Exception:
                    pass
            os._exit(0)
        except Exception as e:
            _PROGRESS["status"] = "error"
            _PROGRESS["error"] = str(e)

    threading.Thread(target=_run, daemon=True).start()
