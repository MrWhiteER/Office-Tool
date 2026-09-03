r"""
Drives a physically-connected scanner via Windows Image Acquisition (WIA) —
built into Windows, works with virtually any scanner that has a driver
(which is effectively all of them), no paid API/service and no extra
install beyond pywin32 (already in requirements.txt). Replaces the old
"scan externally with the scanner's own software, then browse to find the
file" two-step for linking a signed Delivery Order to a Submission (see
app.py's /api/scanner-* routes and linkScannedDo()/openScanNow() in the
page script) with one in-app "Scan Now" action.

---- Flow ----
1. /api/scanner-list — enumerate connected scanners (almost always exactly
   one; the UI only shows a picker if there's more than one).
2. /api/scanner-scan-page — scans ONE page via WIA, saves it to a temp PNG,
   returns a session_id (first call) + a preview thumbnail. Call again
   (same session_id) to add more pages — real Delivery Orders are
   sometimes 2 pages (a signature page after the main one).
3. /api/scanner-finalize — combines every captured page (in order) into
   ONE PDF via Pillow, named after the submission it's for (see
   _build_scan_filename()), saved into the Scanned Delivery Orders folder
   (brand_settings()['scanned_do_folder']) — same folder/shape
   /api/browse-scanned-do already produces, so submissions-link-scanned-do
   needs zero changes.
4. /api/scanner-cancel — discards an in-progress session's captured pages.

Session state (SCAN_SESSIONS) is a plain in-memory dict, not a database —
this is a single-user local desktop app, one scan happens at a time, and
nothing here needs to survive a restart.
"""
import base64
import datetime
import io
import os
import tempfile
import uuid

from PIL import Image

WIA_DEVICE_TYPE_SCANNER = 1
WIA_FORMAT_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
# Well-known WIA scan-item property IDs — set defensively (try/except) since
# support varies by driver; scanning still works with driver defaults if
# these are rejected. 6146=current intent (1 color/2 gray/4 B&W text),
# 6147/6148=horizontal/vertical DPI.
_WIA_PROP_INTENT = 6146
_WIA_PROP_DPI_X = 6147
_WIA_PROP_DPI_Y = 6148

SCAN_SESSIONS = {}  # session_id -> {"pages": [tmp_png_path, ...], "dir": tmp_dir}


def _com():
    """WIA is COM — each thread that touches it needs its own
    CoInitialize() (Flask's dev server hands each request a fresh OS
    thread). Returns the pywin32 modules so callers don't need their own
    top-level `import win32com.client` (keeps the import — and the clear
    "pywin32 not installed" error — inside functions that actually need it)."""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    return pythoncom, win32com.client


def _list_scanner_infos(win32com_client):
    """Isolated in its own function so every COM reference it creates
    (manager, each info, the loop variable) goes out of scope and gets
    released the moment it returns — BEFORE the caller's CoUninitialize()
    runs. Doing that release after CoUninitialize is harmless but makes
    pywin32 log a noisy "Win32 exception releasing IUnknown" to stderr."""
    manager = win32com_client.Dispatch("WIA.DeviceManager")
    out = []
    for info in manager.DeviceInfos:
        if info.Type == WIA_DEVICE_TYPE_SCANNER:
            try:
                name = info.Properties("Name").Value
            except Exception:
                name = "Scanner"
            out.append({"id": info.DeviceID, "name": name})
    return out


def list_scanners():
    """Returns [{"id", "name"}, ...] — empty list (not an error) if none
    connected, so the UI can show a friendly "no scanner found" state."""
    try:
        pythoncom, win32com_client = _com()
        try:
            return _list_scanner_infos(win32com_client)
        finally:
            pythoncom.CoUninitialize()
    except Exception:
        return []


def _connect_device(win32com_client, device_id):
    """Same isolation reasoning as _list_scanner_infos() — everything
    except the final Connect()'d device itself is released before this
    returns."""
    manager = win32com_client.Dispatch("WIA.DeviceManager")
    for info in manager.DeviceInfos:
        if info.Type == WIA_DEVICE_TYPE_SCANNER and (device_id is None or info.DeviceID == device_id):
            return info.Connect()
    raise RuntimeError("No scanner found — check it's connected, turned on, and its driver is installed.")


def _capture_page_to_file(win32com_client, device_id, dest_path):
    """Isolated for the same reason as _list_scanner_infos() — connects,
    scans, AND saves to disk all in here, so every COM reference
    (device/item/image) is released before this returns and the caller
    can CoUninitialize() cleanly with nothing left dangling.

    Requesting WIA_FORMAT_PNG is a polite ask, not a guarantee — plenty of
    real scanner drivers ignore it and hand back their own native format
    regardless (confirmed live: this machine's own scanner returns BMP
    bytes even though PNG was requested). Pillow re-saves whatever comes
    back as a real PNG at dest_path, so every downstream consumer (the
    preview data: URI, the PDF assembly in finalize_session()) can rely on
    dest_path always actually being a PNG, never guessing from the
    driver's mood."""
    device = _connect_device(win32com_client, device_id)
    item = device.Items[1]
    for prop_id, value in ((_WIA_PROP_INTENT, 1), (_WIA_PROP_DPI_X, 200), (_WIA_PROP_DPI_Y, 200)):
        try:
            item.Properties(prop_id).Value = value
        except Exception:
            pass  # driver doesn't support this property — scan with its own default instead
    image = item.Transfer(WIA_FORMAT_PNG)
    raw_path = dest_path + ".raw"
    image.SaveFile(raw_path)
    with Image.open(raw_path) as im:
        im.convert("RGB").save(dest_path, "PNG")
    os.remove(raw_path)


def scan_one_page(device_id=None, session_id=None):
    """
    Scans a single page and appends it to a scan session (creating one if
    session_id is None/unknown). Returns {"session_id", "page_count",
    "preview"} — preview is a small base64 PNG data URI so the UI can show
    a thumbnail of what was just captured, same pattern as the cloud photo
    gallery's own thumbnails.
    """
    if session_id not in SCAN_SESSIONS:
        session_id = uuid.uuid4().hex
        SCAN_SESSIONS[session_id] = {"pages": [], "dir": tempfile.mkdtemp(prefix="officetool_scan_")}
    session = SCAN_SESSIONS[session_id]
    page_path = os.path.join(session["dir"], "page_{}.png".format(len(session["pages"]) + 1))

    pythoncom, win32com_client = _com()
    try:
        _capture_page_to_file(win32com_client, device_id, page_path)
    finally:
        pythoncom.CoUninitialize()
    session["pages"].append(page_path)

    # The full-resolution scan (a real document page can easily be 8-9MB as
    # a lossless PNG) is what goes into the final PDF — the UI only ever
    # needs a small thumbnail to confirm the right page got captured, so
    # build that separately rather than shipping the whole file over HTTP.
    with Image.open(page_path) as im:
        thumb = im.copy()
        thumb.thumbnail((240, 320))
        thumb_io = io.BytesIO()
        thumb.save(thumb_io, "JPEG", quality=70)
        preview = "data:image/jpeg;base64," + base64.b64encode(thumb_io.getvalue()).decode("ascii")
    return {"session_id": session_id, "page_count": len(session["pages"]), "preview": preview}


def remove_last_page(session_id):
    session = SCAN_SESSIONS.get(session_id)
    if not session or not session["pages"]:
        return {"page_count": 0}
    path = session["pages"].pop()
    try:
        os.remove(path)
    except OSError:
        pass
    return {"page_count": len(session["pages"])}


def cancel_session(session_id):
    session = SCAN_SESSIONS.pop(session_id, None)
    if session:
        _cleanup_dir(session["dir"])


def _cleanup_dir(path):
    try:
        for name in os.listdir(path):
            try:
                os.remove(os.path.join(path, name))
            except OSError:
                pass
        os.rmdir(path)
    except OSError:
        pass


def build_scan_filename(brand, do_number, company, kind="DO"):
    """Mirrors engine.py's own BRAND_TYPE_NUMBER..._DATE convention, marked
    SCANNED so it's obviously distinct from the original generated DO.
    kind defaults to "DO" (the original Submissions-linked flow, unchanged
    for every existing caller); the standalone Scanner tool (Menu -> Scanner,
    not tied to any submission) passes kind="SCAN" and company as a free-text
    label instead, so a generic scan doesn't read as if it were a delivery
    order."""
    today = datetime.date.today().isoformat()
    parts = [p for p in (brand, kind, str(do_number or "").strip(), "SCANNED") if p]
    safe_company = "".join(c if c.isalnum() or c in "-_ " else "" for c in (company or "")).strip().replace(" ", "-")
    if safe_company:
        parts.append(safe_company)
    parts.append(today)
    return "_".join(parts) + ".pdf"


def finalize_session(session_id, dest_folder, filename):
    """Combines every captured page (in scan order) into one PDF at
    dest_folder/filename, then cleans up the session's temp files. Returns
    the saved path. Raises ValueError if the session has no pages."""
    session = SCAN_SESSIONS.get(session_id)
    if not session or not session["pages"]:
        raise ValueError("No pages scanned yet.")
    os.makedirs(dest_folder, exist_ok=True)
    dest_path = os.path.join(dest_folder, filename)
    images = [Image.open(p).convert("RGB") for p in session["pages"]]
    first, rest = images[0], images[1:]
    first.save(dest_path, save_all=bool(rest), append_images=rest)
    for img in images:
        img.close()
    del SCAN_SESSIONS[session_id]
    _cleanup_dir(session["dir"])
    return dest_path
