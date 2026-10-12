r"""
Login/logout activity for Admin Tools > User Activity: who logged in or out,
when, on which PC, and who is using the app right now.

---- How it works ----
- Every app session (a login, or the app opening with a remembered login)
  is ONE small JSON object in the same R2 bucket accounts.json lives in:
  "system/activity/<YYYY-MM-DD>/<session id>.json" (date = UTC start day).
  Each PC only ever writes its own sessions' objects, so two PCs can never
  overwrite each other's records (unlike one shared log file would).
- While a session is open, a background loop rewrites its object every
  HEARTBEAT_SECONDS. The admin view decides "active now" from R2's own
  LastModified time on that object (R2's clock, not the PC's), so a PC
  that crashed, lost power or went offline drops out of "active" after
  ACTIVE_WINDOW_SECONDS instead of staying active forever.
- Log out / switching account / tray Exit write ended_at + end_reason.
  Closing the window only hides the app to the tray, so the user stays
  logged in (and active) until they log out or exit.
- Sessions older than KEEP_DAYS are deleted whenever an admin opens the view.
- No R2 configured: nothing is recorded (there is no other PC to show).
"""
import datetime
import json
import os
import socket
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor

import photo_store
from version import APP_VERSION

PREFIX = "system/activity/"
HEARTBEAT_SECONDS = 60
ACTIVE_WINDOW_SECONDS = 180
KEEP_DAYS = 90

_LOCK = threading.Lock()
_PUSH_LOCK = threading.Lock()
_LIVE = {}            # session id -> record, for sessions this process started
_DIRTY = set()        # session ids whose latest state isn't in R2 yet
_WAKE = threading.Event()
# key -> (etag, record): the admin view re-downloads only objects that changed.
_READ_CACHE = {}


def _now():
    return datetime.datetime.now(datetime.timezone.utc)


def _iso(dt):
    return dt.strftime("%Y-%m-%dT%H:%M:%SZ")


def pc_name():
    return os.environ.get("COMPUTERNAME") or socket.gethostname() or "Unknown PC"


def _windows_user():
    return os.environ.get("USERNAME") or os.environ.get("USER") or ""


def _key(rec):
    return PREFIX + rec["started_at"][:10] + "/" + rec["id"] + ".json"


def _push(sid):
    with _LOCK:
        rec = _LIVE.get(sid)
        if rec is None:
            return
        if not rec.get("ended_at"):
            rec["last_seen"] = _iso(_now())
        body = json.dumps(rec, ensure_ascii=False).encode("utf-8")
        key = _key(rec)
        ended = bool(rec.get("ended_at"))
    with _PUSH_LOCK:
        photo_store.put_bytes(key, body, "application/json", allow_bundled_write=True)
    if ended:
        with _LOCK:
            _LIVE.pop(sid, None)


def _flush():
    with _LOCK:
        sids = list(_DIRTY)
        _DIRTY.clear()
    for sid in sids:
        try:
            _push(sid)
        except Exception:
            with _LOCK:
                if sid in _LIVE:
                    _DIRTY.add(sid)  # retried on the next tick


def _loop():
    while True:
        _WAKE.wait(HEARTBEAT_SECONDS)
        _WAKE.clear()
        if not photo_store.is_configured():
            continue
        with _LOCK:
            _DIRTY.update(_LIVE.keys())
        _flush()


threading.Thread(target=_loop, daemon=True, name="activity-heartbeat").start()


def start(username, how="password"):
    """Opens a new session record and returns its id. how: "password",
    "google", or "resumed" (app opened with a remembered login)."""
    now = _iso(_now())
    sid = uuid.uuid4().hex
    rec = {"id": sid, "username": username, "pc": pc_name(), "windows_user": _windows_user(),
           "app_version": APP_VERSION, "how": how, "started_at": now, "last_seen": now,
           "ended_at": None, "end_reason": None}
    with _LOCK:
        _LIVE[sid] = rec
        _DIRTY.add(sid)
    _WAKE.set()
    return sid


def end(sid, reason):
    """reason: "logout", "switched account" or "app exit"."""
    with _LOCK:
        rec = _LIVE.get(sid)
        if rec is None or rec.get("ended_at"):
            return
        rec["ended_at"] = _iso(_now())
        rec["end_reason"] = reason
        _DIRTY.add(sid)
    _WAKE.set()


def is_live(sid):
    with _LOCK:
        return bool(sid) and sid in _LIVE and not _LIVE[sid].get("ended_at")


def ensure(flask_session):
    """Called on every logged-in request: a remembered login that survived
    an app restart has no live record in this process yet, so start one."""
    user = flask_session.get("user")
    if user and not is_live(flask_session.get("activity_id")):
        flask_session["activity_id"] = start(user, "resumed")


def end_all(reason="app exit", timeout=4.0):
    """Tray Exit/Restart: mark every open session ended and try to get
    that into R2 before the process dies (bounded, never blocks exit long)."""
    with _LOCK:
        for sid, rec in _LIVE.items():
            if not rec.get("ended_at"):
                rec["ended_at"] = _iso(_now())
                rec["end_reason"] = reason
                _DIRTY.add(sid)
    if not photo_store.is_configured():
        return
    t = threading.Thread(target=_flush, daemon=True)
    t.start()
    t.join(timeout)


def _list_objects():
    client = photo_store._client()
    bucket = photo_store._bucket()
    out, token = [], None
    while True:
        kwargs = {"Bucket": bucket, "Prefix": PREFIX}
        if token:
            kwargs["ContinuationToken"] = token
        resp = client.list_objects_v2(**kwargs)
        out.extend(resp.get("Contents", []))
        if not resp.get("IsTruncated"):
            return out
        token = resp.get("NextContinuationToken")


def _prune(objs):
    cutoff = (_now() - datetime.timedelta(days=KEEP_DAYS)).strftime("%Y-%m-%d")
    keep = []
    for o in objs:
        day = o["Key"][len(PREFIX):len(PREFIX) + 10]
        if day < cutoff:
            try:
                photo_store.delete_key(o["Key"], allow_bundled_write=True)
            except Exception:
                pass
            _READ_CACHE.pop(o["Key"], None)
        else:
            keep.append(o)
    return keep


def _read(obj):
    key, etag = obj["Key"], obj.get("ETag")
    hit = _READ_CACHE.get(key)
    if hit and hit[0] == etag:
        return hit[1]
    data, _ct = photo_store.get_bytes(key)
    rec = json.loads(data.decode("utf-8"))
    _READ_CACHE[key] = (etag, rec)
    return rec


def list_sessions(days=14):
    """Every session that started in the last `days` days, newest first,
    each with "status": "active", "logged_out" or "closed" (no heartbeat
    for ACTIVE_WINDOW_SECONDS: app closed, crashed, or the PC went offline)."""
    if not photo_store.is_configured():
        raise RuntimeError("No cloud connection is set up on this PC, so activity from other PCs can't be shown.")
    objs = _prune(_list_objects())
    since = (_now() - datetime.timedelta(days=days)).strftime("%Y-%m-%d")
    objs = [o for o in objs if o["Key"][len(PREFIX):len(PREFIX) + 10] >= since]
    with ThreadPoolExecutor(max_workers=8) as ex:
        recs = list(ex.map(lambda o: (o, _safe_read(o)), objs))
    now = _now()
    out = []
    for o, rec in recs:
        if not rec:
            continue
        rec = dict(rec)
        modified = o.get("LastModified")
        age = (now - modified).total_seconds() if modified else None
        if rec.get("ended_at"):
            rec["status"] = "logged_out"
        elif age is not None and age <= ACTIVE_WINDOW_SECONDS:
            rec["status"] = "active"
        else:
            rec["status"] = "closed"
        out.append(rec)
    out.sort(key=lambda r: r.get("started_at") or "", reverse=True)
    return out


def _safe_read(o):
    try:
        return _read(o)
    except Exception:
        return None
