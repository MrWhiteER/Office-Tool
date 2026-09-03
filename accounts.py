r"""
Multi-user accounts + per-user permissions for Office Tool, synced across
every separate install via the SAME Cloudflare R2 bucket photo_store.py
already uses — one connection, no extra service, and no git/GitHub
involved at all (earlier versions of this file synced accounts.json
through the public GitHub repo; moved to R2 on 2026-09-03 once the admin
had R2 set up anyway, since it's both simpler — any install with a
write-capable R2 token can publish, not just a git checkout — and more
private, since R2 buckets aren't world-readable the way the GitHub repo is).

---- How it works ----
- The canonical accounts.json lives as one object (key "accounts.json")
  in the SAME R2 bucket configured in Settings > Shared Product Photos —
  same account_id/bucket, reusing photo_store.py's client (see
  photo_store.get_bytes()/put_bytes()). No separate credentials needed.
- Every running copy caches its own local copy at
  `<DATA_BASE>/accounts_cache.json` (see engine.py's BASE/DATA_BASE split)
  and refreshes it from R2 on every login attempt AND on app launch
  (best-effort — falls back to the last cached copy if offline/not yet
  configured, so login still works without internet once an install has
  synced at least once). Refreshing on every login (not just launch)
  means a brand-new user the admin just added can log in within seconds,
  no restart needed anywhere.
- Any install with WRITE-capable R2 credentials (the admin's own token,
  see photo_store.py) can publish — not just a git checkout. Everyone
  else (read-only R2 token) can only read.

---- Security note (be upfront about this with the admin) ----
Passwords are salted + hashed (werkzeug's PBKDF2-SHA256, same as Flask's
own recommended default) — never stored or transmitted in plain text.
Storing this in R2 rather than the public GitHub repo means it's no
longer world-downloadable — R2 buckets are private by default and every
reader needs a real (even if read-only) R2 credential to fetch anything,
same access boundary as the photo library.

---- Permission model ----
Each user record:
  {"username": "...", "password_hash": "...", "role": "admin"|"user",
   "brand_lock": null | "SOLOLUCE" | "ARTEMIS" | "ADS" | "WATT",
   "blocked_tools": ["settings", "clients", "submissions", "statement", "alldocs"]}
- role "admin": brand_lock/blocked_tools are ignored entirely — full access.
- brand_lock: non-admin user can only ever operate in that one brand;
  enforced server-side in app.py (before_request), not just hidden in the UI.
- blocked_tools: which of the 5 non-document-generation views this user
  can't open — also enforced server-side via URL-prefix checks, not just a
  hidden nav button.
"""
import json
import os

from werkzeug.security import generate_password_hash, check_password_hash

import engine
import photo_store

# "system/" is a deliberate namespace for app data, kept separate from the
# photo library's own key structure (which mirrors the admin's local
# folders, e.g. "2. Index Pictures/2. Outdoor Collection/.../STAGNA-MS.png")
# so the R2 dashboard's file browser reads as two clearly distinct trees —
# your content vs. the app's own bookkeeping — instead of one flat mix.
# Any future app-level data (not a product photo) should live under this
# same "system/" prefix.
ACCOUNTS_KEY = "system/accounts.json"
# Marks "this machine has local account edits not yet published to R2" —
# see save_accounts()/publish_to_cloud()/refresh_from_cloud() below. Without
# this, an admin who adds a user and logs in again (or anyone else logs in)
# BEFORE clicking Publish would have their own unpublished edit silently
# overwritten by the older remote copy — refresh_from_cloud() runs on
# every login attempt (see verify_login()), so this bit has to exist.
DIRTY_FLAG = os.path.join(engine.DATA_BASE, "accounts_dirty.flag")
LOCAL_CACHE = os.path.join(engine.DATA_BASE, "accounts_cache.json")

BLOCKABLE_TOOLS = ("settings", "clients", "submissions", "statement", "alldocs")
BRAND_CODES = tuple(engine.BRANDS.keys())


def _default_accounts():
    # First-run fallback so the app still works before accounts.json has
    # ever been published: a single admin account, username/password both
    # "admin" — deliberately weak, meant to be changed immediately via the
    # Settings > Users screen (see app.py's /api/accounts-save).
    return {"users": [{
        "username": "admin",
        "password_hash": generate_password_hash("admin"),
        "role": "admin",
        "brand_lock": None,
        "blocked_tools": [],
    }]}


def _read_json(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _write_json(path, data):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def refresh_from_cloud():
    """
    Best-effort pull of the latest accounts.json from R2. Never raises —
    R2 not configured yet, offline, or the object not existing yet all
    just mean "keep using the last cached copy" (or the hardcoded
    single-admin default, on a machine that's never synced at all).

    Skips entirely if this machine has unpublished local edits (see
    DIRTY_FLAG) — otherwise an admin's own in-progress edit would get
    silently clobbered by the older remote copy the very next time anyone
    (including that same admin) logs in, since this runs on every login.
    """
    if os.path.exists(DIRTY_FLAG):
        return False
    if not photo_store.is_configured():
        return False
    try:
        data_bytes, _ct = photo_store.get_bytes(ACCOUNTS_KEY)
        data = json.loads(data_bytes.decode("utf-8"))
        if isinstance(data, dict) and isinstance(data.get("users"), list):
            _write_json(LOCAL_CACHE, data)
            return True
    except Exception:
        pass
    return False


def load_accounts():
    data = _read_json(LOCAL_CACHE)
    if data is None or not isinstance(data.get("users"), list):
        data = _default_accounts()
    return data


def save_accounts(data):
    """Writes the local cache immediately (so the admin sees their own
    edit right away) and marks it dirty — publish_to_cloud() is the
    separate, explicit step that both makes it visible to everyone else
    AND clears the dirty flag (see refresh_from_cloud())."""
    _write_json(LOCAL_CACHE, data)
    try:
        with open(DIRTY_FLAG, "w", encoding="utf-8") as f:
            f.write("1")
    except Exception:
        pass  # worst case: a later refresh could overwrite this edit — not fatal, just annoying


def find_user(username):
    username = (username or "").strip().lower()
    for u in load_accounts().get("users", []):
        if u.get("username", "").strip().lower() == username:
            return u
    return None


def verify_login(username, password):
    """Returns the user record (never including password_hash) on success, else None.
    Refreshes from R2 first (see refresh_from_cloud()'s own comment on why
    this runs on every login, not just app launch) so a user the admin
    JUST added/edited can log in immediately without anyone restarting."""
    refresh_from_cloud()
    u = find_user(username)
    if not u or not password:
        return None
    if not check_password_hash(u.get("password_hash", ""), password):
        return None
    return _public(u)


def _public(u):
    return {
        "username": u.get("username", ""),
        "role": u.get("role", "user"),
        "brand_lock": u.get("brand_lock"),
        "blocked_tools": u.get("blocked_tools", []),
    }


def list_users():
    return [_public(u) for u in load_accounts().get("users", [])]


def upsert_user(username, role, brand_lock, blocked_tools, password=None):
    """Create or update a user. password=None on an edit keeps the existing hash."""
    username = (username or "").strip()
    if not username:
        raise ValueError("Username required.")
    if role not in ("admin", "user"):
        raise ValueError("Invalid role.")
    if brand_lock and brand_lock not in BRAND_CODES:
        raise ValueError("Invalid brand.")
    blocked_tools = [t for t in (blocked_tools or []) if t in BLOCKABLE_TOOLS]
    data = load_accounts()
    users = data.setdefault("users", [])
    existing = next((u for u in users if u.get("username", "").lower() == username.lower()), None)
    if existing:
        existing["role"] = role
        existing["brand_lock"] = brand_lock or None
        existing["blocked_tools"] = blocked_tools
        if password:
            existing["password_hash"] = generate_password_hash(password)
    else:
        if not password:
            raise ValueError("Password required for a new user.")
        users.append({
            "username": username,
            "password_hash": generate_password_hash(password),
            "role": role,
            "brand_lock": brand_lock or None,
            "blocked_tools": blocked_tools,
        })
    save_accounts(data)
    return _public(existing or users[-1])


def delete_user(username):
    data = load_accounts()
    users = data.get("users", [])
    remaining = [u for u in users if u.get("username", "").lower() != (username or "").strip().lower()]
    if len(remaining) == len(users):
        raise ValueError("No such user.")
    if not any(u.get("role") == "admin" for u in remaining):
        raise ValueError("Can't remove the last admin account.")
    data["users"] = remaining
    save_accounts(data)


def publish_to_cloud():
    """
    Pushes the local accounts.json up to R2 so every other install picks
    up the change (they refresh on every login attempt — see
    verify_login()). Needs a WRITE-capable R2 token (the admin's own —
    see photo_store.py); a read-only token gets a clear permission error
    back from R2 rather than silently doing nothing.
    """
    if not photo_store.is_configured(require_write=True):
        return {"ok": False, "error": "Set up your Admin key (Admin Tools > Cloud Storage) first — accounts sync uses the same connection."}
    try:
        data = load_accounts()
        photo_store.put_bytes(ACCOUNTS_KEY, json.dumps(data, indent=2, ensure_ascii=False).encode("utf-8"), "application/json")
        if os.path.exists(DIRTY_FLAG):
            os.remove(DIRTY_FLAG)
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}
