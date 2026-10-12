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
- Every account change (admin Add/Save/Delete User, a user changing
  their own password) is applied to a FRESH copy pulled from R2, pushed
  straight back and read back to confirm, and only then cached locally
  (_apply_change()). If the cloud step fails the change is not made
  anywhere and the caller gets a real error. There used to be a separate
  "Publish" step, and saves that were never published silently stayed on
  the admin's PC only (fixed 2026-10-12).

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
   "brand_lock": [] | ["SOLOLUCE"] | ["SOLOLUCE", "ARTEMIS"] | ...,
   "blocked_tools": ["settings", "clients", "submissions", "statement", "alldocs"]}
- role "admin": brand_lock/blocked_tools are ignored entirely — full access.
- brand_lock: a LIST of brand codes this user can operate in (empty list
  or null = unrestricted, sees every brand) — per explicit request: "the
  locked to brand should have option to multiple brands" (was a single
  brand or none). Enforced server-side (app.py's /api/set-brand and
  /api/brands), not just hidden in the UI. _public() below always
  normalizes this to a real list even for a legacy record that still has
  the OLD single-string shape on disk (every real account had this unset
  by the time multi-brand shipped, so there was nothing to migrate in
  practice, but the read path stays tolerant of either shape regardless).
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
# Per-document-type access, independent of and finer-grained than the
# "alldocs" tool block above — per explicit request: "if i want for one
# user to be able to have access to the Invoices folder in all docs i
# will be able to give him the access or no... for everyone and every
# documents type". A user with "alldocs" in blocked_tools can't open All
# Docs at all regardless of this; a user who CAN open All Docs may still
# have specific document types hidden from it via this list. QTN2 covers
# both the current Quotation pipeline and legacy QTN (see /api/index's
# own "QTN2 matches both" comment) — one entry for what's conceptually
# one filter tab in the UI.
BLOCKABLE_DOC_TYPES = ("INV", "DO", "QTN2", "PI", "RV", "CN", "EXP", "CAT")
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
        "blocked_doc_types": [],
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


class CloudSyncError(Exception):
    """An account change could not be saved to R2, so no other PC would
    see it. Raised instead of reporting success: before 2026-10-12, Add /
    Save User only wrote this PC's local cache and waited for a separate
    "Publish" click, so the admin saw "Saved" while every other install
    kept rejecting the new user/password (real report from Edy)."""


# Last error from refresh_from_cloud(), or "" after a good refresh. Lets
# /api/login say "couldn't check the cloud" instead of a bare "wrong
# password" when this PC simply never received the newest accounts.json.
LAST_REFRESH_ERROR = ""


def _is_missing_key_error(e):
    code = ""
    try:
        code = (getattr(e, "response", None) or {}).get("Error", {}).get("Code", "")
    except Exception:
        pass
    return code in ("NoSuchKey", "404", "NotFound")


def _fetch_remote():
    """The current accounts.json in R2. Returns None when it has never been
    published (first-ever save); raises CloudSyncError on any other failure
    (offline, bad key, ...) so callers never write over R2 from a stale base."""
    try:
        data_bytes, _ct = photo_store.get_bytes(ACCOUNTS_KEY)
    except Exception as e:
        if _is_missing_key_error(e):
            return None
        raise CloudSyncError("Couldn't read the user list from the cloud: %s" % e)
    try:
        data = json.loads(data_bytes.decode("utf-8"))
    except Exception as e:
        raise CloudSyncError("The user list in the cloud is unreadable: %s" % e)
    if not isinstance(data, dict) or not isinstance(data.get("users"), list):
        raise CloudSyncError("The user list in the cloud has an unexpected shape.")
    return data


def _push_remote(data):
    """Writes accounts.json to R2 and reads it back to confirm it landed.
    allow_bundled_write=True: every caller is either admin-gated at the
    route level (/api/accounts-*) or touches only the caller's own record,
    and each one applies its change to a FRESH copy pulled from R2 — never
    a blind push of this PC's local cache, which is what the old admin-only
    restriction on publish_to_cloud() was guarding against."""
    body = json.dumps(data, indent=2, ensure_ascii=False).encode("utf-8")
    try:
        photo_store.put_bytes(ACCOUNTS_KEY, body, "application/json", allow_bundled_write=True)
    except Exception as e:
        raise CloudSyncError("Couldn't save the user list to the cloud: %s" % e)
    try:
        check = _fetch_remote()
    except CloudSyncError as e:
        raise CloudSyncError("Saved to the cloud but couldn't confirm it: %s" % e)
    if check != data:
        raise CloudSyncError("The cloud copy didn't match after saving — another PC may have changed users at the same moment. Try again.")


_SELF_SERVICE_FIELDS = ("settings", "google_email")


def _merge_pending(local, remote):
    """Local accounts.json with unpublished admin edits (DIRTY_FLAG set by
    an older version of this app) layered over the current R2 copy: the
    local user list and admin-managed fields win (that was the admin's
    intent), but each user's own self-service fields (theme/profile,
    linked Google account) are taken from R2, since those were saved
    directly to R2 from other PCs after this PC stopped syncing."""
    if not remote:
        return local
    remote_by_name = {(u.get("username") or "").strip().lower(): u for u in remote.get("users", [])}
    merged = json.loads(json.dumps(local))
    for u in merged.get("users", []):
        r = remote_by_name.get((u.get("username") or "").strip().lower())
        if not r:
            continue
        for f in _SELF_SERVICE_FIELDS:
            if f in r:
                u[f] = r[f]
    return merged


def _clear_dirty():
    try:
        if os.path.exists(DIRTY_FLAG):
            os.remove(DIRTY_FLAG)
    except Exception:
        pass


def has_pending_changes():
    return os.path.exists(DIRTY_FLAG)


def _apply_change(mutate):
    """Pull the freshest accounts.json from R2, apply `mutate(data)` to it,
    push it back and confirm, THEN update this PC's cache. Nothing is
    changed anywhere if the cloud step fails, so the caller can honestly
    report an error. Returns whatever `mutate` returns.

    No R2 connection configured at all (a standalone install that never
    synced) just edits the local cache — there's no other PC to reach."""
    if not photo_store.is_configured():
        data = load_accounts()
        result = mutate(data)
        _write_json(LOCAL_CACHE, data)
        return result
    remote = _fetch_remote()
    if has_pending_changes():
        data = _merge_pending(load_accounts(), remote)
    elif remote is not None:
        data = remote
    else:
        data = load_accounts()
    result = mutate(data)
    _push_remote(data)
    _write_json(LOCAL_CACHE, data)
    _clear_dirty()
    return result


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
    global LAST_REFRESH_ERROR
    if os.path.exists(DIRTY_FLAG):
        LAST_REFRESH_ERROR = ""
        return False
    if not photo_store.is_configured():
        LAST_REFRESH_ERROR = ""
        return False
    try:
        data = _fetch_remote()
    except CloudSyncError as e:
        LAST_REFRESH_ERROR = str(e)
        return False
    except Exception as e:
        LAST_REFRESH_ERROR = str(e)
        return False
    LAST_REFRESH_ERROR = ""
    if data is None:
        return False
    try:
        _write_json(LOCAL_CACHE, data)
    except Exception:
        return False
    return True


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


def _normalize_brand_lock(v):
    """Always returns a real list — tolerant of a legacy record that
    still has the OLD single-string shape (null or one bare brand code)
    on disk from before multi-brand locking existed. See this module's
    own docstring."""
    if not v:
        return []
    if isinstance(v, str):
        return [v]
    return list(v)


def _public(u):
    return {
        "username": u.get("username", ""),
        "role": u.get("role", "user"),
        "brand_lock": _normalize_brand_lock(u.get("brand_lock")),
        "blocked_tools": u.get("blocked_tools", []),
        "blocked_doc_types": u.get("blocked_doc_types", []),
        # Per-user preferences (currently just "theme") — see
        # save_user_setting() below. Included here so login and
        # /api/current-user hand it to the frontend for free, no extra
        # round-trip needed to apply it right away.
        "settings": u.get("settings") or {},
        # Which Google account (if any) can sign this user in — see
        # google_oauth.py + link_google_account() below. Never a secret,
        # safe to hand to the frontend so the profile popup can show
        # "Connected as x@gmail.com" / offer to disconnect.
        "google_email": u.get("google_email") or "",
    }


def list_users():
    return [_public(u) for u in load_accounts().get("users", [])]


def upsert_user(username, role, brand_lock, blocked_tools, password=None, blocked_doc_types=None):
    """Create or update a user. password=None on an edit keeps the
    existing hash. brand_lock: a list of brand codes (or a single bare
    string, accepted for backward compatibility — see
    _normalize_brand_lock()); invalid codes are silently dropped rather
    than rejecting the whole save, same treatment blocked_tools/
    blocked_doc_types already get below."""
    username = (username or "").strip()
    if not username:
        raise ValueError("Username required.")
    if role not in ("admin", "user"):
        raise ValueError("Invalid role.")
    brand_lock = [b for b in _normalize_brand_lock(brand_lock) if b in BRAND_CODES]
    blocked_tools = [t for t in (blocked_tools or []) if t in BLOCKABLE_TOOLS]
    blocked_doc_types = [t for t in (blocked_doc_types or []) if t in BLOCKABLE_DOC_TYPES]
    def mutate(data):
        users = data.setdefault("users", [])
        existing = next((u for u in users if u.get("username", "").lower() == username.lower()), None)
        if existing:
            existing["role"] = role
            existing["brand_lock"] = brand_lock
            existing["blocked_tools"] = blocked_tools
            existing["blocked_doc_types"] = blocked_doc_types
            if password:
                existing["password_hash"] = generate_password_hash(password)
            return _public(existing)
        if not password:
            raise ValueError("Password required for a new user.")
        users.append({
            "username": username,
            "password_hash": generate_password_hash(password),
            "role": role,
            "brand_lock": brand_lock,
            "blocked_tools": blocked_tools,
            "blocked_doc_types": blocked_doc_types,
        })
        return _public(users[-1])
    # Saved straight to R2 (see _apply_change()) — raises CloudSyncError
    # rather than pretending success if other PCs wouldn't see it.
    return _apply_change(mutate)


def delete_user(username):
    def mutate(data):
        users = data.get("users", [])
        remaining = [u for u in users if u.get("username", "").lower() != (username or "").strip().lower()]
        if len(remaining) == len(users):
            raise ValueError("No such user.")
        if not any(u.get("role") == "admin" for u in remaining):
            raise ValueError("Can't remove the last admin account.")
        data["users"] = remaining
    _apply_change(mutate)


def publish_to_cloud():
    """
    Pushes unpublished local edits up to R2. Adding, editing and deleting
    users now save to R2 directly (see _apply_change()), so this is only
    needed for edits an older version of the app left behind on this PC
    (DIRTY_FLAG still set). Merges them over the current R2 copy via
    _merge_pending() rather than blindly overwriting it, so users' own
    theme/profile/Google-link saves from other PCs survive.
    """
    if not photo_store.is_configured():
        return {"ok": False, "error": "No cloud connection is set up on this PC (Admin Tools > Cloud Storage)."}
    try:
        _apply_change(lambda data: None)
        return {"ok": True}
    except CloudSyncError as e:
        return {"ok": False, "error": str(e)}
    except Exception as e:
        return {"ok": False, "error": str(e)}


# Whitelist of settings any logged-in user can persist to their OWN
# account via save_user_setting() below — deliberately explicit rather
# than accepting an arbitrary key, so this can never become a backdoor
# for writing something unrelated into the shared accounts.json. "theme"
# was the first (per explicit request: "let the theme be linked to the
# user account"); "full_name"/"phone"/"personal_email"/"company_email"
# are the standard profile fields behind the account-avatar popup (per
# explicit request: "standard information to fill... phone details and
# etc." then "add more thing. Email Personal, company Email").
# "profile_picture_version" is NOT the picture itself (that's a separate
# R2 object — see photo_store.upload_profile_picture()) — just a plain
# timestamp string set whenever one is uploaded/removed, so the frontend
# knows whether to even try fetching /api/profile-picture at all, and can
# cache-bust the <img> src when it changes.
USER_SETTABLE_KEYS = ("theme", "full_name", "phone", "personal_email", "company_email", "profile_picture_version")


def save_user_setting(username, key, value):
    """
    Lets ANY logged-in user (not just the admin) persist one of their own
    settings so it follows them to every PC they log into — unlike
    publish_to_cloud() (a blind push of the whole local file, admin-
    only), this is safe for a regular user to trigger themselves: it can
    only ever touch USER_SETTABLE_KEYS inside THEIR OWN user record, and
    does so by pulling the freshest copy straight from R2 first (not the
    local cache, which can be up to 8s stale and isn't gated by
    DIRTY_FLAG here) and merging just that one field in — minimizing the
    odds of clobbering a concurrent edit from another PC, and never
    touching any other user's data or any other field of this user's own
    record (password_hash, role, brand_lock, blocked_tools).
    """
    if key not in USER_SETTABLE_KEYS:
        return {"ok": False, "error": "Unknown setting."}
    username_norm = (username or "").strip().lower()
    if not username_norm:
        return {"ok": False, "error": "Not logged in."}
    try:
        data_bytes, _ct = photo_store.get_bytes(ACCOUNTS_KEY)
        data = json.loads(data_bytes.decode("utf-8"))
        if not isinstance(data, dict) or not isinstance(data.get("users"), list):
            raise ValueError("unexpected shape")
    except Exception:
        # Not published yet / offline / R2 not configured — fall back to
        # whatever's cached locally so this doesn't hard-fail outright;
        # worst case the value only sticks on this one machine until the
        # next successful publish.
        data = load_accounts()
    users = data.setdefault("users", [])
    user = next((u for u in users if u.get("username", "").strip().lower() == username_norm), None)
    if not user:
        return {"ok": False, "error": "User not found."}
    settings = user.setdefault("settings", {})
    settings[key] = value
    _write_json(LOCAL_CACHE, data)  # apply locally right away regardless of whether the publish below succeeds
    try:
        photo_store.put_bytes(
            ACCOUNTS_KEY, json.dumps(data, indent=2, ensure_ascii=False).encode("utf-8"),
            "application/json", allow_bundled_write=True,
        )
        return {"ok": True}
    except Exception as e:
        # Saved locally at least — this machine keeps the new value even
        # if the publish itself failed (offline, R2 briefly down, etc.).
        return {"ok": False, "error": str(e)}


def change_own_password(username, current_password, new_password):
    """Lets a logged-in user change their OWN password from the account
    avatar popup — a step up in sensitivity from save_user_setting()
    above, so this ALWAYS re-verifies the current password against the
    freshest copy first (never trusts the session alone), and only then
    touches password_hash on that one user's own record. Same narrow
    pull-merge-push pattern as save_user_setting() otherwise — never
    touches role/brand_lock/blocked_tools, never any other user's record,
    allow_bundled_write=True for the same reason that function has it:
    the bundled read-only key really does carry write access today (see
    photo_store.py's own comment), and this is exactly the kind of
    narrow, self-service write that's safe to let every install make."""
    username_norm = (username or "").strip().lower()
    if not username_norm:
        return {"ok": False, "error": "Not logged in."}
    if not new_password or len(new_password) < 4:
        return {"ok": False, "error": "New password must be at least 4 characters."}
    def mutate(data):
        users = data.setdefault("users", [])
        user = next((u for u in users if u.get("username", "").strip().lower() == username_norm), None)
        if not user:
            raise ValueError("User not found.")
        if not check_password_hash(user.get("password_hash", ""), current_password or ""):
            raise ValueError("Current password is incorrect.")
        user["password_hash"] = generate_password_hash(new_password)
    try:
        _apply_change(mutate)
        return {"ok": True}
    except ValueError as e:
        return {"ok": False, "error": str(e)}
    except CloudSyncError as e:
        return {"ok": False, "error": str(e).rstrip(".") + ". Your password was NOT changed."}
    except Exception as e:
        return {"ok": False, "error": str(e)}


def find_user_by_google_email(google_email):
    """Used by app.py's /api/oauth/google-finish (intent="login") to turn
    a verified Google email back into an Office Tool account — only ever
    matches an account that was explicitly linked via
    link_google_account() below; Google sign-in never auto-creates one.
    Returns the PUBLIC shape (never password_hash), same as find_user()'s
    callers get by convention elsewhere in this module."""
    google_email = (google_email or "").strip().lower()
    if not google_email:
        return None
    for u in load_accounts().get("users", []):
        if (u.get("google_email") or "").strip().lower() == google_email:
            return _public(u)
    return None


def link_google_account(username, google_email):
    """Self-service — same narrow pull-freshest-copy-then-write-one-field
    pattern as save_user_setting()/change_own_password() above, just
    touching google_email instead. google_email="" unlinks. Refuses if
    that Google email is already linked to a DIFFERENT user (one Google
    account can't silently take over two Office Tool logins)."""
    username_norm = (username or "").strip().lower()
    if not username_norm:
        return {"ok": False, "error": "Not logged in."}
    google_email_norm = (google_email or "").strip().lower()
    if google_email_norm:
        existing = find_user_by_google_email(google_email_norm)
        if existing and existing.get("username", "").strip().lower() != username_norm:
            return {"ok": False, "error": "That Google account is already linked to a different Office Tool user."}
    try:
        data_bytes, _ct = photo_store.get_bytes(ACCOUNTS_KEY)
        data = json.loads(data_bytes.decode("utf-8"))
        if not isinstance(data, dict) or not isinstance(data.get("users"), list):
            raise ValueError("unexpected shape")
    except Exception:
        data = load_accounts()
    users = data.setdefault("users", [])
    user = next((u for u in users if u.get("username", "").strip().lower() == username_norm), None)
    if not user:
        return {"ok": False, "error": "User not found."}
    user["google_email"] = google_email_norm
    _write_json(LOCAL_CACHE, data)
    try:
        photo_store.put_bytes(
            ACCOUNTS_KEY, json.dumps(data, indent=2, ensure_ascii=False).encode("utf-8"),
            "application/json", allow_bundled_write=True,
        )
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}
