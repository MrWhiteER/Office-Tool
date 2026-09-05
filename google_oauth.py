r"""
Google Sign-In (OAuth 2.0, "Desktop app" client type — RFC 8252 loopback
redirect) — lets a user click "Continue with Google" instead of typing a
password, once their account has been linked to a Google email (see
accounts.link_google_account()). Per explicit request, scoped down from a
much bigger "scan a QR code on your phone" idea (that would need a whole
new always-on relay server this app has never had) to this instead: one
click, opens the user's REAL system browser, approves on Google's own
page, comes back to this same app.

Google explicitly blocks OAuth requests coming from an embedded webview's
user agent (pywebview's WebView2 window included) — "disallowed_useragent"
— so this can never happen inside the app's own window. Python's
webbrowser.open() is what actually launches a real system browser here;
a JS window.open() from inside the app would just open ANOTHER embedded
webview window and get blocked the same way.

---- One-time setup (the admin does this once, in Google Cloud Console —
     nothing here can do this on your behalf, it needs your own Google
     account) ----
1. console.cloud.google.com -> select/create a project -> APIs & Services
   -> Credentials -> Create Credentials -> OAuth client ID -> application
   type "Desktop app".
2. Note the Client ID + Client Secret it gives you.
3. APIs & Services -> Credentials -> that client -> Authorized redirect
   URIs -> add exactly: http://127.0.0.1:5000/api/oauth/google-callback
   (5000 is this app's fixed port — see app.py's `port = ...` line; every
   install runs on the same port, so this one URI covers everyone).
4. Paste the Client ID + Secret into Admin Tools > Google Sign-In, Save,
   then "Publish to installer" (same dev-checkout-only bundling as
   photo_store.py's save_readonly_config() — the same two values ship in
   every future build; there's no per-user credential the way R2 has,
   since a Desktop-app OAuth client secret isn't confidential to begin
   with per RFC 8252).

---- The flow itself ----
1. /api/oauth/google-start (POST, {intent: "login"|"link"}) calls
   start_flow() below: builds a fresh PKCE code_verifier/challenge + a
   random state, remembers them in _PENDING (in-memory — this app only
   ever has ONE person mid-flow on THIS machine, nothing heavier is
   needed), opens the real system browser at Google's consent URL.
2. Google redirects the SYSTEM browser (a separate process/cookie jar
   from the app's own pywebview window!) to
   GET /api/oauth/google-callback?code=...&state=... — handle_callback()
   exchanges the code for a token and fetches the email from Google's
   userinfo endpoint, stashes the result on the matching _PENDING entry,
   and returns a plain "you can close this tab" HTML page. Deliberately
   does NOT touch the Flask session here — a cookie set on this response
   would land in the system browser's cookie jar, not the app window's.
3. The app's own window (a SEPARATE cookie jar) was polling
   /api/oauth/google-finish?state=... the whole time — once it sees
   status "done", THAT request (made by the app's own window) is what
   actually sets the Flask session (intent=login) or links the account
   (intent=link), because it's the one running in the right cookie
   context. See app.py's own api_oauth_google_finish().
"""
import base64
import hashlib
import json
import os
import secrets
import time
import urllib.error
import urllib.parse
import urllib.request
import webbrowser

import engine

GOOGLE_CONFIG_PATH = os.path.join(engine.BASE, "google_oauth.json")
AUTH_ENDPOINT = "https://accounts.google.com/o/oauth2/v2/auth"
TOKEN_ENDPOINT = "https://oauth2.googleapis.com/token"
USERINFO_ENDPOINT = "https://www.googleapis.com/oauth2/v3/userinfo"
# Fixed — matches app.py's own port = int(os.environ.get("PORT", 5000)).
REDIRECT_URI = "http://127.0.0.1:5000/api/oauth/google-callback"

_REQUIRED_KEYS = ("client_id", "client_secret")

_load_cfg = None
_save_cfg = None


def configure(load_cfg_fn, save_cfg_fn):
    global _load_cfg, _save_cfg
    _load_cfg = load_cfg_fn
    _save_cfg = save_cfg_fn


def _admin_cfg_block():
    return (_load_cfg() or {}).get("google_oauth") or {} if _load_cfg else {}


def _bundled_cfg_block():
    try:
        with open(GOOGLE_CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        return {}


def _cfg_block():
    # Same precedence as photo_store._cfg_block(): the admin's own local
    # config wins if fully filled in (so testing a NEW key from a dev
    # checkout doesn't need a rebuild first), else fall back to whatever
    # shipped bundled in this install.
    admin = _admin_cfg_block()
    if all(admin.get(k) for k in _REQUIRED_KEYS):
        return admin
    return _bundled_cfg_block()


def is_configured():
    c = _cfg_block()
    return all(c.get(k) for k in _REQUIRED_KEYS)


def get_public_config():
    """Safe-to-return-to-frontend view of the ADMIN's own config (Admin
    Tools screen) — never includes the secret."""
    c = _admin_cfg_block()
    return {"configured": is_configured(), "client_id": c.get("client_id", ""), "has_secret": bool(c.get("client_secret"))}


def save_config(client_id, client_secret):
    cfg = _load_cfg()
    block = cfg.setdefault("google_oauth", {})
    block["client_id"] = (client_id or "").strip()
    if client_secret:  # blank means "keep the existing secret" (edit form pattern)
        block["client_secret"] = client_secret.strip()
    _save_cfg(cfg)


def save_bundled_config():
    """Mirrors photo_store.save_readonly_config() exactly — dev-checkout
    only, writes google_oauth.json at the project root so build.bat's
    --add-data ships it in every future .exe."""
    import sys
    if getattr(sys, "frozen", False):
        raise RuntimeError("This only works from the project's own dev copy — the bundled config ships when you rebuild+republish the app (see build.bat), not from an installed .exe.")
    admin = _admin_cfg_block()
    if not all(admin.get(k) for k in _REQUIRED_KEYS):
        raise RuntimeError("Save a Client ID and Client Secret first.")
    with open(GOOGLE_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump({"client_id": admin["client_id"], "client_secret": admin["client_secret"]}, f, indent=2, ensure_ascii=False)


# ---- PKCE + pending-request bookkeeping ----
_PENDING = {}
_PENDING_TTL_SECONDS = 300  # 5 minutes — plenty of time to approve on Google's page


def _cleanup_pending():
    now = time.time()
    for k in [k for k, v in _PENDING.items() if now - v["created"] > _PENDING_TTL_SECONDS]:
        _PENDING.pop(k, None)


def start_flow(intent, username=None):
    """intent="login": no account required yet, matched by Google email
    once the callback lands. intent="link": username must be the
    CURRENTLY logged-in user — app.py checks the session BEFORE calling
    this and passes that username through, so google-finish's link step
    never has to re-derive "who is allowed to link this" from a
    possibly-different later request."""
    if not is_configured():
        raise RuntimeError("Google Sign-In isn't set up yet — ask your admin (Admin Tools > Google Sign-In).")
    _cleanup_pending()
    state = secrets.token_urlsafe(24)
    code_verifier = secrets.token_urlsafe(64)
    code_challenge = base64.urlsafe_b64encode(hashlib.sha256(code_verifier.encode()).digest()).decode().rstrip("=")
    _PENDING[state] = {
        "created": time.time(), "code_verifier": code_verifier,
        "intent": intent, "username": username, "status": "pending",
    }
    c = _cfg_block()
    params = {
        "client_id": c["client_id"], "redirect_uri": REDIRECT_URI, "response_type": "code",
        "scope": "openid email profile", "state": state,
        "code_challenge": code_challenge, "code_challenge_method": "S256",
        "access_type": "online", "prompt": "select_account",
    }
    webbrowser.open(AUTH_ENDPOINT + "?" + urllib.parse.urlencode(params))
    return state


def handle_callback(code, state):
    """Runs inside GET /api/oauth/google-callback — hit by the SYSTEM
    browser Google redirected, a different cookie jar from the app's own
    window (see module docstring). Only exchanges the code and records
    the result; never touches the Flask session. Returns (ok, message)
    for that route's own plain HTML response — success/failure detail
    for the person looking at the browser tab, not what drives the app."""
    entry = _PENDING.get(state)
    if not entry:
        return False, "This sign-in link has expired. Close this tab and try again from Office Tool."
    if not code:
        entry["status"] = "error"
        entry["error"] = "Google did not send back an authorization code."
        return False, entry["error"]
    c = _cfg_block()
    try:
        token_req = urllib.request.Request(
            TOKEN_ENDPOINT,
            data=urllib.parse.urlencode({
                "code": code, "client_id": c["client_id"], "client_secret": c["client_secret"],
                "redirect_uri": REDIRECT_URI, "grant_type": "authorization_code",
                "code_verifier": entry["code_verifier"],
            }).encode(),
            headers={"Content-Type": "application/x-www-form-urlencoded"},
        )
        with urllib.request.urlopen(token_req, timeout=15) as resp:
            token_data = json.loads(resp.read().decode("utf-8"))
        access_token = token_data["access_token"]
        info_req = urllib.request.Request(USERINFO_ENDPOINT, headers={"Authorization": "Bearer " + access_token})
        with urllib.request.urlopen(info_req, timeout=15) as resp:
            info = json.loads(resp.read().decode("utf-8"))
        email = (info.get("email") or "").strip().lower()
        if not email:
            raise ValueError("Google did not return an email address.")
        entry["status"] = "done"
        entry["email"] = email
        return True, "Signed in as " + email + " — you can close this tab and return to Office Tool."
    except Exception as e:
        entry["status"] = "error"
        entry["error"] = str(e)
        return False, "Something went wrong talking to Google: " + str(e)


def consume(state):
    """One-shot read: returns the _PENDING entry (or None) and removes it
    — called by app.py's /api/oauth/google-finish once status is no
    longer "pending", so a finished flow can never be replayed by hitting
    that endpoint again with the same state."""
    _cleanup_pending()
    entry = _PENDING.get(state)
    if not entry or entry["status"] == "pending":
        return entry
    return _PENDING.pop(state, None)
