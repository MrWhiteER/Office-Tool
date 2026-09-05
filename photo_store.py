r"""
Shared product-photo storage on Cloudflare R2 — free (first 10GB storage,
zero egress/bandwidth cost ever, no matter how many installs download how
often), so every separate install can share the same product photo library
without a rented server. Wired in as a background sync layer that fills
each machine's own LOCAL product_photos_folder (see app.py's SETTINGS_FIELDS
and engine.match_product_photo) — the existing filename-matching code that
looks up a product's photo for a datasheet is completely unchanged; it just
finds files that got there via R2 instead of a manual copy.

---- One admin-managed key, two storage locations (as of 2026-09-03) ----
Originally every install — admin or not — pasted SOME R2 credential into
Settings themselves. Changed on explicit request: Cloudflare settings
should be admin-only to even see, and non-admins should need to do
nothing at all. A brief in-between design used two DIFFERENT keys (an
admin read/write one + a separate read-only one for everyone else) — the
admin then said just use ONE standard key everywhere, so:

The single key the admin enters in Admin Tools gets saved to BOTH:
1. THIS machine's config.json (`cfg["r2_photo_store"]`) — used directly
   here, per-machine, never synced anywhere, never committed.
2. `r2_readonly.json` at the project root — picked up by build.bat's
   --add-data so it ships INSIDE every future .exe/installer. Every
   non-admin install falls back to this automatically for everything
   (sync, the cloud photo gallery, pulling accounts.json) — they never
   see a Cloudflare field at all. Only takes effect for OTHER installs
   once they're on a build made after the admin saves it, so rotating the
   key needs a rebuild+republish (see build.bat) — accepted as fine since
   this is a rare, one-time-ish setup step, unlike accounts (which change
   often and sync live via R2 itself instead). Loaded from engine.BASE
   (bundled resource root — read-only at runtime once frozen, same as
   templates/static), NOT engine.DATA_BASE. Despite the filename (kept
   from the earlier two-key design), this now holds the SAME key as #1,
   not a separately-scoped one — the module still exposes a
   `require_write` distinction on every call (see _cfg_block() below)
   because it's a cheap, harmless safety net if the two ever do diverge
   again, not because they're expected to today.

   They DID diverge once (2026-09-05, real incident): the admin's own
   installed .exe had an older secret cached in ITS OWN config.json from
   before the key was last rotated, while the bundled r2_readonly.json
   (and the dev checkout's config.json) had the current one — since the
   admin's own local key always wins over the bundle when present (see
   _cfg_block()), every R2 call on that one install failed with
   SignatureDoesNotMatch, even reads, even though every OTHER install
   (using the bundled key) was fine. Fixed by hand (copying the correct
   secret into that install's config.json) — there's no auto-repair for
   this today; if it happens again, compare the two files directly.

---- 10GB hard stop ----
Every upload checks current bucket usage + the new file's size against
HARD_LIMIT_BYTES first and is rejected client-side (before any data
transfer) if it would cross the line — so this never silently becomes a
paid Cloudflare account.
"""
import datetime
import json
import os
import sys
import mimetypes

import engine

HARD_LIMIT_BYTES = 10 * 1024 * 1024 * 1024  # 10 GB — see module docstring

# App data (currently just accounts.py's accounts.json) lives under this
# prefix in the SAME bucket as the photo library, deliberately separate
# from the photo keys (which mirror the admin's own local folder
# structure) — see accounts.py's own comment. Every function here that
# enumerates "the photo library" (usage, the gallery list, sync-down)
# excludes this prefix so app bookkeeping never shows up as a "photo".
SYSTEM_PREFIX = "system/"

# Image extensions actual photos can have — used by list_photos() below as
# defense in depth alongside the SYSTEM_PREFIX/DOCUMENTS_PREFIX exclusions,
# so any FUTURE top-level key prefix (another non-photo feature added later
# and forgetting to update this list) still can't leak into the photo
# gallery just by not matching a known non-photo prefix.
IMAGE_EXTENSIONS = (".png", ".jpg", ".jpeg", ".webp", ".gif")

# Bundled read-only fallback — see module docstring's "Two credential
# tiers". Lives at the project root (like VERSION) so build.bat's
# --add-data picks it up; NOT in .gitignore's sense of "never commit" the
# way secret_key.txt is, but IS gitignored anyway (see .gitignore) since
# it's still a real credential and doesn't belong in the PUBLIC repo even
# though it ships inside the compiled installer — those are different
# audiences (anyone browsing GitHub vs. someone reverse-engineering a
# .exe they already chose to run).
READONLY_CONFIG_PATH = os.path.join(engine.BASE, "r2_readonly.json")

CONFIG = None  # set by app.py via configure(), avoids a circular import on load_cfg/save_cfg


def configure(load_cfg_fn, save_cfg_fn):
    global _load_cfg, _save_cfg
    _load_cfg = load_cfg_fn
    _save_cfg = save_cfg_fn


def _admin_cfg_block():
    return (_load_cfg() or {}).get("r2_photo_store") or {}


def _readonly_cfg_block():
    try:
        with open(READONLY_CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        return {}


_REQUIRED_KEYS = ("account_id", "access_key_id", "secret_access_key", "bucket")


def _cfg_block(require_write=False, allow_bundled_write=False):
    """The admin's own local (write-capable) config always wins if fully
    filled in — including on the admin's OWN machine for read operations,
    so they're always working against their own real credentials rather
    than a possibly-stale bundled fallback. require_write=True normally
    never falls back to the bundled key — but allow_bundled_write=True
    (document backups, see upload_document() below) deliberately opts
    back in: per this module's own docstring, the bundled r2_readonly.json
    holds the SAME value as the admin's write key today (there's only
    ever one key, mirrored to both places), so this isn't opening a new
    hole, it's a conscious choice to let every install use write access
    that's already technically sitting in their own install folder,
    specifically for the one feature (document backup) that needs every
    install — not just the admin's — to be able to write."""
    admin = _admin_cfg_block()
    if all(admin.get(k) for k in _REQUIRED_KEYS):
        return admin
    if require_write and not allow_bundled_write:
        return {}
    return _readonly_cfg_block()


def is_configured(require_write=False):
    c = _cfg_block(require_write=require_write)
    return all(c.get(k) for k in _REQUIRED_KEYS)


def get_public_config():
    """Safe-to-return-to-frontend view of the ADMIN's own write-capable
    config (Admin Tools screen) — never includes the secret key."""
    c = _admin_cfg_block()
    return {
        "configured": is_configured(require_write=True),
        "account_id": c.get("account_id", ""),
        "bucket": c.get("bucket", ""),
        "access_key_id": c.get("access_key_id", ""),
        "has_secret": bool(c.get("secret_access_key")),
    }


def save_config(account_id, access_key_id, secret_access_key, bucket):
    cfg = _load_cfg()
    block = cfg.setdefault("r2_photo_store", {})
    block["account_id"] = (account_id or "").strip()
    block["access_key_id"] = (access_key_id or "").strip()
    block["bucket"] = (bucket or "").strip()
    if secret_access_key:  # blank means "keep the existing secret" (edit form pattern)
        block["secret_access_key"] = secret_access_key.strip()
    _save_cfg(cfg)


def get_resolved_admin_config():
    """The admin's config INCLUDING the real secret — never exposed via
    any API response (get_public_config() masks it), only used internally
    right after save_config() to mirror the exact same, already-resolved
    key into the bundle (see /api/photostore-config's own comment) —
    this way leaving the secret field blank ("keep existing") mirrors
    correctly too, instead of the bundle's own independent blank-stays-
    blank fallback maybe never getting seeded with a real secret at all."""
    return dict(_admin_cfg_block())


def save_readonly_config(account_id, access_key_id, secret_access_key, bucket):
    """Mirrors the same key into the bundled r2_readonly.json — called
    right after save_config() with the identical values (see
    /api/photostore-config's own comment). Only meaningful on the admin's
    own dev checkout (engine.BASE is a read-only extracted temp dir once
    frozen) — same constraint accounts.publish_to_cloud() used to have
    with git. Raises a plain Exception with a clear message otherwise;
    caller (app.py) reports that separately without failing the whole save."""
    if getattr(sys, "frozen", False):
        raise RuntimeError("This only works from the project's own dev copy — the bundled key ships when you rebuild+republish the app (see build.bat), not from an installed .exe.")
    existing = _readonly_cfg_block()
    data = {
        "account_id": (account_id or "").strip() or existing.get("account_id", ""),
        "access_key_id": (access_key_id or "").strip() or existing.get("access_key_id", ""),
        "secret_access_key": (secret_access_key or "").strip() or existing.get("secret_access_key", ""),
        "bucket": (bucket or "").strip() or existing.get("bucket", ""),
    }
    with open(READONLY_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def _client(require_write=False, allow_bundled_write=False):
    import boto3
    c = _cfg_block(require_write=require_write, allow_bundled_write=allow_bundled_write)
    if not all(c.get(k) for k in _REQUIRED_KEYS):
        raise RuntimeError(
            "No write-capable R2 key is set up (Admin Tools)." if require_write
            else "Shared photo storage isn't set up yet — ask your admin."
        )
    endpoint = "https://{}.r2.cloudflarestorage.com".format(c["account_id"])
    return boto3.client(
        "s3",
        endpoint_url=endpoint,
        aws_access_key_id=c["access_key_id"],
        aws_secret_access_key=c["secret_access_key"],
        region_name="auto",
    )


def _bucket(require_write=False, allow_bundled_write=False):
    return _cfg_block(require_write=require_write, allow_bundled_write=allow_bundled_write)["bucket"]


def get_usage():
    """Returns {bytes_used, count, limit_bytes} by listing every object —
    R2/S3 has no cheap single "bucket size" call, so this pages through the
    listing. Fine at photo-library scale (thousands of objects, not
    millions); each page is 1000 objects."""
    client = _client()
    bucket = _bucket()
    total = 0
    count = 0
    token = None
    while True:
        kwargs = {"Bucket": bucket}
        if token:
            kwargs["ContinuationToken"] = token
        resp = client.list_objects_v2(**kwargs)
        for obj in resp.get("Contents", []):
            if obj["Key"].startswith(SYSTEM_PREFIX) or obj["Key"].startswith(DOCUMENTS_PREFIX):
                continue
            total += obj["Size"]
            count += 1
        if not resp.get("IsTruncated"):
            break
        token = resp.get("NextContinuationToken")
    return {"bytes_used": total, "count": count, "limit_bytes": HARD_LIMIT_BYTES}


def get_photo_bytes(key):
    """One photo's raw bytes + content type, for the cloud gallery picker
    (thumbnails AND the full-res click-to-select — see app.py's
    /api/photostore-fetch and openCloudPhotoPicker() in the page script).
    Any logged-in user can call this (read-only), not just the admin."""
    return get_bytes(key)


def get_bytes(key):
    """Generic read of one object's raw bytes + content type — also used
    by accounts.py to store accounts.json in this SAME bucket (a tiny JSON
    object, not a photo, but reusing this one already-configured R2
    connection instead of standing up a second sync channel). Raises on
    any failure (missing key, bad credentials, etc.) — callers decide how
    to handle "not there yet" vs a real error."""
    client = _client()
    resp = client.get_object(Bucket=_bucket(), Key=key)
    return resp["Body"].read(), resp.get("ContentType") or "application/octet-stream"


def put_bytes(key, data, content_type="application/octet-stream", allow_bundled_write=False):
    """Generic write, bypassing the 10GB photo-library check (upload_photo
    enforces that for actual photos; accounts.json is a few KB and isn't
    part of that budget) — used by accounts.py's publish_to_cloud() (full-
    file admin publish, allow_bundled_write left False on purpose — see
    _cfg_block()'s own comment on why that stays admin-only) and by
    accounts.py's save_user_setting() (allow_bundled_write=True — a
    regular user persisting their OWN settings is the same narrow, safe
    case upload_document() already opened this up for)."""
    client = _client(require_write=True, allow_bundled_write=allow_bundled_write)
    client.put_object(Bucket=_bucket(require_write=True, allow_bundled_write=allow_bundled_write), Key=key, Body=data, ContentType=content_type)


# Who uploaded each photo — per explicit request: "I want every user to
# have his own place in cloudflare... the system will see it all in one
# just somewhere it will be mentioned that its created or imported... by
# which user" and "it will show as groups, by whom it was uploaded". A
# single small index file (not per-photo R2 object Metadata, which would
# need one HEAD request PER PHOTO to read back for a 575+-photo gallery —
# far too slow) mapping key -> {uploaded_by, uploaded_at}, same
# read-modify-write-then-reupload shape as accounts.json. Lives under
# SYSTEM_PREFIX (app bookkeeping, not a photo) so it's excluded from
# get_usage()/list_photos() the same way accounts.json already is.
PHOTO_ATTRIBUTION_KEY = SYSTEM_PREFIX + "photo_attribution.json"


def _read_photo_attribution():
    try:
        data_bytes, _ct = get_bytes(PHOTO_ATTRIBUTION_KEY)
        data = json.loads(data_bytes.decode("utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def get_uploader(key):
    """Who uploaded this specific photo, or None if it predates
    attribution tracking (the admin's original library, or an upload that
    happened before this index existed). Used by app.py's Cloud Manager
    tool (/api/photostore-delete) to let a normal user delete only photos
    THEY uploaded — anything with no recorded uploader stays admin-only to
    remove, same as before this existed."""
    return (_read_photo_attribution().get(key) or {}).get("uploaded_by")


# One profile picture per user, under SYSTEM_PREFIX like accounts.json —
# app bookkeeping, not part of the shared product-photo library (excluded
# from list_photos()/get_usage() the same way). Per explicit request:
# "make it so the user can upload his own profile picture." Always a
# .png (app.py normalizes whatever the user picked before uploading), one
# stable key per username so the frontend never needs to look up a URL —
# it just requests /api/profile-picture?username=X and gets a 404 if
# that person never set one.
PROFILE_PICTURE_PREFIX = SYSTEM_PREFIX + "profile_pictures/"


def profile_picture_key(username):
    return PROFILE_PICTURE_PREFIX + (username or "").strip().lower() + ".png"


def upload_profile_picture(local_path, username):
    """Any logged-in user can set their OWN profile picture — narrow,
    self-service write, same allow_bundled_write=True reasoning as
    save_user_setting() (a regular user's bundled read-only key really
    does carry write access today — see put_bytes()'s own comment)."""
    with open(local_path, "rb") as f:
        put_bytes(profile_picture_key(username), f.read(), "image/png", allow_bundled_write=True)


def get_profile_picture(username):
    """Raises (like get_photo_bytes()) if this user never set one, or the
    connection isn't configured — app.py's route turns that into a plain
    404 rather than a 502, since "no picture yet" is the overwhelmingly
    common case, not an error."""
    return get_bytes(profile_picture_key(username))


def delete_profile_picture(username):
    client = _client(require_write=True, allow_bundled_write=True)
    client.delete_object(Bucket=_bucket(require_write=True, allow_bundled_write=True), Key=profile_picture_key(username))


# Which Sololuce Datasheet photo zone (main/lifestyle/diagram/extra1/
# extra2/extra3 — see app.py's CAT_IMG_ZONE_NAME) a photo was uploaded
# FOR, via the globe-icon menu next to that zone — per explicit request:
# "when the user clicks the globe in main product photo and if he adds a
# picture it will save in cloudflare in Main Product Picture and if it's
# Application Photo it will add in Application Photo. same with the
# rest!" Same one-small-shared-index shape as PHOTO_ATTRIBUTION_KEY
# (not per-object R2 Metadata, for the same "one HEAD request per photo"
# reason) — deliberately a SEPARATE index rather than folding "zone" into
# the attribution entry, since a photo can be uploaded through the
# regular Cloud Manager/Settings path (no zone at all) just as easily as
# through a specific zone's globe menu; most photos will never have one.
PHOTO_ZONE_KEY = SYSTEM_PREFIX + "photo_zones.json"


def _read_photo_zones():
    try:
        data_bytes, _ct = get_bytes(PHOTO_ZONE_KEY)
        data = json.loads(data_bytes.decode("utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def get_zone(key):
    """Which zone slot this photo was uploaded for (e.g. "lifestyle"), or
    None if it wasn't uploaded through a zone's globe menu at all — the
    overwhelming majority of the library. Used by app.py's cloud photo
    picker to default-filter to just this zone's own previously-uploaded
    photos when opened from a specific slot."""
    return _read_photo_zones().get(key)


def _record_photo_zone(key, zone):
    """Best-effort, same shape as _record_photo_attribution() — a failed
    write here should never fail the upload itself."""
    try:
        zones = _read_photo_zones()
        zones[key] = zone
        put_bytes(PHOTO_ZONE_KEY, json.dumps(zones, indent=2, ensure_ascii=False).encode("utf-8"),
                   "application/json", allow_bundled_write=True)
    except Exception:
        pass


def _record_photo_attribution(key, username):
    """Best-effort — a failed attribution write should never fail the
    actual photo upload it's attached to. allow_bundled_write=True: any
    logged-in user recording who THEY uploaded is the same narrow, safe
    case documents already opened this up for (see put_bytes' own
    comment)."""
    try:
        attribution = _read_photo_attribution()
        attribution[key] = {"uploaded_by": username, "uploaded_at": datetime.datetime.now().isoformat()}
        put_bytes(PHOTO_ATTRIBUTION_KEY, json.dumps(attribution, indent=2, ensure_ascii=False).encode("utf-8"),
                   "application/json", allow_bundled_write=True)
    except Exception:
        pass


def list_photos():
    client = _client()
    bucket = _bucket()
    out = []
    token = None
    while True:
        kwargs = {"Bucket": bucket}
        if token:
            kwargs["ContinuationToken"] = token
        resp = client.list_objects_v2(**kwargs)
        for obj in resp.get("Contents", []):
            key = obj["Key"]
            if key.startswith(SYSTEM_PREFIX) or key.startswith(DOCUMENTS_PREFIX):
                continue
            if not key.lower().endswith(IMAGE_EXTENSIONS):
                continue
            out.append({"key": key, "size": obj["Size"], "modified": obj["LastModified"].isoformat()})
        if not resp.get("IsTruncated"):
            break
        token = resp.get("NextContinuationToken")
    # Merge in who-uploaded-what — one read of the small shared index,
    # not one request per photo. Photos uploaded before this feature
    # existed (the admin's original 575) simply have no entry, which
    # callers should treat as "uploaded by the admin" (they were, before
    # any of this existed) rather than "unknown".
    attribution = _read_photo_attribution()
    zones = _read_photo_zones()
    for o in out:
        info = attribution.get(o["key"])
        if info:
            o["uploaded_by"] = info.get("uploaded_by")
            o["uploaded_at"] = info.get("uploaded_at")
        zone = zones.get(o["key"])
        if zone:
            o["zone"] = zone
    return sorted(out, key=lambda o: o["key"].lower())


def upload_photo(local_path, filename, running_usage=None, uploaded_by=None, allow_bundled_write=False, zone=None):
    """
    Used to be strictly admin-only; per explicit request, any logged-in
    user can now upload new product photos (allow_bundled_write=True from
    the caller — see put_bytes()'s own comment on why this is a
    conscious, narrow choice rather than an accidental hole). Deletion
    (delete_photo() below) deliberately stays admin-only — "anyone can
    add, only the admin removes" is a reasonable moderation default that
    was never asked to change. Rejects BEFORE transferring any data if
    the upload would cross HARD_LIMIT_BYTES.

    running_usage: pass a single mutable {"bytes_used": N} dict shared
    across a whole batch of uploads (a folder upload can be hundreds of
    files) so this only lists the entire bucket ONCE per batch — via
    get_usage() — instead of once per file, and just adds each file's size
    to that running total as it succeeds. Omit it (None) to have this
    function fetch fresh usage itself, for a single one-off upload.
    """
    size = os.path.getsize(local_path)
    if running_usage is None:
        running_usage = get_usage()
    if running_usage["bytes_used"] + size > HARD_LIMIT_BYTES:
        free = HARD_LIMIT_BYTES - running_usage["bytes_used"]
        raise ValueError(
            "That would exceed the free 10GB limit (only {:.0f} MB left) — "
            "remove some unused photos first, or this stays a paid Cloudflare "
            "feature which the admin chose not to enable.".format(max(free, 0) / 1024 / 1024)
        )
    client = _client(require_write=True, allow_bundled_write=allow_bundled_write)
    content_type = mimetypes.guess_type(filename)[0] or "application/octet-stream"
    with open(local_path, "rb") as f:
        client.put_object(Bucket=_bucket(require_write=True, allow_bundled_write=allow_bundled_write), Key=filename, Body=f, ContentType=content_type)
    running_usage["bytes_used"] += size
    if uploaded_by:
        _record_photo_attribution(filename, uploaded_by)
    if zone:
        _record_photo_zone(filename, zone)


def delete_photo(filename, allow_bundled_write=False):
    """allow_bundled_write=True: needed for the Cloud Manager tool's
    "delete your own upload" case (app.py's api_photostore_delete()) — a
    genuine non-admin user's machine only ever has the bundled read-only
    config, never r2_photo_store, so without this every one of their
    deletes would fail with "No write-capable R2 key is set up" even
    though the permission check already passed. Same reasoning
    upload_photo() already has this for — see its own comment."""
    client = _client(require_write=True, allow_bundled_write=allow_bundled_write)
    client.delete_object(Bucket=_bucket(require_write=True, allow_bundled_write=allow_bundled_write), Key=filename)


# Generated documents (Quotations, Tax Invoices, Delivery Orders, Sololuce
# Datasheets, Expense Reports) live under this prefix — separate from both
# the photo keys (SYSTEM_PREFIX is excluded above) and the accounts.json
# system data (SYSTEM_PREFIX), so none of the existing photo-library
# listings/usage/gallery code picks these up by accident. Unlike photos
# (admin-curated, uploaded deliberately through Admin Tools) every install
# writes here automatically right after generating a document — see
# upload_document()'s own docstring for why that needs allow_bundled_write.
DOCUMENTS_PREFIX = "documents/"


def upload_document(local_path, key_suffix):
    """
    Best-effort cloud backup of ONE generated document (called from
    app.py's /api/generate on a background thread right after the local
    save succeeds — see that call site's own comment for why it's
    fire-and-forget). key_suffix is the caller-built logical path under
    DOCUMENTS_PREFIX, e.g. "SOLOLUCE/CAT/AURA-ECO-Rev0.pdf" — a clean,
    brand/doctype/filename key, NOT the user's real local folder path
    (which is an arbitrary, potentially identifying Windows path chosen
    in Settings and different on every install).

    Uses allow_bundled_write=True: this is the one write path every
    install needs, not just the admin's, so it deliberately falls back to
    the bundled r2_readonly.json key when there's no local admin config —
    see _cfg_block()'s own comment for why that's a conscious choice, not
    an accidental hole (the bundled key already holds the same value as
    the admin's write key). Never raises — the caller doesn't want a
    document's local save to look like it failed just because this
    machine is offline or R2 is briefly down; it just tries again next
    time that same document is generated.
    """
    try:
        client = _client(require_write=True, allow_bundled_write=True)
        bucket = _bucket(require_write=True, allow_bundled_write=True)
    except Exception:
        return False
    key = DOCUMENTS_PREFIX + key_suffix.replace(os.sep, "/")
    content_type = mimetypes.guess_type(key)[0] or "application/octet-stream"
    try:
        with open(local_path, "rb") as f:
            client.put_object(Bucket=bucket, Key=key, Body=f, ContentType=content_type)
        return True
    except Exception:
        return False


def sync_down_documents(local_base_folder, on_progress=None):
    """
    Pulls every object under DOCUMENTS_PREFIX into local_base_folder,
    stripping the "documents/" prefix itself so a key like
    "documents/SOLOLUCE/CAT/AURA-ECO-Rev0.pdf" lands at
    local_base_folder/SOLOLUCE/CAT/AURA-ECO-Rev0.pdf — the exact layout
    app.py's folder_for()/all_doc_folders() now auto-generate paths
    into (see their own comments), so every existing local-file-based
    system (All Docs' scan_all(), Submissions' submittal PDF merge,
    previous_for_company()'s continuation lookups, sidecar read/write)
    keeps working completely unchanged: this just keeps that local
    mirror caught up with whatever every OTHER install has generated,
    the same role sync_down() already plays for the product photo
    library, and the same skip-if-already-present-with-matching-size
    freshness check.

    Uses the SAME read access every install already has for documents
    (no allow_bundled_write needed — this only reads) — every logged-in
    user already gets bundled read access via the readonly-fallback
    credential the same way they already read the shared photo library.
    Never raises for "not configured"/offline — callers (the startup
    sync and the periodic background loop) should just skip quietly.
    """
    if not is_configured():
        return {"ok": False, "error": "not_configured"}
    os.makedirs(local_base_folder, exist_ok=True)
    try:
        client = _client()
        bucket = _bucket()
        downloaded = 0
        token = None
        while True:
            kwargs = {"Bucket": bucket, "Prefix": DOCUMENTS_PREFIX}
            if token:
                kwargs["ContinuationToken"] = token
            resp = client.list_objects_v2(**kwargs)
            for obj in resp.get("Contents", []):
                key = obj["Key"]
                rest = key[len(DOCUMENTS_PREFIX):]
                if not rest:
                    continue
                dest = os.path.join(local_base_folder, *rest.split("/"))
                if os.path.isfile(dest) and os.path.getsize(dest) == obj["Size"]:
                    continue
                os.makedirs(os.path.dirname(dest) or local_base_folder, exist_ok=True)
                client.download_file(bucket, key, dest)
                downloaded += 1
                if on_progress:
                    on_progress(downloaded)
            if not resp.get("IsTruncated"):
                break
            token = resp.get("NextContinuationToken")
        return {"ok": True, "downloaded": downloaded}
    except Exception as e:
        return {"ok": False, "error": str(e)}


def sync_down(local_folder, on_progress=None):
    """
    Pulls every photo from R2 into local_folder, skipping files that are
    already present with a matching size (cheap, good-enough freshness
    check — product photos are uploaded once and rarely re-uploaded under
    the same filename with different content). Read-only credentials are
    enough for this — see module docstring on why non-admin installs should
    use a separate read-only API token rather than the admin's read/write one.
    Never raises for "not configured" — callers (the on-launch background
    sync) should just skip quietly, same as update_checker's own pattern.
    """
    if not is_configured():
        return {"ok": False, "error": "not_configured"}
    os.makedirs(local_folder, exist_ok=True)
    try:
        client = _client()
        bucket = _bucket()
        downloaded = 0
        token = None
        while True:
            kwargs = {"Bucket": bucket}
            if token:
                kwargs["ContinuationToken"] = token
            resp = client.list_objects_v2(**kwargs)
            for obj in resp.get("Contents", []):
                key = obj["Key"]
                if key.startswith(SYSTEM_PREFIX):
                    continue  # app data (accounts.json etc.), not a photo — never syncs into the local photo folder
                dest = os.path.join(local_folder, *key.split("/"))  # R2 keys are always "/"-separated regardless of OS
                if os.path.isfile(dest) and os.path.getsize(dest) == obj["Size"]:
                    continue  # already have this exact file
                os.makedirs(os.path.dirname(dest) or local_folder, exist_ok=True)
                client.download_file(bucket, key, dest)
                downloaded += 1
                if on_progress:
                    on_progress(downloaded)
            if not resp.get("IsTruncated"):
                break
            token = resp.get("NextContinuationToken")
        return {"ok": True, "downloaded": downloaded}
    except Exception as e:
        return {"ok": False, "error": str(e)}
