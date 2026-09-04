"""
HTML/PDF document engine
------------------------
Renders the new pixel-fidelity document designs (ported from the
"Company Templates Rebuild" design handoff) via a real headless-Chromium
render (Playwright), instead of the openpyxl/LibreOffice xlsx pipeline used
by QTN/INV/DO. See engine.HTML_DOC_TYPES for which doc types use this path.

Fonts (Archivo, IBM Plex Sans) are loaded from Google Fonts at render time —
this machine has normal internet access, matching how the original design
pack previews. If offline reliability ever matters, self-host the two font
files and swap the <link> tags in templates_html/*.html for a local
@font-face instead.
"""
import os, re, math, base64, datetime, json, threading, atexit

# Must happen before `import playwright...` (below) ever runs, and before
# any PyInstaller-frozen build's own bundled playwright driver gets a
# chance to resolve its default browser path. A frozen build's bundled
# driver package looks for browsers in ITS OWN ".local-browsers"
# subfolder next to itself — empty, since freezing never re-downloads
# Chromium — instead of the real per-user cache
# (%LOCALAPPDATA%\ms-playwright) `playwright install` actually populated
# on this machine (confirmed directly: an unset PLAYWRIGHT_BROWSERS_PATH
# in a frozen .exe raised "Executable doesn't exist at ...\_internal\
# playwright\driver\package\.local-browsers\..."). Pinning this env var
# forces Playwright to always use that one real, shared cache regardless
# of how/where the calling code was packaged — harmless and correct in
# normal `python app.py` use too, where it's simply already the default.
os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", os.path.expandvars(r"%LOCALAPPDATA%\ms-playwright"))

from collections import Counter
from jinja2 import Environment, FileSystemLoader
from markupsafe import Markup, escape
import engine
import runtime_manager

# Reuses engine.BASE (not its own recomputation) so both modules agree on
# one resource root — see engine.py's own comment on why this can't just
# be os.path.dirname(__file__) once the app is a frozen .exe.
BASE = engine.BASE
TEMPLATES_HTML = os.path.join(BASE, "templates_html")
DOC_HTML_STATIC = os.path.join(BASE, "static", "doc_html")
FONTS_DIR = os.path.join(BASE, "static", "fonts")

_jinja_env = Environment(loader=FileSystemLoader(TEMPLATES_HTML))

_font_cache = {}
def _font_data_uri(filename):
    """Self-hosted webfont (see static/fonts/) inlined as a data URI, same
    pattern as the logo/doc-page.js — cached in memory since these never
    change during the process's lifetime. Was previously a Google Fonts
    <link> fetched at render time; that per-render network round-trip (even
    when cached by Chromium) was the single biggest cost in a render, often
    ~500ms — inlining removes it entirely, on top of not depending on the
    machine having internet access."""
    if filename not in _font_cache:
        with open(os.path.join(FONTS_DIR, filename), "rb") as f:
            b64 = base64.b64encode(f.read()).decode("ascii")
        _font_cache[filename] = f"data:font/woff2;base64,{b64}"
    return _font_cache[filename]

# Company masthead info — constant across every document for a brand.
# Only Artemis is filled in for now (this design pack is Artemis-specific);
# other brands fall back to Artemis's info until their own is supplied.
COMPANY_INFO = {
    "ARTEMIS": {
        "name": "Artemis Electricals Est",
        "poBox": "21547",
        "city": "Dubai, UAE",
        "mobile": "+971 50 6431352",
        "tel": "+9714 452 8239",
        "trn": "100294715600003",
    },
}

def company_info(brand):
    brand = (brand or "ARTEMIS").upper()
    return COMPANY_INFO.get(brand, COMPANY_INFO["ARTEMIS"])

# Bank details block on the Invoice — ported verbatim from the bundled
# INV.xlsx templates (cells C23/C24), which use this exact same static
# text for every brand (ARTEMIS/SOLOLUCE/ADS/WATT all carry byte-identical
# C23/C24 content — confirmed by reading each template directly), same
# "only Artemis is really configured, everyone else inherits it" state as
# COMPANY_INFO above.
BANK_INFO = {
    "ARTEMIS": {
        "heading": "ARTEMIS ELECTRICAL EST BANK INFO:",
        "account_name": "Artemis Electrical Est",
        "account_no": "3708438234701",
        "iban": "AE740340003708438234701",
        "bank_name": "Emirates Islamic Bank",
        "branch": "El Al TWAR",
        "swift_code": "MEBLAEAD",
    },
}

def bank_info(brand):
    brand = (brand or "ARTEMIS").upper()
    return BANK_INFO.get(brand, BANK_INFO["ARTEMIS"])

# ----------------------------------------------------------------------------
# Shared money math for the Discount & VAT card (percent / fixed / target-price
# for discount; percent / fixed for VAT) — same rules as the xlsx QTN/INV
# path (engine._write_summary_block), reimplemented here in plain Python
# since there's no spreadsheet formula engine in this pipeline.
# ----------------------------------------------------------------------------
def _f(v, default=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default

def compute_totals(items, discount=None, vat=None):
    discount = discount or {}
    vat = vat or {}
    subtotal = sum(_f(it.get("qty")) * _f(it.get("price")) for it in items)

    discount_enabled = bool(discount.get("enabled"))
    discount_amount = 0.0
    if discount_enabled:
        mode = discount.get("mode", "percent")
        value = _f(discount.get("value"))
        if mode == "percent":
            discount_amount = subtotal * value / 100
        elif mode == "fixed":
            discount_amount = value
        elif mode == "target":
            discount_amount = max(subtotal - value, 0.0)
    price_after_discount = subtotal - discount_amount
    base_for_vat = price_after_discount if discount_enabled else subtotal

    vat_enabled = bool(vat.get("enabled"))
    vat_amount = 0.0
    vat_pct = None
    if vat_enabled:
        vmode = vat.get("mode", "percent")
        vvalue = _f(vat.get("value"))
        if vmode == "percent":
            vat_amount = base_for_vat * vvalue / 100
            vat_pct = vvalue
        else:
            vat_amount = vvalue

    total = base_for_vat + vat_amount
    return {
        "subtotal": subtotal,
        "discount_enabled": discount_enabled,
        "discount_amount": discount_amount,
        "price_after_discount": price_after_discount,
        "vat_enabled": vat_enabled,
        "vat_amount": vat_amount,
        "vat_pct": vat_pct,
        "total": total,
    }

def money(v):
    return f"{v:,.2f}"

def _fmt_date(date_iso):
    """DD.MM.YYYY, matching the xlsx pipeline's engine._fmt_date — falls
    back to the raw string if it isn't a valid ISO date (e.g. blank draft)."""
    try:
        return engine._fmt_date(date_iso)
    except (ValueError, TypeError):
        return date_iso or ""

def num_display(v):
    """Whole numbers show as '20', not '20.0' (qty/etc — money() always
    wants 2dp so it's kept separate)."""
    f = _f(v)
    return str(int(f)) if f == int(f) else str(f)

# ----------------------------------------------------------------------------
# Amount in words (AED) — no existing converter in the codebase (the xlsx
# template leaves "Amount in words" blank for manual entry), so this is a
# small self-contained one, not a port of anything.
# ----------------------------------------------------------------------------
_ONES = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
         "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
         "Seventeen", "Eighteen", "Nineteen"]
_TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]

def _three_digit_words(n):
    words = []
    if n >= 100:
        words.append(_ONES[n // 100] + " Hundred")
        n %= 100
    if n >= 20:
        tens_word = _TENS[n // 10]
        if n % 10:
            tens_word += "-" + _ONES[n % 10].lower()
        words.append(tens_word)
    elif n > 0:
        words.append(_ONES[n])
    return " ".join(words)

def _int_to_words(n):
    if n == 0:
        return "Zero"
    scales = [(1_000_000_000, "Billion"), (1_000_000, "Million"), (1_000, "Thousand"), (1, "")]
    parts = []
    for scale, name in scales:
        if n >= scale:
            count = n // scale
            n %= scale
            chunk = _three_digit_words(count)
            parts.append(f"{chunk} {name}".strip() if name else chunk)
    return " ".join(parts)

def number_to_words(amount, currency="AED", subunit="Fils"):
    amount = round(_f(amount), 2)
    whole = int(amount)
    fils = round((amount - whole) * 100)
    words = _int_to_words(whole) + f" {currency}"
    if fils:
        words += f" and {_int_to_words(fils)} {subunit}"
    return words + " Only"

def _logo_data_uri():
    path = os.path.join(DOC_HTML_STATIC, "artemis-logo.png")
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    return f"data:image/png;base64,{b64}"

# Brand logo for the Sololuce datasheet header — reuses the same PNG the brand
# switcher already uses (static/logos/<brand>.png), unlike the QTN2 pipeline's
# own baked-in Artemis-only logo above.
STATIC_LOGOS = os.path.join(BASE, "static", "logos")
BRAND_WEBSITE = {"SOLOLUCE": "www.sololucelightings.com"}

def _brand_logo_data_uri(brand):
    path = os.path.join(STATIC_LOGOS, f"{(brand or 'sololuce').lower()}.png")
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    return f"data:image/png;base64,{b64}"

# ----------------------------------------------------------------------------
# Sololuce Product Datasheet (CAT) — spec-badge icon library.
# Every badge is a real image, not a hand-drawn approximation: the user
# supplied the actual PNG set Sololuce uses (39 of them, one per icon), copied
# read-only from F:\...\4. SOLOLUCE\1. CATALOGUE\Edgar\4. Back Pages\ICONS
# into static/cat_badges/ (see app.py's cat_badge_library config default for
# the key/label/filename seed list). Each PNG is fully self-contained — icon,
# caption text, and border are already baked in by the original designer
# (e.g. "IP54" or "DALI" is printed inside the image itself) — so there is no
# per-product value to overlay at render time; picking a badge just means
# picking the one specific pre-made image that already says the right thing
# (e.g. a product rated IP65 uses a different library entry than one rated
# IP44, rather than one generic "IP" badge with a number typed in after).
# The library lives in config.json, not hardcoded here, specifically so the
# user's own "+ Add Custom Badge" uploads (any new icon they add, with
# whatever name/details they like) show up next to the original 39 with zero
# code changes — see app.py's /api/cat-badges-add.
# ----------------------------------------------------------------------------
CAT_BADGES_DIR = os.path.join(BASE, "static", "cat_badges")

def _load_badge_library():
    try:
        with open(os.path.join(BASE, "config.json"), encoding="utf-8") as f:
            cfg = json.load(f)
        return cfg.get("cat_badge_library", [])
    except (OSError, ValueError):
        return []

def _badge_image_data_uri(filename):
    path = os.path.join(CAT_BADGES_DIR, filename)
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    ext = os.path.splitext(filename)[1].lstrip(".").lower() or "png"
    return f"data:image/{ext};base64,{b64}"

def badges_for(selected):
    """selected: [{key}] as posted from the Build form -> [{key, label, src}]
    for the template, resolving each key against the persisted badge library
    (config.json) so a user-added custom badge works identically to one of
    the original 39. Silently skips a key that no longer resolves (library
    entry deleted / image file missing) rather than breaking the render."""
    library = {b["key"]: b for b in _load_badge_library() if b.get("key")}
    out = []
    for item in (selected or []):
        key = (item.get("key") or "").strip()
        b = library.get(key)
        if not b:
            continue
        try:
            src = _badge_image_data_uri(b["filename"])
        except OSError:
            continue
        out.append({"key": key, "label": b.get("label", key), "src": src})
    return out

def _doc_page_js_data_uri():
    path = os.path.join(DOC_HTML_STATIC, "doc-page.js")
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    return f"data:text/javascript;base64,{b64}"

_tls = threading.local()

def _get_browser():
    """One persistent headless Chromium, launched lazily and kept alive for
    the process's lifetime, instead of a fresh launch (~1s) per render.
    Playwright's sync API is pinned to whichever thread starts it, so this
    is cached thread-local — safe and effective ONLY because every caller
    now goes through the dedicated worker thread below (_worker_loop),
    which is the single, real, long-lived thread that ever calls this.
    That distinction mattered in practice: confirmed directly (live
    timing + a thread-name/id print) that Flask's dev/request threads are
    NOT stable across requests — each /api/preview-draft call landed on a
    genuinely different Thread-N, so a naive threading.local() cache here
    was silently useless (a "cache" that's never hit twice is just
    overhead with extra steps) and every render was paying full
    browser.new_page() cost (~600-750ms) regardless. Funneling all
    rendering through one dedicated thread is what makes this cache
    (and _get_page()'s) actually pay off."""
    browser = getattr(_tls, "browser", None)
    if browser is not None:
        try:
            if browser.is_connected():
                return browser
        except Exception:
            pass
    from playwright.sync_api import sync_playwright
    _tls.playwright = sync_playwright().start()
    # The packaged app used to depend on a SEPARATELY-installed Chromium
    # (`playwright install chromium`, living outside the app entirely in
    # %LOCALAPPDATA%\ms-playwright) — confirmed directly, across many
    # live reproduction attempts against the real installed app (including
    # a real user's own normal launch, not just automated testing), to be
    # fundamentally unreliable: Playwright's own default resolution picks
    # the wrong browser variant (chrome-headless-shell, which has no
    # PDF/print support at all) in a frozen build, AND — more
    # fundamentally — that external folder can become entirely invisible
    # to this app in some launch contexts even though it demonstrably
    # exists (confirmed independently from outside the process, and
    # confirmed by BOTH Python's own checks AND Playwright's separate
    # Node-side check agreeing it "doesn't exist"). Whatever the real
    # cause, depending on anything outside the app's own installed files
    # is the actual problem.
    #
    # Fix: bundle the real chrome.exe INTO the app itself, read through a
    # known local path the same proven-reliable way templates_html/static
    # already are, rather than depending on anything outside the app's
    # own installed files. Falls back to Playwright's own default
    # resolution only for a dev checkout that's never populated a local
    # runtime at all.
    #
    # Checked in order:
    # 1. runtime_manager.BUNDLED_BROWSER_DIR — engine.DATA_BASE/runtime/
    #    bundled_browser, OUTSIDE {app}\_internal\, so ordinary app
    #    updates never touch it (see runtime_manager.py's own module
    #    docstring for the full "why split this out" reasoning: the
    #    ~400MB Chromium payload almost never changes between app
    #    releases, so shipping it in the SAME installer as the app's own
    #    code — which changes almost every release — meant re-downloading
    #    and re-installing the whole browser on every single app update
    #    for nothing). app.py's startup calls
    #    runtime_manager.ensure_runtime() before this is ever reached, so
    #    in the normal case this is already populated by the time any
    #    render happens.
    # 2. BASE/bundled_browser — the OLD bundled-at-PyInstaller-build-time
    #    location. Kept as a fallback (not removed) for a dev checkout
    #    that populated it the old way and hasn't set up RUNTIME_VERSION/
    #    the separate runtime download yet.
    candidate = None
    for bundled_root in (
        runtime_manager.BUNDLED_BROWSER_DIR,
        os.path.join(BASE, "bundled_browser"),
    ):
        if not os.path.isdir(bundled_root):
            continue
        for name in os.listdir(bundled_root):
            p = os.path.join(bundled_root, name, "chrome-win64", "chrome.exe")
            if os.path.isfile(p):
                candidate = p
                break
        if candidate:
            break
    if candidate:
        _tls.browser = _tls.playwright.chromium.launch(executable_path=candidate)
    else:
        try:
            _tls.browser = _tls.playwright.chromium.launch()
        except Exception as e:
            msg = str(e)
            if "Executable doesn't exist" not in msg:
                raise
            browsers_dir = os.environ.get("PLAYWRIGHT_BROWSERS_PATH") or os.path.expandvars(r"%LOCALAPPDATA%\ms-playwright")
            m = re.search(r"chromium_headless_shell-(\d+)", msg)
            if not m:
                raise
            fallback = os.path.join(browsers_dir, "chromium-" + m.group(1), "chrome-win64", "chrome.exe")
            _tls.browser = _tls.playwright.chromium.launch(executable_path=fallback)
    atexit.register(_close_thread_browser, _tls)
    return _tls.browser

def _get_page():
    """One persistent tab, reused across every render on this thread,
    instead of browser.new_page() + page.close() around each one —
    measured directly (live timing added while chasing a real "the
    preview feels slow" report): new_page() alone was costing
    ~600-750ms PER RENDER, not the "~10s of ms" this module used to
    (wrongly, never actually measured) claim — by far the single
    biggest cost in the whole pipeline, ahead of the actual PDF export
    itself (~270ms). set_content() fully replaces a page's DOM/state on
    every call, so reusing one tab is safe — there's nothing left over
    from the previous render for a fresh set_content() not to already
    overwrite."""
    browser = _get_browser()
    page = getattr(_tls, "page", None)
    if page is not None:
        try:
            if not page.is_closed():
                return page
        except Exception:
            pass
    _tls.page = browser.new_page()
    return _tls.page

def _close_thread_browser(tls):
    try:
        tls.browser.close()
    except Exception:
        pass
    try:
        tls.playwright.stop()
    except Exception:
        pass

import queue as _queue

# Every actual Playwright call (browser/page creation, set_content, pdf())
# happens on exactly ONE dedicated background thread, started lazily on
# first use and kept alive for the process's lifetime — never on whichever
# Flask request thread happened to receive the HTTP call. This is what
# makes _get_browser()/_get_page()'s thread-local caching above actually
# work: confirmed directly (a thread-name print, live) that Flask's own
# request threads are a fresh Thread-N every single call, so caching
# Playwright objects against THAT thread was never going to hit twice —
# every render was silently paying full browser.new_page() cost
# regardless of the cache. Request threads never touch Playwright
# directly; they submit a job here and block on its own result queue.
_render_job_queue = _queue.Queue()
_render_worker_lock = threading.Lock()
_render_worker_started = False

def _render_worker_loop():
    while True:
        html, out_path, result_q = _render_job_queue.get()
        try:
            _render_html_to_pdf(html, out_path)
            result_q.put((True, out_path))
        except Exception as e:
            result_q.put((False, e))

def _ensure_render_worker():
    global _render_worker_started
    if _render_worker_started:
        return
    with _render_worker_lock:
        if not _render_worker_started:
            threading.Thread(target=_render_worker_loop, daemon=True, name="html-render-worker").start()
            _render_worker_started = True

def render_pdf(template_name, context, out_path):
    """Render a templates_html/<template_name> Jinja2 template with context,
    then print it to PDF via headless Chromium. Logo, doc-page.js, and every
    webfont used by any HTML doc template are all inlined as data URIs, so
    the render has zero filesystem/network dependency — set on every call
    regardless of which fonts the specific template actually references
    (an unused @font-face costs nothing; browsers only fetch a font when
    something on the page actually uses it, and here there's no fetch at
    all since it's already inlined).

    The Jinja2 render itself happens right here, on the CALLING (Flask
    request) thread — cheap, thread-safe, no reason to hop threads for it.
    Only the actual browser work is handed off to the dedicated worker
    thread (see _ensure_render_worker()/_render_worker_loop() above);
    this call blocks until that finishes, so callers see it as a normal
    synchronous function either way."""
    context = dict(context)
    context.setdefault("logo_src", _logo_data_uri())
    context.setdefault("doc_page_js", _doc_page_js_data_uri())
    context.setdefault("archivo_src", _font_data_uri("archivo.woff2"))
    context.setdefault("ibmplexsans_src", _font_data_uri("ibm-plex-sans.woff2"))
    context.setdefault("bodonimoda_src", _font_data_uri("bodoni-moda.woff2"))
    context.setdefault("onest_src", _font_data_uri("onest.woff2"))
    context.setdefault("neuropol_src", _font_data_uri("neuropol.woff2"))
    context.setdefault("gothamblack_src", _font_data_uri("gotham-black.woff2"))
    context.setdefault("gothambook_src", _font_data_uri("gotham-book.woff2"))
    context.setdefault("gothambold_src", _font_data_uri("gotham-bold.woff2"))
    html = _jinja_env.get_template(template_name).render(**context)
    _ensure_render_worker()
    result_q = _queue.Queue()
    _render_job_queue.put((html, out_path, result_q))
    ok, payload = result_q.get()
    if not ok:
        raise payload
    return payload

def _render_html_to_pdf(html, out_path):
    """The actual Chromium work — ONLY ever called from the dedicated
    worker thread (_render_worker_loop above), never directly. See
    render_pdf()'s own docstring for why this split exists."""
    page = _get_page()
    try:
        # emulate_media("print") — page.pdf() itself always switches to print
        # CSS internally right before it rasterizes, but everything else
        # (our own page.evaluate() calls, including the photo-pan-fix
        # <script> that's part of the HTML itself) otherwise runs under
        # Chromium's default SCREEN media, which doc-page.js's own injected
        # `@page`/`@media print` rules (see that file) can make genuinely
        # differently-shaped from the real printed page — confirmed
        # directly: a photo-pan getBoundingClientRect() measured under
        # screen media gave a box ~1.77x narrower than the same element
        # under print media (158px vs 280px, same page/photo), enough to
        # visibly mis-place a photo that was pan-adjusted right at its
        # letterbox-gap extreme. Set BEFORE set_content, not after —
        # set_content() itself runs the page's own inline <script>s as
        # part of loading, so the photo-pan-fix script (which measures and
        # self-resolves synchronously against decoded photos) would already
        # be resolved under stale screen-media geometry by the time a
        # later emulate_media() call took effect, same class of bug as
        # calling it too late ever would be. Same underlying screen-vs-
        # print divergence already called out below for `ch`-based grid
        # metrics, just hitting a JS measurement instead of a pure-CSS one.
        page.emulate_media(media="print")
        page.set_content(html, wait_until="load")
        # Viewport-vs-print-page-width mismatch — a second, distinct
        # divergence from the screen-vs-print MEDIA one already handled
        # above (that one's about which CSS rules apply; this one's about
        # available LAYOUT WIDTH even once print rules already apply).
        # Playwright's page never had its viewport resized to match the
        # doc's own @page size doc-page.js injects — it stays whatever the
        # browser context's default viewport is (materially wider than an
        # A4/Letter page's real printed width). Chromium's page.pdf() does
        # its own dedicated print-layout pass at the TRUE paper width
        # regardless of that viewport, so the final PDF itself was never
        # actually wrong — but any window.__panReady-style <script> that
        # calls getBoundingClientRect() (see sololuce_datasheet.html's own
        # script) is measuring the WRONG-width screen-viewport layout, not
        # the one page.pdf() is about to produce, and can silently
        # disagree with the real output — confirmed directly, chasing an
        # RSPT/Application-Photo alignment bug: page.evaluate() reported
        # 0.0px misalignment while the real generated PDF's own native
        # vector coordinates showed an 11.25pt gap; re-running the exact
        # same template/data with the viewport pre-sized to the doc's true
        # printed pixel dimensions made page.evaluate() and the real PDF
        # agree, and independently confirmed (via the same vector-
        # coordinate check on content this resize never touches — Main
        # Product Photo's own box) that the resize changes nothing about
        # what page.pdf() actually renders; it only makes the measurement
        # trustworthy. Scoped behind the same window.__panReady existence
        # check as the wait below, for the same reason: most templates
        # have no post-layout JS measurement to correct in the first
        # place, so paying for a second full set_content() reload (needed
        # because a mid-flight viewport resize doesn't retroactively fix
        # the synchronous measurements __panReady's own script already
        # took while parsing at the old width) is only ever worth it for
        # the ones that do.
        needs_repaint = page.evaluate("typeof window.__panReady !== 'undefined'")
        if needs_repaint:
            page_px = page.evaluate("""() => {
                var dp = document.querySelector('doc-page');
                if (!dp) return null;
                function toPx(len){
                    var d = document.createElement('div');
                    d.style.cssText = 'position:absolute;visibility:hidden;width:' + len + ';height:' + len + ';';
                    document.body.appendChild(d);
                    var px = d.getBoundingClientRect().width;
                    document.body.removeChild(d);
                    return px;
                }
                return {w: Math.ceil(toPx(dp.pageWidth)), h: Math.ceil(toPx(dp.pageHeight))};
            }""")
            if page_px and page_px.get("w") and page_px.get("h"):
                page.set_viewport_size({"width": page_px["w"], "height": page_px["h"]})
                page.set_content(html, wait_until="load")
        # Was wait_until="networkidle" on both set_content() calls above —
        # measured directly (chasing a real "the preview feels slow"
        # report) at costing ~500ms per call for zero actual benefit: every
        # resource this app's templates use (logo, doc-page.js, every
        # @font-face, even user-picked photos — see _font_data_uri's own
        # docstring and render_pdf's own docstring) is inlined as a data:
        # URI specifically so there's no real network fetch to ever wait
        # for, which is exactly why networkidle's mandatory "500ms of no
        # network activity" quiet-window was pure dead time here rather
        # than protecting against anything. "load" only waits for the
        # page's own load event — real font-shaping/layout completion
        # still gets its own explicit, correct wait right below
        # (document.fonts.ready), which is the actual signal that matters
        # for `ch`-based grid metrics (see that comment, unchanged) — this
        # was never networkidle's job to begin with, confirmed directly:
        # screen-mode layout (after fonts had visibly settled) and print-
        # mode layout of the exact same HTML disagreed on column widths
        # widely enough to push a table's last two columns fully off the
        # page in print while screen showed all of them comfortably
        # fitting, well before this switch away from networkidle existed.
        page.evaluate("document.fonts.ready")
        # window.__panReady: sololuce_datasheet.html's own photo-pan-fix
        # <script> (see its comment there) — re-measures every pannable
        # photo's real rendered box + natural size and overwrites its
        # transform once loaded, since the Jinja-computed inline transform
        # alone can't reach an aspect-ratio-mismatch letterbox gap, only
        # zoom-created overflow. Waiting on it here is exactly why it's a
        # Promise in the first place — Playwright's page.evaluate() awaits
        # a returned Promise before continuing, same pattern as
        # document.fonts.ready above. `|| Promise.resolve()` because this
        # function is shared by every doc type's render — most templates
        # never define window.__panReady at all, and a bare `undefined`
        # would resolve immediately anyway, but the explicit fallback
        # keeps this line correct even if a future template defines a
        # non-Promise value under the same name by mistake.
        page.evaluate("window.__panReady || Promise.resolve()")
        # prefer_css_page_size makes Playwright actually honor the @page
        # size doc-page.js injects (see that file's own module docstring)
        # instead of silently defaulting to US Letter — every template here
        # declares <doc-page size="a4">, so without this every "a4" document
        # this app generates was really coming out as Letter (612x792pt)
        # regardless of what the template said.
        page.pdf(path=out_path, print_background=True, prefer_css_page_size=True)
    except Exception:
        # Discard the page rather than reuse it after a failed render —
        # _get_page() normally keeps one tab alive across every call (see
        # its own docstring for why), but a page that broke mid-render is
        # a bad bet to hand back out on the next call. Closing it here
        # just means the next _get_page() call opens a fresh one instead
        # of reusing a possibly-wedged tab; the happy path below never
        # pays this cost.
        try:
            page.close()
        except Exception:
            pass
        _tls.page = None
        raise
    return out_path

def render_quotation_pdf(data, out_path, brand=None):
    """data keys: number, rev, date, project, area, company(customer name),
    customer_attn, customer_address, status, items[{type,description,unit,
    qty,price,photo}], discount{enabled,mode,value}, vat{enabled,mode,value},
    terms{delivery,payment,warranty}"""
    items_in = data.get("items", []) or []
    totals = compute_totals(items_in, data.get("discount"), data.get("vat"))
    vat_pct = totals["vat_pct"]

    rows = []
    for i, it in enumerate(items_in):
        qty = _f(it.get("qty"))
        price = _f(it.get("price"))
        amount = qty * price
        row_vat = amount * vat_pct / 100 if (totals["vat_enabled"] and vat_pct is not None) else None
        rows.append({
            "no": str(i + 1).zfill(2),
            "type": it.get("type", ""),
            "desc": it.get("description", ""),
            "unit": it.get("unit", "PCS"),
            "qty": num_display(qty),
            "price": money(price),
            "vat": money(row_vat) if row_vat is not None else "",
            "amount": money(amount),
            "photoUrl": it.get("photo") or "",
            "rowStyle": "background:#f8fafb;" if i % 2 else "background:#ffffff;",
        })

    status = (data.get("status") or "Draft").strip()
    project = data.get("project", "")
    area = data.get("area", "")
    project_area = " · ".join(p for p in (project, area) if p)
    terms = data.get("terms") or {}

    context = {
        "company": company_info(brand),
        "customer": {
            "name": data.get("company", ""),
            "attn": data.get("customer_attn", ""),
            "address": data.get("customer_address", ""),
        },
        "qtn_number": data.get("number", ""),
        "rev": data.get("rev", "0"),
        "date": _fmt_date(data.get("date", "")),
        "project_area": project_area,
        "show_status": status.lower() != "none",
        "status": status,
        "items": rows,
        "vat_col_visible": totals["vat_enabled"] and vat_pct is not None,
        "vat_col_label": f"VAT {vat_pct:g}%" if vat_pct is not None else "VAT",
        "subtotal": money(totals["subtotal"]),
        "discount_visible": totals["discount_enabled"],
        "discount_amount": money(totals["discount_amount"]),
        "price_after_discount": money(totals["price_after_discount"]),
        "vat_visible": totals["vat_enabled"],
        "vat_label": f"VAT {vat_pct:g}%" if vat_pct is not None else "VAT",
        "vat_amount": money(totals["vat_amount"]),
        "total": f"AED {money(totals['total'])}",
        "amount_words": number_to_words(totals["total"]),
        "terms": {
            "delivery": terms.get("delivery") or "10-12 weeks",
            "payment": terms.get("payment") or "50% advance, 50% upon delivery",
            "warranty": terms.get("warranty") or "5 years warranty",
        },
    }
    return render_pdf("quotation.html", context, out_path)

def render_invoice_pdf(data, out_path, brand=None):
    """data keys: number, date, qtn_number, lpo_number, project, type,
    company(customer name), customer_block(plain text, \\n-separated —
    the same composed block app.py's customerBlockForXlsx()/the xlsx
    pipeline's _write_customer_block already use), items[{description,
    unit,qty,price,photo}], discount{enabled,mode,value},
    vat{enabled,mode,value}. Row/totals math mirrors render_quotation_pdf
    exactly (invoice.html visually matches quotation.html — see that
    template's own comments for the design tokens both share)."""
    items_in = data.get("items", []) or []
    totals = compute_totals(items_in, data.get("discount"), data.get("vat"))
    vat_pct = totals["vat_pct"]

    rows = []
    for i, it in enumerate(items_in):
        qty = _f(it.get("qty"))
        price = _f(it.get("price"))
        amount = qty * price
        row_vat = amount * vat_pct / 100 if (totals["vat_enabled"] and vat_pct is not None) else None
        rows.append({
            "no": str(i + 1).zfill(2),
            "desc": it.get("description", ""),
            "unit": it.get("unit", "PCS"),
            "qty": num_display(qty),
            "price": money(price),
            "vat": money(row_vat) if row_vat is not None else "",
            "amount": money(amount),
            "photoUrl": it.get("photo") or "",
            "rowStyle": "background:#f8fafb;" if i % 2 else "background:#ffffff;",
        })

    context = {
        "company": company_info(brand),
        "bank": bank_info(brand),
        "customer_name": data.get("company", ""),
        "customer_lines": [ln for ln in (data.get("customer_block") or "").split("\n") if ln.strip()],
        "inv_number": data.get("number", ""),
        "qtn_number": data.get("qtn_number", ""),
        "lpo_number": data.get("lpo_number", ""),
        "project": data.get("project", ""),
        "type_": data.get("type", ""),
        "date": _fmt_date(data.get("date", "")),
        "items": rows,
        "vat_col_visible": totals["vat_enabled"] and vat_pct is not None,
        "vat_col_label": f"VAT {vat_pct:g}%" if vat_pct is not None else "VAT",
        "subtotal": money(totals["subtotal"]),
        "discount_visible": totals["discount_enabled"],
        "discount_amount": money(totals["discount_amount"]),
        "price_after_discount": money(totals["price_after_discount"]),
        "vat_visible": totals["vat_enabled"],
        "vat_label": f"VAT {vat_pct:g}%" if vat_pct is not None else "VAT",
        "vat_amount": money(totals["vat_amount"]),
        "total": f"AED {money(totals['total'])}",
        # Left blank exactly like the old xlsx template's "Amount in
        # words:" cell — manual/handwritten entry, never auto-computed
        # there — so this restyle doesn't change what the document says,
        # only how it looks.
        "amount_words": "",
    }
    return render_pdf("invoice.html", context, out_path)

def render_delivery_order_pdf(data, out_path, brand=None):
    """data keys: number, date, project, lpo_number, company(customer
    name), customer_block(plain text), items[{description,unit,lpo_qty,
    prev_delivery,delivered,photo}], optional status (Draft/Delivered/
    Partial/None — matches the "Company Templates Rebuild" design
    handoff's own enum, see delivery_order.html). No Unit column: the old
    xlsx filler wrote a unit-annotated quantity into the LPO Quantity
    cell and then immediately overwrote that same cell with the bare
    number right after (engine.fill_delivery_order's two back-to-back
    loops both target column F) — Unit never actually reached the real
    output, so this keeps matching what the document has actually always
    shown. Status defaults to hidden: the app has no DO status concept
    yet, so an absent `status` key changes nothing about what today's
    callers produce — this only activates once/if a caller starts
    passing one."""
    items_in = data.get("items", []) or []

    def _num_or_blank(v):
        return num_display(v) if v not in (None, "") else ""

    rows = []
    for i, it in enumerate(items_in):
        rows.append({
            "no": str(i + 1).zfill(2),
            "desc": it.get("description", ""),
            "lpo_qty": _num_or_blank(it.get("lpo_qty")),
            "prev_delivery": _num_or_blank(it.get("prev_delivery")),
            "delivered": _num_or_blank(it.get("delivered")),
            "photoUrl": it.get("photo") or "",
            "rowStyle": "background:#f8fafb;" if i % 2 else "background:#ffffff;",
        })

    status = (data.get("status") or "").strip()
    context = {
        "company": company_info(brand),
        "customer_name": data.get("company", ""),
        "customer_lines": [ln for ln in (data.get("customer_block") or "").split("\n") if ln.strip()],
        "do_number": data.get("number", ""),
        "lpo_number": data.get("lpo_number", ""),
        "project": data.get("project", ""),
        "date": _fmt_date(data.get("date", "")),
        "items": rows,
        "show_status": bool(status) and status.lower() != "none",
        "status": status,
    }
    return render_pdf("delivery_order.html", context, out_path)

def render_proforma_invoice_pdf(data, out_path, brand=None):
    """New doc type, ported straight from the "Company Templates Rebuild"
    design handoff's Proforma Invoice.dc.html (PI) — a UAE-trading-industry
    standard document generated from an approved Quotation, ahead of the
    real Delivery Order/Tax Invoice, mainly for advance-payment/export
    reference (see the "Generate ▾" menu on a Quotation's All Docs row).
    Unlike Quotation/Tax Invoice, the real design file has no discount row
    and no per-item `type` column — Subtotal/VAT 5%/Total only, VAT always
    shown (not conditionally hidden) — so this intentionally does NOT
    reuse render_quotation_pdf's row/discount shape.
    data keys: number, date, qtn_number, valid_until, company(customer
    name), customer_block(plain text), project, items[{description,unit,
    qty,price,photo}], vat{enabled,mode,value} (defaults to UAE's
    standard 5% if omitted), optional status (Draft/Sent/Accepted/None,
    defaults to "Sent" per the design)."""
    items_in = data.get("items", []) or []
    vat_cfg = data.get("vat") or {"enabled": True, "mode": "percent", "value": 5}
    totals = compute_totals(items_in, None, vat_cfg)
    vat_pct = totals["vat_pct"] if totals["vat_pct"] is not None else 5

    rows = []
    for i, it in enumerate(items_in):
        qty = _f(it.get("qty"))
        price = _f(it.get("price"))
        amount = qty * price
        row_vat = amount * vat_pct / 100
        rows.append({
            "no": str(i + 1).zfill(2),
            "desc": it.get("description", ""),
            "unit": it.get("unit", "PCS"),
            "qty": num_display(qty),
            "price": money(price),
            "vat": money(row_vat),
            "amount": money(amount),
            "photoUrl": it.get("photo") or "",
            "rowStyle": "background:#f8fafb;" if i % 2 else "background:#ffffff;",
        })

    status = (data.get("status") or "Sent").strip()
    context = {
        "company": company_info(brand),
        "bank": bank_info(brand),
        "customer_name": data.get("company", ""),
        "customer_lines": [ln for ln in (data.get("customer_block") or "").split("\n") if ln.strip()],
        "project": data.get("project", ""),
        "pi_number": data.get("number", ""),
        "qtn_number": data.get("qtn_number", ""),
        "date": _fmt_date(data.get("date", "")),
        "valid_until": _fmt_date(data.get("valid_until", "")) if data.get("valid_until") else "",
        "items": rows,
        "subtotal": money(totals["subtotal"]),
        "vat_total": money(totals["vat_amount"]),
        "total": f"AED {money(totals['total'])}",
        "amount_words": number_to_words(totals["total"]),
        "show_status": status.lower() != "none",
        "status": status,
    }
    return render_pdf("proforma_invoice.html", context, out_path)

def render_payment_receipt_pdf(data, out_path, brand=None):
    """New doc type "RV", ported from the design handoff's Payment
    Receipt.dc.html — generated from a Tax Invoice's own data (see the
    "Generate ▾" menu on an Invoice's All Docs row), for the payment
    actually received against it. The real design has no line-item table
    at all — just one Amount Received panel and an "Applied To" allocation
    table — so the total is computed from the source invoice's items/vat
    (same compute_totals() math every other renderer uses) rather than
    trusting a precomputed string passed in.
    data keys: number, date, payment_method, reference, company(payer
    name), customer_block(plain text), project, items[{qty,price}] (used
    only to compute the total actually received), vat{enabled,mode,value},
    invoice_number, invoice_date, optional status (Received/Cleared/
    Pending/None, defaults to "Received"). The real design hardcodes the
    status pill green regardless of which of those four values is chosen
    (a receipt reads as inherently positive-toned) — ported faithfully,
    not invented."""
    items_in = data.get("items", []) or []
    totals = compute_totals(items_in, None, data.get("vat") or {"enabled": True, "mode": "percent", "value": 5})
    total_str = f"AED {money(totals['total'])}"

    allocations = [{
        "inv": data.get("invoice_number", ""),
        "date": _fmt_date(data.get("invoice_date", "")),
        "invAmount": money(totals["total"]),
        "applied": money(totals["total"]),
        "rowStyle": "background:#ffffff;",
    }] if data.get("invoice_number") else []

    status = (data.get("status") or "Received").strip()
    context = {
        "company": company_info(brand),
        "customer_name": data.get("company", ""),
        "customer_lines": [ln for ln in (data.get("customer_block") or "").split("\n") if ln.strip()],
        "project": data.get("project", ""),
        "receipt_number": data.get("number", ""),
        "date": _fmt_date(data.get("date", "")),
        "payment_method": data.get("payment_method", ""),
        "reference": data.get("reference", ""),
        "total": total_str,
        "amount_words": number_to_words(totals["total"]),
        "allocations": allocations,
        "outstanding_balance": "AED 0.00",
        "show_status": status.lower() != "none",
        "status": status,
        "status_color": "#1f7a54",
    }
    return render_pdf("payment_receipt.html", context, out_path)

def render_credit_note_pdf(data, out_path, brand=None):
    """New doc type "CN", ported from the design handoff's Credit
    Note.dc.html — generated from a Tax Invoice's own data (see the
    "Generate ▾" menu on an Invoice's All Docs row). Unlike DO/INV/PI,
    the credited items and the reason genuinely can't be inferred from
    the source invoice alone (a credit is a deliberate, partial
    correction) — the caller (see runRowGenerate's 'CN' case in app.py)
    collects the reason from the user first; items default to the full
    invoice unless the caller trims them.
    data keys: number, date, against_invoice, invoice_date, company
    (customer name), customer_block(plain text), project, items[{
    description,unit,qty,price,photo}], vat{enabled,mode,value}, reason
    (free text), optional status (Draft/Issued/Applied/None, defaults to
    "Issued")."""
    items_in = data.get("items", []) or []
    totals = compute_totals(items_in, None, data.get("vat") or {"enabled": True, "mode": "percent", "value": 5})
    vat_pct = totals["vat_pct"] if totals["vat_pct"] is not None else 5

    rows = []
    for i, it in enumerate(items_in):
        qty = _f(it.get("qty"))
        price = _f(it.get("price"))
        amount = qty * price
        row_vat = amount * vat_pct / 100
        rows.append({
            "no": str(i + 1).zfill(2),
            "desc": it.get("description", ""),
            "unit": it.get("unit", "PCS"),
            "qty": num_display(qty),
            "price": money(price),
            "vat": money(row_vat),
            "amount": money(amount),
            "photoUrl": it.get("photo") or "",
            "rowStyle": "background:#f8fafb;" if i % 2 else "background:#ffffff;",
        })

    status = (data.get("status") or "Issued").strip()
    context = {
        "company": company_info(brand),
        "customer_name": data.get("company", ""),
        "customer_lines": [ln for ln in (data.get("customer_block") or "").split("\n") if ln.strip()],
        "project": data.get("project", ""),
        "cn_number": data.get("number", ""),
        "date": _fmt_date(data.get("date", "")),
        "against_invoice": data.get("against_invoice", ""),
        "invoice_date": _fmt_date(data.get("invoice_date", "")),
        "items": rows,
        "subtotal": money(totals["subtotal"]),
        "vat_total": money(totals["vat_amount"]),
        "total": f"AED {money(totals['total'])}",
        "amount_words": number_to_words(totals["total"]),
        "reason": data.get("reason", ""),
        "show_status": status.lower() != "none",
        "status": status,
    }
    return render_pdf("credit_note.html", context, out_path)

def _photo_ctx(data, key):
    """Each photo slot carries zoom/x/y (CSS translate()+scale(), for pan/zoom
    within its mask — see photo_cell's own comment in sololuce_datasheet.html
    for the translate math) and mask (the box-inset %, user-adjustable — used
    to be a hardcoded 70 for all three). Defaults reproduce the original
    fixed centered/cropped-at-70% look for old drafts saved before this existed.

    placeholder is the "show a dashed box + name here when this slot has no
    photo yet" checkbox (app.py's CAT_IMG[slot].placeholder, one per slot) —
    defaults to True (today's long-standing look) when absent, so a draft
    saved before this checkbox existed renders exactly as it always has,
    and the user only ever has to touch it to turn a slot OFF."""
    return {
        key: data.get(key) or "",
        f"{key}_zoom": data.get(f"{key}_zoom") or 1,
        f"{key}_x": data.get(f"{key}_x") if data.get(f"{key}_x") is not None else 50,
        f"{key}_y": data.get(f"{key}_y") if data.get(f"{key}_y") is not None else 50,
        f"{key}_mask": data.get(f"{key}_mask") or 100,
        f"{key}_placeholder": data.get(f"{key}_placeholder", True),
    }

# Bolds "by request" wherever it appears in the note text (case-insensitive)
# — covers the default note ("...available by request.") and any custom note
# a user types containing the same phrase — without requiring the note
# textarea itself to hold raw HTML (it's plain text; the user would otherwise
# see literal <b> tags while editing). Safe to do unconditionally: this
# template's Jinja environment has autoescape off (see _jinja_env), same as
# every other rich-text-from-plain-field spot in this app (e.g. QTN2's own
# Attn/Address boxes), so the inserted tag renders as real bold, not escaped
# text.
_BY_REQUEST_RE = re.compile(r"by request", re.I)
def _bold_by_request(text):
    return _BY_REQUEST_RE.sub(lambda m: f"<b>{m.group(0)}</b>", text) if text else text

# Datasheet footer shows the bare domain, no "www." — scoped to this one
# call site (not BRAND_WEBSITE itself) since other doc types' own templates
# still expect the full "www.sololucelightings.com" form.
_WWW_RE = re.compile(r"^www\.", re.I)
def _strip_www(url):
    return _WWW_RE.sub("", url or "")

# Glues a number to a trailing unit so e.g. "100*50 mm" can't wrap into an
# orphaned "mm" on its own line in the Ordering Table's narrow columns.
# Applied here (not just at typing time in the frontend) so it also fixes
# values saved before this normalization existed — those still have a plain
# space baked into the stored string, which a JS input-side fix alone could
# never retroactively reach.
#
# Originally glued with an actual U+00A0 NO-BREAK SPACE character instead of
# a CSS no-wrap span. Byte-verified correct on disk (Â , valid UTF-8
# for U+00A0) and confirmed to render fine in a generic sans-serif font —
# but confirmed, via a direct render+pixel-crop check, to make this
# template's self-hosted Gotham webfont specifically DROP the unit that
# follows it in Chromium's print pipeline ("1500 lm" printed as bare
# "1500", the "lm" simply never drawn — not a wrap, not a glyph swapped for
# tofu, just gone) even though the exact same "lm" text renders correctly
# elsewhere in the same PDF when preceded by a plain U+0020 space. Whatever
# Gotham's own kerning/ligature table is doing with that specific pairing, a
# CSS `white-space:nowrap` span gets the identical "never wrap here" result
# without depending on this font having a working glyph for it, so that's
# the fix rather than hunting for a different special-space character that
# might just fail the same way. Returns Markup (the template renders
# cell.first/cell.rest with `|safe`) since escaping now has to happen here,
# before the nowrap span's own HTML is added — any user-typed HTML-special
# character elsewhere in the value still gets escaped exactly as Jinja2's
# own autoescaping would have done.
_ORD_UNIT_RE = re.compile(r"(\d)\s+(mm|lm|W|V|K|deg|°)$")
def _ord_glue_unit(text):
    escaped = str(escape(text or ""))
    return Markup(_ORD_UNIT_RE.sub(lambda m: m.group(1) + '<span style="white-space:nowrap"> ' + m.group(2) + '</span>', escaped))

# Cell text: font-size:9px, line-height:1.3 (sololuce_datasheet.html's own
# data-cell CSS) -> each line is 9*1.3=11.7px tall, converted px->pt at
# 0.75pt/px (96dpi->72pt, the same conversion used throughout this file's
# own width calibration comments) since every other size on this page is
# pt-based. Deliberately does NOT add the cell's own 3px+3px padding on
# top — CSS min-height applies to the CONTENT box only under this
# template's default box-sizing:content-box, so the browser already adds
# that padding on top of whatever min-height is set, same as it does for
# a row's own natural (non-minimum) content-driven height. Adding padding
# into this constant too double-counted it — confirmed directly: a row
# already exactly at the 2-line standard grew anyway when Align Rows was
# toggled on, because the applied min-height was quietly larger than the
# row's own real height by exactly one padding's worth.
_ORD_LINE_HEIGHT_PT = 9 * 1.3 * 0.75

def _ord_row_min_height_pt(rows):
    """The "Align Rows" button's own height target: NOT the tallest row in
    the table (that would make every short row balloon out to match one
    rare outlier — confirmed against a real example, only 1 of 8 real rows
    in a real datasheet needed 3 lines anywhere, the other 7 all topped
    out at 2) and NOT a hardcoded line count either (fragile — breaks the
    moment a product's own data shape differs). Instead: the MODE — the
    line count most of this table's own rows actually need — confirmed
    directly to land on exactly the same value a human reviewer picked by
    eye ("keep the standard the same as the 21W row") when checked against
    real data, since a "normal" row for a given product is, definitionally,
    whatever most of its own rows already look like.
    This is a MINIMUM (CSS min-height, not a fixed height or max-height) —
    a row needing MORE than the mode (like this same real example's 5W row,
    3 lines from Ring+Diffuser) still renders every line; only rows already
    AT OR BELOW the mode get pulled up to it. No content is ever hidden or
    compressed by this — see the docstring on build_ordering_table's own
    align_rows param for why that matters here specifically."""
    if not rows:
        return 0.0
    row_line_counts = []
    for row in rows:
        line_count = max((1 + len(cell.get("rest") or []) for cell in row), default=1)
        row_line_counts.append(line_count)
    standard_lines = Counter(row_line_counts).most_common(1)[0][0]
    return standard_lines * _ORD_LINE_HEIGHT_PT

def build_ordering_table(raw_cols, default_weights=None, align_rows=False):
    """Shared by render_datasheet_pdf (below) and the /api/cat-ordering-widths
    route (app.py) that feeds the sidebar's draggable column-width widget —
    single source of truth for the auto-weight algorithm so the widget's
    "what the auto width would be" never drifts from what the PDF actually
    renders. raw_cols: ordering_columns [{label, values: [text, ...]},
    optionally "width": <positive number>]. A column with an explicit
    "width" uses it verbatim as its weight (set by the user dragging that
    column's edge in the widget) instead of the auto-computation below.
    default_weights: optional {label: weight} — the saved "standard" widths
    (config key cat_ordering_default_widths, both callers load and pass
    this in), used as the fallback BASELINE for any column with no manual
    "width" of its own, ahead of the content-based auto-computation further
    down. This is what makes "Reset Widths" in the sidebar revert to the
    saved standard rather than raw content-sizing once one exists — Reset
    just clears every column's manual "width" and re-renders, so it
    automatically lands on whichever of these two layers still applies.
    align_rows: the "Align Rows" button's own toggle (per explicit
    request — "auto align the crafting boxes... row alignment/baseline").
    A row's own height is otherwise sized purely to its own tallest cell
    (default CSS Grid behavior for a shared grid, see col_template's own
    comment on why every row lives in ONE grid), so a row where every
    field happens to be short renders visibly shorter than a row that
    needed a 3-line cell — real, not a bug, but reads as an uneven table
    when several rows are all short and only one or two are genuinely
    tall. This computes a per-cell CSS min-height (not a hard cap — see
    _ord_row_min_height_pt's own docstring for why a cap would risk
    hiding real content) baked into every DATA cell (never the header
    row), pulling short rows UP to a shared "standard" height without
    ever compressing a genuinely-tall row's own content."""
    labels = [c.get("label", "") for c in raw_cols]
    col_values = [c.get("values") or [] for c in raw_cols]
    # CCT and Beam Angle print as ONE shared cell spanning every data row,
    # not one value per row like every other column — per explicit
    # request: a customer reading the printed sheet could easily read the
    # normal one-cell-per-row layout as "this row's Power only comes in
    # this row's CCT/Beam Angle", when in reality every value listed for
    # these two columns is available across every Power/Size variant, not
    # paired row-by-row. Just hiding the cell borders (tried first, in
    # discussion) wouldn't actually fix that — the values would still sit
    # row-aligned directly across from specific Power rows, which is what
    # implies the pairing, border or no border. A real spanning cell (see
    # `merged` below, and grid-row/grid-column in the template) reads
    # unambiguously as "these apply across the board" instead. Matched by
    # label (same normalize+compare the front end's own
    # isCatOrdCctColumn/isCatOrdBeamAngleColumn use), not a fixed column
    # index — this table's column order is fully user-configurable, and
    # only these two get this treatment, every other column (including
    # ones that might happen to repeat the same value on every row too)
    # keeps printing one-per-row exactly as before, per explicit scope.
    merge_col_idxs = {i for i, lbl in enumerate(labels) if _norm_spec_label(lbl) in ("cct", "beamangle")}
    # Cut Out is a plain dimension spec, not a code+description pair like
    # CCT/Controls/Size/Finish Options — bolding its one and only line never
    # had anything to visually contrast against, so it's excluded here
    # rather than sharing every other column's bold-first-line treatment.
    # Lumen is excluded for the same reason once it holds more than one
    # line (one figure per Luminare Efficacy grade, see
    # recomputeCatSpecialOrdColumns in app.py) — those lines are peers, not
    # a code+description pair, so bolding only the first one read as an
    # arbitrary distinction between two equally-important numbers.
    bold_flags = [c.get("label", "").strip() not in ("Cut Out", "Lumen") for c in raw_cols]
    max_rows = max((len(v) for v in col_values), default=0)
    rows = []
    for i in range(max_rows):
        row = []
        for c_idx, (vals, bold) in enumerate(zip(col_values, bold_flags)):
            if c_idx in merge_col_idxs:
                continue  # rendered once, merged — see `merged` below, not per-row
            raw = vals[i] if i < len(vals) else ""
            lines = [_ord_glue_unit(l) for l in str(raw or "").split("\n")]
            # col_index (0-based, true position among ALL columns, merged
            # ones included) is what the template uses to tell "is this
            # really the last column" for its right-border — loop.last
            # within this row's own (now possibly shorter) cell list would
            # answer that wrong whenever a merged column comes after this
            # one, since this row's array skips it entirely above.
            row.append({"first": lines[0], "rest": lines[1:], "bold": bold, "col_index": c_idx})
        # align_spacer: whether a bold-less cell (Cut Out/Lumen) in THIS row
        # should reserve the blank spacer line that pushes its value onto
        # line 2 (see the template's own comment on this). Only when the
        # row actually HAS a taller sibling to align with — some bold cell
        # here genuinely shows a 2-line code+description — not unconditionally
        # for every row: a real product found live left every bold column
        # (Model No., Size, CCT...) blank and only filled Power/Lumen, so no
        # cell in any row was ever 2 lines to begin with; reserving the
        # spacer anyway added 1 dead line to all 6 rows for nothing to align
        # against, enough extra height to push the whole Ordering Table onto
        # a second page. Computed once per row (not per cell) since it's a
        # row-wide question — does ANY cell here need the second line — not
        # a per-cell one.
        row_needs_second_line = any(c["bold"] and c["rest"] for c in row)
        for cell in row:
            if not cell["bold"]:
                cell["align_spacer"] = row_needs_second_line
        rows.append(row)
    # The merged cell itself: one shared cell per merged column, spanning
    # every data row (grid-row in the template) — its own DISTINCT values
    # only (first-seen order, never repeated), since the whole point is
    # showing what's available, not how many rows happen to repeat it.
    merged = []
    for c_idx in sorted(merge_col_idxs):
        seen = []
        for raw in col_values[c_idx]:
            raw = (raw or "").strip()
            if raw and raw not in seen:
                seen.append(raw)
        options = []
        for raw in seen:
            lines = [_ord_glue_unit(l) for l in raw.split("\n")]
            options.append({"first": lines[0], "rest": lines[1:]})
        # col is 1-based (CSS grid-column line numbering) — matches
        # col_template's own track order below, position c_idx+1.
        # "options", not "items" — a dict's own .items() bound method
        # shadows a same-named "items" key under Jinja's getattr-first `.`
        # attribute resolution (confirmed directly: {% for x in m.items %}
        # raised "'builtin_function_or_method' object is not iterable"),
        # so the key needs a name that isn't also a dict method.
        merged.append({"col": c_idx + 1, "options": options})
    # Every column used to get an identical 1/N share of the table's width,
    # which meant a genuinely long value (e.g. Controls' "Non-Dimmable") had
    # no more room than a short one (CCT's "830") and wrapped or truncated
    # even though a neighboring column was sitting on unused space. Instead,
    # each column's width is weighted by the longest line it actually holds
    # — deliberately based on its VALUES, not its label: a label like
    # "Finish Options" or "Input Voltage" is long but its real values ("BK",
    # "220-240V") are short, and sizing by the label would permanently steal
    # width from short-labeled-but-long-valued columns like CCT for no
    # benefit (the label is fixed chrome, not the data the reader needs).
    # The label only sets the width when the column has no values yet to go
    # by (capped low — a still-empty column shouldn't claim a full label's
    # width). Clamped to a floor (so an empty short-label column doesn't
    # collapse to nothing) and a ceiling (so one pathological custom value
    # can't squeeze every other column down to illegible). text-overflow:
    # ellipsis in the template stays as the final safety net for whatever
    # still doesn't fit. A user-dragged column skips all of this — its
    # "width" is used as-is, no floor/ceiling clamp, since the user just
    # explicitly chose it.
    col_weights = []
    for c_idx, label in enumerate(labels):
        manual = raw_cols[c_idx].get("width") if c_idx < len(raw_cols) else None
        if isinstance(manual, (int, float)) and manual > 0:
            col_weights.append(manual)
            continue
        standard = (default_weights or {}).get((label or "").strip())
        if isinstance(standard, (int, float)) and standard > 0:
            col_weights.append(standard)
            continue
        # Measured straight off the RAW values (col_values), not off
        # rows[...]["first"]/["rest"] — those went through _ord_glue_unit,
        # which wraps a trailing unit ("170 mm" -> '170<span style="white-
        # space:nowrap"> mm</span>') to stop the font from dropping it. That
        # wrapper is ~35 extra characters of HTML that were never meant to
        # be *measured*, only rendered — counting them as "content" made
        # any value ending in mm/lm/W/V/K/deg/° (Cut Out, CCT, Size, Lumen…)
        # look 5-8x longer than what's actually on the page, slamming the
        # column straight into the 16-char ceiling for genuinely short data
        # like "170 mm". Bug, not a real width need — confirmed by direct
        # fitz measurement against the real AQUA draft: Cut Out/CCT were
        # rendering at the same inflated width as columns with real long
        # values, while their own actual text sat in the middle of the box
        # with tens of points of unused margin on both sides.
        # Digits count as 1.3 "characters" each, not 1 — confirmed by direct
        # fitz measurement against a real render: "2700" (all-digit) measured
        # 4.38pt/char against this algorithm's own ~3.87pt-per-weight-unit
        # conversion at a typical 12-column total, i.e. a digit needs
        # roughly 1.13x a flat character's worth of room; rounded up to 1.3
        # for a small safety margin rather than leaving values right back at
        # the ragged edge this was measured fixing. Deliberately NOT a
        # blanket multiplier on the whole line (unlike the header's ×1.5) —
        # that would re-inflate letter-heavy values like Ring's "Square
        # Frame" that already size correctly, recreating the exact
        # "column 2x too big for its data" bug this file's column_weights
        # history already fixed once. Only digit-heavy values (CCT's
        # "2700 K", Cut Out's "170 mm") were actually measured running with
        # almost no right margin (as little as 0.19pt) — this targets that.
        longest = 0
        for raw in col_values[c_idx]:
            for line in str(raw or "").split("\n"):
                weighted = sum(1.3 if ch.isdigit() else 1 for ch in line)
                longest = max(longest, math.ceil(weighted))
        # The label itself has to fit too, not just the values — and unlike
        # a value's "\n"-separated lines, the label only gets to wrap where
        # CSS finds a space to break at. A single-word label (CCT, Ring,
        # Diffuser, Lumen...) can't wrap at all, so its full length is a
        # hard floor on the column; a multi-word one (Input Voltage, Finish
        # Options...) only ever needs room for its longest single word once
        # wrapped. Skipping this used to let a short-valued-but-long-
        # labeled column (Lumen's "500" vs the word "Lumen" itself) get
        # sized purely off its short values, so the header word had nowhere
        # to go and visibly overflowed its own cell ("Lumen" -> "Lume").
        label_words = label.split()
        label_floor = max((len(w) for w in label_words), default=0) if len(label_words) > 1 else len(label)
        # ×1.5 on top of that: every header renders bold (font-weight:700,
        # see the template) while this whole character-count model was
        # tuned primarily against the body cells' mostly-regular-weight
        # text — bold glyphs run measurably wider per character than a flat
        # +1 accounted for. Confirmed by direct fitz measurement against a
        # real 12-column render (AQUA): "Model"/"Power" — both real, already
        # rendering header words — measured ~5.2pt wide per character at
        # this size/weight, against this algorithm's own ~3.6pt-per-weight-
        # unit conversion at a typical 12-column total, i.e. each label
        # character needs roughly 1.5 weight-units' worth of room, not 1.
        # word-break stays "normal" here per an explicit "never cut a
        # header word in two" rule, so there's no wrap-based safety net to
        # fall back on if the floor comes up short — this buffer has to be
        # generous enough on its own, hence rounding UP (math.ceil), not to
        # nearest.
        if label_floor:
            label_floor = math.ceil(label_floor * 1.5)
        longest = max(longest, label_floor)
        if not longest:
            longest = min(len(label), 8)
        # +2 "characters" flat, on every column, to stand in for the
        # cell's own horizontal padding (12.4px header / 12px body — a
        # FIXED per-column cost, not proportional to content length, so it
        # has to be added post-clamp or a short column's weight wouldn't
        # actually reserve enough room for it once normalized against a
        # dozen others). Confirmed necessary by direct measurement: without
        # it, a weight exactly equal to a label's own character count
        # reserved a percentage share that, in real points, landed a few pt
        # short of that label's actual rendered width — enough to force an
        # otherwise-unnecessary line wrap on a short single-word column.
        col_weights.append(max(4, min(longest, 16)) + 2)
    label_objs = [{"text": t, "weight": w} for t, w in zip(labels, col_weights)]
    for row in rows:
        for cell in row:
            # cell["col_index"] (set above), not enumerate(row) — a row's
            # own cell list now skips merged columns entirely, so its
            # position in THIS list no longer matches its true column
            # position once any column before it is merged.
            cell["weight"] = col_weights[cell["col_index"]]
    for m in merged:
        m["weight"] = col_weights[m["col"] - 1]
    # Grid tracks: plain percentages, not `fr` and not `ch`. Three earlier
    # attempts each broke a different way — worth recording since it's not
    # an obvious choice at first glance:
    #   `minmax(min-content, Nfr)` — the obviously "correct" one: never
    #   shrink a column below what its own content needs. Backfires on a
    #   real, densely-filled table: min-content is a HARD floor, and once
    #   the SUM of every column's min-content (across a dozen-plus columns
    #   and every row sharing this one grid — see "ONE grid spanning
    #   header + every row" below) exceeds the table's actual width, the
    #   grid has no choice but to grow past its container to honor that
    #   floor, silently overflowing past the rounded-corner wrapper's
    #   overflow:hidden — every column pinned at its own min-content
    #   regardless of fr weight (confirmed: two renders with different fr
    #   values for the same column produced byte-identical column widths,
    #   because neither had any spare space left to redistribute).
    #   `minmax(0, Nfr)` — never overflows, but nothing stops an
    #   individual track shrinking below its own header word's width, so
    #   short single-word labels ("Model", "Power", "Beam") got broken
    #   mid-letter under pressure.
    #   `minmax(Nch, Nfr)` — bounded and predictable in THEORY (N
    #   characters at this font's width, using the same weight number
    #   col_weights already computes) — until confirmed directly that
    #   Chromium's print pipeline resolves `ch` measurably differently
    #   than screen for this self-hosted font: identical HTML, screen
    #   layout fit all 12 columns with room to spare, print pushed the
    #   last two off the page entirely. Whatever the exact cause, `ch` is
    #   a font-metric-dependent unit and print vs. screen metrics
    #   evidently aren't guaranteed to agree.
    # A plain percentage has none of these failure modes: it's pure
    # arithmetic (each column's share of the weight total), identical in
    # every rendering context because there's no content- or font-
    # dependent measurement involved in resolving it at all, and the
    # percentages are computed to sum to exactly 100% by construction, so
    # the table can neither overflow nor leave a gap. `overflow-wrap` on
    # the label/cell text in the template is what protects an individual
    # column that ends up too narrow for its own content — no longer a
    # rare last-resort case as it was with the ch/min-content floors, but
    # the one part of this that reliably behaves the same in every
    # browser context regardless of unit-resolution specifics.
    total_weight = sum(col_weights) or 1
    # this is what let one particular column balloon to ~270px on its own
    # in testing), `Nch` is bounded and predictable — it's "N characters
    # at this font's average width" using the SAME weight number
    # col_weights already computed (already clamped 4-16, i.e. already a
    # deliberately-bounded character-count estimate). The SUM of every
    col_template = " ".join(f"{w / total_weight * 100:.4f}%" for w in col_weights)
    row_min_height_pt = _ord_row_min_height_pt(rows) if align_rows else None
    return {"labels": label_objs, "rows": rows, "col_template": col_template, "col_weights": col_weights,
            "row_min_height_pt": row_min_height_pt, "merged": merged}

def _norm_spec_label(label):
    return re.sub(r"\s+", "", (label or "").strip().lower())

def _compact_multivalue_spec_line(value):
    """Collapse a '\\n'-joined multi-value spec (e.g. Luminare Efficacy's
    "100 lm/W\\n120 lm/W", one figure per Power tier) into one tidy printed
    line, same "3 or fewer -> comma list, more -> min-max range" rule the
    frontend's formatCatSpecPowerValue (app.py) already uses for Power's own
    multi-wattage display. Applied here, read-only, at render time only —
    the value handed in (and the saved draft it came from) keeps every
    value on its own line; see collectCatData()'s comment on why Efficacy
    can't afford Power's save-time compacting the way Power itself does."""
    vals = [v.strip() for v in (value or "").split("\n") if v.strip()]
    if not vals:
        return ""
    if len(vals) <= 3:
        return ", ".join(vals)
    nums, unit = [], None
    for v in vals:
        m = re.search(r"[\d.]+", v)
        nums.append(float(m.group()) if m else None)
    if any(n is None for n in nums):
        return ", ".join(vals)
    unit_match = re.search(r"[^\d.]+$", vals[0])
    unit = unit_match.group().strip() if unit_match else ""
    return f"{min(nums):g}-{max(nums):g}{unit}"

def render_datasheet_pdf(data, out_path, brand=None):
    """data keys: product_name, series, description, main_photo, lifestyle_photo,
    dimension_diagram (data-URI strings, any may be blank) each paired with
    <key>_zoom/_x/_y/_mask/_placeholder (see _photo_ctx), badges [{key,value}],
    specs [{label,value}], note, finish_colors [{hex,label}], ordering_code_example,
    ordering_columns [{label, values: [text, ...]}] — the form lets each column
    grow its own independent list of values (one column can have more entries
    than another, e.g. Model No. has 1 while Power has 5 wattage variants), but
    the table is rendered as genuine rows here: row i pulls column[c].values[i]
    for every column, left blank where a column has fewer values than the
    tallest one. A value's text may contain '\\n' for the two-line bold-code/
    plain-value style seen in the real sheets.

    <key>_placeholder (default True if absent, so a draft saved before this
    field existed keeps today's look) is a per-slot, user-owned checkbox —
    not an automatic preview-vs-generate switch — see _photo_ctx's own
    comment. Same value renders in the live preview and the real generated
    PDF alike, true WYSIWYG like every other photo-adjust control here."""
    brand = (brand or "SOLOLUCE").upper()

    # cat_ordering_default_widths (the saved "standard" column widths, see
    # load_cfg()'s own comment on that config key) is a GLOBAL app setting,
    # not part of a document's own saved fields — the two app.py call sites
    # inject it into a throwaway copy of `data` right before calling this
    # renderer, specifically so it never gets baked into a saved draft/
    # sidecar JSON (which would freeze a stale copy of a setting that's
    # meant to keep reflecting whatever the CURRENT standard is).
    ordering_table = build_ordering_table(
        data.get("ordering_columns") or [],
        default_weights=data.get("cat_ordering_default_widths"),
        align_rows=bool(data.get("ordering_align_rows")))

    context = {
        "website": _strip_www(BRAND_WEBSITE.get(brand, "")),
        "product_name": data.get("product_name", ""),
        "series": data.get("series", ""),
        "description": data.get("description", ""),
        "badges": badges_for(data.get("badges")),
        "specs": [{
            "label": s.get("label", ""),
            "value": _compact_multivalue_spec_line(s.get("value", ""))
                if _norm_spec_label(s.get("label", "")) == "luminareefficacy" else s.get("value", ""),
        } for s in (data.get("specs") or [])],
        "note": _bold_by_request(data.get("note", "")),
        "finish_colors": [{"hex": c.get("hex", "#ffffff"), "label": c.get("label", "")} for c in (data.get("finish_colors") or [])],
        "ordering_code_example": data.get("ordering_code_example", ""),
        "ordering_table": ordering_table,
        # The 3 free zones of the 2x2 "Dimension Diagram" grid (bottom-right
        # stays fixed as the diagram itself, via _photo_ctx below) — each a
        # generic photo with its own user-typed label (Optional Accessories,
        # a second angle, whatever that product's sheet needs there).
        "extra_photo_1_label": data.get("extra_photo_1_label", ""),
        "extra_photo_2_label": data.get("extra_photo_2_label", ""),
        "extra_photo_3_label": data.get("extra_photo_3_label", ""),
        # Dimension Diagram's own optional caption — same "printed only if
        # typed, no reserved space otherwise" contract as the 3 extra zones
        # above, see photo_cell_captioned's own comment in the template.
        "dimension_diagram_label": data.get("dimension_diagram_label", ""),
    }
    context.update(_photo_ctx(data, "main_photo"))
    context.update(_photo_ctx(data, "lifestyle_photo"))
    context.update(_photo_ctx(data, "dimension_diagram"))
    context.update(_photo_ctx(data, "extra_photo_1"))
    context.update(_photo_ctx(data, "extra_photo_2"))
    context.update(_photo_ctx(data, "extra_photo_3"))
    # Mask anchor — WHERE the mask box itself sits within its frame when
    # Mask Size is under 100% (independent of mask_x/mask_y's own job,
    # which is how the PHOTO pans *inside* the mask — two different boxes,
    # two different positions). 0/50/100 on each axis, same convention
    # photo pan already uses, converted to flex-start/center/flex-end in
    # the template (see photo_cell_captioned's own comment). Default
    # 100/100 (bottom-right) matches the fixed hardcoded position these 4
    # zones always used before this was configurable, so an existing
    # saved product with no anchor fields yet renders identically to
    # before. Scoped to just these 4 — same "Dimension Diagram + the 3
    # Extra Photo zones" set the Align buttons themselves are scoped to —
    # Main/Application Photo keep their own fixed anchor, not configurable.
    for _key in ("dimension_diagram", "extra_photo_1", "extra_photo_2", "extra_photo_3"):
        context[f"{_key}_mask_anchor_x"] = data.get(f"{_key}_mask_anchor_x", 100)
        context[f"{_key}_mask_anchor_y"] = data.get(f"{_key}_mask_anchor_y", 100)
    # The top row of the 2x2 grid (Extra Zone 1 top-left, Extra Zone 2
    # top-right) is only worth its own vertical space when there's
    # something there — a photo, or the user explicitly checked "Show
    # this zone" to reserve the space ahead of adding one later. With
    # neither, per explicit request, the whole top row is omitted rather
    # than rendering two empty dashed placeholders, and the space goes
    # back to the Ordering Table below (the Application Photo and the
    # Extra Zone 3 + Dimension Diagram row both already anchor to the
    # bottom of their column, so they simply move up to fill the gap).
    context["extra_top_row_active"] = bool(
        context["extra_photo_1"] or data.get("extra_photo_1_show")
        or context["extra_photo_2"] or data.get("extra_photo_2_show")
    )
    # Merge lives on Top Left (extra_photo_1) specifically — checking it
    # combines Top Left + Top Right into one full-width zone using Top
    # Left's own photo/zoom/pan/mask, for a panoramic shot too wide/short
    # for either square-ish zone alone. Top Right's own upload (if any)
    # just goes unused while merged, never touched, so unmerging brings it
    # right back.
    context["extra_top_merged"] = bool(data.get("extra_photo_1_merged"))
    # Optional, only meaningful once merged — lets the merged zone's own
    # box height follow its real photo's aspect ratio instead of always
    # claiming its default flex share, freeing up leftover space for
    # Bottom Left/Bottom Right below (see the template's own
    # autoSizeTopRowToPicture comment for the full mechanics). Off by
    # default — an old saved draft with no such field yet renders exactly
    # as it always has.
    context["extra_top_autosize"] = bool(data.get("extra_photo_1_autosize"))
    # Manual override, on top of the automatic aspect-ratio math — explicit
    # request: auto-size can only ever hug the uploaded PICTURE FILE's own
    # outer pixel dimensions, it has no way to know whether that file has
    # real content baked right up to its edges or a wide white margin
    # around it, so there's no way to make the automatic value "smarter"
    # about that. This just hands over a direct number the user can type
    # (in app.py's own popover) that wins outright over the computed one —
    # see the template's own autoSizeTopRowToPicture comment for the
    # clamping/fallback details. 0/absent keeps pure automatic behavior.
    try:
        context["extra_top_autosize_h"] = int(data.get("extra_photo_1_autosize_h") or 0)
    except (TypeError, ValueError):
        context["extra_top_autosize_h"] = 0
    # Same merge, same reasoning, one row down — Bottom Left (extra_photo_3)
    # + Dimension Diagram this time. Unlike the top row, this bottom row is
    # never skipped/collapsed regardless of merge state (Bottom Left +
    # Dimension Diagram always show), so merging here only ever changes
    # what's INSIDE the row, never whether it renders at all.
    context["extra_bottom_merged"] = bool(data.get("extra_photo_3_merged"))
    context.setdefault("logo_src", _brand_logo_data_uri(brand))
    return render_pdf("sololuce_datasheet.html", context, out_path)

# ----------------------------------------------------------------------------
# Full Catalog Builder (see catalog_builder.py) — Index and Pre-index, the
# only pages in the assembled book this app still generates itself. Front/
# back matter (cover, introduction, ending) and every family's own divider
# page are files the user uploads and catalog_builder.py inserts verbatim —
# this app only decides where they go and how the surrounding pages are
# numbered/tabbed, it doesn't invent their content. Sololuce-only, like the
# datasheet itself, so there's no brand parameter here — always SOLOLUCE's
# logo/website. These are called directly by catalog_builder.py, never
# through RENDERERS (that dict is doc_type-keyed for /api/generate's
# one-document-at-a-time flow; these are internal pieces of one
# multi-document assembly, not a doc_type of their own). Index and
# Pre-index are deliberately dumb/stateless — catalog_builder.py calls each
# one twice (placeholder numbers to measure page count, then real numbers
# once every page's final position is known) rather than either render
# function knowing anything about "placeholder vs final" itself.
# ----------------------------------------------------------------------------

def render_catalog_index_pdf(sections, out_path):
    """sections: [{label, start_page}] — already filtered to non-empty
    sections by catalog_builder.py before this is called."""
    context = {
        "website": BRAND_WEBSITE.get("SOLOLUCE", ""),
        "sections": sections,
    }
    context.setdefault("logo_src", _brand_logo_data_uri("SOLOLUCE"))
    return render_pdf("catalog_index.html", context, out_path)

def render_catalog_index_grid_pdf(section_label, categories, out_path):
    """The photo-grid Index for one whole SECTION — every one of its
    categories rendered as a stacked block (label+color-bar, then its own
    product grid) in ONE continuous flow, so short categories can share a
    physical page rather than each one always starting fresh (page
    insertion is page-granular — see catalog_builder.py's build_full_catalog
    Phase C/E/F — so "categories share a page" is only achievable by
    rendering them together like this, not as separate per-category PDFs).
    categories: [{number, label, tab_color, rows}], each rows:
    [{product_name, page_number, main_photo, main_photo_zoom, main_photo_x,
    main_photo_y, main_photo_mask}] — already in Index display order (and
    already excluded-filtered) by catalog_builder.py's compute_index_rows.
    number is 1-indexed within this section (resets per section), rendered
    zero-padded as "01", "02"..."""
    context = {
        "website": BRAND_WEBSITE.get("SOLOLUCE", ""),
        "section_label": section_label,
        "categories": categories,
    }
    context.setdefault("logo_src", _brand_logo_data_uri("SOLOLUCE"))
    return render_pdf("catalog_preindex.html", context, out_path)

# ----------------------------------------------------------------------------
# Expense Report (EXP) — unlike CAT, this is a normal multi-brand doc type
# (see engine.FOLDER_KEYS: no Sololuce-style brand lock), but the user
# explicitly asked for each brand to get its own genuinely distinct visual
# design rather than one shared layout with a swapped logo — so each brand
# gets its own template file instead of one template branching on `brand`.
# Falls back to Artemis's template for any future brand not in this map yet,
# same convention as engine.template_path()'s own brand fallback.
# ----------------------------------------------------------------------------
EXPENSE_TEMPLATES = {
    "ARTEMIS": "expense_artemis.html",
    "SOLOLUCE": "expense_sololuce.html",
    "ADS": "expense_ads.html",
    "WATT": "expense_watt.html",
}

def render_expense_pdf(data, out_path, brand=None):
    """data keys: company(employee name), category, number, date, period_from,
    period_to, currency, rows[{date,product,description,payment_method,amount}].
    Reuses the "company"/"project" filename slots the same way collectCatData()
    reuses "company" for product_name — see collectExpData() in app.py."""
    brand = (brand or "ARTEMIS").upper()
    currency = (data.get("currency") or "AED").strip() or "AED"

    raw_rows = data.get("rows") or []
    rows = []
    total = 0.0
    for i, r in enumerate(raw_rows):
        amount = _f(r.get("amount"))
        total += amount
        rows.append({
            "no": i + 1,
            "date": _fmt_date(r.get("date", "")),
            "product": r.get("product", ""),
            "description": r.get("description", ""),
            "payment_method": r.get("payment_method", ""),
            "amount": f"{currency} {money(amount)}",
        })

    context = {
        "brand_label": engine.BRANDS.get(brand, brand),
        "website": BRAND_WEBSITE.get(brand, ""),
        "employee": data.get("company", ""),
        "category": data.get("category", ""),
        "report_number": data.get("number", ""),
        "date": _fmt_date(data.get("date", "")),
        "period_from": _fmt_date(data.get("period_from", "")),
        "period_to": _fmt_date(data.get("period_to", "")),
        "rows": rows,
        "total": f"{currency} {money(total)}",
        # Reuses the same converter Quotation's own "Amount in Words" already
        # uses — Fils is AED-specific, so anything else falls back to the
        # generic "Cents" subunit name.
        "amount_words": number_to_words(total, currency=currency, subunit="Fils" if currency == "AED" else "Cents"),
    }
    context.setdefault("logo_src", _brand_logo_data_uri(brand))
    template_name = EXPENSE_TEMPLATES.get(brand, EXPENSE_TEMPLATES["ARTEMIS"])
    return render_pdf(template_name, context, out_path)

RENDERERS = {"QTN2": render_quotation_pdf, "CAT": render_datasheet_pdf, "EXP": render_expense_pdf,
             "INV": render_invoice_pdf, "DO": render_delivery_order_pdf, "PI": render_proforma_invoice_pdf,
             "RV": render_payment_receipt_pdf, "CN": render_credit_note_pdf}
