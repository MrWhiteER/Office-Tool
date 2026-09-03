"""
"Import from PDF" for the Sololuce Datasheet builder (CAT doc type) — local,
free, no AI/API key
--------------------------------------------------------------------------------
Reads an arbitrary manufacturer product-datasheet PDF (unknown layout, many
pages/photos) using plain PyMuPDF text/image heuristics — no external API call
of any kind, so there's no key to manage and no usage cost, ever. This
replaced an earlier Gemini-based version: Gemini's free tier turned out to
require billing/region eligibility this app's user didn't have, and paid API
usage was a hard no, so this trades extraction quality for zero cost and zero
external dependency.

Same philosophy as this codebase's other heuristic extractors
(engine.extract_product_options, engine.extract_legacy_contact): never guess
a field that isn't confidently found — an empty/missing field is fine, a
wrong value silently written into a document is not. Expect real vendor PDFs
to only get partially filled; that's the deliberate tradeoff of this being
free and offline instead of AI-read.

Boxes still work the same way as before: every field that resolves carries
the *exact* line of text it came from, and `resolve_boxes` finds that text
for real on the rasterized page with `page.search_for()` — so a field with
no box just means the matched line couldn't be re-found verbatim (rare, since
here *we* extracted the snippet directly from the same page's own text).
"""
import re

import fitz

# This used to mirror html_engine.BADGES.keys(), back when badges were a
# fixed hand-drawn icon set. Badges are now a user-editable image library
# (see html_engine.badges_for / app.py's cat_badge_library) with no fixed
# vocabulary, so detection here stays its own small, independent list of
# generic *concepts* worth flagging in an arbitrary foreign PDF — matching
# one back to a specific library image (if one exists) happens in the
# frontend's applyCatImport(), not here.
BADGE_KEYS = ["ce", "rohs", "ground", "weee", "house", "ip", "energy", "cri",
              "ugr", "warranty", "dali", "em", "sdcm", "dimmable"]
_SCALAR_SOURCE_FIELDS = ["product_name", "series", "description", "ordering_code_example"]

# Known Technical-Specifications label keywords to search for — matched as a
# whole line, case-insensitive. Covers this app's own default template
# (see app.py's CAT_DEFAULT_SPEC_LABELS) plus a handful of other common
# lighting-datasheet labels, so a foreign PDF using slightly different
# wording still has a chance of matching something.
_SPEC_KEYWORDS = [
    "Wattage", "Lifespan", "Life Span", "Light Source", "Luminaire Efficacy",
    "Luminare Efficacy", "Lamp Efficacy", "Power Factor", "Ambient Temperature",
    "Body Material", "Material", "Diffuser", "Mounting Type", "IP Rating",
    "Driver", "CCT", "Beam Angle", "Input Voltage", "Voltage", "Color Temperature",
    "Dimensions", "Size", "Cut Out", "Lumen", "Lumens", "CRI", "Controls", "Finish",
]

# Fixed badge vocabulary (must match html_engine.BADGES) detected via regex
# across the whole document's text. Deliberately conservative — "energy
# class" (A++/A+ etc) and "indoor use" aren't included here because a bare
# "A+" or the word "indoor" show up too often in ordinary marketing copy to
# match safely without a real vision read; the user adds those two by hand
# if applicable, same as any spec the heuristics miss.
_BADGE_PATTERNS = [
    ("ce", re.compile(r"\bCE\b")),
    ("rohs", re.compile(r"\bRoHS\b", re.I)),
    ("dali", re.compile(r"\bDALI\b", re.I)),
    ("dimmable", re.compile(r"\bdimmable\b", re.I)),
    ("ip", re.compile(r"\bIP\s?(\d{2})\b")),
    ("cri", re.compile(r"\bCRI\s*[:\-]?\s*(>?\s*\d+)", re.I)),
    ("ugr", re.compile(r"\bUGR\s*(<?\s*\d+)", re.I)),
    ("warranty", re.compile(r"(\d+)\s*[- ]?years?\s*warranty", re.I)),
    ("sdcm", re.compile(r"\bSDCM\s*[:\-]?\s*(\d+)", re.I)),
    ("ground", re.compile(r"\bClass\s*II\b|double[\s-]insulated", re.I)),
    ("em", re.compile(r"\bemergency\b", re.I)),
]


def _page_lines(page):
    return [l.strip() for l in page.get_text().split("\n")]


def _largest_text_span(page):
    """The biggest-font text on the page — very often the product name/title
    in a real datasheet's header. Best-effort only; never assumed to be
    right, just a reasonable starting guess the user can correct."""
    d = page.get_text("dict")
    best_text, best_size = None, 0
    for block in d.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "").strip()
                if text and span.get("size", 0) > best_size:
                    best_size, best_text = span["size"], text
    return best_text


_TABLE_HEADER_WORDS = {kw.lower() for kw in _SPEC_KEYWORDS} | {
    "model no.", "model", "power", "options", "controls", "finish color", "finish colour",
    "mounting options", "mounting", "lumen", "lumens", "no. of led", "mini cut",
}

# A few spec keywords legitimately have a value that looks like one of the
# fixed badge patterns (e.g. "IP Rating"'s value is naturally "IP65", which
# matches the 'ip' badge regex) — those are expected matches, not bleed-
# through from the badge strip, so _looks_like_real_value only rejects a
# badge-pattern hit when it belongs to a *different* badge than the one this
# keyword itself maps to.
_KEYWORD_OWN_BADGE = {"ip rating": "ip", "cri": "cri"}

def _looks_like_real_value(value, keyword=None):
    """Plain-text extraction flattens a table's columns into a linear
    sequence of lines, so a spec-keyword search can wander into the
    ordering table's own header row (e.g. matching 'CCT' there and grabbing
    'Beam Angle' — the next column header — as if it were CCT's value).
    Rejecting a candidate whose "value" is itself just another known
    label/header word catches this without needing real table structure.
    Also rejects a value that looks like a *different* fixed badge pattern
    (e.g. "CRI"'s value coming back as 'UGR< 19') — the icon-badge strip's
    own captions/values sit near each other in the page and can bleed into
    an unrelated spec's "next line" the same way table headers do."""
    v = value.strip()
    if v.lower() in _TABLE_HEADER_WORDS:
        return False
    own_badge = _KEYWORD_OWN_BADGE.get((keyword or "").lower())
    if any(pattern.search(v) and badge_key != own_badge for badge_key, pattern in _BADGE_PATTERNS):
        return False
    return True

def _find_label_value(pages_lines, keyword):
    """Search every page's lines for `keyword` as a label. Two conventions,
    both seen in real vendor sheets: 'Label: Value' on one line, or a bare
    'Label' line with the value on the next non-empty line. Scans every
    occurrence of the keyword (not just the first) and returns the first
    one whose value passes _looks_like_real_value — e.g. a datasheet's icon
    strip can repeat a label-ish word out of order, so the first raw match
    isn't always the real spec row. Returns (page_no, value, snippet), or
    None if nothing plausible was found."""
    kw_re = re.compile(r"^\s*" + re.escape(keyword) + r"\s*[:\-]?\s*(.*)$", re.I)
    for page_no, lines in pages_lines:
        for i, line in enumerate(lines):
            m = kw_re.match(line)
            if not m:
                continue
            remainder = m.group(1).strip()
            if remainder and _looks_like_real_value(remainder, keyword):
                return page_no, remainder, line
            if not remainder:
                for nxt in lines[i + 1:i + 3]:
                    if nxt.strip() and _looks_like_real_value(nxt.strip(), keyword):
                        return page_no, nxt.strip(), nxt.strip()
    return None


def extract_datasheet(pdf_path):
    """Returns a dict in the same shape the app's box-resolution/frontend
    already expect (see resolve_boxes below and app.py's
    normalizeExtractedFields): product_name/series/description (each with
    a source_page+source_snippet sibling), specs[], badges[], finish_colors
    (always empty — colors aren't reliably readable from plain text),
    ordering_code_example, ordering_columns/rows (always empty — a real
    table's column structure doesn't survive plain-text extraction, same
    known limitation documented on engine.extract_product_options; the user
    builds the ordering table by hand for these), photo_candidates."""
    doc = fitz.open(pdf_path)
    try:
        pages_lines = [(i + 1, _page_lines(doc[i])) for i in range(doc.page_count)]
        full_text = "\n".join(l for _, lines in pages_lines for l in lines)

        result = {
            "product_name": None, "product_name_source_page": None, "product_name_source_snippet": None,
            "series": None, "series_source_page": None, "series_source_snippet": None,
            "description": None, "description_source_page": None, "description_source_snippet": None,
            "specs": [], "badges": [], "finish_colors": [],
            "ordering_code_example": None, "ordering_code_example_source_page": None,
            "ordering_code_example_source_snippet": None,
            "ordering_columns": [], "ordering_rows": [], "ordering_table_source_page": None,
            "photo_candidates": [],
        }

        title = ((doc.metadata or {}).get("title") or "").strip()
        generic_titles = {"untitled", "document", "document1", "new document"}
        if title and title.lower() not in generic_titles and not re.search(r"\.(pdf|ai|indd)$", title, re.I):
            result["product_name"] = title
            result["product_name_source_page"] = 1
            result["product_name_source_snippet"] = title
        else:
            big = _largest_text_span(doc[0]) if doc.page_count else None
            if big:
                result["product_name"] = big
                result["product_name_source_page"] = 1
                result["product_name_source_snippet"] = big

        for kw in _SPEC_KEYWORDS:
            found = _find_label_value(pages_lines, kw)
            if found:
                page_no, value, snippet = found
                result["specs"].append({"label": kw, "value": value, "source_page": page_no, "source_snippet": snippet})

        for key, pattern in _BADGE_PATTERNS:
            m = pattern.search(full_text)
            if not m:
                continue
            page_no = next((pn for pn, lines in pages_lines if any(pattern.search(l) for l in lines)), None)
            result["badges"].append({
                "key": key, "value": (m.group(1) if m.groups() else "") or "",
                "source_page": page_no, "source_snippet": m.group(0),
            })

        m = re.search(r"Ordering\s+Code\s+Example\s*[:\-]?\s*([A-Za-z0-9][A-Za-z0-9\-]{4,})", full_text, re.I)
        if m:
            page_no = next((pn for pn, lines in pages_lines if any("ordering code example" in l.lower() for l in lines)), 1)
            result["ordering_code_example"] = m.group(1)
            result["ordering_code_example_source_page"] = page_no
            result["ordering_code_example_source_snippet"] = m.group(0)

        # Photos: largest embedded image per page, ranked by area — page 1's
        # biggest image is offered as "main", the next-biggest elsewhere as
        # "lifestyle". No diagram guess (too unreliable without a real
        # vision read) — the user draws that box manually, same as any
        # product photo the heuristic guessed wrong.
        photo_pages = []
        for i in range(doc.page_count):
            infos = doc[i].get_image_info()
            if not infos:
                continue
            best = max(infos, key=lambda im: (im["bbox"][2] - im["bbox"][0]) * (im["bbox"][3] - im["bbox"][1]))
            area = (best["bbox"][2] - best["bbox"][0]) * (best["bbox"][3] - best["bbox"][1])
            photo_pages.append((i + 1, area))
        photo_pages.sort(key=lambda x: -x[1])
        if photo_pages:
            result["photo_candidates"].append({"page": photo_pages[0][0], "role": "main"})
        if len(photo_pages) > 1:
            result["photo_candidates"].append({"page": photo_pages[1][0], "role": "lifestyle"})

        return result
    finally:
        doc.close()


def _resolve_one(doc, dpi, page_no, snippet):
    """page_no is 1-indexed. Returns a pixel-space box dict or None if the
    page/snippet doesn't resolve to a real match."""
    if not page_no or not snippet:
        return None
    idx = page_no - 1
    if idx < 0 or idx >= doc.page_count:
        return None
    page = doc[idx]
    hits = page.search_for(snippet.strip())
    if not hits:
        return None
    r = hits[0]
    scale = dpi / 72
    return {"page": page_no, "x0": r.x0 * scale, "y0": r.y0 * scale, "x1": r.x1 * scale, "y1": r.y1 * scale}


def _largest_image_box(doc, dpi, page_no):
    idx = page_no - 1
    if idx < 0 or idx >= doc.page_count:
        return None
    page = doc[idx]
    infos = page.get_image_info()
    if not infos:
        return None
    best = max(infos, key=lambda im: (im["bbox"][2] - im["bbox"][0]) * (im["bbox"][3] - im["bbox"][1]))
    x0, y0, x1, y1 = best["bbox"]
    scale = dpi / 72
    return {"page": page_no, "x0": x0 * scale, "y0": y0 * scale, "x1": x1 * scale, "y1": y1 * scale}


def resolve_boxes(pdf_path, extracted, dpi):
    """Augments `extracted` with a 'box' key on every field/list-item that
    has a source_page+source_snippet, and on every photo_candidate (via its
    largest embedded image on that page). Never raises on a field that
    doesn't resolve — that field just gets box=None, surfaced in the UI as
    "value only, location not found"."""
    doc = fitz.open(pdf_path)
    try:
        for field in _SCALAR_SOURCE_FIELDS:
            page_no = extracted.get(f"{field}_source_page")
            snippet = extracted.get(f"{field}_source_snippet")
            extracted[f"{field}_box"] = _resolve_one(doc, dpi, page_no, snippet)

        for key in ("specs", "badges", "finish_colors"):
            for item in extracted.get(key) or []:
                item["box"] = _resolve_one(doc, dpi, item.get("source_page"), item.get("source_snippet"))

        ord_page = extracted.get("ordering_table_source_page")
        if ord_page:
            page = doc[ord_page - 1] if 0 <= ord_page - 1 < doc.page_count else None
            extracted["ordering_table_box"] = (
                {"page": ord_page, "x0": 0, "y0": 0, "x1": page.rect.width * dpi / 72, "y1": page.rect.height * dpi / 72}
                if page else None
            )

        for cand in extracted.get("photo_candidates") or []:
            cand["box"] = _largest_image_box(doc, dpi, cand.get("page"))
    finally:
        doc.close()
    return extracted


def _crop_page_png(doc, dpi, page_no, rect_px):
    page = doc[page_no - 1]
    scale = dpi / 72
    clip = fitz.Rect(rect_px["x0"] / scale, rect_px["y0"] / scale, rect_px["x1"] / scale, rect_px["y1"] / scale)
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=clip)
    return pix.tobytes("png")


def crop_png(pdf_path, page_no, rect_px, dpi):
    """Plain crop, no AI call — used to pull an image-type field (main/
    lifestyle/diagram photo) out of the source PDF as a real PNG once the
    user has the box positioned where they want it."""
    doc = fitz.open(pdf_path)
    try:
        return _crop_page_png(doc, dpi, page_no, rect_px)
    finally:
        doc.close()


def recapture_region(pdf_path, page_no, rect_px, dpi, prompt_hint=""):
    """Reads the literal text inside `rect_px` (pixel-space dict
    {x0,y0,x1,y1} at `dpi`) directly off the page — no AI needed here at
    all, since the user has already told us exactly where to look by
    drawing/nudging the box; PyMuPDF's own clipped text extraction is more
    reliable than a vision-model guess would be for this. `prompt_hint`
    (the target field kind) isn't needed for extraction itself, only kept
    in the signature so callers don't need to change."""
    doc = fitz.open(pdf_path)
    try:
        page = doc[page_no - 1]
        scale = dpi / 72
        clip = fitz.Rect(rect_px["x0"] / scale, rect_px["y0"] / scale, rect_px["x1"] / scale, rect_px["y1"] / scale)
        return page.get_text("text", clip=clip).strip()
    finally:
        doc.close()
