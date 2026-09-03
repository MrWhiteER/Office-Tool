"""
Full Catalog Builder — combines every already-generated Sololuce Datasheet
into one bound book: whatever front-matter the user uploaded (Cover,
Introduction, and any custom pages — one reorderable list, in whatever order
the user set on the Full Catalog Builder screen), Index (Outdoor/Indoor/
Striplight), a Pre-index divider page per category, then that category's own
datasheets (alphabetical by product name, with any product family clustered
together — each cluster preceded by that family's own uploaded divider file,
if one was uploaded), repeated per section, then whatever Ending file the
user uploaded (always last, not part of the reorderable front-matter list).
This module never invents page content itself: Index and Pre-index are the
only pages it actually renders (pure page-number bookkeeping, not narrative
content) — front-matter/Ending/family-dividers are files the user supplied
via the Full Catalog Builder screen (see app.py's /api/full-catalog/extras,
/api/full-catalog/front-matter*, and /api/full-catalog/family-divider
routes, and CATALOG_EXTRAS_DIR below), inserted verbatim. A slot/item with
nothing uploaded simply contributes zero pages — no page appears in the
finished book that the user didn't either generate (a datasheet) or supply
(an extra page) themselves.

Why this is always a from-scratch pass, never incremental: the user wants
products grouped by category then alphabetical within it (and families
clustered together within that), so inserting one new product can shift
every later page number in an 800-page book. A page number/tab burned onto
an individual datasheet the moment it's generated can only be a guess
dressed up as a final answer — so individual Generate no longer stamps
anything (see app.py's /api/generate CAT branch); this module is the only
place page numbers and tabs get decided, and it only ever touches TEMP
copies of each datasheet, never the canonical per-product PDF the Datasheet
Builder itself saved.

Page-layout math needs no iterative remeasure loop. Cover/Introduction/
Ending/family-dividers are all already-finished files — their page count is
just however many pages they actually have, known the instant they're
opened, nothing about them depends on their own eventual page number. The
Index (<=3 rows) and every category's Pre-index are the only pieces that
show OTHER pages' numbers, so those still need the placeholder-then-final
two-pass treatment (render once with placeholder numbers to measure length,
then again for real once every page's final position is known); everything
else renders/resolves once. The whole layout is computed in one
deterministic pass, then Index/Pre-index are re-rendered once for real and
asserted to still match the measured page count before anything is stamped
or assembled.
"""
import os, shutil, tempfile, datetime
import fitz
import engine
import html_engine

SECTION_VALUES = ("Outdoor", "Indoor", "Striplight")
UNASSIGNED_SECTION = "Unassigned Product Type"
UNCATEGORIZED_CATEGORY = "Additional Products (No Category Set)"
UNCATEGORIZED_COLOR = "#9a9a9a"
# User-uploaded front/back matter + per-family divider files — persistent
# (never scratch/deletable, unlike _cs_draft/_cs_import), since these are
# real content the user supplied, not throwaway renders. Filenames on disk
# are always a generated uuid4 hex, never derived from user input (the
# original filename is kept only as display text in config.json) — see
# app.py's _save_catalog_extra_pdf, the one place that writes into this
# folder. Single source of truth here; app.py imports this constant rather
# than redefining its own copy of the path.
CATALOG_EXTRAS_DIR = os.path.join(engine.BASE, "catalog_extras")


def _page_count(path):
    with fitz.open(path) as doc:
        return doc.page_count


def _iter_products_raw(catalogue_folder):
    """Yields (rec, sidecar) for every CAT/.pdf file in the folder — sidecar
    is {} if missing/unreadable, so each caller decides how to handle that
    rather than this shared helper silently dropping it. Factored out so
    search_products (which only needs names/family for the Family linking
    UI, not page counts) doesn't have to open every PDF via fitz."""
    for rec in engine.index_folder(catalogue_folder):
        if rec["type"].upper() != "CAT" or rec["ext"].lower() != "pdf":
            continue
        yield rec, engine.read_sidecar(rec["path"])


def gather_products(catalogue_folder):
    """Every generated Sololuce Datasheet (CAT/.pdf) in the folder, paired
    with its sidecar data (engine.read_sidecar) and its own real page count.
    Never guesses a missing field — a product with a blank category/type is
    still included (handled by group_and_order's catch-all buckets), only a
    genuinely unreadable file is skipped, and that's warned about rather than
    silently dropped."""
    products, warnings = [], []
    for rec, sidecar in _iter_products_raw(catalogue_folder):
        if not sidecar:
            warnings.append(f'No saved data found for {os.path.basename(rec["path"])} — skipped.')
            continue
        try:
            page_count = _page_count(rec["path"])
        except Exception as e:
            warnings.append(f'Could not open {os.path.basename(rec["path"])}: {e} — skipped.')
            continue
        product_name = (sidecar.get("product_name") or sidecar.get("company") or "").strip() \
            or rec["company_label"]
        products.append({
            "path": rec["path"],
            "product_name": product_name,
            "category": (sidecar.get("series") or sidecar.get("project") or "").strip(),
            "product_type": (sidecar.get("product_type") or "").strip(),
            "family": (sidecar.get("family") or "").strip(),
            "page_count": page_count,
            # Same fields html_engine._photo_ctx already knows how to turn
            # into an <img> for the datasheet itself — reused as-is for the
            # Index grid's thumbnail (see compute_index_rows below), no
            # separate upload/lookup needed.
            "main_photo": sidecar.get("main_photo") or "",
            "main_photo_zoom": sidecar.get("main_photo_zoom") or 1,
            "main_photo_x": sidecar.get("main_photo_x") if sidecar.get("main_photo_x") is not None else 50,
            "main_photo_y": sidecar.get("main_photo_y") if sidecar.get("main_photo_y") is not None else 50,
            "main_photo_mask": sidecar.get("main_photo_mask") or 70,
        })
    return products, warnings


def search_products(catalogue_folder, query="", family="", exclude_rel=""):
    """Lightweight product lookup for the Family Tree linking UI — by name
    substring and/or exact family match. Never opens a PDF (no page count
    needed just to search/link). Returns [{"product_name","rel","family"}],
    alphabetical. `rel` is folder-relative (never a raw absolute path — the
    caller that turns a posted `rel` back into a real path, api_cat_family_link
    in app.py, re-validates containment before ever writing to disk)."""
    query = (query or "").strip().lower()
    results = []
    for rec, sidecar in _iter_products_raw(catalogue_folder):
        if not sidecar:
            continue
        rel = os.path.relpath(rec["path"], catalogue_folder).replace(os.sep, "/")
        if exclude_rel and rel == exclude_rel:
            continue
        product_name = (sidecar.get("product_name") or sidecar.get("company") or "").strip() \
            or rec["company_label"]
        fam = (sidecar.get("family") or "").strip()
        if family and fam != family:
            continue
        if query and query not in product_name.lower():
            continue
        results.append({"product_name": product_name, "rel": rel, "family": fam})
    results.sort(key=lambda r: r["product_name"].strip().lower())
    return results


def _build_physical_items(products_sorted, global_family_counts):
    """products_sorted: one category's products, already alphabetical by
    product name. Returns (physical_items, warnings), where physical_items
    is a list of either {"kind":"product","product":p} or
    {"kind":"family_divider","family":name,"members":[p,...]} (members
    themselves alphabetical), ordered so a family's divider always
    immediately precedes its own members and the sequence otherwise still
    reads roughly A-Z (a cluster sorts into place by its alphabetically-
    first member's name; a standalone product sorts by its own name).

    A family only clusters — and gets a divider — when 2+ of its members
    land in THIS category bucket. A family with just 1 member here (even if
    it has more members elsewhere, in a different category — family is
    independent of category, see the docstring above) renders that one
    member as a plain standalone product: no orphan one-item divider page.
    Warns only when a family is a TRUE global singleton (per
    global_family_counts, computed once across every product in every
    bucket before this runs) — that's the "probably a typo, nobody else uses
    this family anywhere" case, distinct from "the rest of this family lives
    in another category," which needs no warning at all.

    Crucially reuses the SAME product dicts passed in (never copies) — the
    page-arithmetic pass sets start_page while walking physical_items, and
    that mutation must be visible through the category's own `products`
    list too (Pre-index reads `products`, unchanged, and expects it)."""
    warnings = []
    by_family = {}
    standalone = []
    for p in products_sorted:
        fam = p.get("family") or ""
        if fam:
            by_family.setdefault(fam, []).append(p)
        else:
            standalone.append(p)

    units = []  # (sort_key, item_dict)
    for fam, members in by_family.items():
        if len(members) >= 2:
            units.append((members[0]["product_name"].strip().lower(), {
                "kind": "family_divider", "family": fam,
                "members": sorted(members, key=lambda p: p["product_name"].strip().lower()),
            }))
        else:
            if global_family_counts.get(fam, 0) <= 1:
                warnings.append(f'"{fam}" is set as a family but only one product uses it '
                                 f'anywhere — rendered as a standalone product, no divider page.')
            units.append((members[0]["product_name"].strip().lower(),
                          {"kind": "product", "product": members[0]}))
    for p in standalone:
        units.append((p["product_name"].strip().lower(), {"kind": "product", "product": p}))

    units.sort(key=lambda u: u[0])
    return [item for _, item in units], warnings


def group_and_order(products, category_order, section_order):
    """Buckets products into section -> category -> alphabetical-by-name,
    then within each category further into `physical_items` (see
    _build_physical_items) for family clustering. Returns (sections,
    warnings). sections is a list (already filtered to non-empty ones — a
    section/category with zero products gets no page at all), each:
    {"label", "categories": [{"label", "products": [...], "physical_items":
    [...]}]}. `products` stays pure alphabetical (Pre-index keeps reading
    this, unchanged, regardless of family clustering); `physical_items` is
    what the page-arithmetic pass actually walks. Category order within a
    section follows category_order (cat_series_labels' own array — the
    canonical sequence, not a second list that could disagree with it), with
    anything present in real data but absent from that list (e.g. a category
    since removed from Manage Lists) appended alphabetically before the
    final no-category catch-all. Section order follows section_order, same
    reasoning."""
    warnings = []
    by_section = {s: {} for s in SECTION_VALUES}
    by_section[UNASSIGNED_SECTION] = {}

    for p in products:
        sec = p["product_type"] if p["product_type"] in SECTION_VALUES else None
        if sec is None:
            if p["product_type"]:
                warnings.append(f'"{p["product_name"]}" has an unrecognized Product Type '
                                 f'({p["product_type"]!r}) — placed in "{UNASSIGNED_SECTION}".')
            else:
                warnings.append(f'"{p["product_name"]}" has no Product Type set — '
                                 f'placed in "{UNASSIGNED_SECTION}".')
            sec = UNASSIGNED_SECTION
        cat = p["category"] or UNCATEGORIZED_CATEGORY
        if not p["category"]:
            warnings.append(f'"{p["product_name"]}" has no Category/Series set — '
                             f'placed in "{UNCATEGORIZED_CATEGORY}".')
        by_section[sec].setdefault(cat, []).append(p)

    seen = {}
    for p in products:
        seen.setdefault(p["product_name"].strip().lower(), []).append(p["product_name"])
    for names in seen.values():
        if len(names) > 1:
            warnings.append(f'"{names[0]}" appears {len(names)} times across the datasheets folder.')

    global_family_counts = {}
    for p in products:
        fam = p.get("family") or ""
        if fam:
            global_family_counts[fam] = global_family_counts.get(fam, 0) + 1

    def ordered_category_labels(cat_map):
        known = [c for c in category_order if c in cat_map]
        rest = sorted(c for c in cat_map if c not in category_order and c != UNCATEGORIZED_CATEGORY)
        tail = [UNCATEGORIZED_CATEGORY] if UNCATEGORIZED_CATEGORY in cat_map else []
        return known + rest + tail

    sections = []
    section_labels = [s for s in section_order if s in by_section] + \
        [s for s in by_section if s not in section_order]
    for sec_label in section_labels:
        cat_map = by_section[sec_label]
        if not cat_map:
            continue
        cat_labels = sorted(cat_map.keys()) if sec_label == UNASSIGNED_SECTION else ordered_category_labels(cat_map)
        categories = []
        for cat_label in cat_labels:
            sorted_products = sorted(cat_map[cat_label], key=lambda p: p["product_name"].strip().lower())
            physical_items, pi_warnings = _build_physical_items(sorted_products, global_family_counts)
            warnings.extend(pi_warnings)
            categories.append({"label": cat_label, "products": sorted_products, "physical_items": physical_items})
        sections.append({"label": sec_label, "categories": categories})

    return sections, warnings


def compute_index_rows(category_products, category, cfg):
    """category_products: one category's products, already alphabetical
    (what group_and_order's `products` list already is per category).
    Overlays cfg["catalog_index_order"][category] — an explicit
    product-name order, those items first in that order, then anything not
    yet listed there appended alphabetically — and tags every item with
    whether it's in cfg["catalog_index_excluded"]. Returns the SAME product
    dicts (never copies, same reasoning as _build_physical_items), each
    with an "excluded" bool added, in full Index display order — including
    excluded ones, so the management UI can still show and re-include them.
    The actual grid render (build_full_catalog) filters excluded ones out
    itself; /api/full-catalog/index-order returns this as-is. Single
    source of truth so the two never disagree about what the Index shows."""
    order = cfg.get("catalog_index_order", {}).get(category, [])
    excluded = set(cfg.get("catalog_index_excluded", []))
    by_name = {p["product_name"]: p for p in category_products}
    rows, seen = [], set()
    for name in order:
        p = by_name.get(name)
        if p and name not in seen:
            rows.append(p)
            seen.add(name)
    for p in category_products:  # already alphabetical — anything not yet explicitly ordered
        if p["product_name"] not in seen:
            rows.append(p)
            seen.add(p["product_name"])
    for p in rows:
        p["excluded"] = p["product_name"] in excluded
    return rows


def _index_row(p, page_number):
    """One product's row for the photo-grid Index template — a fresh dict
    (not the product dict itself, same reasoning render_catalog_index_grid_pdf's
    docstring already gives: page_number varies between the measure pass
    and the final pass, everything else doesn't)."""
    return {"product_name": p["product_name"], "page_number": page_number,
            "main_photo": p["main_photo"], "main_photo_zoom": p["main_photo_zoom"],
            "main_photo_x": p["main_photo_x"], "main_photo_y": p["main_photo_y"],
            "main_photo_mask": p["main_photo_mask"]}


def summarize(catalogue_folder, cfg):
    """Cheap, read-only preview (no rendering) for a live estimate panel
    before committing to a real build. estimated_pages is a lower bound —
    it's every datasheet's own page count, not counting Intro/Index/
    Pre-index/dividers/Ending, whose real length is only known once actually
    rendered."""
    products, warnings = gather_products(catalogue_folder)
    sections, group_warnings = group_and_order(
        products, cfg.get("cat_series_labels", []), cfg.get("catalog_section_order", list(SECTION_VALUES)))
    return {
        "product_count": len(products),
        "section_count": len(sections),
        "category_count": sum(len(s["categories"]) for s in sections),
        "estimated_datasheet_pages": sum(p["page_count"] for p in products),
        "warnings": warnings + group_warnings,
    }


def _resolve_extra(info):
    """info: an item dict that may carry "stored_as" — a front_matter list
    entry, catalog_extras["ending"], or catalog_extras["family_dividers"]
    [family]. Returns the real path to that uploaded file if it's both
    recorded AND still present on disk, else None — a page deleted by hand
    outside the app, or a stale/corrupt config entry, degrades to "nothing
    uploaded" rather than a build-time crash."""
    if not info or not info.get("stored_as"):
        return None
    path = os.path.join(CATALOG_EXTRAS_DIR, info["stored_as"])
    return path if os.path.exists(path) else None


def build_full_catalog(catalogue_folder, output_path, cfg):
    """The full assembly. cfg is the app's already-loaded config dict — only
    read from here (category/section order, series colors, uploaded
    cover/introduction/ending/family-divider files), never saved; persisting
    catalog_last_build is the caller's job. Returns {"total_pages",
    "sections":[{label,start_page,end_page,category_count,product_count}],
    "warnings", "built_at"}."""
    products, warnings = gather_products(catalogue_folder)
    if not products:
        raise ValueError("No Sololuce Datasheets found in that folder yet — "
                          "generate at least one before building the catalog.")

    category_order = cfg.get("cat_series_labels", [])
    section_order = cfg.get("catalog_section_order", list(SECTION_VALUES))
    series_colors = cfg.get("cat_series_colors", {})
    # What actually PRINTS for a section — sec["label"] itself stays the
    # raw internal value (Outdoor/Indoor/Striplight) everywhere else in
    # this function, since grouping/matching needs it stable; only render
    # call sites and the returned summary use the user's own display name.
    section_labels = cfg.get("catalog_section_labels", {})
    disp = lambda lbl: section_labels.get(lbl, lbl)
    sections, group_warnings = group_and_order(products, category_order, section_order)
    warnings = warnings + group_warnings
    extras = cfg.get("catalog_extras", {})
    family_dividers = extras.get("family_dividers", {})

    with tempfile.TemporaryDirectory(prefix="sololuce_catalog_") as tmp:
        # ---- Phase B: front matter — Cover always first (its own fixed
        # slot, like Ending is fixed-last — not something that makes sense
        # to drag to a different position), then Introduction and any
        # custom pages the user added, in the order set on the Full Catalog
        # Builder screen (catalog_extras["front_matter"], a reorderable
        # list — see /api/full-catalog/front-matter-move). An item with
        # nothing uploaded contributes zero pages, not a placeholder — it
        # just doesn't affect the sequence.
        front_matter = extras.get("front_matter", [])
        cover_path = _resolve_extra(extras.get("cover"))
        intro_block_paths = ([cover_path] if cover_path else []) + \
            [p for p in (_resolve_extra(item) for item in front_matter) if p]
        intro_pages = sum(_page_count(p) for p in intro_block_paths)

        ending_path = _resolve_extra(extras.get("ending"))
        ending_pages = _page_count(ending_path) if ending_path else 0

        index_path = os.path.join(tmp, "index.pdf")
        html_engine.render_catalog_index_pdf(
            [{"label": disp(s["label"]), "start_page": 1} for s in sections], index_path)
        index_pages = _page_count(index_path)

        # ---- Phase C: measure each SECTION's photo-grid Index (placeholder
        # numbers — its length depends on which products/photos are in it,
        # not the numbers themselves) and resolve every family's divider
        # file (nothing about which file it is depends on any page number,
        # so — unlike the Index — one lookup suffices). A family with no
        # uploaded divider gets render_path=None, page_count=0 — its
        # members still cluster together physically, just with no page in
        # front of them. Index display order/inclusion is a SEPARATE
        # concern from physical book order (cat["physical_items"], family
        # clustering) — compute_index_rows overlays the user's own
        # Index Order reordering/exclusions on top of cat["products"]
        # (still plain alphabetical), same as the original hand-made index
        # files were never in the same order as the printed clustering.
        #
        # Every category in a section is rendered together as ONE PDF (not
        # one per category) so short categories can share a physical page —
        # inserting a separate per-category PDF always lands on a fresh
        # page (fitz.insert_pdf is page-granular, it can never make two
        # sources share one physical page), so packing categories together
        # is only possible by having them flow through the same render
        # call. See render_catalog_index_grid_pdf's docstring.
        index_measure_paths = {}
        for sec in sections:
            cat_number = 0
            cats_payload = []
            for cat in sec["categories"]:
                cat_number += 1
                cat["number"] = cat_number
                color = series_colors.get(cat["label"], UNCATEGORIZED_COLOR)
                cat["color"] = color
                cat["index_products"] = [p for p in compute_index_rows(cat["products"], cat["label"], cfg)
                                          if not p["excluded"]]
                cats_payload.append({"number": cat_number, "label": cat["label"], "tab_color": color,
                                      "rows": [_index_row(p, 1) for p in cat["index_products"]]})

                for item in cat["physical_items"]:
                    if item["kind"] == "family_divider":
                        div_path = _resolve_extra(family_dividers.get(item["family"]))
                        item["render_path"] = div_path
                        item["page_count"] = _page_count(div_path) if div_path else 0

            measure_path = os.path.join(tmp, f'index_measure_{id(sec)}.pdf')
            html_engine.render_catalog_index_grid_pdf(disp(sec["label"]), cats_payload, measure_path)
            sec["index_pages"] = _page_count(measure_path)
            index_measure_paths[id(sec)] = measure_path

        # ---- Phase D: pure arithmetic, every physical page's real number.
        page_cursor = 1 + intro_pages + index_pages
        for sec in sections:
            sec["start_page"] = page_cursor
            page_cursor += sec["index_pages"]
            for cat in sec["categories"]:
                for item in cat["physical_items"]:
                    if item["kind"] == "family_divider":
                        item["start_page"] = page_cursor
                        page_cursor += item["page_count"]
                        for p in item["members"]:
                            p["start_page"] = page_cursor
                            page_cursor += p["page_count"]
                    else:
                        p = item["product"]
                        p["start_page"] = page_cursor
                        page_cursor += p["page_count"]
        ending_start = page_cursor
        total_pages = ending_start + ending_pages - 1

        # ---- Phase E: final render (real numbers) + assert page count held.
        html_engine.render_catalog_index_pdf(
            [{"label": disp(s["label"]), "start_page": s["start_page"]} for s in sections], index_path)
        if _page_count(index_path) != index_pages:
            raise RuntimeError("Index page count changed between measurement and final render "
                                "— aborting rather than ship a mis-paginated catalog.")

        index_final_paths = {}
        for sec in sections:
            cats_payload = [{"number": cat["number"], "label": cat["label"], "tab_color": cat["color"],
                              "rows": [_index_row(p, p["start_page"]) for p in cat["index_products"]]}
                             for cat in sec["categories"]]
            final_path = os.path.join(tmp, f'index_final_{id(sec)}.pdf')
            html_engine.render_catalog_index_grid_pdf(disp(sec["label"]), cats_payload, final_path)
            if _page_count(final_path) != sec["index_pages"]:
                raise RuntimeError(f'Index page count changed for "{sec["label"]}" '
                                    f'between measurement and final render — aborting.')
            index_final_paths[id(sec)] = final_path

        # ---- Phase F: stamp TEMP copies of every datasheet, assemble.
        out = fitz.open()
        for p in intro_block_paths:
            with fitz.open(p) as src:
                out.insert_pdf(src)
        with fitz.open(index_path) as src:
            out.insert_pdf(src)
        for sec in sections:
            with fitz.open(index_final_paths[id(sec)]) as src:
                out.insert_pdf(src)
            for cat in sec["categories"]:
                def _stamp_and_insert(p, _cat=cat):
                    stamped = os.path.join(tmp, f'stamped_{id(p)}.pdf')
                    shutil.copyfile(p["path"], stamped)
                    engine.stamp_catalogue_page_numbers(
                        stamped, p["start_page"], tab_color=_cat["color"], total_pages=total_pages)
                    with fitz.open(stamped) as src:
                        out.insert_pdf(src)

                for item in cat["physical_items"]:
                    if item["kind"] == "family_divider":
                        if item["render_path"]:
                            with fitz.open(item["render_path"]) as src:
                                out.insert_pdf(src)
                        for p in item["members"]:
                            _stamp_and_insert(p)
                    else:
                        _stamp_and_insert(item["product"])
        if ending_path:
            with fitz.open(ending_path) as src:
                out.insert_pdf(src)

        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        tmp_out = output_path + ".tmp"
        out.save(tmp_out)
        out.close()

    if os.path.exists(output_path):
        try:
            shutil.copyfile(output_path, output_path + ".bak")
        except OSError:
            pass
    os.replace(tmp_out, output_path)

    section_results = []
    for i, s in enumerate(sections):
        end_page = (sections[i + 1]["start_page"] - 1) if i + 1 < len(sections) else (ending_start - 1)
        section_results.append({
            "label": disp(s["label"]), "start_page": s["start_page"], "end_page": end_page,
            "category_count": len(s["categories"]),
            "product_count": sum(len(c["products"]) for c in s["categories"]),
        })

    return {
        "total_pages": total_pages,
        "sections": section_results,
        "warnings": warnings,
        "built_at": datetime.datetime.now().isoformat(timespec="seconds"),
    }
