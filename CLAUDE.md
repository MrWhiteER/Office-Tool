# Office Tool

A local Flask app that generates Quotations, Tax Invoices (INV), and
Delivery Orders (DO) for **multiple lighting/electrical brands** run by the
same user: Artemis Lightings, Sololuce Lightings, ADS Lightings, Watt
Electricals. Two generation pipelines coexist: Tax Invoice/Delivery Order
(and legacy Quotations, code `QTN`) are Excel + PDF from a company template;
the current-standard Quotation (code `QTN2`) is a pixel-fidelity HTML design
rendered straight to PDF — see "Pixel-fidelity HTML document pipeline" below.

Run with `python app.py` (or via the preview tool's `launch.json`, already
configured). Needs LibreOffice installed (used for all xlsx→PDF/PNG
conversions), Pillow (for openpyxl image handling), PyMuPDF (`fitz`, for
multi-page preview rendering — see below), and Playwright with Chromium
installed (`pip install playwright && playwright install chromium` — for the
QTN2 HTML/PDF pipeline, see below).

## File map

- **`app.py`** — the entire Flask backend *and* the frontend. The frontend is
  a single big HTML/CSS/JS string (`PAGE` variable) served at `/`. There is
  no separate template engine or JS bundler — edit the string directly.
- **`engine.py`** — document generation logic for the **xlsx pipeline**
  (QTN/INV/DO): filling templates via openpyxl, the folder scanner for "All
  Docs", product photo and datasheet catalog matching, brand/template
  resolution. `engine.HTML_DOC_TYPES` marks which doc types instead use...
- **`html_engine.py`** — the **HTML/PDF pipeline** (currently just `QTN2`,
  the standard "Quotation"): renders a Jinja2 template
  (`templates_html/<name>.html`) to HTML, then prints it to PDF via
  Playwright/headless Chromium. PDF-only — no `.xlsx` sibling. See
  "Pixel-fidelity HTML document pipeline" below for the full picture; this
  is a second, parallel generation pipeline living alongside the xlsx one,
  not a replacement.
- **`templates/<BRAND>/{QTN,INV,DO}.xlsx`** — one template set per brand.
  Falls back to `templates/ARTEMIS/*` if a brand has no template of its own.
  (Unrelated to `templates_html/` above, despite the similar name — this one
  is xlsx data files, not Jinja2 templates.)
- **`config.json`** — all persisted settings. Structure:
  ```
  {
    "brand": "ARTEMIS",            // currently active brand
    "units": [...],                 // unit dropdown options (PCS/MTR/...)
    "brand_settings": {
      "ARTEMIS": { "qtn_folder", "inv_folder", "do_folder",
                   "product_photos_folder", "datasheets_folder",
                   "templates_folder" },
      "SOLOLUCE": {...}, "ADS": {...}, "WATT": {...}
    }
  }
  ```
  Every brand has **independent** folders — Artemis's document folders are
  separate from Sololuce's, and each brand's product photo/datasheet catalog
  is its own.
- **`_cs_cache/`, `_cs_draft/`** — internal render caches (thumbnails, live
  preview drafts). Never written into the user's own document folders. Safe
  to delete; they'll be regenerated.
- **`static/logos/{artemis,sololuce,ads,watt}.png`** — real brand logos used
  in the brand switcher (top-left button, dropdown cards, Settings header) in
  place of the old plain colored sphere. Served automatically by Flask's
  default `static` folder — no route code needed. Sourced from
  `F:\1. Office\1. EDGAR FILES\1. GRAPHIC\<brand>\88. LOGO\` (Artemis's came
  from the logo already embedded in `templates/ARTEMIS/*.xlsx`; Sololuce's
  was cropped out of a logo-approval mockup sheet that had a mat/border baked
  in; Watt's is `WATT-ELECTRICALS-REMASTERED.png` cropped to drop the long
  decorative swoosh tail, already background-transparent so no keying
  needed). All four brands are now wired into `BRAND_LOGOS` in `app.py` — no
  brand still falls back to the plain colored `.bulb` sphere. ADS's tile uses
  `ADS-LOGO.png` (stacked "LIGHTING L.L.C" under the mark), not
  `ADS-LOGO-OFFICIAL.png` (side-by-side) used initially.
- **`static/logos/{artemis,sololuce}_watermark.png`** — background-keyed
  (transparent) versions used for the faint full-page watermark behind the
  main content (`#watermark` div, `BRAND_WATERMARKS` map in `app.py`). Made
  by thresholding out the flat background color from the tile logos (dark
  navy for Artemis, white for Sololuce) since those still have their
  background baked in — using the tile PNGs directly at low opacity would've
  shown a faint translucent box instead of just the wordmark. ADS's and
  Watt's watermarks reuse `ads.png`/`watt.png` directly since those are
  already background-transparent — no separate `_watermark.png` needed for
  either. All four brands now have a working watermark.

## Current real-world config (as of last session)

- **Artemis**: real QTN/INV/DO folders configured (three *separate* physical
  folders — "CLIENT QUOTATION 2023", "CLIENT INVOICES", "CLIENT DELIVERY
  NOTE"). Now also has product photo/datasheet catalog configured — points at
  the same Sololuce catalog folders on F:.
- **Sololuce**: product photo catalog (566 PNGs) and datasheet catalog (500+
  PDFs, some `.ai`) configured. No document folders yet.
- **ADS / Watt**: completely blank, not set up.

## Architecture notes worth knowing before changing things

- **Brand switcher** (top-left) changes `config.json`'s `"brand"` and
  everything else — active template, active document folders, active
  product catalog — follows from that.
- **Legacy files**: thousands of pre-existing documents don't follow the
  app's naming convention (`BRAND_TYPE_NUMBER_R#_COMPANY_PROJECT_DATE.ext`).
  `engine.scan_all()` still surfaces them via heuristic parsing
  (`_guess_type`, `_guess_date`, `read_items_from_doc`) so "All Docs" isn't
  empty. Untagged legacy docs show under every brand (deliberate decision).
- **"Open in CS"**: for app-generated files, opens the full edit form. For
  legacy files, does a best-effort import (pulls line items + any embedded
  photos via `read_items_from_doc`, leaves header fields blank) — it does
  **not** overwrite the original on save, since the layout isn't the same.
- **Photo/datasheet auto-matching**: typing a line-item description and
  blurring the field searches the current brand's catalog for a matching
  product name and shows a *suggestion* (not auto-applied) for photos, and
  direct links for datasheets. Matching is substring + word-boundary based
  on the catalog filenames (see `match_product_photo` / `match_datasheets`).
- **Preview rendering**: an embedded native PDF viewer (`<iframe>`) renders
  as a solid black box in this environment — dead end, don't reintroduce it.
  Hover preview and the "Open in CS" modal render a single page-1 PNG via
  LibreOffice (`engine.to_png`, cached in `_cs_cache`) and display it as
  `<img>` — deliberately page-1-only, the CS modal even says so in its title.
  The **live Build-view preview pane** (`#previewbox`) is different: it's
  multi-page. `/api/preview-draft` converts the draft xlsx to PDF
  (`engine.to_pdf`) then rasterizes *every* page via PyMuPDF
  (`engine.to_png_pages`, into `_cs_draft/draft_<n>.png`), and
  `/draft-preview?page=<n>` serves them individually. The frontend
  (`showPreviewPages`/`renderPreviewPages` in `app.py`'s JS) lays all pages
  out in a `#previewpages` flex container — stacked in "Single" mode, two
  per row with a small gap in "Double" mode (like Adobe's spread view).
  Double is the default whenever a draft has more than one page
  (`previewModeAuto`) but a user's explicit Single/Double click sticks until
  reload. Zoom (`previewZoom`, 40–300%) and the page width itself are
  recomputed on every render/resize (`computePageWidth`) so pages hug the
  pane's real width instead of leaving big empty margins. The `#previewimg`
  single-`<img>` element from before still exists alongside `#previewpages`
  and is used by the non-multi-page flows (`setPreviewImage`, called after
  Generate and when opening an app-generated file via "Open in CS") — only
  one of the two containers is visible at a time.
- **Resizable split**: the `.left`/`.right` columns in the Build view have a
  drag handle (`#resizer`) between them controlling a `--leftw` CSS variable
  on `<html>`, persisted to `localStorage` (`cs_leftw`) so it survives
  reloads. Disabled below the existing 980px responsive breakpoint (columns
  stack instead).
- **`#previewbox` must keep a bounded `height` (not `min-height`/`flex:1`)**.
  It was briefly changed to `flex:1;min-height:...` while fixing the oversized
  side-margins, which seemed harmless but silently broke internal scrolling —
  with no height cap the pane just grows to fit all pages and the whole
  *document* scrolls instead, which also makes the pan/hand-tool math
  (`box.scrollLeft`/`scrollTop`) a no-op since there's nothing to scroll
  internally. Fixed back to `height:calc(100vh - 150px)`. If the preview area
  ever needs resizing again, keep it a fixed/capped height, not a flexed one.
- **Preview pane viewer tools** (toolbar above `#previewbox`): hand/pan tool
  (`#pm-hand`, `togglePanTool`) drags via `scrollLeft`/`scrollTop`, holding
  **Space** pans temporarily without toggling the persistent tool (standard
  Photoshop/Acrobat/Figma convention — see `spacePanning` vs `panToolOn`), and
  **Ctrl+scroll** over the pane zooms instead of scrolling. These plus
  Single/Double, zoom −/+/Fit, and the resizer are meant to be the complete
  "PDF-viewer" toolkit for this pane — before adding another viewer control,
  check this list first.
- **`.pvpage` needs `flex-shrink:0;min-width:0`**. `#previewpages` is a flex
  container, and flex items default to an automatic `min-width` capped at
  their *intrinsic* content size — for an `<img>` that's its natural pixel
  resolution (the render DPI in `engine.to_png_pages`). Without
  `flex-shrink:0`, zoom silently stopped working above ~140% (the browser was
  quietly shrinking the explicit JS-computed width back down to the image's
  natural size, e.g. 1158px at the old 140dpi) even though `img.style.width`
  kept increasing correctly — `getComputedStyle` vs `element.style` disagreed,
  which is what gave it away. Render DPI is now 170 (`engine.to_png_pages`)
  so pages stay reasonably sharp further into the zoom range before
  upscaling blur becomes noticeable.
- **Datasheet pills must never navigate the page** (`.dspill` in the
  Line Items / Datasheets panel). They used to be `<a href=... target=_blank>`
  — if the browser ever routes that to the *same* tab instead of a new one
  (popup-blocker settings, embedded/preview contexts, etc.), it fully
  unloads this single-page app and wipes every unsaved field, since all
  state is in-memory JS with no persistence. Fixed by making them `<button>`s
  that call `openDatasheetCS(rel, name)`, which opens the datasheet as a
  page-1 PNG inside the existing `csmodal` (same pattern as `openCS()` for
  documents, via a new `/datasheet-thumb` route mirroring `/cs-thumb`) —
  never a real navigation. Don't reintroduce a plain link/`target=_blank`
  here even though it'd look simpler.
- **Autocomplete datalists (`#clients`/`#projects`/`#attns`/`#addresses`)
  are populated by `loadClients()`**, called at page init, after `generate()`,
  and on brand switch — not just when visiting All Docs. This used to be a
  real bug: the company-history datalist was only ever filled inside
  `loadIndex()` (the All Docs loader), so on a fresh Build session — now the
  default landing view — the "companies quoted before" dropdown was always
  empty until you happened to click into All Docs first. `attns`/`addresses`
  come from `QTN2` sidecar files (`engine.read_sidecar`) since those fields
  aren't part of the filename convention; `clients`/`projects` come from
  filenames the same way All Docs already did. If a new free-text field
  wants the same prediction treatment, extend `/api/index`'s aggregation and
  add a matching `<datalist>` + `list=` attribute — don't hand-roll a
  separate fetch.
- **`window.prompt()` / `alert()` / `confirm()` are unreliable in this
  preview environment** — they silently no-op instead of showing a dialog.
  Any "ask the user for a value" UI must be a real inline element (a text
  input that appears in place, a small popover menu), not a native dialog.
  The custom-unit input and the file-op context menu are the reference
  patterns for this.
- **Optional Discount & VAT (QTN + INV only — DO has no price column)**: a
  "Discount & VAT" card in the Build view (`#discvat-card`, hidden for DO)
  sends `discount:{enabled,mode,value}` / `vat:{enabled,mode,value}` (mode is
  `'percent'|'fixed'`, discount also has `'target'` — type a trailing `%` to
  mean percent, otherwise it's a flat amount; Target Price mode means "make
  the final price this number" and the discount needed is back-solved,
  `=MAX(subtotal-target,0)`, never a negative discount). VAT is computed on
  **Price After Discount**, not the raw subtotal, matching the requested row
  order. The summary block (`engine._write_summary_block`) grows/shrinks
  between 2 and 5 rows (`Subtotal` always, `Discount`+`Price After Discount`
  if discount is on, `Vat` if VAT is on, `Total` always) by inserting/deleting
  rows right after Subtotal (`engine._expand_rows` / the new
  `engine._remove_rows`) and rewriting every row's label+formula fresh each
  time — it never tries to patch the template's original 3 rows in place.
  The per-line "Vat 5%" item-table column and its header label follow the
  same master VAT checkbox (blank when off or when VAT is a fixed amount,
  since a lump sum can't be split per line). Defaults (`engine._norm_vat`
  when the key is absent) reproduce the old always-on 5% VAT so any caller
  that doesn't send these fields still gets identical output to before this
  feature existed. **Known gap**: reopening a saved file via "Open in CS"
  does not restore its discount/VAT settings (`read_full_record` was not
  extended) — it resets to the defaults, same as starting a new document.

## Client database (Clients tab)

A per-brand address book — `clients/<BRAND>.json` (array of records: `id`,
`name`, `category`, `attn`, `address`, `phone`, `landline`, `email`,
`website`, `trn`, `notes`, `logo`, `updated`), managed via
`load_clients`/`save_clients` in `app.py` and `GET /api/clients` /
`POST /api/clients` (create-or-update, keyed on `id`) /
`POST /api/clients-delete`. Follows the same per-brand-isolation pattern as
folders/catalogs/templates everywhere else in this app — switching brands
shows a completely different client list.

- **`category` = "Section"** in the editor UI: a free-text field (with a
  datalist of previously-used values so existing sections are easy to
  reuse) rather than a fixed enum — trading businesses' client taxonomies
  (Contractor / Consultant / Government / Individual / Distributor, or
  whatever this brand actually uses) aren't known in advance, so this lets
  the user define sections as they go instead of the app guessing wrong
  ones up front. The Clients grid (`renderClientsGrid`) groups cards under
  a `.clientsectionhead` per distinct category (blank → "Uncategorized",
  always sorted last), each collapsible via `toggleClientSection` —
  collapsed state lives only in the in-memory `COLLAPSED_SECTIONS` set, not
  persisted, since it's a viewing convenience rather than real data.
- **`landline`/`website`/`trn` were added alongside `phone`/`email`** as the
  other fields a real client record in this line of business typically
  needs — `phone` doubles as mobile, `landline` for office numbers, `trn`
  for the Trade License/Tax Registration Number invoices reference (same
  concept as this app's own `COMPANY_INFO.trn` in `html_engine.py`).

- **Logo storage**: stored as a base64 data URL directly inside the JSON
  record (same approach this app already uses for line-item photos), not as
  a separate file + serving route. Simpler, and fine at the expected scale
  (tens–low hundreds of clients); revisit if that ever becomes thousands.
- **Two entry points, one shared editor** (`#clientmodal`,
  `openClientEditor(id)`/`saveClientEditor()`): the full "Clients" tab
  (`#v-clients`, `loadClientsView`/`renderClientsGrid` — a card grid with
  search, "+ New Client", and per-card "Use in Build"/"Delete") and a
  compact quick-picker popover at the top of Build's "Client / Company" card
  (`#clientpicker`, `openClientPicker`/`applyClientPick`/`saveCurrentAsClient`)
  both funnel into the same modal for creating/editing.
- **The quick-picker is type-aware**: `applyClientPick`/`saveCurrentAsClient`
  only touch Attn/Address when `TYPE==='QTN2'` (those fields don't exist for
  QTN/INV/DO) — Company name is the only field filled for the other types,
  via `setCompanyVal`/`companyVal` (see the rich-text section below for why
  Company needs those two accessor functions instead of a plain `.value`).
  `saveCurrentAsClient` also checks for a same-name existing client and
  updates it in place rather than creating a duplicate.
- **Delete uses the double-click-to-confirm pattern** (`deleteClient(id,btn)`
  → arms on first click, `actuallyDeleteClient` on the second within ~2.5s),
  not `window.confirm()` — same reasoning as everywhere else in this app
  that needs a yes/no from the user (see the `window.prompt()`/`alert()`/
  `confirm()` note further down): it silently no-ops in the preview/testing
  environment. Reused in both the grid cards and the editor modal's Delete
  button.
- **"Import from documents" button** (`#clientsimportbtn` → `importClients()`
  → `POST /api/clients-import`, `api_clients_import` in `app.py`): bulk-
  populate from the brand's existing document history, for onboarding a
  brand that already has years of real files but an empty Clients tab.
  Reuses `engine.scan_all()` over `all_doc_folders()` (the same call
  `/api/index` makes) to collect every distinct `company_label` — for
  matched/convention filenames that's the company token baked into the name;
  for everything else it's the top-level subfolder name, since that's how
  the real pre-app archives are organized (folder-per-client, mostly).
  Existing clients (matched by name, case-insensitive) are left alone for
  the *name* — only missing companies get added — but the contact fields
  (attn/address/phone/landline/email/website/trn) get backfilled on BOTH
  new and pre-existing records, one field at a time, never overwriting one
  that's already non-blank. Safe to re-run at any time.
  **Folder-name noise**: real archive folders aren't all client names —
  year bins, "CASH", per-document-type buckets like "ABU DHABI QTN" show up
  too. `engine.looks_like_non_client_label()` filters the obvious cases
  (pure numbers, a small non-client denylist, and anything matching the
  existing `_TYPE_KEYWORDS` doc-type regexes) before import, but it's a
  heuristic, not exhaustive — expect some near-duplicate variants too (e.g.
  "ALTA DERMA" vs. "ALTA DERMA CLINIC", same real client filed under two
  folder names over the years). The user reviews and prunes/merges via the
  normal grid Delete button after each import; nothing here tries to
  auto-merge, since guessing wrong would silently conflate two different
  companies.
- **Contact-field backfill has two sources, tried in priority order per
  company** (both only fill fields that are still blank):
  1. QTN2 sidecars (`engine.read_sidecar`) for `attn`/`address` — exact,
     structured, but only exist for documents made with the new Quotation
     pipeline (`QTN2`), which is brand new, so real archives barely have any
     yet.
  2. `engine.extract_legacy_contact(xlsx_path)` — the actual workhorse, since
     nearly all real history is legacy QTN/INV/DO. **Key discovery**: these
     hand-built xlsx files never put the customer's company/attn/phone/
     address/email into real cell values at all (`fill_quotation`/
     `fill_invoice`/`fill_delivery_order` never write a customer-info cell,
     confirmed by inspecting several real files directly) — instead the
     "letterhead" block (company's own info + a "From"/"To" pair + the
     customer's block) was typed into free-floating DrawingML text-box
     *shapes*, which `openpyxl`'s normal `ws.cell(...)` reading can't see at
     all (it silently drops shape/drawing content — even prints a
     "DrawingML support is incomplete" warning). `extract_legacy_contact`
     reads `xl/drawings/drawing*.xml` straight out of the xlsx zip instead,
     regexes out each shape's `<a:t>` text runs, discards the shape that's
     our own letterhead (matched by brand-name token — checks all 4 brands,
     since one brand's folder tree can contain another brand's documents,
     e.g. `CLIENT INVOICES/2024/ADS INVOICES/...` living under Artemis's
     configured folder) and any shape that's just a bare "From"/"To"/"Buyer"/
     "TAX INVOICE"-style label, then classifies the one remaining shape's
     text fragments field-by-field via keyword/regex (email has an `@`,
     website has `www.`, TRN has the literal word, address matches PO
     Box/Emirate names, phone/landline/fax need the label word anchored to
     the *start* of the fragment). Returns `{}` — never guesses — whenever
     more or fewer than exactly one un-claimed shape remains, or nothing
     classified.
  - `api_clients_import` tries a company's most recent 5 xlsx files (via
    `xlsx_by_company`, sorted by `scan_all`'s date) and merges results
    field-by-field (first successful file wins each field), stopping early
    once every field is filled — bounds the work since some companies have
    hundreds of historical documents.
  - **Hard-won gotcha**: phone/landline/fax matching MUST anchor the label
    word to the start of the fragment, not `.search()` it anywhere in the
    string — a free substring search for "cell" (meant to catch "Cell:")
    matches inside the ordinary word "Excellence", which silently corrupted
    a real client's phone field with the text "ITALIAN EXCELLENCE" before
    this was caught by spot-checking a few of the 55 first-run results
    against the source files. Anchoring the regex to `^` fixed it and is
    covered by the "existing_names"-style re-run test in this file's
    history if this ever needs re-verifying.
  - Also expect the very occasional wrong pickup where a company's own real
    file just doesn't have the customer's info in a discoverable shape
    (e.g. a supplier's own catalog letterhead ends up being the only
    un-claimed shape) — this produces a real but *wrong* value, not a
    crash. Same review-after-import expectation as the name-only import
    above; nothing here is meant to be 100% precision on 100% of a
    multi-year hand-built archive.
- **`CLIENT_FIELDS`/`load_clients()` normalization** (`app.py`): older client
  records (saved before `category`/`landline`/`website`/`trn` existed) don't
  have those keys — `load_clients()` fills in `""` for any missing field on
  every read so callers never need scattered `.get(f, '')` everywhere.
- **Build's Company/Attn/Address prediction is unified with the Clients
  database**, not just document history. Two independent name sources
  existed before this: `CLIENTS`/`ATTNS`/`ADDRESSES` (from `/api/index` —
  every company/attn/address ever typed into a *generated document*) and
  `CLIENT_RECORDS` (from `/api/clients` — the curated Clients-tab address
  book). A brand-new client added only via the Clients tab wouldn't show up
  while typing in Build, and picking a document-history suggestion never
  filled Attn/Address, because the two lists were never cross-referenced.
  Fixed by:
  - `allCompanyNames()`/`allAttns()`/`allAddresses()` (`app.py`) union both
    sources (deduped) — used as the `listGetter` for the Company/Attn/
    Address `attachAutocomplete()` calls (QTN2's richboxes) and to refresh
    the plain `#clients` `<datalist>` (INV/DO's non-QTN2 Company input) via
    `refreshClientNamesDatalist()`. Called after every place `CLIENT_RECORDS`
    changes (`loadClientsView`, `importClients`, `actuallyDeleteClient`, the
    picker's lazy fetch, `saveCurrentAsClient`) so Build always reflects the
    latest Clients-tab state without needing a page reload.
  - `CLIENT_RECORDS` now also loads at page init (`loadClientsView()`
    alongside `loadClients()`), not just when the Clients tab or the
    picker is first opened — otherwise Build's autocomplete would be
    missing Clients-tab-only entries until the user happened to visit that
    tab first.
  - `fillFromClientIfBlank(name)`: on every Company edit (`onCompany()`,
    which already fires on every keystroke/paste/autocomplete-pick), if the
    typed name exactly matches a client record (case-insensitive) and
    Attn/Address are still blank, fill them from that record — mirrors what
    the "Client…" picker button already did explicitly
    (`applyClientPick`), but now happens implicitly just from typing/
    picking a known company name, same as the picker. Blank-checked so it
    never clobbers something the user already typed.
  - **Bug fixed in passing**: `saveCurrentAsClient` (the picker's "+ Save
    current info as client" quick-save) only ever sent
    `name/attn/address/phone/email/notes/logo` in its POST body — since
    `api_save_client` writes whatever fields the client sends and defaults
    anything missing to `""`, using this quick-save on an *existing* client
    would have silently wiped their `category`/`landline`/`website`/`trn`
    (exactly the fields the "Import from documents" backfill just spent
    effort filling in). Fixed by carrying those fields over from the
    existing record when one exists, only ever overwriting attn/address
    (the two fields Build actually edits).

## Client Excel export (Settings + Clients tab)

`GET /api/clients-export` (optional `?id=<client id>`) exports the client
database as a real `.xlsx` — `engine.build_clients_workbook(clients)` builds
one workbook layout (`CLIENT_EXPORT_HEADERS`: Company Name/Section/Attn/
Address/PO Box/City/Country/Phone/Landline/Email/Website/TRN/Notes/Last
Updated/Logo — Country is written via `_country_display()`, flag emoji +
full name, using a small `COUNTRY_NAMES` dict that's just enough for a
readable export column; the authoritative full list for the picker lives
client-side, see below) that both call sites share:
- **Settings tab** → "Company Profiles" card → "Export All Client Profiles
  (Excel)" (`exportAllClients()`) exports every client for the current brand.
- **Clients tab** → each card's "Export" button (`exportClient(id)`) exports
  just that one client, filename `<ClientName>.xlsx` (via `engine._slug`).
- **Logo embedding reuses `_insert_photo`** (the same helper that puts item
  photos into generated QTN/INV/DO Excel files) — decodes the client's logo
  data URL and anchors it into the row's Logo column, so the export isn't
  just a flat text dump when a client has a logo saved.
- Both routes are `window.location.href = ...` navigations (not `fetch`), so
  the browser turns them into a normal file download once it sees the
  response's `Content-Disposition: attachment` header — this shows up as
  `net::ERR_ABORTED` in network logs/dev tools, which is expected and not a
  bug (the "navigation" was always going to end as a download, never a page
  load).

## Structured address (PO Box / City / Country + Google Maps) and the card Edit button

- **Explicit Edit button on client cards** (`.clientcardedit`, top-right of
  `.clientcard`, `position:absolute`): the whole card was already clickable
  to edit, but that wasn't discoverable — a visible pencil button makes it
  obvious without changing the whole-card-click behavior (both call the same
  `openClientEditor(id)`, `event.stopPropagation()` on the button so it
  doesn't also trigger the card's own handler).
- **Address split into `address`/`po_box`/`city`/`country`** (previously one
  free-text `address` field holding everything). `country` stores a 2-letter
  code (e.g. `"AE"`); `countryFlag(code)` derives the flag emoji at render
  time from Unicode regional-indicator math
  (`String.fromCodePoint(...[...code].map(ch=>127397+ch.charCodeAt(0)))`) —
  no emoji table to maintain, works for any of the ~190 entries in the
  client-side `COUNTRIES` array.
- **Country picker** (`#countrypicker`, `openCountryPicker`/
  `renderCountryPicker`/`pickCountry`): a searchable popover, same
  `position:fixed` + `getBoundingClientRect()`-positioned pattern as
  `#clientpicker` (the Build-tab quick client picker) — typing filters,
  clicking sets the hidden `#cf-country` input plus the flag/label on the
  trigger button. UAE is listed first in `COUNTRIES` (not alphabetical)
  since this is a UAE-based business and it's the overwhelmingly common
  pick.
- **Google Maps location** (`refreshClientMapPreview()`, called on every
  address/PO-Box/city input and after picking a country): builds a query
  string from whichever of address/PO Box/city/country are filled, and
  shows a keyless embedded map (`https://www.google.com/maps?q=<query>&output=embed`
  — the older `/maps?q=...&output=embed` URL works without an API key,
  unlike the newer Maps Embed API) plus a "Copy Maps Link" button
  (`navigator.clipboard.writeText` of a `/maps/search/?api=1&query=...`
  URL — built for sharing, e.g. pasting into WhatsApp) and an "Open in
  Google Maps" link. Shows an empty-state message instead of an iframe
  when there's nothing to search for yet.
- **Legacy-contact extraction (`engine.py`) now also fills po_box/city/
  country**, and fixed two data-quality problems the user caught by
  inspecting real imported records:
  1. **Company name leaking into `attn`** — e.g. "ATL" the client showed
     attn "ATL ELECTRICALS LLC Giorgio Palermo Unit 608 Business Point
     Tower" (company name AND part of the address, jammed in with the
     actual contact name). Fixed two ways: `_strip_known_company_prefix()`
     removes the trusted company name from the front of the leftover attn
     text when it's a literal (case-insensitive, whitespace-tolerant)
     prefix match — note this is a `re.match`-anchored strip, not the
     word-by-word greedy accumulation an earlier version of this function
     used, which had a bug where once the accumulated text merely *started
     with* the target, it kept accepting every further word forever. And
     address-like fragments (PO Box, a UAE city/country name, or a
     building/unit/floor/tower/villa/plot word *next to a number*) are now
     classified into `address`/`po_box`/`city`/`country` instead of
     falling through to the leftover "attn" pile — "Unit 608 Business
     Point Tower" no longer ends up in Attn.
  2. **Two new false-positive traps found and fixed while building this**:
     requiring the building/unit/floor words to be *adjacent to a number*
     matters — a bare substring search for "building"/"industrial" also
     matches inside ordinary company names ("AL GHAITH BUILDING
     CONSTRUCTION", "GULF INDUSTRIAL SERVICES"), which briefly misfiled
     those company names as an address. And the bare-phone fallback (a
     number with no preceding "Tel:"/"Mob:" label at all) needs to require
     phone-like formatting (`_BARE_PHONE_RE`: a leading "+" or an internal
     separator) — without that, a solid 15-digit TRN/reference number with
     no separators matched the same loose digit-string pattern as a real
     phone number and got misfiled as one.
  - Both the initial jammed-data batch and these two follow-up fixes
    required the same "clear the 7ish contact fields via the normal save
    API, then re-run Import from documents" cycle (never a raw file
    overwrite) to actually correct the already-saved real client records,
    not just the extraction code for future imports.

## Client database storage moved to an external spreadsheet (Settings > Company Profiles > Clients Spreadsheet)

Same philosophy as every document folder in this app now: the user points
`clients_file` at a real `.xlsx` on disk (native file picker via
`/api/browse` with `field=clients_file`, branching to
`filedialog.asksaveasfilename` instead of `askdirectory` — lets the user
either pick an existing spreadsheet or type a brand-new filename to create
one), and that file becomes the actual database — not something living
inside the app.

- **`load_clients`/`save_clients` (`app.py`) branch on `clients_file_path()`**:
  if set, read/write via `engine.read_clients_workbook()` /
  `engine.build_clients_workbook()` directly against that path; if blank,
  fall back to the original `clients/<BRAND>.json` (keeps the feature
  usable out of the box for a brand that hasn't configured a location yet,
  and doubles as the migration source — see below). Every other route
  (`/api/clients`, `-import`, `-export`, `-delete`) already went through
  these two functions, so nothing else needed to change.
- **One workbook layout serves three jobs now**: the Settings "Export All
  Client Profiles" button, the Clients tab's per-card "Export" button, AND
  the live database file itself, all via the same
  `engine.build_clients_workbook()`/`CLIENT_SHEET_HEADERS`. Column A ("ID")
  is a real column — needed so the app can recognize "this row is the same
  client as before" across saves even if a human reorders rows or edits
  other columns by hand in Excel — but `ws.column_dimensions["A"].hidden =
  True`, so a person opening the file just sees a normal sheet starting at
  Company Name.
- **`engine.read_clients_workbook(path)`** is the reverse of
  `build_clients_workbook` — reads the ID/name/etc. columns back into client
  dicts, and reconstructs each row's logo from the embedded image via
  `openpyxl`'s `ws._images`, matching `img.anchor._from.row` (0-indexed) to
  the 1-indexed data row, then `img._data()` + `img.format` to rebuild the
  base64 data URL. Wrapped in a broad `except Exception -> []` — this reads
  a file that could be open in Excel (locked), hand-edited into something
  malformed, or just not there yet, and one bad file shouldn't crash the
  Clients tab.
- **Country round-trips through a name, not a bare code** — `_country_display()`
  writes the full country name into the cell (e.g. "United Arab Emirates"),
  since this is a real file a human opens directly and a bare "AE" is less
  legible than the picker's flag+name is on screen. `_country_code_from_display()`
  reverses it via `_COUNTRY_NAME_TO_CODE` (a name→code map built from the
  same `COUNTRY_NAMES` dict used for display) — falls back to keeping
  whatever raw text is there if it doesn't match a known name, since a user
  might hand-type something non-standard directly into the spreadsheet and
  that should still be usable as free text even though it won't resolve to
  a flag in the picker. `COUNTRY_NAMES` mirrors the client-side `COUNTRIES`
  list in app.py (kept in sync manually — no shared-data mechanism between
  the embedded JS and Python in this single-file-frontend app).
- **First-time migration, never silent data loss**:
  `_maybe_migrate_clients_to_file(brand, old_path, new_path)` runs from both
  `/api/browse` (picking via the native dialog) and `/api/settings` POST
  (pasting a path manually) — if the newly-set path doesn't already exist,
  it seeds the new file from whatever's in the legacy JSON store, so
  switching to file-based storage doesn't look like the client list
  vanished. If the target file *already exists* (the user pointed at a
  spreadsheet they already had, e.g. from a previous Export), it's left
  untouched — that file is authoritative now, not the JSON. The legacy
  JSON itself is never deleted, just stops being read once `clients_file`
  is set, so it doubles as an incidental backup.
- Verified end-to-end against the real 67-client Artemis database in a
  scratch location (never a real F:\ path, per the standing rule about
  this app running on real production data): migration created the file
  correctly, all 67 rows read back byte-for-byte (including a temporary
  test logo image round-tripping through `_insert_photo`/`read_clients_workbook`
  correctly), a new client create + delete round-tripped through the xlsx
  backend, then `clients_file` was reverted to blank and confirmed the app
  fell back to the untouched original JSON with all 67 clients still intact.

## Small fixes bundled with the above

- **Color picker now shows a 🎨 icon** instead of a plain filled swatch
  (which just looked like a black square with the default color) —
  `.rtcolorwrap` overlays the real `<input type=color>` transparently on
  top of a visible palette-emoji `<span>`, so it's still a real native
  color input functionally (click opens the OS color dialog) but reads as
  "this is the color tool" at a glance, the way Photoshop/Illustrator-style
  toolbars use a recognizable icon rather than a raw color swatch.
- **Preview pan tool couldn't reach the left edge when zoomed in** — classic
  CSS gotcha: `.pvpages` used `justify-content:center`, and browsers don't
  let you scroll into a *centered* flex container's start-side overflow
  (only the end side), so once zoomed past the container width the hand
  tool could pan right to reveal the right edge but never left far enough
  to reach the true left edge. Fixed with `justify-content:safe center` —
  centers normally when everything fits, but falls back to reachable
  start-alignment once content overflows (the CSS Box Alignment spec's
  `safe`/`unsafe` keywords exist specifically for this).
- **Status segmented control now has its own white card**, matching every
  other section on the Build tab (Header, Client/Company, Line items,
  etc.) — it was previously a bare `.seg` sitting directly on the page
  background. `id=qtn2-statusseg-top` moved from the inner `.seg` div to
  the new outer `.card` wrapper (so `setType()`'s hide/show toggle still
  hides the whole card, not just the buttons inside it).
- **Clients tab grouping can now split by City as well as Section**
  (`#clientgroupseg`, `setClientGrouping()`) — `renderClientsGrid()`
  generalized to group by either `c.category` or `c.city` depending on
  `CLIENT_GROUPING`, sorted with blank/unset last under an "Unspecified"
  header (mirrors how "Uncategorized" sorts last for Section grouping).
  Originally shipped as Section/**Country**, but swapped to City after real
  use showed Country grouping wasn't useful here — nearly every client is
  UAE-based, so it just produced one giant "United Arab Emirates" bucket
  and an "Unspecified" bucket; City (Dubai vs. Abu Dhabi vs. Sharjah, etc.)
  is the dimension that's actually distinguishing for this business.

## "+ New" escape hatch and Drafts (save in-progress work on any doc type)

- **`startNewDocument()` (`app.py`)** — a `+ New` button next to the type
  tabs, always visible regardless of what's currently loaded. Fixes a real
  navigation trap: `openDoc()` (the All Docs "Open in CS" action) switches
  Build to whatever type/data the opened document has, but there was no
  obvious way back to a blank form afterward — the escape hatch (clicking
  the already-active type tab, which `setType()` already resets on
  non-silent calls) existed but wasn't discoverable. `startNewDocument()`
  is just `setType(TYPE)` — re-running the existing non-silent reset
  explicitly, as a clearly-labeled action instead of a hidden behavior.
- **Drafts**: save whatever's filled in on *any* doc type (QTN/INV/DO/QTN2)
  before hitting Generate, and resume it later — distinct from `EDITING`
  (which tracks overwriting an already-*generated* file) via a separate
  `EDITING_DRAFT` id. Storage is `drafts/<BRAND>.json` (simple internal
  JSON, unlike Clients — drafts are transient working state a user doesn't
  need to open in Excel, so the "keep it outside the system" treatment
  didn't apply here) via `load_drafts`/`save_drafts` and
  `GET/POST /api/drafts`, `POST /api/drafts-delete`.
  - `collectDocData()` (`app.py`) is the single source of "what a
    document's form state is" — extracted out of `generate()` so both it
    and `saveDraftFromForm()` build the exact same payload shape
    (header fields, items, discount/vat, QTN2's attn/address/status),
    instead of maintaining two copies that could drift apart.
  - **UI**: a `Drafts` button (next to `+ New`) opens a popover — same
    `position:fixed` + `getBoundingClientRect()` popover pattern as the
    client/country pickers, reusing their `.clientpicker`/`.cplist`/`.cpitem`
    CSS classes directly rather than defining new ones — listing saved
    drafts (sorted newest-first) with a per-row delete `✕` and a
    "+ Save current as draft" button that calls `saveDraftFromForm()`.
  - `loadDraft(id)` mirrors `openDoc()`'s field-population logic (`setType`,
    `setFieldVal` over `HEAD[type]`, items, QTN2 extras) plus a new
    `restoreDiscVat(discount, vat)` — the inverse of `collectDiscVat()`,
    reconstructing the Discount & VAT card's checkbox/mode-seg/value-input
    state from the saved `{enabled, mode, value}` shape (target mode sets
    `discMode='target'` with a plain number; percent/fixed both live under
    `discMode='amount'`, differing only by whether the value string has a
    trailing `%`).
  - **Lifecycle**: loading a draft sets the page title to "Draft: …" and
    leaves `EDITING` (the overwrite-on-Generate target) at `null` — a draft
    isn't a real file yet. Once `generate()` succeeds, if `EDITING_DRAFT`
    was set it auto-deletes that draft (its job is done, it's now a real
    document) via the same `deleteDraft()` the picker's `✕` button uses,
    just called without a click event. `EDITING_DRAFT` is also cleared
    (without deleting anything) by `startNewDocument()` and `openDoc()`, so
    abandoning a loaded draft to start something else doesn't leave stale
    state around.
  - Verified end-to-end: save → reload page (clearing all JS state) →
    reopen via the picker → every field including a `target`-mode discount
    came back exactly as saved; delete removes it from `drafts/ARTEMIS.json`
    immediately.

## Product Builder — compose an item description from its datasheet's own option tables

Instead of typing a product code from memory, matching a line item's
description to a datasheet (already-existing feature — see
`match_datasheets`) now also unlocks a "🛠 Build" button that reads that
datasheet's own spec tables and lets the user pick Wattage/CCT/Beam Angle/
Size/Finish Color/Controls from dropdowns, then composes both a
human-readable description line and a `CODE:` line, ready to insert.

- **Why this was even possible**: sampled ~5 real datasheets across
  different product categories (downlights, wall washers, spotlights)
  *before* writing any regex, and found the whole catalog follows one real
  convention — a "Technical Specifications" block with labeled option
  groups, and an "Ordering Code Example: SLAQU-5W-D2-30-12-22-ND-WH" line
  that hands you the model-code prefix directly. This is what makes
  `engine.extract_product_options()` (in `engine.py`, next to
  `match_datasheets`) more than a shot in the dark.
- **How the regexes work**: PyMuPDF's plain-text extraction from a PDF
  table collapses all column structure, but turns each cell into its own
  line — so a spec table's "code / value" pairs (e.g. `"30\n3000K"` for
  CCT, `"WH\nWhite"` for color, `"ND\nNon-dimmable"` for controls) reliably
  end up as two *adjacent* lines. Every pair-extracting regex
  (`_CCT_PAIR_RE`, `_BEAM_PAIR_RE`, `_CONTROL_PAIR_RE`, `_COLOR_PAIR_RE`)
  anchors on exactly that adjacency with `re.M` + `^...$\n^...$`, rather
  than trying to parse real table structure (which the extracted text
  doesn't have anymore). The model prefix and `order_code_example` come
  straight from splitting the "Ordering Code Example" line on its first
  `-`, since that line is reliably present verbatim.
- **This is heuristic, not a real PDF-table parser** — same "best-effort,
  never guess a category that isn't confidently found" philosophy as
  `extract_legacy_contact`. One of the 5 sampled datasheets (DONI.pdf) had
  its first Technical-Specifications section extract as garbled
  single-character lines (an embedded-font quirk in that specific source
  file) and came back with no `wattage`/`ip_rating` — but its CCT/Beam
  Angle/Controls/Color/Size/model/order-code, which live further down in a
  cleanly-extracted section of the same file, all still worked. The UI
  just omits a dropdown for whatever category wasn't found (`pbSelectRow`
  returns `''` for an empty/missing list) rather than erroring — the user
  can still type the description by hand, or open the datasheet itself
  (the existing 📄 pill, unchanged) to read it visually.
- **Size is a flat list, not mapped to a specific wattage** — real
  datasheets group multiple wattages under one cutout/dimension (e.g.
  AQUA's 5-7W share a `90×70mm` housing, 18-24W share `190×75mm`), and
  PyMuPDF's linear text doesn't preserve which table row a given size
  belonged to reliably enough to auto-pair them. `_SIZE_RE` just collects
  every distinct `NNN×NNNmm`-shaped token found anywhere in the document;
  the user cross-checks against the datasheet (still one click away) if
  the precise wattage↔size pairing matters for a given order.
- **Route**: `GET /api/product-options?rel=<datasheet path>` (reuses
  `_resolve_datasheet_pdf`, same resolution/validation as `/datasheet` and
  `/datasheet-thumb`) → `engine.extract_product_options(pdf_path)`.
- **UI** (`#productbuilder` modal, reuses `.clientmodal`/`.clientmodalbox`/
  `.clientmodalbar`/`.clientmodalbody` chrome — no new modal CSS needed):
  `openProductBuilder(itemIndex, rel, name)` fetches the options and calls
  `renderProductBuilderBody()`, which builds one `<select>` per category
  that came back non-empty via `pbSelectRow()`. `updateProductBuilderPreview()`
  recomputes `composeProductBuilderText()` (model + wattage + CCT + beam +
  IP + color on the first line, `CODE:` line built from whichever
  code-halves of the pairs are available) on every dropdown change.
  "Insert into Description" writes straight into `items[i].description`.
  **The QTN2-vs-textarea distinction already established elsewhere in this
  file applies here too**: QTN2's description cell is a contenteditable
  richbox (needs real `<br>` tags for line breaks), every other type is a
  plain `<textarea>` (a literal `\n` already renders as a line break there)
  — `insertProductBuilderDescription()` branches on `TYPE==='QTN2'` to
  `escHtml(text).replace(/\n/g,'<br>')` vs. the raw text, exactly like
  `collectQtn2Extra()`'s composed-address handling.
- **The "Build" button sits right next to the "📄 <name>" preview pill,
  inside that item's own card** (`itemCardHtml()` builds both from
  `it.datasheets` directly — see the card-layout section below; there's no
  separate Datasheets card/panel anymore, that got folded into each item).
- Verified end-to-end against real catalog files (AQUA.pdf): matched →
  built (12W/3000K/38°/White) → inserted → correctly appeared in the live
  QTN2 PDF preview with a real `<br>` between the description and CODE
  line; separately confirmed the plain-textarea (INV) path produces a
  literal `\n` instead, as intended.
- **Product Builder dropdowns now have a "Custom…" option** (`onPbSelectChange`),
  same interaction pattern as the existing Unit dropdown's Custom entry —
  swaps that one `<select>` for a text `<input>` in place (never
  `window.prompt()`, which is documented elsewhere in this file as a
  silent no-op in this environment). Picked values now come in three
  possible shapes (bare string like `"12W"`, a `[code, label]` pair like
  `["30", "3000K"]`, or `{custom: "..."}`), so `pbSelected()`/`pbLabel()`/
  `pbCode()` normalize all three instead of `composeProductBuilderText()`
  needing to know which one it's holding.
  - **Going custom used to be a one-way trip** (caught immediately after
    shipping, by the user) — the input permanently replaced the select
    with no way back to the option list for that category. Fixed with a
    `PB_ROWS` registry (`id -> {label, list, fmt}`, populated by
    `pbSelectRow` as it builds each row) so `revertPbCustom(id)` can
    rebuild *just that one row's* `<select>` from scratch (`pbSelectInner(id)`)
    without touching any other field's current selection — the naive fix
    (re-render the whole builder body) would have reset every other
    dropdown back to its first option too, silently discarding whatever
    else the user had already picked. The custom `<input>` gets a small
    "▾" button next to it for this.
- **A "Find Product" step now exists before "Build"** — see the Line Items
  card-layout section below for the full picture; the short version is
  that Item Description's job changed from "type anything, hope it
  auto-matches a datasheet" to "search the catalog on purpose, then
  configure what you found," with manual typing kept as a fallback for
  anything not in the catalog.
- **Build vs. Edit is state-aware, not a fixed label**: `isProductBuilt(desc)`
  (a one-line `/CODE:/i` test — matches both the app's own composed
  descriptions and the pre-existing real convention historical descriptions
  already used, e.g. `"...CODE:SLCEL-10W-30"`) decides whether a datasheet's
  pill reads "🛠 Build — Name" (nothing configured yet) or "✏️ Edit — Name"
  (re-opening what's already there) — same check drives both the item
  card's pill label and the Product Builder modal's own title, so they
  never disagree. The 📄-preview action moved to its own small icon-only
  button (`.dspillicon`) right after Build/Edit, rather than sitting
  between "Find Product" and "Build" as two visually-equal buttons — Find
  and Build/Edit are the two actions that matter and read as a pair now;
  preview is secondary.
- **More breathing room between item cards**: `.itemslist`'s `gap` (10px →
  18px) and the list's `margin-bottom` before the "+ Add Line Item" bar —
  the cards' own borders were getting lost against each other at the
  tighter spacing.

## Build tab: PO Box/City/Country, preview toolbar order, line items layout

- **Preview toolbar reordered**: `Fit` moved next to the hand-tool button
  (start of the toolbar), ahead of Single/Double — grouping the two "reset/
  control how you're viewing this" actions together, separate from the
  zoom level controls (−/100%/+) that follow Single/Double. Pure markup
  reorder, no behavior change.
- **Line Items rewritten from a `<table>` to a stack of item cards**
  (`renderItems()`/`itemCardHtml()`, `#itemslist` instead of `<table
  class=it>`/`#itbody`/`#ithead` — `renderItemHead()` is gone entirely,
  there's no header row to build since every field now carries its own
  inline `<label>`). Each `.itemcard` is: a compact top row (photo
  thumbnail, Type input if the doc type has one, Remove `×` — all
  small/inline), a "find" row (🔍 Find Product button + any matched
  datasheet's 📄/🛠 pills), the description itself (full-width, richbox
  for QTN2 / textarea otherwise — the same distinction as before, just
  wider and less cramped now that it isn't squeezed into one column of a
  7-column table), and a bottom row of whatever non-photo/description/type
  fields that doc type actually has (Unit + Qty + Price for QTN/QTN2/INV;
  Unit + LPO Qty + Prev. Delivery + Delivered for DO) via `COLS[TYPE]`,
  same generic-over-doc-type approach the old table used.
  - **"+ Add Line Item" is a full-width dashed bar** below the card stack
    (`.additembar`), not a small button — closer to how Word/Notion-style
    "add a row" affordances read as an obvious, easy target rather than a
    small button tucked in a card header.
  - **The Datasheets card is gone** — its pills (📄 preview / 🛠 Build) now
    render directly inside the item card they belong to (`it.datasheets`
    read straight off each item in `itemCardHtml`), so there's no separate
    section to scroll to just to find the button for the item you're
    already looking at.
  - **"Find Product" (`openProductFinder`/`pickProduct`) is the new front
    door**: a `.clientpicker`-style popover (search + list, reusing
    `.cplist`/`.cpitem`, same as the client/country/draft pickers) listing
    every product in the datasheet catalog (`GET /api/product-list` →
    `engine.list_datasheet_products()`, a flat name-sorted dedupe over the
    same catalog index `match_datasheets` already builds). Picking one:
    attaches that datasheet to the item, fills the description with just
    the product name **only if the description is still blank** (never
    clobbers text someone already typed), and immediately opens the
    Product Builder for it — "find, then build" as one continuous action,
    not two things the user has to remember to do separately. Typing a
    description by hand and letting it auto-match (the original mechanism)
    still works unchanged for anything not worth browsing the catalog for.
- **Build's Client/Company card now has the same PO Box/City/Country split
  as the Clients editor** (QTN2 only, same restriction as Attn/Address
  already had) — `#customer_pobox`/`#customer_city` plain inputs plus a
  `#customer-countrybtn` Country picker, deliberately **without** the
  Google Maps location block (not relevant to what gets printed on a
  document, unlike the Clients tab where it's for finding/sharing an
  address).
  - **Country picker generalized to multiple targets** instead of being
    hardcoded to the Client editor's `cf-*` fields — `countryPickerTarget`
    (a module-level variable, same pattern as `richFocused` for the rich-
    text toolbar or `countryPickerTarget` itself) tracks which caller
    opened it. `openCountryPicker(ev, target)` takes an optional second
    arg (`'customer'` from Build, defaults to `'cf'` from the Client
    editor); `setCountryValue(prefix, code)` and `pickCountry(code)` route
    through it. Every target's fields must follow `<prefix>-country`/
    `-countryflag`/`-countrylabel`/`-countrybtn` naming (note the Build
    tab's other customer fields use underscores — `customer_pobox`,
    `customer_attn` — but the *country* ones specifically use hyphens,
    `customer-country`, to fit this shared prefix+suffix scheme).
  - **The PDF only has one Address slot, so PO Box/City/Country are
    composed into `customer_address` at collection time**, not sent as
    separate template fields — `composedCustomerAddress()` appends
    `PO Box X, City, Country Name` as a second line (`<br>`-separated) onto
    whatever's in the Address richbox, only when at least one of the three
    is filled in. **Critical gotcha this created**: the composed string
    must never be fed back into the Address box itself (e.g. when
    restoring a saved draft) — doing so and then re-composing again would
    print the PO Box/City/Country line twice. `collectQtn2Extra()` returns
    *both* `customer_address` (composed, sent to the PDF renderer) and
    `customer_address_raw` (just the box's own text, used only to restore
    the Address field itself in `loadDraft()`) to keep these separate.
  - **Every existing Client↔Build sync point updated to carry the three
    new fields alongside attn/address**: `setBuildAddressExtras(c)` is the
    shared helper (sets `customer_pobox`/`customer_city`/country from a
    client record) called by `fillFromClientIfBlank` (blank-check gated,
    like attn/address), `useClientInBuild`, and `applyClientPick`
    (unconditional, like attn/address there). `saveCurrentAsClient` now
    also reads them from the Build fields when `TYPE==='QTN2'`, falling
    back to the existing client's values otherwise — same "never silently
    wipe a field Build doesn't edit" reasoning already documented for
    category/landline/website/trn.
  - Verified end-to-end: composed address renders correctly in the live
    PDF preview; save-as-draft → reload confirmed the raw/composed split
    prevents double-printing the PO Box/City/Country line; Client editor's
    own country picker (`cf` target) still works unaffected by the
    generalization.

## Rich text on QTN2's free-text fields

Company, Attn, Address, Project, Area, and each line item's Item Description
(QTN2 / new Quotation only — not QTN/INV/DO, which still use plain
`<input>`/`<textarea>` since the xlsx pipeline has no equivalent rich-text
support) support Bold/Italic/Underline/Strikethrough/Color via a shared
floating toolbar (`#richtoolbar`, `richCmd()` in `app.py`'s JS).

- **Why contenteditable, not `<textarea>`**: only a `contenteditable` element
  can hold inline formatting tags (`<b>`, `<i>`, `<u>`, `<strike>`,
  `<font color=...>`). These fields are `.richbox` divs; `document.execCommand`
  (still functional in Chromium despite being deprecated) applies the actual
  formatting. `richText(el)` reads `el.innerHTML` back out (normalizing an
  empty contenteditable's stray `<br>` to `''`).
- **One shared toolbar, not one per box**: `richFocused` tracks whichever
  richbox last had focus; toolbar buttons use `onmousedown="event.preventDefault()"`
  so clicking them doesn't blur the richbox first. The color `<input
  type=color>` is the one exception — it's a native popup that *does* steal
  focus no matter what, so `selectionchange` continuously snapshots the live
  selection into `richRange` while a richbox is focused, and `richCmd`
  restores that saved range before running `execCommand` even if focus
  bounced away to the color picker and back.
  - **Real bug this caused**: the `focusout` handler that hides the toolbar
    only checked `document.activeElement !== richFocused` — so the instant
    focus moved to the color `<input>` (which lives *inside* the toolbar),
    the toolbar hid itself out from under the color picker the user was
    about to use. Fixed by also checking
    `!$('richtoolbar').contains(document.activeElement)`, and by listening
    for `focusout` on the toolbar itself (not just the richbox) so the
    toolbar still hides correctly once focus later leaves the color input
    for something outside both. The color input also uses `onchange`
    (fires once, when picking finishes) rather than `oninput` (fires
    continuously while dragging), since `execCommand`, run repeatedly
    against a range that a prior call already mutated, is asking for
    trouble.
  - **Recently-used color swatches** (`RECENT_COLORS`, `renderRichSwatches`/
    `pinRichColor` — `#richswatches` in the toolbar): one-click reuse for
    colors you've already picked, so you don't have to reopen the native
    picker for the same color twice. Seeded with 5 sensible defaults on a
    fresh browser (ink/amber/red/blue/green); picking any new color via
    `#richcolor`'s `onchange` pins it to the front (deduped, capped at 6),
    persisted in `localStorage['cs_recentColors']` so it survives reloads.
- **Native `<datalist>` doesn't work on `contenteditable`** — only on
  `<input>`. Converting Company/Attn/Address/Project to richboxes would have
  silently killed the autocomplete-prediction feature built the session
  before this one. Fixed with a small hand-rolled `attachAutocomplete(el,
  listGetter)` (filtered popup, click/arrow-keys/Enter to pick) reusing the
  same `CLIENTS`/`PROJECTS`/`ATTNS`/`ADDRESSES` arrays `loadClients()` already
  fetches. QTN/INV/DO's plain Company/Project inputs still use the real
  `<datalist>` (`#clients`/`#projects`) — both mechanisms coexist.
- **Company is two elements, not one**: `#company` (plain input, used by
  QTN/INV/DO) and `#company-rich` (richbox, used by QTN2) sit side by side;
  `setType()` toggles which is visible via `.hide`, and `companyVal()`/
  `setCompanyVal()` read/write whichever is active. Project/Area and Item
  Description don't need this duplication since `renderHead()`/`renderItems()`
  already regenerate their markup fresh per type — they just emit a richbox
  instead of a plain input/textarea when `TYPE==='QTN2'`.
- **`engine._slug()` strips HTML tags first, not just `<`/`>`** — a real bug
  caught during testing: stripping only angle brackets left the tag *name*
  behind as ordinary text (`"<b>Foo</b>"` slugged to `"bFoo-b"`, since `b` is
  alphanumeric and passed the whitelist filter once the brackets were gone).
  Any future field that can carry rich HTML and also feeds a filename must
  go through `_slug()` (or something that strips `<[^>]*>` first) — never
  assume angle-bracket-stripping alone is enough.
- Jinja2's `Environment` here has `autoescape=False` (the default) — rich
  HTML renders as real markup in the PDF, not escaped text. That's
  intentional and required for this feature; don't "fix" it by turning
  autoescape on without also marking these fields `|safe`.

## Pixel-fidelity HTML document pipeline (QTN2, and future doc types)

A professional design handoff ("Company Templates Rebuild") specced 10 new
business documents as high-fidelity HTML/CSS mockups (Google-Fonts
typography, gradients, pill badges, `border-radius`) that Excel/openpyxl
cannot reproduce. Rather than reskin the xlsx pipeline, these are recreated
as **real HTML rendered to PDF via headless Chromium (Playwright)** — a
second pipeline living alongside the xlsx one, registered per doc type via
`engine.HTML_DOC_TYPES`. Only **Quotation → `QTN2`** is built so far; the
other 9 (Proforma Invoice, Tax Invoice, Credit Note, Payment Receipt,
Statement of Account, Delivery Note, Purchase Order, Material Request, Job
Completion Certificate) are a deliberate follow-up, not yet started. The
original design-pack files (for reference when doing the next one) aren't in
the repo — they came from `Company Templates Rebuild.zip`, unzipped to a
scratch temp dir during that session; ask the user for the zip again if it's
needed and no longer around.

**`QTN2` is now *the* Quotation** — the old xlsx-based Quotation creation UI
(`#typeseg`'s old "Quotation"/`QTN` button) has been removed from Build, and
`QTN2` is the default type on load (`setType('QTN2')` in the init line).
`QTN`'s code, template, and folder setting are all still fully in place and
untouched — deliberately *disconnected from the create-new UI, not
deleted* — so:
- **All Docs still fully browses/previews/opens the thousands of existing
  legacy Excel quotations** (`qtn_folder`, relabeled "Quotations (Legacy
  Excel)" in Settings to distinguish it from the new `qtn2_folder`
  "Quotations (PDF)").
- **"Open in CS" on a legacy QTN file still works** even with no `QTN`
  button in `#typeseg` — `setType('QTN', true)` doesn't require a matching
  tab button to exist, it just drives `HEAD`/`COLS`/`LABEL` lookups by key,
  so `openDoc()` sets `TYPE='QTN'` and the form renders correctly with no
  tab visibly highlighted (a deliberate, acceptable quirk — you're in an
  editing mode that isn't one of the three create-new tabs).
- If old-style Quotation creation ever needs to come back (or needs
  removing for real), the button markup is the only thing that was deleted;
  everything else (`engine.fill_quotation`, `templates/<BRAND>/QTN.xlsx`,
  `HEAD.QTN`/`COLS.QTN`/`LABEL.QTN`, `FOLDER_KEYS.QTN`) is intact.

- **PDF-only, no xlsx**: `QTN2` (and any future doc in `HTML_DOC_TYPES`) has
  no spreadsheet twin — `/api/generate` and `/api/preview-draft` in `app.py`
  branch on `doc_type in engine.HTML_DOC_TYPES` to call
  `html_engine.RENDERERS[doc_type](...)` instead of
  `engine.FILLERS`/`engine.to_pdf`. The resulting PDF still feeds the *same*
  `engine.to_png_pages` call as the xlsx path, so the multi-page live
  preview (zoom/pan/double-page/resizer) needed **zero** changes — it's
  generic over "a PDF landed in `_cs_draft/`", regardless of which engine
  produced it.
- **Template porting pattern**: the design pack's own `<x-dc>`/`support.js`/
  `<sc-for>`/`<sc-if>` scaffolding is the *design tool's* preview harness —
  explicitly not for production (its own README says so). What *is* reused
  verbatim is `doc-page.js` (copied into `static/doc_html/`): a
  framework-free custom element that repeats a header/footer on every
  printed page via `@page{margin:0}` + a table with repeating
  `<thead>/<tfoot>` spacers. That trick works unmodified under headless
  Chromium, so `templates_html/quotation.html` still uses a real
  `<doc-page size=a4 margin=13mm>` element — only the design tool's own
  templating layer was swapped for Jinja2 (`{{ field }}`/`{% for %}` map
  almost 1:1 onto the mockup's own `{{ }}`/`<sc-for>` syntax already).
- **Playwright + Flask threading gotcha**: Playwright's *sync* API is pinned
  to the thread that started it. A tempting "launch once, reuse the
  browser" module-level singleton (mirroring `engine._find_soffice()`)
  breaks with `cannot switch to a different thread` the moment Flask's dev
  server handles two requests on different threads. `html_engine.render_pdf`
  instead runs the whole `sync_playwright()` start→render→close lifecycle
  inside one call — costs a Chromium launch per render (~hundreds of ms) but
  is correct regardless of threading. Don't reintroduce a cached
  browser/playwright instance without solving that properly first.
- **Company masthead info** (`html_engine.COMPANY_INFO`) is hardcoded for
  Artemis only for now, matching the design pack (name/P.O. box/phone/TRN) —
  unlike the xlsx pipeline's per-brand template files, there's no per-brand
  config UI for this yet. Fine while only Artemis uses `QTN2`; revisit if
  another brand needs its own HTML-pipeline documents.
- **Discount & VAT math is duplicated, deliberately**: `html_engine.compute_totals`
  reimplements the same percent/fixed/target-price rules as the xlsx path's
  `engine._write_summary_block` (see below), in plain Python since there's no
  spreadsheet formula engine here. If the discount/VAT rules ever change,
  both places need updating.
- **Fonts load from Google Fonts at render time** (this machine has normal
  internet access) — not self-hosted. If offline reliability ever matters,
  download Archivo/IBM Plex Sans once and swap the `<link>` tags in
  `templates_html/*.html` for local `@font-face` rules.
- **No round-trip editing yet**: reopening a saved `QTN2` PDF via "Open in
  CS" only offers the generic page-1 preview (`editable` is correctly
  `False` since `QTN2` isn't in `engine.FILLERS`) — there's no form
  repopulation like xlsx docs get. Deliberately out of scope for the first
  pass.
- **JSON sidecar** (`engine.save_sidecar`/`read_sidecar_items`,
  `<stem>.json` next to the `.pdf`): `HTML_DOC_TYPES` have no `.xlsx` to read
  line items back out of the way `read_items_from_doc` does for QTN/INV, so
  `/api/generate` saves the exact generation payload as a small JSON file
  alongside the PDF. Currently only powers "previous quotations for this
  company" (`previous_for_company` matches on `.pdf` instead of `.xlsx` for
  these types) — it's not read anywhere else yet, but it's there if
  round-trip editing gets built later.
- Numbers rendered into the template go through `html_engine.num_display`
  (qty — whole numbers show as `20`, not `20.0`) and `money` (2dp,
  thousands-separated) as appropriate; dates go through a local `_fmt_date`
  wrapping `engine._fmt_date` (`DD.MM.YYYY`, matching the xlsx pipeline) —
  don't pass raw form values straight into the template.
- Filename convention regexes (`engine._FN_RE`/`_FN_RE_BRANDED`) had to be
  loosened from `[A-Za-z]+` to `[A-Za-z]+\d*` for the type token to allow
  `QTN2` — worth remembering if another numbered type code gets added.

## Known gaps / not yet done

- ADS and Watt Electricals have no folders or templates configured.
- `templates_folder` (per-brand external template override) exists in
  config/code but hasn't been exercised with a real external folder yet.
- The F: drive (where Sololuce's catalogs live) has intermittently dropped
  its connection during past sessions — `os.walk`-based scanning already
  degrades gracefully (skips unreachable subfolders) rather than crashing.
- 9 of the 10 design-pack documents (everything except Quotation/`QTN2`)
  haven't been built yet — see "Pixel-fidelity HTML document pipeline" above.
- Client logos are stored full-size as uploaded (no client-side resize/
  compression before base64-encoding) — fine for the expected scale, but a
  user uploading several large, high-res logos could bloat
  `clients/<BRAND>.json` noticeably. Worth adding a canvas-based downscale
  on upload if that becomes a real problem.

## Line-item card gap, Clients Country/City grouping, "+ New" dead-end fixes

- **The item-card gap CSS was correct but never applied**: the previous
  pass added `.itemslist{gap:18px;margin-bottom:18px}` but the actual
  `<div id=itemslist>` markup never had `class=itemslist` on it — so none
  of that spacing rendered, and the cards visually ran together with no
  breathing room before the "+ Add Line Item" bar. Fixed by adding the
  missing class to the element. Lesson: when a CSS-only "fix" doesn't show
  up live, check that the selector is actually attached to the markup
  before assuming the rule itself is wrong.
- **Clients grouping "Section" replaced with "Country"**: the toggle used
  to group by `c.category` (a free-text "Section" field like Contractor/
  Consultant/Government) or `c.city`. Per the user's request this became
  **Country** (grouping on the existing `c.country` field, already
  populated via the Country picker) and **City**, side by side — the old
  `category` grouping mode is gone (the `cf-category`/"Section" field on
  the client editor itself is untouched, it's just no longer one of the
  two grouping tabs).
- **"+ New" was a dead end when editing a legacy `QTN` file**: opening an
  existing legacy-Excel quotation for edit/import (`openDoc()`) sets
  `TYPE='QTN'`, but the visible type segmented control only has buttons
  for `QTN2`/`INV`/`DO` — `QTN` (the old xlsx pipeline, superseded by QTN2
  for new documents but still openable from All Docs) has no tab of its
  own. That left the control showing *no* active tab, and — the actual
  bug — pressing "+ New" called `setType(TYPE)`, i.e. `setType('QTN')`,
  which silently rebuilt another blank **legacy QTN** form: still no tab
  lit up, so the user had no way to tell anything had happened and no
  path back to a normal, visibly-selected document type. Fixed with a
  `VISIBLE_TYPES=['QTN2','INV','DO']` allowlist: `startNewDocument()` now
  falls back to `QTN2` whenever the current `TYPE` isn't one of the three
  tabs the user can actually click, so "+ New" always lands on a real,
  highlighted tab instead of silently re-entering the same orphaned state.

## "Open in CS" is now its own full-window editing mode, not the Build tab

The `startNewDocument()`/`VISIBLE_TYPES` fix above patched the symptom (no
way back out), but the user then asked for the real fix: editing an
existing document shouldn't share the Build tab/state at all — it should
feel like its own dedicated mode, "similar to Build option, just more
detailed editing," as a separate window.

- **`#buildwrap`** is the existing Build `.wrap` (type tabs, header,
  client/company, line items, discount/VAT, Generate button, and the live
  preview pane) — now given an id so it can be *moved* between two hosts
  rather than duplicated. Vanilla single-file inline-script apps like this
  one can't easily run two independent copies of the same form (every
  field is a hardcoded global id like `#company`/`#itemslist`), so
  re-parenting the one real copy of the DOM subtree is the trick that
  gets a genuinely separate mode without rewriting the whole form as a
  parameterized component.
- **`#editmodal`** is a new full-screen overlay (`position:fixed;inset:0`,
  same pattern as `#csmodal`/`#productbuilder`) with its own bar (a plain
  "Editing/Import — <label>" title + a single Close button) and an empty
  `#editmodalbody` slot.
- **`enterDocEditMode(label)`**: sets the modal's title, `appendChild`s
  `#buildwrap` into `#editmodalbody` (moving it out of the normal Build
  tab), hides `#modebar` (the type-tabs/Drafts/+New bar — irrelevant once
  you're editing one specific existing file rather than choosing what to
  create), and shows the modal. Crucially this does **not** call
  `view('build')` — whatever tab was showing underneath (typically All
  Docs) stays exactly as it was, so the modal is a true overlay, not a
  navigation.
- **`exitEditMode()`** (bound to the modal's Close button) reverses it:
  moves `#buildwrap` back into `#v-build`, re-shows `#modebar`, hides the
  modal, and calls `setType('QTN2')` to fully reset the form — so closing
  the editor always leaves Build in a clean, blank state rather than
  carrying over whatever was being edited.
- **`openDoc(rel)`** (the xlsx-backed "Open in CS" path — legacy `QTN`,
  `INV`, `DO`) is unchanged apart from swapping `view('build')` for
  `enterDocEditMode(label)` at the end; all the existing field-population
  logic (`HEAD[...].forEach`, `renderItems()`, `matchPhotoForItem`,
  `matchDatasheetsForItem`, `setPreviewImage`/`schedulePreview`) still
  targets the same ids, which now just happen to live inside the modal.
- **`openCS(rel)`** (the read-only thumbnail preview for PDF-only docs
  that have no xlsx to edit, e.g. `QTN2` or a `DO` without a spreadsheet)
  is untouched — that's a genuinely different, simpler case (just a
  picture of page 1) and still uses the original `#csmodal`.
- Verified against real production files (read-only — no Generate/Save
  was run against F:\ data, per the standing rule): opened an INV via
  "Open in CS" from All Docs → full-window editor appeared with no type
  tabs, header/line items populated correctly, live preview rendering;
  clicked Close → landed back on All Docs exactly as it was; switched to
  Build tab → confirmed it was a fresh, untouched "New Quotation" (QTN2
  tab active) rather than showing leftover edit state; separately
  confirmed a PDF-only DO's "Open in CS" still opens the old read-only
  `#csmodal` thumbnail unchanged.

## Edit mode: keep the real document as-is until the user actually changes something

The user's follow-up after the full-window editor above: don't touch the
document's own information or template on open — show it exactly as it
is, and only start doing anything "live" once the user actually edits a
field. Two real gaps surfaced here, both in `/api/doc`'s "best-effort
import" branch (used for any `.xlsx` that doesn't match the app's own
generated filename convention — i.e. almost every real historical file,
since they predate this app and use older/looser naming):

- **Wrong doc type, blank header, for real historical files.** The
  branch used to hardcode `doc_type = "QTN"` whenever `parse_filename`
  failed (which it does for anything not shaped like
  `TYPE_NUMBER_R{rev}_COMPANY_PROJECT_DATE.ext`), and left `number`/
  `date`/`company`/`project` all blank. In practice this meant opening a
  real legacy **Tax Invoice** like `ARTEMIS_ TAX INVOICE-10147_LTS_...xlsx`
  via "Open in CS" silently mislabeled it as a Quotation internally (wrong
  header fields shown) with every field empty — even though the All Docs
  listing, one screen over, already correctly showed it as "INV 10147"
  grouped under "LTS". That's because `scan_all`/`index_folder` (which
  power the listing) already run a generic keyword guesser
  (`engine._guess_type`, matching QTN/INV/DO/LPO/PO keywords in the
  filename) for anything `parse_filename` can't parse, but `/api/doc` had
  never been wired to the same guesser — same underlying file, two
  different, disagreeing interpretations. Fixed by reusing
  `engine._guess_type`/`_guess_number`/`_guess_date` (the exact same
  fallbacks `scan_all` already uses) plus the on-disk top-level folder
  name as the company (`rel.split("/")[0]`, matching how `scan_all`
  derives `company_label` for unmatched files) — so `/api/doc` and the
  All Docs listing now agree, and only genuinely unknowable fields
  (project, area) are left blank rather than guessed outright, keeping
  the existing "best-effort, never guess, user reviews" philosophy intact.
- **The preview switched templates before the user touched anything.**
  `openDoc()` used to call `schedulePreview()` immediately for any
  imported (non-`editable`) doc — which re-renders a brand-new preview
  from the app's own live template using whatever (often still-blank)
  header data had loaded, discarding the view of the actual original file
  before the user had done anything. Fixed by always calling
  `setPreviewImage('/cs-thumb?f=...')` on open (the real file's own
  thumbnail, exactly like the non-imported path already did) and never
  calling `schedulePreview()` proactively — the existing per-field
  `oninput`/`onchange` handlers already call `schedulePreview()` the
  moment the user changes anything, so the switch to the live template
  preview now happens exactly when it should and not before.
- Verified against a real historical Tax Invoice (`ARTEMIS_ TAX
  INVOICE-10147_LTS_...xlsx`, read-only — no Generate/Save run against
  F:\ data): opening it now shows "Import Tax Invoice 10147" (correct
  type), INV Number/Date auto-filled from the filename, Company
  auto-filled as "LTS" from the folder, and the preview pane showing the
  actual original invoice image untouched; typing into the Project field
  correctly triggered the switch to the live re-rendered template preview
  a moment later, confirming the "wait for a real edit" behavior without
  breaking normal editing.

## "Open in CS" can now edit a Delivery Order that only exists as a PDF

Follow-up request: PDF-only documents were still stuck behind the
read-only `#csmodal` thumbnail. Scoped this specifically to **Delivery
Orders** rather than all PDF-only files, for a real reason found while
investigating: the DO/QTN/INV folders (walked recursively by
`scan_all`) are full of unrelated PDFs sharing the same tree — client
drawings, LPOs, datasheets — and DOs are the one type this business
routinely keeps *only* as a PDF (kept as-is after physical delivery/
signature; the xlsx often just isn't retained). Genuine PDF-only QTN/INV
files turned out to be rare/noisy by comparison, so this stays DO-only
for now — the same door is open to extend later if needed.

- **`engine.read_do_pdf(path)`** (new): opens the PDF with PyMuPDF
  (lazy `import fitz`, same pattern as `extract_product_options`/
  `to_png_pages`), concatenates every page's text, and:
  - Reads the header via a new generic `_read_pdf_label_block(lines,
    label_keys, stop_at)` — matches "Label" lines (colon optional —
    real templates are inconsistent: some say "Rev:", others just
    "Rev") and takes every following non-empty line up to the next
    known label as the value, joined with a space. That multi-line
    join mattered in practice: one real DO's Project value line-wraps
    across 3 lines ("LA ROSA 03 & " / "04 / VILLA " / "NOVA") — an
    earlier version that only grabbed exactly 1 line after the label
    truncated it to "LA ROSA 03 &". A label whose next line is itself
    another label (or the stop marker) means that field was genuinely
    left blank on the original — correctly omitted rather than
    swallowing the next label's text.
  - Splits items by locating sequential bare-integer lines (`1`, `2`,
    `3`, …) as row starts — the same "detect the row number sequence"
    trick used nowhere else in this codebase yet, chosen because a
    PDF's linear text extraction has no cell/table structure to key
    off like the xlsx readers do. Everything between one row-number
    line and the next belongs to that item; the *last* line inside a
    row matching `<number><unit>` (e.g. "3 ROLLS", "5PCS" — spacing
    varies) is the qty+unit, everything before it is the description.
  - **Real gotcha, caught by testing a 2-page DO**: the printed
    "DELIVERY ORDER" title + company letterhead block repeats after
    the item table on *every* page (a layout artifact, not a true
    end-of-list marker) — an earlier version treated it as a stop
    marker and silently dropped every item after page 1. Fixed by only
    stopping at "Delivered By:", which really does appear exactly once
    on the true last page.
  - Returns `{}` (no text layer at all) for a scanned/signed PDF —
    there's no OCR here, same "best-effort, never guess" rule as
    everywhere else; the route below turns that into a plain-language
    error rather than silently showing an empty form.
- **`app.py`'s `_api_doc_from_pdf(rel, path, meta)`** (new): gated to
  `doc_type == "DO"` (via `parse_filename` or the same
  `engine._guess_type` keyword guesser `scan_all` uses) — anything else
  returns a 400 telling the user only DOs support this. A 422 with a
  human-readable message covers the no-text-layer case. On success,
  `imported: true` always (never overwrites the original PDF; Generate
  saves a brand-new DO xlsx+pdf in the app's own template, same as
  every other best-effort import in this app).
- **`/api/doc`** now branches on the request's own extension
  (`ext in (".xlsx", ".pdf")`) rather than requiring `.xlsx`.
- **All Docs row rendering**: `isDoPdf = !hasXlsx && hasPdf &&
  r.type==='DO'` — when true, the "Open in CS" button calls `openDoc(pdfRel)`
  (the same function used for xlsx editing) instead of falling back to
  `openCS(rel)` (the old read-only thumbnail, still used for every other
  PDF-only case).
- Verified against several real DO PDFs (read-only, no Generate/Save
  against F:\ data): a single-page one, a two-page one (confirmed all
  11 items across both pages after the trailer fix, not just page 1's
  7), one with genuinely blank Project/LPO fields (correctly left
  blank, not guessed), and one scanned/signed copy (correctly rejected
  with the "no selectable text" message rather than showing a blank
  form). Also confirmed a DO that has *both* a `.pdf` and an `.xlsx` for
  the same stem still goes through the existing xlsx path (higher
  fidelity, since xlsx cell layout is more reliable than PDF text
  extraction) — the PDF path only ever engages when there's truly no
  xlsx twin.

## Adaptive xlsx header/customer reading + real Attn/Address/PO Box/City/Country for QTN/INV/DO

Follow-up complaint against the real INV-10188 file: opening it for edit
showed INV Number/Date/Company but left QTN Number, LPO Number, and the
entire customer/company block blank — even though the real xlsx plainly
has all of that. Ask was explicit: make the reader adapt to whatever
layout a document actually uses, and give QTN/INV/DO real editable
company-detail fields (not just read-only).

- **Two real storage mechanisms discovered by unzipping actual files**:
  (1) some legacy files share the app's own cell layout almost exactly
  (just a different filename convention) — e.g. INV-10188's "QTN
  Number"/"LPO Number" labels sit at the *exact* cells (`I5`/`I6`) the
  app's own `HEADER_CELLS` already expects; (2) the "To" customer block in
  many real files isn't a cell at all — it's a floating Excel **shape**
  (a drawn textbox), invisible to `openpyxl`'s normal `ws.cell()` API,
  confirmed by unzipping a real file and finding the customer's name/
  address/TRN in `xl/drawings/drawing1.xml`, not any cell. Both needed
  their own reader.
- **`engine.read_xlsx_header_labels(ws, label_keys, max_row=40)`** (new):
  scans every cell in the top ~40 rows for text matching a known label
  (colon optional — `_norm_pdf_label`, shared with the PDF DO reader,
  handles "Rev:" vs "Rev"), takes the next non-empty cell to the right
  (skipping past it if that cell is itself another label — the field was
  left blank) or the cell below as the value. `_XLSX_HEADER_LABELS` gives
  per-type label→field maps (QTN/INV/DO each have their own set of
  recognized labels). This is the exact same "find the label, take what's
  next to it" idea as `_read_pdf_label_block`, just against a cell grid
  instead of PDF text lines — proven correct by the fact it works whether
  or not the file happens to share the app's own cell positions.
- **`engine.read_xlsx_customer_shape(path)`** (new): unzips the xlsx,
  parses every `xl/drawings/drawingN.xml` for shape anchors + text,
  finds a shape whose own text is just a "To"/"Bill To"/"Customer"/
  "Client" label, then returns the text of the nearest shape at or after
  it (same column range, same row or below) — the customer block is
  almost always the very next shape geometrically, per every real sample
  checked. Returns `""` (not a guess) if no label shape is found.
- **`/api/doc`'s best-effort branch** now merges both: `qtn_number`,
  `lpo_number`, `type`, `area`, `rev` come from the generic label scan;
  `customer_block` tries the shape scan first, falling back to the plain
  `CUSTOMER_CELL` cell value for files that do use a real cell. The
  **editable** (filename-matched) path gets the same `customer_block`
  treatment inside `read_full_record`, so reopening an app-generated file
  shows its customer block too.
- **Real editable Attn/Address/PO Box/City/Country for QTN/INV/DO**: the
  `#qtn2-extra` block (previously hidden for every type except QTN2) is
  now always shown — `setType()` no longer toggles it by doc type. The
  same richbox `#customer_attn`/`#customer_address` elements are reused
  for every type (a richbox is just a styled contenteditable div; xlsx
  types simply read `.textContent` off it instead of the HTML QTN2 needs
  for its own template), so no elements were duplicated.
- **`customerBlockForXlsx()`** (new): QTN/INV/DO's bundled xlsx templates
  only have *one* free-text "To" cell (confirmed by inspecting them —
  `C3` for QTN/INV, `F3` for DO — no separate name/address/PO-box
  fields the way QTN2's own HTML template has), so Company + Attn +
  Address + "PO Box X, City, Country" all compose into one plain-text
  (no HTML — `collectQtn2Extra()`/`composedCustomerAddress()` stay
  HTML-based and untouched, only used for QTN2) multi-line block, sent as
  `customer_block` in `collectDocData()`/`runPreview()`'s payload
  whenever `TYPE!=='QTN2'`.
- **`engine._write_customer_block(ws, doc_type, text)`** (new, called
  from `fill_quotation`/`fill_invoice`/`fill_delivery_order`): writes
  into `CUSTOMER_CELL[doc_type]` with `wrap_text=True` — **without it the
  cell's existing center alignment clips a multi-line block equally from
  both ends** (caught by rendering a test file: "TEST CUSTOMER CO
  LLC..." came out as "CO LLCAttn: Mr. Test PersonAl Quoz Industrial
  Area 1PO Box 12345, Dubai, Un" — truncated on *both* sides, no
  linebreaks). Skips writing entirely when blank, so an empty field
  doesn't blank out the template's own border/placeholder design.
- **The merged "To" cell is only 3 rows tall by default and doesn't
  auto-grow** — a real 7-line customer block (name + 2-line address +
  contact person + phone + TRN, typical for this business's real
  documents) got clipped after ~4 lines even with `wrap_text` on, since
  Excel/LibreOffice don't expand a merged cell's row height to fit
  content. Fixed by measuring the block's line count after writing and,
  if it exceeds what the merge's current row heights can hold, growing
  each spanned row's height proportionally (`ws.row_dimensions[r].height`)
  — confirmed by re-rendering the same 7-line test block and seeing every
  line fit with no clipping.
- **Round-trip de-duplication**: since `customerBlockForXlsx()` always
  puts the company name first, reopening an app-generated file (via
  `openDoc()` or `loadDraft()`) strips the block's first line back off
  when it exactly matches the Company field — otherwise re-Generating
  would duplicate the company name a little further down the Address box
  every time the doc gets reopened and resaved. A legacy file's own first
  line essentially never matches the (usually shorter, folder-derived)
  company name exactly, so real historical content is left untouched.
  Attn/PO Box/City/Country are never guessed apart from the recovered
  blob — there's no reliable way to tell which line is which — so on
  reopen the whole thing lands in Address and Attn/PO Box/City/Country
  stay blank for the user to redistribute if they want to.
- `fillFromClientIfBlank`/`applyClientPick`/`saveCurrentAsClient` had
  their `TYPE==='QTN2'`-only guards removed, since the fields they touch
  are now visible (and meaningful) for every type.
- Verified against the real INV-10188 file (read-only, no Generate/Save
  against F:\ data): QTN Number (30227) and LPO Number
  (LTTR-PO-00241-1), previously blank, now populate; the full 7-line
  customer block (LTS TRADING LLC's name/address/contact/TRN) appears in
  Address; opening a fresh new Tax Invoice shows the same fields empty
  and ready to type into, with the live preview correctly composing them
  into the To box; confirmed QTN2 is completely unaffected (still its
  own richbox/HTML path, `#qtn2-extra` still visible, no console errors).
  Generation itself was verified only via standalone scratch-directory
  scripts calling `engine.fill_invoice` directly (never through the
  running app's Generate button, and never against a real F:\ folder),
  per the standing rule against exercising Generate/Save on production
  data.

## Two-step confirmation before overwriting an existing document

Requested directly: editing a previously-created document should carry a
persistent warning, and saving should require reviewing a change summary
plus a separate red "are you sure" step before anything actually gets
written — `window.confirm()`/`alert()`-style native dialogs aren't enough
here (and `confirm()`/`prompt()` are already documented elsewhere in this
file as silent no-ops in this environment anyway), so this is custom UI.

- **`.editwarnbanner`**: a persistent amber strip under the edit modal's
  bar, always visible for the lifetime of the modal (not a toast that
  fades) — "You're editing a document that was already created…".
- **`EDIT_SNAPSHOT`**: a deep clone (`JSON.parse(JSON.stringify(...))`)
  of `collectDocData()`, taken at the end of `enterDocEditMode()` once
  every field has been populated. **Must be a deep clone, not just a
  fresh object** — `collectDocData()`'s `items` array is `items.filter(...)`,
  which returns a new array but the *same item objects* `upd(i,k,v){items[i][k]=v}`
  mutates in place; a shallow snapshot would silently drift to match
  every edit as it happened, making the diff always show "no changes."
  Confirmed this would have been a real bug by checking `EDIT_SNAPSHOT`
  after mutating a live item — the clone stayed at its original value.
- **`onGenerateClick()`**: the Generate button's new entry point —
  `requestSave()` (the two-step flow) when `EDIT_MODE` is true, otherwise
  the existing plain `generate()` for brand-new documents, which need
  none of this ceremony.
- **`diffDocData(orig, cur)`**: field-by-field comparison against a fixed
  set of header keys (number/date/project/qtn_number/lpo_number/etc,
  company, customer_block) plus discount/vat and a per-index item
  comparison (`JSON.stringify` equality — good enough for "did this
  change," not trying to describe *how*). Returns human-readable strings
  like `Company: "LTS" → "LTS Trading"` or `Item 2 changed`.
- **Step 1** (`requestSave()`): validates company/number are filled (same
  checks `generate()` itself does, just fired earlier so the confirmation
  dialog doesn't open on an invalid form), then shows the diff in
  `#saveconfirmmodal` with a plain "Continue to Save" button.
- **Step 2** (`showSaveConfirmStep2()`): swaps the same modal's content to
  bold red (`#b91c28`) warning text and a red "Yes, Save" button next to
  Cancel. The wording is conditional on `EDITING`: a true overwrite
  ("cannot be undone") vs. an imported best-effort doc, which actually
  saves as a *new* file ("the original file's layout doesn't match the
  app's own template, so it won't be modified") — the two cases have
  genuinely different consequences and shouldn't share one warning.
  `confirmAndGenerate()` only calls the real `generate()` after this step.
- **Real bug caught while verifying**: `#saveconfirmmodal` reuses
  `.clientmodal`'s styling, which was `z-index:210` — one *lower* than
  `.editmodal`'s `220`. Since this dialog only ever opens from inside the
  edit modal, it was rendering completely hidden behind it (confirmed via
  `classList.contains('hide')` returning `false` — i.e. genuinely "shown,"
  just stacked underneath). Fixed by bumping `.clientmodal` to `225` —
  which incidentally also fixes the *same* pre-existing bug for the
  Product Builder and Client editor modals, both `.clientmodal` too and
  both reachable from inside the edit modal's line-items card, that
  nobody had noticed yet.
- Verified against the real INV-10188 file (read-only up to the final
  step — no Generate/Save run against F:\ data, per the standing rule):
  opened it, changed a line item's price, clicked Generate → "Review
  changes" modal correctly showed "Item 1 changed" only; clicked
  "Continue to Save" → red confirmation correctly explained this
  particular file saves as a new document (not an overwrite), since it's
  a best-effort import; clicked **Cancel** (not "Yes, Save") to stop
  short of touching the real file, confirmed via the network log that no
  `/api/generate` request was ever sent.

## Submissions: QTN → LPO → DO → scanned DO → INV, bundled and tracked

New top-level feature per the user's own description of their real
workflow: a confirmed Quotation gets an LPO from the customer; the
LPO-confirmed quantities immediately produce a real Delivery Order; once
the material is delivered and the signed DO is scanned back in, the
system links it and generates the Invoice from those same quantities —
Quotation, LPO, DO, and Invoice end up tracked together as one record.
Scoped via upfront clarifying questions rather than guessed: LPO is a
real uploaded file (no OCR — LPO layouts vary too much per client to
parse reliably, same reasoning as everywhere else in this app); matching
is a human reviewing/adjusting per-item quantities, not an automatic
hard block; the DO generates for real immediately, not as a draft; and
linking the scanned DO is a manual file pick from a **dedicated new
Settings folder** for scanned DOs (the user's own answer — not reusing
`do_folder`, since in practice the scanned copies land somewhere else
before being filed).

- **New setting**: `scanned_do_folder`, same pattern as every other
  document folder setting (`SETTINGS_FIELDS`, Settings UI card, native
  folder-choose dialog via the existing `/api/browse`).
- **`submissions/<BRAND>.json`** (new, mirrors `drafts/<BRAND>.json`'s
  storage pattern): one record per submission — `qtn_rel`/`qtn_number`,
  `company`/`project`, `items` (each with its own `lpo_qty`, the
  LPO-confirmed quantity, alongside the quotation's original `qty`),
  `lpo_number`/`lpo_filename`/`lpo_saved_name`, `do_number`/`do_rel`,
  `scanned_do_rel`, `inv_number`/`inv_rel`, and `stage` (`do_generated` →
  `delivered` → `invoiced`). The uploaded LPO file itself is **not**
  base64-stored in the JSON (would bloat it for a multi-page PDF) — it's
  saved to disk at `submissions/<BRAND>/<id>/lpo.<ext>` and only the
  filename is referenced, same "real files, not blobs in JSON" instinct
  as every other document folder in this app.
- **`POST /api/submissions`**: creates the record **and** immediately
  generates a real DO via `engine.generate("DO", ...)` — the same
  function the normal Build/Generate flow uses — with `lpo_qty` (not the
  quotation's original `qty`) as the DO's quantity, so a partial
  delivery is reflected correctly from the start. Numbering reuses
  `_next_number_for()`, a small refactor pulled out of the existing
  `/api/next-number` route so both share one implementation.
- **`POST /api/submissions-generate-invoice`**: generates the Invoice
  from the **same** `lpo_qty` values used for the DO (not the
  quotation's original quantities) — DO and INV always agree on what was
  actually ordered/delivered because they're both built from the one
  shared items list on the submission record, never re-derived from
  re-reading either generated file.
- **`POST /api/browse-scanned-do`**: a native file-*open* dialog (not the
  existing `/api/browse`, which only does folders or the one special
  save-as case for `clients_file`) rooted at `scanned_do_folder` via
  `initialdir`. Returns a path relative to that folder when the pick is
  inside it, or falls back to an absolute path if the user browses
  elsewhere — `_resolve_scanned_do()` handles either form when opening it
  back.
- **`/api/doc` extended to read QTN2 sidecars**: the Submissions picker
  needs a confirmed quotation's items regardless of whether it's a
  legacy xlsx or a `QTN2` PDF — `_api_doc_from_pdf()` gained a `QTN2`
  branch that just returns `engine.read_sidecar(path)` (the exact
  generation payload already saved next to every QTN2 PDF) rather than
  re-parsing anything, since QTN2 never needed round-trip reading before
  this.
- **Real bug caught while testing**: a JS string had `customer\\'s`
  (double backslash) instead of `customer\'s` (single) — the extra
  backslash closed the string early, breaking every function
  declaration after it in the whole script (`typeof view` came back
  `"undefined"` — the entire inline script had failed to parse, not just
  one function). Traced by extracting the served page's inline
  `<script>` block and confirming the exact malformed sequence in the
  file, since the browser console gave no error at all (a hard parse
  failure before any code runs doesn't always surface as a catchable
  runtime error the console-log tool can see). Fixed both occurrences.
  Lesson: when hand-writing JS string literals with an apostrophe inside
  single-quoted strings, it's `\'` (one backslash) — `\\'` is an escaped
  backslash *followed by* a real, string-ending quote.
- **Real UI-chrome bug caught while testing**: `#saveconfirmmodal` (see
  the two-step confirmation section above) reusing `.clientmodal`
  surfaced that its `z-index:210` sat *below* `.editmodal`'s `220` —
  fixed there already by bumping `.clientmodal` to `225`, which this
  feature's own confirmation dialogs (and the QTN-picker/item-confirm
  modal, also `.clientmodal`) benefit from too.
- The "Import Scanned DO" native file dialog can't be driven by browser
  automation (it's an OS-level window, not part of the page) — verified
  its backend (`/api/submissions-link-scanned-do`) directly instead by
  placing a dummy file in a scratch "scanned DOs" folder and calling the
  endpoint with its relative path, then confirming the UI picked up the
  `delivered` stage correctly after a refresh.
- **Verified fully end-to-end using a temporary settings swap to scratch
  folders** (per the standing rule against exercising Generate/Save on
  real F:\ data): backed up `config.json`, pointed `do_folder`/
  `inv_folder`/`scanned_do_folder` at throwaway scratch directories,
  ran the complete flow through the real running app — picked a real
  quotation (LTS QTN 30268, read-only search against real data), 
  confirmed an LPO quantity, created the submission (a real DO
  generated into the scratch folder, confirmed on disk), linked a dummy
  scanned DO, generated the Invoice (confirmed on disk in the scratch
  INV folder), deleted the test submission, then restored the original
  `config.json` and confirmed via `/api/settings` that the real F:\
  paths were back before continuing. No real document folder was ever
  written to.

## Fixed: stray "Could not render a preview." text under a perfectly good preview

User-reported, with a screenshot circling the text sitting in the empty
canvas below an otherwise correctly-rendered single-page preview. The
string only exists in one place — `setPreviewImage()`'s `img.onerror`
handler — but that function is only called from `generate()`/`openDoc()`,
neither of which had run in the reported scenario (a brand-new,
never-generated Quotation). Root cause: `img.onload`/`img.onerror` are
assigned directly onto the `#previewimg` element and **persist across
calls** — if an earlier `setPreviewImage()` call's image request is slow
or fails, and a *newer* `showPreviewPages()` call renders the correct
multi-page preview before that old request resolves, the stale
`onerror` still fires when it eventually does, unhiding `#previewempty`
with the error text *underneath* the now-correctly-showing
`#previewpages` — both simply stack in normal document flow since
neither is absolutely positioned over the other.

Fixed with a monotonic `previewImgToken`: `setPreviewImage()` captures
the token at call time and its `onload`/`onerror` callbacks bail out
immediately if the token has since changed; `showPreviewPages()` bumps
the token too, so it invalidates any in-flight `setPreviewImage()`
request the moment a newer preview supersedes it. Verified by
deliberately reproducing the race — called `setPreviewImage()` with a
URL guaranteed to 404, immediately followed by `showPreviewPages(1)`,
then waited for the failed request to resolve — confirmed
`#previewempty` stayed hidden with its default text and `#previewpages`
stayed visible, where before the fix this exact sequence would have
surfaced the bug.

## Submissions rework: auto-triggered from Quotation approval, DO+INV together upfront

Follow-up request: automate the Submissions pipeline built earlier rather
than requiring a manual "+ New Submission" pick every time. Scoped via
clarifying questions before touching anything, since this genuinely
changed the sequencing of what was already built: LPO moves to the *end*
of the flow instead of the start; the DO and Invoice generate together,
immediately, from the quoted (not LPO-confirmed) quantities; "in progress"
is a real, visible state (red in All Docs) rather than an internal-only
notion; and "build the submittal" produces an actual merged PDF, not just
a status flip.

- **New stage sequence**: `in_progress` (QTN Approved → DO + INV both
  generated for real, right away, from the quotation's own quantities) →
  `delivered` (scanned DO linked — the "delivery confirmed" moment) →
  `submittal_built` (LPO attached, quantities reconciled, one merged PDF
  produced). Replaces the old `do_generated`/`delivered`/`invoiced`
  sequence from the first pass — LPO no longer gates DO creation at all.
- **Mandatory status gate**: `onGenerateClick()` now routes QTN2 through
  `openStatusGate()` before ever calling `generate()` — a modal listing
  all five statuses with descriptions, defaulting to nothing pre-selected
  as "confirmed" (every click is an explicit choice, not a passive
  acceptance of whatever the segmented control already showed). Picking
  **Approved** shows a second prompt — "generate the DO and Invoice now
  too?" — before proceeding; any other status just generates normally.
  `generate()` was changed to actually `return` its result (previously a
  fire-and-forget side-effecting function) so the Approved-flow can read
  back the just-generated PDF's filename and reuse it as `qtn_rel`.
- **`_generate_do_and_inv()`** (new, in `app.py`): the shared core both
  the automatic (post-Approval) and manual ("+ New Submission") paths
  call — generates DO and INV **together**, from one shared items list,
  optionally with explicit `do_number`/`inv_number` passed in so a later
  correction overwrites the same files instead of minting new ones.
- **All Docs red highlight**: `/api/index` cross-references
  `load_submissions()` — any DO/INV belonging to a submission that
  hasn't reached `submittal_built` gets `in_progress: true`, matched by
  **`(type, number)`, not file path** (numbers are the stable identifier;
  paths can be stored differently depending on how a record was reached).
  **Real bug caught here**: the first version compared submission numbers
  (plain ints, e.g. `1`) against All Docs' filename-derived numbers
  (zero-padded strings, e.g. `"0001"`) as raw strings — `"1" != "0001"`
  never matches, silently no rows ever highlighted. Fixed by comparing
  numerically (`int(num)`) on both sides. `.row.inprogress` gets a red
  tint/border and an "In Progress" pill in the row.
- **`engine.build_submittal_pdf(parts, out_path)`** (new): merges the
  Quotation + LPO + scanned DO + Invoice — any mix of PDF and image
  (jpg/png) — into one PDF via PyMuPDF. Verified directly first that
  `fitz.Document.insert_pdf()` flatly refuses a non-PDF source
  ("source or target not a PDF") before writing the real function, so an
  image part gets its own new page sized to the image instead of being
  inserted as pages. Skips any part that's missing, unreadable, or not a
  real file (confirmed in testing: a dummy scanned "DO" that was actually
  a plain text file with a `.pdf` extension was silently skipped rather
  than failing the whole merge) — a submittal built from 2 of 4 real
  parts is still useful.
- **`/api/submissions-build-submittal`** (new, replaces the old
  `-generate-invoice` route): this is where the LPO shows up for the
  first time. If the confirmed delivered quantities differ from what was
  quoted, it calls `_generate_do_and_inv()` again with the *same*
  `do_number`/`inv_number` — overwriting the existing files in place —
  before merging. `_resolve_qtn_pdf()` handles getting a PDF for the
  Quotation regardless of whether it's a legacy xlsx (uses its `.pdf`
  sibling, or converts on the fly) or already a QTN2 PDF.
- **Real reliability bug caught while testing `_generate_do_and_inv()`**:
  generating DO then INV back-to-back triggered a genuine LibreOffice
  headless conversion failure — `soffice.exe` exited non-zero with no
  useful stderr. Retrying the *exact same* conversion by hand a few
  seconds later succeeded immediately, confirming it's LibreOffice's
  per-profile lock still being held from the first conversion, not a bad
  file. `engine.to_pdf()` now retries once after a 2-second delay before
  raising — a small, targeted resilience fix directly motivated by a
  failure this feature's own back-to-back-generation pattern made far
  more likely to hit than before.
- Manual "+ New Submission" flow simplified to match: step 2 just shows
  the quoted items read-only and generates DO+Invoice immediately (no
  LPO fields anymore — moved entirely to the Build Submittal step).
- Verified the complete chain against real data end-to-end (temporary
  settings swap to scratch `qtn2_folder`/`do_folder`/`inv_folder`/
  `scanned_do_folder`, config backed up and restored after, per the
  standing rule): created a real Quotation → mandatory status modal →
  Approved → confirmation prompt → DO 1 + INV 1 generated for real and
  confirmed red/"In Progress" in All Docs → linked a scanned DO → Build
  Submittal modal → merged `submittal.pdf` confirmed on disk with the
  expected page count → stage correctly reached `submittal_built` →
  deleted the test submission and restored the original folder settings.

## Nav reorder, colorized QTN2 status everywhere except the document, Terms & Conditions builder

Three smaller, independent requests handled together:

- **Nav reorder**: `#n-clients` moved to sit directly above `#n-settings`,
  below the `flex:1` spacer — was previously between All Docs and
  Submissions. Pure markup reorder, no JS/logic change.
- **Colorized QTN2 status, everywhere except the printed PDF**: the user
  was explicit that colorizing is welcome throughout the app UI but must
  **not** touch the document itself. `templates_html/quotation.html`'s
  own status pill was already a single neutral navy-outline style
  regardless of status value — confirmed unchanged, left exactly as is.
  Added a shared color scheme (Draft=gray, Sent=blue, Approved=green,
  Revised=amber — "None" renders no badge at all, since it means no
  status tracking) applied in three places:
  - **All Docs**: `/api/index` now reads `status` out of the QTN2
    sidecar (same `read_sidecar()` already used for attn/address) and
    attaches it to the record; `renderList()` renders a `.statuspill`
    next to the row.
  - **Build tab's own status segmented control** (`#qtn2-statusseg-top`):
    the *active* button now takes its category color instead of the
    generic dark highlight — `#qtn2-statusseg-top button.on[data-s=X]`
    attribute-selector rules, no JS change needed since `data-s` was
    already on each button.
  - **The mandatory status-confirm modal** (see the two-step-confirmation
    section above): each option button gets a `border-left` accent in
    its category color via a new `STATUS_COLOR` JS map.
- **Terms & Conditions builder** (QTN2 only — `html_engine.py` already
  accepted a `data.terms{delivery,payment,warranty}` dict with sensible
  defaults; this just adds the UI to actually set it instead of always
  falling back to the hardcoded defaults):
  - **Delivery**: a preset dropdown (the user's own "ranged table" —
    1-2/2-4/4-6/6-8/8-10/10-12 weeks, **10-12 standard/default**) with
    the same "Custom…" escape-hatch pattern used elsewhere (Unit, Product
    Builder dropdowns) for anything outside the presets.
  - **Payment**: a variable list of `{percent, label}` stages (2 by
    default — Advance/Upon Delivery, matching the prior hardcoded
    standard — add/remove for e.g. a 3-stage Advance/Interim/Final
    split), composed into `"50% Advance, 50% Upon Delivery"` on the fly.
  - **Warranty**: single-select segmented control among 3/5/7/10 years
    (5 is standard/default), despite being described as "checkboxes" in
    the request — a quotation can only carry one warranty duration at a
    time, so a radio-style single-select is what "standard box plus
    additional boxes" actually means functionally.
  - **`collectQtn2Extra()`** now also returns `terms` (the composed
    strings `html_engine.py` expects) and `terms_ui` (the raw
    delivery/payment-array/warranty state) — deliberately two different
    shapes, since round-tripping the *composed* string back into
    structured rows on draft-reload would be lossy guessing, while
    `terms_ui` lets `loadDraft()` restore the exact editable state
    without re-parsing anything.
  - `resetTerms()` (delivery back to 10-12 weeks, payment back to the
    2-stage standard, warranty back to 5 years) wired into `setType()`'s
    reset branch and the initial page-load `setType('QTN2')` call, so a
    fresh document always starts from the standard terms.
- Verified live: nav order correct; generating a Quotation with status
  set to Draft/Sent/Approved/Revised each showed the matching color on
  the Build tab's own segmented control and (via a scratch-folder test,
  config backed up and restored after) the correct colored pill in All
  Docs — confirmed "Approved" rendered green with no red "In Progress"
  pill when auto-DO/INV generation was declined; changed Delivery to
  "2-4 weeks" and Warranty to 10 years and confirmed the live PDF preview
  updated its Terms & Conditions block to match, while the status pill
  printed on the document itself stayed its original plain, uncolored
  style throughout.

## Fixed: New Submission showed blank company and blank quantities for legacy QTNs

User-reported via a screenshot: "New Submission — " with nothing after
the dash, and every item showing "Qty" with no number. Both traced back
to the same best-effort xlsx-reading path used when a quotation's
filename doesn't match this app's own convention (the common case for
real historical files) — and both were real, non-cosmetic bugs, not
just missing niceties.

- **Blank quantities**: `engine.read_items_from_doc()` — the best-effort
  line-item reader — only ever extracted `description` (plus a photo if
  present), never `unit`/`qty`/`price`, regardless of layout. Confirmed
  by inspecting the exact file from the report: it turned out to use
  this app's own precise column labels (`No/Type/Photo/Item Description/
  Unit/Qty/Price/Vat 5%/Amount` in the header row) despite not matching
  the filename convention — the data was sitting right there, just never
  read. Fixed by having the function scan that same header row (already
  located to find the description column) for cells labeled Unit/Qty/
  Price/Type — a few label variants each, case-insensitive — and reading
  those columns per item row too. Each column is independently optional,
  so a layout missing one still gets whatever it does have.
- **Blank company**: worse than just missing — the *existing* fallback
  (top-level folder name) was actively wrong for a large share of real
  files. Checked 15 real quotations directly: files sitting in the
  folder root had no folder name to fall back to at all (blank), and
  files organized under a folder like `"ABU DHABI QTN"` or
  `"ADS QUOTATION"` inherited that as their "company" — but those are
  **region/category buckets holding many different clients' files**, not
  a client name at all. Fixed by preferring the customer block's own
  first line (the real company name, already being read via
  `read_xlsx_customer_shape()`/the cell fallback for the `customer_block`
  field) as the primary source, falling back to the folder name only
  when no customer block could be found at all. This also happens to
  make `openDoc()`'s existing first-line-dedup logic (documented
  earlier: strip the block's first line from Address when it matches
  Company, so re-Generating doesn't duplicate it) *more* correct than
  before, not less — now that Company is genuinely derived from that
  same first line, the two always agree.
- Verified against the exact file from the report
  (`ARTEMIS_ QTN-30265_LTS_J2 VILLA_DECORATIVE_05.08.2025_rev1.xlsx`,
  read-only): `/api/doc` now returns `company: "LTS TRADING LLC"`
  (previously `"LTS QTN"`, the enclosing folder, not the real client)
  and both items carry their real `qty`/`unit`/`price`; reproduced the
  same result live through the actual New Submission picker — title
  correctly reads "New Submission — LTS TRADING LLC" and both items show
  "Qty 1" instead of blank.

## Statement of Account — income, outstanding balances, a chart

New top-level feature. Scoped via clarifying questions first, since the
central design question — how to handle payment history for thousands of
real historical invoices with zero payment records — determines whether
the feature is trustworthy on day one or not: manual Paid/Unpaid toggle
per invoice (no bank integration exists to do this automatically), and
**only track from today forward** — historical invoices are excluded
from the whole feature entirely rather than defaulting to "unpaid" (which
would make the outstanding balance look enormous and almost entirely
wrong, since most old invoices were presumably already settled).

- **`finance/<BRAND>.json`** (new, mirrors `drafts/`/`submissions/`'s
  per-brand JSON pattern): one entry per Invoice generated from today
  onward — `rel`, `number`, `company`, `date`, `subtotal`/`vat`/
  `discount`/`total`, `paid`, `paid_date`. Nothing here is backfilled
  from the historical archive; the ledger starts empty and only grows as
  new Invoices are generated.
- **`record_invoice_in_ledger(data, xlsx_path, brand, inv_folder)`**:
  computes totals directly in Python from the same items/discount/vat
  the invoice was just filled with, via `html_engine.compute_totals()` —
  the exact same function QTN2 already uses for its own printed totals,
  now shared rather than duplicated. Deliberately does **not** read the
  total back out of the xlsx's formula cells — confirmed by checking a
  real historical invoice that those DO carry cached values (Excel
  computes and saves them whenever a human opens the file), but a
  brand-new file this app just wrote via openpyxl has never been opened
  by anything that computes formulas, so its cached values would still
  be blank. Called from both `/api/generate` (the normal Build/Generate
  flow, gated to `dtype=="INV"`) and `_generate_do_and_inv()` (the
  Submissions auto-generation) — every real path that produces a real
  Invoice feeds the same ledger.
- **Editing an existing invoice and renaming it** (different number/
  company/date changes the filename) drops the stale ledger entry for
  the old, now-deleted file rather than leaving an orphan pointing at
  nothing — handled right where `/api/generate`'s existing `replace`
  logic already deletes the old xlsx/pdf.
- **`GET /api/finance/ledger`** / **`POST /api/finance/mark-paid`**:
  the only two routes needed — aggregation (by month, by company) happens
  client-side over the raw ledger, since it's small and grows slowly.
- **Chart**: hand-rolled inline SVG (no charting library, consistent
  with the rest of this app), built per the dataviz skill's procedure —
  grouped bars (Invoiced/Collected) by month, legend, gridlines, a hover
  tooltip, direct `data-tip` labels. Colors reuse the status-color pair
  already established elsewhere this session (blue=Sent/informational,
  green=Approved/good) rather than inventing a new palette. Skill also
  calls for running `scripts/validate_palette.js` on any chart palette —
  **couldn't**, no Node.js in this environment (confirmed:
  `node --check`/`node --version` both fail) — so this pair is a
  judgment call, not a validated one; compensated with secondary
  encoding (legend, direct tooltip labels, fixed left/right bar position
  per series) so identity was never color-alone regardless.
  - **Real bug caught while verifying**: the gridline builder wrote
    `stroke-width=1/>` (no space before the self-closing slash). Because
    unquoted HTML attribute values can legally contain `/`, the browser
    parsed that as `stroke-width="1/"` with **no actual close** — every
    `<line>` stayed open, silently swallowing all the `<rect>` bars and
    subsequent `<line>`s as its own (invalid) children, none of which
    render inside an SVG `<line>`. The chart looked completely empty
    (just the legend) with zero console errors, since this is valid-
    enough markup to parse without throwing — only caught by inspecting
    the actual rendered DOM (`outerHTML`) and noticing five stray
    `</line>` closing tags bunched at the very end, all nested inside
    each other. Fixed by quoting the value and spacing the slash
    (`stroke-width="1" />`). Lesson: self-closing SVG tags built via
    string concatenation need a space before `/>`, or an unquoted
    numeric attribute right before it will absorb the slash.
- Verified end-to-end in a scratch `inv_folder` (config backed up and
  restored after, per the standing rule): generated two real test
  invoices for two different companies, confirmed both landed in the
  ledger with correct computed totals (1000 + 5% VAT = 1050 each);
  opened Statement of Account and confirmed the KPI tiles, chart, and
  company table all matched; toggled one invoice to Paid and confirmed
  Collected/Outstanding updated live, the chart grew a green bar, and the
  company table re-sorted by remaining balance — then cleared the test
  ledger entries and restored the original settings.
