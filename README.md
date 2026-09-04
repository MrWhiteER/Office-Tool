# Office Tool

A desktop app for Artemis Group that generates the company's documents —
Quotations, Tax Invoices, Delivery Orders, Expense Reports, and Sololuce
Datasheets — as pixel-accurate PDFs, and keeps track of the paperwork
around each job from quote to signed delivery.

It's a real installed Windows app (not a website, no browser tab to keep
open) with a live preview, so what you see on screen while you're typing is
exactly the PDF you'll get.

## Install

Download the latest installer from
**[Releases](https://github.com/MrWhiteER/Office-Tool/releases/latest)** —
`OfficeTool-Setup.exe`. Run it, then open **Office Tool** from the Start
Menu or your desktop. That's it; nothing else to install alongside it.

Already have it installed? It checks for updates on its own and installs
them with one click from the **Update** button in Settings — no need to
download a fresh installer each time.

## What it does

**Documents.** Quotation, Tax Invoice, Delivery Order, and Expense Report,
each with a live preview pane that renders as you type. Every document is
auto-numbered, saved as PDF into the right folder, and shows up in **All
Docs** and the client's own history automatically.

**Submissions.** Once a quotation is approved, it becomes a Submission —
one place that tracks the Delivery Order, the signed/scanned copy of it,
the Invoice, the client's LPO, and the final submittal pack, all tied
together for that one job.

**Scanner.** A physically connected scanner can be driven straight from the
app — scan a page, add more, and either save the result as a plain file or
link it directly to a Submission's Delivery Order.

**Sololuce Datasheets & Full Catalog Builder.** Product datasheets for the
Sololuce brand, plus a tool that combines a set of datasheets into one
paginated catalog with a cover, index, and reorderable page order.

**Multi-user accounts.** Each person signs in with their own account. An
admin manages who has access to what — which brands, which tools — from
one place, and everyone's account list stays in sync across every install.

**Records.** All Docs, Statement, and Clients give a searchable view across
everything that's been generated, with each client's history and totals in
one place.

## Your data

Everything you generate is saved locally, into folders you choose (set once
in Settings). The shared pieces — the account list and the product photo
library — sync through a private cloud store so every install sees the same
thing; nothing else leaves your PC unless you explicitly export or send it.

## For developers

This is a Flask app wrapped as a native window (via `pywebview`), packaged
with PyInstaller and Inno Setup. To run it from source:

```bash
pip install -r requirements.txt
python app.py
```

To build the installer yourself, see `build.bat` — it also explains the
one-time setup (`playwright install chromium`) the build needs.
