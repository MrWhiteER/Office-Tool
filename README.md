# Office Tool

A small program that runs on **your own PC**. It builds your Quotations, Tax
Invoices and Delivery Orders straight from the company Excel templates, saves
each one as **both Excel and PDF** into your documents folder, names them using
the company convention, and shows you the **last 3 quotations** for whichever
company you're quoting.

The preview you see on screen is the PDF.

---

## What you need once
1. **Python** — free from https://python.org/downloads (during install on
   Windows, tick "Add Python to PATH").
2. **LibreOffice** — free from https://libreoffice.org. This is what turns the
   Excel into a PDF. (If your PC has Microsoft Excel only, the Excel files still
   save fine; install LibreOffice to also get the automatic PDF + preview.)

## How to start it
- **Windows:** double-click **run.bat**
- **Mac:** double-click **run.command**

The first run installs two small Python add-ons, then your browser opens the
app at http://127.0.0.1:5000. Leave the little black window open while you work;
close it to stop the app.

## How to use it
1. Click **Choose folder…** and pick the folder where all your documents live.
   The app reads everything there and sorts it by the file name.
2. Pick **Quotation / Tax Invoice / Delivery Order**.
3. The document number is suggested automatically (next in sequence). Fill the
   header and the line items.
4. Type the **Company** name — if you've quoted them before, the last 3
   quotations appear with their items.
5. Press **Generate Excel + PDF**. Both files are saved into your folder and the
   PDF shows on the right.

## File naming convention
```
TYPE_NUMBER_REV_COMPANY_PROJECT_DATE.ext
QTN_0042_R0_Resinal-Developments_Facade-Lighting_2026-06-30.xlsx
```
The app both **creates** names in this format and **reads** existing files that
follow it, so the "All Docs" tab and the company history work automatically.
Files that don't follow the convention are ignored (they won't be touched).

## Adding more document types or templates
Drop a new template into the `templates/` folder and tell me the field
positions — I'll add it. Current types: QTN, INV, DO.

## Note on the "always connected to AI" idea
This tool sorts and fills documents with plain logic — it doesn't need an AI
connection for everyday use, so it works offline and your files never leave your
PC. If you later want AI help for the messy cases (e.g. reading old scanned PDFs
to pull their items, or auto-tidying oddly-named files), that can be added as an
optional step that calls the Claude API only when you ask it to.
