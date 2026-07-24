# Remote Deposit Check Extractor

A self-service web page that reads **scanned check PDFs** and builds a
**remote-deposit import sheet** with three columns:

| Date | Description | Amount |
|------|-------------|--------|
| check date | check / reference number | check amount |

Drop in one or more check PDFs, let the page OCR every page in your browser,
review the extracted values against the on-screen check images, and download a
formatted `.xlsx` (or CSV). **Everything runs in the browser** — no server, and
no check images or financial data ever leave your machine.

> There is also a second tool in this repo — **[Bank Deposit → QBO Import](qbo-import.html)** —
> for turning a bank deposits export into a QuickBooks import workbook.

---

## How to use

1. Open the page (see **Deploy** below). It must be **served over http/https**
   — GitHub Pages, or `python3 -m http.server` locally. Opening `index.html` as
   a bare `file://` blocks the OCR web worker.
2. **Step 1 — Upload the check PDFs.** Add one or more scanned check PDFs. A
   single PDF can hold many checks; each check becomes its own row.
3. **Step 2 — Extract.** The page renders every page, runs OCR, and pulls the
   check date, check number, and amount. Backs of checks and remittance
   summaries are skipped automatically.
4. **Step 3 — Review & download.** Every row shows the check image so you can
   confirm the values. Rows the tool isn't sure about are highlighted and
   carry a **RED FLAG** reason. Click any value to edit it, add or delete rows,
   then download `Remote_Deposit_<n>_checks.xlsx`.

## How the extraction works (and where it's careful)

Scanned checks are messy, so the tool is deliberately conservative:

- **One row per physical check.** A check's remittance voucher and its check
  face are collapsed into a single row (matched on the check number, ignoring
  leading zeros and prefixes like the `S2` in `S2 70002553`).
- **Check number** is taken from a labelled *"CHECK NO."* when present, then a
  prominent top-right number, then the MICR line at the bottom. Numbers
  labelled TIN / NPI / contract / phone are never mistaken for a check number.
- **Amount** is cross-checked against the **written ("legal") amount**
  (*"Forty Four and 40/100"*). When the figures and the words disagree, the
  words win and the row is flagged to verify. Single-check PDFs are also
  cross-checked against the deposit total in the file name.
- **Date** prefers a *"Check Date"* / *"Issue Date"* label over a generic date.
- **Review flags.** A row is marked **OK** only when the check number, date and
  amount are all confidently read. Anything unclear, missing, or cross-checked
  gets a **RED FLAG – …** reason. The tool never guesses a value it can't read —
  it flags it for you.

## Output workbook (2 sheets)

- **Remote Deposit** — the import file: `Date | Description | Amount`, with the
  date as a real date, the check number stored as **text** (leading zeros
  preserved), the amount numeric, and a `=SUM` total row.
- **Review** — an audit trail: `Date | Description (Check #) | Amount | Source
  file | Page | Review Flag`, so the three-column import sheet stays clean.

Formatting is GATP house style: navy `#1F3864` header, white bold Arial,
alternating `#D9E1F2` fills, thin borders, frozen header row.

## Deploy

It's a static site — host the folder anywhere:

- **GitHub Pages:** repo → Settings → Pages → deploy from this branch, root.
- **Any static host / intranet:** copy the folder and serve it.
- **Local:** run `python3 -m http.server` in the folder and open
  `http://localhost:8000`.

No build step. All libraries are **vendored** in `vendor/` so the page works
fully offline (nothing is fetched from a CDN):
[pdf.js](https://mozilla.github.io/pdf.js/) renders the PDF pages,
[Tesseract.js](https://tesseract.projectnaptha.com/) does the OCR (the English
model ships in `vendor/tessdata/`), and
[ExcelJS](https://github.com/exceljs/exceljs) writes the workbook.

Because the OCR model is a few MB and OCR is CPU-bound, a large multi-check PDF
can take a few minutes — a progress bar shows where it's up to.

## Project layout

```
index.html                 # the Remote Deposit Check Extractor page
qbo-import.html            # the earlier Bank Deposit → QBO Import tool
assets/rd-styles.css       # app-specific styling (extends styles.css)
assets/rd-core.js          # extraction logic (OCR-text parsing, dedupe) — DOM-free, unit-tested
assets/rd-app.js           # UI controller: pdf.js render + Tesseract OCR + review table + download
assets/rd-xlsx.js          # builds the 2-sheet workbook (ExcelJS)
assets/styles.css          # shared GATP styling
assets/core.js, app.js, xlsx-build.js   # the QBO Import tool's logic
vendor/pdf.min.mjs, pdf.worker.min.mjs  # pdf.js (renders PDF pages)
vendor/tesseract.min.js, tesseract-worker.min.js, tess-core/, tessdata/   # Tesseract.js + English model
vendor/exceljs.min.js, xlsx.full.min.js # ExcelJS (write) / SheetJS (read, QBO tool)
```
