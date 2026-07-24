# GATP → QuickBooks Online Import Tools

Two self-service web pages that turn source documents into
**QuickBooks Online–ready import workbooks**. **Everything runs in the browser** —
no server, and no financial data ever leaves the user's machine.

| Page | Turns… | …into |
|---|---|---|
| [`bills.html`](bills.html) | vendor invoice **PDFs** (Cardinal, AmerisourceBergen, CuraScript SD, McKesson) + a GL mapping | a **QBO Bill import** workbook |
| [`index.html`](index.html) | a **bank deposits export** + Customers list | a **QBO Deposit import** workbook |

The two pages share the same styling and vendored libraries and link to each
other from the header.

---

# AP Vendor Bills → QBO Import (`bills.html`)

Turns vendor invoice **PDFs** into a **QuickBooks Online bill-import workbook**,
following the *Healthnomics AP* prompts. Drop in one or more invoices, add your
GL mapping, and download a formatted `.xlsx` with a review summary.

### How it works (SOP rules)

1. **Extract** from each invoice — vendor is detected automatically:
   | Vendor | Invoice # | Date | Due date | Amount | Lookup key |
   |---|---|---|---|---|---|
   | **Cardinal** | `INVOICE` | `INVOICE DATE` | Net terms | `GRAND TOTAL` | `SHIP TO` number |
   | **AmerisourceBergen** | `Invoice Number` | `Invoice Date` | on the total line | `Total Amount` | `CUSTOMER NUMBER` |
   | **CuraScript SD** | `INVOICE NO.` | `INVOICE DATE` | `NET DUE DATE` | `TOTAL` | `Customer #` |
   | **McKesson** | `Billing No.` | `Billing Date` | statement date | `NET PAYABLE` | Customer (Route/Stop) |
2. **Look up** the ship-to / customer number in the uploaded **GL reference**
   (columns auto-detected by header name) to pull **Class**, **Category**
   (expense account), **Memo/Description**, **Mailing Address** and **Vendor**.
3. **Assemble** one QBO bill row per invoice. Anything missing or unmatched is
   flagged **`[REVIEW NEEDED]`** — the tool never guesses a value.

### Output workbook (3 sheets, GATP house style)

- **QBO Import** — `Bill No. | Vendor | Bill Date | Due Date | Terms | Account |
  Amount | Memo/Description | Class`, with a `=SUM` total. Cells needing a fix
  are shown in red.
- **Summary** — the SOP's brief review table (source file, vendor, invoice #,
  date, amount, ship-to/customer #, class, category, matched-in-GL?).
- **Exceptions** — every invoice carrying a `[REVIEW NEEDED]` / unmatched flag,
  with the reason.

The GL file just needs recognisable headers — e.g. a *Ship-To Number* /
*Customer Number* / *Account Number* column plus *Class*, *Category*, *Memo*,
*Mailing Address* and *Vendor Name*. No GL? The page still runs and flags every
lookup for review.

---

# Bank Deposit → QBO Import Sheet (`index.html`)

A self-service web page that turns a **bank deposits export** into a
**QuickBooks Online–ready import workbook** with a built-in reconciliation —
following the GATP SOP *"Bank Deposit to QBO Import Sheet (with Reconciliation)."*

Anyone on the team can open the page, drop in two files, set three values, and
download a formatted `.xlsx`.

---

## How to use

1. Open `index.html` (locally or from wherever you host it — see **Deploy** below).
2. **Step 1 — Upload two files**
   - **Bank transactions** (`.csv`, `.xls`, or `.xlsx`) with columns
     `Posted Date`, `Full description`, `Amount`.
   - **Customers list** (`.xls`/`.xlsx` exported from QuickBooks) — the `Name`
     column is the master customer list.
3. **Step 2 — Set parameters**
   - **Starting SNo.** — first serial number (increments by 1 per row).
   - **Accounts** — e.g. `Undeposited Funds` (same for every row).
   - **Class** — e.g. `Corporate` (leave blank if unused).
   - **Client** — used in the output file name.
4. **Step 3 — Review & download**
   - See the reconciliation (Processed + Left = Bank total → **OK/MISMATCH**).
   - Assign a customer to any row flagged **⚠ SELECT CUSTOMER**.
   - Download `‹Client›_QBO_Deposit_Import_‹range›.xlsx`.

## What the tool does (SOP rules)

- **Processes only** positive amounts whose memo contains a TRN in the pattern
  `TRN*1*<number>*`. Everything else (debits/transfers out, deposits with no TRN)
  goes to the **Not Processed** sheet with a plain-language reason — never dropped.
- **Customer name** is chosen **only** from your Customers list. Built-in mappings
  for this client:
  | Bank memo contains | Customer |
  |---|---|
  | CAREFIRST GHMSI / OF MD / BLUECH / ADVANT / MEDGAP, or plain GHMSI | **BCBS** |
  | CFMI / CFBC **FEP Non-Pos** | **FEDERAL EMPLOYEE** |
  | CFMI **FEP Postal** | **Fep Postal** |
  | **Wellpoint** MD5C | **Wellpoint** |
  | **36 TREAS** | **36 TREAS** |

  Unrecognised payers are matched by name where possible, otherwise flagged for a
  manual pick (the tool never invents a name that isn't on your list).
- **Description** and **Memo** = the TRN number, stored as **text** so leading
  digits are preserved.

## Output workbook (3 sheets, GATP house style)

- **QBO Import** — `SNo. | Date | Customer name | Accounts | Description | Amount |
  Memo | Class`, with a `=SUM` total row.
- **Not Processed** — `Date | Bank Description | Amount | Reason` (reason also as a
  cell comment).
- **Reconciliation** — all-formula cross-checks pulling from the other two sheets,
  ending in an `OK / MISMATCH` check cell.

Formatting: navy `#1F3864` header, white bold Arial, alternating `#D9E1F2` fills,
thin borders, frozen header row, gridlines off, Arial 10 body, `#,##0.00` amounts.

## Deploy

It's a static site — host the folder anywhere:

- **GitHub Pages:** repo → Settings → Pages → deploy from this branch, root.
- **Any static host / intranet:** copy the folder and serve it.
- **Local:** just open `index.html`, or run `python3 -m http.server` in the folder.

No build step. Libraries ([SheetJS](https://sheetjs.com) for reading,
[ExcelJS](https://github.com/exceljs/exceljs) for writing) are vendored in
`vendor/` so the page works fully offline.

## Project layout

```
bills.html               # AP Vendor Bills → QBO Import page
index.html               # Bank Deposit → QBO Import page
assets/styles.css        # shared GATP styling

assets/bills-core.js     # bills: vendor parsing + GL lookup — DOM-free, unit-testable
assets/bills-build.js    # bills: builds the styled 3-sheet workbook (ExcelJS)
assets/bills-app.js      # bills: UI controller (pdf.js text extraction, review, download)

assets/core.js           # deposits: SOP logic (parsing, TRN, customer mapping) — DOM-free
assets/xlsx-build.js     # deposits: builds the styled 3-sheet workbook (ExcelJS)
assets/app.js            # deposits: UI controller

vendor/xlsx.full.min.js  # SheetJS (reads .csv/.xls/.xlsx)
vendor/exceljs.min.js    # ExcelJS (writes styled .xlsx with formulas & comments)
vendor/pdf.min.mjs       # pdf.js (reads invoice PDFs in-browser)
vendor/pdf.worker.min.mjs # pdf.js worker
```

## Note on serving

`bills.html` loads `pdf.js` as an ES module, so it must be served over
`http(s)://` (not opened as a `file://` path). Use GitHub Pages, any static
host, or `python3 -m http.server` in the folder. `index.html` works either way.
