# Bank Deposit → QBO Import Sheet

A self-service web page that turns a **bank deposits export** into a
**QuickBooks Online–ready import workbook** with a built-in reconciliation —
following the GATP SOP *"Bank Deposit to QBO Import Sheet (with Reconciliation)."*

Anyone on the team can open the page, drop in two files, set three values, and
download a formatted `.xlsx`. **Everything runs in the browser** — no server, and
no financial data ever leaves the user's machine.

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
index.html              # the page
assets/styles.css       # GATP styling
assets/core.js          # SOP logic (parsing, TRN, customer mapping) — DOM-free, unit-testable
assets/xlsx-build.js    # builds the styled 3-sheet workbook (ExcelJS)
assets/app.js           # UI controller (file reading, review table, download)
vendor/xlsx.full.min.js # SheetJS (reads .csv/.xls/.xlsx)
vendor/exceljs.min.js   # ExcelJS (writes styled .xlsx with formulas & comments)
```
