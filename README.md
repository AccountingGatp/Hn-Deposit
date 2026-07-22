# GATP QuickBooks Import Tools · Healthnomics

Self-service web pages that turn source exports into **QuickBooks Online–ready
import workbooks**. Open the page, drop in the files, download a formatted,
validated `.xlsx`. **Everything runs in the browser** — no server, and no
financial data ever leaves the user's machine.

`index.html` is a hub linking the tools:

| Tool | Page | Input → Output |
|---|---|---|
| **Monthly Services → QBO Import (340B Split)** | `monthly-services.html` | Monthly Services + JQ Code List + Patient List → 5-sheet Sales/Credit-Note (340B split) + Journal Ledger |
| **Bank Deposit → QBO Import Sheet** | `bank-deposit.html` | Bank deposits export + Customers list → QBO import sheet with reconciliation |

---

## Monthly Services → QBO Import (340B Split)

Turns a **Healthnomics / Allscripts Monthly Services** export into one workbook
with **exactly five sheets**, ready for SaasAnt import — following the GATP
reusable prompt *"Monthly Services → QuickBooks Import (340B Split)."*

### How to use

1. Open `index.html` → **Monthly Services → QBO Import (340B Split)**.
2. **Step 1 — Upload three files**
   - **Monthly Services** (`.xlsx`) — sheet `Monthly Services`, one row per service line.
   - **JQ Code List** (`.xlsx`) — column `HCPCS`, the 340B drug/procedure codes
     (J-, G-, Q- **and** CPT `7xxxx` codes — the whole list, not just J-codes).
   - **Patient List** (`.xlsx`) — column `Patient Name` in `LNAME,FNAME` (or
     `LNAME,FNAME MI`) form. The EMR Source column and `(blank)` rows are ignored.
3. **Step 2 — Build.** Optionally set the file-name prefix. Click **Build**.
4. **Step 3 — Review & download.** Confirm every source row was placed once and
   the journal balances, then download the `.xlsx`.

### What the tool does (the rules)

**Line items — two per source row** (an *original* and an *Adjustment* line;
header fields repeated on both so each row stands alone):

| Output field | Source / rule |
|---|---|
| Invoice Number | Voucher Number (col B) + Service ID (col C), stored as **text** |
| Customer | Original Ins Category Desc (col BC) |
| Invoice Date | Post Date (col W), `mm/dd/yyyy` |
| Memo / Message / Product-Service Description | = Invoice Number |
| Product/Service | Procedure Code (col BL) + " - " + Procedure Descr (col BM); Adjustment line prefixed `Adjustment - ` |
| Product/Service Amount | Original = Fee (col BP); Adjustment = **−Fee × 67.67%**, `ROUND_HALF_UP`, `0.00` |
| Product/Service Taxable | False |
| Location Name | Place of Service Descr (col Y) |

**Classify 340B** — an invoice is **340B only if both** are true (else Non-340B):
- Procedure Code (col BL) appears in the JQ Code List, **and**
- the patient (built from LName/FName/MI, cols G/H/I) matches the Patient List
  (case-insensitive; tries `LNAME,FNAME` then `LNAME,FNAME MI`).

**Split Sales vs Credit Note** — net the two amounts per invoice:
- Net **≥ 0** → **Sales** sheet for its class (0 is *not* negative → stays in Sales).
- Net **< 0** → **Credit Note** sheet for its class (negative-Fee rows land here).

**Journal Ledger** — one 3-line, balanced entry **per credit-note invoice**:
`Journal No | Date | Memo | Account | Debit | Credit | Description | Name | Class`
- Line 1 (revenue) DEBIT = |original amount|
- Line 2 (adjustment) CREDIT = |Adjustment amount|
- Line 3 `Undeposited Funds-Allscript EOB` CREDIT = Line 1 − Line 2
- Debit = sum of Credits (the entry always ties out). Account strings differ for
  340B vs Non-340B and are used verbatim.

Every source row lands in exactly one sheet; voided rows are kept.

### Output workbook (5 sheets, GATP house style)

`Sales - Non-340B` · `Sales - 340B` · `Credit Note - Non-340B` ·
`Credit Note - 340B` · `Journal Ledger`.

Formatting: Arial 10; navy `#1F3864` header, white bold, frozen top row; thin
gridlines; per-invoice `#D9E1F2` banding; amounts `0.00`; dates `mm/dd/yyyy`;
invoice numbers as **text**; per-sheet total rows and a journal `BALANCED` check.

---

## Bank Deposit → QBO Import Sheet

Turns a **bank deposits export** into a QBO import workbook with a built-in
reconciliation — following the GATP SOP *"Bank Deposit to QBO Import Sheet."*
Upload the bank export + your Customers list, set the parameters, download a
3-sheet workbook (`QBO Import` · `Not Processed` · `Reconciliation`). See the
in-page hints for the full column and mapping rules.

---

## Deploy

Static site — host the folder anywhere:

- **GitHub Pages:** repo → Settings → Pages → deploy from this branch, root.
- **Any static host / intranet:** copy the folder and serve it.
- **Local:** open `index.html`, or run `python3 -m http.server` in the folder.

No build step. Libraries ([SheetJS](https://sheetjs.com) for reading,
[ExcelJS](https://github.com/exceljs/exceljs) for writing) are vendored in
`vendor/` so the pages work fully offline.

## Project layout

```
index.html                # hub linking the tools
monthly-services.html     # Monthly Services → QBO (340B Split) tool
bank-deposit.html         # Bank Deposit → QBO tool
assets/styles.css         # shared GATP styling
assets/ms-core.js         # Monthly Services logic (classify, split, journal) — DOM-free, testable
assets/ms-build.js        # builds the 5-sheet workbook (ExcelJS)
assets/ms-app.js          # Monthly Services UI controller
assets/core.js            # Bank Deposit logic
assets/xlsx-build.js      # Bank Deposit workbook builder
assets/app.js             # Bank Deposit UI controller
vendor/xlsx.full.min.js   # SheetJS (reads .csv/.xls/.xlsx)
vendor/exceljs.min.js     # ExcelJS (writes styled .xlsx with formulas)
```
