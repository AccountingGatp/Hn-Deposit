# Deposit Reconciliation (Batch · Transactions · Bank)

A self-service web page that reconciles **three files** describing the same money
at three stages of settlement, and produces a formatted Excel workbook with a
built-in three-way tie-out — in GATP house style.

Anyone on the team can open the page, drop in the three files, and download a
`.xlsx`. **Everything runs in the browser** — no server, and no financial data
ever leaves the user's machine.

---

## The three files

| # | File | What it is | Key columns |
|---|------|-----------|-------------|
| 1 | **Batch file** | One row per settlement **batch**, with the batch total. | `Batch #`, `MID`, `Total Amount`, `Location`, `Batch Close/Submit/Settlement Date`, `Status` |
| 2 | **Transaction export** | The individual card **transactions (splits)** that roll up into each batch. | `Date`, `Amount`, `MID`, `Batch #`, `Transaction`, `Result`, `Payment Method`, … |
| 3 | **Bank deposit register** | What actually **landed in the bank** (one row per MID per settlement). | `Date`, `Description` (contains the MID), `Received` |

## The two ties

- **Reconciliation 1 — Transactions → Batches.**
  For each batch, sum the transaction splits that carry its `Batch #` and compare
  to the batch `Total Amount`. Any difference is flagged as a **split mismatch**.
  Transactions with **no** batch number (authorizations, cash/check, declined,
  voids) are never dropped — they are listed and labelled on the *Transactions*
  sheet.
- **Reconciliation 2 — Batches → Bank Deposits.**
  Each batch is matched to a bank deposit by **MID + amount**. Matched batches are
  *Settled & deposited*; unmatched batches are *Awaiting deposit* (a timing
  difference — submitted but not yet settled in the bank). Any bank deposit with
  **no** matching batch is raised as an **exception**.

---

## How to use

1. Open `index.html` (locally or from wherever you host it — see **Deploy**).
2. **Step 1 — Upload three files** (`.csv`, `.xls`, or `.xlsx`). Columns are
   detected automatically.
3. **Step 2 — Confirm the output** — set the client name used in the file name.
4. **Step 3 — Review & download** — read the three-way summary, scan the batch
   table and any exceptions, then download
   `‹Client›_Deposit_Reconciliation_‹range›.xlsx`.

## Output workbook (5 sheets, GATP house style)

- **Summary** — the three-way tie-out. **All amounts are live formulas** pulling
  from the other sheets, ending in three `OK / MISMATCH` check cells:
  1. batch totals = sum of their splits,
  2. deposited-batch total = matched bank-deposit total,
  3. every bank deposit is matched to a batch.
- **Batch Recon** — one row per batch: batch total, split count & sum, variance,
  whether it was deposited, the deposit amount/date, and a recon status.
- **Transactions** — every split, flagged with the batch it rolls into (or a
  plain-language note for the ones with no batch).
- **Bank Deposits** — every deposit, with its matched batch and status.
- **Exceptions** — anything that did not tie, in simple language (also as cell
  comments). Shows "None" when all three files reconcile.

Formatting: navy `#1F3864` header, white bold Arial, alternating `#D9E1F2` fills,
thin borders, frozen header row, gridlines off, Arial 10 body, `#,##0.00` amounts.

## What the sample data reconciles to

Running the three included sample files (batch `06/17–06/28`, bank export through
`06/27`):

- **80 batches** totalling **$21,442.53** — every batch total ties exactly to the
  sum of its splits (**0 mismatches**).
- **513 transactions** ($26,727.90); 409 roll into batches ($21,442.53), 104 carry
  no batch (auth / cash / check / declined, $5,285.37).
- **64 bank deposits** totalling **$16,630.19** — **all** match a batch.
- **64 batches settled & deposited** ($16,630.19); **16 awaiting deposit**
  ($4,812.34) that settle `06/29–06/30`, after the bank export window.
- All three CHECK cells → **OK**.

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
index.html               # the page (three dropzones + results)
assets/styles.css        # GATP styling
assets/core.js           # reconciliation engine (parsing, ties, matching) — DOM-free, unit-testable
assets/xlsx-build.js     # builds the styled 5-sheet workbook (ExcelJS)
assets/app.js            # UI controller (file reading, column detection, review tables, download)
vendor/xlsx.full.min.js  # SheetJS (reads .csv/.xls/.xlsx)
vendor/exceljs.min.js    # ExcelJS (writes styled .xlsx with formulas & comments)
```
