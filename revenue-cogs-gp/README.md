# Revenue · COGS · Gross &amp; Net Profit Report

A self-service web page that turns two QuickBooks **Transaction Reports** — a
**Sales / Revenue** report and a **COGS** report — into a clean, 3-tab
**Profit &amp; Loss** report. Everything runs in the browser; no data leaves the
user's machine.

Live at `…/revenue-cogs-gp/` when the repo is served (GitHub Pages, intranet, or
`python3 -m http.server`).

## What it produces

Three tabs on screen (and three sheets in the downloadable workbook):

1. **Revenue transactions** — every line from the Sales report, with a `TOTAL`.
2. **COGS transactions** — every line from the COGS report, with a `TOTAL`.
3. **Gross &amp; Net Profit** — a Profit &amp; Loss statement plus a margin table:

   | Line | Calculation |
   |---|---|
   | Revenue (Sales) | from the Revenue report total |
   | Less: Cost of Goods Sold | from the COGS report total |
   | **Gross Profit** | Sales − COGS |
   | Less: Total Operating Expenses | entered by the user (Step 2) |
   | **Net Profit** | Gross Profit − Operating Expenses |
   | **Gross Profit % (GP %)** | Gross Profit ÷ Sales × 100 |
   | **Net Profit % (NP %)** | Net Profit ÷ Sales × 100 |

   Negatives show in **(parentheses)** and red so a loss is obvious at a glance.

## How to use

1. Open the page.
2. **Step 1** — drop in the **Revenue** and **COGS** `.xlsx` files. The app
   auto-detects the header row, data rows and the report total, and auto-fills
   the company name and period.
3. **Step 2** — enter **Total operating expenses** (everything below the
   gross-profit line). Leave it `0` for a Gross-Profit-only view.
4. **Step 3** — review the three tabs and click **Download** for a formatted
   `.xlsx`. The workbook uses **live formulas** (Gross Profit, Net Profit and the
   two percentages recalculate in Excel), so the numbers are fully auditable.

## Notes on the data

The two source files carry Revenue and COGS only, so **Operating Expenses** is a
user input — that's the only figure needed to complete Net Profit and the Net
Profit margin.

## Layout

```
revenue-cogs-gp/
  index.html     # the page
  styles.css     # GATP styling + tabs + P&L formatting
  app.js         # parser, live P&L, and .xlsx export
../vendor/       # SheetJS (read) + ExcelJS (write), shared with the root tool
```

No build step, no CDN — the libraries are vendored so the page works offline.
