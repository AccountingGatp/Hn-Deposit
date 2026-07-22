/*
 * xlsx-build.js — builds the styled reconciliation workbook with ExcelJS.
 *
 * Sheets (GATP house style: navy header, alternating fills, thin borders,
 * frozen header, gridlines off, Arial 10, #,##0.00 amounts):
 *   1. Summary        — three-way tie-out; all amounts are live formulas.
 *   2. Batch Recon    — one row per batch: total vs splits vs bank deposit.
 *   3. Transactions   — every split, flagged with the batch it rolls into.
 *   4. Bank Deposits  — every deposit, flagged with the batch it settles.
 *   5. Exceptions     — anything that did not tie (plain-language).
 *
 * DOM-free. Takes the ExcelJS constructor as an argument so the same code runs
 * in the browser (vendored ExcelJS) and under Node (npm exceljs) for tests.
 */
(function (root) {
  'use strict';

  const NAVY = 'FF1F3864';
  const ALT = 'FFD9E1F2';
  const WHITE = 'FFFFFFFF';
  const RED = 'FFC00000';
  const AMBER = 'FF9C6500';
  const MONEY = '#,##0.00';

  function thin() {
    const c = { style: 'thin', color: { argb: 'FFBFBFBF' } };
    return { top: c, left: c, bottom: c, right: c };
  }

  function styleHeader(ws, rowNum, ncols) {
    const row = ws.getRow(rowNum);
    for (let c = 1; c <= ncols; c++) {
      const cell = row.getCell(c);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
      cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: WHITE } };
      cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
      cell.border = thin();
    }
    row.height = 26;
  }

  function styleBody(ws, first, last, ncols) {
    for (let r = first; r <= last; r++) {
      const row = ws.getRow(r);
      const alt = (r - first) % 2 === 1;
      for (let c = 1; c <= ncols; c++) {
        const cell = row.getCell(c);
        if (!cell.font) cell.font = { name: 'Arial', size: 10 };
        else cell.font = Object.assign({ name: 'Arial', size: 10 }, cell.font);
        cell.border = thin();
        if (alt && !cell.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ALT } };
        cell.alignment = Object.assign({ vertical: 'middle' }, cell.alignment);
      }
    }
  }

  function frozen() { return { views: [{ state: 'frozen', ySplit: 1, showGridLines: false }] }; }

  // ---------------------------------------------------------------------------
  function buildWorkbook(ExcelJS, R, opts) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'GATP Deposit Reconciliation';
    wb.created = new Date(2026, 0, 1);
    opts = opts || {};

    const batches = R.batches, txns = R.transactions, deps = R.deposits, exc = R.exceptions;
    const nb = batches.length, nt = txns.length, nd = deps.length;

    // sheet names referenced by formulas
    const S_BR = "'Batch Recon'";
    const S_TX = "'Transactions'";
    const S_BD = "'Bank Deposits'";

    // last data rows (guard empty)
    const brLast = Math.max(nb + 1, 2);
    const txLast = Math.max(nt + 1, 2);
    const bdLast = Math.max(nd + 1, 2);

    // ---------------- Sheet 1: Summary (created first so it is the front tab) ----------------
    buildSummary(wb, R, { S_BR, S_TX, S_BD, brLast, txLast, bdLast });

    // ---------------- Sheet 2: Batch Recon ----------------
    // A Batch # | B MID | C Location | D Close Date | E Settlement Date | F Status
    // G Batch Total | H Txn Count | I Split Sum | J Variance | K Deposited | L Deposit Amt | M Deposit Date | N Recon Status
    const br = wb.addWorksheet('Batch Recon', frozen());
    br.addRow(['Batch #', 'MID', 'Location', 'Close Date', 'Settlement Date', 'Status',
      'Batch Total', 'Txn Count', 'Split Sum', 'Variance', 'Deposited?',
      'Deposit Amount', 'Deposit Date', 'Recon Status']);
    br.columns = [
      { width: 20 }, { width: 12 }, { width: 26 }, { width: 12 }, { width: 14 }, { width: 11 },
      { width: 13 }, { width: 10 }, { width: 12 }, { width: 11 }, { width: 11 },
      { width: 14 }, { width: 12 }, { width: 18 },
    ];
    batches.forEach((b) => {
      const row = br.addRow([
        b.batchNo, b.mid, b.location, b.closeDate, b.settleDate, b.status,
        b.batchTotal, b.txnCount, b.txnSum, b.variance, b.deposited ? 'Yes' : 'No',
        b.deposited ? b.depositAmount : '', b.depositDate, b.reconStatus,
      ]);
      row.getCell(1).numFmt = '@';
      row.getCell(2).numFmt = '@';
      [7, 9, 10, 12].forEach((c) => { row.getCell(c).numFmt = MONEY; });
      if (Math.abs(b.variance) > R.EPS) {
        row.getCell(10).font = { name: 'Arial', size: 10, bold: true, color: { argb: RED } };
        row.getCell(14).font = { name: 'Arial', size: 10, bold: true, color: { argb: RED } };
      } else if (!b.deposited) {
        row.getCell(14).font = { name: 'Arial', size: 10, color: { argb: AMBER } };
      }
    });
    styleHeader(br, 1, 14);
    if (nb) styleBody(br, 2, nb + 1, 14);

    // ---------------- Sheet 3: Transactions ----------------
    // A Date | B Number | C Patient | D Location | E Payment Method | F Type | G Result
    // H Amount | I MID | J Batch # | K Batched? | L Note
    const tx = wb.addWorksheet('Transactions', frozen());
    tx.addRow(['Date', 'Number', 'Patient', 'Location', 'Payment Method', 'Type', 'Result',
      'Amount', 'MID', 'Batch #', 'Batched?', 'Note']);
    tx.columns = [
      { width: 11 }, { width: 12 }, { width: 24 }, { width: 24 }, { width: 18 }, { width: 14 },
      { width: 15 }, { width: 12 }, { width: 12 }, { width: 20 }, { width: 10 }, { width: 52 },
    ];
    txns.forEach((t) => {
      const row = tx.addRow([
        t.date, t.number, t.patient, t.location, t.paymentMethod, t.type, t.result,
        t.amount, t.mid, t.batchNo, t.batched ? 'Yes' : 'No', t.note,
      ]);
      row.getCell(2).numFmt = '@';
      row.getCell(8).numFmt = MONEY;
      row.getCell(9).numFmt = '@';
      row.getCell(10).numFmt = '@';
      row.getCell(12).alignment = { wrapText: true, vertical: 'top' };
      if (!t.batched) row.getCell(11).font = { name: 'Arial', size: 10, color: { argb: AMBER } };
      if (t.note) row.getCell(12).note = { texts: [{ text: t.note }] };
    });
    styleHeader(tx, 1, 12);
    if (nt) styleBody(tx, 2, nt + 1, 12);

    // ---------------- Sheet 4: Bank Deposits ----------------
    // A Date | B Description | C MID | D Received | E Matched Batch # | F Status
    const bd = wb.addWorksheet('Bank Deposits', frozen());
    bd.addRow(['Date', 'Description', 'MID', 'Received', 'Matched Batch #', 'Status']);
    bd.columns = [{ width: 11 }, { width: 60 }, { width: 12 }, { width: 14 }, { width: 20 }, { width: 20 }];
    deps.forEach((d) => {
      const row = bd.addRow([d.date, d.description, d.mid, d.received, d.matchedBatch, d.status]);
      row.getCell(3).numFmt = '@';
      row.getCell(4).numFmt = MONEY;
      row.getCell(5).numFmt = '@';
      if (!d.matchedBatch) row.getCell(6).font = { name: 'Arial', size: 10, bold: true, color: { argb: RED } };
    });
    styleHeader(bd, 1, 6);
    if (nd) styleBody(bd, 2, nd + 1, 6);

    // ---------------- Sheet 5: Exceptions ----------------
    const ex = wb.addWorksheet('Exceptions', frozen());
    ex.addRow(['Type', 'Reference', 'Amount', 'Explanation']);
    ex.columns = [{ width: 30 }, { width: 34 }, { width: 14 }, { width: 90 }];
    if (exc.length) {
      exc.forEach((e) => {
        const row = ex.addRow([e.type, e.reference, e.amount, e.explanation]);
        row.getCell(3).numFmt = MONEY;
        row.getCell(4).alignment = { wrapText: true, vertical: 'top' };
        row.getCell(4).note = { texts: [{ text: e.explanation }] };
      });
      styleBody(ex, 2, exc.length + 1, 4);
    } else {
      const row = ex.addRow(['None', '', '', 'All three files reconcile: every batch total equals the sum of its transaction splits, and every bank deposit matches a batch. Batches shown as "Awaiting deposit" on the Batch Recon sheet are simply a timing difference (submitted but not yet settled in the bank), not an exception.']);
      row.getCell(4).alignment = { wrapText: true, vertical: 'top' };
      styleBody(ex, 2, 2, 4);
    }
    styleHeader(ex, 1, 4);

    return wb;
  }

  // ---------------------------------------------------------------------------
  function buildSummary(wb, R, ref) {
    const { S_BR, S_TX, S_BD, brLast, txLast, bdLast } = ref;
    const sm = wb.addWorksheet('Summary', { views: [{ state: 'frozen', ySplit: 1, showGridLines: false }] });
    sm.addRow(['Particulars', 'Count', 'Amount']);
    sm.columns = [{ width: 62 }, { width: 12 }, { width: 18 }];

    // formula fragments over detail sheets
    const BR_G = `${S_BR}!G2:G${brLast}`;
    const BR_J = `${S_BR}!J2:J${brLast}`;
    const BR_K = `${S_BR}!K2:K${brLast}`;
    const BR_A = `${S_BR}!A2:A${brLast}`;
    const TX_A = `${S_TX}!A2:A${txLast}`;
    const TX_H = `${S_TX}!H2:H${txLast}`;
    const TX_K = `${S_TX}!K2:K${txLast}`;
    const BD_A = `${S_BD}!A2:A${bdLast}`;
    const BD_D = `${S_BD}!D2:D${bdLast}`;
    const BD_E = `${S_BD}!E2:E${bdLast}`;

    const rows = [];
    const section = (t) => rows.push({ section: t });
    const line = (label, count, amount, opt) => rows.push(Object.assign({ label, count, amount }, opt || {}));

    section('FILE TOTALS');
    line('Batch file — settlement batches', { formula: `COUNTA(${BR_A})` }, { formula: `SUM(${BR_G})` });
    line('Transaction export — card transactions (splits)', { formula: `COUNTA(${TX_A})` }, { formula: `SUM(${TX_H})` });
    line('Bank deposit register — deposits', { formula: `COUNTA(${BD_A})` }, { formula: `SUM(${BD_D})` });

    section('RECONCILIATION 1 — TRANSACTIONS → BATCHES');
    line('Transactions that roll into a batch', { formula: `COUNTIF(${TX_K},"Yes")` }, { formula: `SUMIF(${TX_K},"Yes",${TX_H})` });
    line('Transactions with no batch (auth / cash / check / declined)', { formula: `COUNTIF(${TX_K},"No")` }, { formula: `SUMIF(${TX_K},"No",${TX_H})` });
    line('Sum of batch totals', { formula: `COUNTA(${BR_A})` }, { formula: `SUM(${BR_G})` });
    line('CHECK: batch totals = sum of their splits', null, { formula: `IF(ROUND(SUM(${BR_J}),2)=0,"OK","MISMATCH")` }, { check: true });

    section('RECONCILIATION 2 — BATCHES → BANK DEPOSITS');
    line('Batches settled & deposited in the bank', { formula: `COUNTIF(${BR_K},"Yes")` }, { formula: `SUMIF(${BR_K},"Yes",${BR_G})` });
    line('Batches submitted, awaiting deposit (timing)', { formula: `COUNTIF(${BR_K},"No")` }, { formula: `SUMIF(${BR_K},"No",${BR_G})` });
    line('Bank deposits matched to a batch', { formula: `COUNTIF(${BD_E},"?*")` }, { formula: `SUMIF(${BD_E},"?*",${BD_D})` });
    line('Bank deposits with no matching batch', { formula: `COUNTBLANK(${BD_E})` }, { formula: `SUMIF(${BD_E},"",${BD_D})` });
    line('CHECK: deposited-batch total = matched bank-deposit total', null,
      { formula: `IF(ROUND(SUMIF(${BR_K},"Yes",${BR_G})-SUMIF(${BD_E},"?*",${BD_D}),2)=0,"OK","MISMATCH")` }, { check: true });
    line('CHECK: every bank deposit is matched to a batch', null,
      { formula: `IF(COUNTBLANK(${BD_E})=0,"OK","MISMATCH")` }, { check: true });

    // write rows
    rows.forEach((r) => {
      if (r.section) {
        const row = sm.addRow([r.section, '', '']);
        row.getCell(1).font = { name: 'Arial', size: 10, bold: true, color: { argb: WHITE } };
        for (let c = 1; c <= 3; c++) {
          row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2b4a80' } };
          row.getCell(c).border = thin();
        }
        return;
      }
      const row = sm.addRow([r.label, r.count == null ? '' : r.count, r.amount == null ? '' : r.amount]);
      row.getCell(2).alignment = { horizontal: 'right' };
      if (!r.check) row.getCell(3).numFmt = MONEY;
      row.getCell(3).alignment = { horizontal: 'right' };
      row.getCell(1).font = { name: 'Arial', size: 10, bold: !!r.check };
      row.getCell(3).font = { name: 'Arial', size: 10, bold: !!r.check };
      row.getCell(1).border = thin(); row.getCell(2).border = thin(); row.getCell(3).border = thin();
      if (r.check) {
        row.getCell(3).alignment = { horizontal: 'center' };
      }
    });
    styleHeader(sm, 1, 3);
    // light zebra on non-section rows
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].section) continue;
      const rn = i + 2;
      const row = sm.getRow(rn);
      if ((rn % 2) === 1) {
        for (let c = 1; c <= 3; c++) {
          if (!row.getCell(c).fill) row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ALT } };
        }
      }
    }

    // date range caption under the table
    if (R.dateLabel) {
      const cap = sm.addRow(['Date range: ' + R.dateLabel, '', '']);
      cap.getCell(1).font = { name: 'Arial', size: 9, italic: true, color: { argb: 'FF5B6577' } };
    }

    return wb;
  }

  const api = { buildWorkbook };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.ReconXlsx = api;
})(typeof self !== 'undefined' ? self : this);
