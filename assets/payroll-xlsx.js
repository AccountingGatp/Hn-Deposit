/*
 * payroll-xlsx.js — builds the styled Payroll Journal Entry workbook (ExcelJS).
 * DOM-free. Takes the ExcelJS constructor as an argument so the same code runs
 * in the browser (vendored ExcelJS) and under Node (npm exceljs) for tests.
 *
 * Sheets:
 *   1. JE Import    — ACCOUNT | DEBITS | CREDITS | DESCRIPTION | NAME  (matches
 *                     the client's proven QBO import layout; one JE per check date)
 *   2. Summary      — Type × Department totals
 *   3. Journal Entry— human-readable JE, debits first, with a =SUM total row
 *   4. Validation   — debit=credit, identity, exceptions
 */
(function (root) {
  'use strict';

  const NAVY = 'FF1F3864';
  const ALT = 'FFD9E1F2';
  const WHITE = 'FFFFFFFF';
  const RED = 'FFC00000';
  const GREEN = 'FF1A7F45';

  function thin() {
    return {
      top: { style: 'thin', color: { argb: 'FFBFBFBF' } },
      left: { style: 'thin', color: { argb: 'FFBFBFBF' } },
      bottom: { style: 'thin', color: { argb: 'FFBFBFBF' } },
      right: { style: 'thin', color: { argb: 'FFBFBFBF' } },
    };
  }
  function styleHeader(ws, rowNum, ncols) {
    const row = ws.getRow(rowNum);
    for (let c = 1; c <= ncols; c++) {
      const cell = row.getCell(c);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
      cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: WHITE } };
      cell.alignment = { vertical: 'middle', horizontal: 'left' };
      cell.border = thin();
    }
    row.height = 20;
  }
  function styleBody(ws, firstDataRow, lastDataRow, ncols) {
    for (let r = firstDataRow; r <= lastDataRow; r++) {
      const row = ws.getRow(r);
      const alt = (r - firstDataRow) % 2 === 1;
      for (let c = 1; c <= ncols; c++) {
        const cell = row.getCell(c);
        cell.font = { name: 'Arial', size: 10 };
        cell.border = thin();
        if (alt) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ALT } };
        cell.alignment = Object.assign({ vertical: 'middle' }, cell.alignment);
      }
    }
  }
  const MONEY = '#,##0.00';

  function buildWorkbook(ExcelJS, data, opts) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'GATP Payroll Register → QBO Journal Entry';
    wb.created = new Date(2026, 0, 1); // fixed for reproducibility

    const entries = data.entries;      // from PJCore.buildJournalEntries
    const summary = data.summary;      // from PJCore.summarise
    const totals = data.totals;        // from PJCore.totals
    const validation = data.validation; // from PJCore.validate
    const client = (opts && opts.client) || 'Payroll';

    // ---------------- Sheet 1: JE Import ----------------
    const s1 = wb.addWorksheet('JE Import', {
      views: [{ state: 'frozen', ySplit: 1, showGridLines: false }],
    });
    s1.addRow(['ACCOUNT', 'DEBITS', 'CREDITS', 'DESCRIPTION', 'NAME']);
    s1.columns = [{ width: 46 }, { width: 14 }, { width: 14 }, { width: 60 }, { width: 12 }];

    let rowPtr = 1;
    entries.forEach((e, ei) => {
      e.lines.forEach((ln) => {
        const row = s1.addRow([
          ln.account,
          ln.debit != null ? ln.debit : '',
          ln.credit != null ? ln.credit : '',
          e.memo,
          '',
        ]);
        row.getCell(2).numFmt = MONEY;
        row.getCell(3).numFmt = MONEY;
        row.getCell(4).alignment = { wrapText: false, vertical: 'middle' };
        rowPtr++;
      });
      // blank separator row between multiple JEs (keeps each entry visually distinct)
      if (ei < entries.length - 1) { s1.addRow([]); rowPtr++; }
    });
    styleHeader(s1, 1, 5);
    styleBody(s1, 2, rowPtr, 5);

    // ---------------- Sheet 2: Summary ----------------
    const s2 = wb.addWorksheet('Summary', {
      views: [{ state: 'frozen', ySplit: 1, showGridLines: false }],
    });
    s2.addRow(['Type', 'Department', 'Lines', 'Earning', 'Deductions', 'Employee Tax', 'Employer Tax', 'Net Pay']);
    s2.columns = [{ width: 16 }, { width: 13 }, { width: 8 }, { width: 13 }, { width: 13 }, { width: 14 }, { width: 14 }, { width: 13 }];
    summary.forEach((s) => {
      const row = s2.addRow([s.type, s.dept, s.count, s.earning, s.deductions, s.eeTax, s.erTax, s.net]);
      [4, 5, 6, 7, 8].forEach((c) => { row.getCell(c).numFmt = MONEY; });
    });
    const sTotRow = s2.addRow(['TOTAL', '', summary.reduce((a, b) => a + b.count, 0),
      totals.earning, totals.deductions, totals.eeTax, totals.erTax, totals.net]);
    [4, 5, 6, 7, 8].forEach((c) => { sTotRow.getCell(c).numFmt = MONEY; });
    styleHeader(s2, 1, 8);
    styleBody(s2, 2, 1 + summary.length, 8);
    [1, 2, 3, 4, 5, 6, 7, 8].forEach((c) => { sTotRow.getCell(c).font = { name: 'Arial', size: 10, bold: true }; });
    sTotRow.eachCell((cell) => { cell.border = thin(); });

    // ---------------- Sheet 3: Journal Entry (human readable) ----------------
    const s3 = wb.addWorksheet('Journal Entry', {
      views: [{ state: 'frozen', ySplit: 2, showGridLines: false }],
    });
    s3.columns = [{ width: 48 }, { width: 15 }, { width: 15 }, { width: 60 }];
    let r3 = 1;
    entries.forEach((e) => {
      const title = s3.getRow(r3);
      title.getCell(1).value = e.memo;
      title.getCell(1).font = { name: 'Arial', size: 11, bold: true, color: { argb: NAVY } };
      s3.mergeCells(r3, 1, r3, 4);
      r3++;
      s3.getRow(r3).values = ['Account', 'Debit', 'Credit', 'Memo'];
      styleHeader(s3, r3, 4);
      const headerRow = r3;
      r3++;
      const firstData = r3;
      e.lines.forEach((ln) => {
        const row = s3.getRow(r3);
        row.values = [ln.account, ln.debit != null ? ln.debit : null, ln.credit != null ? ln.credit : null, e.memo];
        row.getCell(2).numFmt = MONEY;
        row.getCell(3).numFmt = MONEY;
        r3++;
      });
      const lastData = r3 - 1;
      // total row with formulas (never hardcoded)
      const tot = s3.getRow(r3);
      tot.getCell(1).value = 'TOTAL';
      tot.getCell(2).value = { formula: `SUM(B${firstData}:B${lastData})` };
      tot.getCell(3).value = { formula: `SUM(C${firstData}:C${lastData})` };
      tot.getCell(2).numFmt = MONEY; tot.getCell(3).numFmt = MONEY;
      [1, 2, 3].forEach((c) => { tot.getCell(c).font = { name: 'Arial', size: 10, bold: true }; });
      tot.getCell(2).border = { top: { style: 'double' }, bottom: { style: 'double' } };
      tot.getCell(3).border = { top: { style: 'double' }, bottom: { style: 'double' } };
      styleBody(s3, firstData, lastData, 4);
      // wrap memo
      for (let rr = firstData; rr <= lastData; rr++) s3.getRow(rr).getCell(4).alignment = { vertical: 'middle' };
      void headerRow;
      r3 += 2; // gap before next entry
    });

    // ---------------- Sheet 4: Validation ----------------
    const s4 = wb.addWorksheet('Validation', {
      views: [{ state: 'frozen', ySplit: 1, showGridLines: false }],
    });
    s4.addRow(['Check', 'Value', 'Status']);
    s4.columns = [{ width: 62 }, { width: 26 }, { width: 12 }];
    const rows = [];
    validation.checks.forEach((c) => rows.push([c.label, c.value, c.pass ? 'PASS' : 'FAIL', c.pass]));
    rows.push(['Register — Earning total', totals.earning, '', null]);
    rows.push(['Register — Deductions total', totals.deductions, '', null]);
    rows.push(['Register — Employee Tax total', totals.eeTax, '', null]);
    rows.push(['Register — Employer Tax total', totals.erTax, '', null]);
    rows.push(['Register — Net Pay total', totals.net, '', null]);
    rows.push(['Rows dropped (blank presence flag)', data.dropped, '', null]);
    rows.push(['Employees', data.employeeCount, '', null]);
    rows.push(['Check dates', data.checkDates.join(', '), '', null]);
    rows.forEach((r) => {
      const row = s4.addRow([r[0], r[1], r[2]]);
      if (typeof r[1] === 'number') row.getCell(2).numFmt = MONEY;
      if (r[3] === true) row.getCell(3).font = { name: 'Arial', size: 10, bold: true, color: { argb: GREEN } };
      if (r[3] === false) row.getCell(3).font = { name: 'Arial', size: 10, bold: true, color: { argb: RED } };
    });
    styleHeader(s4, 1, 3);
    styleBody(s4, 2, 1 + rows.length, 3);

    // ---------------- Sheet 5: Exceptions (only if any) ----------------
    if (validation.unclassified.length || validation.negatives.length) {
      const s5 = wb.addWorksheet('Exceptions', {
        views: [{ state: 'frozen', ySplit: 1, showGridLines: false }],
      });
      s5.addRow(['Type', 'Employee', 'Description', 'Earning', 'Deductions', 'EE Tax', 'ER Tax']);
      s5.columns = [{ width: 16 }, { width: 22 }, { width: 16 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }];
      const add = (label, r) => {
        const row = s5.addRow([label, r.name, r.desc, r.earning, r.deductions, r.eeTax, r.erTax]);
        [4, 5, 6, 7].forEach((c) => { row.getCell(c).numFmt = MONEY; });
      };
      validation.unclassified.forEach((r) => add('Unclassified', r));
      validation.negatives.forEach((r) => add('Negative amount', r));
      styleHeader(s5, 1, 7);
      styleBody(s5, 2, 1 + validation.unclassified.length + validation.negatives.length, 7);
    }

    void client;
    return wb;
  }

  const api = { buildWorkbook };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.PJXlsx = api;
})(typeof self !== 'undefined' ? self : this);
