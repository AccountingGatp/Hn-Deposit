/*
 * bills-build.js — builds the styled QBO Bills Import workbook with ExcelJS.
 *
 * DOM-free. Takes the ExcelJS constructor as an argument so the same code runs
 * in the browser (vendored ExcelJS) and under Node (npm exceljs) for tests.
 *
 * Three sheets, GATP house style:
 *   1. QBO Import   — the columns a QuickBooks Online bill import expects, one
 *                     row per invoice, with a =SUM total on Amount.
 *   2. Summary      — the SOP's brief review table (Invoice #, Date, Amount,
 *                     Ship-To/Customer #, Class, Category, Vendor, Matched?).
 *   3. Exceptions   — every row carrying a [REVIEW NEEDED] / unmatched flag.
 */
(function (root) {
  'use strict';

  var NAVY = 'FF1F3864';
  var ALT = 'FFD9E1F2';
  var WHITE = 'FFFFFFFF';
  var RED = 'FFC00000';

  function thin() {
    return {
      top: { style: 'thin', color: { argb: 'FFBFBFBF' } },
      left: { style: 'thin', color: { argb: 'FFBFBFBF' } },
      bottom: { style: 'thin', color: { argb: 'FFBFBFBF' } },
      right: { style: 'thin', color: { argb: 'FFBFBFBF' } }
    };
  }

  function styleHeader(ws, rowNum, ncols) {
    var row = ws.getRow(rowNum);
    for (var c = 1; c <= ncols; c++) {
      var cell = row.getCell(c);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
      cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: WHITE } };
      cell.alignment = { vertical: 'middle', horizontal: 'left' };
      cell.border = thin();
    }
    row.height = 20;
  }

  function styleBody(ws, firstDataRow, lastDataRow, ncols) {
    for (var r = firstDataRow; r <= lastDataRow; r++) {
      var row = ws.getRow(r);
      var alt = (r - firstDataRow) % 2 === 1;
      for (var c = 1; c <= ncols; c++) {
        var cell = row.getCell(c);
        if (!cell.font) cell.font = { name: 'Arial', size: 10 };
        cell.border = thin();
        if (alt) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ALT } };
        cell.alignment = Object.assign({ vertical: 'middle' }, cell.alignment);
      }
    }
  }

  var REVIEW = '[REVIEW NEEDED]';

  function isReview(v) { return v === REVIEW || v == null || v === ''; }

  function buildWorkbook(ExcelJS, rows, opts) {
    opts = opts || {};
    var wb = new ExcelJS.Workbook();
    wb.creator = 'GATP AP Vendor Bills → QBO Import';
    wb.created = new Date(2026, 0, 1); // fixed for reproducibility
    var N = rows.length;

    /* ---------------- Sheet 1: QBO Import ---------------- */
    var s1 = wb.addWorksheet('QBO Import', {
      views: [{ state: 'frozen', ySplit: 1, showGridLines: false }]
    });
    var h1 = ['Bill No.', 'Vendor', 'Bill Date', 'Due Date', 'Terms',
              'Account', 'Amount', 'Memo/Description', 'Class'];
    s1.addRow(h1);
    s1.columns = [
      { width: 16 }, { width: 22 }, { width: 12 }, { width: 12 }, { width: 10 },
      { width: 26 }, { width: 14 }, { width: 34 }, { width: 16 }
    ];

    rows.forEach(function (p) {
      var row = s1.addRow([
        p.billNo, p.vendor, p.billDate, p.dueDate, p.terms,
        p.account, (p.amount == null ? 0 : p.amount), p.memo, p.klass
      ]);
      row.getCell(1).numFmt = '@';           // Bill No. as TEXT (keep leading digits)
      row.getCell(1).value = String(p.billNo);
      row.getCell(7).numFmt = '#,##0.00';    // Amount
      // Red-flag any [REVIEW NEEDED] cells so they are impossible to miss.
      [1, 2, 3, 6].forEach(function (ci) {
        if (isReview(row.getCell(ci).value)) {
          row.getCell(ci).font = { name: 'Arial', size: 10, bold: true, color: { argb: RED } };
          if (row.getCell(ci).value == null || row.getCell(ci).value === '') row.getCell(ci).value = REVIEW;
        }
      });
      if (p.amount == null) {
        row.getCell(7).font = { name: 'Arial', size: 10, bold: true, color: { argb: RED } };
      }
    });

    // Total row (formula, never hardcoded).
    var totalRowNum = N + 2;
    var tRow = s1.getRow(totalRowNum);
    tRow.getCell(6).value = 'TOTAL';
    tRow.getCell(6).font = { name: 'Arial', size: 10, bold: true };
    tRow.getCell(6).alignment = { horizontal: 'right' };
    tRow.getCell(7).value = N >= 1 ? { formula: 'SUM(G2:G' + (N + 1) + ')' } : 0;
    tRow.getCell(7).numFmt = '#,##0.00';
    tRow.getCell(7).font = { name: 'Arial', size: 10, bold: true };
    tRow.getCell(7).border = { top: { style: 'double' }, bottom: { style: 'double' } };

    styleHeader(s1, 1, 9);
    if (N >= 1) styleBody(s1, 2, N + 1, 9);

    s1.getCell('F1').note = {
      texts: [{ text:
        'Account = the expense Category pulled from the GL for this Ship-To / ' +
        'Customer number. Class also comes from the GL. Rows shown in red need a ' +
        'manual fix before importing to QuickBooks Online.' }]
    };

    /* ---------------- Sheet 2: Summary ---------------- */
    var s2 = wb.addWorksheet('Summary', {
      views: [{ state: 'frozen', ySplit: 1, showGridLines: false }]
    });
    s2.addRow(['Source File', 'Vendor', 'Invoice #', 'Invoice Date', 'Amount',
               'Ship-To / Customer #', 'Class', 'Category', 'Matched in GL?']);
    s2.columns = [
      { width: 26 }, { width: 20 }, { width: 15 }, { width: 13 }, { width: 14 },
      { width: 20 }, { width: 16 }, { width: 24 }, { width: 14 }
    ];
    rows.forEach(function (p) {
      var row = s2.addRow([
        p.sourceName || '', p.vendorLabel || p.vendor, p.billNo, p.billDate,
        (p.amount == null ? 0 : p.amount), p.key, p.klass, p.account,
        p.matched ? 'Yes' : 'No'
      ]);
      row.getCell(5).numFmt = '#,##0.00';
      row.getCell(3).numFmt = '@';
      row.getCell(3).value = String(p.billNo);
      var mCell = row.getCell(9);
      mCell.font = { name: 'Arial', size: 10, bold: true, color: { argb: p.matched ? 'FF1A7F45' : RED } };
    });
    var sumTotRow = s2.getRow(N + 2);
    sumTotRow.getCell(4).value = 'TOTAL';
    sumTotRow.getCell(4).font = { name: 'Arial', size: 10, bold: true };
    sumTotRow.getCell(4).alignment = { horizontal: 'right' };
    sumTotRow.getCell(5).value = N >= 1 ? { formula: 'SUM(E2:E' + (N + 1) + ')' } : 0;
    sumTotRow.getCell(5).numFmt = '#,##0.00';
    sumTotRow.getCell(5).font = { name: 'Arial', size: 10, bold: true };
    sumTotRow.getCell(5).border = { top: { style: 'double' }, bottom: { style: 'double' } };
    styleHeader(s2, 1, 9);
    if (N >= 1) styleBody(s2, 2, N + 1, 9);

    /* ---------------- Sheet 3: Exceptions ---------------- */
    var exceptions = rows.filter(function (p) { return p.needsReview; });
    var s3 = wb.addWorksheet('Exceptions', {
      views: [{ state: 'frozen', ySplit: 1, showGridLines: false }]
    });
    s3.addRow(['Source File', 'Vendor', 'Invoice #', 'Ship-To / Customer #', 'What needs review']);
    s3.columns = [{ width: 28 }, { width: 20 }, { width: 15 }, { width: 22 }, { width: 70 }];
    if (exceptions.length) {
      exceptions.forEach(function (p) {
        var reason = (p.flags && p.flags.length) ? p.flags.join('; ') : 'Missing required field';
        var row = s3.addRow([p.sourceName || '', p.vendorLabel || p.vendor,
                             String(p.billNo), p.key, reason]);
        row.getCell(3).numFmt = '@';
        row.getCell(5).alignment = { wrapText: true, vertical: 'top' };
        row.getCell(5).note = { texts: [{ text: reason }] };
      });
    } else {
      var okRow = s3.addRow(['', '', '', '', 'No exceptions — every invoice was extracted and matched in the GL.']);
      okRow.getCell(5).font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF1A7F45' } };
    }
    styleHeader(s3, 1, 5);
    var exN = Math.max(exceptions.length, 1);
    styleBody(s3, 2, exN + 1, 5);

    return wb;
  }

  var api = { buildWorkbook: buildWorkbook };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.APBillsXlsx = api;
})(typeof self !== 'undefined' ? self : this);
