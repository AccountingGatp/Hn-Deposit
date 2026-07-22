/*
 * ms-build.js — builds the styled 5-sheet Monthly Services → QBO workbook.
 *
 * Sheets (exact order required by the prompt):
 *   1. Sales - Non-340B
 *   2. Sales - 340B
 *   3. Credit Note - Non-340B
 *   4. Credit Note - 340B
 *   5. Journal Ledger
 *
 * DOM-free. Takes the ExcelJS constructor as an argument so the same code runs
 * in the browser (vendored ExcelJS) and under Node (npm exceljs) for tests.
 * GATP house style: Arial 10, navy #1F3864 header, per-invoice #D9E1F2 banding,
 * thin gridlines, frozen header, amounts 0.00, dates mm/dd/yyyy, invoice # text.
 */
(function (root) {
  'use strict';

  var NAVY = 'FF1F3864';
  var BAND = 'FFD9E1F2';
  var WHITE = 'FFFFFFFF';

  function thin() {
    return {
      top: { style: 'thin', color: { argb: 'FFBFBFBF' } },
      left: { style: 'thin', color: { argb: 'FFBFBFBF' } },
      bottom: { style: 'thin', color: { argb: 'FFBFBFBF' } },
      right: { style: 'thin', color: { argb: 'FFBFBFBF' } }
    };
  }

  function styleHeader(ws, ncols) {
    var row = ws.getRow(1);
    for (var c = 1; c <= ncols; c++) {
      var cell = row.getCell(c);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
      cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: WHITE } };
      cell.alignment = { vertical: 'middle', horizontal: 'left' };
      cell.border = thin();
    }
    row.height = 20;
  }

  // The 10 line-item columns shared by all four Sales / Credit Note sheets.
  var LINE_HEADERS = [
    'Invoice Number', 'Customer', 'Invoice Date', 'Memo',
    'Message displayed on invoice', 'Product/Service', 'Product/Service Description',
    'Product/Service Amount', 'Product/Service Taxable', 'Location Name'
  ];
  var LINE_WIDTHS = [18, 26, 13, 18, 22, 46, 20, 16, 12, 22];

  // Build one Sales / Credit Note sheet. Two rows per invoice, banded by invoice.
  function addLineSheet(wb, name, invoices) {
    var ws = wb.addWorksheet(name, {
      views: [{ state: 'frozen', ySplit: 1, showGridLines: false }]
    });
    ws.addRow(LINE_HEADERS);
    ws.columns = LINE_WIDTHS.map(function (w) { return { width: w }; });

    var rownum = 1;
    invoices.forEach(function (inv, i) {
      var band = i % 2 === 1; // alternate the fill per invoice (both its lines)
      // original line, then Adjustment line
      var lines = [
        { ps: inv.psName, amt: inv.origAmt },
        { ps: inv.adjName, amt: inv.adjAmt }
      ];
      lines.forEach(function (ln) {
        rownum++;
        var row = ws.addRow([
          inv.invNo, inv.customer, inv.date, inv.invNo, inv.invNo,
          ln.ps, inv.invNo, ln.amt, false, inv.location
        ]);
        // Invoice Number / Memo / Message / Description as TEXT
        [1, 4, 5, 7].forEach(function (ci) {
          row.getCell(ci).numFmt = '@';
          row.getCell(ci).value = String(inv.invNo);
        });
        row.getCell(3).numFmt = 'mm/dd/yyyy';
        row.getCell(8).numFmt = '0.00';
        // styling
        for (var c = 1; c <= LINE_HEADERS.length; c++) {
          var cell = row.getCell(c);
          cell.font = { name: 'Arial', size: 10 };
          cell.border = thin();
          if (band) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BAND } };
          cell.alignment = Object.assign({ vertical: 'middle' }, cell.alignment);
        }
      });
    });

    // Total row (formula over the Amount column) for a quick eyeball.
    if (invoices.length) {
      var last = rownum;
      var tRow = ws.getRow(last + 1);
      tRow.getCell(7).value = 'TOTAL';
      tRow.getCell(7).font = { name: 'Arial', size: 10, bold: true };
      tRow.getCell(8).value = { formula: 'SUM(H2:H' + last + ')' };
      tRow.getCell(8).numFmt = '0.00';
      tRow.getCell(8).font = { name: 'Arial', size: 10, bold: true };
      tRow.getCell(8).border = { top: { style: 'double' }, bottom: { style: 'double' } };
    }

    styleHeader(ws, LINE_HEADERS.length);
    return ws;
  }

  var J_HEADERS = ['Journal No', 'Date', 'Memo', 'Account', 'Debit', 'Credit', 'Description', 'Name', 'Class'];
  var J_WIDTHS = [18, 13, 18, 58, 14, 14, 18, 26, 22];

  function addJournalSheet(wb, name, journal) {
    var ws = wb.addWorksheet(name, {
      views: [{ state: 'frozen', ySplit: 1, showGridLines: false }]
    });
    ws.addRow(J_HEADERS);
    ws.columns = J_WIDTHS.map(function (w) { return { width: w }; });

    var rownum = 1;
    journal.forEach(function (l, i) {
      var band = Math.floor(i / 3) % 2 === 1; // band per 3-line entry
      rownum++;
      var row = ws.addRow([
        l.journalNo, l.date, l.memo, l.account,
        l.debit === null ? null : l.debit,
        l.credit === null ? null : l.credit,
        l.description, l.name, l.klass
      ]);
      [1, 3, 7].forEach(function (ci) { row.getCell(ci).numFmt = '@'; row.getCell(ci).value = String(l.journalNo); });
      row.getCell(2).numFmt = 'mm/dd/yyyy';
      row.getCell(5).numFmt = '0.00';
      row.getCell(6).numFmt = '0.00';
      for (var c = 1; c <= J_HEADERS.length; c++) {
        var cell = row.getCell(c);
        cell.font = { name: 'Arial', size: 10 };
        cell.border = thin();
        if (band) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BAND } };
        cell.alignment = Object.assign({ vertical: 'middle' }, cell.alignment);
      }
    });

    if (journal.length) {
      var last = rownum;
      var tRow = ws.getRow(last + 1);
      tRow.getCell(4).value = 'TOTAL';
      tRow.getCell(4).font = { name: 'Arial', size: 10, bold: true };
      tRow.getCell(5).value = { formula: 'SUM(E2:E' + last + ')' };
      tRow.getCell(6).value = { formula: 'SUM(F2:F' + last + ')' };
      [5, 6].forEach(function (c) {
        tRow.getCell(c).numFmt = '0.00';
        tRow.getCell(c).font = { name: 'Arial', size: 10, bold: true };
        tRow.getCell(c).border = { top: { style: 'double' }, bottom: { style: 'double' } };
      });
      // balance check cell
      tRow.getCell(7).value = { formula: 'IF(ROUND(E' + (last + 1) + '-F' + (last + 1) + ',2)=0,"BALANCED","OUT")' };
      tRow.getCell(7).font = { name: 'Arial', size: 10, bold: true };
    }

    styleHeader(ws, J_HEADERS.length);
    return ws;
  }

  function buildWorkbook(ExcelJS, data) {
    var wb = new ExcelJS.Workbook();
    wb.creator = 'GATP Monthly Services → QBO Import (340B Split)';
    wb.created = new Date(2026, 0, 1); // fixed for reproducibility

    var b = data.buckets;
    addLineSheet(wb, 'Sales - Non-340B', b.salesNon);
    addLineSheet(wb, 'Sales - 340B', b.sales340);
    addLineSheet(wb, 'Credit Note - Non-340B', b.cnNon);
    addLineSheet(wb, 'Credit Note - 340B', b.cn340);
    addJournalSheet(wb, 'Journal Ledger', data.journal);
    return wb;
  }

  var api = { buildWorkbook: buildWorkbook };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.MSXlsx = api;
})(typeof self !== 'undefined' ? self : this);
