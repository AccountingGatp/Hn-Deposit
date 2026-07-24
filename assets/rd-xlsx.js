/*
 * rd-xlsx.js — builds the Remote Deposit workbook with ExcelJS. DOM-free; takes
 * the ExcelJS constructor as an argument so it runs in the browser (vendored
 * ExcelJS) and under Node (npm exceljs) for tests.
 *
 * Sheet 1 "Remote Deposit" is the import file: exactly three columns —
 * Date | Description | Amount — with a =SUM total row.
 * Sheet 2 "Review" keeps the audit trail (source file, page, review flag) so the
 * three-column import sheet stays clean.
 */
(function (root) {
  'use strict';

  var NAVY = 'FF1F3864', ALT = 'FFD9E1F2', WHITE = 'FFFFFFFF', RED = 'FFC00000';

  function thin() {
    var b = { style: 'thin', color: { argb: 'FFBFBFBF' } };
    return { top: b, left: b, bottom: b, right: b };
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

  function styleBody(ws, first, last, ncols) {
    for (var r = first; r <= last; r++) {
      var row = ws.getRow(r);
      var alt = (r - first) % 2 === 1;
      for (var c = 1; c <= ncols; c++) {
        var cell = row.getCell(c);
        cell.font = cell.font || { name: 'Arial', size: 10 };
        cell.border = thin();
        if (alt && !cell.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ALT } };
        cell.alignment = Object.assign({ vertical: 'middle' }, cell.alignment);
      }
    }
  }

  function dateSerial(d) {
    // Excel serial date (1900 system) from {yr,mo,day} — avoids timezone drift.
    if (!d) return null;
    var utc = Date.UTC(d.yr, d.mo - 1, d.day);
    var epoch = Date.UTC(1899, 11, 30);
    return Math.round((utc - epoch) / 86400000);
  }

  function buildWorkbook(ExcelJS, rows, opts) {
    opts = opts || {};
    var wb = new ExcelJS.Workbook();
    wb.creator = 'GATP Remote Deposit Check Extractor';
    wb.created = new Date(2026, 0, 1);

    var N = rows.length;

    // ---------------- Sheet 1: Remote Deposit (3 columns) ----------------
    var s1 = wb.addWorksheet('Remote Deposit', {
      views: [{ state: 'frozen', ySplit: 1, showGridLines: false }]
    });
    s1.addRow(['Date', 'Description', 'Amount']);
    s1.columns = [{ width: 14 }, { width: 22 }, { width: 16 }];

    rows.forEach(function (r) {
      var row = s1.addRow([null, null, null]);
      // Date as a real date value when we have one; else keep the raw string.
      var serial = dateSerial(r.dateObj);
      if (serial != null) {
        row.getCell(1).value = serial;
        row.getCell(1).numFmt = 'mm/dd/yyyy';
      } else {
        row.getCell(1).value = r.date || '';
      }
      // Description = check number, stored as TEXT so leading zeros survive.
      row.getCell(2).numFmt = '@';
      row.getCell(2).value = String(r.checkNo || '');
      // Amount, numeric.
      row.getCell(3).numFmt = '#,##0.00';
      row.getCell(3).value = (r.amount == null ? null : r.amount);
      if (r.flagStatus && r.flagStatus !== 'OK') {
        row.getCell(2).font = { name: 'Arial', size: 10, color: { argb: RED } };
        if (!r.checkNo) row.getCell(2).value = 'REVIEW';
      }
    });

    var totalRowNum = N + 2;
    var tRow = s1.getRow(totalRowNum);
    tRow.getCell(2).value = 'TOTAL';
    tRow.getCell(2).font = { name: 'Arial', size: 10, bold: true };
    tRow.getCell(3).value = N >= 1 ? { formula: 'SUM(C2:C' + (N + 1) + ')' } : 0;
    tRow.getCell(3).numFmt = '#,##0.00';
    tRow.getCell(3).font = { name: 'Arial', size: 10, bold: true };
    tRow.getCell(3).border = { top: { style: 'double' }, bottom: { style: 'double' } };

    styleHeader(s1, 1, 3);
    if (N >= 1) styleBody(s1, 2, N + 1, 3);

    // ---------------- Sheet 2: Review (audit trail) ----------------
    var s2 = wb.addWorksheet('Review', {
      views: [{ state: 'frozen', ySplit: 1, showGridLines: false }]
    });
    s2.addRow(['Date', 'Description (Check #)', 'Amount', 'Source file', 'Page', 'Review Flag']);
    s2.columns = [
      { width: 14 }, { width: 22 }, { width: 14 },
      { width: 52 }, { width: 8 }, { width: 42 }
    ];
    rows.forEach(function (r) {
      var row = s2.addRow([
        r.date || '', String(r.checkNo || ''), (r.amount == null ? null : r.amount),
        r.file || '', (r.pageIndex == null ? '' : r.pageIndex + 1),
        r.flagLabel || 'OK'
      ]);
      row.getCell(2).numFmt = '@';
      row.getCell(3).numFmt = '#,##0.00';
      row.getCell(6).alignment = { wrapText: true, vertical: 'middle' };
      if (r.flagStatus && r.flagStatus !== 'OK') {
        row.getCell(6).font = { name: 'Arial', size: 10, bold: true, color: { argb: RED } };
      } else {
        row.getCell(6).font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF1A7F45' } };
      }
    });
    styleHeader(s2, 1, 6);
    if (N >= 1) styleBody(s2, 2, N + 1, 6);

    return wb;
  }

  var api = { buildWorkbook: buildWorkbook, _dateSerial: dateSerial };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.RDXlsx = api;
})(typeof self !== 'undefined' ? self : this);
