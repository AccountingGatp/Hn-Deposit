/*
 * xlsx-build.js — builds the styled 3-sheet QBO Import workbook with ExcelJS.
 * DOM-free. Takes the ExcelJS constructor as an argument so the same code runs
 * in the browser (vendored ExcelJS) and under Node (npm exceljs) for tests.
 */
(function (root) {
  'use strict';

  const NAVY = 'FF1F3864';
  const ALT = 'FFD9E1F2';
  const WHITE = 'FFFFFFFF';

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

  function buildWorkbook(ExcelJS, data, opts) {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'GATP Bank Deposit → QBO Import';
    wb.created = new Date(2026, 0, 1); // fixed for reproducibility

    const processed = data.processed;
    const notProcessed = data.notProcessed;
    const N = processed.length;
    const M = notProcessed.length;

    // ---------------- Sheet 1: QBO Import ----------------
    const s1 = wb.addWorksheet('QBO Import', {
      views: [{ state: 'frozen', ySplit: 1, showGridLines: false }],
    });
    const h1 = ['SNo.', 'Date', 'Customer name', 'Accounts', 'Description', 'Amount', 'Memo', 'Class'];
    s1.addRow(h1);
    s1.columns = [
      { width: 8 }, { width: 12 }, { width: 26 }, { width: 20 },
      { width: 16 }, { width: 14 }, { width: 16 }, { width: 14 },
    ];

    processed.forEach((p) => {
      const row = s1.addRow([
        p.sno, p.date, p.customer || '', opts.accounts,
        p.description, p.amount, p.memo, opts.klass,
      ]);
      row.getCell(5).numFmt = '@';        // Description as TEXT
      row.getCell(5).value = String(p.description);
      row.getCell(6).numFmt = '#,##0.00'; // Amount
      row.getCell(7).numFmt = '@';        // Memo as TEXT
      row.getCell(7).value = String(p.memo);
      if (!p.customer) {
        row.getCell(3).font = { name: 'Arial', size: 10, color: { argb: 'FFC00000' } };
        row.getCell(3).value = '⚠ SELECT CUSTOMER';
      }
    });

    // Total row (formula, never hardcoded)
    const totalRowNum = N + 2;
    const tRow = s1.getRow(totalRowNum);
    tRow.getCell(5).value = 'TOTAL';
    tRow.getCell(5).font = { name: 'Arial', size: 10, bold: true };
    if (N >= 1) {
      tRow.getCell(6).value = { formula: `SUM(F2:F${N + 1})` };
    } else {
      tRow.getCell(6).value = 0;
    }
    tRow.getCell(6).numFmt = '#,##0.00';
    tRow.getCell(6).font = { name: 'Arial', size: 10, bold: true };
    tRow.getCell(6).border = { top: { style: 'double' }, bottom: { style: 'double' } };

    styleHeader(s1, 1, 8);
    if (N >= 1) styleBody(s1, 2, N + 1, 8);

    // Mapping-assumptions note on the Customer name header
    if (data.assumptions && data.assumptions.length) {
      const noteText = 'Customer-mapping logic applied:\n' +
        data.assumptions.map((a) => `• (${a.count}) ${a.note}`).join('\n') +
        '\n\nNames are chosen only from the Customers list. Rows marked ' +
        '"SELECT CUSTOMER" need a manual pick.';
      s1.getCell('C1').note = { texts: [{ text: noteText }], margins: { insetmode: 'auto' } };
    }

    // ---------------- Sheet 2: Not Processed ----------------
    const s2 = wb.addWorksheet('Not Processed', {
      views: [{ state: 'frozen', ySplit: 1, showGridLines: false }],
    });
    s2.addRow(['Date', 'Bank Description', 'Amount', 'Reason']);
    s2.columns = [{ width: 12 }, { width: 60 }, { width: 14 }, { width: 70 }];
    notProcessed.forEach((n) => {
      const row = s2.addRow([n.date, n.description, n.amount, n.reason]);
      row.getCell(3).numFmt = '#,##0.00';
      row.getCell(4).alignment = { wrapText: true, vertical: 'top' };
      row.getCell(4).note = { texts: [{ text: n.reason }] };
    });
    styleHeader(s2, 1, 4);
    if (M >= 1) styleBody(s2, 2, M + 1, 4);

    // ---------------- Sheet 3: Reconciliation ----------------
    const s3 = wb.addWorksheet('Reconciliation', {
      views: [{ state: 'frozen', ySplit: 1, showGridLines: false }],
    });
    s3.addRow(['Particulars', 'Count', 'Amount']);
    s3.columns = [{ width: 58 }, { width: 12 }, { width: 16 }];

    const qEnd = Math.max(N + 1, 2);
    const npEnd = Math.max(M + 1, 2);
    const QA = `'QBO Import'!A2:A${qEnd}`;
    const QF = `'QBO Import'!F${totalRowNum}`;
    const NA = `'Not Processed'!A2:A${npEnd}`;
    const NC = `'Not Processed'!C2:C${npEnd}`;

    const recRows = [
      [`Total transactions in bank file (${data.dateLabel || 'range'})`,
        { formula: `COUNT(${QA})+COUNTA(${NA})` },
        { formula: `${QF}+SUM(${NC})` }],
      ['Processed — deposits with TRN number',
        { formula: `COUNT(${QA})` },
        { formula: `${QF}` }],
      ['Left / not processed',
        { formula: `COUNTA(${NA})` },
        { formula: `SUM(${NC})` }],
      ['   of which: debits / transfers out',
        { formula: `COUNTIF(${NC},"<0")` },
        { formula: `SUMIF(${NC},"<0")` }],
      ['   of which: deposits without TRN — manual entry needed',
        { formula: `COUNTIF(${NC},">0")` },
        { formula: `SUMIF(${NC},">0")` }],
      ['Check: Processed + Left = Bank file total',
        { formula: `IF(B3+B4-B2=0,"OK","MISMATCH")` },
        { formula: `IF(ROUND(C3+C4-C2,2)=0,"OK","MISMATCH")` }],
    ];
    recRows.forEach((r) => {
      const row = s3.addRow(r);
      row.getCell(3).numFmt = '#,##0.00';
    });
    styleHeader(s3, 1, 3);
    styleBody(s3, 2, 1 + recRows.length, 3);
    // Emphasise the check row
    const checkRow = s3.getRow(1 + recRows.length);
    [1, 2, 3].forEach((c) => { checkRow.getCell(c).font = { name: 'Arial', size: 10, bold: true }; });

    return wb;
  }

  const api = { buildWorkbook };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.BDXlsx = api;
})(typeof self !== 'undefined' ? self : this);
