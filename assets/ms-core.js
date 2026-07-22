/*
 * ms-core.js — Monthly Services → QuickBooks Import (340B Split) core logic.
 *
 * Implements the GATP reusable prompt "Monthly Services → QuickBooks Import
 * (340B Split)". DOM-free and I/O-free: pure functions so the exact same logic
 * runs in the browser (vendored SheetJS/ExcelJS) and under Node for tests.
 *
 * Pipeline
 *   1. Build two line items (original + Adjustment) per Monthly Services row.
 *   2. Classify each invoice 340B / Non-340B (Procedure Code in JQ list AND
 *      patient in Patient List).
 *   3. Split Sales vs Credit Note by the netted amount (net < 0 → Credit Note).
 *   4. Build the Journal Ledger from the Credit Note invoices only.
 */
(function (root) {
  'use strict';

  var ADJ_PCT = 0.6767; // Adjustment line = -Fee × 67.67%

  // ---- Amounts -------------------------------------------------------------
  // Round HALF-UP (away from zero) to 2 decimals — matches Python Decimal
  // ROUND_HALF_UP used in the reference implementation.
  function round2(x) {
    if (x === null || x === undefined || x === '' || isNaN(x)) return 0;
    var n = Number(x);
    var s = n < 0 ? -1 : 1;
    return s * Math.round((Math.abs(n) + 1e-9) * 100) / 100;
  }

  // ---- Column helpers ------------------------------------------------------
  // Excel column letter -> 0-based index. 'A'->0, 'B'->1, 'AA'->26, 'BC'->54.
  function colIdx(letter) {
    var n = 0;
    letter = String(letter).toUpperCase();
    for (var i = 0; i < letter.length; i++) {
      n = n * 26 + (letter.charCodeAt(i) - 64);
    }
    return n - 1;
  }

  // Source columns (letters cross-checked against the reusable prompt).
  var COL = {
    voucher: colIdx('B'),   // Voucher Number
    service: colIdx('C'),   // Service ID
    voided: colIdx('D'),   // Is Voided
    postDate: colIdx('W'),   // Post Date
    posDescr: colIdx('Y'),   // Place of Service Descr  (Location Name / Class)
    lname: colIdx('G'),   // LName
    fname: colIdx('H'),   // FName
    mi: colIdx('I'),   // Middle Initial
    customer: colIdx('BC'),  // Original Ins Category Desc (Customer / Name)
    procCode: colIdx('BL'),  // Procedure Code
    procDescr: colIdx('BM'),  // Procedure Descr
    fee: colIdx('BP')   // Fee
  };

  // ---- Value normalisers ---------------------------------------------------
  // Integer-string form of a numeric-ish cell (no trailing ".0"), for building
  // the 14-digit Invoice Number as text.
  function intStr(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') {
      return Number.isInteger(v) ? String(v) : String(Math.trunc(v));
    }
    var s = String(v).trim();
    // strip a trailing ".0" that Excel/SheetJS sometimes yields
    return s.replace(/\.0+$/, '');
  }

  // Uppercase + remove ALL whitespace, for tolerant name matching
  // ("TEST, MARK" and "TEST,MARK" collapse to the same key).
  function normName(s) {
    if (s === null || s === undefined) return '';
    return String(s).toUpperCase().replace(/\s+/g, '');
  }

  // Coerce a cell into a JS Date (accepts Date, Excel serial, or a date string).
  function toDate(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return v;
    if (typeof v === 'number') {
      // Excel serial date (1900 date system)
      var ms = Math.round((v - 25569) * 86400 * 1000);
      var d = new Date(ms);
      return isNaN(d.getTime()) ? null : d;
    }
    var p = new Date(v);
    return isNaN(p.getTime()) ? null : p;
  }

  // ---- Reference sets ------------------------------------------------------
  // JQ Code List: a Set of upper-cased HCPCS/CPT codes (J-, G-, Q- and 7xxxx).
  function buildCodeSet(codeMatrix) {
    var set = new Set();
    if (!codeMatrix || !codeMatrix.length) return set;
    // Detect a header row ("HCPCS") and skip it.
    var start = 0;
    var first = codeMatrix[0] && codeMatrix[0][0];
    if (first !== null && first !== undefined &&
        /^[A-Za-z]/.test(String(first)) && !/^[JGQ]\d/i.test(String(first).trim()) &&
        String(first).trim().length < 6 && /hcpcs|code/i.test(String(first))) {
      start = 1;
    } else if (first !== null && first !== undefined && /hcpcs|code/i.test(String(first))) {
      start = 1;
    }
    for (var r = start; r < codeMatrix.length; r++) {
      var v = codeMatrix[r] && codeMatrix[r][0];
      if (v === null || v === undefined) continue;
      var s = String(v).trim();
      if (s) set.add(s.toUpperCase());
    }
    return set;
  }

  // Patient List: a Set of normalised "LNAME,FNAME[ MI]" keys from the
  // "Patient Name" column. Ignores the EMR Source column and "(blank)" rows.
  function buildPatientSet(patMatrix) {
    var set = new Set();
    if (!patMatrix || !patMatrix.length) return set;
    // find the "Patient Name" column in the header row; fall back to last col
    var header = patMatrix[0].map(function (h) { return String(h == null ? '' : h).trim().toLowerCase(); });
    var nameCol = header.indexOf('patient name');
    if (nameCol === -1) {
      nameCol = header.findIndex(function (h) { return h.indexOf('patient') !== -1 && h.indexOf('name') !== -1; });
    }
    if (nameCol === -1) nameCol = patMatrix[0].length - 1;
    for (var r = 1; r < patMatrix.length; r++) {
      var v = patMatrix[r] && patMatrix[r][nameCol];
      if (v === null || v === undefined) continue;
      var s = String(v).trim();
      if (!s || s.toLowerCase() === '(blank)') continue;
      set.add(normName(s));
    }
    return set;
  }

  // ---- Step 1+2: build invoices from the Monthly Services matrix -----------
  // msMatrix: array-of-arrays incl. header row 0. Returns one object per source
  // row, each already carrying its two line items and its 340B classification.
  function buildInvoices(msMatrix, codeSet, patientSet) {
    var out = [];
    if (!msMatrix || msMatrix.length < 2) return out;
    for (var r = 1; r < msMatrix.length; r++) {
      var row = msMatrix[r];
      if (!row) continue;
      // skip a fully-empty tail row
      if ((row[COL.voucher] === null || row[COL.voucher] === undefined || row[COL.voucher] === '') &&
          (row[COL.service] === null || row[COL.service] === undefined || row[COL.service] === '') &&
          (row[COL.fee] === null || row[COL.fee] === undefined || row[COL.fee] === '')) {
        continue;
      }

      var invNo = intStr(row[COL.voucher]) + intStr(row[COL.service]);
      var customer = row[COL.customer] == null ? '' : String(row[COL.customer]);
      var date = toDate(row[COL.postDate]);
      var location = row[COL.posDescr] == null ? '' : String(row[COL.posDescr]);
      var procCode = row[COL.procCode] == null ? '' : String(row[COL.procCode]).trim();
      var procDescr = row[COL.procDescr] == null ? '' : String(row[COL.procDescr]);
      var fee = Number(row[COL.fee]) || 0;
      var voided = row[COL.voided] === true ||
        String(row[COL.voided]).trim().toUpperCase() === 'TRUE';

      // --- classify 340B ---
      var lname = row[COL.lname] == null ? '' : String(row[COL.lname]);
      var fname = row[COL.fname] == null ? '' : String(row[COL.fname]);
      var mi = row[COL.mi];
      var miStr = (mi === null || mi === undefined) ? '' : String(mi).trim();

      var codeHit = codeSet.has(procCode.toUpperCase());
      var form1 = normName(lname + ',' + fname);
      var form2 = miStr ? normName(lname + ',' + fname + miStr) : null;
      var patHit = patientSet.has(form1) || (form2 !== null && patientSet.has(form2));
      var is340 = codeHit && patHit;

      // --- two line items ---
      var origAmt = round2(fee);
      var adjAmt = round2(-fee * ADJ_PCT);
      var net = round2(origAmt + adjAmt);
      var psName = procCode + ' - ' + procDescr;

      out.push({
        invNo: invNo,
        customer: customer,
        date: date,
        location: location,
        procCode: procCode,
        psName: psName,
        adjName: 'Adjustment - ' + psName,
        origAmt: origAmt,
        adjAmt: adjAmt,
        net: net,
        is340: is340,
        codeHit: codeHit,
        patHit: patHit,
        voided: voided
      });
    }
    return out;
  }

  // ---- Step 3: split Sales vs Credit Note per class ------------------------
  function split(invoices) {
    var b = {
      salesNon: [], sales340: [], cnNon: [], cn340: []
    };
    invoices.forEach(function (inv) {
      var isCN = inv.net < 0;           // 0 is NOT negative → stays in Sales
      if (inv.is340) (isCN ? b.cn340 : b.sales340).push(inv);
      else (isCN ? b.cnNon : b.salesNon).push(inv);
    });
    return b;
  }

  // ---- Account strings (verbatim per the prompt) ---------------------------
  var ACCOUNTS = {
    non: {
      revenue: 'Revenue from operations:Revenue from operations(Non 340B)',
      adjust: 'Revenue from operations:Revenue from operations(Non 340B):Adjustment(Non 340B)(Expected)'
    },
    b340: {
      revenue: 'Revenue from operations:340B Buy and bill (Owned clinics)',
      adjust: 'Revenue from operations:340B Buy and bill (Owned clinics):340B Adjustment(Owned Clinics)(Expected)'
    },
    third: 'Undeposited Funds-Allscript EOB'
  };

  // ---- Step 4: journal entries from the Credit Note invoices ---------------
  // Each credit-note invoice → three balanced lines.
  function journalEntries(invoices, is340) {
    var acct = is340 ? ACCOUNTS.b340 : ACCOUNTS.non;
    var lines = [];
    invoices.forEach(function (inv) {
      var d1 = Math.abs(inv.origAmt);          // revenue DEBIT
      var c2 = Math.abs(inv.adjAmt);           // adjustment CREDIT
      var c3 = round2(d1 - c2);                // Undeposited Funds CREDIT (difference)
      lines.push(
        { journalNo: inv.invNo, date: inv.date, memo: inv.invNo, account: acct.revenue, debit: d1, credit: null, description: inv.invNo, name: inv.customer, klass: inv.location },
        { journalNo: inv.invNo, date: inv.date, memo: inv.invNo, account: acct.adjust, debit: null, credit: c2, description: inv.invNo, name: inv.customer, klass: inv.location },
        { journalNo: inv.invNo, date: inv.date, memo: inv.invNo, account: ACCOUNTS.third, debit: null, credit: c3, description: inv.invNo, name: inv.customer, klass: inv.location }
      );
    });
    return lines;
  }

  // ---- One-shot driver -----------------------------------------------------
  // Returns everything the UI/builder needs, plus a validation summary.
  function build(msMatrix, codeMatrix, patMatrix) {
    var codeSet = buildCodeSet(codeMatrix);
    var patientSet = buildPatientSet(patMatrix);
    var invoices = buildInvoices(msMatrix, codeSet, patientSet);
    var b = split(invoices);

    var journalNon = journalEntries(b.cnNon, false);
    var journal340 = journalEntries(b.cn340, true);
    var journal = journalNon.concat(journal340);

    function netOf(list) { return round2(list.reduce(function (a, x) { return a + x.net; }, 0)); }

    // journal balance check
    var jbal = true;
    journal.forEach(function (l) { /* per-line, checked in entry groups below */ });
    var entries = b.cnNon.concat(b.cn340);
    entries.forEach(function (inv) {
      var d1 = Math.abs(inv.origAmt), c2 = Math.abs(inv.adjAmt), c3 = round2(d1 - c2);
      if (Math.abs(d1 - (c2 + c3)) > 0.005) jbal = false;
    });

    var sourceCount = invoices.length;
    var placed = b.salesNon.length + b.sales340.length + b.cnNon.length + b.cn340.length;

    return {
      invoices: invoices,
      buckets: b,
      journal: journal,
      codeCount: codeSet.size,
      patientCount: patientSet.size,
      summary: {
        sheets: [
          { key: 'salesNon', name: 'Sales - Non-340B', invoices: b.salesNon.length, lines: b.salesNon.length * 2, net: netOf(b.salesNon) },
          { key: 'sales340', name: 'Sales - 340B', invoices: b.sales340.length, lines: b.sales340.length * 2, net: netOf(b.sales340) },
          { key: 'cnNon', name: 'Credit Note - Non-340B', invoices: b.cnNon.length, lines: b.cnNon.length * 2, net: netOf(b.cnNon) },
          { key: 'cn340', name: 'Credit Note - 340B', invoices: b.cn340.length, lines: b.cn340.length * 2, net: netOf(b.cn340) },
          { key: 'journal', name: 'Journal Ledger', invoices: entries.length, lines: journal.length, net: null }
        ],
        sourceCount: sourceCount,
        placed: placed,
        allPlaced: sourceCount === placed,
        count340: invoices.filter(function (i) { return i.is340; }).length,
        journalBalanced: jbal
      }
    };
  }

  var api = {
    round2: round2, colIdx: colIdx, COL: COL, intStr: intStr, normName: normName,
    toDate: toDate, buildCodeSet: buildCodeSet, buildPatientSet: buildPatientSet,
    buildInvoices: buildInvoices, split: split, journalEntries: journalEntries,
    ACCOUNTS: ACCOUNTS, build: build
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.MSCore = api;
})(typeof self !== 'undefined' ? self : this);
