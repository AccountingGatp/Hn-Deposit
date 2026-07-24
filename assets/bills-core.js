/*
 * bills-core.js — AP Vendor Bills → QBO Import: parsing & lookup logic.
 *
 * DOM-free and dependency-free so the exact same code runs in the browser
 * and under Node for unit tests. It never touches PDF binaries: the caller
 * (bills-app.js in the browser, the test harness under Node) is responsible
 * for turning a PDF into plain text and handing it here.
 *
 * Follows the Healthnomics AP SOP:
 *   1. Extract Invoice #, Invoice Date, Due Date, Amount and the
 *      Ship-To / Customer number from each vendor invoice.
 *   2. Look that number up in the uploaded General Ledger reference to pull
 *      Class, Category, Memo/Description, Mailing Address and Vendor Name.
 *   3. Assemble one QBO-import row per invoice, flagging anything missing or
 *      unmatched with [REVIEW NEEDED] instead of guessing.
 */
(function (root) {
  'use strict';

  var REVIEW = '[REVIEW NEEDED]';

  /* ---------------------------------------------------------------- helpers */

  function clean(s) {
    return String(s == null ? '' : s).replace(/ /g, ' ').replace(/[ \t]+/g, ' ').trim();
  }

  // Parse a money string ("1,900.02", "$18,949.32", "(12.34)") to a Number.
  function money(s) {
    if (s == null) return null;
    var t = String(s).replace(/[$,\s]/g, '');
    var neg = /^\(.*\)$/.test(t);
    t = t.replace(/[()]/g, '');
    var n = parseFloat(t);
    if (!isFinite(n)) return null;
    return neg ? -n : n;
  }

  // Normalise an MM/DD/YYYY (or M/D/YYYY) date to a zero-padded MM/DD/YYYY.
  function normDate(s) {
    if (!s) return '';
    var m = String(s).match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (!m) return '';
    var mm = ('0' + m[1]).slice(-2);
    var dd = ('0' + m[2]).slice(-2);
    var yy = m[3].length === 2 ? '20' + m[3] : m[3];
    return mm + '/' + dd + '/' + yy;
  }

  function toDateObj(mmddyyyy) {
    var m = /(\d{2})\/(\d{2})\/(\d{4})/.exec(mmddyyyy || '');
    if (!m) return null;
    return new Date(Number(m[3]), Number(m[1]) - 1, Number(m[2]));
  }

  // Whole-day difference between two MM/DD/YYYY strings (b - a).
  function dayDiff(a, b) {
    var da = toDateObj(a), db = toDateObj(b);
    if (!da || !db) return null;
    return Math.round((db - da) / 86400000);
  }

  // Add n days to an MM/DD/YYYY string, returning MM/DD/YYYY.
  function addDays(mmddyyyy, n) {
    var d = toDateObj(mmddyyyy);
    if (!d) return '';
    d.setDate(d.getDate() + n);
    return ('0' + (d.getMonth() + 1)).slice(-2) + '/' +
           ('0' + d.getDate()).slice(-2) + '/' + d.getFullYear();
  }

  /* --------------------------------------------------------- vendor detect */

  var VENDORS = {
    cardinal:    { label: 'Cardinal Health', keyLabel: 'Ship-To #' },
    amerisource: { label: 'AmerisourceBergen', keyLabel: 'Customer #' },
    curascript:  { label: 'CuraScript SD', keyLabel: 'Customer #' },
    mckesson:    { label: 'McKesson', keyLabel: 'Customer #' }
  };

  function detectVendor(text) {
    var u = (text || '').toUpperCase();
    // Order matters: Curascript/Priority is distinct from Amerisource/Cencora.
    if (/CURASCRIPT|PRIORITY HEALTHCARE/.test(u)) return 'curascript';
    if (/MCKESSON/.test(u)) return 'mckesson';
    if (/AMERISOURCE|CENCORA|ABDC/.test(u)) return 'amerisource';
    if (/CARDINAL HEALTH|CARDINAL/.test(u)) return 'cardinal';
    return null;
  }

  /* ------------------------------------------------------ per-vendor parse */

  function firstMatch(text, re, group) {
    var m = re.exec(text);
    return m ? clean(m[group || 1]) : '';
  }

  function parseCardinal(text) {
    var r = { invoiceNo: '', invoiceDate: '', dueDate: '', amount: null, key: '', terms: '' };
    r.invoiceDate = normDate(firstMatch(text, /INVOICE DATE\s+(\d{1,2}\/\d{1,2}\/\d{4})/i));
    // "INVOICE DATE 07/13/2026 INVOICE 7482278219" — the number after INVOICE.
    r.invoiceNo = firstMatch(text, /INVOICE\s+(\d{6,})/i);
    r.key = firstMatch(text, /SHIP\s*TO\s+(\d{5,})/i);
    // Grand Total is the payable figure; fall back to Net Amount / Total Due.
    var amt = firstMatch(text, /GRAND TOTAL\s*\$?([\d,]+\.\d{2})/i) ||
              firstMatch(text, /TOTAL DUE\s*\$?([\d,]+\.\d{2})/i) ||
              firstMatch(text, /NET AMOUNT\s*\$?([\d,]+\.\d{2})/i);
    r.amount = money(amt);
    var net = /NET\s+(\d{1,3})\s+DAYS/i.exec(text);
    if (net && r.invoiceDate) { r.terms = 'Net ' + net[1]; r.dueDate = addDays(r.invoiceDate, Number(net[1])); }
    return r;
  }

  function parseAmerisource(text) {
    var r = { invoiceNo: '', invoiceDate: '', dueDate: '', amount: null, key: '', terms: '' };
    r.invoiceNo = firstMatch(text, /Invoice Number:\s*(\d+)/i);
    r.invoiceDate = normDate(firstMatch(text, /Invoice Date:?\s*(\d{1,2}\/\d{1,2}\/\d{4})/i));
    r.amount = money(firstMatch(text, /Total Amount:\s*\$?([\d,]+\.\d{2})/i) ||
                     firstMatch(text, /DOCUMENT TOTAL[\s\S]{0,40}?([\d,]+\.\d{2})/i));
    // Customer number appears as "100564152 / 100564152" under CUSTOMER NUMBER.
    r.key = firstMatch(text, /(\d{6,10})\s*\/\s*\d{6,10}/);
    // Due date sits on the "STATE LIC: … <amount> <date>" line.
    r.dueDate = normDate(firstMatch(text, /[\d,]+\.\d{2}\s+(\d{1,2}\/\d{1,2}\/\d{4})/));
    var net = /Invoice Due in\s*(\d{1,3})\s*days/i.exec(text);
    if (net) {
      r.terms = 'Net ' + net[1];
      if (!r.dueDate && r.invoiceDate) r.dueDate = addDays(r.invoiceDate, Number(net[1]));
    }
    return r;
  }

  function parseCurascript(text, lines) {
    var r = { invoiceNo: '', invoiceDate: '', dueDate: '', amount: null, key: '', terms: '' };
    r.invoiceNo = firstMatch(text, /INVOICE NO\.?:?\s*(\d+)/i);
    r.invoiceDate = normDate(firstMatch(text, /INVOICE DATE\.?:?\s*(\d{1,2}\/\d{1,2}\/\d{4})/i));
    r.dueDate = normDate(firstMatch(text, /NET DUE DATE:?\s*(\d{1,2}\/\d{1,2}\/\d{4})/i));
    // Payable total. On some layouts "TOTAL:" and its figure land on separate
    // lines, so fall back to the "…IF PAID BY CHECK/EFT" amount (marked "X")
    // and then "Invoice Total". Never SUBTOTAL / TOTAL TAX.
    r.amount = money(
      firstMatch(text, /(?:^|[^A-Z])TOTAL\s*:\s*\$([\d,]+\.\d{2})/i) ||
      firstMatch(text, /\$([\d,]+\.\d{2})\s*X\b/) ||
      firstMatch(text, /Invoice Total\s*:?\s*\$([\d,]+\.\d{2})/i)
    );
    // Customer number is the first token on the row beneath the header line
    // "Customer # P.O. # Terms Order # Rep. # State Reg. # DEA #".
    if (lines) {
      for (var i = 0; i < lines.length; i++) {
        if (/Customer #/i.test(lines[i]) && /P\.?O\.? #/i.test(lines[i])) {
          var next = lines[i + 1] || '';
          var m = /(\d{5,})/.exec(next);
          if (m) { r.key = m[1]; }
          var t = /\b(DIRECT DEBIT|NET \d+|\d+)\b/i.exec(next.replace(/^\s*\d+\s+\S+\s+/, ''));
          break;
        }
      }
    }
    if (!r.key) r.key = firstMatch(text, /Customer #[\s\S]{0,60}?(\d{6,})/i);
    if (r.invoiceDate && r.dueDate) {
      var d = dayDiff(r.invoiceDate, r.dueDate);
      if (d != null) r.terms = 'Net ' + d;
    }
    return r;
  }

  function parseMckesson(text) {
    var r = { invoiceNo: '', invoiceDate: '', dueDate: '', amount: null, key: '', terms: '' };
    r.invoiceNo = firstMatch(text, /Billing No\.?:?\s*(\d+)/i);
    r.invoiceDate = normDate(firstMatch(text, /Billing Date:?\s*(\d{1,2}\/\d{1,2}\/\d{4})/i));
    // "331106 LV0 095 1 of 3"  → customer = first token (Customer Route Stop Page).
    r.key = firstMatch(text, /(\d{4,8})\s+\S{2,4}\s+\d{2,4}\s+\d+\s+of\s+\d+/i);
    // "NET PAYABLE BY STATEMENT DATE0 7/25/2026: $13,213.63"
    var np = /NET PAYABLE BY STATEMENT DATE\s*(\d?)\s*(\d{1,2}\/\d{1,2}\/\d{4})\s*:\s*\$?\s*([\d,]+\.\d{2})/i.exec(text);
    if (np) {
      r.dueDate = normDate((np[1] || '') + np[2]);
      r.amount = money(np[3]);
    }
    if (r.amount == null) {
      r.amount = money(firstMatch(text, /NET PAYABLE[\s\S]{0,40}?:\s*\$?([\d,]+\.\d{2})/i) ||
                       firstMatch(text, /TOTAL RX PURCHASES:\s*\$?([\d,]+\.\d{2})/i));
    }
    if (r.invoiceDate && r.dueDate) {
      var d = dayDiff(r.invoiceDate, r.dueDate);
      if (d != null && d >= 0) r.terms = 'Net ' + d;
    }
    return r;
  }

  var PARSERS = {
    cardinal: parseCardinal,
    amerisource: parseAmerisource,
    curascript: parseCurascript,
    mckesson: parseMckesson
  };

  // Parse one invoice's text. `lines` is the optional array of reconstructed
  // text lines (used by Curascript's positional lookup); `text` is the same
  // content joined with newlines.
  function parseInvoice(text, lines, forcedVendor) {
    var vendor = forcedVendor || detectVendor(text);
    var out = {
      vendor: vendor,
      vendorLabel: vendor ? VENDORS[vendor].label : '',
      keyLabel: vendor ? VENDORS[vendor].keyLabel : 'Ship-To / Customer #',
      invoiceNo: '', invoiceDate: '', dueDate: '', amount: null, key: '', terms: '',
      flags: []
    };
    if (!vendor) { out.flags.push('Vendor not recognised'); return out; }
    var parsed = PARSERS[vendor](text, lines);
    out.invoiceNo = parsed.invoiceNo;
    out.invoiceDate = parsed.invoiceDate;
    out.dueDate = parsed.dueDate;
    out.amount = parsed.amount;
    out.key = parsed.key;
    out.terms = parsed.terms;

    if (!out.invoiceNo) out.flags.push('Invoice # not found');
    if (!out.invoiceDate) out.flags.push('Invoice date not found');
    if (out.amount == null) out.flags.push('Amount not found');
    if (!out.key) out.flags.push(out.keyLabel + ' not found');
    return out;
  }

  /* -------------------------------------------------------------- GL index */

  // Header keyword sets → logical field. First matching header wins.
  var GL_FIELDS = {
    key:      [/ship\s*[-_ ]?to.*(no|num|#)/i, /customer.*(no|num|#)/i, /account.*(no|num|#)/i, /ship\s*to/i, /customer/i, /account/i],
    vendor:   [/vendor.*name/i, /vendor/i, /supplier/i, /payee/i],
    klass:    [/class/i],
    category: [/category/i, /account$/i, /expense.*account/i, /gl.*account/i, /account.*name/i],
    memo:     [/memo/i, /description/i, /desc/i, /note/i],
    address:  [/mailing.*address/i, /address/i, /mail.*addr/i]
  };

  function detectGLColumns(headerRow) {
    var map = {};
    var headers = headerRow.map(function (h) { return clean(h); });
    Object.keys(GL_FIELDS).forEach(function (field) {
      var patterns = GL_FIELDS[field];
      for (var p = 0; p < patterns.length; p++) {
        for (var c = 0; c < headers.length; c++) {
          if (map[field] != null) break;
          if (patterns[p].test(headers[c]) && Object.keys(map).map(function (k) { return map[k]; }).indexOf(c) === -1) {
            map[field] = c;
            break;
          }
        }
        if (map[field] != null) break;
      }
    });
    return { map: map, headers: headers };
  }

  function normKey(v) {
    return String(v == null ? '' : v).replace(/\.0$/, '').replace(/[\s\-]/g, '').toUpperCase();
  }

  // Build a lookup index from a sheet given as an array-of-arrays (row 0 = header).
  function buildGLIndex(aoa) {
    if (!aoa || !aoa.length) return { rows: {}, columns: { map: {}, headers: [] }, count: 0 };
    // Find the header row: the first row that yields a usable "key" column.
    var headerIdx = 0, cols = detectGLColumns(aoa[0]);
    if (cols.map.key == null) {
      for (var h = 0; h < Math.min(aoa.length, 10); h++) {
        var c = detectGLColumns(aoa[h]);
        if (c.map.key != null) { headerIdx = h; cols = c; break; }
      }
    }
    var idx = {};
    var count = 0;
    for (var r = headerIdx + 1; r < aoa.length; r++) {
      var row = aoa[r];
      if (!row || cols.map.key == null) continue;
      var rawKey = row[cols.map.key];
      if (rawKey == null || clean(rawKey) === '') continue;
      var rec = {
        vendor:   cols.map.vendor   != null ? clean(row[cols.map.vendor])   : '',
        klass:    cols.map.klass    != null ? clean(row[cols.map.klass])    : '',
        category: cols.map.category != null ? clean(row[cols.map.category]) : '',
        memo:     cols.map.memo     != null ? clean(row[cols.map.memo])     : '',
        address:  cols.map.address  != null ? clean(row[cols.map.address])  : '',
        rawKey:   clean(rawKey)
      };
      var nk = normKey(rawKey);
      if (!idx[nk]) { idx[nk] = rec; count++; }
      // Also index the zero-stripped form so "062792" ↔ "62792" both match.
      var stripped = nk.replace(/^0+/, '');
      if (stripped && !idx[stripped]) idx[stripped] = rec;
    }
    return { rows: idx, columns: cols, count: count };
  }

  function glLookup(glIndex, key) {
    if (!glIndex || !key) return null;
    var nk = normKey(key);
    if (glIndex.rows[nk]) return glIndex.rows[nk];
    var stripped = nk.replace(/^0+/, '');
    if (stripped && glIndex.rows[stripped]) return glIndex.rows[stripped];
    return null;
  }

  /* --------------------------------------------------- assemble import row */

  // Merge a parsed invoice with its GL match into the final import record.
  // opts: { defaultAccount, defaultClass }
  function assembleRow(parsed, glIndex, opts) {
    opts = opts || {};
    var gl = glLookup(glIndex, parsed.key);
    var flags = parsed.flags.slice();
    var matched = !!gl;

    var vendor = (gl && gl.vendor) || parsed.vendorLabel || '';
    var klass = (gl && gl.klass) || opts.defaultClass || '';
    var category = (gl && gl.category) || opts.defaultAccount || '';
    var memo = (gl && gl.memo) || '';
    var address = (gl && gl.address) || '';

    if (glIndex && glIndex.count > 0 && parsed.key && !matched) {
      flags.push(parsed.keyLabel + ' ' + parsed.key + ' not matched in GL');
    }
    var out = {
      vendorKey: parsed.vendor,
      vendorLabel: parsed.vendorLabel,
      keyLabel: parsed.keyLabel,
      billNo: parsed.invoiceNo || REVIEW,
      vendor: vendor || REVIEW,
      billDate: parsed.invoiceDate || REVIEW,
      dueDate: parsed.dueDate || '',
      terms: parsed.terms || '',
      account: category || REVIEW,
      amount: parsed.amount,
      memo: memo || '',
      klass: klass || '',
      key: parsed.key || REVIEW,
      address: address || '',
      matched: matched,
      flags: flags,
      needsReview: flags.length > 0 || !category || !vendor || parsed.amount == null
    };
    return out;
  }

  // Full pipeline: array of {text, lines, name} → assembled rows.
  function processInvoices(invoices, glIndex, opts) {
    return invoices.map(function (inv) {
      var parsed = parseInvoice(inv.text, inv.lines, inv.vendor);
      var row = assembleRow(parsed, glIndex, opts);
      row.sourceName = inv.name || '';
      return row;
    });
  }

  var api = {
    REVIEW: REVIEW,
    VENDORS: VENDORS,
    money: money,
    normDate: normDate,
    addDays: addDays,
    dayDiff: dayDiff,
    detectVendor: detectVendor,
    parseInvoice: parseInvoice,
    detectGLColumns: detectGLColumns,
    buildGLIndex: buildGLIndex,
    glLookup: glLookup,
    assembleRow: assembleRow,
    processInvoices: processInvoices
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.APBillsCore = api;
})(typeof self !== 'undefined' ? self : this);
