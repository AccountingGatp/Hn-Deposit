/*
 * bills-app.js — UI controller for AP Vendor Bills → QBO Import.
 * Ties the DOM to APBillsCore (logic) and APBillsXlsx (workbook), and uses the
 * vendored pdf.js to turn each invoice PDF into plain text lines.
 */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };

  var state = {
    invoices: [],   // [{ name, text, lines }]
    glIndex: null,  // APBillsCore.buildGLIndex output
    glName: '',
    rows: null,     // assembled import rows
    opts: {},
    fmt2: function (n) { return (n == null || isNaN(n)) ? '' : Number(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  };

  var pdfReady = !!window.pdfjsLib;
  window.addEventListener('pdfjs-ready', function () { pdfReady = true; refreshReady(); });

  /* ------------------------------------------------------------ PDF → text */

  // Reconstruct reading-order lines by grouping text items on their y-position
  // (top of page first) and sorting each line left-to-right by x.
  function itemsToLines(items) {
    var buckets = {};
    items.forEach(function (it) {
      if (!it.str) return;
      var y = Math.round(it.transform[5]);
      (buckets[y] = buckets[y] || []).push({ x: it.transform[4], s: it.str });
    });
    var ys = Object.keys(buckets).map(Number).sort(function (a, b) { return b - a; });
    var lines = [];
    ys.forEach(function (y) {
      var line = buckets[y].sort(function (a, b) { return a.x - b.x; })
        .map(function (o) { return o.s; }).join(' ').replace(/\s+/g, ' ').trim();
      if (line) lines.push(line);
    });
    return lines;
  }

  function readPdf(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function (e) {
        var task = window.pdfjsLib.getDocument({ data: new Uint8Array(e.target.result) });
        task.promise.then(function (doc) {
          var lines = [];
          var chain = Promise.resolve();
          for (var p = 1; p <= doc.numPages; p++) {
            (function (pageNum) {
              chain = chain.then(function () {
                return doc.getPage(pageNum).then(function (page) {
                  return page.getTextContent().then(function (tc) {
                    lines = lines.concat(itemsToLines(tc.items));
                  });
                });
              });
            })(p);
          }
          chain.then(function () {
            resolve({ name: file.name, lines: lines, text: lines.join('\n') });
          }).catch(reject);
        }).catch(reject);
      };
      reader.onerror = function () { reject(new Error('Could not read ' + file.name)); };
      reader.readAsArrayBuffer(file);
    });
  }

  /* -------------------------------------------------------------- handlers */

  function handlePdfs(files) {
    if (!files || !files.length) return;
    if (!window.pdfjsLib) { showErr('The PDF engine is still loading — try again in a moment.'); return; }
    showErr('');
    $('pdfFileName').textContent = '… reading ' + files.length + ' file(s)';
    var list = Array.prototype.slice.call(files).filter(function (f) { return /\.pdf$/i.test(f.name); });
    Promise.all(list.map(readPdf)).then(function (parsed) {
      // Merge with any already-loaded invoices, de-duplicating by file name.
      var byName = {};
      state.invoices.forEach(function (inv) { byName[inv.name] = inv; });
      parsed.forEach(function (inv) { byName[inv.name] = inv; });
      state.invoices = Object.keys(byName).map(function (k) { return byName[k]; });
      var recognised = state.invoices.filter(function (inv) { return APBillsCore.detectVendor(inv.text); }).length;
      $('pdfFileName').textContent = '✓ ' + state.invoices.length + ' invoice(s) · ' + recognised + ' vendor(s) recognised';
      showErr('');
      refreshReady();
    }).catch(function (err) {
      $('pdfFileName').textContent = '';
      showErr('Could not parse a PDF: ' + err.message);
    });
  }

  function handleGl(file) {
    state.glName = file.name;
    $('glFileName').textContent = '✓ ' + file.name;
    var ext = file.name.split('.').pop().toLowerCase();
    var done = function (matrix) {
      try {
        state.glIndex = APBillsCore.buildGLIndex(matrix);
        showGlInfo();
        showErr('');
      } catch (e) { state.glIndex = null; showErr('GL file: ' + e.message); }
      refreshReady();
    };
    if (ext === 'csv') {
      var reader = new FileReader();
      reader.onload = function (e) { done(csvToMatrix(e.target.result)); };
      reader.onerror = function () { showErr('Could not read the GL CSV.'); };
      reader.readAsText(file);
    } else {
      sheetToMatrix(file, done, function (e) { showErr('GL file: ' + e.message); });
    }
  }

  function showGlInfo() {
    var box = $('glInfo');
    if (!state.glIndex) { box.style.display = 'none'; return; }
    var m = state.glIndex.columns.map;
    var headers = state.glIndex.columns.headers;
    var label = function (k) { return m[k] != null ? '<b>' + esc(headers[m[k]]) + '</b>' : '<span class="muted">— not found —</span>'; };
    box.style.display = '';
    box.innerHTML = '<b>GL loaded:</b> ' + state.glIndex.count + ' mapping row(s). Detected columns → ' +
      'Key: ' + label('key') + ' · Vendor: ' + label('vendor') + ' · Class: ' + label('klass') +
      ' · Category: ' + label('category') + ' · Memo: ' + label('memo') + ' · Address: ' + label('address') +
      (m.key == null ? '<br><span style="color:var(--warn)">⚠ No Ship-To / Customer number column detected — lookups will not match.</span>' : '');
  }

  /* --------------------------------------------------------- file readers */

  function sheetToMatrix(file, cb, onErr) {
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var wb = XLSX.read(e.target.result, { type: 'array' });
        var ws = wb.Sheets[wb.SheetNames[0]];
        var matrix = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
        cb(matrix.filter(function (r) { return r.some(function (v) { return String(v).trim() !== ''; }); }));
      } catch (err) { onErr(err); }
    };
    reader.onerror = function () { onErr(new Error('Could not read the file.')); };
    reader.readAsArrayBuffer(file);
  }

  function csvToMatrix(text) {
    var rows = [], row = [], field = '', i = 0, q = false;
    while (i < text.length) {
      var c = text[i];
      if (q) {
        if (c === '"') { if (text[i + 1] === '"') { field += '"'; i += 2; continue; } q = false; i++; continue; }
        field += c; i++; continue;
      }
      if (c === '"') { q = true; i++; continue; }
      if (c === ',') { row.push(field); field = ''; i++; continue; }
      if (c === '\r') { i++; continue; }
      if (c === '\n') { row.push(field); rows.push(row); row = []; field = ''; i++; continue; }
      field += c; i++;
    }
    if (field.length || row.length) { row.push(field); rows.push(row); }
    return rows.filter(function (r) { return r.some(function (v) { return String(v).trim() !== ''; }); });
  }

  /* ---------------------------------------------------------------- ready */

  function refreshReady() {
    var ok = state.invoices.length > 0 && (pdfReady || window.pdfjsLib);
    $('processBtn').disabled = !ok;
    if (!state.invoices.length) {
      $('readyHint').textContent = 'Upload at least one invoice PDF to enable.';
    } else if (!state.glIndex) {
      $('readyHint').textContent = state.invoices.length + ' invoice(s) loaded · no GL yet (lookups will be flagged for review).';
    } else {
      $('readyHint').textContent = state.invoices.length + ' invoice(s) · ' + state.glIndex.count + ' GL mapping(s) loaded.';
    }
  }

  function showErr(msg) {
    var box = $('errBox');
    if (!msg) { box.classList.add('hidden'); box.textContent = ''; return; }
    box.classList.remove('hidden'); box.textContent = '⚠ ' + msg;
  }

  /* -------------------------------------------------------------- process */

  function runProcess() {
    showErr('');
    if (!state.invoices.length) { showErr('Upload at least one invoice PDF.'); return; }
    state.opts = {
      defaultAccount: $('defAccount').value.trim(),
      defaultClass: $('defClass').value.trim(),
      client: ($('client').value.trim() || 'Client')
    };
    try {
      state.rows = APBillsCore.processInvoices(state.invoices, state.glIndex, state.opts);
    } catch (e) { showErr('Could not process the invoices: ' + e.message); return; }
    renderResults();
    $('resultsCard').classList.remove('hidden');
    $('resultsCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  /* --------------------------------------------------------------- render */

  function renderResults() {
    var rows = state.rows;
    var total = rows.reduce(function (a, b) { return a + (b.amount || 0); }, 0);
    var matched = rows.filter(function (r) { return r.matched; }).length;
    var review = rows.filter(function (r) { return r.needsReview; }).length;
    var unrecognised = rows.filter(function (r) { return !r.vendorKey; }).length;

    $('tiles').innerHTML = [
      tile('Bills parsed', rows.length),
      tile('Total payable', '$' + state.fmt2(total), 'good'),
      tile('Matched in GL', matched + ' / ' + rows.length, matched === rows.length ? 'good' : ''),
      tile('Need review', review, review ? 'warnv' : 'good')
    ].join('');

    // Vendor breakdown note
    var byVendor = {};
    rows.forEach(function (r) { var k = r.vendorLabel || 'Unrecognised'; byVendor[k] = (byVendor[k] || 0) + 1; });
    var ab = $('assumpBox');
    ab.classList.remove('hidden');
    ab.innerHTML = '<b>By vendor:</b><ul>' +
      Object.keys(byVendor).map(function (k) { return '<li>(' + byVendor[k] + ') ' + esc(k) + '</li>'; }).join('') +
      '</ul>' + (unrecognised ? '<span style="color:var(--warn)">⚠ ' + unrecognised + ' file(s) not recognised as a supported vendor.</span>'
        : 'Values shown come from the invoice PDF and your GL. Rows with ' + APBillsCore.REVIEW + ' need a fix before importing.');

    // QBO Import preview
    $('procCount').textContent = rows.length;
    var ph = '<thead><tr><th>Source</th><th>Bill No.</th><th>Vendor</th><th>Bill Date</th><th>Due Date</th>' +
      '<th>Ship-To/Cust #</th><th>Account</th><th class="num">Amount</th><th>Class</th></tr></thead><tbody>';
    rows.forEach(function (p) {
      ph += '<tr class="' + (p.needsReview ? 'needs' : '') + '">' +
        '<td class="reason">' + esc(p.sourceName) + '</td>' +
        '<td>' + cell(p.billNo) + '</td>' +
        '<td>' + cell(p.vendor) + '</td>' +
        '<td>' + cell(p.billDate) + '</td>' +
        '<td>' + esc(p.dueDate) + '</td>' +
        '<td>' + cell(p.key) + '</td>' +
        '<td>' + cell(p.account) + '</td>' +
        '<td class="num">' + (p.amount == null ? flag('—') : state.fmt2(p.amount)) + '</td>' +
        '<td>' + esc(p.klass) + '</td></tr>';
    });
    ph += '</tbody>';
    $('procTable').innerHTML = ph;

    // Exceptions
    var exceptions = rows.filter(function (r) { return r.needsReview; });
    $('npCount').textContent = exceptions.length;
    if (exceptions.length) {
      var nh = '<thead><tr><th>Source</th><th>Vendor</th><th>Bill No.</th><th>Ship-To/Cust #</th><th>What needs review</th></tr></thead><tbody>';
      exceptions.forEach(function (n) {
        nh += '<tr><td class="reason">' + esc(n.sourceName) + '</td><td>' + esc(n.vendorLabel || n.vendor) + '</td>' +
          '<td>' + esc(n.billNo) + '</td><td>' + esc(n.key) + '</td>' +
          '<td class="reason">' + esc((n.flags || []).join('; ')) + '</td></tr>';
      });
      nh += '</tbody>';
      $('npTable').innerHTML = nh;
    } else {
      $('npTable').innerHTML = '<thead><tr><th>Status</th></tr></thead><tbody><tr><td style="color:var(--ok);font-weight:700;">✓ No exceptions — every invoice extracted and matched.</td></tr></tbody>';
    }

    updateDlBadge(review);
  }

  function updateDlBadge(review) {
    var b = $('dlBadge');
    if (review > 0) { b.className = 'badge bad'; b.textContent = review + ' bill(s) need review'; }
    else { b.className = 'badge ok'; b.textContent = '✓ All bills ready to import'; }
  }

  function cell(v) {
    return (v === APBillsCore.REVIEW) ? flag(v) : esc(v);
  }
  function flag(v) { return '<span style="color:var(--warn);font-weight:700;">' + esc(v) + '</span>'; }
  function tile(k, v, cls) {
    return '<div class="tile ' + (cls || '') + '"><div class="k">' + k + '</div><div class="v">' + v + '</div></div>';
  }
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  /* ------------------------------------------------------------- download */

  function download() {
    try {
      var wb = APBillsXlsx.buildWorkbook(ExcelJS, state.rows, state.opts);
      wb.xlsx.writeBuffer().then(function (buf) {
        var blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        var dates = state.rows.map(function (r) { return r.billDate; })
          .filter(function (d) { return d && d !== APBillsCore.REVIEW; }).sort();
        var stamp = dates.length ? dates[0].replace(/\//g, '-') : 'export';
        var name = (state.opts.client || 'Client').replace(/[^\w\-]+/g, '_') + '_QBO_Bills_Import_' + stamp + '.xlsx';
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url; a.download = name; document.body.appendChild(a); a.click();
        setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 1000);
      });
    } catch (e) {
      showErr('Could not build the workbook: ' + e.message);
    }
  }

  /* --------------------------------------------------------------- wiring */

  function wireDrop(dropId, inputId, handler, multi) {
    var drop = $(dropId), input = $(inputId);
    input.addEventListener('change', function (e) {
      if (e.target.files && e.target.files.length) handler(multi ? e.target.files : e.target.files[0]);
    });
    ['dragover', 'dragenter'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('dragover'); }); });
    ['dragleave', 'drop'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('dragover'); }); });
    drop.addEventListener('drop', function (e) {
      if (e.dataTransfer.files && e.dataTransfer.files.length) handler(multi ? e.dataTransfer.files : e.dataTransfer.files[0]);
    });
  }

  document.addEventListener('DOMContentLoaded', function () {
    wireDrop('dropPdf', 'pdfInput', handlePdfs, true);
    wireDrop('dropGl', 'glInput', handleGl, false);
    $('processBtn').addEventListener('click', runProcess);
    $('downloadBtn').addEventListener('click', download);
    refreshReady();
  });
})();
