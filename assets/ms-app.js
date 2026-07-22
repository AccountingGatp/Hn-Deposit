/* ms-app.js — UI controller for Monthly Services → QBO Import (340B Split).
   Ties the DOM to MSCore (logic) and MSXlsx (workbook). All in-browser. */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };
  var state = {
    ms: null, msName: '',       // Monthly Services matrix
    jq: null, jqName: '',       // JQ Code List matrix
    pt: null, ptName: '',       // Patient List matrix
    result: null,
    fmt2: function (n) { return (n === null || isNaN(n)) ? '' : Number(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  };

  // ---------- read a workbook to a matrix (array of arrays) ----------
  function readMatrix(file, preferredSheet, cb, onErr) {
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        var sheetName = wb.SheetNames[0];
        if (preferredSheet) {
          var found = wb.SheetNames.find(function (n) { return n.toLowerCase() === preferredSheet.toLowerCase(); });
          if (found) sheetName = found;
        }
        var ws = wb.Sheets[sheetName];
        var matrix = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
        cb(matrix);
      } catch (err) { onErr(err); }
    };
    reader.onerror = function () { onErr(new Error('Could not read the file.')); };
    reader.readAsArrayBuffer(file);
  }

  function handleMS(file) {
    state.msName = file.name; $('msFileName').textContent = '✓ ' + file.name;
    readMatrix(file, 'Monthly Services', function (m) {
      // sanity: needs enough columns to reach Fee (col BP = 68)
      if (!m || m.length < 2 || (m[0] && m[0].length < 68)) {
        state.ms = null; showErr('That doesn\'t look like a Monthly Services export (need columns through "Fee" / col BP).');
      } else { state.ms = m; showErr(''); }
      refreshReady();
    }, function (e) { state.ms = null; showErr(e.message); refreshReady(); });
  }

  function handleJQ(file) {
    state.jqName = file.name; $('jqFileName').textContent = '✓ ' + file.name;
    readMatrix(file, null, function (m) {
      if (!m || !m.length) { state.jq = null; showErr('The JQ Code List looks empty.'); }
      else { state.jq = m; showErr(''); }
      refreshReady();
    }, function (e) { state.jq = null; showErr(e.message); refreshReady(); });
  }

  function handlePT(file) {
    state.ptName = file.name; $('ptFileName').textContent = '✓ ' + file.name;
    readMatrix(file, null, function (m) {
      if (!m || m.length < 2) { state.pt = null; showErr('The Patient List looks empty.'); }
      else { state.pt = m; showErr(''); }
      refreshReady();
    }, function (e) { state.pt = null; showErr(e.message); refreshReady(); });
  }

  function refreshReady() {
    var ok = state.ms && state.jq && state.pt;
    $('processBtn').disabled = !ok;
    if (ok) {
      $('readyHint').textContent = (state.ms.length - 1) + ' service rows · ' +
        (state.jq.length - 1) + ' codes · ' + (state.pt.length - 1) + ' patients loaded.';
    } else {
      $('readyHint').textContent = 'Upload all three files to enable.';
    }
  }

  function showErr(msg) {
    var box = $('errBox');
    if (!msg) { box.classList.add('hidden'); box.textContent = ''; return; }
    box.classList.remove('hidden'); box.textContent = '⚠ ' + msg;
  }

  // ---------- process ----------
  function runProcess() {
    showErr('');
    try {
      state.result = MSCore.build(state.ms, state.jq, state.pt);
    } catch (e) {
      showErr('Could not process the files: ' + e.message); return;
    }
    renderResults();
    $('resultsCard').classList.remove('hidden');
    $('resultsCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function tile(k, v, cls) {
    return '<div class="tile ' + (cls || '') + '"><div class="k">' + k + '</div><div class="v">' + v + '</div></div>';
  }
  function fmtDate(d) {
    if (!(d instanceof Date) || isNaN(d)) return '';
    return (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
  }

  function renderResults() {
    var r = state.result, sm = r.summary;

    // validation tiles
    $('tiles').innerHTML = [
      tile('Source rows', sm.sourceCount),
      tile('All rows placed', sm.allPlaced ? 'OK' : 'MISMATCH', sm.allPlaced ? 'good' : 'warnv'),
      tile('340B invoices', sm.count340),
      tile('Journal balances', sm.journalBalanced ? 'OK' : 'OUT', sm.journalBalanced ? 'good' : 'warnv')
    ].join('');

    // per-sheet summary table
    var sh = '<thead><tr><th>Sheet</th><th class="num">Invoices</th><th class="num">Lines</th><th class="num">Net total</th></tr></thead><tbody>';
    sm.sheets.forEach(function (s) {
      sh += '<tr><td>' + esc(s.name) + '</td><td class="num">' + s.invoices + '</td><td class="num">' + s.lines +
        '</td><td class="num">' + (s.net === null ? '—' : state.fmt2(s.net)) + '</td></tr>';
    });
    var totInv = sm.sheets.slice(0, 4).reduce(function (a, s) { return a + s.invoices; }, 0);
    sh += '<tr><td><b>Total (Sales + Credit Note)</b></td><td class="num"><b>' + totInv +
      '</b></td><td class="num"></td><td class="num"></td></tr></tbody>';
    $('sumTable').innerHTML = sh;

    // reconciliation line
    $('reconLine').innerHTML = 'Source rows: <b>' + sm.sourceCount + '</b> · placed in sheets: <b>' + sm.placed +
      '</b> → <span class="badge ' + (sm.allPlaced ? 'ok">✓ every row placed exactly once' : 'bad">✗ mismatch') + '</span>';

    // credit notes preview
    var cn = r.buckets.cnNon.concat(r.buckets.cn340);
    $('cnCount').textContent = cn.length;
    var ch = '<thead><tr><th>Invoice #</th><th>Class</th><th>Customer</th><th>Date</th><th>Product/Service</th>' +
      '<th class="num">Original</th><th class="num">Adjustment</th><th class="num">Net</th></tr></thead><tbody>';
    cn.forEach(function (inv) {
      ch += '<tr><td>' + esc(inv.invNo) + '</td><td>' + (inv.is340 ? '340B' : 'Non-340B') + '</td><td>' + esc(inv.customer) +
        '</td><td>' + fmtDate(inv.date) + '</td><td>' + esc(inv.psName) + '</td><td class="num">' + state.fmt2(inv.origAmt) +
        '</td><td class="num">' + state.fmt2(inv.adjAmt) + '</td><td class="num">' + state.fmt2(inv.net) + '</td></tr>';
    });
    if (!cn.length) ch += '<tr><td colspan="8" class="muted">No credit notes in this period (no negative-net invoices).</td></tr>';
    ch += '</tbody>';
    $('cnTable').innerHTML = ch;

    // journal preview
    $('jCount').textContent = r.journal.length;
    var jh = '<thead><tr><th>Journal No</th><th>Date</th><th>Account</th><th class="num">Debit</th><th class="num">Credit</th><th>Name</th><th>Class</th></tr></thead><tbody>';
    r.journal.forEach(function (l) {
      jh += '<tr><td>' + esc(l.journalNo) + '</td><td>' + fmtDate(l.date) + '</td><td class="reason">' + esc(l.account) +
        '</td><td class="num">' + (l.debit === null ? '' : state.fmt2(l.debit)) + '</td><td class="num">' +
        (l.credit === null ? '' : state.fmt2(l.credit)) + '</td><td>' + esc(l.name) + '</td><td>' + esc(l.klass) + '</td></tr>';
    });
    if (!r.journal.length) jh += '<tr><td colspan="7" class="muted">No journal entries (journal is built only from credit notes).</td></tr>';
    jh += '</tbody>';
    $('jTable').innerHTML = jh;
  }

  // ---------- download ----------
  async function download() {
    try {
      var wb = MSXlsx.buildWorkbook(ExcelJS, state.result);
      var buf = await wb.xlsx.writeBuffer();
      var blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      var client = ($('client').value.trim() || 'Healthnomics').replace(/[^\w\-]+/g, '_');
      var name = client + '_QBO_Monthly_Services_340B_Split.xlsx';
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url; a.download = name; document.body.appendChild(a); a.click();
      setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 1000);
    } catch (e) {
      showErr('Could not build the workbook: ' + e.message);
    }
  }

  // ---------- wire up ----------
  function wireDrop(dropId, inputId, handler) {
    var drop = $(dropId), input = $(inputId);
    input.addEventListener('change', function (e) { if (e.target.files[0]) handler(e.target.files[0]); });
    ['dragover', 'dragenter'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('dragover'); }); });
    ['dragleave', 'drop'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('dragover'); }); });
    drop.addEventListener('drop', function (e) { if (e.dataTransfer.files[0]) handler(e.dataTransfer.files[0]); });
  }

  document.addEventListener('DOMContentLoaded', function () {
    wireDrop('dropMS', 'msInput', handleMS);
    wireDrop('dropJQ', 'jqInput', handleJQ);
    wireDrop('dropPT', 'ptInput', handlePT);
    $('processBtn').addEventListener('click', runProcess);
    $('downloadBtn').addEventListener('click', download);
  });
})();
