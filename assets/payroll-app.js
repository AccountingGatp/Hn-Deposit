/* payroll-app.js — UI controller. Ties the DOM to PJCore (logic) and PJXlsx (workbook). */
(function () {
  'use strict';

  const $ = (id) => document.getElementById(id);
  const state = {
    matrix: null,     // raw sheet rows (array of arrays)
    fileName: '',
    result: null,     // computed bundle
    fmt2: (n) => (isNaN(n) ? '' : Number(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })),
  };

  // ---------- CSV parsing (RFC-4180-ish) ----------
  function parseCSV(text) {
    const rows = [];
    let row = [], field = '', i = 0, q = false;
    while (i < text.length) {
      const c = text[i];
      if (q) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
          q = false; i++; continue;
        }
        field += c; i++; continue;
      }
      if (c === '"') { q = true; i++; continue; }
      if (c === ',') { row.push(field); field = ''; i++; continue; }
      if (c === '\r') { i++; continue; }
      if (c === '\n') { row.push(field); rows.push(row); row = []; field = ''; i++; continue; }
      field += c; i++;
    }
    if (field.length || row.length) { row.push(field); rows.push(row); }
    return rows;
  }

  function sheetToMatrix(file, cb, onErr) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
        cb(matrix);
      } catch (err) { onErr(err); }
    };
    reader.onerror = () => onErr(new Error('Could not read the file.'));
    reader.readAsArrayBuffer(file);
  }

  // ---------- File handler ----------
  function handleFile(file) {
    state.fileName = file.name;
    $('regFileName').textContent = '✓ ' + file.name;
    const ext = file.name.split('.').pop().toLowerCase();
    const done = (matrix) => {
      try {
        if (!matrix || !matrix.length) throw new Error('The register file appears to be empty.');
        state.matrix = matrix;
        showErr('');
      } catch (e) { state.matrix = null; showErr(e.message); }
      refreshReady();
    };
    if (ext === 'csv') {
      const reader = new FileReader();
      reader.onload = (e) => done(parseCSV(e.target.result));
      reader.onerror = () => showErr('Could not read the register CSV.');
      reader.readAsText(file);
    } else {
      sheetToMatrix(file, done, (e) => { showErr(e.message); });
    }
  }

  function refreshReady() {
    const ok = state.matrix && state.matrix.length;
    $('processBtn').disabled = !ok;
    if (ok) {
      const preview = PJCore.normalise(state.matrix);
      $('readyHint').textContent = preview.rows.length + ' payroll lines · ' + preview.dropped + ' spacer rows dropped'
        + (preview.hadHeader ? ' · headers detected' : ' · positional layout');
    } else {
      $('readyHint').textContent = 'Upload the register to enable.';
    }
  }

  function showErr(msg) {
    const box = $('errBox');
    if (!msg) { box.classList.add('hidden'); box.textContent = ''; return; }
    box.classList.remove('hidden'); box.textContent = '⚠ ' + msg;
  }

  // ---------- Physician rules from the textarea ----------
  function readPhysicianRules() {
    const raw = ($('physRules').value || '').split(/[\n,]/).map((s) => s.trim()).filter(Boolean);
    if (!raw.length) return PJCore.defaultPhysicianRules();
    return raw.map((v) => {
      // "OBrian*" or "starts:OBrian" → startsWith; otherwise contains
      if (v.endsWith('*')) return { type: 'startsWith', value: v.slice(0, -1).trim() };
      const m = v.match(/^starts:(.*)$/i);
      if (m) return { type: 'startsWith', value: m[1].trim() };
      return { type: 'contains', value: v };
    });
  }

  function readAccounts() {
    return {
      wagesEmployee: $('acctWagesEmp').value.trim() || 'Payroll Wages & Salaries:Employee',
      wagesPhysicians: $('acctWagesPhys').value.trim() || 'Payroll Wages & Salaries:Physicians',
      employerTax: $('acctErTax').value.trim() || 'Payroll Taxes:Employee',
      deductionParent: $('acctDedParent').value.trim() || 'Payroll wages and tax to pay',
      netPay: $('acctNet').value.trim() || 'Payroll wages and tax to pay:Wages to pay',
      payTaxPayable: $('acctTaxPay').value.trim() || 'Payroll wages and tax to pay:Payroll tax to pay',
    };
  }

  // ---------- Process ----------
  function runProcess() {
    showErr('');
    try {
      const norm = PJCore.normalise(state.matrix);
      if (!norm.rows.length) throw new Error('No payroll lines found after dropping spacer rows.');
      const physRules = readPhysicianRules();
      const classified = PJCore.classify(norm.rows, physRules);
      const accounts = readAccounts();
      const summary = PJCore.summarise(classified);
      const totals = PJCore.totals(classified);
      const entries = PJCore.buildJournalEntries(classified, { accounts });
      const validation = PJCore.validate(classified, entries);
      const checkDates = PJCore.distinctCheckDates(classified);
      const employeeCount = new Set(classified.map((r) => r.name)).size;

      state.result = {
        classified, summary, totals, entries, validation,
        dropped: norm.dropped, checkDates, employeeCount,
        client: ($('client').value.trim() || 'Payroll'),
      };
      renderResults();
      $('resultsCard').classList.remove('hidden');
      $('resultsCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (e) {
      showErr(e.message);
    }
  }

  // ---------- Rendering ----------
  function tile(k, v, cls) {
    return '<div class="tile ' + (cls || '') + '"><div class="k">' + k + '</div><div class="v">' + v + '</div></div>';
  }
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, (c) =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  function renderResults() {
    const r = state.result;
    const allPass = r.validation.checks.every((c) => c.pass);
    const je0 = r.entries[0];

    // tiles
    $('tiles').innerHTML = [
      tile('Employees', r.employeeCount),
      tile('Total earnings', '$' + state.fmt2(r.totals.earning), 'good'),
      tile('Net pay', '$' + state.fmt2(r.totals.net), 'good'),
      tile('Balances', allPass ? 'OK' : 'CHECK', allPass ? 'good' : 'warnv'),
    ].join('');

    // summary table
    let sh = '<thead><tr><th>Type</th><th>Dept</th><th class="num">Lines</th><th class="num">Earning</th><th class="num">Deductions</th><th class="num">EE Tax</th><th class="num">ER Tax</th><th class="num">Net Pay</th></tr></thead><tbody>';
    r.summary.forEach((s) => {
      sh += '<tr><td>' + esc(s.type) + '</td><td>' + esc(s.dept) + '</td><td class="num">' + s.count + '</td><td class="num">' + state.fmt2(s.earning) + '</td><td class="num">' + state.fmt2(s.deductions) + '</td><td class="num">' + state.fmt2(s.eeTax) + '</td><td class="num">' + state.fmt2(s.erTax) + '</td><td class="num">' + state.fmt2(s.net) + '</td></tr>';
    });
    sh += '<tr><td><b>TOTAL</b></td><td></td><td class="num"><b>' + r.classified.length + '</b></td><td class="num"><b>' + state.fmt2(r.totals.earning) + '</b></td><td class="num"><b>' + state.fmt2(r.totals.deductions) + '</b></td><td class="num"><b>' + state.fmt2(r.totals.eeTax) + '</b></td><td class="num"><b>' + state.fmt2(r.totals.erTax) + '</b></td><td class="num"><b>' + state.fmt2(r.totals.net) + '</b></td></tr>';
    sh += '</tbody>';
    $('summaryTable').innerHTML = sh;

    // journal entry table(s)
    let jh = '';
    r.entries.forEach((e) => {
      jh += '<div class="je-memo">' + esc(e.memo) + '</div>';
      jh += '<div class="tablewrap"><table class="data"><thead><tr><th>Account</th><th class="num">Debit</th><th class="num">Credit</th></tr></thead><tbody>';
      e.lines.forEach((ln) => {
        jh += '<tr><td>' + esc(ln.account) + '</td><td class="num">' + (ln.debit != null ? state.fmt2(ln.debit) : '') + '</td><td class="num">' + (ln.credit != null ? state.fmt2(ln.credit) : '') + '</td></tr>';
      });
      const bal = Math.abs(e.totalDebit - e.totalCredit) < 0.005;
      jh += '<tr><td><b>TOTAL</b></td><td class="num"><b>' + state.fmt2(e.totalDebit) + '</b></td><td class="num"><b>' + state.fmt2(e.totalCredit) + '</b></td></tr>';
      jh += '</tbody></table></div>';
      jh += '<div class="je-check ' + (bal ? 'ok' : 'bad') + '">' + (bal ? '✓ Debits = Credits ($' + state.fmt2(e.totalDebit) + ')' : '✗ Out of balance by $' + state.fmt2(Math.abs(e.totalDebit - e.totalCredit))) + '</div>';
    });
    $('jeBlocks').innerHTML = jh;

    // validation table
    let vh = '<thead><tr><th>Check</th><th>Value</th><th>Status</th></tr></thead><tbody>';
    r.validation.checks.forEach((c) => {
      vh += '<tr><td>' + esc(c.label) + '</td><td>' + esc(c.value) + '</td><td><span class="badge ' + (c.pass ? 'ok">✓ PASS' : 'bad">✗ FAIL') + '</span></td></tr>';
    });
    vh += '</tbody>';
    $('validationTable').innerHTML = vh;

    // exceptions
    const ex = r.validation.unclassified.concat(r.validation.negatives);
    const exBox = $('exceptionsBox');
    if (ex.length) {
      exBox.classList.remove('hidden');
      exBox.innerHTML = '<b>Exceptions to review</b><ul>' +
        r.validation.unclassified.map((u) => '<li>Unclassified: ' + esc(u.name) + ' — ' + esc(u.desc) + '</li>').join('') +
        r.validation.negatives.map((u) => '<li>Negative amount: ' + esc(u.name) + ' — ' + esc(u.desc) + '</li>').join('') +
        '</ul>';
    } else {
      exBox.classList.remove('hidden');
      exBox.innerHTML = '<b>✓ No exceptions</b> — no unclassified lines, no negative amounts, no dropped data lost.';
    }

    void je0;
  }

  // ---------- Download ----------
  async function download() {
    try {
      const r = state.result;
      const wb = PJXlsx.buildWorkbook(ExcelJS, {
        entries: r.entries, summary: r.summary, totals: r.totals,
        validation: r.validation, dropped: r.dropped,
        checkDates: r.checkDates, employeeCount: r.employeeCount,
      }, { client: r.client });
      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const datePart = (r.checkDates[0] || 'export').replace(/\//g, '-');
      const name = (r.client || 'Payroll').replace(/[^\w\-]+/g, '_') + '_Payroll_JE_' + datePart + '.xlsx';
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = name; document.body.appendChild(a); a.click();
      setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
    } catch (e) {
      showErr('Could not build the workbook: ' + e.message);
    }
  }

  // ---------- Wire up ----------
  function wireDrop(dropId, inputId, handler) {
    const drop = $(dropId), input = $(inputId);
    input.addEventListener('change', (e) => { if (e.target.files[0]) handler(e.target.files[0]); });
    ['dragover', 'dragenter'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('dragover'); }));
    ['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('dragover'); }));
    drop.addEventListener('drop', (e) => { if (e.dataTransfer.files[0]) handler(e.dataTransfer.files[0]); });
  }

  document.addEventListener('DOMContentLoaded', () => {
    wireDrop('dropReg', 'regInput', handleFile);
    $('processBtn').addEventListener('click', runProcess);
    $('downloadBtn').addEventListener('click', download);
    // advanced settings toggle
    const adv = $('advToggle');
    if (adv) adv.addEventListener('click', () => { $('advPanel').classList.toggle('hidden'); });
  });
})();
