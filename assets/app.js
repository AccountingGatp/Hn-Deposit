/* app.js — UI controller. Ties the DOM to BDCore (logic) and BDXlsx (workbook). */
(function () {
  'use strict';

  const $ = (id) => document.getElementById(id);
  const state = {
    bankRows: null,     // [{date, description, amountRaw}]
    bankName: '',
    customers: null,    // [names]
    custName: '',
    result: null,       // BDCore.process output
    dateLabel: '',
    fmt2: (n) => (isNaN(n) ? '' : n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })),
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
    return rows.filter((r) => r.some((v) => String(v).trim() !== ''));
  }

  function findCol(headers, candidates) {
    const H = headers.map((h) => String(h).trim().toLowerCase());
    for (const cand of candidates) {
      const idx = H.findIndex((h) => h === cand);
      if (idx !== -1) return idx;
    }
    for (const cand of candidates) {
      const idx = H.findIndex((h) => h.indexOf(cand) !== -1);
      if (idx !== -1) return idx;
    }
    return -1;
  }

  function rowsToBank(matrix) {
    if (!matrix.length) throw new Error('The bank file appears to be empty.');
    const headers = matrix[0];
    const di = findCol(headers, ['posted date', 'date']);
    const desci = findCol(headers, ['full description', 'description', 'memo', 'details']);
    const ai = findCol(headers, ['amount', 'value']);
    if (di === -1 || desci === -1 || ai === -1) {
      throw new Error('Could not find the required columns. Expected headers like "Posted Date", "Full description", "Amount". Found: ' + headers.join(', '));
    }
    const out = [];
    for (let r = 1; r < matrix.length; r++) {
      const row = matrix[r];
      const date = (row[di] ?? '').toString().trim();
      const description = (row[desci] ?? '').toString().trim();
      const amountRaw = (row[ai] ?? '').toString().trim();
      if (date === '' && description === '' && amountRaw === '') continue;
      out.push({ date, description, amountRaw });
    }
    return out;
  }

  function sheetToMatrix(file, cb, onErr) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
        cb(matrix.filter((r) => r.some((v) => String(v).trim() !== '')));
      } catch (err) { onErr(err); }
    };
    reader.onerror = () => onErr(new Error('Could not read the file.'));
    reader.readAsArrayBuffer(file);
  }

  // ---------- File handlers ----------
  function handleBank(file) {
    state.bankName = file.name;
    $('bankFileName').textContent = '✓ ' + file.name;
    const ext = file.name.split('.').pop().toLowerCase();
    const done = (matrix) => {
      try { state.bankRows = rowsToBank(matrix); showErr(''); }
      catch (e) { state.bankRows = null; showErr(e.message); }
      refreshReady();
    };
    if (ext === 'csv') {
      const reader = new FileReader();
      reader.onload = (e) => done(parseCSV(e.target.result));
      reader.onerror = () => showErr('Could not read the bank CSV.');
      reader.readAsText(file);
    } else {
      sheetToMatrix(file, done, (e) => { showErr(e.message); });
    }
  }

  function handleCust(file) {
    state.custName = file.name;
    $('custFileName').textContent = '✓ ' + file.name;
    sheetToMatrix(file, (matrix) => {
      try {
        if (!matrix.length) throw new Error('Customers file is empty.');
        const ni = findCol(matrix[0], ['name', 'customer', 'display name']);
        const col = ni === -1 ? 0 : ni;
        const names = [];
        for (let r = 1; r < matrix.length; r++) {
          const v = (matrix[r][col] ?? '').toString().trim();
          if (v) names.push(v);
        }
        if (!names.length) throw new Error('No customer names found in the Customers file.');
        state.customers = names;
        showErr('');
      } catch (e) { state.customers = null; showErr(e.message); }
      refreshReady();
    }, (e) => { showErr(e.message); });
  }

  function refreshReady() {
    const ok = state.bankRows && state.bankRows.length && state.customers && state.customers.length;
    $('processBtn').disabled = !ok;
    $('readyHint').textContent = ok
      ? state.bankRows.length + ' bank rows · ' + state.customers.length + ' customers loaded.'
      : 'Upload both files to enable.';
  }

  function showErr(msg) {
    const box = $('errBox');
    if (!msg) { box.classList.add('hidden'); box.textContent = ''; return; }
    box.classList.remove('hidden'); box.textContent = '⚠ ' + msg;
  }

  // ---------- Process ----------
  function runProcess() {
    showErr('');
    const startSNo = parseInt($('startSNo').value, 10);
    if (isNaN(startSNo)) { showErr('Please enter a valid Starting SNo.'); return; }
    const accounts = $('accounts').value.trim();
    if (!accounts) { showErr('Please enter the Accounts value.'); return; }
    const klass = $('klass').value.trim();

    const res = BDCore.process(state.bankRows, state.customers, { startSNo, accounts, klass });
    state.dateLabel = BDCore.dateRange(state.bankRows).label;
    state.result = res;
    state.opts = { accounts, klass, client: ($('client').value.trim() || 'Client') };

    renderResults();
    $('resultsCard').classList.remove('hidden');
    $('resultsCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // ---------- Rendering ----------
  function bankNet() {
    return state.bankRows.reduce((a, b) => a + (BDCore.parseAmount(b.amountRaw) || 0), 0);
  }

  function renderResults() {
    const r = state.result;
    const procTotal = r.processed.reduce((a, b) => a + b.amount, 0);
    const npTotal = r.notProcessed.reduce((a, b) => a + b.amount, 0);
    const net = bankNet();
    const check = Math.abs(procTotal + npTotal - net) < 0.005;
    const unassigned = r.processed.filter((p) => !p.customer).length;

    // tiles
    $('tiles').innerHTML = [
      tile('Rows processed', r.processed.length),
      tile('Total deposited', '$' + state.fmt2(procTotal), 'good'),
      tile('Left / not processed', r.notProcessed.length),
      tile('Reconciles', check ? 'OK' : 'MISMATCH', check ? 'good' : 'warnv'),
    ].join('');

    // assumptions
    const ab = $('assumpBox');
    if (r.assumptions && r.assumptions.length) {
      ab.classList.remove('hidden');
      ab.innerHTML = '<b>Customer-mapping assumptions applied</b> (names chosen only from your Customers list):<ul>' +
        r.assumptions.map((a) => '<li>(' + a.count + ') ' + esc(a.note) + '</li>').join('') + '</ul>';
    } else { ab.classList.add('hidden'); }

    // reconciliation table
    const debits = r.notProcessed.filter((n) => n.amount < 0);
    const depNoTrn = r.notProcessed.filter((n) => n.amount >= 0);
    const rec = [
      ['Total transactions in bank file (' + (state.dateLabel || 'range') + ')', state.bankRows.length, net],
      ['Processed — deposits with TRN number', r.processed.length, procTotal],
      ['Left / not processed', r.notProcessed.length, npTotal],
      ['&nbsp;&nbsp;&nbsp;of which: debits / transfers out', debits.length, debits.reduce((a, b) => a + b.amount, 0)],
      ['&nbsp;&nbsp;&nbsp;of which: deposits without TRN — manual entry', depNoTrn.length, depNoTrn.reduce((a, b) => a + b.amount, 0)],
    ];
    let recHtml = '<thead><tr><th>Particulars</th><th class="num">Count</th><th class="num">Amount</th></tr></thead><tbody>';
    rec.forEach((row) => {
      recHtml += '<tr><td>' + row[0] + '</td><td class="num">' + row[1] + '</td><td class="num">' + state.fmt2(row[2]) + '</td></tr>';
    });
    recHtml += '<tr><td><b>Check: Processed + Left = Bank file total</b></td><td class="num"></td><td class="num"><span class="badge ' +
      (check ? 'ok">✓ OK' : 'bad">✗ MISMATCH') + '</span></td></tr></tbody>';
    $('reconTable').innerHTML = recHtml;

    // processed table (with editable customer selects)
    $('procCount').textContent = r.processed.length;
    const opts = state.customers.slice().sort((a, b) => a.localeCompare(b));
    let ph = '<thead><tr><th>SNo.</th><th>Date</th><th>Customer name</th><th>Accounts</th><th>Description (TRN)</th><th class="num">Amount</th><th>Memo</th><th>Class</th></tr></thead><tbody>';
    r.processed.forEach((p, idx) => {
      const unset = !p.customer;
      const sel = '<select data-idx="' + idx + '" class="custSel' + (unset ? ' unset' : '') + '">' +
        '<option value="">— select —</option>' +
        opts.map((n) => '<option' + (n === p.customer ? ' selected' : '') + '>' + esc(n) + '</option>').join('') +
        '</select>';
      ph += '<tr class="' + (unset ? 'needs' : '') + '">' +
        '<td>' + p.sno + '</td><td>' + esc(p.date) + '</td><td>' + sel + '</td>' +
        '<td>' + esc(state.opts.accounts) + '</td><td>' + esc(p.description) + '</td>' +
        '<td class="num">' + state.fmt2(p.amount) + '</td><td>' + esc(p.memo) + '</td>' +
        '<td>' + esc(state.opts.klass) + '</td></tr>';
    });
    ph += '</tbody>';
    $('procTable').innerHTML = ph;
    Array.prototype.forEach.call(document.querySelectorAll('.custSel'), (s) => {
      s.addEventListener('change', function () {
        const i = parseInt(this.getAttribute('data-idx'), 10);
        state.result.processed[i].customer = this.value || null;
        state.result.processed[i].matched = !!this.value;
        this.classList.toggle('unset', !this.value);
        this.closest('tr').classList.toggle('needs', !this.value);
        updateDlBadge();
      });
    });

    // not processed table
    $('npCount').textContent = r.notProcessed.length;
    let nh = '<thead><tr><th>Date</th><th>Bank Description</th><th class="num">Amount</th><th>Reason</th></tr></thead><tbody>';
    r.notProcessed.forEach((n) => {
      nh += '<tr><td>' + esc(n.date) + '</td><td>' + esc(n.description) + '</td>' +
        '<td class="num">' + state.fmt2(n.amount) + '</td><td class="reason">' + esc(n.reason) + '</td></tr>';
    });
    nh += '</tbody>';
    $('npTable').innerHTML = nh;

    updateDlBadge();
  }

  function updateDlBadge() {
    const left = state.result.processed.filter((p) => !p.customer).length;
    const b = $('dlBadge');
    if (left > 0) { b.className = 'badge bad'; b.textContent = left + ' row(s) still need a customer'; }
    else { b.className = 'badge ok'; b.textContent = '✓ All rows have a customer'; }
  }

  function tile(k, v, cls) {
    return '<div class="tile ' + (cls || '') + '"><div class="k">' + k + '</div><div class="v">' + v + '</div></div>';
  }
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, (c) =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  // ---------- Download ----------
  async function download() {
    try {
      const r = state.result;
      const wb = BDXlsx.buildWorkbook(ExcelJS, {
        processed: r.processed, notProcessed: r.notProcessed,
        assumptions: r.assumptions, dateLabel: state.dateLabel,
      }, state.opts);
      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const range = (state.dateLabel || '').replace(/\//g, '-').replace(/\s+/g, '');
      const name = (state.opts.client || 'Client').replace(/[^\w\-]+/g, '_') +
        '_QBO_Deposit_Import_' + (range || 'export') + '.xlsx';
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
    wireDrop('dropBank', 'bankInput', handleBank);
    wireDrop('dropCust', 'custInput', handleCust);
    $('processBtn').addEventListener('click', runProcess);
    $('downloadBtn').addEventListener('click', download);
  });
})();
