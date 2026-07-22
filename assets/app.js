/* app.js — UI controller. Ties the DOM to Recon (logic) and ReconXlsx (workbook). */
(function () {
  'use strict';

  const $ = (id) => document.getElementById(id);
  const state = {
    batch: null, batchName: '',   // [{batchNo,mid,total,location,closeDate,submitDate,settleDate,status}]
    txn: null, txnName: '',       // [{date,number,patient,location,paymentMethod,type,result,amount,mid,batchNo}]
    dep: null, depName: '',       // [{date,description,received,payee}]
    result: null, opts: null,
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
    return rows.filter((r) => r.some((v) => String(v).trim() !== ''));
  }

  // exact header match first, then substring
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

  const g = (row, i) => (i === -1 ? '' : (row[i] == null ? '' : String(row[i]).trim()));

  // ---------- file → matrix ----------
  function readMatrix(file, cb, onErr) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext === 'csv') {
      const reader = new FileReader();
      reader.onload = (e) => { try { cb(parseCSV(e.target.result)); } catch (err) { onErr(err); } };
      reader.onerror = () => onErr(new Error('Could not read ' + file.name));
      reader.readAsText(file);
    } else {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: 'array' });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
          cb(matrix.filter((r) => r.some((v) => String(v).trim() !== '')));
        } catch (err) { onErr(err); }
      };
      reader.onerror = () => onErr(new Error('Could not read ' + file.name));
      reader.readAsArrayBuffer(file);
    }
  }

  // ---------- parsers per file type ----------
  function parseBatch(matrix) {
    if (!matrix.length) throw new Error('The batch file appears to be empty.');
    const h = matrix[0];
    const iBatch = findCol(h, ['batch #', 'batch number', 'batch']);
    const iMid = findCol(h, ['mid']);
    const iTotal = findCol(h, ['total amount', 'amount', 'total']);
    const iLoc = findCol(h, ['location']);
    const iClose = findCol(h, ['batch close date', 'close date']);
    const iSubmit = findCol(h, ['batch submit date', 'submit date']);
    const iSettle = findCol(h, ['expected settlement date', 'settlement date', 'settlement']);
    const iStatus = findCol(h, ['status']);
    if (iBatch === -1 || iTotal === -1) {
      throw new Error('Batch file: could not find "Batch #" / "Total Amount" columns. Found: ' + h.join(', '));
    }
    const out = [];
    for (let r = 1; r < matrix.length; r++) {
      const row = matrix[r];
      const batchNo = g(row, iBatch);
      if (!batchNo) continue;
      out.push({
        batchNo, mid: g(row, iMid), total: g(row, iTotal), location: g(row, iLoc),
        closeDate: g(row, iClose), submitDate: g(row, iSubmit), settleDate: g(row, iSettle),
        status: g(row, iStatus),
      });
    }
    if (!out.length) throw new Error('Batch file: no batch rows found.');
    return out;
  }

  function parseTxn(matrix) {
    if (!matrix.length) throw new Error('The transaction export appears to be empty.');
    const h = matrix[0];
    const iDate = findCol(h, ['date']);
    const iNum = findCol(h, ['number']);
    const iPatient = findCol(h, ['patient']);
    const iAmt = findCol(h, ['amount']);
    const iPm = findCol(h, ['payment method']);
    const iType = findCol(h, ['transaction']);
    const iResult = findCol(h, ['result']);
    const iLoc = findCol(h, ['location name', 'location']);
    const iMid = findCol(h, ['mid']);
    const iBatch = findCol(h, ['batch #', 'batch']);
    if (iAmt === -1 || iBatch === -1) {
      throw new Error('Transaction export: could not find "Amount" / "Batch #" columns. Found: ' + h.join(', '));
    }
    const out = [];
    for (let r = 1; r < matrix.length; r++) {
      const row = matrix[r];
      const date = g(row, iDate), num = g(row, iNum), amt = g(row, iAmt);
      if (date === '' && num === '' && amt === '') continue;
      out.push({
        date, number: num, patient: g(row, iPatient), location: g(row, iLoc),
        paymentMethod: g(row, iPm), type: g(row, iType), result: g(row, iResult),
        amount: amt, mid: g(row, iMid), batchNo: g(row, iBatch),
      });
    }
    if (!out.length) throw new Error('Transaction export: no transaction rows found.');
    return out;
  }

  function parseDep(matrix) {
    if (!matrix.length) throw new Error('The bank deposit register appears to be empty.');
    const h = matrix[0];
    const iDate = findCol(h, ['date']);
    const iDesc = findCol(h, ['description', 'memo', 'details']);
    const iRecv = findCol(h, ['received', 'credit', 'deposit', 'amount']);
    const iPayee = findCol(h, ['payee']);
    if (iDate === -1 || iDesc === -1 || iRecv === -1) {
      throw new Error('Bank deposit register: could not find "Date" / "Description" / "Received" columns. Found: ' + h.join(', '));
    }
    const out = [];
    for (let r = 1; r < matrix.length; r++) {
      const row = matrix[r];
      const date = g(row, iDate), desc = g(row, iDesc), recv = g(row, iRecv);
      if (date === '' && desc === '' && recv === '') continue;
      if (recv === '') continue; // skip non-deposit lines with no Received amount
      out.push({ date, description: desc, received: recv, payee: g(row, iPayee) });
    }
    if (!out.length) throw new Error('Bank deposit register: no deposit rows found.');
    return out;
  }

  // ---------- file handlers ----------
  function makeHandler(kind, labelId, parser, storeKey, nameKey) {
    return function (file) {
      state[nameKey] = file.name;
      $(labelId).textContent = '✓ ' + file.name;
      readMatrix(file, (matrix) => {
        try { state[storeKey] = parser(matrix); showErr(''); }
        catch (e) { state[storeKey] = null; showErr(e.message); }
        refreshReady();
      }, (e) => { state[storeKey] = null; showErr(e.message); refreshReady(); });
    };
  }
  const handleBatch = makeHandler('batch', 'batchFileName', parseBatch, 'batch', 'batchName');
  const handleTxn = makeHandler('txn', 'txnFileName', parseTxn, 'txn', 'txnName');
  const handleDep = makeHandler('dep', 'depFileName', parseDep, 'dep', 'depName');

  function refreshReady() {
    const ok = state.batch && state.txn && state.dep;
    $('processBtn').disabled = !ok;
    $('readyHint').textContent = ok
      ? state.batch.length + ' batches · ' + state.txn.length + ' transactions · ' + state.dep.length + ' deposits loaded.'
      : 'Upload all three files to enable.';
  }

  function showErr(msg) {
    const box = $('errBox');
    if (!msg) { box.classList.add('hidden'); box.textContent = ''; return; }
    box.classList.remove('hidden'); box.textContent = '⚠ ' + msg;
  }

  // ---------- process ----------
  function runProcess() {
    showErr('');
    try {
      const res = Recon.reconcile(state.batch, state.txn, state.dep);
      res.EPS = Recon.EPS;
      state.result = res;
      state.opts = { client: ($('client').value.trim() || 'Client') };
      renderResults();
      $('resultsCard').classList.remove('hidden');
      $('resultsCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (e) {
      showErr('Reconciliation failed: ' + e.message);
    }
  }

  // ---------- rendering ----------
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, (c) =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  function tile(k, v, cls) {
    return '<div class="tile ' + (cls || '') + '"><div class="k">' + k + '</div><div class="v">' + v + '</div></div>';
  }

  function renderResults() {
    const R = state.result, T = R.totals, C = R.checks;
    const allOk = C.splitsTieToBatches && C.depositsTieToBatches && C.allDepositsMatched;

    // tiles
    $('tiles').innerHTML = [
      tile('Batches', T.batchCount),
      tile('Batch total', '$' + state.fmt2(T.batchTotal), 'good'),
      tile('Deposited / awaiting', T.depositedBatchCount + ' / ' + T.awaitingBatchCount),
      tile('Reconciles', allOk ? 'OK' : 'CHECK', allOk ? 'good' : 'warnv'),
    ].join('');

    // note box
    const ab = $('assumpBox');
    ab.classList.remove('hidden');
    ab.innerHTML = '<b>How the tie-out works.</b> ' +
      'Recon 1 sums each batch’s transaction splits and compares to the batch total. ' +
      'Recon 2 matches every batch to a bank deposit by <b>MID + amount</b>. ' +
      'Batches with no deposit yet are <b>awaiting settlement</b> (a timing difference, not an error); ' +
      'their ' + T.awaitingBatchCount + ' batch(es) total $' + state.fmt2(T.awaitingBatchTotal) + '. ' +
      T.noBatchTxnCount + ' transaction(s) totalling $' + state.fmt2(T.noBatchTxnTotal) +
      ' carry no batch number (authorizations, cash/check, or declined) and are listed on the Transactions sheet.';

    // summary table
    const money = (n) => state.fmt2(n);
    const chk = (b) => '<span class="badge ' + (b ? 'ok">✓ OK' : 'bad">✗ CHECK') + '</span>';
    const secRow = (t) => '<tr class="sec"><td colspan="3">' + esc(t) + '</td></tr>';
    const dRow = (l, c, a, isChk) => '<tr><td>' + l + '</td><td class="num">' + (c === '' ? '' : c) +
      '</td><td class="num">' + (isChk ? a : money(a)) + '</td></tr>';
    let s = '<thead><tr><th>Particulars</th><th class="num">Count</th><th class="num">Amount</th></tr></thead><tbody>';
    s += secRow('FILE TOTALS');
    s += dRow('Batch file — settlement batches', T.batchCount, T.batchTotal);
    s += dRow('Transaction export — card transactions (splits)', T.txnCount, T.txnTotal);
    s += dRow('Bank deposit register — deposits', T.depositCount, T.depositTotal);
    s += secRow('RECONCILIATION 1 — TRANSACTIONS → BATCHES');
    s += dRow('Transactions that roll into a batch', T.batchedTxnCount, T.batchedTxnTotal);
    s += dRow('Transactions with no batch (auth / cash / check / declined)', T.noBatchTxnCount, T.noBatchTxnTotal);
    s += dRow('Sum of batch totals', T.batchCount, T.batchTotal);
    s += dRow('<b>CHECK: batch totals = sum of their splits</b>', '', chk(C.splitsTieToBatches), true);
    s += secRow('RECONCILIATION 2 — BATCHES → BANK DEPOSITS');
    s += dRow('Batches settled &amp; deposited in the bank', T.depositedBatchCount, T.depositedBatchTotal);
    s += dRow('Batches submitted, awaiting deposit (timing)', T.awaitingBatchCount, T.awaitingBatchTotal);
    s += dRow('Bank deposits matched to a batch', T.matchedDepositCount, T.matchedDepositTotal);
    s += dRow('Bank deposits with no matching batch', T.unmatchedDepositCount, T.unmatchedDepositTotal);
    s += dRow('<b>CHECK: deposited-batch total = matched deposit total</b>', '', chk(C.depositsTieToBatches), true);
    s += dRow('<b>CHECK: every bank deposit is matched to a batch</b>', '', chk(C.allDepositsMatched), true);
    s += '</tbody>';
    $('summaryTable').innerHTML = s;

    // batch table
    $('batchCount').textContent = T.batchCount;
    let b = '<thead><tr><th>Batch #</th><th>MID</th><th>Location</th><th>Settle</th>' +
      '<th class="num">Batch Total</th><th class="num">Txns</th><th class="num">Split Sum</th>' +
      '<th class="num">Var.</th><th>Deposited</th><th>Status</th></tr></thead><tbody>';
    R.batches.forEach((x) => {
      const bad = Math.abs(x.variance) > R.EPS;
      const cls = bad ? ' class="needs"' : '';
      b += '<tr' + cls + '><td>' + esc(x.batchNo) + '</td><td>' + esc(x.mid) + '</td>' +
        '<td>' + esc(x.location) + '</td><td>' + esc(x.settleDate) + '</td>' +
        '<td class="num">' + money(x.batchTotal) + '</td><td class="num">' + x.txnCount + '</td>' +
        '<td class="num">' + money(x.txnSum) + '</td><td class="num">' + money(x.variance) + '</td>' +
        '<td>' + (x.deposited ? 'Yes' : '<span class="amber">No</span>') + '</td>' +
        '<td>' + esc(x.reconStatus) + '</td></tr>';
    });
    b += '</tbody>';
    $('batchTable').innerHTML = b;

    // exceptions table
    $('excCount').textContent = R.exceptions.length ? R.exceptions.length : 'none 🎉';
    let e = '<thead><tr><th>Type</th><th>Reference</th><th class="num">Amount</th><th>Explanation</th></tr></thead><tbody>';
    if (R.exceptions.length) {
      R.exceptions.forEach((x) => {
        e += '<tr><td>' + esc(x.type) + '</td><td>' + esc(x.reference) + '</td>' +
          '<td class="num">' + money(x.amount) + '</td><td class="reason">' + esc(x.explanation) + '</td></tr>';
      });
    } else {
      e += '<tr><td>None</td><td></td><td class="num"></td><td class="reason">All three files reconcile. ' +
        'Every batch total equals the sum of its splits, and every bank deposit matches a batch.</td></tr>';
    }
    e += '</tbody>';
    $('excTable').innerHTML = e;

    const badge = $('dlBadge');
    if (allOk) { badge.className = 'badge ok'; badge.textContent = '✓ All checks pass'; }
    else { badge.className = 'badge bad'; badge.textContent = 'Review exceptions before use'; }
  }

  // ---------- download ----------
  async function download() {
    try {
      const wb = ReconXlsx.buildWorkbook(ExcelJS, state.result, state.opts);
      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const range = (state.result.dateLabel || '').replace(/\//g, '-').replace(/\s+/g, '');
      const name = (state.opts.client || 'Client').replace(/[^\w\-]+/g, '_') +
        '_Deposit_Reconciliation_' + (range || 'export') + '.xlsx';
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = name; document.body.appendChild(a); a.click();
      setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
    } catch (e) {
      showErr('Could not build the workbook: ' + e.message);
    }
  }

  // ---------- wire up ----------
  function wireDrop(dropId, inputId, handler) {
    const drop = $(dropId), input = $(inputId);
    input.addEventListener('change', (e) => { if (e.target.files[0]) handler(e.target.files[0]); });
    ['dragover', 'dragenter'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('dragover'); }));
    ['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('dragover'); }));
    drop.addEventListener('drop', (e) => { if (e.dataTransfer.files[0]) handler(e.dataTransfer.files[0]); });
  }

  document.addEventListener('DOMContentLoaded', () => {
    wireDrop('dropBatch', 'batchInput', handleBatch);
    wireDrop('dropTxn', 'txnInput', handleTxn);
    wireDrop('dropDep', 'depInput', handleDep);
    $('processBtn').addEventListener('click', runProcess);
    $('downloadBtn').addEventListener('click', download);
  });
})();
