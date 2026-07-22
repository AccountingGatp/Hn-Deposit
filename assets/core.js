/*
 * core.js — Three-file Deposit Reconciliation engine (DOM-free, testable in Node & browser)
 *
 * Reconciles three files that describe the same money at three stages:
 *   A) Batch file            — one row per settlement batch, with a Total Amount.
 *   B) Transaction export    — the individual card transactions (the "splits")
 *                              that roll up into each batch (linked by Batch #).
 *   C) Bank deposit register — the amounts that actually landed in the bank
 *                              (one row per MID per settlement, RECEIVED amount).
 *
 * Two ties are checked:
 *   Recon 1  (A <-> B)  batch Total Amount  ==  sum of its transaction splits.
 *   Recon 2  (A <-> C)  each batch settles as a bank deposit (match by MID + amount);
 *                       leftover batches are "awaiting deposit", leftover deposits
 *                       are exceptions.
 *
 * No DOM, no I/O — pure functions so the same logic runs in the browser and under
 * Node for unit tests.
 */
(function (root) {
  'use strict';

  const EPS = 0.005; // money tolerance (half a cent)

  // ---- Amount parsing -------------------------------------------------------
  // Handles: "$1,762.71", "($45.00)", "-$45.00", "44.07", "$40.00 ", "-45.00"
  function parseAmount(raw) {
    if (raw === null || raw === undefined) return 0;
    let s = String(raw).trim();
    if (s === '') return 0;
    let neg = false;
    if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
    s = s.replace(/[$,\s]/g, '');
    if (s.endsWith('-')) { neg = true; s = s.slice(0, -1); }
    if (s.startsWith('-')) { neg = true; s = s.slice(1); }
    if (s === '' || isNaN(Number(s))) return 0;
    const n = Number(s);
    return neg ? -n : n;
  }

  function round2(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }

  // ---- MID extraction from a bank-deposit description -----------------------
  // "Phreesia SV9T/8886547473 MID011041237 011041237 Healthnomic" -> "011041237"
  function extractMID(desc) {
    if (!desc) return '';
    const m = String(desc).match(/MID\s*(\d{4,})/i);
    if (m) return m[1];
    const m2 = String(desc).match(/\b(\d{6,})\b/); // fallback: first long number
    return m2 ? m2[1] : '';
  }

  // Normalise a batch number / MID for keying.
  function key(v) { return String(v == null ? '' : v).trim(); }

  // ---- US date parsing (accepts a trailing time) ----------------------------
  function parseUSDate(s) {
    if (!s) return null;
    const m = String(s).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (!m) return null;
    let [, mo, da, yr] = m;
    yr = yr.length === 2 ? '20' + yr : yr;
    const d = new Date(Number(yr), Number(mo) - 1, Number(da));
    return isNaN(d.getTime()) ? null : d;
  }

  function fmtDate(d) {
    return (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
  }

  // date part only (strip a trailing time) for display
  function dateOnly(s) {
    const m = String(s || '').trim().match(/^(\d{1,2}\/\d{1,2}\/\d{2,4})/);
    return m ? m[1] : String(s || '').trim();
  }

  function dateRange(rowsOfDateStrings) {
    const ds = rowsOfDateStrings.map(parseUSDate).filter(Boolean).sort((a, b) => a - b);
    if (!ds.length) return { min: null, max: null, label: '' };
    return { min: ds[0], max: ds[ds.length - 1], label: fmtDate(ds[0]) + ' to ' + fmtDate(ds[ds.length - 1]) };
  }

  // ---- Plain-language note for a transaction with no batch number -----------
  function noBatchNote(t) {
    const type = String(t.type || '').toLowerCase();
    const result = String(t.result || '').toLowerCase();
    const pm = String(t.paymentMethod || '').toLowerCase();
    if (/declin|insufficient|fail/.test(result)) {
      return 'Declined / failed transaction — no money was captured, so it never settled into a batch.';
    }
    if (/auth/.test(type) || (round2(t.amount) === 0 && /credit application|authorization/.test(type))) {
      return 'Authorization / $0 credit application — no funds were captured, so it is not part of any batch.';
    }
    if (/cash/.test(pm)) {
      return 'Cash payment — collected at the office and deposited outside the card processor, so it is not in a card batch.';
    }
    if (/check|cheque/.test(pm)) {
      return 'Check payment — deposited outside the card processor, so it is not in a card batch.';
    }
    if (/void/.test(type)) {
      return 'Void — reverses a sale before settlement, so it is not carried in a batch.';
    }
    return 'No batch number on this transaction — it did not settle into a card batch (e.g. authorization, cash/check, or declined).';
  }

  // ---- Main reconciliation --------------------------------------------------
  // batchRows:   [{ batchNo, mid, total(raw or num), location, closeDate, submitDate, settleDate, status }]
  // txnRows:     [{ date, number, patient, location, paymentMethod, type, result, amount(raw or num), mid, batchNo }]
  // depositRows: [{ date, description, mid?, received(raw or num), payee }]
  function reconcile(batchRows, txnRows, depositRows) {
    // --- normalise transactions & group by batch # ---
    const txnByBatch = new Map();
    const transactions = txnRows.map((t) => {
      const amount = typeof t.amount === 'number' ? t.amount : parseAmount(t.amount);
      const batchNo = key(t.batchNo);
      const rec = {
        date: dateOnly(t.date), number: key(t.number), patient: String(t.patient || '').trim(),
        location: String(t.location || '').trim(), paymentMethod: String(t.paymentMethod || '').trim(),
        type: String(t.type || '').trim(), result: String(t.result || '').trim(),
        amount: amount, mid: key(t.mid), batchNo: batchNo,
        batched: batchNo !== '', note: '',
      };
      if (batchNo) {
        if (!txnByBatch.has(batchNo)) txnByBatch.set(batchNo, []);
        txnByBatch.get(batchNo).push(rec);
      } else {
        rec.note = noBatchNote(rec);
      }
      return rec;
    });

    // --- normalise deposits, build a match pool keyed by MID|amount ---
    const deposits = depositRows.map((d) => {
      const received = typeof d.received === 'number' ? d.received : parseAmount(d.received);
      const mid = key(d.mid) || extractMID(d.description);
      return {
        date: dateOnly(d.date), description: String(d.description || '').trim(),
        mid: mid, received: received, matchedBatch: '', status: '',
      };
    });
    const depPool = new Map(); // "mid|amt" -> [indices, in file order]
    deposits.forEach((d, i) => {
      const k = d.mid + '|' + round2(d.received).toFixed(2);
      if (!depPool.has(k)) depPool.set(k, []);
      depPool.get(k).push(i);
    });

    // --- batches: tie to splits (A<->B) then to a deposit (A<->C) ---
    const batchNosSeen = new Set();
    const batches = batchRows.map((b) => {
      const batchNo = key(b.batchNo);
      const mid = key(b.mid);
      const batchTotal = typeof b.total === 'number' ? b.total : parseAmount(b.total);
      const splits = txnByBatch.get(batchNo) || [];
      const txnSum = round2(splits.reduce((a, s) => a + s.amount, 0));
      const variance = round2(batchTotal - txnSum);
      batchNosSeen.add(batchNo);
      return {
        batchNo, mid, batchTotal,
        location: String(b.location || '').trim(),
        closeDate: dateOnly(b.closeDate), submitDate: dateOnly(b.submitDate),
        settleDate: dateOnly(b.settleDate), status: String(b.status || '').trim(),
        txnCount: splits.length, txnSum, variance,
        deposited: false, depositAmount: 0, depositDate: '', reconStatus: '',
      };
    });

    // match each batch to a deposit (stable, by batchNo order preserved from file)
    for (const b of batches) {
      const k = b.mid + '|' + round2(b.batchTotal).toFixed(2);
      const pool = depPool.get(k);
      if (pool && pool.length) {
        const idx = pool.shift();
        deposits[idx].matchedBatch = b.batchNo;
        deposits[idx].status = 'Matched to batch';
        b.deposited = true;
        b.depositAmount = deposits[idx].received;
        b.depositDate = deposits[idx].date;
      }
      // recon status
      if (Math.abs(b.variance) > EPS) b.reconStatus = 'Split mismatch';
      else if (b.deposited) b.reconStatus = 'Settled & deposited';
      else b.reconStatus = 'Awaiting deposit';
    }
    deposits.forEach((d) => { if (!d.matchedBatch) d.status = 'No matching batch'; });

    // orphan transactions: reference a batch # that is not in the batch file
    for (const t of transactions) {
      if (t.batched && !batchNosSeen.has(t.batchNo)) {
        t.note = 'References batch ' + t.batchNo + ' which is not in the batch file.';
      }
    }

    // --- exceptions (true anomalies; awaiting-deposit is a timing bucket, listed too) ---
    const exceptions = [];
    for (const b of batches) {
      if (Math.abs(b.variance) > EPS) {
        exceptions.push({
          type: 'Batch split mismatch', reference: b.batchNo, amount: b.variance,
          explanation: 'Batch total ' + b.batchTotal.toFixed(2) + ' does not equal the sum of its ' +
            b.txnCount + ' transaction split(s) (' + b.txnSum.toFixed(2) + '). Difference ' +
            b.variance.toFixed(2) + '. Check for a missing or extra transaction in the export.',
        });
      }
    }
    for (const d of deposits) {
      if (!d.matchedBatch) {
        exceptions.push({
          type: 'Deposit with no batch', reference: d.description || ('MID ' + d.mid), amount: d.received,
          explanation: 'A bank deposit of ' + d.received.toFixed(2) + ' for MID ' + d.mid +
            ' has no batch in the batch file with the same MID and amount. It may belong to a batch ' +
            'outside this file\'s date range, or the amounts differ.',
        });
      }
    }
    const orphanTxns = transactions.filter((t) => t.batched && !batchNosSeen.has(t.batchNo));
    for (const t of orphanTxns) {
      exceptions.push({
        type: 'Transaction references unknown batch', reference: t.number + ' → ' + t.batchNo,
        amount: t.amount,
        explanation: 'Transaction ' + t.number + ' (' + t.amount.toFixed(2) + ') references batch ' +
          t.batchNo + ', which is not present in the batch file.',
      });
    }

    // --- totals ---
    const sum = (arr, f) => round2(arr.reduce((a, x) => a + f(x), 0));
    const batchedTxns = transactions.filter((t) => t.batched);
    const noBatchTxns = transactions.filter((t) => !t.batched);
    const depositedBatches = batches.filter((b) => b.deposited);
    const awaitingBatches = batches.filter((b) => !b.deposited && Math.abs(b.variance) <= EPS);
    const matchedDeposits = deposits.filter((d) => d.matchedBatch);
    const unmatchedDeposits = deposits.filter((d) => !d.matchedBatch);

    const totals = {
      batchCount: batches.length, batchTotal: sum(batches, (b) => b.batchTotal),
      txnCount: transactions.length, txnTotal: sum(transactions, (t) => t.amount),
      depositCount: deposits.length, depositTotal: sum(deposits, (d) => d.received),
      batchedTxnCount: batchedTxns.length, batchedTxnTotal: sum(batchedTxns, (t) => t.amount),
      noBatchTxnCount: noBatchTxns.length, noBatchTxnTotal: sum(noBatchTxns, (t) => t.amount),
      depositedBatchCount: depositedBatches.length, depositedBatchTotal: sum(depositedBatches, (b) => b.batchTotal),
      awaitingBatchCount: awaitingBatches.length, awaitingBatchTotal: sum(awaitingBatches, (b) => b.batchTotal),
      matchedDepositCount: matchedDeposits.length, matchedDepositTotal: sum(matchedDeposits, (d) => d.received),
      unmatchedDepositCount: unmatchedDeposits.length, unmatchedDepositTotal: sum(unmatchedDeposits, (d) => d.received),
      splitMismatchCount: batches.filter((b) => Math.abs(b.variance) > EPS).length,
      totalVariance: sum(batches, (b) => b.variance),
    };

    // checks (booleans; the workbook re-derives these as live formulas)
    const checks = {
      splitsTieToBatches: totals.splitMismatchCount === 0,
      depositsTieToBatches: Math.abs(totals.depositedBatchTotal - totals.matchedDepositTotal) <= EPS,
      allDepositsMatched: totals.unmatchedDepositCount === 0,
    };

    const dLabel = dateRange(
      batches.map((b) => b.closeDate).concat(deposits.map((d) => d.date))
    ).label;

    return { batches, transactions, deposits, exceptions, totals, checks, dateLabel: dLabel };
  }

  const api = {
    parseAmount, round2, extractMID, parseUSDate, dateRange, dateOnly, fmtDate,
    noBatchNote, reconcile, EPS,
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.Recon = api;
})(typeof self !== 'undefined' ? self : this);
