/*
 * core.js — Bank Deposit → QBO Import core logic (DOM-free, testable in Node & browser)
 *
 * Implements the GATP SOP "Bank Deposit to QBO Import Sheet (with Reconciliation)".
 * No DOM, no I/O. Pure functions so the same logic runs in the browser and under Node
 * for unit tests.
 */
(function (root) {
  'use strict';

  // ---- Amount parsing -------------------------------------------------------
  // Handles: "$7,966.40 ", "($29,800.00)", "$699.34 ", "-1234.5", "1,234.00-"
  function parseAmount(raw) {
    if (raw === null || raw === undefined) return NaN;
    let s = String(raw).trim();
    if (s === '') return NaN;
    let neg = false;
    if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
    s = s.replace(/[$,\s]/g, '');
    if (s.endsWith('-')) { neg = true; s = s.slice(0, -1); }
    if (s.startsWith('-')) { neg = true; s = s.slice(1); }
    if (s === '' || isNaN(Number(s))) return NaN;
    const n = Number(s);
    return neg ? -n : n;
  }

  // ---- TRN extraction -------------------------------------------------------
  // "first number after TRN*1* and before the next *"
  // e.g. ...TRN*1*321784050*7530078070 -> "321784050"
  function extractTRN(desc) {
    if (!desc) return null;
    const m = String(desc).match(/TRN\*1\*(\d+)\*/i);
    return m ? m[1] : null;
  }

  // ---- Default customer mapping rules --------------------------------------
  // Each rule: { pattern: RegExp, name: <exact customer name>, note: <why> }
  // First match wins. More specific rules first.
  function defaultRules() {
    return [
      { pattern: /FEP\s*Postal/i,        name: 'Fep Postal',       note: 'CFMI FEP Postal → Fep Postal' },
      { pattern: /FEP\s*Non-?Pos/i,      name: 'FEDERAL EMPLOYEE',  note: 'CFMI/CFBC FEP Non-Pos → FEDERAL EMPLOYEE' },
      { pattern: /Wellpoint/i,           name: 'Wellpoint',         note: 'Wellpoint MD5C → Wellpoint' },
      { pattern: /36\s*TREAS/i,          name: '36 TREAS',          note: '36 TREAS → 36 TREAS' },
      { pattern: /CAREFIRST|GHMSI|CFMI|CFBC/i, name: 'BCBS',        note: 'CareFirst (GHMSI/CFMI/CFBC/OF MD/BLUECH/ADVANT/MEDGAP/plain GHMSI) → BCBS (CareFirst BlueCross BlueShield)' },
    ];
  }

  // Normalise a string for loose token matching.
  function norm(s) {
    return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  }

  // Try to resolve a customer from the description.
  // Returns { name, note, matched:true } or { name:null, note, matched:false }.
  function resolveCustomer(desc, customerNames, rules) {
    const names = customerNames || [];
    const nameSet = new Set(names.map((n) => n.toLowerCase()));

    // 1) explicit rules (only accept if the target exists in the customer list)
    for (const r of rules) {
      if (r.pattern.test(desc)) {
        if (nameSet.has(r.name.toLowerCase())) {
          const exact = names.find((n) => n.toLowerCase() === r.name.toLowerCase());
          return { name: exact, note: r.note, matched: true, rule: true };
        }
        // rule matched but target not on list — surface it, still flag for review
        return { name: null, note: r.note + ' (target "' + r.name + '" not found in Customers list)', matched: false, rule: true };
      }
    }

    // 2) generic: longest customer name whose normalised form appears in the memo
    const nd = norm(desc);
    let best = null;
    for (const n of names) {
      const nn = norm(n);
      if (nn.length >= 4 && nd.indexOf(nn) !== -1) {
        if (!best || nn.length > norm(best).length) best = n;
      }
    }
    if (best) return { name: best, note: 'Matched by name found in bank memo', matched: true, rule: false };

    return { name: null, note: 'Payer not recognised — needs manual customer selection', matched: false, rule: false };
  }

  // ---- Reason text for excluded lines --------------------------------------
  function reasonFor(row) {
    if (row.amount < 0) {
      const acct = (String(row.description).match(/\*{2,}(\d+)/) || [])[1];
      return 'This is a debit (money out)' + (/transfer/i.test(row.description)
          ? ', a bank transfer out' + (acct ? ' to account ' + acct : '') : '') +
        '. We only process the deposit (money-in) side, so this line is not part of the ' +
        'QBO import. It also has no TRN reference number.';
    }
    // positive but no TRN
    return 'This is a deposit (money in), but the bank memo has no TRN reference number ' +
      '(pattern TRN*1*<number>*). Without a TRN we cannot tie it to a claim, so it needs ' +
      'manual review / entry before import.';
  }

  // ---- Main processing ------------------------------------------------------
  // bankRows: [{ date, description, amountRaw }]  (already split from CSV)
  // customerNames: array of strings (the "Name" column)
  // opts: { startSNo, accounts, klass, rules }
  function process(bankRows, customerNames, opts) {
    const rules = opts.rules || defaultRules();
    const start = parseInt(opts.startSNo, 10);
    const processed = [];
    const notProcessed = [];
    const assumptions = new Map(); // note -> count

    let sno = start;
    for (const raw of bankRows) {
      const amount = parseAmount(raw.amountRaw);
      const desc = raw.description || '';
      const trn = extractTRN(desc);
      const isDeposit = !isNaN(amount) && amount > 0;

      if (isDeposit && trn) {
        const res = resolveCustomer(desc, customerNames, rules);
        if (res.rule && res.note) assumptions.set(res.note, (assumptions.get(res.note) || 0) + 1);
        processed.push({
          sno: sno++,
          date: raw.date,
          customer: res.name,       // may be null -> needs review
          matched: res.matched,
          description: trn,          // TRN number (text)
          amount: amount,
          memo: trn,
          bankDescription: desc,
        });
      } else {
        const rowObj = { date: raw.date, description: desc, amount: isNaN(amount) ? 0 : amount };
        rowObj.reason = reasonFor(rowObj);
        notProcessed.push(rowObj);
      }
    }

    return {
      processed: processed,
      notProcessed: notProcessed,
      assumptions: Array.from(assumptions.entries()).map(([note, count]) => ({ note, count })),
    };
  }

  // ---- Date range helper ----------------------------------------------------
  function dateRange(bankRows) {
    const ds = bankRows
      .map((r) => parseUSDate(r.date))
      .filter((d) => d)
      .sort((a, b) => a - b);
    if (!ds.length) return { min: null, max: null, label: '' };
    const fmt = (d) => (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
    return { min: ds[0], max: ds[ds.length - 1], label: fmt(ds[0]) + ' to ' + fmt(ds[ds.length - 1]) };
  }

  function parseUSDate(s) {
    if (!s) return null;
    const m = String(s).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (!m) return null;
    let [, mo, da, yr] = m;
    yr = yr.length === 2 ? '20' + yr : yr;
    return new Date(Number(yr), Number(mo) - 1, Number(da));
  }

  // ---- Customer-wise counts (summary) --------------------------------------
  function customerCounts(processed) {
    const map = new Map();
    for (const r of processed) {
      const key = r.customer || '(unassigned)';
      if (!map.has(key)) map.set(key, { count: 0, amount: 0 });
      const e = map.get(key);
      e.count += 1;
      e.amount += r.amount;
    }
    return Array.from(map.entries())
      .map(([customer, v]) => ({ customer, count: v.count, amount: v.amount }))
      .sort((a, b) => b.amount - a.amount);
  }

  const api = {
    parseAmount, extractTRN, defaultRules, resolveCustomer, reasonFor,
    process, dateRange, parseUSDate, customerCounts, norm,
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.BDCore = api;
})(typeof self !== 'undefined' ? self : this);
