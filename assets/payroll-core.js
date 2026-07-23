/*
 * payroll-core.js — Payroll Register → QBO Journal Entry core logic
 * (DOM-free, testable in Node & browser).
 *
 * Implements the GATP prompt "Eldersburg Payroll Journal Entry":
 * normalise a payroll register export, classify each line, summarise by
 * Type × Department, build a balanced journal entry per check date, and
 * validate it — all as pure functions so the same logic runs in the browser
 * and under Node for unit tests.
 */
(function (root) {
  'use strict';

  // ---- Number parsing -------------------------------------------------------
  // "2,193.50", "$1,450.20", "(31.81)", "", " 0.00 " → number (blank → 0)
  function num(raw) {
    if (raw === null || raw === undefined) return 0;
    let s = String(raw).replace(/"/g, '').trim();
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

  function clean(raw) {
    return String(raw === null || raw === undefined ? '' : raw).replace(/"/g, '').trim();
  }

  // ---- Column map -----------------------------------------------------------
  // The register export is positional (F1..F23). Column index = F-number − 1.
  // If the sheet has real headers we match by header name instead.
  const POS = {
    f2: 1,   // presence flag — drop row if blank
    name: 2, // F3
    gross: 4, // F5
    pstart: 5, // F6
    pend: 6,  // F7
    checkdate: 7, // F8
    checkno: 8,   // F9
    type2: 9,     // F10
    type1: 10,    // F11 — D / E / T
    desc: 11,     // F12
    hrs: 12,      // F13
    rate: 13,     // F14
    earning: 14,  // F15
    deductions: 17, // F18
    eeTaxable: 18,  // F19
    erTaxable: 19,  // F20
    eeTax: 20,      // F21
    erTax: 21,      // F22
    checkno2: 22,   // F23
  };

  const HEADER_ALIASES = {
    name: ['employee name', 'name', 'employee'],
    gross: ['rate / gross', 'gross', 'rate/gross'],
    pstart: ['pay period start', 'period start', 'period begin'],
    pend: ['pay period end', 'period end'],
    checkdate: ['check date', 'checkdate'],
    checkno: ['check number', 'check no', 'check #'],
    type2: ['type 2', 'type2'],
    type1: ['type1', 'type 1', 'type'],
    desc: ['description', 'desc', 'item'],
    hrs: ['hours', 'hrs'],
    rate: ['rate'],
    earning: ['earning', 'earnings'],
    deductions: ['deductions', 'deduction'],
    eeTaxable: ['employee taxable', 'ee taxable'],
    erTaxable: ['employer taxable', 'er taxable'],
    eeTax: ['employee tax', 'ee tax'],
    erTax: ['employer tax', 'er tax'],
  };

  function looksLikeHeader(row) {
    const joined = row.map((c) => clean(c).toLowerCase()).join(' ');
    return /employee/.test(joined) && /(earning|tax|deduction)/.test(joined);
  }

  function buildColMap(headerRow) {
    if (!headerRow || !looksLikeHeader(headerRow)) return Object.assign({}, POS);
    const H = headerRow.map((c) => clean(c).toLowerCase());
    const map = {};
    for (const key of Object.keys(HEADER_ALIASES)) {
      let idx = -1;
      for (const alias of HEADER_ALIASES[key]) {
        idx = H.findIndex((h) => h === alias);
        if (idx === -1) idx = H.findIndex((h) => h.indexOf(alias) !== -1);
        if (idx !== -1) break;
      }
      if (idx !== -1) map[key] = idx;
    }
    // presence flag: any non-blank in the row → keep. Use name column.
    map.f2 = map.name !== undefined ? map.name : POS.f2;
    // fill any missing keys from positional defaults
    for (const k of Object.keys(POS)) if (map[k] === undefined) map[k] = POS[k];
    return map;
  }

  // ---- Normalise the register ----------------------------------------------
  // matrix: array of arrays (already read from the sheet).
  // Returns { rows:[recs], dropped:count, colMap, hadHeader:bool }
  function normalise(matrix) {
    const rows = [];
    let dropped = 0;
    let start = 0;
    let colMap = Object.assign({}, POS);
    let hadHeader = false;
    if (matrix.length && looksLikeHeader(matrix[0])) {
      colMap = buildColMap(matrix[0]);
      hadHeader = true;
      start = 1;
    }
    for (let r = start; r < matrix.length; r++) {
      const raw = matrix[r];
      if (!raw) continue;
      const at = (k) => raw[colMap[k]];
      // Drop rule: F2 (presence flag) blank → header / spacer / total.
      if (clean(at('f2')) === '') { dropped++; continue; }
      rows.push({
        name: clean(at('name')),
        gross: clean(at('gross')),
        pstart: clean(at('pstart')),
        pend: clean(at('pend')),
        checkdate: clean(at('checkdate')),
        checkno: clean(at('checkno')),
        type2: clean(at('type2')),
        type1: clean(at('type1')).toUpperCase(),
        desc: clean(at('desc')),
        hrs: num(at('hrs')),
        rate: num(at('rate')),
        earning: num(at('earning')),
        deductions: num(at('deductions')),
        eeTaxable: num(at('eeTaxable')),
        erTaxable: num(at('erTaxable')),
        eeTax: num(at('eeTax')),
        erTax: num(at('erTax')),
      });
    }
    return { rows, dropped, colMap, hadHeader };
  }

  // ---- Classify -------------------------------------------------------------
  // Bucket order (first match wins), per prompt §2.
  function bucketOf(rec) {
    if (rec.eeTax > 0) return 'Employee Tax';
    if (rec.erTax > 0) return 'Employer Tax';
    if (rec.type1 === 'E') return 'Earning';
    if (rec.type1 === 'D') return 'Deduction';
    if (rec.type1 === 'T') return 'Employee Tax';
    return 'Unclassified';
  }

  // Department from the employee name. Physician rules are configurable so new
  // physician names can be added without touching code.
  function departmentOf(name, physicianRules) {
    const rules = physicianRules || defaultPhysicianRules();
    const n = clean(name);
    for (const r of rules) {
      if (r.type === 'startsWith' && n.toLowerCase().startsWith(r.value.toLowerCase())) return 'Physicians';
      if (r.type === 'contains' && n.toLowerCase().indexOf(r.value.toLowerCase()) !== -1) return 'Physicians';
    }
    return 'Employee';
  }

  function defaultPhysicianRules() {
    return [
      { type: 'startsWith', value: 'OBrian' },
      { type: 'contains', value: 'Niculescu' },
    ];
  }

  function classify(rows, physicianRules) {
    return rows.map((rec) => {
      const bucket = bucketOf(rec);
      const dept = departmentOf(rec.name, physicianRules);
      const net = rec.earning - rec.deductions - rec.eeTax;
      return Object.assign({}, rec, { bucket, dept, net });
    });
  }

  // ---- Summary: Type × Department ------------------------------------------
  function summarise(classified) {
    const map = new Map();
    const order = ['Earning', 'Deduction', 'Employee Tax', 'Employer Tax', 'Unclassified'];
    for (const r of classified) {
      const key = r.bucket + '||' + r.dept;
      if (!map.has(key)) {
        map.set(key, { type: r.bucket, dept: r.dept, earning: 0, deductions: 0, eeTax: 0, erTax: 0, net: 0, count: 0 });
      }
      const e = map.get(key);
      e.earning += r.earning; e.deductions += r.deductions;
      e.eeTax += r.eeTax; e.erTax += r.erTax; e.net += r.net; e.count += 1;
    }
    return Array.from(map.values()).sort((a, b) => {
      const oa = order.indexOf(a.type), ob = order.indexOf(b.type);
      if (oa !== ob) return oa - ob;
      return a.dept.localeCompare(b.dept);
    });
  }

  // ---- Company totals -------------------------------------------------------
  function totals(classified) {
    const t = { earning: 0, deductions: 0, eeTax: 0, erTax: 0, net: 0 };
    for (const r of classified) {
      t.earning += r.earning; t.deductions += r.deductions;
      t.eeTax += r.eeTax; t.erTax += r.erTax; t.net += r.net;
    }
    return t;
  }

  // Deductions grouped by description (each becomes its own credit line).
  function deductionsByDesc(classified) {
    const map = new Map();
    for (const r of classified) {
      if (r.deductions > 0) {
        const key = r.desc || 'Deduction';
        map.set(key, (map.get(key) || 0) + r.deductions);
      }
    }
    return Array.from(map.entries()).map(([desc, amount]) => ({ desc, amount }));
  }

  // ---- Journal entry (one per check date) ----------------------------------
  // opts: { accounts:{ wagesEmployee, wagesPhysicians, employerTax,
  //          payTaxPayable, deductionParent, netPay }, deductionLabels }
  function defaultAccounts() {
    return {
      wagesEmployee: 'Payroll Wages & Salaries:Employee',
      wagesPhysicians: 'Payroll Wages & Salaries:Physicians',
      employerTax: 'Payroll Taxes:Employee',
      deductionParent: 'Payroll wages and tax to pay',
      netPay: 'Payroll wages and tax to pay:Wages to pay',
      payTaxPayable: 'Payroll wages and tax to pay:Payroll tax to pay',
    };
  }

  // Friendly label for a deduction description → account leaf name.
  function deductionLeaf(desc) {
    const d = clean(desc).toUpperCase();
    if (d === '401K') return '401k deduction';
    if (d === 'ROTH' || d === '401KR') return 'Roth deduction';
    return clean(desc) + ' deduction';
  }

  function round2(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }

  function buildJournalEntries(classified, opts) {
    const acct = Object.assign(defaultAccounts(), (opts && opts.accounts) || {});
    // group by check date, keeping period info
    const groups = new Map();
    for (const r of classified) {
      const key = r.checkdate || '(no date)';
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(r);
    }
    const entries = [];
    for (const [checkdate, recs] of groups) {
      const t = totals(recs);
      const earnEmp = recs.filter((r) => r.dept === 'Employee').reduce((a, b) => a + b.earning, 0);
      const earnPhys = recs.filter((r) => r.dept === 'Physicians').reduce((a, b) => a + b.earning, 0);
      const deds = deductionsByDesc(recs);
      const pstart = (recs[0] && recs[0].pstart) || '';
      const pend = (recs[0] && recs[0].pend) || '';
      const memo = 'Payroll — pay period ' + pstart + ' to ' + pend + ', check date ' + checkdate;

      const lines = [];
      if (earnPhys > 0) lines.push({ account: acct.wagesPhysicians, debit: round2(earnPhys), credit: null });
      if (earnEmp > 0) lines.push({ account: acct.wagesEmployee, debit: round2(earnEmp), credit: null });
      if (t.erTax > 0) lines.push({ account: acct.employerTax, debit: round2(t.erTax), credit: null });
      deds.forEach((d) => {
        lines.push({ account: acct.deductionParent + ':' + deductionLeaf(d.desc), debit: null, credit: round2(d.amount) });
      });
      if (t.net !== 0) lines.push({ account: acct.netPay, debit: null, credit: round2(t.net) });
      const taxPayable = t.eeTax + t.erTax;
      if (taxPayable > 0) lines.push({ account: acct.payTaxPayable, debit: null, credit: round2(taxPayable) });

      const totalDebit = round2(lines.reduce((a, b) => a + (b.debit || 0), 0));
      const totalCredit = round2(lines.reduce((a, b) => a + (b.credit || 0), 0));
      entries.push({ checkdate, pstart, pend, memo, lines, totalDebit, totalCredit, totals: t, earnEmp: round2(earnEmp), earnPhys: round2(earnPhys), deductions: deds });
    }
    return entries;
  }

  // ---- Validation -----------------------------------------------------------
  function validate(classified, entries) {
    const t = totals(classified);
    const checks = [];
    const near = (a, b) => Math.abs(round2(a) - round2(b)) < 0.005;

    entries.forEach((e) => {
      checks.push({
        label: 'JE ' + e.checkdate + ' — Debits = Credits',
        value: round2(e.totalDebit) + ' vs ' + round2(e.totalCredit),
        pass: near(e.totalDebit, e.totalCredit),
      });
    });

    const lhs = t.earning + t.erTax;                 // Earning + Employer Tax
    const rhs = t.net + t.eeTax + t.erTax + t.deductions; // Net + EE + ER + Deductions
    checks.push({
      label: 'Identity: Earning + Employer Tax = Net + EE Tax + ER Tax + Deductions',
      value: round2(lhs) + ' = ' + round2(rhs),
      pass: near(lhs, rhs),
    });

    const unclassified = classified.filter((r) => r.bucket === 'Unclassified');
    checks.push({ label: 'Unclassified lines', value: unclassified.length, pass: unclassified.length === 0 });

    const negatives = classified.filter((r) => Math.min(r.earning, r.deductions, r.eeTax, r.erTax) < 0);
    checks.push({ label: 'Negative amounts', value: negatives.length, pass: negatives.length === 0 });

    return { checks, totals: t, unclassified, negatives };
  }

  // ---- Distinct meta helpers ------------------------------------------------
  function distinctCheckDates(classified) {
    return Array.from(new Set(classified.map((r) => r.checkdate).filter(Boolean)));
  }
  function periodLabel(classified) {
    const p = classified.find((r) => r.pstart || r.pend);
    return p ? (p.pstart + ' to ' + p.pend) : '';
  }

  const api = {
    num, clean, normalise, classify, bucketOf, departmentOf, defaultPhysicianRules,
    summarise, totals, deductionsByDesc, deductionLeaf, defaultAccounts,
    buildJournalEntries, validate, distinctCheckDates, periodLabel, round2,
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.PJCore = api;
})(typeof self !== 'undefined' ? self : this);
