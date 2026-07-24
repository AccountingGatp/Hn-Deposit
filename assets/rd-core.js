/*
 * rd-core.js — Remote Deposit Check Extractor: parsing engine.
 *
 * Pure, DOM-free logic so it can be unit-tested under Node. It takes the OCR
 * result of a single scanned page (plain text + optional word boxes) and pulls
 * out the three fields we care about for a remote-deposit slip:
 *
 *     Date        -> the check date
 *     Description -> the check / reference number
 *     Amount      -> the check amount
 *
 * A scanned check PDF is messy: the same check often appears twice (a remittance
 * / EOB voucher on top and the negotiable check below), backs of checks carry no
 * money at all, and the check number may be a labelled "CHECK NO.", an unlabelled
 * number in the top-right corner, or only the MICR line at the bottom. The parser
 * is deliberately best-effort and marks anything it is unsure about with a review
 * flag — a human confirms the value against the on-screen image before export.
 */
(function (root) {
  'use strict';

  var MONTHS = {
    jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
    jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12
  };

  function clean(s) {
    return String(s == null ? '' : s).replace(/\r/g, '').replace(/ /g, ' ');
  }

  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  /* ---- Amount --------------------------------------------------------- */

  // A currency-looking token: optional $, optional * padding, then digits with
  // optional thousands separators and exactly two decimals (e.g. 4,346.17).
  var MONEY = '([0-9]{1,3}(?:,[0-9]{3})+\\.[0-9]{2}|[0-9]+\\.[0-9]{2})';

  function toNumber(raw) {
    var v = parseFloat(String(raw).replace(/,/g, ''));
    return isNaN(v) ? null : v;
  }

  function allMoney(text) {
    var re = new RegExp(MONEY, 'g'), m, out = [];
    while ((m = re.exec(text))) {
      var v = toNumber(m[1]);
      if (v != null) out.push({ raw: m[1], value: v, index: m.index });
    }
    return out;
  }

  // Pick the check amount. Prefer a value sitting next to an AMOUNT / PAY /
  // CHECK AMT label; then the largest $-prefixed value; then the largest value.
  function parseAmount(text) {
    text = clean(text);
    var labelled = new RegExp(
      '(?:CHECK\\s*AM(?:OUN)?T|AMOUNT|PAY)\\b[^0-9\\n]{0,40}?\\$?\\s*\\*{0,}\\s*' + MONEY, 'i'
    );
    var m = text.match(labelled);
    if (m) return { value: toNumber(m[1]), raw: m[1], source: 'label' };

    var dollar = new RegExp('\\$\\s*\\*{0,}\\s*' + MONEY, 'g'), d, best = null;
    while ((d = dollar.exec(text))) {
      var v = toNumber(d[1]);
      if (v != null && (!best || v > best.value)) best = { value: v, raw: d[1], source: 'dollar' };
    }
    if (best) return best;

    var any = allMoney(text);
    if (!any.length) return null;
    any.sort(function (a, b) { return b.value - a.value; });
    return { value: any[0].value, raw: any[0].raw, source: 'max' };
  }

  // The written / legal amount ("FORTY FOUR AND 40/100", "Two Thousand Seven
  // Hundred Dollars and 00/100"). Words survive scanning far better than the
  // cramped courtesy box, so this is our cross-check on the numeric amount.
  var NUMWORD = {
    zero: 0, one: 1, two: 2, three: 3, four: 4, five: 5, six: 6, seven: 7,
    eight: 8, nine: 9, ten: 10, eleven: 11, twelve: 12, thirteen: 13,
    fourteen: 14, fifteen: 15, sixteen: 16, seventeen: 17, eighteen: 18,
    nineteen: 19, twenty: 20, thirty: 30, forty: 40, fifty: 50, sixty: 60,
    seventy: 70, eighty: 80, ninety: 90
  };

  function wordsToNumber(tokens) {
    var result = 0, current = 0, saw = false;
    tokens.forEach(function (w) {
      if (NUMWORD[w] != null) { current += NUMWORD[w]; saw = true; }
      else if (w === 'hundred') { current = (current || 1) * 100; saw = true; }
      else if (w === 'thousand') { result += (current || 1) * 1000; current = 0; saw = true; }
      else if (w === 'million') { result += (current || 1) * 1000000; current = 0; saw = true; }
    });
    return saw ? result + current : null;
  }

  function parseWrittenAmount(text) {
    var low = clean(text).toLowerCase().replace(/[^a-z0-9/ \n]/g, ' ');
    var lines = low.split('\n');
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      var hasSlash = /\d{1,2}\s*\/\s*100/.test(line);
      var hasDollars = /\bdollars?\b/.test(line);
      var hasCents = /\bcents?\b/.test(line);
      if (!hasSlash && !hasDollars && !hasCents) continue;

      // Cents: "40/100", or "and No cents", or spelled out "and Forty One Cents".
      var cents = null, cm = line.match(/(\d{1,2})\s*\/\s*100/);
      if (cm) cents = parseInt(cm[1], 10);
      else if (/\b(?:and\s+)?no\s+(?:\/?\s*100|cents?)/.test(line)) cents = 0;
      else if (hasCents) {
        var cw = line.match(/\band\s+([a-z ]+?)\s+cents?/);
        if (cw) { var cv = wordsToNumber(cw[1].split(/[^a-z]+/).filter(Boolean)); if (cv != null) cents = cv; }
      }

      // Dollar words: everything before "dollars" (or before the "/100").
      var head = hasDollars ? line.split(/\bdollars?\b/)[0] : line.split(/\d{1,2}\s*\/\s*100/)[0];
      var tokens = head.split(/[^a-z]+/).filter(Boolean).filter(function (w) { return w !== 'and'; });
      var dollars = wordsToNumber(tokens);
      if (dollars == null) continue;
      return { value: Math.round((dollars + (cents || 0) / 100) * 100) / 100, source: 'written' };
    }
    return null;
  }

  var AMT_RANK = { agree: 0, written: 1, label: 2, dollar: 3, max: 4 };
  function amtRank(src) { return AMT_RANK[src] != null ? AMT_RANK[src] : 5; }

  /* ---- Date ----------------------------------------------------------- */

  var MDY = '(0?[1-9]|1[0-2])[\\/\\-](0?[1-9]|[12][0-9]|3[01])[\\/\\-]((?:19|20)?[0-9]{2})';
  var MONTH_WORD =
    '(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|' +
    'Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)' +
    '\\s+([0-3]?[0-9]),?\\s*((?:19|20)[0-9]{2})';

  function normYear(y) {
    y = parseInt(y, 10);
    if (y < 100) y += 2000;
    return y;
  }

  function buildDate(mo, day, yr) {
    mo = parseInt(mo, 10); day = parseInt(day, 10); yr = normYear(yr);
    if (mo < 1 || mo > 12 || day < 1 || day > 31) return null;
    return { mo: mo, day: day, yr: yr, mdY: pad2(mo) + '/' + pad2(day) + '/' + yr };
  }

  function matchAnyDate(segment) {
    var mw = new RegExp(MONTH_WORD, 'i').exec(segment);
    if (mw) {
      var mo = MONTHS[mw[1].slice(0, 3).toLowerCase()];
      var d = buildDate(mo, mw[2], mw[3]);
      if (d) return d;
    }
    var md = new RegExp(MDY).exec(segment);
    if (md) {
      var d2 = buildDate(md[1], md[2], md[3]);
      if (d2) return d2;
    }
    return null;
  }

  // Prefer a date next to a check-date label; otherwise the first plausible date.
  // The source records which label matched so enrichment can favour the stronger
  // "CHECK DATE" over a generic "DATE".
  var DATE_LABELS = [
    { re: 'CHECK\\s*DATE', src: 'checkdate' },
    { re: 'ISSUE\\s*DATE', src: 'issuedate' },
    { re: 'DATE\\s*CHECK\\s*ISSUED', src: 'dateissued' },
    { re: 'DATE', src: 'date' }
  ];
  function dateRank(src) {
    return { checkdate: 0, issuedate: 1, dateissued: 1, date: 2, scan: 3 }[src] != null
      ? { checkdate: 0, issuedate: 1, dateissued: 1, date: 2, scan: 3 }[src] : 4;
  }
  function parseDate(text) {
    text = clean(text);
    for (var i = 0; i < DATE_LABELS.length; i++) {
      var re = new RegExp(DATE_LABELS[i].re + '\\b[^0-9A-Za-z\\n]{0,12}(' +
        MONTH_WORD + '|' + MDY + ')', 'i');
      var m = text.match(re);
      if (m) {
        var d = matchAnyDate(m[1]);
        if (d) { d.source = DATE_LABELS[i].src; return d; }
      }
    }
    var any = matchAnyDate(text);
    if (any) { any.source = 'scan'; return any; }
    return null;
  }

  /* ---- Check / reference number -------------------------------------- */

  function normKey(s) {
    return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  }

  // A labelled check number. Capture the remainder of the line after the label
  // and keep the leading alphanumeric run (allowing a single internal space, as
  // in "S2 70002553").
  function labelledCheckNo(text) {
    var re = /CHECK\s*(?:NO\.?|NUMBER|#)\s*[:.]?\s*([A-Za-z0-9][A-Za-z0-9 \-]{2,24})/i;
    var m = text.match(re);
    if (!m) return null;
    var tail = m[1];
    // Keep first token, optionally a second numeric token separated by one space.
    var mm = tail.match(/^([A-Za-z]{0,3}\d[\dA-Za-z]*)(?:\s(\d[\d]*))?/);
    if (!mm) return null;
    var val = mm[1] + (mm[2] ? ' ' + mm[2] : '');
    val = val.replace(/\s+/g, ' ').trim();
    if (normKey(val).length < 4) return null;
    return val;
  }

  // Numbers that are explicitly something else (TIN, NPI, contract, phone,
  // routing, payee id) so the positional finder never mistakes them for a check
  // number on a remittance box.
  function labeledNumbers(text) {
    var set = {}, re = /(?:TIN|NPI|CONTRACT|PHONE|ROUTING|PAYEE\s*NUMBER)\s*[:#.]?\s*([A-Za-z]?\d[\d ]{3,})/ig, m;
    while ((m = re.exec(text))) set[m[1].replace(/\D/g, '')] = true;
    return set;
  }

  // Unlabelled: a prominent number token in the top-right region of the page.
  function positionalCheckNo(page) {
    if (!page || !page.words || !page.width || !page.height) return null;
    var W = page.width, H = page.height, cands = [];
    var excluded = labeledNumbers(clean(page.text));
    for (var i = 0; i < page.words.length; i++) {
      var w = page.words[i];
      var t = (w.text || '').trim();
      if (!/^[A-Za-z]{0,2}\d{5,12}$/.test(t)) continue;      // 5-12 digit id
      if (/\.\d{2}$/.test(t)) continue;                       // not money
      if (excluded[t.replace(/\D/g, '')]) continue;           // TIN/NPI/etc.
      var bb = w.bbox || {};
      var cx = ((bb.x0 || 0) + (bb.x1 || 0)) / 2;
      var cy = ((bb.y0 || 0) + (bb.y1 || 0)) / 2;
      if (cx < W * 0.5) continue;                             // right half only
      if (cy > H * 0.5) continue;                             // top half only
      cands.push({ text: t, cy: cy, cx: cx, conf: w.conf == null ? 0 : w.conf });
    }
    if (!cands.length) return null;
    // Highest on the page wins (check number usually sits at the very top-right).
    cands.sort(function (a, b) { return a.cy - b.cy || b.cx - a.cx; });
    return cands[0].text.replace(/\s+/g, '').trim();
  }

  // Last resort: the MICR line. For these business checks the auxiliary on-us
  // field (the first digit group) is the check number.
  function micrCheckNo(text) {
    var lines = clean(text).split('\n');
    for (var i = lines.length - 1; i >= 0; i--) {
      var ln = lines[i];
      var digitCount = (ln.match(/\d/g) || []).length;
      if (digitCount < 12) continue;                          // MICR lines are digit-dense
      if (/\.\d{2}\b/.test(ln)) continue;                     // skip amount lines
      var groups = ln.match(/\d{5,}/g);
      if (groups && groups.length) return groups[0];
    }
    return null;
  }

  function parseCheckNumber(page) {
    var text = clean(page && page.text);
    var v = labelledCheckNo(text);
    if (v) return { value: v, source: 'label', confident: true };
    v = positionalCheckNo(page);
    if (v) return { value: v, source: 'position', confident: false };
    v = micrCheckNo(text);
    if (v) return { value: v, source: 'micr', confident: false };
    return null;
  }

  // Is there a MICR-style code line on the page? Every negotiable check face has
  // one (routing/account groups of 8–9 digits). Requiring a run of >=7 digits
  // separates a true MICR line from the shorter digit groups on remittance grids.
  function hasMicrLine(text) {
    var lines = clean(text).split('\n');
    for (var i = 0; i < lines.length; i++) {
      var ln = lines[i];
      if (/\.\d{2}\b/.test(ln)) continue;
      // A real MICR line carries several long digit groups (aux on-us, routing,
      // account). A lone TIN/NPI on a remittance page has just one, so require
      // at least two groups of >=5 digits with one of them >=7.
      var groups = (ln.match(/\d{5,}/g) || []);
      if (groups.length >= 2 && groups.some(function (g) { return g.length >= 7; })) return true;
    }
    return false;
  }

  // A loose identity key for a check number: its longest digit run with leading
  // zeros stripped. Lets "0000470852", "00470852" and a MICR "…470852…" match,
  // and ignores alpha prefixes like the "S2" in "S2 70002553".
  function matchKey(s) {
    var digits = String(s || '').match(/\d{4,}/g);
    if (!digits || !digits.length) return '';
    digits.sort(function (a, b) { return b.length - a.length; });
    return digits[0].replace(/^0+/, '') || digits[0];
  }

  // A page is treated as a check (front face or its remittance voucher) only if
  // it has money AND some check-identifying evidence: a labelled check number, a
  // labelled amount, or a MICR line. This keeps generic EOB line-item pages —
  // which also contain dollar figures — from becoming phantom rows.
  function isCheckLike(page, amount) {
    var text = clean(page && page.text);
    if (!amount) return false;
    if (labelledCheckNo(text)) return true;
    if (amount.source === 'label') return true;
    if (hasMicrLine(text)) return true;
    return false;
  }

  /* ---- Per-page extraction ------------------------------------------- */

  // Turn one OCR'd page into a candidate check record, or null if the page has
  // no money on it (backs of checks, blank pages, endorsement stubs).
  function extractCheck(page, meta) {
    page = page || {};
    var numeric = parseAmount(page.text);         // courtesy box (may be null)
    var written = parseWrittenAmount(page.text);  // legal/word amount (may be null)
    if (!numeric && !written) return null;         // no money => not a check face

    // Decide the amount. Words are the legal amount on a real check: if figures
    // and words disagree, trust the words and flag; if the figures were lost in
    // OCR (a dropped decimal point), the words rescue the row on their own.
    var value, amtSource, amtMismatch = false;
    if (numeric && written) {
      if (Math.abs(written.value - numeric.value) < 0.005) { value = numeric.value; amtSource = 'agree'; }
      else { value = written.value; amtSource = 'written'; amtMismatch = true; }
    } else if (numeric) {
      value = numeric.value; amtSource = numeric.source;
    } else {
      value = written.value; amtSource = 'written';
    }

    if (!isCheckLike(page, { source: amtSource })) return null;  // EOB / summary page, skip
    var date = parseDate(page.text);
    var chk = parseCheckNumber(page);
    var checkNo = chk ? chk.value : '';

    return {
      file: meta && meta.file || '',
      pageIndex: meta && meta.pageIndex != null ? meta.pageIndex : null,
      pageLabel: meta && meta.pageLabel || '',
      thumb: meta && meta.thumb || null,
      isFace: hasMicrLine(page.text),
      matchKey: matchKey(checkNo || (chk && chk.source === 'micr' ? chk.value : '')),
      date: date ? date.mdY : '',
      dateObj: date ? { mo: date.mo, day: date.day, yr: date.yr } : null,
      dateSource: date ? date.source : null,
      checkNo: checkNo,
      checkNoSource: chk ? chk.source : null,
      checkNoConfident: chk ? chk.confident : false,
      amount: value,
      amountRaw: numeric ? numeric.raw : (written ? String(written.value) : ''),
      amountSource: amtSource,
      amountMismatch: amtMismatch
    };
  }

  // Turn every candidate page of ONE file into final records. The reliable unit
  // is a check face (has a MICR line); a check's remittance voucher rides on the
  // same or a neighbouring page and carries the clean labelled values. So: make
  // one record per face, then enrich it with labelled check-number / date /
  // amount from any candidate sharing its identity key. If no faces are found
  // (e.g. a remittance-only PDF), fall back to plain check-number de-duplication.
  function assembleRecords(candidates) {
    candidates = (candidates || []).filter(Boolean);
    var faces = candidates.filter(function (c) { return c.isFace; });
    if (!faces.length) return dedupeRecords(candidates);

    // Collapse duplicate face scans of the same check.
    var byKey = {}, order = [];
    faces.forEach(function (f) {
      var k = f.matchKey || ('_p' + f.pageIndex);
      if (!byKey[k]) { byKey[k] = f; order.push(k); }
      else byKey[k] = mergeRecords(byKey[k], f);
    });

    return order.map(function (k) {
      var face = byKey[k];
      candidates.forEach(function (c) {
        if (c === face) return;
        if (!c.matchKey || c.matchKey !== face.matchKey) return;
        face = enrichFromVoucher(face, c);
      });
      return face;
    });
  }

  // Fill a face's gaps from a matching voucher/label candidate. Labelled and
  // human-trusted values win over positional/MICR guesses.
  function enrichFromVoucher(face, v) {
    var out = {};
    for (var k in face) out[k] = face[k];
    if (v.checkNoSource === 'label' && out.checkNoSource !== 'label') {
      out.checkNo = v.checkNo; out.checkNoSource = 'label'; out.checkNoConfident = true;
    }
    if (v.date && (dateRank(v.dateSource) < dateRank(out.dateSource) || !out.date)) {
      out.date = v.date; out.dateObj = v.dateObj; out.dateSource = v.dateSource;
    }
    if (amtRank(v.amountSource) < amtRank(out.amountSource)) {
      out.amount = v.amount; out.amountRaw = v.amountRaw;
      out.amountSource = v.amountSource; out.amountMismatch = v.amountMismatch;
    }
    return out;
  }

  /* ---- De-duplication ------------------------------------------------- */

  // The same physical check can appear on several pages (voucher + check face).
  // Collapse records that share a check number; then drop any number-less record
  // whose (amount,date) already exists on a numbered one.
  function dedupeRecords(records) {
    var byKey = {}, order = [], noKey = [];

    records.forEach(function (r) {
      var key = normKey(r.checkNo);
      if (!key) { noKey.push(r); return; }
      if (!byKey[key]) { byKey[key] = r; order.push(key); }
      else { byKey[key] = mergeRecords(byKey[key], r); }
    });

    var out = order.map(function (k) { return byKey[k]; });

    noKey.forEach(function (r) {
      var dup = out.some(function (o) {
        return o.amount === r.amount && o.date && r.date && o.date === r.date;
      });
      if (!dup) out.push(r);
    });

    return out;
  }

  function mergeRecords(a, b) {
    // Prefer confident/label sources and non-empty fields; keep a stable base.
    var out = {};
    for (var k in a) out[k] = a[k];
    if (!out.date && b.date) { out.date = b.date; out.dateObj = b.dateObj; out.dateSource = b.dateSource; }
    if ((!out.checkNoConfident && b.checkNoConfident) || (!out.checkNo && b.checkNo)) {
      out.checkNo = b.checkNo; out.checkNoSource = b.checkNoSource; out.checkNoConfident = b.checkNoConfident;
    }
    if (amtRank(b.amountSource) < amtRank(out.amountSource)) {
      out.amount = b.amount; out.amountRaw = b.amountRaw;
      out.amountSource = b.amountSource; out.amountMismatch = b.amountMismatch;
    }
    if (!out.thumb && b.thumb) { out.thumb = b.thumb; out.pageIndex = b.pageIndex; out.pageLabel = b.pageLabel; }
    return out;
  }

  /* ---- Review flags --------------------------------------------------- */

  function reviewFlag(r) {
    var problems = [];
    if (!r.checkNo) problems.push('CHECK NUMBER MISSING');
    else if (!r.checkNoConfident) problems.push('CHECK NUMBER UNCLEAR');
    if (!r.date) problems.push('DATE MISSING');
    if (r.amount == null) problems.push('AMOUNT MISSING');
    else if (r.amountMismatch) problems.push('AMOUNT UNCLEAR – VERIFY');
    if (!problems.length) return { status: 'OK', label: 'OK', detail: '' };
    return { status: 'RED', label: 'RED FLAG – ' + problems.join(' / '), detail: problems.join(', ') };
  }

  var API = {
    parseAmount: parseAmount,
    parseDate: parseDate,
    parseCheckNumber: parseCheckNumber,
    extractCheck: extractCheck,
    dedupeRecords: dedupeRecords,
    assembleRecords: assembleRecords,
    matchKey: matchKey,
    reviewFlag: reviewFlag,
    normKey: normKey,
    isCheckLike: isCheckLike,
    _internal: { labelledCheckNo: labelledCheckNo, micrCheckNo: micrCheckNo, matchAnyDate: matchAnyDate, hasMicrLine: hasMicrLine }
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = API;
  root.RDCore = API;
})(typeof self !== 'undefined' ? self : this);
