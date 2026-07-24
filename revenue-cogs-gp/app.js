/* Revenue · COGS · Gross & Net Profit report — GATP
   Runs entirely in the browser. Reads two QuickBooks Transaction Reports,
   builds a 3-tab P&L, and exports a styled .xlsx with live formulas.        */
(function () {
  "use strict";

  var NAVY = "FF1F3864", BAND = "FFD9E1F2", BAND2 = "FFEEF2FB";
  var state = { rev: null, cogs: null };

  /* ---------- helpers ---------- */
  var $ = function (id) { return document.getElementById(id); };
  function isNum(v) { return typeof v === "number" && isFinite(v); }
  function s(v) { return v == null ? "" : String(v).trim(); }

  // 1,234.56 ; negatives shown in (parentheses)
  function money(n) {
    if (!isNum(n)) return "";
    var neg = n < 0, a = Math.abs(n)
      .toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    return neg ? "(" + a + ")" : a;
  }
  function pct(n) {
    if (!isNum(n)) return "—";
    return n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + "%";
  }
  function esc(t) {
    return s(t).replace(/[&<>"]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
    });
  }

  /* ---------- parse a QuickBooks Transaction Report ---------- */
  function parseReport(wb) {
    var ws = wb.Sheets[wb.SheetNames[0]];
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });

    // Header row = the row that contains a cell reading "Amount"
    var hIdx = -1, header = null;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i] || [];
      for (var c = 0; c < r.length; c++) {
        if (s(r[c]).toLowerCase() === "amount") { hIdx = i; header = r; break; }
      }
      if (hIdx >= 0) break;
    }
    if (hIdx < 0) throw new Error("Couldn't find the header row (no 'Amount' column). Is this a QuickBooks Transaction Report?");

    // Locate columns by header text
    var col = { date: -1, type: -1, num: -1, name: -1, desc: -1, account: -1, amount: -1 };
    header.forEach(function (h, idx) {
      var t = s(h).toLowerCase();
      if (!t) return;
      if (col.date < 0 && /date/.test(t)) col.date = idx;
      else if (col.type < 0 && /type/.test(t)) col.type = idx;
      else if (col.num < 0 && t === "num") col.num = idx;
      else if (col.name < 0 && t === "name") col.name = idx;
      else if (col.desc < 0 && (/description/.test(t) || /memo/.test(t))) col.desc = idx;
      else if (col.account < 0 && (t === "account" || t === "full name")) col.account = idx;
      else if (col.amount < 0 && t === "amount") col.amount = idx;
    });
    if (col.amount < 0) throw new Error("Couldn't find the Amount column.");

    // Titles above the header (single-cell rows) → company + period
    var titles = [];
    for (var t2 = 0; t2 < hIdx; t2++) {
      var first = s((rows[t2] || [])[0]);
      if (first) titles.push(first);
    }
    var company = titles[0] || "Company";
    var period = titles.length ? titles[titles.length - 1] : "";

    // Data rows: column A empty AND a numeric Amount
    var data = [], grandTotal = null;
    for (var k = hIdx + 1; k < rows.length; k++) {
      var row = rows[k] || [];
      var labelA = s(row[0]);
      var amt = row[col.amount];
      if (/^total$/i.test(labelA)) { if (isNum(amt)) grandTotal = amt; continue; }
      if (labelA) continue;                    // category / subtotal / "Total for ..."
      if (!isNum(amt)) continue;               // skip blanks
      data.push({
        date: s(row[col.date]),
        type: s(row[col.type]),
        num: col.num >= 0 ? s(row[col.num]) : "",
        name: col.name >= 0 ? s(row[col.name]) : "",
        desc: col.desc >= 0 ? s(row[col.desc]) : "",
        account: col.account >= 0 ? s(row[col.account]) : "",
        amount: amt
      });
    }
    if (grandTotal == null) grandTotal = data.reduce(function (a, r) { return a + r.amount; }, 0);

    return { company: company, period: period, rows: data, total: grandTotal };
  }

  /* ---------- read a file into a workbook ---------- */
  function readFile(file, cb) {
    var fr = new FileReader();
    fr.onload = function (e) {
      try { cb(null, parseReport(XLSX.read(new Uint8Array(e.target.result), { type: "array" }))); }
      catch (err) { cb(err); }
    };
    fr.onerror = function () { cb(new Error("Could not read the file.")); };
    fr.readAsArrayBuffer(file);
  }

  /* ---------- file inputs ---------- */
  function wireDrop(inputId, dropId, nameId, key) {
    $(inputId).addEventListener("change", function () {
      var f = this.files[0]; if (!f) return;
      $(nameId).textContent = "reading…";
      readFile(f, function (err, rep) {
        if (err) {
          state[key] = null;
          $(dropId).classList.remove("set");
          $(nameId).textContent = "";
          showErr(f.name + ": " + err.message);
          refreshReady(); return;
        }
        state[key] = rep; state[key].fileName = f.name;
        $(dropId).classList.add("set");
        $(nameId).textContent = "✓ " + f.name + " · " + rep.rows.length + " rows · total " + money(rep.total);
        // Auto-fill company/period from whichever file we have
        if (!$("client").value) $("client").value = rep.company;
        if (!$("period").value) $("period").value = rep.period;
        clearErr(); refreshReady();
      });
    });
  }

  function refreshReady() {
    var ok = state.rev && state.cogs;
    $("processBtn").disabled = !ok;
    $("readyHint").textContent = ok ? "Ready — click to build." : "Upload both files to enable.";
  }
  function showErr(m) { var b = $("errBox"); b.textContent = "⚠ " + m; b.classList.remove("hidden"); }
  function clearErr() { $("errBox").classList.add("hidden"); }

  /* ---------- compute the P&L ---------- */
  function computePL() {
    var sales = state.rev.total, cogs = state.cogs.total;
    var expenses = parseFloat($("totExp").value) || 0;
    var gp = sales - cogs;
    var np = gp - expenses;
    return {
      sales: sales, cogs: cogs, gp: gp, expenses: expenses, np: np,
      gpPct: sales !== 0 ? gp / sales * 100 : null,
      npPct: sales !== 0 ? np / sales * 100 : null
    };
  }

  /* ---------- render everything ---------- */
  function render() {
    var p = computePL();
    var company = s($("client").value) || state.rev.company || "Company";
    var period = s($("period").value) || state.rev.period || "";
    var scope = (company ? company : "") + (period ? " · " + period : "");

    /* tiles */
    $("tiles").innerHTML = [
      tile("Revenue (Sales)", money(p.sales), ""),
      tile("Cost of Goods Sold", money(p.cogs), ""),
      tile("Gross Profit", money(p.gp), pct(p.gpPct) + " margin", p.gp >= 0 ? "good" : "bad"),
      tile("Net Profit", money(p.np), pct(p.npPct) + " margin", p.np >= 0 ? "good" : "bad")
    ].join("");

    /* P&L statement */
    $("plScope").textContent = scope ? "— " + scope : "";
    function moneyCell(n) { return '<td class="money' + (n < 0 ? " neg" : "") + '">' + money(n) + "</td>"; }
    $("plTable").innerHTML =
      "<tbody>" +
      "<tr><td class='label'>Revenue (Sales)</td>" + moneyCell(p.sales) + "</tr>" +
      "<tr class='sub'><td class='label'>Less: Cost of Goods Sold</td>" + moneyCell(-p.cogs) + "</tr>" +
      "<tr class='gp'><td class='label'>Gross Profit</td>" + moneyCell(p.gp) + "</tr>" +
      "<tr class='sub'><td class='label'>Less: Total Operating Expenses</td>" + moneyCell(-p.expenses) + "</tr>" +
      "<tr class='np'><td class='label'>Net Profit</td>" + moneyCell(p.np) + "</tr>" +
      "</tbody>";

    /* percentages */
    $("pctTable").innerHTML =
      "<thead><tr><th>Metric</th><th>Calculation</th><th class='num'>Result</th></tr></thead><tbody>" +
      "<tr><td class='metric'>Gross Profit Margin (GP %)</td><td class='formula'>Gross Profit ÷ Sales × 100 = " +
        money(p.gp) + " ÷ " + money(p.sales) + " × 100</td><td class='val'>" + pct(p.gpPct) + "</td></tr>" +
      "<tr><td class='metric'>Net Profit Margin (NP %)</td><td class='formula'>Net Profit ÷ Sales × 100 = " +
        money(p.np) + " ÷ " + money(p.sales) + " × 100</td><td class='val'>" + pct(p.npPct) + "</td></tr>" +
      "</tbody>";

    /* transaction tables */
    $("revCount").textContent = state.rev.rows.length + " rows · total " + money(state.rev.total);
    $("cogsCount").textContent = state.cogs.rows.length + " rows · total " + money(state.cogs.total);
    $("revTable").innerHTML = txTable(state.rev);
    $("cogsTable").innerHTML = txTable(state.cogs);

    $("resultsCard").classList.remove("hidden");
    $("dlBadge").textContent = "";
    $("resultsCard").scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function tile(k, v, sub, cls) {
    return "<div class='tile " + (cls || "") + "'><div class='k'>" + esc(k) + "</div><div class='v'>" +
      esc(v) + "</div>" + (sub ? "<div class='sub'>" + esc(sub) + "</div>" : "") + "</div>";
  }

  function txTable(rep) {
    var head = "<thead><tr><th>Date</th><th>Type</th><th>Num</th><th>Name</th>" +
      "<th>Description / Memo</th><th>Account</th><th class='num'>Amount</th></tr></thead>";
    var body = rep.rows.map(function (r) {
      return "<tr><td>" + esc(r.date) + "</td><td>" + esc(r.type) + "</td><td>" + esc(r.num) +
        "</td><td>" + esc(r.name) + "</td><td class='desc'>" + esc(r.desc) + "</td><td>" +
        esc(r.account) + "</td><td class='num'>" + money(r.amount) + "</td></tr>";
    }).join("");
    var total = "<tr class='total'><td colspan='6'>TOTAL</td><td class='num'>" + money(rep.total) + "</td></tr>";
    return head + "<tbody>" + body + total + "</tbody>";
  }

  /* ---------- tab switching ---------- */
  Array.prototype.forEach.call(document.querySelectorAll(".tab"), function (btn) {
    btn.addEventListener("click", function () {
      document.querySelectorAll(".tab").forEach(function (b) { b.classList.remove("active"); });
      document.querySelectorAll(".tabpane").forEach(function (p) { p.classList.remove("active"); });
      btn.classList.add("active");
      $("pane-" + btn.dataset.tab).classList.add("active");
    });
  });

  /* ---------- Excel export (ExcelJS, live formulas) ---------- */
  function buildTxSheet(wb, title, rep) {
    var ws = wb.addWorksheet(title, { views: [{ state: "frozen", ySplit: 5 }] });
    ws.columns = [
      { width: 13 }, { width: 12 }, { width: 12 }, { width: 22 },
      { width: 52 }, { width: 22 }, { width: 15 }
    ];
    ws.mergeCells("A1:G1"); ws.getCell("A1").value = rep.company;
    ws.getCell("A1").font = { bold: true, size: 14, color: { argb: NAVY } };
    ws.mergeCells("A2:G2"); ws.getCell("A2").value = title + " — Transaction Report";
    ws.getCell("A2").font = { bold: true, size: 11 };
    ws.mergeCells("A3:G3"); ws.getCell("A3").value = rep.period;
    ws.getCell("A3").font = { italic: true, color: { argb: "FF667085" } };

    var hdr = ["Date", "Type", "Num", "Name", "Description / Memo", "Account", "Amount"];
    var hRow = ws.getRow(5);
    hdr.forEach(function (h, i) {
      var c = hRow.getCell(i + 1);
      c.value = h;
      c.font = { bold: true, color: { argb: "FFFFFFFF" }, name: "Arial", size: 10 };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NAVY } };
      c.alignment = { vertical: "middle" };
      c.border = thin();
    });

    var start = 6, r = start;
    rep.rows.forEach(function (t, i) {
      var row = ws.getRow(r);
      [t.date, t.type, t.num, t.name, t.desc, t.account, t.amount].forEach(function (v, ci) {
        var c = row.getCell(ci + 1);
        c.value = v;
        c.font = { name: "Arial", size: 10 };
        c.border = thin();
        if (ci === 4) c.alignment = { wrapText: true, vertical: "top" };
        if (ci === 6) { c.numFmt = "#,##0.00"; c.alignment = { horizontal: "right" }; }
      });
      if (i % 2 === 1) row.eachCell(function (c) {
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BAND2 } };
      });
      r++;
    });

    var totRow = ws.getRow(r);
    totRow.getCell(1).value = "TOTAL";
    var totCell = "G" + r;
    totRow.getCell(7).value = rep.rows.length
      ? { formula: "SUM(G" + start + ":G" + (r - 1) + ")" }
      : 0;
    totRow.eachCell(function (c, cn) {
      c.font = { bold: true, name: "Arial", size: 10 };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BAND } };
      c.border = { top: { style: "medium", color: { argb: NAVY } } };
      if (cn === 7) c.numFmt = "#,##0.00";
    });
    return totCell; // e.g. "G28"
  }

  function buildPLSheet(wb, revTot, cogsTot, expenses, company, period) {
    var ws = wb.addWorksheet("Gross & Net Profit");
    ws.columns = [{ width: 34 }, { width: 18 }, { width: 42 }];
    ws.mergeCells("A1:C1"); ws.getCell("A1").value = company;
    ws.getCell("A1").font = { bold: true, size: 14, color: { argb: NAVY } };
    ws.mergeCells("A2:C2"); ws.getCell("A2").value = "Profit & Loss — Gross & Net Profit";
    ws.getCell("A2").font = { bold: true, size: 11 };
    ws.mergeCells("A3:C3"); ws.getCell("A3").value = period;
    ws.getCell("A3").font = { italic: true, color: { argb: "FF667085" } };

    // Reference the totals on the other two sheets (live cross-sheet formulas)
    var salesRef = "'Revenue'!" + revTot;      // e.g. 'Revenue'!G9
    var cogsRef = "'COGS'!" + cogsTot;

    var rowN = 5;
    function line(label, valObj, opts) {
      opts = opts || {};
      var row = ws.getRow(rowN);
      var a = row.getCell(1), b = row.getCell(2);
      a.value = label;
      b.value = valObj;
      b.numFmt = opts.pct ? '0.00"%"' : "#,##0.00";
      b.alignment = { horizontal: "right" };
      a.font = { name: "Arial", size: 10, bold: !!opts.bold };
      b.font = { name: "Arial", size: 10, bold: !!opts.bold };
      if (opts.fill) {
        [a, b].forEach(function (c) {
          c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: opts.fill } };
        });
      }
      if (opts.top) [a, b].forEach(function (c) {
        c.border = { top: { style: opts.top, color: { argb: NAVY } } };
      });
      var n = rowN; rowN++; return n;
    }

    var rSales = line("Revenue (Sales)", { formula: salesRef });
    var rCogs = line("Less: Cost of Goods Sold", { formula: cogsRef });
    var rGp = line("Gross Profit", { formula: "B" + rSales + "-B" + rCogs }, { bold: true, fill: BAND2, top: "thin" });
    var rExp = line("Less: Total Operating Expenses", expenses);
    var rNp = line("Net Profit", { formula: "B" + rGp + "-B" + rExp }, { bold: true, fill: BAND, top: "medium" });

    rowN++; // blank line
    // Section header for margins
    var mh = ws.getRow(rowN); mh.getCell(1).value = "Margin analysis (percentages)";
    mh.getCell(1).font = { bold: true, color: { argb: NAVY }, size: 11 }; rowN++;

    // GP% and NP% with transparent formulas + a calc note in column C
    var gpRow = ws.getRow(rowN);
    gpRow.getCell(1).value = "Gross Profit Margin (GP %)";
    gpRow.getCell(2).value = { formula: "IF(B" + rSales + "=0,0,B" + rGp + "/B" + rSales + "*100)" };
    gpRow.getCell(2).numFmt = '0.00"%"'; gpRow.getCell(2).alignment = { horizontal: "right" };
    gpRow.getCell(3).value = "= Gross Profit ÷ Sales × 100";
    styleRow(gpRow, true); rowN++;

    var npRow = ws.getRow(rowN);
    npRow.getCell(1).value = "Net Profit Margin (NP %)";
    npRow.getCell(2).value = { formula: "IF(B" + rSales + "=0,0,B" + rNp + "/B" + rSales + "*100)" };
    npRow.getCell(2).numFmt = '0.00"%"'; npRow.getCell(2).alignment = { horizontal: "right" };
    npRow.getCell(3).value = "= Net Profit ÷ Sales × 100";
    styleRow(npRow, true); rowN++;

    function styleRow(row, bold) {
      row.getCell(1).font = { name: "Arial", size: 10, bold: bold };
      row.getCell(2).font = { name: "Arial", size: 11, bold: true, color: { argb: NAVY } };
      row.getCell(3).font = { name: "Arial", size: 9, italic: true, color: { argb: "FF667085" } };
      [1, 2, 3].forEach(function (i) { row.getCell(i).border = thin(); });
    }
  }

  function thin() {
    var s2 = { style: "thin", color: { argb: "FFD6DDEA" } };
    return { top: s2, left: s2, right: s2, bottom: s2 };
  }

  function meta() {
    return {
      company: s($("client").value) || state.rev.company || "Company",
      period: s($("period").value) || state.rev.period || ""
    };
  }
  function periodClean(period) { return s(period).replace(/,/g, "").replace(/\s+/g, " ").trim(); }
  function fileNameFor(company) {
    var safe = company.replace(/[^a-z0-9]+/gi, "_").replace(/^_|_$/g, "") || "Company";
    return safe + "_Revenue_COGS_GP_Report.xlsx";
  }

  // Build the 3-sheet workbook (used by both Download and the Gmail draft)
  function buildWorkbook() {
    var m = meta(), expenses = parseFloat($("totExp").value) || 0;
    var wb = new ExcelJS.Workbook();
    wb.creator = "GATP Revenue-COGS-GP Report";
    var revTot = buildTxSheet(wb, "Revenue", state.rev);
    var cogsTot = buildTxSheet(wb, "COGS", state.cogs);
    buildPLSheet(wb, revTot, cogsTot, expenses, m.company, m.period);
    return wb;
  }

  function download() {
    var m = meta();
    buildWorkbook().xlsx.writeBuffer().then(function (buf) {
      var blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      var a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = fileNameFor(m.company);
      document.body.appendChild(a); a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
      $("dlBadge").textContent = "✓ downloaded";
    }).catch(function (e) { showErr("Export failed: " + e.message); });
  }

  /* ---------- Gmail draft (Google Identity Services + Gmail API) ---------- */
  function mailErr(m) { var b = $("mailErr"); b.textContent = "⚠ " + m; b.classList.remove("hidden"); }
  function mailClear() { $("mailErr").classList.add("hidden"); }

  function bytesToBase64(buf) {
    var bytes = new Uint8Array(buf), bin = "", CH = 0x8000;
    for (var i = 0; i < bytes.length; i += CH) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CH));
    return btoa(bin);
  }
  function toBase64Url(str) {
    var bytes = new TextEncoder().encode(str), bin = "", CH = 0x8000;
    for (var i = 0; i < bytes.length; i += CH) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CH));
    return btoa(bin).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  }
  function buildMime(o) {
    var B = "=_gatp_" + Date.now();
    var L = [];
    L.push("To: " + o.to);
    if (o.cc) L.push("Cc: " + o.cc);
    L.push("Subject: " + o.subject);
    L.push("MIME-Version: 1.0");
    L.push('Content-Type: multipart/mixed; boundary="' + B + '"');
    L.push("");
    L.push("--" + B);
    L.push('Content-Type: text/plain; charset="UTF-8"');
    L.push("Content-Transfer-Encoding: 7bit");
    L.push("");
    L.push(o.body.replace(/\n/g, "\r\n"));
    L.push("");
    L.push("--" + B);
    L.push('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; name="' + o.fileName + '"');
    L.push('Content-Disposition: attachment; filename="' + o.fileName + '"');
    L.push("Content-Transfer-Encoding: base64");
    L.push("");
    L.push(o.b64.replace(/(.{76})/g, "$1\r\n"));
    L.push("--" + B + "--");
    return L.join("\r\n");
  }

  function createDraft() {
    mailClear();
    var clientId = s($("gClientId").value);
    var to = s($("mailTo").value), cc = s($("mailCc").value);
    if (!clientId) { mailErr("Enter your Google OAuth Client ID (see the setup note in the Gmail-draft section)."); return; }
    if (!to) { mailErr("Enter a 'To' address."); return; }
    if (typeof google === "undefined" || !google.accounts || !google.accounts.oauth2) {
      mailErr("Google sign-in library hasn't loaded (needs internet). Check your connection and try again."); return;
    }
    $("mailStatus").textContent = "authorizing…";
    var tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: clientId,
      scope: "https://www.googleapis.com/auth/gmail.compose",
      callback: function (resp) {
        if (!resp || !resp.access_token) {
          $("mailStatus").textContent = "";
          mailErr("Authorization failed" + (resp && resp.error ? ": " + resp.error : ".")); return;
        }
        postDraft(resp.access_token, to, cc);
      }
    });
    try { tokenClient.requestAccessToken({ prompt: "" }); }
    catch (e) { $("mailStatus").textContent = ""; mailErr("Could not start Google authorization: " + e.message); }
  }

  function postDraft(token, to, cc) {
    var m = meta(), per = periodClean(m.period) || "the month";
    var subject = "Revenue Report - " + (periodClean(m.period) || m.company);
    var body = "Hi,\n\nPlease find the attached revenue report for " + per + ".\n\n" +
      "Let us know if you have any questions.\n\nThanks";
    $("mailStatus").textContent = "building draft…";
    buildWorkbook().xlsx.writeBuffer().then(function (buf) {
      var raw = toBase64Url(buildMime({
        to: to, cc: cc, subject: subject, body: body,
        fileName: fileNameFor(m.company), b64: bytesToBase64(buf)
      }));
      return fetch("https://gmail.googleapis.com/gmail/v1/users/me/drafts", {
        method: "POST",
        headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
        body: JSON.stringify({ message: { raw: raw } })
      });
    }).then(function (r) {
      return r.json().then(function (j) { return { ok: r.ok, j: j }; });
    }).then(function (res) {
      if (res.ok) {
        $("mailStatus").innerHTML = '✓ Draft created — <a href="https://mail.google.com/mail/u/0/#drafts" target="_blank" rel="noopener">open Gmail → Drafts</a>';
      } else {
        $("mailStatus").textContent = "";
        var msg = res.j && res.j.error && res.j.error.message ? res.j.error.message : JSON.stringify(res.j);
        mailErr("Gmail API error: " + msg);
      }
    }).catch(function (e) { $("mailStatus").textContent = ""; mailErr("Network error: " + e.message); });
  }

  /* ---------- wire up ---------- */
  wireDrop("revInput", "dropRev", "revFileName", "rev");
  wireDrop("cogsInput", "dropCogs", "cogsFileName", "cogs");
  $("processBtn").addEventListener("click", function () { clearErr(); try { render(); } catch (e) { showErr(e.message); } });
  $("totExp").addEventListener("input", function () { if (!$("resultsCard").classList.contains("hidden")) render(); });
  $("downloadBtn").addEventListener("click", download);
  $("draftBtn").addEventListener("click", createDraft);
  $("setupToggle").addEventListener("click", function (e) {
    e.preventDefault(); $("setupBox").classList.toggle("hidden");
  });
})();
