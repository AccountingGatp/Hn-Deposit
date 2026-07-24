/*
 * rd-app.js — UI controller for the Remote Deposit Check Extractor.
 *
 * Pipeline (all client-side):
 *   PDF file  -> pdf.js renders each page to a canvas
 *             -> Tesseract.js OCRs the canvas (text + word boxes)
 *             -> RDCore.extractCheck turns a page into a candidate record
 *             -> RDCore.dedupeRecords collapses voucher+check duplicates
 *             -> editable review table with a thumbnail per row
 *             -> RDXlsx builds the 3-column workbook for download.
 *
 * pdf.js ships as an ES module, so this file is loaded as type="module".
 * Tesseract, ExcelJS and RDCore/RDXlsx are classic globals.
 */
import * as pdfjsLib from '../vendor/pdf.min.mjs';
pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.mjs';

const $ = (id) => document.getElementById(id);
const RENDER_SCALE = 3.1;   // ~230 dpi equivalent — enough for reliable digit OCR
const THUMB_W = 520;

let files = [];       // {file, name}
let records = [];     // final review records
let ocrWorker = null;

/* ------------------------------- file input ------------------------------ */

const dropPdf = $('dropPdf');
const pdfInput = $('pdfInput');

['dragenter', 'dragover'].forEach((ev) =>
  dropPdf.addEventListener(ev, (e) => { e.preventDefault(); dropPdf.classList.add('dragover'); }));
['dragleave', 'drop'].forEach((ev) =>
  dropPdf.addEventListener(ev, (e) => { e.preventDefault(); dropPdf.classList.remove('dragover'); }));
dropPdf.addEventListener('drop', (e) => {
  const dropped = [...(e.dataTransfer?.files || [])].filter((f) => /\.pdf$/i.test(f.name) || f.type === 'application/pdf');
  addFiles(dropped);
});
pdfInput.addEventListener('change', () => addFiles([...pdfInput.files]));

function addFiles(list) {
  list.forEach((f) => {
    if (!files.some((x) => x.name === f.name && x.file.size === f.size)) files.push({ file: f, name: f.name });
  });
  renderFileList();
}

function renderFileList() {
  const box = $('fileList');
  if (!files.length) { box.classList.add('hidden'); box.innerHTML = ''; }
  else {
    box.classList.remove('hidden');
    box.innerHTML = files.map((f, i) =>
      `<div class="fileitem"><span>🧾 ${escapeHtml(f.name)}</span>` +
      `<button data-i="${i}" class="rmfile" title="Remove">✕</button></div>`).join('');
    box.querySelectorAll('.rmfile').forEach((b) =>
      b.addEventListener('click', () => { files.splice(+b.dataset.i, 1); renderFileList(); }));
  }
  $('pdfFileName').textContent = files.length ? `${files.length} file${files.length > 1 ? 's' : ''} ready` : '';
  $('processBtn').disabled = files.length === 0;
  $('readyHint').textContent = files.length ? 'Ready to extract.' : 'Add at least one PDF to enable.';
}

/* ------------------------------- extraction ------------------------------ */

$('processBtn').addEventListener('click', run);

async function run() {
  hideError();
  $('processBtn').disabled = true;
  $('resultsCard').classList.add('hidden');
  showProgress(0, 'Starting…');

  try {
    // Count pages up front for an accurate progress bar.
    const docs = [];
    let totalPages = 0;
    for (const f of files) {
      const buf = await f.file.arrayBuffer();
      const doc = await pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise;
      docs.push({ name: f.name, doc });
      totalPages += doc.numPages;
    }

    await ensureWorker();

    const raw = [];
    let done = 0;
    for (const { name, doc } of docs) {
      const perFile = [];
      for (let p = 1; p <= doc.numPages; p++) {
        showProgress(done / totalPages, `OCR — ${name} · page ${p}/${doc.numPages}`);
        const { canvas } = await renderPage(doc, p);
        const { data } = await ocrWorker.recognize(canvas, {}, { text: true, blocks: true });
        const page = {
          text: data.text || '',
          words: flattenWords(data),
          width: canvas.width,
          height: canvas.height
        };
        const rec = RDCore.extractCheck(page, {
          file: name,
          pageIndex: p - 1,
          pageLabel: `${name} · p${p}`,
          thumb: makeThumb(canvas)
        });
        if (rec) perFile.push(rec);
        done++;
      }
      // One record per check face, enriched from its voucher page.
      const asm = RDCore.assembleRecords(perFile);
      // Single-check file: the deposit total in the file name is a free amount
      // cross-check. If it disagrees, mark the row so OCR slips don't slip by.
      const hint = amountFromName(name);
      if (asm.length === 1 && hint != null && asm[0].amount != null &&
          Math.abs(hint - asm[0].amount) >= 0.005) {
        asm[0].amountMismatch = true;
      }
      asm.forEach((r) => { r.file = name; });
      raw.push(...asm);
    }

    records = raw.map(finalize);
    showProgress(1, 'Done');
    docs.forEach((d) => d.doc.destroy());
    renderResults();
  } catch (err) {
    console.error(err);
    showError('Extraction failed: ' + (err && err.message ? err.message : err) +
      '. Tip: this page must be served over http (GitHub Pages, or `python3 -m http.server`) — opening the file directly blocks the OCR worker.');
  } finally {
    $('processBtn').disabled = false;
    setTimeout(() => $('progWrap').classList.add('hidden'), 600);
  }
}

async function ensureWorker() {
  if (ocrWorker) return;
  ocrWorker = await Tesseract.createWorker('eng', 1, {
    workerPath: 'vendor/tesseract-worker.min.js',
    corePath: 'vendor/tess-core/',
    langPath: 'vendor/tessdata/',
    gzip: true,
    logger: (m) => {
      if (m.status === 'recognizing text' && typeof m.progress === 'number') {
        $('progText').textContent = $('progText').textContent.replace(/\s+\(\d+%\)$/, '') +
          ` (${Math.round(m.progress * 100)}%)`;
      }
    }
  });
  // Whitelist-free but tuned for single blocks of print.
  await ocrWorker.setParameters({ preserve_interword_spaces: '1' });
}

async function renderPage(doc, pageNum) {
  const page = await doc.getPage(pageNum);
  const viewport = page.getViewport({ scale: RENDER_SCALE });
  const canvas = document.createElement('canvas');
  canvas.width = Math.floor(viewport.width);
  canvas.height = Math.floor(viewport.height);
  const ctx = canvas.getContext('2d', { willReadFrequently: true });
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvasContext: ctx, viewport }).promise;
  page.cleanup();
  preprocess(ctx, canvas.width, canvas.height);
  return { canvas };
}

// Grayscale + contrast stretch. Scanned checks sit on light colored "VOID"
// security tints; pushing lights toward white and darks toward black lets
// Tesseract lock onto the black print instead of the pattern.
function preprocess(ctx, w, h) {
  const img = ctx.getImageData(0, 0, w, h);
  const d = img.data;
  for (let i = 0; i < d.length; i += 4) {
    const g = d[i] * 0.299 + d[i + 1] * 0.587 + d[i + 2] * 0.114;
    // Piecewise contrast: brighten background (>165), keep/darken ink.
    let v;
    if (g > 165) v = 255;
    else if (g < 90) v = Math.max(0, g * 0.65);
    else v = (g - 90) * (255 / 75);           // ramp 90..165 -> 0..255
    d[i] = d[i + 1] = d[i + 2] = v;
  }
  ctx.putImageData(img, 0, 0);
}

// Tesseract v5 returns nested blocks->paragraphs->lines->words. Flatten to the
// {text, conf, bbox} shape RDCore expects for positional check-number hunting.
function flattenWords(data) {
  const out = [];
  const walk = (node) => {
    if (!node) return;
    if (Array.isArray(node.words)) {
      node.words.forEach((w) => out.push({ text: w.text, conf: w.confidence, bbox: w.bbox }));
    }
    ['blocks', 'paragraphs', 'lines'].forEach((k) => {
      if (Array.isArray(node[k])) node[k].forEach(walk);
    });
  };
  if (Array.isArray(data.blocks)) data.blocks.forEach(walk);
  else if (Array.isArray(data.words)) data.words.forEach((w) => out.push({ text: w.text, conf: w.confidence, bbox: w.bbox }));
  return out;
}

function makeThumb(canvas) {
  const scale = Math.min(1, THUMB_W / canvas.width);
  const tc = document.createElement('canvas');
  tc.width = Math.round(canvas.width * scale);
  tc.height = Math.round(canvas.height * scale);
  tc.getContext('2d').drawImage(canvas, 0, 0, tc.width, tc.height);
  return tc.toDataURL('image/jpeg', 0.7);
}

function finalize(r) {
  const f = RDCore.reviewFlag(r);
  return Object.assign({}, r, { flagStatus: f.status, flagLabel: f.label });
}

/* --------------------------------- results ------------------------------- */

function renderResults() {
  $('resultsCard').classList.remove('hidden');

  const total = records.reduce((s, r) => s + (r.amount || 0), 0);
  const flagged = records.filter((r) => r.flagStatus !== 'OK').length;

  $('tiles').innerHTML = [
    tile('Checks found', records.length),
    tile('Total amount', money(total)),
    tile('Ready (OK)', records.length - flagged, 'good'),
    tile('Need review', flagged, flagged ? 'warnv' : 'good')
  ].join('');

  renderReconByFile();
  renderTable();
  $('dlBadge').textContent = '';
}

function renderReconByFile() {
  // Group extracted totals by file so the user can sanity-check each deposit.
  const byFile = {};
  records.forEach((r) => {
    (byFile[r.file] = byFile[r.file] || { count: 0, sum: 0 });
    byFile[r.file].count++; byFile[r.file].sum += (r.amount || 0);
  });
  const shortfalls = [];
  const rows = Object.keys(byFile).map((f) => {
    const g = byFile[f];
    const hint = amountFromName(f);
    const match = hint != null ? (Math.abs(hint - g.sum) < 0.005 ? 'ok' : 'bad') : null;
    if (match === 'bad') shortfalls.push({ f, sum: g.sum, hint });
    return `<tr><td>${escapeHtml(f)}</td><td class="num">${g.count}</td>` +
      `<td class="num">${money(g.sum)}</td>` +
      `<td>${hint != null
        ? `<span class="badge ${match === 'ok' ? 'ok' : 'bad'}">${match === 'ok' ? '✓ matches file total' : '≠ file total ' + money(hint)}</span>`
        : '<span class="muted">—</span>'}</td></tr>`;
  }).join('');

  const banner = $('reviewBanner');
  if (shortfalls.length) {
    banner.classList.remove('hidden');
    banner.innerHTML = '<b>⚠ Some files don\'t reconcile to the total in their file name.</b> ' +
      'OCR may have missed a check or misread an amount — check the image on each row, and ' +
      'use “Add a check row” for anything missing:' +
      '<ul>' + shortfalls.map((s) =>
        `<li>${escapeHtml(s.f)} — extracted ${money(s.sum)} vs ${money(s.hint)} in the file name</li>`).join('') +
      '</ul>';
  } else {
    banner.classList.add('hidden'); banner.innerHTML = '';
  }
  $('reconBox').innerHTML =
    '<h3 class="section-title">Per-file check</h3>' +
    '<div class="tablewrap"><table class="data"><thead><tr>' +
    '<th>Source file</th><th>Checks</th><th>Extracted total</th><th>vs. amount in file name</th>' +
    '</tr></thead><tbody>' + rows + '</tbody></table></div>';
}

function renderTable() {
  const t = $('checksTable');
  $('rowCount').textContent = `${records.length} row${records.length === 1 ? '' : 's'}`;
  t.innerHTML =
    '<thead><tr><th>Check</th><th>Date</th><th>Description (check #)</th>' +
    '<th>Amount</th><th>Review flag</th><th></th></tr></thead><tbody>' +
    records.map((r, i) => rowHtml(r, i)).join('') + '</tbody>';

  t.querySelectorAll('[data-edit]').forEach((el) => {
    el.addEventListener('blur', onEdit);
    el.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); el.blur(); } });
  });
  t.querySelectorAll('.rmrow').forEach((b) =>
    b.addEventListener('click', () => { records.splice(+b.dataset.i, 1); renderResults(); }));
  t.querySelectorAll('.thumb img').forEach((img) =>
    img.addEventListener('click', () => openLightbox(img.src)));
}

function rowHtml(r, i) {
  const needs = r.flagStatus !== 'OK';
  return `<tr class="${needs ? 'needs' : ''}">` +
    `<td class="thumb">${r.thumb ? `<img src="${r.thumb}" alt="check ${i + 1}" />` : '<span class="muted">—</span>'}` +
    `<div class="src muted">${escapeHtml(r.file)}${r.pageIndex != null ? ' · p' + (r.pageIndex + 1) : ''}</div></td>` +
    `<td><span data-edit="date" data-i="${i}" contenteditable="true" class="cell">${escapeHtml(r.date || '')}</span></td>` +
    `<td><span data-edit="checkNo" data-i="${i}" contenteditable="true" class="cell mono">${escapeHtml(r.checkNo || '')}</span></td>` +
    `<td class="num"><span data-edit="amount" data-i="${i}" contenteditable="true" class="cell">${r.amount == null ? '' : Number(r.amount).toFixed(2)}</span></td>` +
    `<td><span class="badge ${needs ? 'bad' : 'ok'}">${escapeHtml(r.flagLabel)}</span></td>` +
    `<td><button class="rmrow" data-i="${i}" title="Delete row">✕</button></td></tr>`;
}

function onEdit(e) {
  const el = e.target;
  const i = +el.dataset.i, field = el.dataset.edit;
  const val = el.textContent.trim();
  const r = records[i];
  if (field === 'date') {
    r.date = val;
    const d = RDCore._internal.matchAnyDate(val);
    r.dateObj = d ? { mo: d.mo, day: d.day, yr: d.yr } : null;
  } else if (field === 'checkNo') {
    r.checkNo = val;
    r.checkNoConfident = true;   // a human typed it → trust it
    r.checkNoSource = 'manual';
  } else if (field === 'amount') {
    const n = parseFloat(val.replace(/[^0-9.\-]/g, ''));
    r.amount = isNaN(n) ? null : n;
  }
  records[i] = finalize(r);
  renderResults();
}

$('addRowBtn').addEventListener('click', () => {
  records.push(finalize({ file: '(manual entry)', pageIndex: null, thumb: null, date: '', dateObj: null, checkNo: '', amount: null }));
  renderResults();
});

/* --------------------------------- download ------------------------------ */

$('downloadBtn').addEventListener('click', () => {
  if (!records.length) return;
  const wb = RDXlsx.buildWorkbook(ExcelJS, records, {});
  wb.xlsx.writeBuffer().then((buf) => {
    downloadBlob(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
      outName('xlsx'));
    $('dlBadge').innerHTML = '<span class="badge ok">✓ workbook downloaded</span>';
  });
});

$('downloadCsvBtn').addEventListener('click', () => {
  if (!records.length) return;
  const head = 'Date,Description,Amount\n';
  const body = records.map((r) =>
    [csv(r.date || ''), csv(r.checkNo || ''), r.amount == null ? '' : Number(r.amount).toFixed(2)].join(',')
  ).join('\n');
  downloadBlob(new Blob([head + body], { type: 'text/csv' }), outName('csv'));
  $('dlBadge').innerHTML = '<span class="badge ok">✓ CSV downloaded</span>';
});

function outName(ext) {
  return `Remote_Deposit_${records.length}_checks.${ext}`;
}

/* --------------------------------- helpers ------------------------------- */

function tile(k, v, cls) {
  return `<div class="tile ${cls || ''}"><div class="k">${k}</div><div class="v">${v}</div></div>`;
}
function money(n) { return '$' + Number(n || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function amountFromName(name) {
  // Deposit total is often encoded in the file name ("… $7807.70.pdf").
  const m = String(name).match(/\$?\s*([0-9]{1,3}(?:,?[0-9]{3})*(?:\.[0-9]{2}))\s*\.pdf$/i) ||
            String(name).match(/([0-9]+\.[0-9]{2})/);
  return m ? parseFloat(m[1].replace(/,/g, '')) : null;
}
function showProgress(frac, text) {
  $('progWrap').classList.remove('hidden');
  $('progFill').style.width = Math.round(frac * 100) + '%';
  $('progText').textContent = text;
}
function showError(msg) { const b = $('errBox'); b.textContent = msg; b.classList.remove('hidden'); }
function hideError() { $('errBox').classList.add('hidden'); }
function downloadBlob(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}
function csv(s) { s = String(s); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }
function escapeHtml(s) { return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

function openLightbox(src) {
  let lb = $('lightbox');
  if (!lb) {
    lb = document.createElement('div'); lb.id = 'lightbox';
    lb.addEventListener('click', () => lb.classList.add('hidden'));
    document.body.appendChild(lb);
  }
  lb.innerHTML = `<img src="${src}" alt="check" />`;
  lb.classList.remove('hidden');
}
