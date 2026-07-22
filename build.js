/*
 * build.js — assembles the static site into dist/ for Vercel.
 *
 * The two runtime libraries (ExcelJS, SheetJS) are NOT committed as build
 * inputs here; they are pulled from npm at their pinned versions (see
 * package.json) and copied in, reproducing the exact same files vendored in
 * ./vendor for GitHub Pages / offline use. Same bytes, same relative paths —
 * so the deployed site serves the libs from its own origin (works offline
 * once loaded, no CDN).
 */
const fs = require('fs');
const path = require('path');

const OUT = 'dist';
const HTML = ['index.html', 'monthly-services.html', 'bank-deposit.html'];
const LIBS = [
  ['node_modules/exceljs/dist/exceljs.min.js', 'vendor/exceljs.min.js'],
  ['node_modules/xlsx/dist/xlsx.full.min.js', 'vendor/xlsx.full.min.js'],
];

fs.rmSync(OUT, { recursive: true, force: true });
fs.mkdirSync(path.join(OUT, 'assets'), { recursive: true });
fs.mkdirSync(path.join(OUT, 'vendor'), { recursive: true });

HTML.forEach((f) => fs.copyFileSync(f, path.join(OUT, f)));
fs.readdirSync('assets').forEach((f) => fs.copyFileSync(path.join('assets', f), path.join(OUT, 'assets', f)));

LIBS.forEach(([src, dst]) => {
  // Prefer the npm copy (Vercel build); fall back to the committed vendor/ copy.
  const from = fs.existsSync(src) ? src : dst;
  fs.copyFileSync(from, path.join(OUT, dst));
});

// Static hosts should not run Jekyll on this output.
fs.writeFileSync(path.join(OUT, '.nojekyll'), '');

console.log('Built ' + OUT + '/ — ' + (HTML.length) + ' pages, assets, and 2 vendored libs.');
