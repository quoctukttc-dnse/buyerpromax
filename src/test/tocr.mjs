/* OCR thử bằng tesseract CLI (tsv) → words → readGenPdf → analyzeFab với inbound cùng thư mục
   node tocr.mjs <thư mục> */
import fs from 'fs'; import path from 'path'; import { execSync } from 'child_process';
import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const ExcelJS = require('exceljs');
const src = fs.readFileSync('fab.js','utf8') + '\n' + fs.readFileSync('gen.js','utf8') + '\nreturn { readGenPdf, analyzeFab, ocrWordsToItems, ocrFixToken, genLinesOf, pdfHeader, qtyOnLine, genPos };';
const lib = new Function('ExcelJS','module',src)(ExcelJS,{exports:{}});
const dir = process.argv[2];
const files = fs.readdirSync(dir).map((f) => path.join(dir, f));
const tmp = '/tmp/ocrtest'; fs.rmSync(tmp, { recursive: true, force: true }); fs.mkdirSync(tmp);
function ocrPdf(f) {
  const base = path.join(tmp, 'p');
  execSync(`pdftoppm -r 300 -png "${f}" "${base}"`);
  const pngs = fs.readdirSync(tmp).filter((x) => x.startsWith('p') && x.endsWith('.png')).sort();
  const pages = [];
  for (const png of pngs) {
    const tsv = execSync(`tesseract "${path.join(tmp, png)}" - -l eng+vie --psm 6 tsv 2>/dev/null`).toString();
    const dims = execSync(`python3 -c "from PIL import Image;im=Image.open('${path.join(tmp, png)}');print(im.size[0],im.size[1])"`).toString().trim().split(' ').map(Number);
    const words = [];
    for (const row of tsv.split('\n').slice(1)) {
      const c = row.split('\t'); if (c.length < 12 || !c[11] || !c[11].trim()) continue;
      words.push({ line: c[1] + '-' + c[2] + '-' + c[3] + '-' + c[4], x0: +c[6], y0: +c[7], x1: +c[6] + +c[8], y1: +c[7] + +c[9], text: c[11], conf: +c[10] });
    }
    const scale = 300 / 72;
    pages.push(lib.ocrWordsToItems(words, scale, dims[1]));
    fs.unlinkSync(path.join(tmp, png));
  }
  return pages;
}
const chars = (p) => p.reduce((a, pg) => a + pg.reduce((b, i) => b + i.s.length, 0), 0);
for (const f of files) {
  if (!/\.pdf$/i.test(f)) continue;
  const txt = execSync(`pdftotext "${f}" - 2>/dev/null | wc -c`).toString().trim();
  if (+txt > 40) { console.log('--', path.basename(f), 'có chữ, bỏ qua'); continue; }
  const pages = ocrPdf(f);
  if (process.argv.includes('-v')) for (const pg of pages) { const L = lib.genLinesOf(pg); const hd = lib.pdfHeader(L, 0); console.log('  header', JSON.stringify(hd)); L.forEach((l, i) => { const q = lib.qtyOnLine(l, hd ? hd.qcol : null, false); console.log('   ', String(i).padStart(3), (q ? `Q=${q.q}${q.unit}` : '').padEnd(12), lib.genPos(l.text).join(','), '|', l.cells.map((c) => Math.round(c.x) + ':' + c.s).join(' ¦ ').slice(0, 160)); }); }
  const d = lib.readGenPdf(pages, path.basename(f), dir);
  console.log('=== OCR', path.basename(f), 'chars', chars(pages));
  if (!d || !d.inv) { console.log('  (không đọc được)'); continue; }
  console.log(`   role=${d.role} invNo=${d.inv.invNo} date=${d.inv.invDate} items=${d.inv.items.length} unit=${d.inv.unitDefault}`);
  for (const it of d.inv.items) console.log(`   • ${it.po || it.poRaw || '(no PO)'} | ${it.material} | ${(it.colorShort || '').slice(0, 50)} | sz=${it.size} | ${it.qty} ${it.unit} | p=${it.price} a=${it.amount}`);
  const inbs = files.filter((x) => /\.xlsx$/i.test(x));
  for (const ib of inbs) {
    const wb = new ExcelJS.Workbook(); await wb.xlsx.load(fs.readFileSync(ib));
    const ws = wb.worksheets[0]; const head = (ws.getRow(1).values || []).join(' ').toLowerCase();
    if (!/purchasing document/.test(head)) continue;
    const H = {}; ws.getRow(1).eachCell((c, i) => { H[String(c.text).trim().toLowerCase()] = i; });
    const pick = (...n) => { for (const k of n) if (H[k]) return H[k]; return 0; };
    const num = (v) => { if (v == null || v === '') return NaN; if (typeof v === 'number') return v; if (typeof v === 'object' && v.result != null) return Number(v.result); const n = parseFloat(String(v).replace(/,/g, '')); return isNaN(n) ? NaN : n; };
    const rows = [];
    for (let r = 2; r <= ws.rowCount; r++) { const row = ws.getRow(r); const poV = String(row.getCell(pick('purchasing document')).text || '').trim(); if (!poV) continue; const g = (k) => { const c = pick(k); return c ? String(row.getCell(c).text || '') : ''; };
      rows.push({ r, poV, desc: g('material description'), spec: g('specification'), material: g('material'), supRef: g('supplier ref') || g('supplier mat. no.'), color: g('color'), lapdip: g('lapdip color'), unit: g('base unit of measure').trim(), overTol: num(row.getCell(pick('over tolerance qty')).value), size: g('size').replace(/\s+/g, '').toUpperCase(), qty: num(row.getCell(pick('quantity')).value), deliv: num(row.getCell(pick('delivered qty')).value) || 0, invQty: num(row.getCell(pick('invoice quantity')).value), price: num(row.getCell(pick('gross price')).value), sur: num(row.getCell(pick('surcharge item')).value) || 0, invNoHad: g('invoice number'), eff: 0, matched: null, setQty: null }); }
    const { lines, VAL } = lib.analyzeFab(d.inv, d.pkl, rows, {});
    console.log(`   ### ⟷ ${path.basename(ib)}: ` + JSON.stringify(lines.reduce((a, l) => { a[l.status] = (a[l.status] || 0) + 1; return a; }, {})) + ` wrote=${VAL.wroteQty}`);
    for (const l of lines) console.log(`      ${l.status.padEnd(20)} ${(l.it.code || '').slice(0, 36).padEnd(36)} ${String(l.sapPo).padEnd(11)} ${String(l.it.qty).padStart(8)} ${l.unitUsed || ''}  ${l.note.slice(0, 90)}`);
    for (const r of rows) { const had = r.invNoHad ? r.invQty : NaN; if (!isNaN(had) && r.setQty != null && Math.abs(had - r.setQty) > 0.01) console.log(`      ✗ ${r.poV} ${r.material} sz=${r.size}: người ${had} / tool ${r.setQty}`); else if (!isNaN(had) && r.setQty == null) console.log(`      ○ ${r.poV} ${r.material} sz=${r.size}: người ${had} / tool KHÔNG`); }
  }
}
