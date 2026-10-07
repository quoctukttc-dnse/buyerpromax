/* Kiểm thử bộ đọc GEN trong Node (không cần trình duyệt):
   node tgen.mjs <thư mục chủ hàng> [--dump]
   - đọc mọi xlsx/pdf, in các dòng hàng đọc được
   - nếu có inbound: chạy analyzeFab và so với cột Invoice Quantity người đã điền (nếu có) */
import fs from 'fs'; import path from 'path';
import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const ExcelJS = require('exceljs');
const pdfjs = require('/home/claude/build/node_modules/pdfjs-dist/legacy/build/pdf.js');

const src = fs.readFileSync('/home/claude/build/fab.js', 'utf8') + '\n' + fs.readFileSync('/home/claude/build/gen.js', 'utf8')
  + '\nreturn { readGenWb, readGenPdf, analyzeFab, poSap, AZ, unitKey, pdfItemXY };';
const lib = new Function('ExcelJS', 'module', src)(ExcelJS, { exports: {} });

const dir = process.argv[2];
const dump = process.argv.includes('--dump');
const files = []; (function walk(d) { for (const f of fs.readdirSync(d, { withFileTypes: true })) { const p = path.join(d, f.name); if (f.isDirectory()) walk(p); else if (!f.name.startsWith('.')) files.push(p); } })(dir);

function headerIndex(ws) {
  const H = {};
  const row = ws.getRow(1);
  for (let i = 1; i <= ws.columnCount; i++) { const v = String(row.getCell(i).text || '').trim().toLowerCase(); if (v) H[v] = i; }
  const pick = (...names) => { for (const n of names) if (H[n]) return H[n]; return 0; };
  return {
    po: pick('purchasing document'), material: pick('material'), desc: pick('material description'),
    size: pick('size'), spec: pick('specification'), price: pick('gross price'), sur: pick('surcharge item'),
    qty: pick('quantity'), deliv: pick('delivered qty'), invQty: pick('invoice quantity', 'input quantity'),
    invNo: pick('invoice number'), invDate: pick('invoice date'),
    supRef: pick('supplier ref', 'supplier mat. no.', 'supplier material number', 'supplier material no.'), color: pick('color', 'colour'), lapdip: pick('lapdip color'),
    overTol: pick('over tolerance qty'), unit: pick('base unit of measure'), cur: pick('currency'),
  };
}
const num = (v) => { if (v == null || v === '') return NaN; if (typeof v === 'number') return v; if (typeof v === 'object' && v.result != null) return Number(v.result); const n = parseFloat(String(v).replace(/,/g, '')); return isNaN(n) ? NaN : n; };
function readInbRows(ws, H) {
  const rows = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r); const poV = String(row.getCell(H.po).text || '').trim(); if (!poV) continue;
    const g = (c) => (c ? String(row.getCell(c).text || '') : '');
    rows.push({ r, poV, desc: g(H.desc), spec: g(H.spec), material: g(H.material), supRef: g(H.supRef), color: g(H.color), lapdip: g(H.lapdip),
      unit: g(H.unit).trim(), cur: g(H.cur).trim(), overTol: H.overTol ? num(row.getCell(H.overTol).value) : NaN,
      size: g(H.size).replace(/\s+/g, '').toUpperCase(), qty: num(row.getCell(H.qty).value), deliv: H.deliv ? num(row.getCell(H.deliv).value) : 0,
      invQty: H.invQty ? num(row.getCell(H.invQty).value) : NaN, price: H.price ? num(row.getCell(H.price).value) : NaN, sur: H.sur ? num(row.getCell(H.sur).value) : 0,
      invNoHad: g(H.invNo), eff: 0, matched: null, setQty: null });
  }
  return rows;
}

async function pdfLinesXY(file) {
  const data = new Uint8Array(fs.readFileSync(file));
  const pdf = await pdfjs.getDocument({ data, isEvalSupported: false, useSystemFonts: true, verbosity: 0 }).promise;
  const pages = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const tc = await page.getTextContent();
    pages.push(tc.items.filter((it) => it.str && it.str.trim()).map((it) => lib.pdfItemXY(it)));
  }
  return pages;
}

const docs = [], inbs = [];
for (const f of files) {
  const ext = path.extname(f).toLowerCase();
  if (ext === '.xlsx' || ext === '.xlsm' || ext === '.xls') {
    let buf = fs.readFileSync(f);
    if (ext === '.xls') { /* Excel 97-2003 → xlsx bằng SheetJS (npm i xlsx@0.18.5) */
      const X = require('xlsx'); const w0 = X.read(buf, { type: 'buffer', cellDates: true }); buf = X.write(w0, { type: 'buffer', bookType: 'xlsx', cellDates: true });
    }
    const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf);
    const head = (wb.worksheets[0].getRow(1).values || []).join(' ').toLowerCase();
    if (/purchasing document/.test(head)) { inbs.push({ f, wb }); continue; }
    const d = lib.readGenWb(wb, path.basename(f), path.dirname(f));
    docs.push({ f, d });
  } else if (ext === '.pdf') {
    const pages = await pdfLinesXY(f);
    const nChars = pages.reduce((a, p) => a + p.reduce((b, i) => b + i.s.length, 0), 0);
    if (nChars < 40) { docs.push({ f, d: null, scan: true }); continue; }
    const d = lib.readGenPdf(pages, path.basename(f), path.dirname(f));
    docs.push({ f, d });
  }
}
for (const x of docs) {
  console.log('\n=== ' + path.basename(x.f) + (x.scan ? '  (SCAN — không có chữ)' : ''));
  if (!x.d) { console.log('   (không đọc được bảng)'); continue; }
  const d = x.d;
  console.log(`   role=${d.role} invNo=${d.inv.invNo || '(trống)'} date=${d.inv.invDate || '(trống)'} items=${d.inv.items.length} unit=${d.inv.unitDefault} total=${d.inv.total} totalQty=${d.inv.totalQty} pklGroups=${d.pkl ? d.pkl.groups.length : 0}`);
  if (dump || true) for (const it of d.inv.items.slice(0, 60)) console.log(`   • ${it.po || it.poRaw || '(no PO)'} | ${it.material || ''} | ${(it.article || '').slice(0, 35)} | ${(it.colorText || '').slice(0, 30)} | sz=${it.size || ''} | ${it.qty} ${it.unit} | p=${it.price} a=${it.amount}`);
  if (d.pkl && dump) for (const g of d.pkl.groups.slice(0, 40)) console.log(`   ▫ PKL ${g.po || g.poRaw} | ${g.article} | ${g.color} | lot=${g.lot} | sz=${g.size} | ${g.total} (${g.rolls.length} cuộn)`);
}
/* đối chiếu với inbound */
let invDocs = docs.filter((x) => x.d && x.d.role === 'inv');
if (!invDocs.length) invDocs = docs.filter((x) => x.d && x.d.pklOnly);
for (const x of invDocs) {
  const d = x.d;
  /* gộp packing list cùng thư mục */
  const pk = docs.filter((y) => y.d && y !== x && y.d.pkl).map((y) => y.d.pkl.groups).flat();
  const pkl = pk.length ? { groups: pk, level: 'lot' } : d.pkl;
  for (const ib of inbs) {
    const ws = ib.wb.worksheets[0]; const H = headerIndex(ws);
    const rows = readInbRows(ws, H);
    const { lines, VAL } = lib.analyzeFab(d.inv, pkl, rows, {});
    console.log(`\n### ${path.basename(x.f)} ⟷ ${path.basename(ib.f)}  (${rows.length} dòng inbound)`);
    const st = {}; lines.forEach((l) => { st[l.status] = (st[l.status] || 0) + 1; });
    console.log('   ' + JSON.stringify(st) + ` docQty=${VAL.docQty} wrote=${VAL.wroteQty} qtyBad=${VAL.qtyBad} totalBad=${VAL.totalBad} inbTotal=${VAL.inbTotal} invTotal=${VAL.invTotal}`);
    for (const l of lines) console.log(`   ${l.status.padEnd(22)} ${(l.it.code || '').slice(0, 40).padEnd(40)} ${String(l.sapPo).padEnd(12)} ${String(l.it.qty).padStart(9)} ${l.unitUsed || ''} pkl=${l.pklTotal}  ${l.matchBy}  ${l.note.slice(0, 110)}`);
    /* so với số người đã điền */
    let same = 0, diff = 0, extra = 0, miss = 0;
    for (const r of rows) {
      const had = r.invNoHad ? r.invQty : NaN;
      const got = r.setQty;
      if (!isNaN(had) && got != null) { if (Math.abs(had - got) < 0.01) same++; else { diff++; console.log(`   ✗ ${r.poV} ${r.material} ${r.color} sz=${r.size}: người điền ${had}, công cụ ${got}`); } }
      else if (!isNaN(had) && got == null) { miss++; console.log(`   ○ ${r.poV} ${r.material} ${r.color} sz=${r.size}: người điền ${had}, công cụ KHÔNG điền`); }
      else if (isNaN(had) && got != null && r.invNoHad === '' && rows.some((z) => z.invNoHad)) { extra++; console.log(`   + ${r.poV} ${r.material} ${r.color} sz=${r.size}: công cụ điền ${got}, người không điền`); }
    }
    if (rows.some((z) => z.invNoHad)) console.log(`   SO VỚI NGƯỜI ĐIỀN: giống ${same} · khác ${diff} · thiếu ${miss} · thừa ${extra}`);
  }
}
