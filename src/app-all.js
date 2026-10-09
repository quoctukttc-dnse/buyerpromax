/* ================= Helpers ================= */
const $ = (s) => document.querySelector(s);
const norm = (s) => String(s == null ? '' : s).toUpperCase().replace(/\s+/g, '');
const SIZE_ORDER = ['XS', 'S', 'M', 'L', 'XL', 'XXL', 'XXXL'];
/* "KHỚP", "KHỚP (giao thiếu)", "KHỚP (trong dung sai)" đều là kết quả đạt */
const isOkK = (x) => String(x == null ? '' : x).indexOf('KHỚP') === 0;
const isOk = (x) => isOkK(x) || x === 'CHƯA CÓ INBOUND';

function num(v) {
  if (v == null) return NaN;
  if (typeof v === 'number') return v;
  let t = String(v).trim().replace(/\s/g, '');
  if (!t) return NaN;
  if (/^\d{1,3}([.,]\d{3})+$/.test(t)) return parseInt(t.replace(/[.,]/g, ''), 10);
  if (/^\d+$/.test(t)) return parseInt(t, 10);
  t = t.replace(/\.(?=\d{3}\b)/g, '').replace(',', '.');
  const n = parseFloat(t);
  return isNaN(n) ? NaN : n;
}
function colLetter(i) { // 1 -> A
  let s = '';
  while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
  return s;
}
const fmt = (n) => (n == null || isNaN(n) ? '—' : Number(n).toLocaleString('vi-VN'));
function log(msg, cls) {
  const el = document.createElement('div');
  el.className = 'logline' + (cls ? ' ' + cls : '');
  el.textContent = msg;
  $('#log').appendChild(el);
  $('#log').scrollTop = $('#log').scrollHeight;
}

/* ================= PDF -> visual lines ================= */
const PDF_PAGES = new Map();
/* ---------- Word 97-2003 (.doc): lấy chữ ngay trong trình duyệt (CFB của SheetJS + bảng piece của Word) ----------
   Chứng từ kiểu Chain Guan là chữ dàn cột cố định (font monospace) → dựng thành "trang PDF" giả: mỗi cụm chữ
   cách nhau ≥ 2 khoảng trắng là một ô, toạ độ x theo số cột ký tự → đưa vào đúng bộ đọc PDF chung. */
async function docText(buf) {
  if (typeof window.loadSheetJS !== 'function') throw new Error('thiếu thư viện SheetJS');
  const X = await window.loadSheetJS();
  if (!X.CFB) throw new Error('thư viện không có bộ đọc CFB');
  const cfb = X.CFB.read(new Uint8Array(buf), { type: 'array' });
  const get = (n) => { const e = X.CFB.find(cfb, n); return e && e.content ? new Uint8Array(e.content) : null; };
  const wd = get('WordDocument');
  if (!wd) throw new Error('không phải file Word 97-2003');
  const dv = new DataView(wd.buffer, wd.byteOffset, wd.byteLength);
  if (dv.getUint16(0, true) !== 0xA5EC) throw new Error('không phải file Word 97-2003');
  const flags = dv.getUint16(0x0A, true);
  if (flags & 0x0100) throw new Error('file Word có mật khẩu');
  const tbl = get((flags & 0x0200) ? '1Table' : '0Table');
  if (!tbl) throw new Error('thiếu bảng 0Table/1Table');
  const fcClx = dv.getUint32(0x01A2, true), lcbClx = dv.getUint32(0x01A6, true);
  const tv = new DataView(tbl.buffer, tbl.byteOffset, tbl.byteLength);
  let i = fcClx;
  while (i < fcClx + lcbClx && tbl[i] === 0x01) i += 3 + tv.getUint16(i + 1, true);   // bỏ Prc
  if (tbl[i] !== 0x02) throw new Error('không đọc được bảng piece');
  const lcb = tv.getUint32(i + 1, true), base = i + 5;
  const n = (lcb - 4) / 12;
  const cp1252 = new TextDecoder('windows-1252');
  let out = '';
  for (let k = 0; k < n; k++) {
    const cp0 = tv.getUint32(base + k * 4, true), cp1 = tv.getUint32(base + (k + 1) * 4, true);
    const pcd = base + (n + 1) * 4 + k * 8;
    const fc = tv.getUint32(pcd + 2, true);
    const len = cp1 - cp0;
    if (fc & 0x40000000) { const o = (fc & 0x3FFFFFFF) / 2; out += cp1252.decode(wd.subarray(o, o + len)); }
    else { let t = ''; for (let j = 0; j < len; j++) t += String.fromCharCode(dv.getUint16(fc + j * 2, true)); out += t; }
  }
  /* trường (field): giữ phần kết quả, bỏ mã lệnh; ô bảng / ngắt dòng → khoảng trắng / xuống dòng */
  out = out.replace(/\x13[^\x14\x15]*\x14([^\x15]*)\x15/g, '$1').replace(/\x13[^\x15]*\x15/g, '')
    .replace(/\x07/g, '   ').replace(/\x0b/g, '\r');
  return out;
}
/* Packing list chữ dàn cột (Word, kiểu Chain Guan): P/O NO:… → "#634 RIO" (màu) → "LOT NO.1-3" → từng cây
   "PR18 150YDS 19.70KGS 19.90KGS"; dòng cộng nhóm dưới đường kẻ đã được đánh dấu TOTAL. Trả về dạng gen.pkl. */
function readMonoPkl(lines, fname) {
  const t = lines.join(' ');
  if (!/PACKING\s*LIST/i.test(t) || /COMMERCIAL\s*INVOICE/i.test(t)) return null;
  let po = '', color = '', lot = '';
  const groups = [];
  for (const raw of lines) {
    const L = raw.trim();
    let m;
    if ((m = L.match(/P\/?O\s*NO\.?\s*[:.]?\s*([A-Z][A-Z&]{1,5}\d{7}|CH\d{8})/i))) { po = m[1].toUpperCase(); color = ''; lot = ''; continue; }
    if ((m = L.match(/^#\s*(\S+)\s+(.+)$/))) { color = (m[1] + ' ' + m[2]).replace(/\s+/g, ' ').trim(); lot = ''; continue; }
    if ((m = L.match(/^LOT\s*NO\.?\s*[:.]?\s*(\S+)/i))) { lot = m[1]; continue; }
    if (/^TOTAL\b/i.test(L) || !po || !color) continue;
    m = L.match(/^(\S+)\s+(\d[\d,]*(?:\.\d+)?)\s*(YDS?|YARDS?|MTRS?|MTS|M|KGS?|PCS)\b/i);
    if (!m || /ROLLS?$/i.test(m[1])) continue;
    const unit = /^Y/i.test(m[3]) ? 'YD' : /^K/i.test(m[3]) ? 'KG' : /^P/i.test(m[3]) ? 'PCS' : 'M';
    const q = Number(m[2].replace(/,/g, ''));
    if (!q) continue;
    const key = po + '|' + AZ(color) + '|' + AZ(lot);
    let g = groups.find((x) => x.key === key);
    if (!g) { g = { key, po, poRaw: po, article: '', color, lot: lot || '(không ghi lô)', unit, rolls: [], total: 0 }; groups.push(g); }
    g.rolls.push({ no: m[1], qty: q });
    g.total = Math.round((g.total + q) * 1000) / 1000;
  }
  if (groups.length < 1 || groups.reduce((a, g) => a + g.rolls.length, 0) < 3) return null;
  const no = (t.match(/INVOICE\s*NO\.?\s*[:.]?\s*([A-Z0-9][A-Z0-9\/-]{3,})/i) || [])[1] || '';
  return {
    profile: 'GEN', supplier: supplierNameOf(lines), role: 'pkl', file: fname || '',
    inv: { no, invNo: no, invDate: '', items: [], currency: '', unitDefault: groups[0].unit, total: NaN, gen: true },
    pkl: { groups, lots: groups, unit: groups[0].unit, level: 'lot', soft: false },
    pklOnly: true,
  };
}

function textToPages(text) {
  const pages = [];
  for (const pg of String(text).split(/\x0c/)) {
    const items = [];
    let prev = '';
    pg.split(/\r\n?|\n/).forEach((line, li) => {
      /* dòng tổng không ghi chữ TOTAL, kẹp dưới đường kẻ "--------" (chữ dàn cột kiểu Chain Guan):
         chỉ có số + đơn vị/tiền tệ → đánh dấu TOTAL để bộ đọc chung không coi là mặt hàng */
      const isTot = /^[\s\-=_]{6,}$/.test(prev) && /\d/.test(line)
        && !line.replace(/USD|US\$|EUR|VND|RMB|HKD|YARDS?|YDS?|MTRS?|MTS|KGS?|PCS|PRS|ROLLS?|DZ/gi, '').replace(/[\d.,\s]/g, '');
      if (line.trim()) prev = line;
      const re = /\S+(?: \S+)*/g;
      let m, first = true;
      while ((m = re.exec(line))) {
        /* số dính đơn vị / tiền tệ: "1,048YDS" → "1,048 YDS", "USD1.10" → "USD 1.10" */
        const t = m[0].replace(/(\d)(YDS?|YARDS?|MTRS?|MTS|KGS?|PCS|PRS|ROLLS?|DZ)\b/gi, '$1 $2')
          .replace(/\b(USD|US\$|EUR|VND|RMB|HKD)\s*(?=\d)/gi, '$1 ');
        items.push({ x: m.index * 6, y: 2000 - li * 12, w: m[0].length * 6, s: (isTot && first ? 'TOTAL ' : '') + t });
        first = false;
      }
    });
    if (items.length) pages.push(items);
  }
  return pages;
}

async function pdfPages(file) {
  if (PDF_PAGES.has(file)) return PDF_PAGES.get(file);
  if (/\.doc$/i.test(file.name)) {
    const pages = textToPages(await docText(await file.arrayBuffer()));
    PDF_PAGES.set(file, pages);
    return pages;
  }
  const data = new Uint8Array(await file.arrayBuffer());
  const pdf = await pdfjsLib.getDocument({ data, isEvalSupported: false }).promise;
  const pages = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const tc = await page.getTextContent();
    pages.push(tc.items.filter((it) => it.str && it.str.trim()).map((it) => pdfItemXY(it)));
  }
  PDF_PAGES.set(file, pages);
  return pages;
}
/* ---------- OCR cho PDF dạng ảnh (tesseract.js, tải từ CDN khi cần; lần đầu ~8 MB) ---------- */
const OCR_CFG = Object.assign({ script: 'https://cdn.jsdelivr.net/npm/tesseract.js@7/dist/tesseract.min.js', langs: 'eng+vie', scale: 3 }, window.OCR_CFG || {});
let OCR_WORKER = null, OCR_FAIL = '';
function ocrStatus(msg) { const el = $('#ocrstat'); if (el) el.textContent = msg || ''; }
async function loadOcrScript() {
  if (window.Tesseract) return true;
  await new Promise((res, rej) => {
    const sc = document.createElement('script'); sc.src = OCR_CFG.script; sc.onload = res; sc.onerror = () => rej(new Error('không tải được ' + OCR_CFG.script));
    document.head.appendChild(sc);
  });
  return !!window.Tesseract;
}
async function ocrWorker() {
  if (OCR_WORKER) return OCR_WORKER;
  await loadOcrScript();
  const opts = { legacyCore: false, legacyLang: false, logger: (m) => { if (m && m.status && /load|init/i.test(m.status)) ocrStatus(`Đang tải bộ OCR: ${m.status} ${m.progress != null ? Math.round(m.progress * 100) + '%' : ''}`); } };
  if (OCR_CFG.workerPath) opts.workerPath = OCR_CFG.workerPath;
  if (OCR_CFG.corePath) opts.corePath = OCR_CFG.corePath;
  if (OCR_CFG.langPath) opts.langPath = OCR_CFG.langPath;
  OCR_WORKER = await Tesseract.createWorker(OCR_CFG.langs, 1, opts);
  await OCR_WORKER.setParameters({ preserve_interword_spaces: '1' });
  return OCR_WORKER;
}
/* vẽ từng trang PDF ra canvas → nhận dạng → trả về pages giống pdfPages() (toạ độ pt, gốc trái-dưới) */
/* ---- Excel 97-2003 (.xls) → .xlsx bằng SheetJS (nhúng sẵn, chỉ nạp khi gặp file .xls) ---- */
const XLS_CONV = new Set();
const DOC_CONV = new Set();
async function xlsToXlsx(buf) {
  if (typeof window.loadSheetJS !== 'function') throw new Error('thiếu thư viện SheetJS');
  const X = await window.loadSheetJS();
  const wb = X.read(new Uint8Array(buf), { type: 'array', cellDates: true, cellNF: true });
  if (!wb || !wb.SheetNames || !wb.SheetNames.length) throw new Error('không có sheet');
  return X.write(wb, { type: 'array', bookType: 'xlsx', cellDates: true });
}
async function ocrPdfPages(file) {
  const data = new Uint8Array(await file.arrayBuffer());
  const pdf = await pdfjsLib.getDocument({ data, isEvalSupported: false }).promise;
  const worker = await ocrWorker();
  const pages = [];
  const scale = OCR_CFG.scale || 3;
  for (let p = 1; p <= pdf.numPages; p++) {
    ocrStatus(`Đang OCR ${file.name} — trang ${p}/${pdf.numPages}…`);
    const page = await pdf.getPage(p);
    const vp = page.getViewport({ scale });
    const canvas = document.createElement('canvas');
    canvas.width = Math.ceil(vp.width); canvas.height = Math.ceil(vp.height);
    await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise;
    const { data: d } = await worker.recognize(canvas, {}, { blocks: true, text: false });
    const words = [];
    let li = 0;
    for (const b of (d.blocks || [])) for (const pg of (b.paragraphs || [])) for (const ln of (pg.lines || [])) {
      li++;
      for (const w of (ln.words || [])) if (w.text && w.text.trim()) words.push({ line: li, x0: w.bbox.x0, y0: w.bbox.y0, x1: w.bbox.x1, y1: w.bbox.y1, text: w.text, conf: w.confidence });
    }
    pages.push(ocrWordsToItems(words, scale, canvas.height));
  }
  ocrStatus('');
  (window.__OCR__ = window.__OCR__ || {})[file.name] = pages;
  return pages;
}
function linesFromPages(pages) {
  const out = [];
  for (const items of pages) {
    const its = items.slice().sort((a, b) => (b.y - a.y) || (a.x - b.x));
    let cur = [], curY = null;
    const flush = () => {
      if (!cur.length) return;
      cur.sort((a, b) => a.x - b.x);
      out.push(cur.map((c) => c.s).join(' ').replace(/\s+/g, ' ').trim());
      cur = [];
    };
    for (const it of its) {
      if (curY === null || Math.abs(it.y - curY) <= 3) { cur.push(it); curY = curY === null ? it.y : curY; }
      else { flush(); cur = [it]; curY = it.y; }
    }
    flush();
  }
  return out;
}
async function pdfLines(file) { return linesFromPages(await pdfPages(file)); }

/* ================= Thiên Gia (in bao bì, trimming) =================
   Hóa đơn GTGT: "2 TAG PAPER (L100xW70MM) PO TGB0052700 PC 3.270 330 1.079.100" — mã hàng là KÍCH THƯỚC.
   Packing list PDF: "THIEN GIA-TGB0052700 Tag Paper (L100xW70mm) 71423.01 pcs 1,000 …" — có mã code
   (khớp cột Specification "code 71423.01 - W26" của inbound) → điền Invoice Quantity từng dòng. */
const isTtgText = (t) => /THI[ÊE]N\s*GIA/i.test(t);
const RE_DIM = /L\s*([\d]+(?:[.,]\d+)?)\s*[x*×]\s*W\s*([\d]+(?:[.,]\d+)?)\s*(MM|CM)?/i;
/* kích thước quy về mm để so: "L23.8xW14.8CM" = "L238xW148MM"; không ghi đơn vị → coi là mm */
function dimKey(t) {
  const m = String(t || '').match(RE_DIM);
  if (!m) return '';
  const k = /CM/i.test(m[3] || '') ? 10 : 1;
  const f = (x) => String(Math.round(parseFloat(String(x).replace(',', '.')) * k * 10) / 10);
  return f(m[1]) + 'x' + f(m[2]);
}
/* mã code dạng 71035.10 → so theo giá trị số (packing list Excel/PDF hay rơi số 0 cuối: 71035.1) */
const ttgCodes = (t) => (String(t || '').match(/\b\d{4,6}\.\d{1,3}\b/g) || []).map((x) => String(parseFloat(x)));
const qtyEn = (x) => { const s = String(x).trim(); return /^\d{1,3}(,\d{3})+$/.test(s) ? Number(s.replace(/,/g, '')) : num(s); };

function parseTtgItems(bodyCut) {
  const RE_Q = /^\s*(\d{1,3})?\s*(.*?)\s*\b(ROL|ROLL|PC|PCS|C[ÁA]I|CU[ỘO]N|SET|B[ỘO]|KG|M|T[ỜO])\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$/i;
  const items = [];
  let buf = [];
  for (const L of bodyCut) {
    if (/6\s*=\s*4\s*x\s*5|\(No\.\)|T[êe]n h[àa]ng h[óo]a/i.test(L)) { buf = []; continue; }
    const m = L.match(RE_Q);
    if (m && RE_Q.test(L) && !isNaN(num(m[4]))) {
      const text = buf.slice(-2).concat([m[2]]).join(' ').trim();
      const it = { code: '', po: '', text: (text + ' ' + L).trim(), qty: num(m[4]), price: num(m[5]), amount: num(m[6]), unit: m[3].toUpperCase() };
      const pm = it.text.match(/\bPO\s*[:\-]?\s*([A-Z]{2,6}\d{7})\b/i);
      if (pm) it.po = pm[1].toUpperCase();
      const dm = it.text.match(RE_DIM);
      if (dm) { it.code = dm[0].replace(/\s+/g, ''); it.dim = dimKey(dm[0]); it.exactVars = [norm(it.code)]; }
      it.type = text.replace(RE_DIM, '').replace(/\bPO\b.*$/i, '').replace(/[()]/g, ' ').replace(/\s+/g, ' ').trim();
      it.priceNarrow = true; it.ttg = true;
      items.push(it); buf = [];
      continue;
    }
    const last = items[items.length - 1];
    const lonePo = L.trim().match(/^(?:PO\s*)?([A-Z]{2,6}\d{7})$/i);
    if (lonePo && last && !last.po) { last.po = lonePo[1].toUpperCase(); last.text += ' ' + L; continue; }
    buf.push(L);
  }
  return items.filter((x) => x.code);
}

function readTtgPkl(lines) {
  const t = lines.join(' ');
  if (!isTtgText(t) || /H[ÓO]A Đ[ƠO]N GI[ÁA] TR[ỊI] GIA T[ĂA]NG/i.test(t)) return null;
  if (!lines.some((L) => /\bPO\b.*T[ÊE]N H[ÀA]NG.*Quantity/i.test(L))) return null;
  const RE = /([A-Z]{2,6}\d{7})\s+(.*?)\s+(Rol|Roll|pcs|pc|c[áa]i|set|kg|m)\s+([\d][\d.,]*)\b/i;
  const rows = [];
  for (const L of lines) {
    const m = L.match(RE);
    if (!m) continue;
    const name = m[2];
    const dm = name.match(RE_DIM);
    const after = dm ? name.slice(name.indexOf(dm[0]) + dm[0].length).replace(/^\s*\)?\s*/, '').trim() : '';
    const code = (ttgCodes(after)[0]) || '';
    const q = qtyEn(m[4]);
    if (!q || isNaN(q)) continue;
    const inv = (L.match(/\b(\d{3,8})\s*$/) || [])[1] || '';
    rows.push({
      material: '', po: m[1].toUpperCase(), name: name.replace(/\s+/g, ' ').trim(), dim: dimKey(dm ? dm[0] : ''),
      code, color: code ? '' : after, spec: code, ref: '', size: '', order: '', qty: q, sheet: '', invNo: inv,
    });
  }
  if (!rows.length) return null;
  const pos = [...new Set(rows.map((r) => r.po))];
  const no = (t.match(/No\.?\s*:?\s*(PGH-[\w-]+)/i) || [])[1] || '';
  return { po: pos[0], pos, rows, ttg: true, docNo: no, total: rows.reduce((a, b) => a + b.qty, 0),
    invNos: [...new Set(rows.map((r) => r.invNo).filter(Boolean))] };
}

/* ================= Invoice ================= */
function parseInvoice(lines) {
  const all = lines.join('\n');
  const flat = all.replace(/\s+/g, ' ');

  let serial = (flat.match(/K[ýy]\s*hi[ệe]u[^:]*:\s*([A-Z0-9]+)/i) || [])[1] || '';
  let no = (flat.match(/S[ốo]\s*\(\s*No\.?\s*\)\s*:?\s*(\d+)/i) || flat.match(/Invoice\s*No\.?\s*\)?\s*:?\s*(\d+)\b/i)
    || flat.match(/\bNo\.?\s*\)\s*:?\s*(\d+)\b/i)
    || flat.match(/S[ốo]\s*:\s*(\d{4,})/i) || [])[1] || '';

  let d = null, m = null, y = null;
  let md = flat.match(/Ng[àa]y\s*\(?\s*date\s*\)?\s*(\d{1,2})\s*th[áa]ng\s*\(?\s*month\s*\)?\s*(\d{1,2})\s*n[ăa]m\s*\(?\s*year\s*\)?\s*(\d{4})/i);
  if (md) { d = md[1]; m = md[2]; y = md[3]; }
  if (!md) {   // "Ngày29tháng09năm2026" (hóa đơn không có chữ "date")
    md = flat.match(/Ng[àa]y\s*(\d{1,2})\s*th[áa]ng\s*(\d{1,2})\s*n[ăa]m\s*(\d{4})/i);
    if (md) { d = md[1]; m = md[2]; y = md[3]; }
  }
  if (!md) {
    md = flat.match(/K[ýy]\s*ng[àa]y\s*(\d{1,2})\/(\d{1,2})\/(\d{4})/i) || flat.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (md) { d = md[1]; m = md[2]; y = md[3]; }
  }
  const invDate = d ? `${String(d).padStart(2, '0')}.${String(m).padStart(2, '0')}.${y}` : '';
  const invNo = serial && no ? `${serial}#${String(no).padStart(8, '0')}` : '';

  // tổng tiền trên hóa đơn
  const g = (re) => { const x = flat.match(re); return x ? num(x[1]) : NaN; };
  const total = g(/C[ộo]ng ti[ềe]n h[àa]ng[^:]{0,30}:?\s*([\d.,]+)/i);
  const vat = g(/Ti[ềe]n thu[ếe] GTGT[^:]{0,30}:?\s*([\d.,]+)/i);
  const payment = g(/T[ổo]ng c[ộo]ng ti[ềe]n thanh to[áa]n[^:]{0,30}:?\s*([\d.,]+)/i);

  // line items — mã hàng nằm giữa dấu "/" đầu tiên và "//": "… / LB 5873 // TRIMMINGVN-0725"
  const rePo = /\b([A-Z]{4,}(?:VN)?)-(\d{3,5})\b/;
  const reQty = /(?:C[áa]i|PCS|Pcs|Chi[ếe]c|PC)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)/;

  const items = [];
  let pendingCode = null, pendingPo = null, cur = null, buf = [];
  const startIdx = lines.findIndex((l) => /T[êe]n h[àa]ng h[óo]a|Description/i.test(l));
  const body = startIdx >= 0 ? lines.slice(startIdx) : lines;

  /* ---- Dạng B (Inkava): PO nằm ở dòng TRƯỚC dòng số lượng, mã hàng ở dòng SAU
     "PODUY0095800- Label L56xW20MM-" / "1 PCS 16.988,00 100 1.698.800" / "ALABPRWV0228"
     Nhận ra khi cả hóa đơn không có dấu "//" mà lại có dòng bắt đầu bằng PO<mã>.   */
  const RE_POPRE = /\bPO\s*([A-Z]{2,6}\d{4,})/i;
  const endIdx = body.findIndex((L) => /C[ộo]ng ti[ềe]n h[àa]ng|T[ổo]ng s[ốo] l[ưu][ợo]ng|Total amount/i.test(L));
  const bodyCut = endIdx > 0 ? body.slice(0, endIdx) : body;
  /* ---- Dạng C (Thiên Gia): "TAG PAPER (L100xW70MM) PO TGB0052700 PC 3.270 330 1.079.100" — mã hàng là kích thước */
  if (isTtgText(flat)) {
    const its = parseTtgItems(bodyCut);
    if (its.length) return { serial, no, invNo, invDate, items: its, total, vat, payment, supplier: 'Thiên Gia' };
  }
  const usePoPre = !bodyCut.some((L) => L.includes('//')) && bodyCut.some((L) => RE_POPRE.test(L));
  if (usePoPre) {
    const codesIn = (t, po) => (String(t).toUpperCase().match(/\b[A-Z]{3,}\d{3,}\b/g) || [])
      .filter((x) => !po || !x.includes(String(po).toUpperCase()));
    let cur2 = null, buf2 = [];
    for (const L of bodyCut) {
      const mq = L.match(reQty);
      if (mq) {
        const ctx = buf2.concat([L]).join(' ');
        const po = (ctx.match(RE_POPRE) || [])[1] || '';
        cur2 = { code: '', po: po.toUpperCase(), text: ctx, qty: num(mq[1]), price: num(mq[2]), amount: num(mq[3]) };
        const c0 = codesIn(ctx, po);
        if (c0.length) cur2.code = c0[c0.length - 1];
        items.push(cur2);
        buf2 = [];
        continue;
      }
      const isPoLine = RE_POPRE.test(L);
      if (cur2 && !cur2.code && !isPoLine) {
        const c = codesIn(L, cur2.po);
        if (c.length) { cur2.code = c[c.length - 1]; cur2.text += ' ' + L; continue; }
      }
      if (isPoLine) { cur2 = null; buf2 = [L]; }
      else if (cur2) cur2.text += ' ' + L;
      else buf2.push(L);
    }
    return { serial, no, invNo, invDate, items, total, vat, payment };
  }

  let frag = null;
  for (const L of body) {
    if (/C[ộo]ng ti[ềe]n h[àa]ng|Total amount/i.test(L)) break;
    // mã hàng nằm giữa "/" và "//", có thể bị xuống dòng: "… / LM-" ⏎ "RFIDST22 // TRIMMINGVN-0708"
    let code = null;
    const k = L.indexOf('//');
    if (k >= 0) {
      const pre = L.slice(0, k);
      if (pre.includes('/')) code = pre.replace(/^[^\/]*\//, '').trim();
      else if (frag) code = (/-$/.test(frag) ? frag + pre.trim() : frag + ' ' + pre.trim()).trim();
      frag = null;
    } else if (L.includes('/')) {
      const tail = L.slice(L.lastIndexOf('/') + 1).trim();
      if (tail && tail.length <= 40) frag = tail;
    }
    const isCode = code && !/T[êe]n h[àa]ng/i.test(L) && code.length <= 40;
    if (isCode) {
      const last = items.length ? items[items.length - 1] : null;
      if (last && !last.code) { last.code = code; last.text += ' ' + L; }   // mã hàng nằm sau dòng số lượng
      else { pendingCode = code; cur = null; buf = [L]; }
    }

    const mq = L.match(reQty);
    if (mq) {
      cur = {
        code: pendingCode || '', po: pendingPo || '', text: buf.concat([L]).join(' '),
        qty: num(mq[1]), price: num(mq[2]), amount: num(mq[3]),
      };
      items.push(cur);
      pendingCode = null; pendingPo = null; buf = [];
    } else if (!isCode) {
      if (cur) cur.text += ' ' + L; else buf.push(L);
    }

    const mp = L.match(rePo);
    if (mp) {
      const po = mp[1] + '-' + mp[2];
      const last = items.length ? items[items.length - 1] : null;
      if (last && !last.po) last.po = po; else pendingPo = po;
    }
  }
  /* Hóa đơn ITL từ 10.2026 ghi thẳng PO ScaF sau "//": "… / VCH101F // TRI0012500" */
  for (const it of items) {
    if (it.po) continue;
    const m1 = (it.text || '').toUpperCase().match(/\/\/\s*([A-Z]{2,6}\d{4,})/);
    if (m1) { it.po = m1[1]; continue; }
    const m2 = (it.text || '').toUpperCase().match(/\b([A-Z]{2,6}\d{7})\b/);
    if (m2) it.po = m2[1];
  }

  /* Dạng mô tả khác (Inkava): "PODUY0095800- Label L56xW20MM- ALABPRWV0228"
     — không có dấu "//", PO dán liền chữ PO, mã hàng là token cuối.            */
  for (const it of items) {
    if (!it.po) {
      const mm = it.text.match(/\bPO\s*[:\-]?\s*([A-Z]{2,6}\d{4,})/i);
      if (mm) it.po = mm[1].toUpperCase();
    }
    if (!it.code) {
      const cands = (it.text.toUpperCase().match(/\b[A-Z]{3,}\d{3,}\b/g) || [])
        .filter((x) => !it.po || !x.includes(it.po.toUpperCase()));
      if (cands.length) it.code = cands[cands.length - 1];
    }
  }
  return { serial, no, invNo, invDate, items, total, vat, payment };
}

/* Các dạng viết khác nhau của cùng một mã hàng, từ chi tiết nhất đến chung nhất.
   "LB 07780 C/1" → LB07780C/1 · LB7780C1 · LB7780C01 · LB7780 · LB07780      */
function codeVariants(raw) {
  const s = norm(raw);
  const out = [s];
  const m = String(raw || '').trim().match(/^([A-Za-z]+)\s*0*(\d+)\s*(.*)$/);
  if (m) {
    const pre = m[1].toUpperCase(), num = m[2], suf = norm(m[3]);
    const nz = String(Number(num));
    const bases = [...new Set([pre + nz, pre + num, pre + nz.padStart(4, '0'), pre + nz.padStart(5, '0')])];
    if (suf) for (const b of bases) out.push(b + suf, b + suf.replace(/\//g, ''), b + suf.replace(/\//g, '0'));
    out.push(...bases);
  }
  return [...new Set(out.filter((x) => x.length >= 3))];
}

/* Mã PO đã đúng dạng SAP: 2–6 chữ cái + đúng 7 chữ số (TRI0012500, TEC0002400, DUY0094000) */
const RE_SAPPO = /^[A-Z]{2,6}\d{7}$/;

/* Tra PO: ưu tiên PO ScaX (TRIMMINGVN-xxxx) → PO ScaF trong file PO →
   nếu bản thân mã trên hóa đơn đã đúng dạng SAP thì dùng luôn (không cần file PO) */
function resolvePo(item, poIdx) {
  const text = (item.text || '') + ' ' + (item.po || '');
  const raw = [...new Set(text.toUpperCase().match(/[A-Z0-9][A-Z0-9\-\/]{3,}/g) || [])];
  const toks = [...new Set([].concat(...raw.map((t) => [t, t.replace(/[-\/]+$/, ''), t.replace(/^PO/, '').replace(/[-\/]+$/, '')])))]
    .filter((t) => t.length >= 4);
  if (poIdx) {
    for (const t of toks) if (poIdx.scax.has(t)) return { scax: t, sap: poIdx.scax.get(t), via: 'ScaX' };
    for (const t of toks) if (poIdx.sap.has(t)) return { scax: '', sap: poIdx.sap.get(t), via: 'ScaF' };
  }
  for (const t of toks) {
    if (!RE_SAPPO.test(t)) continue;
    /* "PODUY0095800" là chữ PO dán liền mã PO → ưu tiên phần sau chữ PO */
    const bare = t.startsWith('PO') && RE_SAPPO.test(t.slice(2)) ? t.slice(2) : t;
    return { scax: '', sap: bare, via: 'ScaF (sẵn trên hóa đơn)' };
  }
  return { scax: item.po || '', sap: '', via: '' };
}

/* ================= Packing list ================= */
const RE_SIZE = /(?:^|\s)([A-Z0-9]{3,6})?\s*(XXXL|XXL|XL|XS|S|M|L)\s*\/\s*(?:XXXG|XXG|XG|XP|P|G|M|L|S)\s*-/;
const RE_PCS = /([\d.,]+)\s*PCS/i;
const RE_JUNK = /(CTN\s*No|Our Ref|Label Ref|Description of|Measure|Colour|Customer|Reference|Quantity|Goods|Descr|N\.W|G\.\s*W|\(kg\)|\(cm\)|PACKING LIST|Despatch Note|Delivery|Invoice Date|Attn|Please note)/i;

// Label Ref của một dòng hàng: ưu tiên mã có khoảng trắng (LB 5873) để không nhầm với Our Ref (VN090131)
function labelOf(line) {
  const cands = [...line.matchAll(/\b([A-Z]{2,3})(\s?)(\d{3,6})\b/g)];
  if (!cands.length) return null;
  const spaced = cands.filter((m) => m[2]);
  const notVN = (spaced.length ? spaced : cands).filter((m) => m[1] !== 'VN');
  const pick = (notVN.length ? notVN : (spaced.length ? spaced : cands))[0];
  return norm(pick[1] + pick[3]);
}

/* Nhóm size của packing list ghi dạng <size nội bộ>/<size quốc tế>, có thể xuống dòng:
   XS/XP → XS · S-DD/P-DD → S-DD · "XL/XXL / XG/XXG" → XL/XXL · "M-DD/M- DD" → M-DD  */
function sizeKey(raw) {
  let s = String(raw || '').trim().replace(/\s*-\s*$/, '');
  const sp = s.split(/\s+\/\s+/);
  const local = sp.length > 1 ? sp[0] : (s.includes('/') ? s.slice(0, s.indexOf('/')) : s);
  return local.replace(/\s+/g, '').toUpperCase();
}

function parsePacking(lines) {
  const groups = [];
  let cur = null;
  for (const L of lines) {
    const isHead = /\bVN\d{5,}/.test(L) && /([\d][\d.,]*)\s*$/.test(L) && !RE_PCS.test(L);
    if (isHead) {
      const clean = L.replace(/\bVN\d{5,}\w*/g, ' ');
      const lb = labelOf(clean);
      if (lb) {
        cur = { label: lb, qty: num(L.match(/([\d][\d.,]*)\s*$/)[1]), po: '', poSap: '', inv: {}, raw: [L], text: L };
        const ms = clean.match(/\b([A-Z]{2,4}\d{6,})\b/);   // PO ScaF ghi thẳng ở dòng đầu (TRI0004800)
        if (ms) cur.poSap = ms[1].toUpperCase();
        groups.push(cur);
        continue;
      }
    }
    if (!cur) continue;
    cur.raw.push(L);
    // bỏ các dòng tiêu đề bảng lặp lại giữa trang (chúng chen vào giữa dòng size và dòng số lượng)
    if (!(RE_JUNK.test(L) && !RE_PCS.test(L))) cur.text += ' ' + L;
    if (!cur.po) {
      const m1 = L.match(/([A-Z]{4,})\s*-\s*(\d{3,5})\b/);
      const m2 = L.match(/-\s*(\d{3,5})\s+[A-Z]{2,3}\s?\d{3,6}/);
      if (m1) cur.po = (m1[1] + '-' + m1[2]).toUpperCase();
      else if (m2) cur.po = m2[1];
    }
  }
  // đọc bảng size trên toàn bộ text của nhóm (size + số lượng có thể nằm ở 2–3 dòng khác nhau)
  for (const g of groups) {
    const t = g.text.replace(/\s+/g, ' ');
    // PO khách hàng có thể bị tách dòng: "PO TRIMMINGVN … -0730"
    if (!g.po) {
      const mp = t.match(/\b([A-Z]{5,})\b[\s\S]{0,150}?[-–]\s*(\d{3,5})\b/);
      if (mp && !/^(PACKING|DESPATCH|INVOICE|LABEL)$/.test(mp[1])) g.po = mp[1] + '-' + mp[2];
    }
    const add = (code, rawSize, q) => {
      const s = sizeKey(rawSize);
      if (!s || isNaN(q)) return false;
      g.inv[code] = g.inv[code] || {};
      g.inv[code][s] = (g.inv[code][s] || 0) + q;
      return true;
    };
    let found = false, m;
    // <mã invoice> <size> - <số lượng> PCS ; mã invoice luôn có cả chữ và số (54A2, 7VWB…)
    const re = /\b([A-Z0-9]{3,6})\s+([A-Z][A-Z0-9\/\- ]{0,18}?)\s*-\s*([\d.,]+)\s*PCS/g;
    while ((m = re.exec(t))) {
      const code = m[1].toUpperCase(), rawSize = m[2];
      const okCode = /[A-Z]/.test(code) && /[0-9]/.test(code);
      const okSize = !/[A-Z]{2}\s?\d{3,6}/.test(rawSize);   // tránh nuốt nhầm "LB 5874" vào phần size
      if (okCode && okSize) { if (add(code, rawSize, num(m[3]))) found = true; }
      else re.lastIndex = m.index + 1;                       // bỏ qua, dò lại từ ký tự kế tiếp
    }
    if (!found) { // packing list không ghi mã invoice trên từng dòng size
      const re2 = /([A-Z][A-Z0-9\/\- ]{0,18}?)\s*-\s*([\d.,]+)\s*PCS/g;
      while ((m = re2.exec(t))) {
        /* phần "size" có thể dính cả Label Ref phía trước: "LB 6742 XS - 100 PCS".
           Cắt bỏ mọi thứ tới hết mã nhãn, phần còn lại mới là size.
           (Trước đây cả cụm bị loại → mất luôn nhóm size đầu tiên của dòng hàng.) */
        const sz = String(m[1]).replace(/.*\b[A-Z]{2,3}\s?\d{3,6}\b/, '').trim();
        if (sz) add('?', sz, num(m[2]));
      }
    }
  }
  return groups;
}

/* ================= PO file (ScaX -> SAP) ================= */
function readPoWb(wb) {
  const ws = wb.worksheets[0];
  let hdrRow = 1, cPo = 0, cScax = 0, cSpec = 0, cSize = 0, cQty = 0, cPrice = 0, cName = 0;
  for (let r = 1; r <= Math.min(15, ws.rowCount); r++) {
    const vals = ws.getRow(r).values || [];
    const idx = (names) => {
      for (let i = 1; i < vals.length; i++) {
        const v = String(vals[i] == null ? '' : vals[i]).replace(/\s+/g, ' ').trim().toLowerCase();
        if (names.includes(v)) return i;
      }
      return 0;
    };
    const a = idx(['po no scax', 'po no. scax']);
    if (a) {
      hdrRow = r; cScax = a;
      cPo = idx(['po no.', 'po no']);
      cSpec = idx(['specification']);
      cSize = idx(['size cup', 'size']);
      cQty = idx(['order quantity']);
      cPrice = idx(['price']);
      cName = idx(['rm name']);
      break;
    }
  }
  if (!cPo) throw new Error('Không tìm thấy cột "PO No." trong file PO SCAF-SCAX.');
  const map = new Map(); const sap = new Map(); const rows = [];
  for (let r = hdrRow + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const scax = cScax ? row.getCell(cScax).text.trim() : '';
    const po = row.getCell(cPo).text.trim();
    if (!po) continue;
    if (scax && !map.has(scax.toUpperCase())) map.set(scax.toUpperCase(), po);
    if (!sap.has(po.toUpperCase())) sap.set(po.toUpperCase(), po);
    rows.push({
      scax: scax.toUpperCase(), po,
      spec: cSpec ? row.getCell(cSpec).text.trim() : '',
      size: cSize ? row.getCell(cSize).text.trim() : '',
      qty: cQty ? num(row.getCell(cQty).value) : NaN,
      price: cPrice ? num(row.getCell(cPrice).value) : NaN,
      name: cName ? row.getCell(cName).text.trim() : '',
    });
  }
  return { map, sap, rows };
}


/* ================= Nhận diện & gom nhóm file ================= */
const KIND_NAME = { inb: 'File inbound', inv: 'Hóa đơn', pkl: 'Packing list', po: 'PO SCAF-SCAX', fab: 'Chứng từ vải (HĐ+PKL)', pklx: 'Packing list Excel (theo PO)' };
const CACHE = new WeakMap();
const STATE = { po: null, items: [], groups: [], busy: false };
let RESULT = null;

const hits = (t, arr) => arr.reduce((n, s) => n + (t.includes(s) ? 1 : 0), 0);
const dirOf = (f) => {
  const p = f.webkitRelativePath || f._relPath || '';
  const i = p.lastIndexOf('/');
  return i > 0 ? p.slice(0, i) : '';
};

function headerIndex(ws) {
  const vals = ws.getRow(1).values || [];
  const H = {};
  for (let i = 1; i < vals.length; i++) {
    const v = String(vals[i] == null ? '' : vals[i]).replace(/\s+/g, ' ').trim().toLowerCase();
    if (v) H[v] = i;
  }
  const pick = (...names) => { for (const n of names) if (H[n]) return H[n]; return 0; };
  return {
    po: pick('purchasing document'), material: pick('material'), desc: pick('material description'),
    size: pick('size'), spec: pick('specification'), price: pick('gross price'), sur: pick('surcharge item'),
    qty: pick('quantity'), deliv: pick('delivered qty'), invQty: pick('invoice quantity', 'input quantity'),
    invNo: pick('invoice number'), invDate: pick('invoice date'),
    supRef: pick('supplier ref', 'supplier mat. no.', 'supplier material number', 'supplier material no.', 'supplier mat no'), color: pick('color', 'colour'), lapdip: pick('lapdip color'),
    overTol: pick('over tolerance qty'), unit: pick('base unit of measure'), cur: pick('currency'),
    last: Math.max(...Object.values(H), 1),
  };
}

/* tên chủ hàng: dòng đầu có CO., LTD / LIMITED / COMPANY / S.P.A … */
/* ================= Chủ hàng đã được huấn luyện =================
   Chứng từ của chủ hàng KHÔNG có trong danh sách này vẫn được bộ đọc chung xử lý, nhưng công cụ cảnh báo và đề
   nghị buyer liên hệ người phụ trách để bổ sung mẫu. Thêm chủ hàng mới: thêm một dòng vào TRAINED_SUPPLIERS. */
const CONTACT = { name: 'anh Quốc Tú', email: 'quoctu.nguyen@blaogroup.com' };
const TRAINED_SUPPLIERS = [
  ['ITL', /\bITL\b/], ['Inkava', /INKAVA/], ['Thiên Gia', /THIEN\s*GIA/], ['Fujian Techwork', /TECHWORK/],
  ['New Style – BLAO', /NEW\s*STYLE/], ['Quanzhou Hengyu', /HENGYU/], ['J&H Yubo', /YUBO/],
  ['Capital Tricot', /CAPITAL\s*TRICOT/], ['Carvico', /CARVICO/], ['Celeb', /\bCELEB\b/], ['Cheung Hing', /CHEUNG\s*HING/],
  ['Chuangjie', /CHUANGJIE/], ['Derun', /\bDERUN\b/], ['DJIC', /\bDJIC\b/], ['Dongguan Uwork', /UWORK/],
  ['Freetex', /FREETEX/], ['Fujian Baikai', /BAIKAI/], ['Fujian Honggang', /HONGGANG/], ['Hing Yip', /HING\s*YIP/],
  ['Hoa Nghiêm', /HOA\s*NGHIEM/], ['Best Pacific', /BEST\s*PACIFIC/], ['Junye', /\bJUNYE\b/], ['Luen Hing', /LUEN\s*HING/],
  ['Pioneer', /\bPIONEER\b/], ['PT Winner', /\bWINNER\b/], ['S&M', /\bS\s*&\s*M\b/], ['Seamless', /\bSEAMLESS\b/],
  ['Stretchline', /STRETCHLINE/], ['SunPo', /\bSUN\s*PO\b/], ['Yibei', /\bYIBEI\b/], ['Brugnoli', /BRUGNOLI/],
  ['AIM High', /\bAIM\s*HIGH\b/], ['Chain Guan', /CHAIN\s*GUAN/], ['Paddies', /PADDIES/], ['Prestige', /\bPRESTIGE\b/],
  ['Vinity', /\bVINITY\b/],
];
const foldVN = (t) => String(t || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[đĐ]/g, 'D').toUpperCase();
function trainedSupplierOf(text) {
  const f = foldVN(text);
  const hit = TRAINED_SUPPLIERS.find(([, re]) => re.test(f));
  return hit ? hit[0] : '';
}
const contactHtml = () => `${esc(CONTACT.name)} (<a href="mailto:${esc(CONTACT.email)}">${esc(CONTACT.email)}</a>)`;
const contactText = () => `${CONTACT.name} (${CONTACT.email})`;
/* chủ hàng của một bộ chứng từ: mẫu riêng (vải) thì đã huấn luyện; còn lại so tên người bán trên chứng từ
   + Partner Name trong file inbound với danh sách */
function supplierCheck(g, cInb) {
  const fabSup = g.inv && g.inv.fab && g.inv.fab.supplier;
  if (g.isFab && !g.isGen) return { trained: fabSup || g.fabName || 'mẫu riêng', name: fabSup || g.fabName || '', kind: 'mẫu riêng' };
  const c = CACHE.get(g.inv.file) || {};
  const docLines = (c.lines || []).slice(0, 30);
  /* Partner Name của đúng các PO trên hóa đơn; chứng từ không ghi PO → toàn bộ partner trong file */
  const byPo = (cInb && cInb.partnerByPo) || {};
  let parts = [...new Set((g.inv.sapPos || []).map((p) => byPo[String(p).toUpperCase()]).filter(Boolean))];
  if (!parts.length) parts = (cInb && cInb.partners) || [];
  const name = fabSup || supplierNameOf(docLines) || parts[0] || '';
  const fromDoc = trainedSupplierOf([fabSup, docLines.join(' ')].join(' '));
  const fromInb = parts.length === 1 ? trainedSupplierOf(parts[0]) : '';   // nhiều chủ hàng trong cùng file → không kết luận từ inbound
  const trained = fromDoc || fromInb;
  return { trained, name: name || parts[0] || g.inv.file.name, kind: !trained ? 'chưa huấn luyện' : (g.isGen ? 'bộ đọc chung' : 'mẫu riêng') };
}

function supplierNameOf(lines) {
  const L = (lines || []).slice(0, 12).map((x) => String(x).trim()).filter(Boolean);
  const notUs = (x) => !/SCAVI|B'?LAO|BLAO SPORT/i.test(x);
  const hit = L.find((x) => /\b(CO\.?,?\s*LTD|LIMITED|COMPANY|CORP|S\.P\.A|INC\b|CÔNG TY|CONG TY|TNHH|GMBH|S\.A\b)/i.test(x) && notUs(x))
    || L.find((x) => /\b(FACTORY|GROUP|INDUSTR|TEXTILE|TRADING)/i.test(x) && notUs(x));
  /* cắt phần địa chỉ/ghi chú đứng sau tên: "Carvico S.p.A. * Sede legale Via…" → "Carvico S.p.A." */
  let s = (hit || L[0] || '').replace(/\s[*|•·]\s.*$/, '').replace(/\s{2,}.*$/, '');
  const m = hit ? s.match(/\b(CO\.?,?\s*LTD\.?|LIMITED|COMPANY|CORP\.?|S\.P\.A\.?|INC\b\.?|GMBH|S\.A\b\.?)/i) : null;
  if (m && !/^(CÔNG TY|CONG TY|TNHH)/i.test(s)) s = s.slice(0, m.index + m[0].length);
  return s.trim().slice(0, 48);
}
function supplierNameOfWb(wb) {
  const lines = [];
  const ws = wb.worksheets[0];
  for (let r = 1; r <= Math.min(12, ws.rowCount); r++) for (let c = 1; c <= 12; c++) { let t = ''; try { t = cellText(ws.getCell(r, c).value).trim(); } catch (e) { t = ''; } if (t) lines.push(t); }
  return supplierNameOf(lines);
}

async function classify(file) {
  if (CACHE.has(file)) return CACHE.get(file);
  const name = (file.name || '').toLowerCase();
  let out;
  if (/\.(pdf|doc)$/.test(name)) {
    let pages;
    try { pages = await pdfPages(file); }
    catch (e) { if (/\.doc$/.test(name)) { out = { kind: 'doc', score: 0, why: e && e.message ? e.message : String(e) }; CACHE.set(file, out); return out; } throw e; }
    if (/\.doc$/.test(name)) {
      DOC_CONV.add(file);
      const mp = readMonoPkl(linesFromPages(pages), file.name);
      if (mp) { out = { kind: 'genpkl', score: 5, gen: mp, lines: linesFromPages(pages), pos: [...new Set(mp.pkl.groups.map((g) => g.po))] }; CACHE.set(file, out); return out; }
    }
    let lines = linesFromPages(pages);
    const nChars = pages.reduce((a, p) => a + p.reduce((b, i) => b + i.s.length, 0), 0);
    let ocr = false;
    if (nChars < 40 && /\.doc$/.test(name)) { out = { kind: 'doc', score: 0, why: 'file Word không có chữ' }; CACHE.set(file, out); return out; }
    if (nChars < 40) {
      const want = $('#ocr') ? $('#ocr').checked : true;
      if (!want) { out = { kind: 'scan', score: 0, lines, why: 'đã tắt OCR' }; CACHE.set(file, out); return out; }
      try {
        pages = await ocrPdfPages(file);
        lines = linesFromPages(pages);
        ocr = true;
        if (lines.join(' ').replace(/\s+/g, '').length < 40) { out = { kind: 'scan', score: 0, lines, why: 'OCR không đọc ra chữ' }; CACHE.set(file, out); return out; }
      } catch (e) {
        OCR_FAIL = e && e.message ? e.message : String(e);
        ocrStatus('');
        out = { kind: 'scan', score: 0, lines, why: 'OCR lỗi: ' + OCR_FAIL + ' (cần mạng để tải bộ OCR)' }; CACHE.set(file, out); return out;
      }
    }
    /* Capital Tricot (vải, Thái Lan): INVOICE + PACKING LIST theo kiện trong cùng một PDF → mẫu riêng (fab.js) */
    if (!ocr && isCapitalText(lines.join(' '))) {
      let cap = null;
      try { cap = readCapital(lines); } catch (e) { cap = null; }
      if (cap && cap.inv && cap.inv.items.length) {
        out = { kind: 'fab', score: 9, fab: cap, lines, pos: [...new Set(cap.inv.items.map((x) => x.po).filter(Boolean))] };
        CACHE.set(file, out); return out;
      }
    }
    /* Carvico (vải, Ý): hóa đơn FATTURA/INVOICE → mẫu riêng; packing list PDF rời → nhóm theo màu/lô, gắn vào hóa đơn cùng số PKL */
    if (!ocr && isCarvicoText(lines.join(' '))) {
      let cv = null;
      try { cv = readCarvico(lines); } catch (e) { cv = null; }
      if (cv && cv.inv && cv.inv.items.length) { out = { kind: 'fab', score: 9, fab: cv, lines, pos: [] }; CACHE.set(file, out); return out; }
      let cp = null;
      try { cp = readCarvicoPkl(lines, file.name); } catch (e) { cp = null; }
      if (cp && cp.pkl && cp.pkl.groups.length) { out = { kind: 'genpkl', score: 6, gen: cp, lines, pos: [] }; CACHE.set(file, out); return out; }
    }
    /* packing list PDF của Thiên Gia: PO + kích thước + mã code → dùng như packing list Excel theo PO */
    if (!ocr) {
      const tp = readTtgPkl(lines);
      if (tp) { out = { kind: 'pklx', score: 8, lines, px: tp, pos: tp.pos }; CACHE.set(file, out); return out; }
    }
    const t = lines.join(' ').toUpperCase();
    const sInv = hits(t, ['HÓA ĐƠN GIÁ TRỊ GIA TĂNG', 'VAT INVOICE', 'KÝ HIỆU', 'TIỀN THUẾ', 'NGƯỜI BÁN HÀNG', 'ĐVT']);
    const sPkl = hits(t, ['PACKING LIST', 'DESPATCH NOTE', 'LABEL REF', 'CTN NO', 'OUR REF', 'DELIVERY METHOD']);
    let kind = sInv === 0 && sPkl === 0 ? '' : (sInv >= sPkl ? 'inv' : 'pkl');
    const note = (t.match(/(?:GHI CHÚ \(NOTE\)\s*:|DESPATCH NOTE\s*:?)\s*([A-Z]{2}[A-Z0-9]{5,})/) || [])[1] || '';
    out = { kind, score: Math.max(sInv, sPkl), lines, note, ocr };
    /* hoá đơn GTGT nhưng không theo mẫu ITL/Inkava (không đọc được dòng hàng) → bộ đọc chung */
    const itlItems = kind === 'inv' ? parseInvoice(lines).items.length : 0;
    /* bản scan (OCR) luôn qua bộ đọc chung: mẫu packing list phụ liệu cần chữ chính xác, OCR không đáp ứng */
    if ((kind === 'inv' && !itlItems) || kind === '' || (kind === 'pkl' && (sPkl < 2 || ocr))) {
      let gen = null;
      try { gen = readGenPdf(pages, file.name, dirOf(file)); } catch (e) { gen = null; }
      if (gen && ocr) { gen.ocr = true; if (gen.inv) gen.inv.ocr = true; }
      if (gen && gen.role === 'proforma') out = { kind: 'proforma', score: 0, lines };
      else if (gen && gen.inv && gen.inv.items.length) {
        gen.supplier = gen.supplier || supplierNameOf(lines);
        gen.vatLines = kind === 'inv' ? lines : null;
        if (kind === 'inv') out = { kind: 'inv', score: sInv, lines, note, genFallback: gen };   // HĐ GTGT: có thể là HĐ kèm PKL (Yubo) hoặc chứng từ độc lập
        else if (gen.role === 'pkl') out = { kind: 'genpkl', score: 5, gen, lines, pos: [...new Set(gen.inv.items.map((x) => x.po).filter(Boolean))] };
        else out = { kind: 'fab', score: 8, fab: gen, lines, pos: [...new Set(gen.inv.items.map((x) => x.po).filter(Boolean))], gen: true };
      } else if (gen && gen.pkl && gen.pkl.groups.length) {
        gen.supplier = supplierNameOf(lines);
        out = { kind: 'genpkl', score: 5, gen, lines, pos: [...new Set(gen.pkl.groups.map((x) => x.po).filter(Boolean))] };
      } else if (kind === 'inv' && !itlItems) {
        const pv = parseInvoice(lines);
        out = pv.invNo ? { kind: 'inv', score: sInv, lines, note, ocr, headerOnly: true } : (ocr ? { kind: 'scan', score: 0, lines, why: 'OCR đọc được chữ nhưng không nhận ra bảng hàng (ảnh mờ / có dấu đè)' } : { kind: '', score: 0, lines });
      } else if (ocr) {
        /* bản scan OCR không ra bảng hàng → không đưa vào luồng phụ liệu (sẽ sai), báo để nhập tay */
        out = { kind: 'scan', score: 0, lines, why: 'OCR đọc được chữ nhưng không nhận ra bảng hàng (ảnh mờ / có dấu đè)' + (gen && gen.inv && gen.inv.invNo ? ' — số HĐ đọc được: ' + gen.inv.invNo : '') };
      }
    }
  } else if (/\.xls[xm]?$/.test(name)) {
    let buf = await file.arrayBuffer();
    let fromXls = false;
    if (/\.xls$/.test(name)) {
      /* Excel 97-2003: chuyển sang .xlsx ngay trong trình duyệt (SheetJS) rồi đọc như file .xlsx */
      try { buf = await xlsToXlsx(buf); fromXls = true; }
      catch (e) { out = { kind: 'xls', score: 0, why: e && e.message ? e.message : String(e) }; CACHE.set(file, out); return out; }
    }
    const wb = new ExcelJS.Workbook();
    try { await wb.xlsx.load(buf.slice(0)); }
    catch (e) { if (fromXls) { out = { kind: 'xls', score: 0, why: 'chuyển sang .xlsx xong nhưng không đọc được: ' + (e && e.message ? e.message : e) }; CACHE.set(file, out); return out; } throw e; }
    if (fromXls) XLS_CONV.add(file);
    const fp = fabProfile(wb);
    if (fp) {
      let fab = null;
      try { fab = readFab(wb); } catch (e) { fab = null; }
      if (fab && fab.inv && fab.inv.items.length) {
        out = { kind: 'fab', score: 9, buf, fab, pos: [...new Set(fab.inv.items.map((x) => x.po).filter(Boolean))] };
        CACHE.set(file, out);
        return out;
      }
    }
    /* packing list Excel theo PO (Inkava): cột "Material Code" + Spec/Order No; số PO ghi "PO No: DUY…"
       hoặc chỉ ghi trơn "DUY0081000" ở dòng đầu (file chủ hàng tự soạn lại), hoặc lấy từ tên file.
       Đọc MỌI sheet: sheet "Sheet" thường là số theo PO, các sheet khác ("89-45", "10-15"…) là số thực giao
       → khi đối chiếu sẽ chọn sheet có tổng khớp hóa đơn. */
    {
      const px = readPklxWb(wb, file.name);
      if (px && px.rows.length) { out = { kind: 'pklx', score: 8, buf, px, pos: px.po ? [px.po] : [] }; CACHE.set(file, out); return out; }
    }
    const ws = wb.worksheets[0];
    let head = '';
    for (let r = 1; r <= Math.min(15, ws.rowCount); r++) {
      head += ' ' + (ws.getRow(r).values || []).map((v) => String(v == null ? '' : v)).join(' | ');
    }
    const t = (head + ' ' + wb.worksheets.map((w) => w.name).join(' ')).replace(/\s+/g, ' ').toLowerCase();
    const sPo = hits(t, ['po no scax', 'po follow up', 'order quantity', 'agreed lead-time', 'pr no.']);
    const sInb = hits(t, ['purchasing document', 'invoice quantity', 'over tolerance qty', 'act. gds mvmnt date', 'external delivery id']);
    const kind = sPo === 0 && sInb === 0 ? '' : (sPo >= sInb ? 'po' : 'inb');
    if (kind === 'inb') {
      const H = headerIndex(ws);
      const pos = new Set();
      const partners = new Set();
      const partnerByPo = {};   // PO → Partner Name: xét chủ hàng theo đúng các PO của hóa đơn (file SAP xuất chung nhiều chủ hàng)
      const cPartner = (() => { const row = ws.getRow(1); for (let i = 1; i <= ws.columnCount; i++) if (/partner name/i.test(String(row.getCell(i).text || ''))) return i; return 0; })();
      for (let r = 2; r <= ws.rowCount; r++) {
        const v = ws.getCell(r, H.po).text.trim().toUpperCase();
        if (v) pos.add(v);
        if (!cPartner) continue;
        let t = ''; try { t = cellText(ws.getCell(r, cPartner).value).trim(); } catch (e) { t = ''; }
        if (!t) continue;
        partners.add(t.toUpperCase());
        if (v && !partnerByPo[v]) partnerByPo[v] = t.toUpperCase();
      }
      out = { kind, score: sInb, buf, pos: [...pos], partners: [...partners], partnerByPo };   // không giữ workbook để đỡ tốn bộ nhớ
    } else if (kind === 'po') {
      out = { kind, score: sPo, buf, wb };
    } else {
      /* bộ đọc chung cho chủ hàng chưa có mẫu riêng */
      let gen = null;
      try { gen = readGenWb(wb, file.name, dirOf(file)); } catch (e) { gen = null; }
      if (gen && gen.inv && gen.inv.items.length && gen.role === 'inv') {
        gen.supplier = gen.supplier || supplierNameOfWb(wb);
        out = { kind: 'fab', score: 7, buf, fab: gen, pos: [...new Set(gen.inv.items.map((x) => x.po).filter(Boolean))], gen: true };
      } else if (gen && gen.pkl && gen.pkl.groups.length) {
        gen.supplier = gen.supplier || supplierNameOfWb(wb);
        out = { kind: 'genpkl', score: 5, buf, gen, pos: [...new Set(gen.pkl.groups.map((x) => x.po).filter(Boolean))] };
      } else out = { kind: '', score: 0, buf };
    }
  } else if (/\.docx$/.test(name)) out = { kind: 'doc', score: 0, why: 'Word .docx chưa hỗ trợ' };
  else out = { kind: '', score: 0 };
  CACHE.set(file, out);
  return out;
}

/* Packing list Excel theo PO (Inkava): No. | Material Code | Description | Supp. Ref. |
   Order No | Reference | [Story] | Size | Spec. | Quantity | Quantity | Thực xuất | …
   Tiêu đề có thể bị dính số 0 ("0 Size", "0 Spec.", "0 Order No") khi chủ hàng sửa tay. */
const PKLX_PO_RE = /^[A-Z]{3}\d{7}$/;   // DUY0081000
function pklxHeaderRow(ws) {
  const txt = (r, c) => String(ws.getCell(r, c).text || '').trim();
  for (let r = 1; r <= Math.min(12, ws.rowCount); r++) {
    let mat = false, qty = false, extra = 0;
    for (let c = 1; c <= 25; c++) {
      const t = txt(r, c);
      if (/^(0\s*)?Material\s*Code$/i.test(t)) mat = true;
      else if (/^Quantity$/i.test(t)) qty = true;
      else if (/^(0\s*)?(Spec\.?|Order\s*No|Supp\.?\s*Ref\.?|Size)$/i.test(t)) extra++;
    }
    if (mat && qty && extra >= 2) return r;
  }
  return 0;
}
function readPklx(ws, fname) {
  const txt = (r, c) => (c ? String(ws.getCell(r, c).text || '').trim() : '');
  const hr = pklxHeaderRow(ws);
  if (!hr) return null;
  let po = '', poFrom = '';
  for (let r = 1; r <= Math.min(10, ws.rowCount) && !po; r++) {
    for (let c = 1; c <= 20; c++) {
      const t = txt(r, c);
      const m = t.match(/PO\s*No\s*[:.]?\s*([A-Z0-9]{5,})/i);
      if (m) { po = m[1].toUpperCase(); poFrom = 'tiêu đề'; break; }
      if (r < hr && PKLX_PO_RE.test(t.toUpperCase())) { po = t.toUpperCase(); poFrom = 'ô đầu sheet'; break; }
    }
  }
  if (!po && fname) {
    const m = String(fname).toUpperCase().match(/(?:^|[^A-Z0-9])([A-Z]{3}\d{7})(?![0-9])/);
    if (m) { po = m[1]; poFrom = 'tên file'; }
  }
  const col = (re) => { for (let c = 1; c <= 25; c++) if (re.test(txt(hr, c))) return c; return 0; };
  const C = {
    mat: col(/Material\s*Code/i), desc: col(/Description/i), ref: col(/Reference/i),
    size: col(/^(0\s*)?Size\b/i), spec: col(/^(0\s*)?Spec/i), qty: col(/^Quantity$/i), order: col(/Order\s*No/i),
  };
  if (!C.mat || !C.qty) return null;
  const rows = [];
  let hidden = 0;
  for (let r = hr + 1; r <= ws.rowCount; r++) {
    if (/^total/i.test(txt(r, C.desc || 1)) || /^total/i.test(txt(r, 3)) || /^total/i.test(txt(r, 4))) break;
    /* dòng bị lọc ẩn: có khi là "không giao" (ô Total dùng SUBTOTAL), có khi chỉ là đang lọc để xem
       → giữ lại, đánh dấu hid để lúc đối chiếu thử cả hai cách tính */
    const hid = !!ws.getRow(r).hidden;
    const mat = txt(r, C.mat);
    const q = num(ws.getCell(r, C.qty).value);
    if (!mat || !/^[A-Z]{4,}[A-Z0-9]*\d{3,}$/i.test(mat.replace(/\s+/g, '')) || isNaN(q) || q <= 0) continue;
    rows.push({
      material: mat.replace(/\s+/g, '').toUpperCase(), ref: txt(r, C.ref), spec: txt(r, C.spec),
      size: txt(r, C.size).replace(/\s+/g, '').toUpperCase(), order: txt(r, C.order).trim(), qty: q, sheet: ws.name, hid,
    });
    if (hid) hidden++;
  }
  return { po, poFrom, rows, hidden, total: rows.reduce((a, b) => a + b.qty, 0) };
}
/* gộp mọi sheet packing list trong một file; mỗi dòng nhớ tên sheet để chọn đúng sheet khi đối chiếu */
function readPklxWb(wb, fname) {
  const parts = [];
  for (const w of wb.worksheets) {
    let px = null;
    try { px = readPklx(w, fname); } catch (e) { px = null; }
    if (px && px.rows.length) parts.push(px);
  }
  if (!parts.length) return null;
  const po = (parts.find((x) => x.poFrom === 'tiêu đề') || parts.find((x) => x.po) || {}).po || '';
  const rows = [].concat(...parts.map((x) => x.rows));
  return {
    po, rows, total: rows.reduce((a, b) => a + b.qty, 0),
    sheets: parts.map((x) => ({ name: x.rows[0].sheet, total: x.total, hidden: x.hidden })),
    poFrom: (parts.find((x) => x.po === po) || {}).poFrom || '',
  };
}

/* đọc cả thư mục khi kéo–thả */
async function filesFromDataTransfer(dt) {
  const out = [];
  const items = dt.items ? [...dt.items] : [];
  const entries = items.map((i) => (i.webkitGetAsEntry ? i.webkitGetAsEntry() : null)).filter(Boolean);
  if (!entries.length) return [...dt.files];
  const walk = async (entry, path) => {
    if (entry.isFile) {
      const f = await new Promise((res, rej) => entry.file(res, rej));
      try { Object.defineProperty(f, '_relPath', { value: path + f.name, configurable: true }); } catch (e) { /* bỏ qua */ }
      out.push(f);
    } else if (entry.isDirectory) {
      const reader = entry.createReader();
      let batch;
      do {
        batch = await new Promise((res, rej) => reader.readEntries(res, rej));
        for (const e of batch) await walk(e, path + entry.name + '/');
      } while (batch.length);
    }
  };
  for (const e of entries) await walk(e, '');
  return out;
}

/* File thả vào trong lúc đang nhận diện / đang chạy lượt trước: trước đây bị bỏ qua mà không báo
   (thả HĐ GTGT PDF rồi thả ngay file PKL + inbound → chỉ còn PDF). Nay xếp hàng, chạy xong lượt trước thì tự thêm. */
const FILE_QUEUE = [];
let ACCEPTING = false;
function flushQueue() {
  if (!FILE_QUEUE.length || STATE.busy || ACCEPTING) return;
  const q = FILE_QUEUE.splice(0);
  setTimeout(() => acceptFiles(q), 0);
}
async function acceptFiles(fileList) {
  const list = [...fileList].filter(Boolean);
  if (!list.length) return;
  if (STATE.busy || ACCEPTING) {
    FILE_QUEUE.push(...list);
    $('#detect').textContent = `Đang xử lý lượt trước — ${FILE_QUEUE.length} file vừa thả sẽ được thêm ngay sau khi xong, không cần thả lại.`;
    return;
  }
  ACCEPTING = true;
  try { await acceptFilesNow(list); } finally { ACCEPTING = false; }
  flushQueue();
}
async function acceptFilesNow(list) {
  $('#detect').textContent = `Đang nhận diện ${list.length} file…`;
  const unknown = [], fresh = new Set(), notices = [];
  for (const f of list) {
    let c; try { c = await classify(f); } catch (e) { c = { kind: '', score: 0 }; }
    if (c.kind === 'scan') { notices.push(`${f.name}: PDF dạng ảnh (scan)${c.why ? ' — ' + c.why : ''} — cần file Excel/PDF gốc hoặc nhập tay`); continue; }
    if (c.kind === 'xls') { notices.push(`${f.name}: Excel 97-2003 (.xls) không chuyển được (${c.why || ''}) — mở bằng Excel rồi Lưu dưới dạng .xlsx`); continue; }
    if (XLS_CONV.has(f)) notices.push(`${f.name}: Excel 97-2003 — đã tự chuyển sang .xlsx để đọc`);
    if (DOC_CONV.has(f) && c.kind && c.kind !== 'doc') notices.push(`${f.name}: Word 97-2003 — đã đọc chữ trong file`);
    if (c.kind === 'doc') { notices.push(`${f.name}: file Word — không đọc được${c.why ? ' (' + c.why + ')' : ''}, cần Excel/PDF`); continue; }
    if (c.kind === 'proforma') { notices.push(`${f.name}: proforma invoice — bỏ qua`); continue; }
    if (!c.kind) { unknown.push(f.name); continue; }
    if (c.kind === 'po') { STATE.po = f; continue; }
    const key = f.name + '|' + dirOf(f) + '|' + f.size;
    if (STATE.items.some((x) => x.key === key)) continue;
    STATE.items.push({ file: f, kind: c.kind, dir: dirOf(f), key });
    fresh.add(key);
  }
  $('#detect').textContent = (unknown.length ? 'Không nhận diện được: ' + unknown.join(', ') : '') + (notices.length ? (unknown.length ? ' · ' : '') + notices.join(' · ') : '')
    + (unknown.length || notices.some((n) => /không đọc được|chưa hỗ trợ|scan/i.test(n)) ? ` · Nếu đây là chứng từ của chủ hàng mới, vui lòng liên hệ ${contactText()} để được bổ sung.` : '');
  await buildGroups();
  // chỉ chọn sẵn các hóa đơn vừa được bổ sung file (thêm inbound cho hóa đơn nào thì chạy hóa đơn đó)
  const touched = STATE.groups.filter((g) => [g.inv, g.pkl, g.inb].some((x) => x && fresh.has(x.key)));
  if (touched.length && touched.length < STATE.groups.length) {
    STATE.groups.forEach((g) => { g.sel = false; });
    touched.forEach((g) => { g.sel = true; });
    $('#detect').textContent = ($('#detect').textContent ? $('#detect').textContent + ' · ' : '')
      + `Đã chọn sẵn ${touched.length} hóa đơn vừa thêm file — tick thêm ở cột đầu nếu muốn xuất lại hóa đơn khác.`;
  }
  renderSlots();
  if (!STATE.busy) { ACCEPTING = false; return run(); }
}

/* bản sao để việc ghép (gắn PKL, điền số HĐ GTGT, tách dòng theo PKL) không sửa vào dữ liệu đã đọc trong CACHE —
   nếu không, ghép lại lần nữa (thả thêm file) sẽ cộng đôi nhóm lô / mang số HĐ của lần trước */
const cloneFab = (f) => f && Object.assign({}, f, {
  inv: f.inv && Object.assign({}, f.inv, { items: (f.inv.items || []).map((x) => Object.assign({}, x)) }),
  pkl: f.pkl && Object.assign({}, f.pkl, { groups: (f.pkl.groups || []).slice(), files: (f.pkl.files || []).slice() }),
});
/* Ghép hóa đơn ↔ packing list ↔ inbound */
async function buildGroups() {
  /* mỗi lần ghép (kể cả khi thả thêm file lượt sau) làm lại từ loại GỐC của từng file: lần ghép trước có thể đã
     đổi HĐ GTGT / packing list rời thành "chứng từ" (kind 'fab') — giữ lại thì lượt sau lỗi "reading 'inv'" */
  for (const x of STATE.items) {
    if (x.origKind) x.kind = x.origKind;
    for (const k of ['fab', 'inv', 'isFab', 'isGen', 'pdfInv', 'fabOwner', 'used', 'pklx', 'tag', 'sapPos', 'px', 'gen', 'note']) delete x[k];
  }
  const invs = STATE.items.filter((x) => x.kind === 'inv');
  const fabs = STATE.items.filter((x) => x.kind === 'fab');
  const pkls = STATE.items.filter((x) => x.kind === 'pkl');
  const pxs = STATE.items.filter((x) => x.kind === 'pklx');
  const inbs = STATE.items.filter((x) => x.kind === 'inb');
  const poIdx = STATE.po ? readPoWb((await classify(STATE.po)).wb) : null;

  for (const it of invs) {
    const c = await classify(it.file);
    it.inv = parseInvoice(c.lines);
    it.note = c.note;
    it.tag = (it.inv.no || '').replace(/^0+/, '');
    it.sapPos = [...new Set(it.inv.items.map((x) => resolvePo(x, poIdx ? { scax: poIdx.map, sap: poIdx.sap } : null).sap).filter(Boolean))];
  }
  for (const p of pkls) p.note = (await classify(p.file)).note;
  for (const p of pxs) p.px = (await classify(p.file)).px;
  /* packing list Excel: gắn theo số PO xuất hiện trên hóa đơn (một hóa đơn nhiều PO) */
  for (const it of invs) {
    const want = [...new Set(it.inv.items.map((x) => String(x.po || '').toUpperCase()).filter(Boolean))];
    it.pklx = pxs.filter((p) => p.px && (p.px.pos || (p.px.po ? [p.px.po] : [])).some((q) => want.some((w) => poSame(w, q))));
    it.pklx.forEach((p) => { p.used = true; });
  }

  const nameHas = (f, tag) => tag && (f.name + ' ' + (f.webkitRelativePath || f._relPath || '')).includes(tag);
  // ghép toàn cục: chấm điểm mọi cặp rồi gán từ cặp điểm cao nhất xuống,
  // để một file không bị hóa đơn đứng trước "giành" mất
  const assign = (cands, extra) => {
    cands.forEach((c) => { c.used = false; });
    const pairs = [];
    for (const inv of invs) {
      for (const c of cands) {
        const sameDir = inv.dir && c.dir && inv.dir === c.dir;
        const byName = nameHas(c.file, inv.tag);
        // đã tổ chức theo thư mục thì không ghép chéo thư mục (trừ khi tên file có số hóa đơn)
        if (inv.dir && c.dir && !sameDir && !byName) continue;
        const s = (sameDir ? 6 : 0) + (byName ? 3 : 0) + extra(inv, c);
        if (s > 0) pairs.push({ inv, c, s });
      }
    }
    pairs.sort((a, b) => b.s - a.s);
    const res = new Map();
    for (const p of pairs) {
      if (res.has(p.inv) || p.c.used) continue;
      res.set(p.inv, p.c); p.c.used = true;
    }
    return res;
  };
  const pklOf = assign(pkls, (a, b) => (a.note && b.note && a.note === b.note ? 4 : 0));
  /* một bộ hàng (Despatch Note) có thể được chia thành nhiều hóa đơn → cho dùng chung packing list */
  for (const it of invs) {
    if (pklOf.get(it) || !it.note) continue;
    const sh = pkls.find((p) => p.note && p.note === it.note);
    if (sh) pklOf.set(it, sh);
  }
  const inbOf = assign(inbs, (a, b) => {
    const c = CACHE.get(b.file);
    if (!c || !c.pos || !a.sapPos.length) return 0;
    const hit = a.sapPos.filter((p) => c.pos.includes(p)).length;
    return (hit / a.sapPos.length) * 3;
  });

  /* ---- chứng từ vải: hóa đơn + packing list nằm trong cùng 1 file; một file
       inbound có thể dùng cho nhiều hóa đơn (SAP xuất chung nhiều PO) ---- */
  for (const it of fabs) {
    const c = await classify(it.file);
    it.fab = cloneFab(c.fab); it.inv = it.fab.inv; it.isFab = true; it.pdfInv = null; it.isGen = !!c.gen;
    it.tag = it.inv.no || '';
    it.sapPos = [...new Set(it.inv.items.map((x) => x.po).filter(Boolean))];
    if (it.isGen && it.fab.pkl && it.fab.pkl.groups) it.fab.pkl = { groups: it.fab.pkl.groups.slice(), unit: it.fab.pkl.unit, level: 'lot', soft: !!it.fab.pkl.soft, files: [] };
  }
  /* Chủ hàng vải có thể gửi kèm HÓA ĐƠN GTGT dạng PDF (Yubo): lấy ký hiệu + số + ngày
     và tổng tiền chính thức từ đó, còn chi tiết dòng/lô vẫn đọc từ file PKL.
     Ghép toàn cục: số HĐ trong ô "HD:" trùng số trên PDF (+5) · cùng ngày (+3) ·
     mỗi PO trùng (+1) · cùng thư mục (+0.5). Phải có ít nhất một trong hai dấu hiệu
     đầu, vì nhiều file PKL của cùng chủ hàng đều trùng hết số PO.                      */
  const pairVat = () => {
    const digits = (x) => String(x == null ? '' : x).replace(/\D/g, '').replace(/^0+/, '');
    const pairs = [];
    for (const it of fabs) {
      if (!it.inv) continue;
      const hd = digits(it.inv.vatNo || it.inv.no);
      const gpkPos = STATE.items.filter((x) => x.kind === 'genpkl' && (x.dir || '') === (it.dir || '') && x.gen && x.gen.pkl).map((x) => x.gen.pkl.groups.map((g) => g.po)).flat();
      const posAll = [...new Set(it.sapPos.concat(it.fab && it.fab.pkl ? it.fab.pkl.groups.map((g) => g.po) : [], gpkPos).filter(Boolean))];
      for (const v of invs) {
        const cv = CACHE.get(v.file);
        if (!cv || !cv.lines || !v.inv || !v.inv.invNo) continue;
        const txt = cv.lines.join(' ').toUpperCase().replace(/\s+/g, '');
        const nPo = posAll.filter((q) => txt.includes(String(q).toUpperCase().replace(/\s+/g, ''))).length;
        const vn = digits(v.inv.no);
        const sameNo = hd && vn && (vn === hd || vn.endsWith(hd) || hd.endsWith(vn));
        const sameDay = it.inv.invDate && v.inv.invDate && it.inv.invDate === v.inv.invDate;
        /* chứng từ chung (Paddies: commercial invoice + HĐ GTGT scan): PO trùng cũng đủ để ghép khi HĐ GTGT không có dòng hàng */
        const cv2 = CACHE.get(v.file);
        const headerOnly = !!(cv2 && cv2.headerOnly);
        if (!sameNo && !sameDay && !(it.isGen && headerOnly && nPo > 0)) continue;
        const sc = (sameNo ? 5 : 0) + (sameDay ? 3 : 0) + nPo
          + (it.dir && v.dir && it.dir === v.dir ? 0.5 : 0);
        pairs.push({ it, v, sc });
      }
    }
    pairs.sort((a, b) => b.sc - a.sc);
    for (const pr of pairs) {
      if (pr.it.pdfInv || pr.v.fabOwner) continue;
      pr.v.fabOwner = pr.it; pr.it.pdfInv = pr.v;
      const pv = pr.v.inv;
      /* chứng từ đã tự ghi số HĐ GTGT ("VAT Invoice#") thì giữ, vì bản PDF OCR có thể đọc thiếu ký hiệu */
      if (pv.invNo && !pr.it.inv.vatNo) { pr.it.inv.invNo = pv.invNo; pr.it.inv.noInvoiceNo = false; pr.it.inv.noSerial = false; }
      if (pv.invDate) pr.it.inv.invDate = pv.invDate;
      pr.it.inv.pdfTotal = pv.total; pr.it.inv.pdfVat = pv.vat; pr.it.inv.pdfPayment = pv.payment;
      pr.it.inv.pdfFile = pr.v.file.name;
    }
    /* chỉ còn đúng MỘT hoá đơn GTGT và MỘT chứng từ chưa có số HĐ, lại trùng PO → ghép dù không trùng số/ngày
       (PKL Yubo mẫu mới không có ô "HD:", ngày ghi dạng khác) */
    const vLeft = invs.filter((v) => !v.fabOwner && v.inv && v.inv.invNo);
    const iLeft = fabs.filter((it) => !it.pdfInv && it.inv && !it.inv.invNo);
    if (vLeft.length === 1 && iLeft.length === 1) {
      const v = vLeft[0], it = iLeft[0], cv = CACHE.get(v.file);
      const txt = cv && cv.lines ? cv.lines.join(' ').toUpperCase().replace(/\s+/g, '') : '';
      const pos = [...new Set(it.sapPos.concat(it.fab && it.fab.pkl ? it.fab.pkl.groups.map((g) => g.po) : []).filter(Boolean))];
      if (pos.some((q) => txt.includes(String(q).toUpperCase().replace(/\s+/g, '')))) {
        v.fabOwner = it; it.pdfInv = v; const pv = v.inv;
        it.inv.invNo = pv.invNo; it.inv.noInvoiceNo = false; it.inv.noSerial = false;
        if (pv.invDate) it.inv.invDate = pv.invDate;
        it.inv.pdfTotal = pv.total; it.inv.pdfVat = pv.vat; it.inv.pdfPayment = pv.payment; it.inv.pdfFile = v.file.name;
      }
    }
  };
  pairVat();

  /* Hoá đơn GTGT không theo mẫu ITL/Inkava và không ghép được với PKL vải nào → tự nó là chứng từ (bộ đọc chung) */
  for (const it of invs.slice()) {
    if (it.fabOwner || it.inv.items.length) continue;
    const c = CACHE.get(it.file);
    if (!c || !c.genFallback) continue;
    it.fab = cloneFab(c.genFallback); it.inv = it.fab.inv; it.isFab = true; it.isGen = true; it.pdfInv = null; it.origKind = it.origKind || it.kind; it.kind = 'fab';
    it.tag = it.inv.no || '';
    it.sapPos = [...new Set(it.inv.items.map((x) => x.po).filter(Boolean))];
    if (it.fab.pkl && it.fab.pkl.groups) it.fab.pkl = { groups: it.fab.pkl.groups.slice(), unit: it.fab.pkl.unit, level: 'lot', soft: !!it.fab.pkl.soft, files: [] };
    fabs.push(it); invs.splice(invs.indexOf(it), 1);
    const ib = inbOf.get(it); if (ib) { ib.used = false; inbOf.delete(it); }
    const pk = pklOf.get(it); if (pk) { pk.used = false; pklOf.delete(it); }
  }
  /* HĐ GTGT chỉ có phần đầu (scan, OCR) có thể thuộc về chứng từ vừa chuyển ở trên (Paddies: CI + HĐ GTGT scan) → ghép lại lần nữa */
  pairVat();
  /* ---- packing list rời (Excel/PDF) của chủ hàng dùng bộ đọc chung: gắn vào chứng từ cùng thư mục,
       hoặc có số hoá đơn / PO trùng. Một hoá đơn có thể có nhiều file packing list (CELEB PKL (1)(2)(3)). ---- */
  const gpks = STATE.items.filter((x) => x.kind === 'genpkl');
  for (const p of gpks) { p.gen = cloneFab((await classify(p.file)).gen); p.used = false; }
  const genFabs = fabs.filter((x) => x.isGen);
  for (const p of gpks) {
    let best = null, bs = 0;
    const ptxt = (p.file.name + ' ' + (p.gen.inv.invNo || '')).toUpperCase().replace(/\s+/g, '');
    /* packing list của mẫu riêng (Carvico) chỉ gắn vào chứng từ cùng mẫu; còn lại gắn vào chứng từ bộ đọc chung */
    const own = p.gen.profile && p.gen.profile !== 'GEN';
    const cands = own ? fabs.filter((x) => x.fab && x.fab.profile === p.gen.profile) : genFabs;
    for (const it of cands) {
      /* số packing list ghi trên hóa đơn (Carvico "00851853 …") trùng số trên file PKL → gắn chắc */
      const byPkl = p.gen.pklNo && it.inv.items.some((x) => x.pklNo && stripZero(x.pklNo) === stripZero(p.gen.pklNo));
      const sameDir = it.dir && p.dir && it.dir === p.dir;
      if (it.dir && p.dir && !sameDir) continue;
      const tag = String(it.tag || '').toUpperCase().replace(/[\s#]+/g, '').replace(/^0+/, '');
      const byNo = tag.length >= 4 && ptxt.includes(tag);
      const pos = (p.gen.pkl ? p.gen.pkl.groups.map((g) => g.po) : []).filter(Boolean);
      const nPo = pos.filter((q) => it.sapPos.some((x) => poSame(x, q))).length;
      const sameSup = it.fab.supplier && p.gen.supplier && AZ(it.fab.supplier).slice(0, 8) === AZ(p.gen.supplier).slice(0, 8);
      /* mã hàng / màu trên packing list có trong hoá đơn */
      const invTxt = AZ(it.inv.items.map((x) => [x.article, x.colorText, x.desc].join(' ')).join(' '));
      const toks = (p.gen.pkl ? p.gen.pkl.groups : []).map((g) => [g.article, g.color]).flat().map(AZ).filter((k) => k.length >= 5);
      const sameArt = toks.length && toks.filter((k) => invTxt.includes(k)).length >= Math.min(2, toks.length);
      const only = genFabs.length === 1 && gpks.length === 1;
      const sc = (sameDir ? 6 : 0) + (byNo ? 3 : 0) + nPo + (sameSup ? 1 : 0) + (sameArt ? 2 : 0) + (only ? 1 : 0) + (byPkl ? 8 : 0);
      if (sc > bs) { bs = sc; best = it; }
    }
    if (best && bs > 0 && p.gen.pkl) {
      if (!best.fab.pkl) best.fab.pkl = { groups: [], unit: p.gen.pkl.unit, level: 'lot', soft: !!p.gen.pkl.soft, files: [] };
      best.fab.pkl.groups = best.fab.pkl.groups.concat(p.gen.pkl.groups);
      best.fab.pkl.soft = best.fab.pkl.soft || !!p.gen.pkl.soft;
      best.fab.pkl.files.push(p.file.name);
      p.used = true;
      /* hoá đơn ghi gộp theo mã hàng, packing list rời ghi theo PO/size → tách lại theo packing list (chỉ bộ đọc chung) */
      if (best.isGen) try { const re = genFromDocsRefresh(best.fab); if (re) best.inv = best.fab.inv = re; } catch (e) { /* bỏ qua */ }
      best.sapPos = [...new Set(best.inv.items.map((x) => x.po).filter(Boolean))];
    }
  }
  /* packing list rời không có hoá đơn (hoá đơn là bản scan): dùng chính packing list làm chứng từ */
  for (const p of gpks) {
    if (p.used || !p.gen || !p.gen.inv || !p.gen.inv.items.length) continue;
    p.fab = cloneFab(p.gen); p.inv = p.fab.inv; p.isFab = true; p.isGen = true; p.pdfInv = null; p.origKind = p.origKind || p.kind; p.kind = 'fab';
    p.inv.noInvoiceNo = !p.inv.invNo;
    p.tag = p.inv.no || '';
    p.sapPos = [...new Set(p.inv.items.map((x) => x.po).filter(Boolean))];
    p.used = true; fabs.push(p); genFabs.push(p);
  }
  /* Bản scan OCR trùng với một chứng từ có chữ (SunPo gửi cả CI scan lẫn Excel): cùng thư mục,
     cùng số hoá đơn hoặc cùng tổng số lượng → bỏ bản OCR, giữ bản có chữ. */
  STATE.dupOcr = [];
  {
    const sumQ = (f) => Math.round(f.inv.items.reduce((a, x) => a + (isNaN(x.qty) ? 0 : x.qty), 0) * 100) / 100;
    const dg = (x) => String(x == null ? '' : x).replace(/\D/g, '').replace(/^0+/, '');
    for (const o of fabs.slice()) {
      if (!o.fab || !o.fab.ocr) continue;
      const twin = fabs.find((t) => t !== o && t.fab && !t.fab.ocr && (t.dir || '') === (o.dir || '')
        && ((dg(t.inv.invNo) && dg(t.inv.invNo) === dg(o.inv.invNo)) || (sumQ(t) > 0 && Math.abs(sumQ(t) - sumQ(o)) <= sumQ(t) * 0.005)));
      if (!twin) continue;
      /* bản có chữ chỉ là packing list (số trên đó là số packing) → lấy số hoá đơn từ bản scan */
      if ((!twin.inv.invNo || twin.fab.pklOnly) && o.inv.invNo) { twin.inv.invNo = twin.inv.no = o.inv.invNo; twin.inv.noInvoiceNo = false; twin.inv.noFrom = 'bản scan ' + o.file.name + ' (OCR)'; }
      if (!twin.inv.invDate && o.inv.invDate) twin.inv.invDate = o.inv.invDate;
      STATE.dupOcr.push({ o, twin });
      fabs.splice(fabs.indexOf(o), 1);
      const gi = genFabs.indexOf(o); if (gi >= 0) genFabs.splice(gi, 1);
    }
  }

  /* tên file SAP xuất ra có dấu thời gian: ZMME0032_20260926042733 → 26.09.2026.
     Thả nhiều bản xuất thì chọn bản gần ngày hóa đơn nhất.                      */
  const fileDay = (f) => {
    const m = String(f.name).match(/(20\d{2})(\d{2})(\d{2})/);
    return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) : NaN;
  };
  const invDay = (d) => {
    const m = String(d || '').match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
    return m ? Date.UTC(+m[3], +m[2] - 1, +m[1]) : NaN;
  };
  const inbFab = new Map();
  for (const it of fabs) {
    let best = null, bs = 0;
    const di = invDay(it.inv.invDate);
    for (const b of inbs) {
      const c = CACHE.get(b.file);
      if (!c || !c.pos || !c.pos.length) continue;
      const hit = it.sapPos.filter((q) => c.pos.some((x) => poSame(x, q))).length;
      const db = fileDay(b.file);
      const gap = (!isNaN(di) && !isNaN(db)) ? Math.abs(di - db) / 86400000 : NaN;
      const near = isNaN(gap) ? 0 : (gap <= 2 ? 2.5 : (gap <= 7 ? 1 : 0));
      /* chủ hàng trên chứng từ trùng Partner Name trong inbound (chứng từ không ghi PO: CARVICO) */
      const supKey = AZ(it.fab.supplier || '').slice(0, 6);
      const supHit = supKey.length >= 4 && (c.partners || []).some((x) => AZ(x).startsWith(supKey));
      const only = inbs.length === 1 && fabs.length === 1;
      const s = (hit / Math.max(1, it.sapPos.length)) * 6
        + (it.dir && b.dir && it.dir === b.dir ? 2 : 0) + (nameHas(b.file, it.tag) ? 3 : 0) + near + (supHit ? 2 : 0) + (only ? 0.5 : 0);
      if (s > bs) { bs = s; best = b; }
    }
    if (best && bs > 0) { inbFab.set(it, best); best.used = true; }
  }

  const prevSel = new Map((STATE.groups || []).map((g) => [g.inv.key, g.sel !== false]));
  STATE.groups = [];
  STATE.headerOnly = [];
  for (const it of invs) {
    if (it.fabOwner) continue;            // đã dùng làm hóa đơn cho chứng từ vải
    const cc = CACHE.get(it.file);
    if (cc && cc.headerOnly && !it.inv.items.length) { STATE.headerOnly.push(it); continue; }   // HĐ GTGT chỉ đọc được số/ngày
    STATE.groups.push({
      inv: it, pkl: pklOf.get(it) || null, inb: inbOf.get(it) || null, pklx: it.pklx || [],
      sel: prevSel.has(it.key) ? prevSel.get(it.key) : true,
    });
  }
  for (const it of fabs) {
    STATE.groups.push({
      inv: it, pkl: null, inb: inbFab.get(it) || null, isFab: true, isGen: !!it.isGen, pdfInv: it.pdfInv || null,
      fabName: (it.fab.ocr ? 'OCR · ' : '') + (it.isGen ? 'bộ đọc chung · ' : '') + (it.fab.supplier || it.fab.profile) + (it.fab.pkl && it.fab.pkl.files && it.fab.pkl.files.length ? ' · PKL: ' + it.fab.pkl.files.join(', ') : ''),
      sel: prevSel.has(it.key) ? prevSel.get(it.key) : true,
    });
  }
  /* chủ hàng của từng bộ — hiện ngay trên bảng file, dùng lại khi chạy */
  for (const g of STATE.groups) {
    const sup = supplierCheck(g, g.inb ? CACHE.get(g.inb.file) : null);
    g.supName = sup.name; g.supTrained = sup.trained; g.supKind = sup.kind;
  }
  STATE.orphanPkl = [...pkls, ...pxs, ...gpks.filter((p) => p.kind === 'genpkl')].filter((p) => !p.used);
  STATE.orphanInb = inbs.filter((p) => !p.used);
  STATE.poIdx = poIdx;
}

function selectedGroups() { return STATE.groups.filter((g) => g.sel); }

function renderSlots() {
  const g = STATE.groups;
  const n = selectedGroups().length;
  const nFab = g.filter((x) => x.isFab).length;
  const nGen = g.filter((x) => x.isGen).length;
  const nNew = g.filter((x) => !x.supTrained).length;
  $('#count').textContent = `${g.length} hóa đơn (chọn ${n})${nFab - nGen ? ` · ${nFab - nGen} chứng từ vải` : ''}${nGen ? ` · ${nGen} chứng từ (bộ đọc chung)` : ''}${nNew ? ` · ${nNew} chủ hàng chưa huấn luyện` : ''} · ${g.filter((x) => x.pkl || x.isFab || (x.pklx && x.pklx.length)).length} packing list · ${g.filter((x) => x.inb).length} inbound${STATE.po ? ' · có file PO' : ' · chưa có file PO (chỉ cần cho trimming / PO hệ cũ)'}`;
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  if (!g.length) { $('#groups').innerHTML = ''; return; }
  let h = `<div class="selbar">Chọn hóa đơn để xuất:
      <a href="#" data-sel="all">tất cả</a> ·
      <a href="#" data-sel="none">bỏ chọn</a> ·
      <a href="#" data-sel="inb">chỉ hóa đơn đã có inbound</a></div>`;
  h += '<table class="files"><thead><tr><th class="c"><input type="checkbox" id="selAll"></th><th>Hóa đơn</th><th>Chủ hàng</th><th>Ngày</th><th class="n">Dòng</th><th>Packing list</th><th>File inbound</th></tr></thead><tbody>';
  g.forEach((x, i) => {
    h += `<tr class="${x.sel ? '' : 'off'}"><td class="c"><input type="checkbox" class="gsel" data-i="${i}"${x.sel ? ' checked' : ''}></td>
      <td><b>${esc(x.inv.inv.invNo || x.inv.file.name)}</b><span class="fn">${esc(x.inv.file.name)}</span></td>
      <td>${x.supTrained ? `<b>${esc(x.supName || x.supTrained)}</b><span class="fn">${esc(x.supKind || '')}</span>` : `<b>${esc(x.supName || '?')}</b><span class="fn"><span class="tag bad">chưa huấn luyện</span></span>`}</td>
      <td>${esc(x.inv.inv.invDate)}</td><td class="n">${x.inv.inv.items.length}</td>
      <td>${x.isFab ? `<span class="${(x.inv.fab && x.inv.fab.pkl) ? 'ok2' : 'miss'}">${(x.inv.fab && x.inv.fab.pkl) ? '✓' : '–'}</span> <span class="fn">${(x.inv.fab && x.inv.fab.pkl && !(x.inv.fab.pkl.files && x.inv.fab.pkl.files.length)) ? 'trong cùng file · ' : ((x.inv.fab && x.inv.fab.pkl) ? '' : 'không có packing list · ')}${esc(x.fabName || 'vải')}${x.pdfInv ? ' · kèm HĐ GTGT ' + esc(x.pdfInv.file.name) : ''}</span>`
        : ((x.pklx && x.pklx.length) ? `<span class="ok2">✓</span> <span class="fn">${x.pklx.length} file ${x.pklx.some((p) => p.px.ttg) ? '' : 'Excel '}theo PO: ${esc([...new Set([].concat(...x.pklx.map((p) => p.px.pos || [p.px.po])))].join(', '))}</span>`
          : (x.pkl ? `<span class="ok2">✓</span> <span class="fn">${esc(x.pkl.file.name)}</span>` : '<span class="miss">chưa có</span>'))}</td>
      <td>${x.inb ? `<span class="ok2">✓</span> <span class="fn">${esc(x.inb.file.name)}</span>` : '<span class="miss">chưa có</span>'}</td></tr>`;
  });
  h += '</tbody></table>';
  if (STATE.orphanPkl.length || STATE.orphanInb.length) {
    h += `<div class="hint">Không ghép được với hóa đơn nào: ${[...STATE.orphanPkl, ...STATE.orphanInb].map((x) => esc(x.file.name)).join(', ')}</div>`;
  }
  if (STATE.dupOcr && STATE.dupOcr.length) {
    h += `<div class="hint">Bản scan (OCR) trùng với chứng từ có chữ nên không dùng: ${STATE.dupOcr.map((d) => esc(d.o.file.name) + ' → dùng ' + esc(d.twin.file.name)).join(', ')}</div>`;
  }
  if (STATE.headerOnly && STATE.headerOnly.length) {
    h += `<div class="hint">Hóa đơn GTGT chỉ đọc được số/ngày, không đọc được dòng hàng (bản scan mờ?) và không ghép được với chứng từ nào: ${STATE.headerOnly.map((x) => esc(x.file.name) + ' (' + esc(x.inv.invNo) + ')').join(', ')}</div>`;
  }
  $('#groups').innerHTML = h;
  $('#selAll').checked = n === g.length;
  const selKey = selectedGroups().map((x) => x.inv.key).join('|');
  $('#stale').textContent = (OUTPUTS.length && STATE.lastRunSel !== undefined && STATE.lastRunSel !== selKey)
    ? '⚠ Danh sách chọn đã thay đổi — bấm "Đối chiếu & xuất file" để cập nhật các file bên dưới.' : '';
  $('#groups').querySelectorAll('.gsel').forEach((cb) => {
    cb.addEventListener('change', (e) => { STATE.groups[+e.target.dataset.i].sel = e.target.checked; renderSlots(); });
  });
  $('#selAll').addEventListener('change', (e) => { STATE.groups.forEach((x) => { x.sel = e.target.checked; }); renderSlots(); });
  $('#groups').querySelectorAll('[data-sel]').forEach((a) => {
    a.addEventListener('click', (e) => {
      e.preventDefault();
      const m = e.target.dataset.sel;
      STATE.groups.forEach((x) => { x.sel = m === 'all' ? true : m === 'none' ? false : !!x.inb; });
      renderSlots();
    });
  });
}

/* Tìm các dòng inbound khớp mã hàng: thử lần lượt các biến thể của mã (chi tiết → chung),
   rồi lọc thêm bằng từ khóa phụ có trong mã (Main / Care / Angel Pink…) */
function matchByCode(R, pool) {
  const textOf = (x) => norm(x.desc) + '|' + norm(x.spec) + '|' + norm(x.material);
  let hit = [], usedVar = '';
  for (const v of R.vars) {
    const h = pool.filter((x) => textOf(x).includes(v));
    if (h.length) { hit = h; usedVar = v; break; }
  }
  R.usedVar = usedVar;
  R.narrowedByWord = false;
  for (const w of R.words.filter((w) => !usedVar.includes(w))) {
    const h = hit.filter((x) => textOf(x).includes(w));
    if (h.length && h.length < hit.length) { hit = h; R.narrowedByWord = true; }
  }
  // các dòng khác cùng "họ" mã hàng (khớp bất kỳ biến thể nào) nhưng không được chọn
  R.broadHit = pool.filter((x) => R.vars.some((v) => textOf(x).includes(v)));
  R.alt = R.broadHit.filter((x) => !hit.includes(x));
  return hit;
}

/* Điểm "quyền sở hữu" một dòng inbound khi nhiều dòng hóa đơn cùng giành:
   biến thể mã càng chi tiết càng cao · có từ khóa của mình +2 · mang từ khóa của item khác −6 */
function claimScore(row, R) {
  const t = norm(row.desc) + '|' + norm(row.spec) + '|' + norm(row.material);
  const vi = R.vars.findIndex((v) => t.includes(v));
  if (vi < 0) return -Infinity;
  let s = 10 - Math.min(vi, 8);
  for (const w of R.words) if (t.includes(w)) s += 2;
  for (const w of (R.rivalWords || [])) if (t.includes(w)) s -= 6;
  return s;
}

/* ================= Phân tích 1 hóa đơn ================= */
function analyze(inv, pkl, rows, po, pklx) {
  const lines = [];
  const resolved = inv.items.map((it) => ({ it, rp: resolvePo(it, po ? { scax: po.map, sap: po.sap } : null) }));
  const priceKey = {};   // để chỉ ghép theo đơn giá khi (PO + đơn giá) là duy nhất trong hóa đơn
  resolved.forEach(({ it, rp }) => { const k = (rp.sap || '') + '|' + it.price; priceKey[k] = (priceKey[k] || 0) + 1; });

  /* ---- Vòng 1: mỗi dòng hóa đơn tự tìm các dòng inbound khớp mã hàng ---- */
  const rowsOf = (sapPo) => (rows && sapPo ? rows.filter((x) => x.poV.toUpperCase() === sapPo.toUpperCase()) : []);
  const textOf = (x) => norm(x.desc) + '|' + norm(x.spec) + '|' + norm(x.material);
  const wordsOf = (code) => [...new Set((String(code).toUpperCase().match(/[A-Z]{3,}/g) || []))];

  /* không tra được qua file PO nhưng chính file inbound có mã PO đó → dùng luôn */
  if (rows) {
    const inbPos = [...new Set(rows.map((x) => x.poV.trim().toUpperCase()))];
    for (const R of resolved) {
      /* đã tra được nhưng mã đó không có trong file inbound → thử lại theo cột A của inbound */
      if (R.rp.sap && inbPos.includes(R.rp.sap.trim().toUpperCase())) continue;
      const cands = [...new Set(((R.it.text || '') + ' ' + (R.it.po || '')).toUpperCase().match(/[A-Z]{2,6}\d{4,}/g) || [])];
      const hitPo = inbPos.find((p) => cands.some((c) => c === p || c.endsWith(p) || p.endsWith(c)));
      if (hitPo) { R.rp = { scax: '', sap: hitPo, via: 'inbound' }; }
    }
  }

  for (const R of resolved) {
    const { it, rp } = R;
    it.poScax = rp.scax; it.poVia = rp.via;
    it.po = rp.scax || rp.sap || it.po;
    R.sapPo = rp.sap;
    R.vars = it.exactVars || codeVariants(it.code);
    R.words = wordsOf(it.code);
    R.inPo = rowsOf(rp.sap);
    R.matchBy = 'mã hàng';
    R.hit = matchByCode(R, R.inPo);
    /* Thiên Gia: cùng PO có thể có hai mặt hàng trùng kích thước (L100xW35 giá 270 và 365) → tách bằng đơn giá */
    if (it.priceNarrow && R.hit.length > 1 && !isNaN(it.price)) {
      const hp = R.hit.filter((x) => x.price === it.price);
      if (hp.length && hp.length < R.hit.length) { R.hit = hp; R.narrowedByWord = true; R.alt = []; R.byPrice = true; }
    }
    // mã hàng trên hóa đơn không đủ để phân biệt: inbound còn dòng khác cùng họ mã mà không có
    // từ khóa nào trong mã để tách ra (vd hóa đơn ghi "LB 5731 C/509" trong khi inbound có cả
    // "Main label LB 5731 C/509" lẫn "LB care label 5731 C/509")
    R.ambiguous = !R.narrowedByWord && R.hit.length > 0 && R.alt.length > 0;
    // còn mơ hồ → nếu chỉ đúng MỘT nhóm có tổng số lượng bằng số lượng hóa đơn thì chọn nhóm đó
    if (R.ambiguous && it.qty) {
      const famKey = (x) => norm(x.desc).replace(/[0-9]/g, '').slice(0, 26);
      const eff = (arr) => {
        const v = arr.reduce((a, x) => a + (isNaN(x.invQty) ? 0 : x.invQty), 0);
        return v > 0 ? v : arr.reduce((a, x) => a + (isNaN(x.qty) ? 0 : x.qty), 0);
      };
      const fams = {};
      R.alt.forEach((x) => { const k = famKey(x); (fams[k] = fams[k] || []).push(x); });
      const cands = [R.hit, ...Object.values(fams)];
      const exact = cands.filter((c) => eff(c) === it.qty);
      if (exact.length === 1 && exact[0] !== R.hit) {
        R.hit = exact[0];
        R.alt = R.broadHit.filter((x) => !R.hit.includes(x));
        R.matchBy = 'mã hàng + số lượng khớp';
        R.ambiguous = false; R.pickedByQty = true;
      } else if (exact.length === 1) { R.ambiguous = false; R.pickedByQty = true; }
    }
  }

  /* ---- Vòng 2: phân xử khi nhiều dòng hóa đơn cùng giành một dòng inbound ----
     Ví dụ "LB 5731 Main C/509" và "LB 5731 C/509" cùng PO: dòng inbound
     "Main label LB 5731 C/509" thuộc về dòng có chữ Main; dòng còn lại đi tìm tiếp.  */
  const byPo = {};
  resolved.forEach((R) => { if (R.sapPo) (byPo[R.sapPo] = byPo[R.sapPo] || []).push(R); });
  for (const list of Object.values(byPo)) {
    if (list.length < 2) continue;
    list.forEach((R) => {
      R.rivalWords = [...new Set([].concat(...list.filter((o) => o !== R).map((o) => o.words)))]
        .filter((w) => !R.words.includes(w));
    });
    const owners = new Map();     // row -> [R…]
    list.forEach((R) => R.hit.forEach((x) => owners.set(x, (owners.get(x) || []).concat(R))));
    const tied = [];
    for (const [row, claimers] of owners) {
      if (claimers.length < 2) continue;
      const scored = claimers.map((R) => ({ R, s: claimScore(row, R) }));
      const best = Math.max(...scored.map((x) => x.s));
      const win = scored.filter((x) => x.s === best);
      if (win.length === 1) {
        scored.filter((x) => x.s !== best).forEach((x) => { x.R.hit = x.R.hit.filter((y) => y !== row); });
        win[0].R.ambiguous = false;                 // đã tách được nhờ item trùng mã bên cạnh
        scored.filter((x) => x.s !== best).forEach((x) => { x.R.resolvedAway = true; });
      } else tied.push(row);
    }
    // dòng hóa đơn bị mất hết dòng inbound → tìm lại trong phần chưa bị ai chiếm
    const owned = new Set();
    list.forEach((R) => R.hit.forEach((x) => owned.add(x)));
    for (const R of list) {
      if (R.hit.length) continue;
      const free = R.inPo.filter((x) => !owned.has(x));
      R.hit = matchByCode(R, free);
      R.hit.forEach((x) => owned.add(x));
      if (R.hit.length) { R.matchBy = 'mã hàng (sau khi tách với item trùng mã)'; R.ambiguous = false; }
    }
    list.forEach((R) => { if (R.resolvedAway && R.hit.length) R.ambiguous = false; });
    if (tied.length) list.forEach((R) => { if (R.hit.some((x) => tied.includes(x))) R.ambiguous = true; });
  }

  for (const R of resolved) {
    const { it, rp } = R;
    const sapPo = R.sapPo, vars = R.vars, inPo = R.inPo;
    const codeN = norm(it.code);
    let hit = R.hit, matchBy = R.matchBy;
    // không tìm được mã hàng trong file inbound → thử ghép theo đơn giá (mã nội bộ của ITL khác mã SAP)
    if (rows && sapPo && !hit.length && it.price && priceKey[(sapPo || '') + '|' + it.price] === 1) {
      const h = inPo.filter((x) => x.price === it.price);
      if (h.length) { hit = h; matchBy = 'đơn giá'; }
    }
    /* ---- packing list Excel theo PO (Inkava): có sẵn Material Code + Size + Spec
       nên phân bổ được số lượng vào từng dòng inbound, không cần điền tay ---- */
    let xRows = null, xNote = '', xOrder = '', ttgDone = false, ttgCov = 0, ttgMiss = 0;
    const ttgFiles = (pklx || []).filter((x) => x.ttg);
    if (ttgFiles.length && sapPo && it.dim) {
      /* packing list Thiên Gia: dòng cùng PO + kích thước; chia cho từng dòng inbound theo mã code trong Specification */
      const rws = [].concat(...ttgFiles.map((x) => x.rows)).filter((rw) => poSame(rw.po, sapPo) && rw.dim === it.dim);
      if (rws.length) {
        ttgDone = true;
        const used = [];
        hit.forEach((h) => {
          const cs = ttgCodes(h.spec + ' ' + h.desc);
          const mine = cs.length ? rws.filter((rw) => rw.code && cs.includes(rw.code) && !used.includes(rw)) : [];
          if (mine.length) { mine.forEach((rw) => { used.push(rw); rw.size = h.size || '?'; }); h.setQty = mine.reduce((a, b) => a + b.qty, 0); }
          else h.setQty = null;
        });
        /* dòng không ghi mã code (sticker, hanger) → chỉ gán khi còn đúng một dòng inbound chưa có số */
        const free = rws.filter((rw) => !used.includes(rw) && !rw.code);
        const open = hit.filter((h) => h.setQty == null);
        if (free.length && open.length === 1) { free.forEach((rw) => { used.push(rw); rw.size = open[0].size || '?'; }); open[0].setQty = free.reduce((a, b) => a + b.qty, 0); }
        const extra = rws.filter((rw) => !used.includes(rw));
        ttgCov = hit.reduce((a, h) => a + (h.setQty || 0), 0);
        ttgMiss = hit.filter((h) => h.setQty == null).length;
        xRows = used.length ? used : null;
        const byCode = used.filter((rw) => rw.code).length;
        if (extra.length && byCode) xNote += `Packing list còn ${extra.length} dòng cùng PO + kích thước không thuộc dòng hóa đơn này (${extra.map((rw) => (rw.code || rw.color || '?') + ': ' + fmt(rw.qty)).join('; ')}). `;
        if (ttgMiss && ttgCov > 0) xNote += `${ttgMiss} dòng inbound không có trong packing list đợt này. `;
      }
    }
    if (!ttgDone && pklx && pklx.length && sapPo) {
      const files = pklx.filter((x) => x.po && (x.po.toUpperCase() === sapPo.toUpperCase()
        || x.po.toUpperCase().endsWith(sapPo.toUpperCase()) || sapPo.toUpperCase().endsWith(x.po.toUpperCase())));
      let rws = [];
      files.forEach((x) => x.rows.forEach((rw) => { if (vars.some((v) => norm(rw.material).includes(v))) rws.push(rw); }));
      if (rws.length) {
        const sum = (a) => a.reduce((x, y) => x + y.qty, 0);
        const tot = sum(rws);
        const sheetNames = [...new Set(rws.map((rw) => rw.sheet || ''))];
        const anyHid = rws.some((rw) => rw.hid);
        if (it.qty && (Math.abs(tot - it.qty) > 0.001 || sheetNames.length > 1)) {
          /* file nhiều sheet: sheet "Sheet" = số theo PO, các sheet khác ("89-45", "10-15"…) = số thực giao.
             Dòng lọc ẩn: thử cả "mọi dòng" và "chỉ dòng đang hiện" → chọn phương án có tổng khớp hóa đơn */
          const sets = [];
          const ordered = sheetNames.slice().sort((x, y) => (x === 'Sheet') - (y === 'Sheet'));   // ưu tiên sheet thực giao
          ordered.forEach((n) => {
            const all = rws.filter((rw) => (rw.sheet || '') === n);
            const vis = all.filter((rw) => !rw.hid);
            if (vis.length && vis.length < all.length) sets.push({ n, lab: `"${n}" (bỏ ${all.length - vis.length} dòng lọc ẩn)`, r: vis });
            sets.push({ n, lab: `"${n}"`, r: all });
          });
          const desc = () => sets.map((x) => `${x.lab}: ${fmt(sum(x.r))}`).join('; ');
          let done = false;
          const ok = sets.find((x) => Math.abs(sum(x.r) - it.qty) < 0.001);
          if (ok) {
            rws = ok.r; done = true;
            if (sets.length > 1) xNote = `Packing list có ${sets.length} cách tính cho mã này (${desc()}) — đã lấy ${ok.lab} = ${fmt(it.qty)} khớp hóa đơn. `;
          }
          if (!done) {
            for (const x of sets) {
              const byOrd = {};
              x.r.forEach((rw) => { byOrd[rw.order || '?'] = (byOrd[rw.order || '?'] || 0) + rw.qty; });
              const okOrd = Object.keys(byOrd).filter((k) => Math.abs(byOrd[k] - it.qty) < 0.001);
              if (okOrd.length === 1) {
                rws = x.r.filter((rw) => (rw.order || '?') === okOrd[0]);
                xOrder = okOrd[0]; done = true;
                xNote = `Packing list${sets.length > 1 ? ` (sheet ${x.lab})` : ''} gộp ${Object.keys(byOrd).length} Order No (tổng ${fmt(sum(x.r))}) — đã lấy đúng nhóm ${okOrd[0]} = ${fmt(it.qty)}. `;
                break;
              }
            }
          }
          if (!done) {
            if (sets.length > 1) {
              const best = sets.slice().sort((x, y) => Math.abs(sum(x.r) - it.qty) - Math.abs(sum(y.r) - it.qty))[0];
              rws = best.r;
              xNote = `⚠ Không cách tính packing list nào khớp hóa đơn ${fmt(it.qty)} (${desc()}) — đang dùng ${best.lab} gần nhất. `;
            } else {
              const byOrd = {};
              rws.forEach((rw) => { byOrd[rw.order || '?'] = (byOrd[rw.order || '?'] || 0) + rw.qty; });
              xNote = `⚠ Packing list tổng ${fmt(tot)} ≠ hóa đơn ${fmt(it.qty)}, không nhóm Order No nào khớp (${Object.entries(byOrd).map(([k, v]) => k + ': ' + fmt(v)).join('; ')}). `;
            }
          }
        } else if (anyHid && !it.qty) { /* không có SL hóa đơn để chọn → giữ mọi dòng */ }
        xRows = rws;
      }
    }
    let xMap = null;
    /* mã trong packing list có thể là mã gốc 12 ký tự (PSTIPAPR0003) còn inbound là mã đầy đủ có đuôi size
       (PSTIPAPR0003011) → so theo tiền tố */
    const xLens = xRows && xRows.length ? [...new Set(xRows.map((rw) => norm(rw.material).length))] : [];
    const matKeys = (m) => { const n = norm(m); return [n, ...xLens.filter((L) => L < n.length).map((L) => n.slice(0, L))]; };
    if (xRows && xRows.length && !ttgDone) {
      xMap = {};
      xRows.forEach((rw) => {
        const k1 = norm(rw.material) + '|' + rw.size + '|' + norm(rw.spec || rw.ref);
        xMap[k1] = (xMap[k1] || 0) + rw.qty;
        const k2 = '~' + norm(rw.material) + '|' + rw.size;
        xMap[k2] = (xMap[k2] || 0) + rw.qty;
      });
    }
    let xCov = 0, xMiss = 0;
    if (ttgDone) {
      if (ttgCov > 0) { xMap = {}; xCov = ttgCov; xMiss = ttgMiss; }
      else { hit.forEach((h) => { h.setQty = null; }); xRows = null; }
    }
    if (xMap && !ttgDone) {
      /* nhiều dòng inbound trùng một khoá packing list (khác MO) → chia theo số inbound, không cộng trùng */
      /* lượt 1: khớp đủ Material + Size + Spec; lượt 2: dòng inbound còn lại lấy phần packing list CÒN DƯ
         cùng Material + Size (không cộng lại phần đã chia ở lượt 1) */
      const sz = (h) => String(h.size || '').replace(/\s+/g, '').toUpperCase();
      const left = {};
      Object.keys(xMap).forEach((k) => { if (k[0] !== '~') left[k] = xMap[k]; });
      const give = (hs, take) => hs.forEach((h, i) => {
        const cap = isNaN(h.qty) ? 0 : h.qty;
        const want = i === hs.length - 1 ? Infinity : cap;
        const v = take(want);
        h.setQty = v > 0 ? v : null; xCov += Math.max(v, 0);
      });
      const g1 = {}, rest = [];
      hit.forEach((h) => {
        h.setQty = null;
        const k = matKeys(h.material).map((m) => m + '|' + sz(h) + '|' + norm(h.spec)).find((x) => left[x] != null);
        if (k) (g1[k] = g1[k] || []).push(h); else rest.push(h);
      });
      Object.entries(g1).forEach(([k, hs]) => give(hs, (want) => { const v = Math.min(left[k], want); left[k] -= v; return v; }));
      const g2 = {};
      rest.forEach((h) => {
        const m = matKeys(h.material).find((mm) => Object.keys(left).some((x) => x.startsWith(mm + '|' + sz(h) + '|')));
        if (m) (g2[m + '|' + sz(h) + '|'] = g2[m + '|' + sz(h) + '|'] || []).push(h);
      });
      Object.entries(g2).forEach(([pre, hs]) => give(hs, (want) => {
        let got = 0;
        for (const x of Object.keys(left)) {
          if (!x.startsWith(pre) || left[x] <= 0) continue;
          const v = Math.min(left[x], want - got); left[x] -= v; got += v;
          if (got >= want) break;
        }
        return got;
      }));
      if (xCov <= 0) { xMap = null; hit.forEach((h) => { h.setQty = null; }); }
      else xMiss = hit.filter((h) => h.setQty == null).length;
    }

    const sumInvQty = hit.reduce((a, b) => a + (isNaN(b.invQty) ? 0 : b.invQty), 0);
    const sumQty = hit.reduce((a, b) => a + (isNaN(b.qty) ? 0 : b.qty), 0);
    const base = xMap ? xCov : (sumInvQty > 0 ? sumInvQty : sumQty);
    const useInv = xMap ? true : sumInvQty > 0;
    const fallbackQty = !useInv && $('#useQty').checked;
    hit.forEach((h) => {
      if (xMap) { h.eff = h.setQty == null ? 0 : h.setQty; if (h.eff > 0) h.matched = it; return; }
      h.eff = useInv ? (isNaN(h.invQty) ? 0 : h.invQty) : (isNaN(h.qty) ? 0 : h.qty);
      if (h.eff > 0 && (useInv || fallbackQty)) h.matched = it;
    });
    const pklCodes = [...new Set(hit.map((h) => (h.spec.match(/\/([A-Z0-9]{3,6})#/) || [])[1]).filter(Boolean))];

    const bySize = {};
    hit.forEach((h) => { const s = h.size || '?'; bySize[s] = (bySize[s] || 0) + h.eff; });

    const inbPrices = [...new Set(hit.map((h) => h.price).filter((v) => !isNaN(v)))].sort((a, b) => a - b);
    const inbAmount = hit.reduce((a, h) => a + h.eff * (isNaN(h.price) ? 0 : h.price) + (isNaN(h.sur) ? 0 : h.sur), 0);
    const priceBad = hit.length > 0 && (inbPrices.length > 1 || (!!it.price && inbPrices.length === 1 && inbPrices[0] !== it.price));
    // chỉ so thành tiền khi số lượng đã khớp — lệch SL thì đương nhiên lệch tiền
    const amtBad = hit.length > 0 && useInv && base === it.qty && !!it.amount && Math.abs(inbAmount - it.amount) > 0.5;

    let pklSize = {}, pklTotal = 0; const pklCodesUsed = new Set();
    const pklFromX = !!(xRows && xRows.length);
    if (pklFromX) {
      xRows.forEach((rw) => {
        const k = rw.size || '?';
        pklSize[k] = (pklSize[k] || 0) + rw.qty;
        pklTotal += rw.qty;
        if (rw.spec || rw.ref) pklCodesUsed.add(rw.spec || rw.ref);
      });
    }
    // packing list PDF
    if (pkl && !pklFromX) {
      const cand = pkl.filter((g) => g.label === codeN || vars.includes(g.label));
      let pg = [];
      if (rp.via === 'ScaX' && rp.scax) pg = cand.filter((g) => g.po && norm(g.po).endsWith(norm(rp.scax).slice(-4)));
      else if (sapPo) pg = cand.filter((g) => g.poSap === sapPo.toUpperCase() || norm(g.text).includes(norm(sapPo)));
      if (!pg.length) pg = cand.filter((g) => !g.po && !g.poSap);
      if (!pg.length) pg = cand;
      pg.forEach((g) => {
        Object.entries(g.inv).forEach(([code, sizes]) => {
          if (pklCodes.length && !pklCodes.includes(code)) return;
          pklCodesUsed.add(code);
          Object.entries(sizes).forEach(([s, q]) => { pklSize[s] = (pklSize[s] || 0) + q; });
        });
      });
      pklTotal = Object.values(pklSize).reduce((a, b) => a + b, 0);
    }

    let poRows = [];
    if (rows && !hit.length && sapPo && po) {
      poRows = po.rows.filter((x) => x.po.toUpperCase() === sapPo.toUpperCase()
        && vars.some((v) => norm(x.spec).includes(v) || norm(x.name).includes(v)));
      const codes = [...pklCodesUsed];
      if (codes.length) {
        const f = poRows.filter((x) => codes.some((c) => norm(x.spec).includes('/' + norm(c) + '#')));
        if (f.length) poRows = f;
      }
    }

    const diffs = [];
    const sizes = [...new Set([...Object.keys(bySize), ...Object.keys(pklSize)])]
      .sort((a, b) => SIZE_ORDER.indexOf(a) - SIZE_ORDER.indexOf(b));
    for (const s of sizes) {
      const a = bySize[s] || 0, b = pklSize[s] || 0;
      if (a !== b) diffs.push({ size: s, inb: a, pkl: b, diff: a - b });
    }

    let status, note;
    if (!sapPo && !po) { status = 'THIẾU FILE PO'; note = `Mã PO trên hóa đơn ("${it.po || 'không đọc được'}") chưa đúng dạng SAP — cần tải file PO SCAF-SCAX để tra chuyển đổi`; }
    else if (!sapPo) { status = 'LỖI'; note = `Không tìm thấy PO "${it.po}" trong file PO SCAF-SCAX (đã dò cả cột PO No ScaX và PO No.)`; }
    else if (!rows) {
      status = 'CHƯA CÓ INBOUND';
      note = pklTotal ? (pklTotal === it.qty ? `Chưa có file inbound — packing list khớp hóa đơn (${fmt(it.qty)})`
        : `Chưa có file inbound — packing list ${fmt(pklTotal)} ≠ hóa đơn ${fmt(it.qty)}`) : 'Chưa có file inbound';
    } else if (!hit.length) {
      status = 'THIẾU DÒNG';
      note = `Không có dòng nào trong file inbound cho PO ${sapPo} + item "${it.code}" (đã thử cả ghép theo đơn giá ${fmt(it.price)})` +
        (poRows.length ? ` — file PO có ${poRows.length} dòng (${fmt(poRows.reduce((a, b) => a + (isNaN(b.qty) ? 0 : b.qty), 0))} pcs)` : '');
    } else if (priceBad || amtBad) {
      status = 'LỆCH GIÁ TRỊ';
      note = (priceBad ? `Đơn giá hóa đơn ${fmt(it.price)} ≠ đơn giá inbound ${inbPrices.map(fmt).join(' / ')}. ` : '') +
        (amtBad ? `Thành tiền hóa đơn ${fmt(it.amount)} ≠ inbound ${fmt(inbAmount)} (lệch ${fmt(inbAmount - it.amount)}). ` : '') +
        (!useInv ? `(Cột Invoice Quantity đang trống, SL theo Quantity là ${fmt(base)}.) ` : '') +
        'CẦN KIỂM TRA LẠI HÓA ĐƠN.';
    } else if (!useInv) {
      status = 'CHƯA ĐIỀN SL HĐ';
      note = `Cột Invoice Quantity của ${hit.length} dòng inbound đang trống — SL theo cột Quantity là ${fmt(base)}, hóa đơn ${fmt(it.qty)}. ` +
        'Hãy điền Invoice Quantity rồi chạy lại (hoặc tick "Dùng cột Quantity khi Invoice Quantity còn trống").';
    } else if (base !== it.qty) {
      status = 'LỆCH SL';
      note = xNote + (xMiss ? `${xMiss} dòng inbound không có trong packing list. ` : '')
        + (it.ttg && !pklFromX && hit.length > 1 ? `Hóa đơn Thiên Gia ghi gộp ${hit.length} mã code — thả kèm packing list PDF của Thiên Gia để chia đúng từng dòng. ` : '')
        + `Inbound ${fmt(base)} vs hóa đơn ${fmt(it.qty)} (lệch ${fmt(base - it.qty)})` +
        (diffs.length ? ' — size lệch: ' + diffs.map((d) => `${d.size}: inbound ${d.inb} / PKL ${d.pkl}`).join('; ') : '');
    } else if ((pkl || pklFromX) && pklTotal && pklTotal !== it.qty) {
      status = 'LỆCH PKL';
      note = `Inbound khớp hóa đơn (${fmt(it.qty)}) nhưng packing list ${fmt(pklTotal)} pcs` +
        (diffs.length ? ' — size lệch: ' + diffs.map((d) => `${d.size}: inbound ${d.inb} / PKL ${d.pkl}`).join('; ') : '');
    } else {
      status = 'KHỚP';
      note = (matchBy === 'đơn giá' ? 'Ghép theo đơn giá (mã hàng không có trong file inbound). ' : '')
        + (matchBy !== 'mã hàng' && matchBy !== 'đơn giá' ? `Ghép theo ${matchBy}. ` : '')
        + (R.pickedByQty ? 'Trong PO có nhiều dòng cùng họ mã — đã chọn nhóm có tổng số lượng khớp hóa đơn. ' : '')
        + (R.ambiguous ? `⚠ Mã hàng chưa đủ phân biệt: inbound còn ${R.alt.length} dòng cùng họ mã `
            + `(${[...new Set(R.alt.map((x) => x.desc.trim().slice(0, 40)))].slice(0, 2).join(' · ')}`
            + `, ${fmt(R.alt.reduce((a, x) => a + (x.qty || 0), 0))} pcs) — kiểm tra lại xem có chọn đúng dòng không. ` : '')
        + (pklFromX ? xNote + (ttgDone ? `Packing list Thiên Gia — đã điền Invoice Quantity cho ${hit.filter((h) => h.setQty != null).length} dòng theo PO + kích thước + mã code` : `Packing list Excel theo PO — đã điền Invoice Quantity cho ${hit.filter((h) => h.setQty != null).length} dòng theo Material Code + Size`) + (xOrder ? ` (Order ${xOrder})` : '') + '. ' : '')
        + (R.byPrice ? `Cùng PO có nhiều mặt hàng trùng kích thước — đã tách bằng đơn giá ${fmt(it.price)}. ` : '')
        + (pkl || pklFromX ? '' : 'Không có packing list — chỉ đối chiếu với hóa đơn');
    }
    if (matchBy !== 'mã hàng' && status !== 'KHỚP') note = `Ghép theo ${matchBy}. ` + note;
    if (R.ambiguous && status !== 'KHỚP') {
      note = `⚠ Mã hàng chưa đủ phân biệt — inbound còn ${R.alt.length} dòng cùng họ mã `
        + `(${[...new Set(R.alt.map((x) => x.desc.trim().slice(0, 40)))].slice(0, 2).join(' · ')}`
        + `, ${fmt(R.alt.reduce((a, x) => a + (x.qty || 0), 0))} pcs). ` + note;
    }

    lines.push({
      it, sapPo, hit, base, sumQty, sumInvQty, pklCodes: [...pklCodesUsed], pklSize, pklTotal,
      bySize, diffs, poRows, status, note, inbPrices, inbAmount, priceBad, amtBad, xOrder, xMiss,
      hasInb: !!rows, hasPkl: !!pkl || pklFromX, pklFromX, useInv, matchBy, ambiguous: !!R.ambiguous, usedVar: R.usedVar,
      pickedByQty: !!R.pickedByQty, altRows: (R.alt || []).length,
    });
  }

  const itemsTotal = lines.reduce((a, l) => a + (isNaN(l.it.amount) ? 0 : l.it.amount), 0);
  const invTotal = isNaN(inv.total) ? itemsTotal : inv.total;
  // chỉ so tổng tiền trên các dòng đã điền Invoice Quantity
  const cmp = lines.filter((l) => l.hasInb && l.useInv);
  const inbTotal = cmp.reduce((a, l) => a + l.inbAmount, 0);
  const invCmpTotal = cmp.reduce((a, l) => a + (isNaN(l.it.amount) ? 0 : l.it.amount), 0);
  const totalDiff = inbTotal - invCmpTotal;
  const valueLines = lines.filter((l) => l.priceBad || l.amtBad);
  const pendingLines = lines.filter((l) => l.status === 'CHƯA ĐIỀN SL HĐ');
  const otherLines = lines.filter((l) => !(l.priceBad || l.amtBad)
    && !['KHỚP', 'LỆCH PKL', 'CHƯA CÓ INBOUND', 'CHƯA ĐIỀN SL HĐ'].includes(l.status));
  const amtOf = (l) => l.inbAmount - (isNaN(l.it.amount) ? 0 : l.it.amount);
  const VAL = {
    inbTotal, itemsTotal, invTotal, invCmpTotal, totalDiff,
    valueBad: valueLines.length > 0,
    totalBad: cmp.length > 0 && Math.abs(totalDiff) > 0.5,
    valueLines, otherLines, pendingLines, cmpCount: cmp.length,
    valueDiff: valueLines.reduce((a, l) => a + amtOf(l), 0),
    otherDiff: otherLines.reduce((a, l) => a + amtOf(l), 0),
    hasInb: !!rows,
  };
  return { lines, VAL };
}

/* ================= Đọc dòng của file inbound ================= */
function readInbRows(ws, H) {
  const rows = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const poV = row.getCell(H.po).text.trim();
    if (!poV) continue;
    rows.push({
      r, poV,
      desc: row.getCell(H.desc).text,
      spec: H.spec ? row.getCell(H.spec).text : '',
      material: H.material ? row.getCell(H.material).text : '',
      supRef: H.supRef ? row.getCell(H.supRef).text : '',
      color: H.color ? row.getCell(H.color).text : '',
      lapdip: H.lapdip ? row.getCell(H.lapdip).text : '',
      unit: H.unit ? row.getCell(H.unit).text.trim() : '',
      cur: H.cur ? row.getCell(H.cur).text.trim() : '',
      overTol: H.overTol ? num(row.getCell(H.overTol).value) : NaN,
      size: H.size ? row.getCell(H.size).text.replace(/\s+/g, '').toUpperCase() : '',
      qty: num(row.getCell(H.qty).value),
      deliv: H.deliv ? num(row.getCell(H.deliv).value) : 0,
      invQty: H.invQty ? num(row.getCell(H.invQty).value) : NaN,
      price: H.price ? num(row.getCell(H.price).value) : NaN,
      sur: H.sur ? num(row.getCell(H.sur).value) : 0,
      eff: 0, matched: null, setQty: null,
    });
  }
  return rows;
}

/* ================= Xuất file ================= */
const OUTPUTS = [];
async function addDownload(book, fname, label, cls) {
  const blob = new Blob([await book.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  OUTPUTS.push({ url: URL.createObjectURL(blob), fname, label, cls });
}
function renderDownloads() {
  $('#dls').innerHTML = OUTPUTS.map((o, i) => `<a class="dl ${o.cls || ''}" id="dlx${i}" href="${o.url}" download="${o.fname}">⬇ ${o.label}</a>`).join('');
  $('#dlall').classList.toggle('hidden', OUTPUTS.length < 2);
}

function fillSapWorkbook(ws, H, rows, inv, onlyMatched) {
  const keep = new Set();
  for (const row of rows) if (row.matched) keep.add(row.r);
  /* file SAP xuất ra đôi khi còn dòng rác (ô lẻ, không có số PO) — bỏ luôn,
     nếu giữ lại thì SAP sẽ báo lỗi hoặc tạo ra một dòng trống khi import */
  const toDelete = [];
  for (let r = 2; r <= ws.rowCount; r++) if (!keep.has(r)) toDelete.push(r);
  for (const row of rows) {
    if (row.matched) {
      ws.getCell(row.r, H.invNo).value = inv.invNo;
      ws.getCell(row.r, H.invDate).value = inv.invDate;
      ws.getCell(row.r, H.invDate).numFmt = '@';
      if (row.setQty != null && H.invQty) ws.getCell(row.r, H.invQty).value = row.setQty;
    }
  }
  /* ô công thức (kể cả công thức dùng chung) → ghi giá trị, vì xoá dòng chủ sẽ làm ExcelJS không ghi được file */
  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    row.eachCell({ includeEmpty: false }, (cell) => {
      const v = cell.value;
      if (v && typeof v === 'object' && !(v instanceof Date) && (v.formula !== undefined || v.sharedFormula !== undefined)) {
        cell.value = v.result === undefined ? null : v.result;
      }
    });
  }
  if (onlyMatched) toDelete.slice().sort((a, b) => b - a).forEach((r) => ws.spliceRows(r, 1));
  return keep.size;
}

/* ================= Chạy ================= */
async function run() {
  if (STATE.busy) return;
  if (!STATE.groups.length) { log('Chưa có hóa đơn nào.', 'err'); return; }
  const groups = selectedGroups();
  if (!groups.length) { log('Chưa chọn hóa đơn nào — tick ở cột đầu của bảng.', 'err'); return; }
  STATE.busy = true; $('#run').disabled = true; $('#runpo').disabled = true;
  $('#log').innerHTML = ''; $('#report').innerHTML = ''; $('#dls').innerHTML = '';
  OUTPUTS.length = 0;
  try {
    const po = STATE.poIdx;
    const needPo = groups.some((g) => !g.isFab);
    if (!po && needPo) log('Chưa có file PO SCAF-SCAX — chỉ cần khi hóa đơn ghi PO kiểu cũ (TRIMMINGVN-xxxx); PO đã đúng dạng SAP thì dùng trực tiếp.', 'ok');
    else if (!po) log('Chứng từ vải dùng PO ScaF sẵn — chưa cần file PO SCAF-SCAX (nếu có PO hệ cũ em sẽ nhắc).', 'ok');
    else log(`PO SCAF-SCAX: ${po.map.size} mã ScaX + ${po.sap.size} mã ScaF`, 'ok');

    const onlyMatched = $('#onlyMatched').checked;
    const perInvoiceReport = $('#perReport').checked;
    const all = [];

    for (const g of groups) {
      const inv = g.inv.inv;
      log(`── Hóa đơn ${inv.invNo || g.inv.file.name} (${inv.items.length} dòng)${g.isFab ? ` — ${g.isGen ? '' : 'vải · '}${g.fabName}` : ((g.pkl || (g.pklx && g.pklx.length)) ? '' : ' — không có packing list')}${g.inb ? '' : ' — CHƯA CÓ INBOUND'}`);
      if (g.isGen && inv.noFrom) log(`  Số hoá đơn "${inv.invNo}" lấy từ ${inv.noFrom} — kiểm lại trước khi import.`, 'warn');
      if (inv.ocr) log('  ⚠ Chứng từ là bản scan, đọc bằng OCR — kiểm tra kỹ số liệu với bản gốc trước khi import.', 'err');
      if (g.isGen && inv.explodedByPkl) log('  Hoá đơn ghi gộp theo mã hàng — đã tách dòng theo packing list (PO/size).', 'ok');
      if (g.isFab && g.pdfInv) log(`  Hóa đơn GTGT: ${inv.invNo} ngày ${inv.invDate} (${g.pdfInv.file.name})`, 'ok');
      const pklGroups = g.isFab ? g.inv.fab.pkl : (g.pkl ? parsePacking((await classify(g.pkl.file)).lines) : null);
      const pxList = (g.pklx || []).map((p) => p.px).filter(Boolean);
      const doAnalyze = (rr) => (g.isFab
        ? analyzeFab(inv, pklGroups, rr, { poIdx: po ? { scax: po.map, sap: po.sap } : null, hasPoFile: !!po })
        : analyze(inv, pklGroups, rr, po, pxList));
      let rows = null, wbSap = null, wsSap = null, H = null, cInb = null;
      if (g.inb) {
        cInb = await classify(g.inb.file);
        wbSap = new ExcelJS.Workbook();
        await wbSap.xlsx.load(cInb.buf.slice(0));
        wsSap = wbSap.worksheets[0];
        H = headerIndex(wsSap);
        if (!H.po || !H.desc || !H.invNo || !H.invDate) { log('  File inbound thiếu cột bắt buộc — bỏ qua.', 'err'); rows = null; }
        else rows = readInbRows(wsSap, H);
      }
      const { lines, VAL } = doAnalyze(rows);
      const sup = g.supTrained !== undefined ? { trained: g.supTrained, name: g.supName } : supplierCheck(g, cInb);
      all.push({ g, inv, lines, VAL, newSup: sup.trained ? '' : sup.name, yuboAlone: g.isGen && sup.trained === 'J&H Yubo' });
      if (g.isGen && sup.trained === 'J&H Yubo') log(`  ⚠ Đây là hóa đơn GTGT của J&H Yubo nhưng chưa có file PKL Excel đi kèm (PKL SCAVI… / file nhập inbound của Yubo) — thả thêm file PKL để công cụ đọc theo lô, chia PO và điền đúng Invoice Quantity.`, 'err');
      if (!sup.trained) log(`  ⚠ Chủ hàng "${sup.name}" chưa được huấn luyện trong công cụ — kết quả đọc bằng bộ đọc chung, cần kiểm tra kỹ. Vui lòng liên hệ ${contactText()} và gửi kèm bộ chứng từ để được bổ sung.`, 'err');

      const bad = lines.filter((l) => !isOk(l.status));
      if (VAL.valueBad) log(`  ⚠ SAI GIÁ TRỊ — lệch ${fmt(VAL.valueDiff)}: ` + VAL.valueLines.map((l) => `${l.it.code}/${l.it.po}`).join(', '), 'err');
      if (bad.length && !VAL.valueBad) log(`  ${bad.length} dòng cần xem lại: ` + bad.map((l) => `${l.it.code}/${l.it.po} (${l.status})`).join(', '), 'err');
      if (!bad.length && rows) log('  Khớp toàn bộ.', 'ok');

      if (rows) {
        const kept = fillSapWorkbook(wsSap, H, rows, inv, onlyMatched);
        const stamp = `${String(inv.invNo || 'CHUA-CO-SO-HD').replace(/[#\\\/:*?"<>|&]/g, '-')}_${inv.invDate}`;
        await addDownload(wbSap, `INB_${stamp}.xlsx`, `Import SAP – ${inv.invNo} (${kept} dòng)`, 'main');
        if (perInvoiceReport) {
          const wbR = new ExcelJS.Workbook();
          await wbR.xlsx.load(cInb.buf.slice(0));
          const wsR = wbR.worksheets[0];
          const rows2 = readInbRows(wsR, H);
          const a2 = doAnalyze(rows2);
          buildDetailSheet(wbR, wsR, H, rows2, a2.lines, inv, a2.VAL, g);
          await addDownload(wbR, `BAOCAO_${stamp}.xlsx`, `Báo cáo chi tiết – ${inv.invNo}`, 'alt');
        }
      }
    }

    const wbSum = new ExcelJS.Workbook();
    buildSummaryWorkbook(wbSum, all);
    await addDownload(wbSum, `BAOCAO_TONGHOP_${new Date().toISOString().slice(0, 10)}.xlsx`, `Báo cáo tổng hợp (${all.length} hóa đơn)`, 'alt');

    renderDownloads();
    renderAll(all);
    RESULT = all;
    window.__RESULT__ = all.map((a) => ({
      invNo: a.inv.invNo, invDate: a.inv.invDate, pkl: !!a.g.pkl, inb: !!a.g.inb,
      invTotal: a.VAL.invTotal, inbTotal: a.VAL.inbTotal, valueBad: a.VAL.valueBad,
      lines: a.lines.map((l) => ({ code: l.it.code, po: l.it.po, via: l.it.poVia, sapPo: l.sapPo, invQty: l.it.qty, inbQty: l.base, pklQty: l.pklTotal, status: l.status, note: l.note })),
    }));
    STATE.lastRunSel = groups.map((x) => x.inv.key).join('|');
    $('#stale').textContent = '';
    log(`Hoàn tất — đã xử lý ${groups.length}/${STATE.groups.length} hóa đơn.`, 'ok');
  } catch (e) {
    log('Lỗi: ' + (e && e.message ? e.message : e), 'err');
    console.error(e);
  } finally {
    STATE.busy = false; $('#run').disabled = false; $('#runpo').disabled = false;
    flushQueue();
  }
}

/* ================= Xuất danh sách PO ================= */
async function runPoList() {
  if (STATE.busy) return;
  if (!STATE.groups.length) { log('Chưa có hóa đơn nào.', 'err'); return; }
  const groups = selectedGroups();
  if (!groups.length) { log('Chưa chọn hóa đơn nào — tick ở cột đầu của bảng.', 'err'); return; }
  STATE.busy = true; $('#runpo').disabled = true;
  try {
    const po = STATE.poIdx;
    const map = new Map();       // sapPo -> {sap, scax, invs:Set, qty, amount}
    const unresolved = [];
    for (const g of groups) {
      const inv = g.inv.inv;
      for (const it of inv.items) {
        const rp = g.isFab ? { scax: '', sap: it.po, via: 'ScaF' }
          : resolvePo(it, po ? { scax: po.map, sap: po.sap } : null);
        if (!rp.sap) { unresolved.push({ invNo: inv.invNo, code: it.code, po: it.po || '(không đọc được)', qty: it.qty }); continue; }
        const k = rp.sap.toUpperCase();
        if (!map.has(k)) map.set(k, { sap: rp.sap, scax: rp.scax || '', invs: new Set(), qty: 0, amount: 0, items: new Set() });
        const e = map.get(k);
        e.invs.add(inv.invNo); e.items.add(it.code);
        e.qty += isNaN(it.qty) ? 0 : it.qty;
        e.amount += isNaN(it.amount) ? 0 : it.amount;
        if (!e.scax && rp.scax) e.scax = rp.scax;
      }
    }
    const list = [...map.values()].sort((a, b) => a.sap.localeCompare(b.sap));

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('PO');
    ws.getColumn(1).width = 16; ws.getColumn(2).width = 22;
    ws.addRow(['PO No.', 'PO No ScaX']);
    ws.getRow(1).font = { bold: true };
    ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
    list.forEach((e) => ws.addRow([e.sap, e.scax]));

    const d = wb.addWorksheet('CHI TIET');
    [16, 22, 22, 14, 14, 12, 16, 16].forEach((w, i) => { d.getColumn(i + 1).width = w; });
    d.addRow(['PO No.', 'PO No ScaX', 'Hóa đơn', 'Ngày', 'Item', 'SL', 'Thành tiền', 'Ghi chú']);
    d.getRow(1).font = { bold: true };
    d.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
    for (const g of groups) {
      const inv = g.inv.inv;
      for (const it of inv.items) {
        const rp = g.isFab ? { scax: '', sap: it.po, via: 'ScaF' }
          : resolvePo(it, po ? { scax: po.map, sap: po.sap } : null);
        d.addRow([rp.sap || '(không tra được)', rp.scax || '', inv.invNo, inv.invDate, it.code,
          it.qty, isNaN(it.amount) ? '' : it.amount, g.inb ? 'đã có inbound' : 'chưa có inbound']);
      }
    }
    if (unresolved.length) {
      const u = wb.addWorksheet('KHONG TRA DUOC');
      [20, 16, 22, 12].forEach((w, i) => { u.getColumn(i + 1).width = w; });
      u.addRow(['Hóa đơn', 'Item', 'Mã PO trên hóa đơn', 'SL']);
      u.getRow(1).font = { bold: true };
      unresolved.forEach((x) => u.addRow([x.invNo, x.code, x.po, x.qty]));
    }
    OUTPUTS.length = 0;
    await addDownload(wb, `DANHSACH_PO_${new Date().toISOString().slice(0, 10)}.xlsx`, `Danh sách PO (${list.length} PO / ${groups.length} hóa đơn)`, 'main');
    renderDownloads();
    log(`Danh sách PO: ${list.length} PO từ ${groups.length} hóa đơn được chọn` + (unresolved.length ? ` — ${unresolved.length} dòng không tra được PO` : ''), 'ok');
    renderPoTable(list, unresolved);
  } catch (e) {
    log('Lỗi: ' + (e && e.message ? e.message : e), 'err');
    console.error(e);
  } finally {
    STATE.busy = false; $('#runpo').disabled = false;
    flushQueue();
  }
}

/* ================= Sheet chi tiết cho 1 hóa đơn ================= */
const EXTRA = ['Amount', 'Balance', 'Check Status', 'Invoice Item', 'PO (hóa đơn)', 'Loại PO', 'Invoice No (check)',
  'PKL Invoice Ref', 'Inbound Qty (item)', 'Invoice Qty (item)', 'PKL Qty (size)', 'Inbound Qty (size)', 'Diff (size)',
  'Đơn giá HĐ', 'Đơn giá inbound', 'Thành tiền HĐ (item)', 'Thành tiền inbound (item)', 'Lệch tiền (item)', 'Ghi chú'];

function buildDetailSheet(wb, ws, H, rows, lines, inv, VAL, g) {
  const L = {
    price: colLetter(H.price), sur: colLetter(H.sur), qty: colLetter(H.qty),
    deliv: colLetter(H.deliv), invQty: colLetter(H.invQty),
  };
  const start = H.last + 1;
  const hRow = ws.getRow(1);
  EXTRA.forEach((t, i) => {
    const c = hRow.getCell(start + i);
    c.value = t;
    c.font = Object.assign({}, hRow.getCell(1).font, { bold: true });
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
    ws.getColumn(start + i).width = i < 2 ? 14 : 18;
  });
  for (const row of rows) {
    const r = row.r, xr = ws.getRow(r), m = row.matched;
    if (m) {
      ws.getCell(r, H.invNo).value = inv.invNo;
      ws.getCell(r, H.invDate).value = inv.invDate;
      ws.getCell(r, H.invDate).numFmt = '@';
      if (row.setQty != null && H.invQty) ws.getCell(r, H.invQty).value = row.setQty;
    }
    xr.getCell(start).value = { formula: `${L.invQty}${r}*${L.price}${r}+${L.sur}${r}` };
    xr.getCell(start + 1).value = { formula: `${L.qty}${r}-(${L.invQty}${r}+${L.deliv}${r})` };
    xr.getCell(start).numFmt = '#,##0';
    xr.getCell(start + 1).numFmt = '#,##0';
    const info = m ? lines.find((x) => x.it === m) : null;
    if (info) {
      const sz = row.size || '?';
      const inbS = info.bySize[sz] || 0, pklS = info.pklSize[sz];
      const dif = pklS == null ? null : inbS - pklS;
      [info.status, m.code, m.po, m.poVia || '', inv.invNo, info.pklCodes.join(', '),
        info.base, m.qty, pklS == null ? '' : pklS, inbS, dif == null ? '' : dif,
        isNaN(m.price) ? '' : m.price, info.inbPrices.join(' / '),
        isNaN(m.amount) ? '' : m.amount, info.inbAmount, isNaN(m.amount) ? '' : info.inbAmount - m.amount,
        info.note || '',
      ].forEach((v, i) => { xr.getCell(start + 2 + i).value = v; });
      const red = info.priceBad || info.amtBad;
      if (red || !isOkK(info.status) || (dif != null && dif !== 0)) {
        for (let i = 0; i < EXTRA.length; i++) {
          xr.getCell(start + i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: red ? 'FFFFC7CE' : 'FFFFF2CC' } };
        }
        xr.getCell(start + 2).font = { bold: true, color: { argb: 'FFC00000' } };
      }
    } else xr.getCell(start + 2).value = 'KHÔNG THUỘC HÓA ĐƠN NÀY';
  }
  const rs = wb.addWorksheet('BAO CAO');
  writeInvoiceReport(rs, 1, inv, lines, VAL, g, true);
}

/* Khối báo cáo của 1 hóa đơn, trả về dòng kế tiếp */
function writeInvoiceReport(rs, R, inv, lines, VAL, g, withWidths) {
  const put = (r, arr, bold, fill) => {
    arr.forEach((v, i) => {
      const c = rs.getCell(r, i + 1); c.value = v;
      if (bold) c.font = { bold: true };
      if (fill) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
    });
  };
  if (withWidths) [22, 18, 10, 14, 12, 12, 12, 10, 11, 14, 15, 15, 12, 15, 60].forEach((w, i) => { rs.getColumn(i + 1).width = w; });
  put(R, [`HÓA ĐƠN ${inv.invNo || g.inv.file.name.replace(/\.[^.]+$/, '')} — ngày ${inv.invDate}`], true, 'FFEEEEEE');
  rs.getCell(R, 1).font = { bold: true, size: 12 }; R++;
  put(R++, ['File hóa đơn', g.inv.file.name]);
  put(R++, ['File packing list', g.pkl ? g.pkl.file.name
    : ((g.pklx && g.pklx.length) ? g.pklx.map((p) => p.file.name).join(' · ')
      : (g.isFab ? '(nằm trong cùng file chứng từ)' : '(không có)'))]);
  put(R++, ['File inbound', g.inb ? g.inb.file.name : '(chưa có)']);
  if (inv.pdfFile) put(R++, ['File hóa đơn GTGT (PDF)', inv.pdfFile, `${inv.invNo} · ${inv.invDate}`]);
  if (VAL.noSerial) {
    put(R, ['⚠ Ô "HD:" chỉ có số, chưa có ký hiệu hóa đơn — thả thêm file PDF hóa đơn GTGT để điền đủ (vd 1C26TYY#00001728).'], true, 'FFFFF2CC');
    R++;
  }
  if (VAL.pdfBad) {
    put(R, [`⚠ Tổng tiền hàng trên PKL (${VAL.itemsTotal}) ≠ hóa đơn GTGT (${VAL.pdfTotal}), lệch ${VAL.pdfDiff} — KIỂM TRA LẠI`], true, 'FFFFC7CE');
    rs.getCell(R, 1).font = { bold: true, color: { argb: 'FFC00000' } };
    R++;
  } else if (!isNaN(VAL.pdfDiff) && VAL.pdfTotal != null) {
    put(R++, ['Đối chiếu với hóa đơn GTGT', VAL.pdfTotal, 'khớp']);
  }
  if (VAL.noInvoiceNo) {
    put(R, [g.isGen ? '⚠ CHỨNG TỪ CHƯA CÓ SỐ HÓA ĐƠN — cột Invoice Number để trống. Hãy điền tay vào file INB trước khi import.'
      : '⚠ CHỨNG TỪ CHƯA CÓ SỐ HÓA ĐƠN — cột Invoice Number để trống. Điền ô "HD:" trong file rồi thả lại.'], true, 'FFFFF2CC');
    R++;
  } else if (VAL.noFrom) {
    put(R, [`⚠ Số hóa đơn "${inv.invNo}" lấy từ ${VAL.noFrom} (trong chứng từ không có nhãn số hóa đơn) — kiểm lại trước khi import.`], true, 'FFFFF2CC');
    R++;
  }
  if (VAL.noAmounts && g.isGen) put(R++, ['Chứng từ không có cột đơn giá/thành tiền — chỉ đối chiếu số lượng, không kiểm giá trị.']);
  if (inv.ocr) {
    put(R, ['⚠ CHỨNG TỪ LÀ BẢN SCAN, ĐỌC BẰNG OCR — số liệu có thể nhận dạng sai (0/O, 1/I, dấu phẩy). ĐỐI CHIẾU TAY VỚI BẢN GỐC TRƯỚC KHI IMPORT.'], true, 'FFFFC7CE');
    rs.getCell(R, 1).font = { bold: true, color: { argb: 'FFC00000' } }; R++;
  }
  if (VAL.valueBad) {
    put(R, ['⚠ HÓA ĐƠN NÀY SAI GIÁ TRỊ – CẦN KIỂM TRA LẠI (đơn giá/thành tiền không khớp PO)'], true, 'FFFFC7CE');
    rs.getCell(R, 1).font = { bold: true, size: 12, color: { argb: 'FFC00000' } }; R++;
  } else if (VAL.totalBad) {
    put(R, ['⚠ Tổng tiền chưa khớp do có dòng thiếu/lệch số lượng trong file inbound (đơn giá vẫn đúng)'], true, 'FFFFF2CC');
    R++;
  }
  if (VAL.qtyBad) {
    put(R, [`⚠ TỔNG SỐ LƯỢNG KHÔNG KHỚP: chứng từ ${VAL.docQty} ${VAL.unitLabel || ''} · đã ghi vào inbound ${VAL.wroteQty} · lệch ${Math.round((VAL.wroteQty - VAL.docQty) * 1000) / 1000} — KIỂM TRA LẠI`], true, 'FFFFC7CE');
    rs.getCell(R, 1).font = { bold: true, size: 12, color: { argb: 'FFC00000' } };
    R++;
  } else if (VAL.docQty != null && !isNaN(VAL.docQty)) {
    put(R++, ['Tổng số lượng (chứng từ / đã ghi vào inbound)', VAL.docQty, VAL.wroteQty, 'khớp']);
  }
  put(R++, ['Cộng tiền hàng (hóa đơn)', VAL.invTotal, VAL.currency || '']);
  if (VAL.surchargeHeader) put(R++, ['Trong đó phụ phí (dòng riêng trên hóa đơn)', VAL.surchargeHeader]);
  if (VAL.hasInb && VAL.cmpCount) {
    if (VAL.pendingLines.length) put(R++, ['Phần đối chiếu được (hóa đơn)', VAL.invCmpTotal,
      `${VAL.pendingLines.length} dòng chưa điền Invoice Quantity nên chưa đối chiếu tiền`]);
    put(R++, ['Tổng thành tiền theo inbound', VAL.inbTotal]);
    put(R++, ['Chênh lệch', VAL.totalDiff, VAL.totalBad ? 'CẦN KIỂM TRA LẠI' : 'Khớp'], true,
      VAL.valueBad ? 'FFFFC7CE' : (VAL.totalBad ? 'FFFFF2CC' : null));
  }
  if (!isNaN(inv.vat)) put(R++, ['Tiền thuế GTGT (hóa đơn)', inv.vat]);
  if (!isNaN(inv.payment)) put(R++, ['Tổng cộng tiền thanh toán (hóa đơn)', inv.payment]);
  R++;
  put(R++, ['Item', 'PO (hóa đơn)', 'Loại PO', 'PO SAP', 'SL hóa đơn', 'SL inbound', 'SL packing', 'Lệch SL',
    'Đơn giá HĐ', 'Đơn giá inbound', 'Thành tiền HĐ', 'Thành tiền inbound', 'Lệch tiền', 'Kết quả', 'Ghi chú'], true, 'FFDDEBF7');
  for (const l of lines) {
    const red = l.priceBad || l.amtBad;
    const rr = R;
    put(R++, [l.it.code, l.it.po, l.it.poVia || '(?)', l.sapPo || '(không tìm thấy)', l.it.qty,
      l.hasInb ? l.base : '', l.pklTotal || '', l.hasInb ? l.base - l.it.qty : '',
      isNaN(l.it.price) ? '' : l.it.price, l.inbPrices.join(' / '),
      isNaN(l.it.amount) ? '' : l.it.amount, l.hasInb ? l.inbAmount : '',
      (l.hasInb && !isNaN(l.it.amount)) ? l.inbAmount - l.it.amount : '',
      l.status, l.note], false,
      isOkK(l.status) ? null : (red ? 'FFFFC7CE' : 'FFFFF2CC'));
    if (red) rs.getCell(rr, 14).font = { bold: true, color: { argb: 'FFC00000' } };
  }
  R++;
  if (lines.some((l) => (l.lots || []).length)) {
    put(R++, ['CHI TIẾT LÔ / CÂY VẢI'], true, 'FFEEEEEE');
    put(R++, ['Item', 'PO SAP', 'Lô (Lot/Batch)', 'Số cây', 'Số lượng', 'Đơn vị', 'SL hóa đơn'], true, 'FFDDEBF7');
    for (const l of lines) {
      for (const L of (l.lots || [])) put(R++, [l.it.code, l.sapPo, L.lot, L.rolls, L.qty, L.unit || '', l.base]);
    }
    R++;
  }
  const bad = lines.filter((l) => !isOk(l.status));
  for (const l of bad) {
    put(R++, [`${l.it.code} / ${l.it.po} → ${l.sapPo}`], true);
    if (l.diffs.length) {
      put(R++, ['', 'PKL Invoice', 'Size', 'SL inbound', 'SL packing', 'Lệch'], true, 'FFEEEEEE');
      for (const d of l.diffs) put(R++, ['', l.pklCodes.join(', '), d.size, d.inb, d.pkl, d.diff]);
    }
    if (l.poRows.length) {
      put(R++, ['', 'Dòng có trong file PO nhưng thiếu trong inbound:'], true);
      put(R++, ['', 'Specification', 'Size', 'SL PO', 'Đơn giá'], true, 'FFEEEEEE');
      for (const p of l.poRows) put(R++, ['', p.spec, p.size, p.qty, p.price]);
      put(R++, ['', 'Tổng', '', l.poRows.reduce((a, b) => a + (isNaN(b.qty) ? 0 : b.qty), 0)], true);
    }
    if (l.priceBad || l.amtBad) {
      put(R++, ['', 'Đơn giá hóa đơn', isNaN(l.it.price) ? '' : l.it.price, 'Đơn giá inbound', l.inbPrices.join(' / ')], false, 'FFFFC7CE');
      put(R++, ['', 'Thành tiền hóa đơn', isNaN(l.it.amount) ? '' : l.it.amount, 'Thành tiền inbound', l.inbAmount,
        'Lệch', isNaN(l.it.amount) ? '' : l.inbAmount - l.it.amount], false, 'FFFFC7CE');
    }
    if (!l.diffs.length && !l.poRows.length && !(l.priceBad || l.amtBad)) put(R++, ['', l.note]);
    R++;
  }
  if (!bad.length) put(R++, ['Tất cả các dòng khớp số lượng và giá trị.'], true);
  return R + 1;
}

/* ================= Báo cáo tổng hợp nhiều hóa đơn ================= */
const invLabel = (a) => a.inv.invNo || (a.g && a.g.inv ? a.g.inv.file.name.replace(/\.[^.]+$/, '') : '(chưa có số HĐ)');

function buildSummaryWorkbook(wb, all) {
  const head = (ws, arr, widths) => {
    widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
    ws.addRow(arr);
    ws.getRow(1).font = { bold: true };
    ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
  };
  const paint = (ws, fill) => { if (fill) ws.getRow(ws.rowCount).eachCell((c) => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } }; }); };

  const s = wb.addWorksheet('TONG HOP');
  head(s, ['Hóa đơn', 'Ngày', 'Packing list', 'Inbound', 'Số dòng', 'Dòng cần xem lại', 'Tiền hàng (HĐ)',
    'Đối chiếu được (HĐ)', 'Theo inbound', 'Lệch tiền', 'Tổng SL chứng từ', 'Tổng SL ghi inbound', 'Lệch SL', 'Kết luận', 'File hóa đơn'],
    [20, 12, 12, 12, 10, 16, 16, 18, 16, 14, 16, 18, 12, 44, 34]);
  for (const a of all) {
    const bad = a.lines.filter((l) => !isOk(l.status));
    const concl = !a.VAL.hasInb ? 'CHƯA CÓ INBOUND — cần tạo inbound trên SAP'
      : a.VAL.qtyBad ? `⚠ LỆCH TỔNG SỐ LƯỢNG ${Math.round((a.VAL.wroteQty - a.VAL.docQty) * 1000) / 1000} – CẦN KIỂM TRA LẠI`
      : a.VAL.valueBad ? '⚠ HÓA ĐƠN SAI GIÁ TRỊ – CẦN KIỂM TRA LẠI'
        : bad.length ? `${bad.length} dòng lệch — xem sheet CHI TIET`
        : (a.VAL.pendingLines.length ? `${a.VAL.pendingLines.length} dòng chưa điền Invoice Quantity` : 'Khớp toàn bộ');
    s.addRow([invLabel(a), a.inv.invDate, a.g.pkl ? 'có' : ((a.g.pklx && a.g.pklx.length) ? `${a.g.pklx.length} file Excel` : (a.g.isFab ? 'cùng file' : 'không')), a.g.inb ? 'có' : 'chưa có',
      a.lines.length, bad.length, a.VAL.invTotal,
      a.VAL.cmpCount ? a.VAL.invCmpTotal : '', a.VAL.cmpCount ? a.VAL.inbTotal : '',
      a.VAL.cmpCount ? a.VAL.totalDiff : '',
      (a.VAL.docQty == null || isNaN(a.VAL.docQty)) ? '' : a.VAL.docQty,
      a.VAL.wroteQty == null ? '' : a.VAL.wroteQty,
      (a.VAL.docQty == null || isNaN(a.VAL.docQty)) ? '' : Math.round((a.VAL.wroteQty - a.VAL.docQty) * 1000) / 1000,
      (a.newSup ? `CHỦ HÀNG MỚI (${a.newSup}) — liên hệ ${contactText()} · ` : '') + concl, a.g.inv.file.name]);
    paint(s, a.VAL.valueBad ? 'FFFFC7CE' : (bad.length || !a.VAL.hasInb ? 'FFFFF2CC' : null));
  }

  const d = wb.addWorksheet('CHI TIET');
  head(d, ['Hóa đơn', 'Ngày', 'Item', 'PO (hóa đơn)', 'Loại PO', 'PO SAP', 'SL hóa đơn', 'SL inbound', 'SL packing',
    'Lệch SL', 'Đơn giá HĐ', 'Đơn giá inbound', 'Thành tiền HĐ', 'Thành tiền inbound', 'Lệch tiền', 'Kết quả', 'Ghi chú'],
    [20, 12, 11, 18, 9, 13, 11, 11, 11, 10, 11, 14, 14, 15, 12, 15, 70]);
  for (const a of all) {
    for (const l of a.lines) {
      d.addRow([invLabel(a), a.inv.invDate, l.it.code, l.it.po, l.it.poVia || '', l.sapPo || '', l.it.qty,
        l.hasInb ? l.base : '', l.pklTotal || '', l.hasInb ? l.base - l.it.qty : '',
        isNaN(l.it.price) ? '' : l.it.price, l.inbPrices.join(' / '),
        isNaN(l.it.amount) ? '' : l.it.amount, l.hasInb ? l.inbAmount : '',
        (l.hasInb && !isNaN(l.it.amount)) ? l.inbAmount - l.it.amount : '', l.status, l.note]);
      paint(d, (l.priceBad || l.amtBad) ? 'FFFFC7CE' : (isOk(l.status) ? null : 'FFFFF2CC'));
    }
  }

  const z = wb.addWorksheet('LECH SIZE-LO');
  head(z, ['Hóa đơn', 'Item', 'PO SAP', 'PKL Invoice / Lô', 'Size / Lô', 'SL inbound', 'SL packing', 'Lệch'],
    [20, 24, 13, 16, 16, 12, 12, 10]);
  for (const a of all) {
    for (const l of a.lines) {
      if (!l.hasInb || !l.hasPkl || !l.useInv) continue;
      for (const df of l.diffs) {
        z.addRow([invLabel(a), l.it.code, l.sapPo, l.pklCodes.join(', '), df.size, df.inb, df.pkl, df.diff]);
        paint(z, 'FFFFF2CC');
      }
    }
  }
  if (z.rowCount === 1) z.addRow(['Không có size/lô nào lệch giữa inbound và packing list.']);

  if (all.some((a) => a.lines.some((l) => (l.alloc || []).length > 1))) {
    const ap = wb.addWorksheet('PHAN BO PO');
    head(ap, ['Hóa đơn', 'Item', 'Mã Material', 'PO được chia', 'SL chia', 'SL của cả dòng', 'Kết quả'],
      [20, 34, 18, 16, 14, 16, 22]);
    for (const a of all) {
      for (const l of a.lines) {
        if ((l.alloc || []).length < 2) continue;
        for (const x of l.alloc) {
          ap.addRow([invLabel(a), l.it.code, x.material || '', x.po, x.qty, l.base, l.status]);
          paint(ap, 'FFFFF2CC');
        }
      }
    }
  }

  if (all.some((a) => a.lines.some((l) => (l.lots || []).length))) {
    const lo = wb.addWorksheet('CHI TIET LO');
    head(lo, ['Hóa đơn', 'PO SAP', 'Item', 'Lô (Lot/Batch)', 'Số cây', 'Số lượng lô', 'Đơn vị', 'SL hóa đơn', 'Kết quả'],
      [20, 13, 30, 18, 9, 14, 8, 13, 22]);
    for (const a of all) {
      for (const l of a.lines) {
        for (const L of (l.lots || [])) {
          lo.addRow([invLabel(a), l.sapPo, l.it.code, L.lot, L.rolls, L.qty, L.unit || '', l.base, l.status]);
          paint(lo, isOkK(l.status) ? null : 'FFFFF2CC');
        }
      }
    }
  }

  const p = wb.addWorksheet('DANH SACH PO');
  head(p, ['PO No.', 'PO No ScaX', 'Hóa đơn', 'SL', 'Thành tiền', 'Đã có inbound'], [16, 22, 30, 12, 15, 14]);
  const m = new Map();
  for (const a of all) {
    for (const l of a.lines) {
      if (!l.sapPo) continue;
      const k = l.sapPo.toUpperCase();
      if (!m.has(k)) m.set(k, { sap: l.sapPo, scax: l.it.poScax || '', invs: new Set(), qty: 0, amt: 0, inb: !!a.g.inb });
      const e = m.get(k);
      e.invs.add(invLabel(a)); e.qty += l.it.qty || 0; e.amt += isNaN(l.it.amount) ? 0 : l.it.amount;
      if (!e.scax && l.it.poScax) e.scax = l.it.poScax;
      if (a.g.inb) e.inb = true;
    }
  }
  [...m.values()].sort((x, y) => x.sap.localeCompare(y.sap))
    .forEach((e) => p.addRow([e.sap, e.scax, [...e.invs].join(', '), e.qty, e.amt, e.inb ? 'có' : 'chưa']));

  for (const a of all) {
    const name = ('HD ' + (a.inv.no || a.inv.invNo || invLabel(a))).slice(0, 28).replace(/[\\\/\?\*\[\]:]/g, '');
    const ws = wb.addWorksheet(name);
    writeInvoiceReport(ws, 1, a.inv, a.lines, a.VAL, a.g, true);
  }
}

/* ================= Hiển thị ================= */
const esc = (s) => String(s == null ? '' : s).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));

function renderAll(all) {
  let h = '';
  const nBad = all.filter((a) => a.VAL.valueBad).length;
  const nNoInb = all.filter((a) => !a.VAL.hasInb).length;
  h += `<div class="sum"><div><span>Hóa đơn</span><b>${all.length}</b></div>
    <div><span>Sai giá trị</span><b class="${nBad ? 'bad' : 'good'}">${nBad}</b></div>
    <div><span>Chưa có inbound</span><b class="${nNoInb ? 'bad' : 'good'}">${nNoInb}</b></div>
    <div><span>Tổng tiền hàng</span><b>${fmt(all.reduce((s, a) => s + (a.VAL.invTotal || 0), 0))}</b></div></div>`;
  for (const a of all) {
    const bad = a.lines.filter((l) => !isOk(l.status));
    const cls = a.VAL.valueBad ? 'bad' : (bad.length || !a.VAL.hasInb ? 'warn' : 'ok');
    h += `<details class="inv ${cls}" ${a.VAL.valueBad || bad.length || a.newSup || a.yuboAlone ? 'open' : ''}>
      <summary><b>${esc(a.inv.invNo)}</b> · ${esc(a.inv.invDate)} · ${a.lines.length} dòng ·
      ${a.VAL.hasInb ? '' : '<span class="tag bad">CHƯA CÓ INBOUND</span> '}
      ${a.g.isFab ? `<span class="tag via">vải · ${esc(a.g.fabName || '')}</span> ` : ''}
      ${a.newSup ? '<span class="tag bad">chủ hàng mới — chưa huấn luyện</span> ' : ''}
      ${(a.g.pkl || a.g.isFab || (a.g.pklx && a.g.pklx.length)) ? '' : '<span class="tag warnt">không có packing list</span> '}
      ${a.VAL.noInvoiceNo ? '<span class="tag warnt">chưa có số HĐ</span> ' : ''}
      ${a.VAL.noFrom ? `<span class="tag warnt">số HĐ lấy từ ${esc(a.VAL.noFrom)}</span> ` : ''}
      ${a.inv.ocr ? '<span class="tag bad">OCR — kiểm tra kỹ</span> ' : ''}
      ${a.VAL.noSerial ? '<span class="tag warnt">thiếu ký hiệu HĐ</span> ' : ''}
      ${a.g.pdfInv ? '<span class="tag good">có hóa đơn GTGT</span> ' : ''}
      ${a.VAL.pdfBad ? `<span class="tag bad">lệch hóa đơn GTGT ${fmt(a.VAL.pdfDiff)}</span> ` : ''}
      ${a.VAL.qtyBad ? `<span class="tag bad">lệch tổng SL ${fmt(a.VAL.wroteQty - a.VAL.docQty)}</span> ` : ''}
      ${a.VAL.valueBad ? '<span class="tag bad">SAI GIÁ TRỊ</span> ' : (bad.length ? `<span class="tag warnt">${bad.length} dòng lệch</span> ` : '<span class="tag good">khớp</span> ')}
      <span class="mono">${fmt(a.VAL.invTotal)}</span></summary>`;
    if (a.g.isGen && a.yuboAlone) h += `<div class="newsup">⚠ Hóa đơn GTGT của <b>J&amp;H Yubo</b> đang đứng một mình — thả thêm <b>file PKL Excel</b> của Yubo (và file inbound) để công cụ đọc theo lô, chia PO và điền đúng Invoice Quantity.</div>`;
    if (a.newSup) h += `<div class="newsup">⚠ Chủ hàng <b>${esc(a.newSup)}</b> chưa được huấn luyện trong công cụ. Kết quả dưới đây do bộ đọc chung tự dò nên có thể sai — <b>kiểm tra kỹ trước khi import</b>, và vui lòng liên hệ ${contactHtml()} kèm bộ chứng từ để được bổ sung vào công cụ.</div>`;
    if (a.VAL.valueBad) {
      h += `<div class="alert"><div class="ttl">⚠ HÓA ĐƠN NÀY SAI GIÁ TRỊ — CẦN KIỂM TRA LẠI</div>
        <div class="lines"><span>Tiền hàng HĐ: <b>${fmt(a.VAL.invTotal)}</b></span><span>Theo inbound: <b>${fmt(a.VAL.inbTotal)}</b></span><span>Lệch: <b>${a.VAL.totalDiff > 0 ? '+' : ''}${fmt(a.VAL.totalDiff)}</b></span></div>
        ${a.VAL.valueLines.map((l) => `<div class="li">• <b>${esc(l.it.code)} / ${esc(l.it.po)}</b>: ${esc(l.note)}</div>`).join('')}</div>`;
    }
    h += '<table><thead><tr><th>Item</th><th>PO</th><th>PO SAP</th><th class="n">SL HĐ</th><th class="n">SL inbound</th><th class="n">SL packing</th><th class="n">Giá HĐ</th><th class="n">Giá inb.</th><th>Kết quả</th><th>Ghi chú</th></tr></thead><tbody>';
    for (const l of a.lines) {
      const ok = isOkK(l.status);
      h += `<tr class="${ok ? '' : 'warn'}"><td>${esc(l.it.code)}</td>
        <td>${esc(l.it.po)}${l.it.poVia ? ` <span class="tag via">${esc(l.it.poVia)}</span>` : ''}</td>
        <td>${esc(l.sapPo || '—')}</td><td class="n">${fmt(l.it.qty)}</td>
        <td class="n">${l.hasInb ? fmt(l.base) : '—'}</td><td class="n">${l.pklTotal ? fmt(l.pklTotal) : '—'}</td>
        <td class="n">${fmt(l.it.price)}</td><td class="n ${l.priceBad ? 'bad' : ''}"><b>${l.inbPrices.map(fmt).join(' / ') || '—'}</b></td>
        <td><span class="tag ${ok ? 'good' : 'bad'}">${esc(l.status)}</span></td><td class="note">${esc(l.note)}</td></tr>`;
    }
    h += '</tbody></table></details>';
  }
  $('#report').innerHTML = h;
}

function renderPoTable(list, unresolved) {
  let h = `<div class="sum"><div><span>Số PO</span><b>${list.length}</b></div><div><span>Hóa đơn</span><b>${STATE.groups.length}</b></div>${unresolved.length ? `<div><span>Không tra được</span><b class="bad">${unresolved.length}</b></div>` : ''}</div>`;
  h += '<table><thead><tr><th>PO No.</th><th>PO No ScaX</th><th>Hóa đơn</th><th class="n">SL</th><th class="n">Thành tiền</th></tr></thead><tbody>';
  for (const e of list) {
    h += `<tr><td><b>${esc(e.sap)}</b></td><td>${esc(e.scax) || '<span class="miss">(ScaF)</span>'}</td>
      <td class="note">${esc([...e.invs].join(', '))}</td><td class="n">${fmt(e.qty)}</td><td class="n">${fmt(e.amount)}</td></tr>`;
  }
  h += '</tbody></table>';
  if (unresolved.length) {
    h += '<div class="hint">Không tra được PO:</div><table class="inner"><thead><tr><th>Hóa đơn</th><th>Item</th><th>Mã PO trên hóa đơn</th></tr></thead><tbody>';
    unresolved.forEach((x) => { h += `<tr><td>${esc(x.invNo)}</td><td>${esc(x.code)}</td><td>${esc(x.po)}</td></tr>`; });
    h += '</tbody></table>';
  }
  $('#report').innerHTML = h;
}

/* ================= wiring ================= */
const zone = $('#dropall'), zoneInput = zone.querySelector('input');
zone.addEventListener('click', () => zoneInput.click());
zoneInput.addEventListener('change', (e) => { acceptFiles(e.target.files); e.target.value = ''; });
['dragenter', 'dragover'].forEach((ev) => zone.addEventListener(ev, (e) => { e.preventDefault(); zone.classList.add('drag'); }));
['dragleave', 'dragend'].forEach((ev) => zone.addEventListener(ev, () => zone.classList.remove('drag')));
document.addEventListener('dragover', (e) => e.preventDefault());
async function onDrop(e) {
  e.preventDefault(); zone.classList.remove('drag');
  if (!e.dataTransfer) return;
  const files = await filesFromDataTransfer(e.dataTransfer);
  acceptFiles(files);
}
zone.addEventListener('drop', onDrop);
document.addEventListener('drop', onDrop);
$('#run').addEventListener('click', run);
$('#runpo').addEventListener('click', runPoList);
$('#clear').addEventListener('click', () => {
  STATE.busy = false; FILE_QUEUE.length = 0;
  STATE.items = []; STATE.groups = []; STATE.po = null; STATE.poIdx = null;
  OUTPUTS.length = 0;
  $('#report').innerHTML = ''; $('#log').innerHTML = ''; $('#dls').innerHTML = ''; $('#groups').innerHTML = ''; $('#detect').textContent = '';
  renderSlots();
});
$('#dlall').addEventListener('click', () => {
  OUTPUTS.forEach((o, i) => setTimeout(() => $('#dlx' + i).click(), i * 400));
});
renderSlots();
