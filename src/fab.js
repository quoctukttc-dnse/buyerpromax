/* =====================================================================
   VẢI — bộ đọc chứng từ Excel cho các chủ hàng vải (Techwork / BLAO / HYU)
   Không dùng DOM, chỉ cần ExcelJS → chạy được cả trong Node để kiểm thử.
   ===================================================================== */

/* ---------- chuẩn hoá ---------- */
const AZ = (s) => String(s == null ? '' : s).toUpperCase().replace(/[^A-Z0-9]/g, '');
const WORDS = (s) => [...new Set((String(s == null ? '' : s).toUpperCase().match(/[A-Z0-9]+/g) || []).filter((w) => w.length >= 3))];

/* PO chuẩn SAP = chữ cái + phần số bù 0 cho đủ 7 chữ số.
   TEC002400 → TEC0002400 · SRV0000400 → SRV0000400 · HYU0001700 → HYU0001700 */
function poSap(raw) {
  const m = String(raw == null ? '' : raw).toUpperCase().match(/([A-Z][A-Z&]{1,7})\s*-?\s*0*(\d{1,10})/);
  if (!m) return String(raw == null ? '' : raw).trim().toUpperCase();
  return m[1] + m[2].padStart(7, '0');
}
const poSame = (a, b) => !!a && !!b && poSap(a) === poSap(b);
/* một ô có thể ghi nhiều PO: "J&H0009000 / J&H0008500" */
const poListOf = (raw) => [...new Set((String(raw == null ? '' : raw).toUpperCase().match(/[A-Z][A-Z&]{1,7}\s*-?\s*\d{4,}/g) || []).map(poSap))];

const UNIT_ALIAS = {
  M: 'M', MT: 'M', MTR: 'M', MTS: 'M', METER: 'M', METERS: 'M', METRE: 'M', METRES: 'M',
  YD: 'YD', YDS: 'YD', YARD: 'YD', YARDS: 'YD', Y: 'YD',
  KG: 'KG', KGS: 'KG', KGM: 'KG', LB: 'LBS', LBS: 'LBS', PCS: 'PCS', PC: 'PCS', PIECE: 'PCS', PIECES: 'PCS',
  PR: 'PAA', PRS: 'PAA', PAIR: 'PAA', PAIRS: 'PAA', PAA: 'PAA', DZ: 'DZ', DOZ: 'DZ', DOZEN: 'DZ', SET: 'SET', SETS: 'SET',
};
const unitKey = (s) => UNIT_ALIAS[AZ(s)] || AZ(s);

/* So màu: hóa đơn và inbound viết khác nhau rất nhiều
   "SILVER SATIN" ↔ "SILVER SATIN-0500552"          (inbound dài hơn)
   "BLACK 19-4201 TSX_NLK107341 SHADE C" ↔ "BLACK 19-4201TSX"   (hóa đơn dài hơn)
   "SOLID11-4002 TSX … Powdered Sugar" ↔ "POWDERED SUGAR 11-4002TSX" (đảo thứ tự)
   Trả về điểm 0–4, 0 = không khớp.                                          */
/* Bỏ phần "phụ" trong tên màu để chỉ còn lõi màu thật:
   "AL5 NIDUS (MATCHING COLOR W F25-LCST-002)" → "AL5 NIDUS"
   "UGW MARINE matching color w F25-LCST-004"  → "UGW MARINE"
   "70V FARINE (Non organic)"                  → "70V FARINE"
   Nếu bỏ hết thì giữ nguyên chuỗi gốc.                                      */
const COLOR_STOP = /\b(MATCHING|MATCH|NON[- ]?ORGANIC|ORGANIC|SHADE|COLOUR|COLOR|SOLID|FABRIC|FAB)\b/;
function colorCore(raw) {
  let t = String(raw == null ? '' : raw).toUpperCase().replace(/[\r\n]+/g, ' ');
  t = t.replace(/\([^)]*\)/g, ' ');                       // chú thích trong ngoặc
  const m = t.search(COLOR_STOP);
  if (m > 0) t = t.slice(0, m);
  t = t.replace(/\b[A-Z]\d{2}-[A-Z]{2,6}-\d+[A-Z0-9]*\b/g, ' ');  // mã style F25-LCST-002
  t = t.replace(/\bW\b/g, ' ').replace(/\s+/g, ' ').trim();
  return t || String(raw == null ? '' : raw).toUpperCase().trim();
}

/* So màu theo thang từ chắc chắn đến mờ nhạt (10 → 0).
   invColor = ô màu trên chứng từ · inbColor = cột Color của inbound
   wide     = thêm mô tả/mã màu của chứng từ, dùng cho phép thử "có chứa"        */
/* phần đứng sau chữ "matching" — mã style mà màu này phải khớp với
   ("…matching color w F25-LCST-004" → F25LCST004). Dùng để phân xử khi lõi màu trùng nhau. */
function colorTail(raw) {
  const t = String(raw == null ? '' : raw).toUpperCase().replace(/[\r\n]+/g, ' ');
  const i = t.search(/\bMATCHING\b|\bMATCH\b/);
  if (i < 0) return '';
  return AZ(t.slice(i).replace(/MATCHING|MATCH|COLOU?R|\bW\b/g, ' '));
}

function colorScore(invColor, inbColor, wide) {
  const a = AZ(invColor), b = AZ(inbColor);
  if (!a || !b) return 0;
  if (a === b) return 10;                                   // giống hệt
  if (a.startsWith(b) || b.startsWith(a)) return 9;         // một bên là tiền tố
  if (a.includes(b) || b.includes(a)) return 8;             // một bên nằm trong bên kia
  const w = AZ(String(wide == null ? '' : wide) + ' ' + String(invColor));
  if (w.includes(b)) return 7;                              // tên màu inbound có trong mô tả
  const ca = AZ(colorCore(invColor)), cb = AZ(colorCore(inbColor));
  if (ca && cb) {
    /* lõi màu trùng nhau → phân xử bằng phần "matching w <style>":
       khớp +1 · khác nhau −2 · một bên có một bên không −1                       */
    const ta = colorTail(invColor), tb = colorTail(inbColor);
    const tie = (ta && tb) ? ((ta === tb || ta.indexOf(tb) >= 0 || tb.indexOf(ta) >= 0) ? 1 : -2) : ((ta || tb) ? -1 : 0);
    if (ca === cb) return 6 + tie;                          // trùng lõi màu
    if (ca.startsWith(cb) || cb.startsWith(ca) || ca.includes(cb) || cb.includes(ca)) return 5 + tie;
  }
  const toks = WORDS(colorCore(inbColor));
  if (!toks.length) return 0;
  const shared = toks.filter((t) => w.includes(t)).length;
  if (shared === toks.length) return 4;                     // đủ mọi từ của lõi màu (kể cả đảo thứ tự)
  if (shared >= 1 && shared / toks.length >= 0.5) return 3;  // quá nửa số từ
  return 0;
}

/* số từ (≥3 ký tự) dùng chung — dùng để ghép lô/cây với dòng hóa đơn */
const shareWords = (a, b) => { const B = WORDS(colorCore(b)); return WORDS(colorCore(a)).filter((w) => B.includes(w)).length; };

/* mã article = phần đầu mô tả, cắt ở " - " (khoảng trắng hai bên) hoặc xuống dòng:
   "DA-DNS2693 Solid" → DA-DNS2693 · "HY-N403420FDY - FAB NY/SP\n85/15…" → HY-N403420FDY */
function artOf(desc) {
  const first = String(desc == null ? '' : desc).split(/[\r\n]/)[0];
  let m = String(first.split(/\s+-\s+|\s{2,}/)[0] || first).trim();
  m = m.replace(/-\s*FAB\b.*$/i, '').replace(/[\s,;]+$/, '');   // "HY-N403420FDY-FAB NY/SP" → HY-N403420FDY
  return m;
}

/* mã article: "DA-DNS2693" ⊂ "DA-DNS2693 Solid Fab NY85 SP15" */
const artHit = (art, ...texts) => {
  const k = AZ(art);
  return !!k && k.length >= 4 && texts.some((t) => AZ(t).includes(k));
};

/* mã tham chiếu của dòng inbound để dò trong chữ của chứng từ (bộ đọc chung):
   Supplier Ref "RC031 Solid" → RC031 · "Slide NC214" → SLIDENC214 · desc "825 SYDNEY ECO Solid Fab…" → 825SYDNEYECO */
const REF_CUT = /\s+(SOLID|FAB|FABRIC|ELA|ELASTIC|STRING|LACE|PADDING|MOULD|WIRE|TAPE|LABEL|TAG|STICKER|RING|SLIDE|HOOK|CUP|JACQUARD|PRINT|MESH|KNIT|WOVEN|ADHESIVE|BUTTON|THREAD|ZIPPER|BOW|CORD|VELCRO|FOAM)\b.*$/i;
function refKeys(r) {
  const out = [];
  const add = (t) => { const k = AZ(String(t == null ? '' : t).replace(REF_CUT, '')); if (k.length >= 4 && /\d/.test(k) && !out.includes(k)) out.push(k); };
  add(r.supRef);
  add(String(r.desc || '').split(/\s+-\s+|,/)[0]);
  const w = String(r.supRef || '').toUpperCase().match(/[A-Z0-9][A-Z0-9\/.#-]*\d[A-Z0-9\/.#-]*/g) || [];
  w.forEach((x) => { const k = AZ(x); if (k.length >= 5 && !out.includes(k)) out.push(k); });
  return out;
}
const refHit = (r, ctx) => { const c = AZ(ctx); return !!c && refKeys(r).some((k) => c.includes(k)); };
/* mã trong Specification / Lapdip (có số, ≥4 ký tự) xuất hiện trong chữ chứng từ */
function codeTokens(t) {
  const u = String(t == null ? '' : t).toUpperCase();
  const a = (u.match(/[A-Z0-9][A-Z0-9\/.#-]*/g) || []).map(AZ);
  const b = (u.match(/[A-Z0-9]+/g) || []);
  return [...new Set(a.concat(b))].filter((k) => k.length >= 4 && /\d/.test(k) && !/^(19|20)\d{2}$/.test(k));
}

/* ---------- ô Excel ----------
   Hai cái bẫy của chứng từ vải: ô rich-text (trả về object, không phải string)
   và ô gộp dọc (dòng dưới của vùng gộp vẫn trả về giá trị của dòng trên, làm
   một dòng hàng bị đếm hai lần). Vì vậy chỉ đọc ô gộp ở đúng dòng "chủ".      */
function cellText(v) {
  if (v == null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number' || typeof v === 'boolean') return String(v);
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : v.toISOString().slice(0, 10);
  if (Array.isArray(v.richText)) return v.richText.map((x) => (x && x.text) || '').join('');
  if (v.text != null) return cellText(v.text);
  if (v.result != null) return cellText(v.result);
  return '';
}
function rawCell(ws, r, c) {
  try {
    const cell = ws.getCell(r, c);
    if (cell.isMerged && cell.master && cell.master.row !== undefined && Number(cell.master.row) !== Number(r)) return null;
    return cell.value;
  } catch (e) { return null; }
}
const T = (ws, r, c) => cellText(rawCell(ws, r, c)).trim();
const r3 = (n) => (isNaN(n) || n == null ? n : Math.round(n * 1000) / 1000);
const V = (ws, r, c) => rawCell(ws, r, c);
function N(ws, r, c) {
  let v = V(ws, r, c);
  if (v != null && typeof v === 'object' && !(v instanceof Date) && v.result != null) v = v.result;
  if (v == null || v === '') return NaN;
  if (typeof v === 'number') return v;
  if (v instanceof Date) return NaN;
  if (typeof v === 'object' && v.result != null) return typeof v.result === 'number' ? v.result : NaN;
  const t = String(v).replace(/[^\d.,\-]/g, '');
  if (!t) return NaN;
  const n = parseFloat(t.replace(/,(?=\d{3}\b)/g, '').replace(',', '.'));
  return isNaN(n) ? NaN : n;
}
const rowText = (ws, r, n = 24) => { let s = ''; for (let c = 1; c <= n; c++) s += ' | ' + T(ws, r, c); return s; };
function findRow(ws, re, max) {
  const lim = Math.min(max || 40, ws.rowCount);
  for (let r = 1; r <= lim; r++) if (re.test(rowText(ws, r))) return r;
  return 0;
}
function colBy(ws, r, re, n = 24) { for (let c = 1; c <= n; c++) if (re.test(T(ws, r, c))) return c; return 0; }

/* ngày: serial Excel · Date · dd/mm/yyyy · "Sep 28th, 2026" */
const MONTHS = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
const dmy = (d, m, y) => `${String(d).padStart(2, '0')}.${String(m).padStart(2, '0')}.${y}`;
function fromSerial(n) {
  const ms = Date.UTC(1899, 11, 30) + Math.round(n) * 86400000;
  const d = new Date(ms);
  return dmy(d.getUTCDate(), d.getUTCMonth() + 1, d.getUTCFullYear());
}
function parseDateCell(raw, yearHint) {
  if (raw == null || raw === '') return '';
  if (raw instanceof Date) return dmy(raw.getUTCDate(), raw.getUTCMonth() + 1, raw.getUTCFullYear());
  if (typeof raw === 'number') return raw > 20000 && raw < 80000 ? fromSerial(raw) : '';
  const s = String(raw);
  let m = s.match(/ng[àa]y\s*(\d{1,2})\s*th[áa]ng\s*(\d{1,2})\s*n[ăa]m\s*(\d{4})/i);   // Ngày 6 tháng 10 Năm 2026 (PKL Yubo)
  if (m) return dmy(m[1], m[2], m[3]);
  m = s.match(/(\d{1,2})\s*[\/\-.]\s*(\d{1,2})\s*[\/\-.]\s*(\d{4})/);
  if (m) return (Number(m[2]) > 12 && Number(m[1]) <= 12) ? dmy(m[2], m[1], m[3]) : dmy(m[1], m[2], m[3]);   // 09/17/2026 kiểu Mỹ (tháng/ngày/năm)
  m = s.match(/(\d{4})\s*[\/\-.]\s*(\d{1,2})\s*[\/\-.]\s*(\d{1,2})/);
  if (m) return dmy(m[3], m[2], m[1]);
  m = s.toUpperCase().match(/([A-Z]{3,9})\.?\s*(\d{1,2})\s*(?:ST|ND|RD|TH)?\s*[,，]?\s*(\d{4})?/);
  if (m) {
    const mi = MONTHS.findIndex((x) => m[1].startsWith(x));
    if (mi >= 0) {
      const y = m[3] || yearHint;
      if (y) return dmy(m[2], mi + 1, y);
    }
  }
  if (/^\d{5}$/.test(s)) return fromSerial(Number(s));
  return '';
}
/* tìm một năm hoặc serial ngày ở bất cứ đâu trong workbook để bù cho ngày thiếu năm */
function yearHintOf(wb) {
  for (const ws of wb.worksheets) {
    for (let r = 1; r <= Math.min(12, ws.rowCount); r++) {
      for (let c = 1; c <= 18; c++) {
        const v = V(ws, r, c);
        if (typeof v === 'number' && v > 40000 && v < 80000) return Number(fromSerial(v).slice(-4));
        if (v instanceof Date) return v.getUTCFullYear();
        const m = String(T(ws, r, c)).match(/\b(20\d{2})\b/);
        if (m) return Number(m[1]);
      }
    }
  }
  return new Date().getUTCFullYear();
}

/* ---------- nhận diện chủ hàng ---------- */
function fabProfile(wb) {
  const names = wb.worksheets.map((w) => String(w.name).trim().toUpperCase());
  let t = '';
  for (const ws of wb.worksheets) for (let r = 1; r <= Math.min(30, ws.rowCount); r++) t += rowText(ws, r);
  t = t.toUpperCase();
  if (/QTY\s*\/\s*M/.test(t) && /SURCHARGE/.test(t) && names.includes('PACKING LIST')) return 'TECHWORK';
  if (/11\.\s*PO NUMBER/.test(t) && /QUANTITY\s*\/\s*UNIT/.test(t)) return 'BLAO';
  if (/MARKS\s*&\s*NOS/.test(t) && /INV\.?\s*NO/.test(t)) return 'HYU';
  if (/SCAVI\s*CODE/.test(t) && /PACKING LIST/.test(t)) return 'YUBO';
  if (/CELEB\s*TEXTILES/.test(t) && /INVOICE\s*NO/.test(t) && /PO\s*NO\./.test(t) && /DESCRIPTION OF GOODS/.test(t)) return 'CELEB';
  if (/CHEUNG\s*HING/.test(t) && /INVOICE/.test(t) && /UNIT\s*PRICE/.test(t) && /DESCRIPTION/.test(t) && !/PACKING\s*LIST/.test(t)) return 'CHEUNGHING';
  return '';
}

/* =====================  TECHWORK  ===================== */
function readTechwork(wb) {
  const wi = wb.worksheets.find((w) => /^invoice$/i.test(String(w.name).trim())) || wb.worksheets[0];
  const wp = wb.worksheets.find((w) => /packing/i.test(String(w.name)));
  const hr = findRow(wi, /PO\s*NO\./i) || 14;
  const C = {
    po: colBy(wi, hr, /^PO\s*NO/i), desc: colBy(wi, hr, /DESCRIPTION/i), art: colBy(wi, hr, /^ITEM$/i),
    ccode: colBy(wi, hr, /COLOUR\s*CODE|COLOR\s*CODE/i), qty: colBy(wi, hr, /^QTY/i),
    price: colBy(wi, hr, /PRICE/i), sur: colBy(wi, hr, /SURCHARGE/i), amt: colBy(wi, hr, /AMOUNT/i),
  };
  C.color = 0;
  for (let c = 1; c <= 24; c++) if (/^COLOUR$|^COLOR$/i.test(T(wi, hr, c))) { C.color = c; break; }
  if (!C.color) C.color = colBy(wi, hr, /COLOUR|COLOR/i);
  const unit = unitKey((T(wi, hr, C.qty).match(/\/\s*([A-Za-z]+)/) || [])[1] || 'M');

  const rNo = findRow(wi, /INVOICE\s*NO/i);
  const cNo = rNo ? colBy(wi, rNo, /INVOICE\s*NO/i) : 0;
  const cDt = rNo ? colBy(wi, rNo, /INVOICE\s*DATE/i) : 0;
  const no = rNo && cNo ? T(wi, rNo + 1, cNo) : '';
  const invDate = rNo && cDt ? parseDateCell(V(wi, rNo + 1, cDt), yearHintOf(wb)) : '';

  const items = [];
  let lastPo = '', total = NaN, totalQty = NaN;
  for (let r = hr + 1; r <= wi.rowCount; r++) {
    const first = T(wi, r, 1);
    if (/^total/i.test(first)) { total = N(wi, r, C.amt); totalQty = N(wi, r, C.qty); break; }
    const qty = N(wi, r, C.qty);
    if (isNaN(qty) || qty <= 0) continue;
    if (first) lastPo = first;
    const art = C.art ? T(wi, r, C.art) : '';
    const color = C.color ? T(wi, r, C.color) : '';
    items.push({
      poRaw: lastPo, po: poSap(lastPo), article: art,
      colorText: color, colorCode: C.ccode ? T(wi, r, C.ccode) : '',
      colorShort: color.replace(/[\r\n]+/g, ' ').trim(),
      desc: [art, color, C.desc ? T(wi, r, C.desc) : ''].filter(Boolean).join(' · '),
      qtyByUnit: { [unit]: qty }, unit, qty,
      price: N(wi, r, C.price), surcharge: C.sur ? (N(wi, r, C.sur) || 0) : 0,
      amount: N(wi, r, C.amt),
      code: art || color,
    });
  }
  /* packing list: PO# | Article No. | Colour | Lot No. | Roll | Rolls # | Ttl/M … */
  const groups = [];
  if (wp) {
    const pr = findRow(wp, /PO\s*#/i) || 12;
    const P = {
      po: colBy(wp, pr, /PO\s*#/i), art: colBy(wp, pr, /ARTICLE/i), color: colBy(wp, pr, /COLOUR|COLOR/i),
      lot: colBy(wp, pr, /LOT/i), roll: colBy(wp, pr, /ROLLS\s*#/i), qty: colBy(wp, pr, /TTL/i),
      nw: colBy(wp, pr, /N\.?W/i), gw: colBy(wp, pr, /G\.?W/i),
    };
    let cp = '', ca = '', cc = '', cl = '';
    for (let r = pr + 1; r <= wp.rowCount; r++) {
      if (/^total/i.test(T(wp, r, 1))) break;
      const q = N(wp, r, P.qty);
      if (isNaN(q) || q <= 0) continue;
      cp = T(wp, r, P.po) || cp; ca = T(wp, r, P.art) || ca;
      cc = T(wp, r, P.color) || cc; cl = (P.lot ? T(wp, r, P.lot) : '') || cl;
      const key = poSap(cp) + '|' + AZ(ca) + '|' + AZ(cc) + '|' + AZ(cl);
      let g = groups.find((x) => x.key === key);
      if (!g) { g = { key, po: poSap(cp), poRaw: cp, article: ca, color: cc, lot: cl, unit, rolls: [], total: 0 }; groups.push(g); }
      g.rolls.push({ no: P.roll ? T(wp, r, P.roll) : String(g.rolls.length + 1), qty: q, nw: N(wp, r, P.nw), gw: N(wp, r, P.gw) });
      g.total = r3(g.total + q);
    }
  }
  return {
    profile: 'TECHWORK', supplier: 'Fujian Techwork',
    inv: { no, invNo: no, invDate, items, currency: 'USD', unitDefault: unit, surchargeHeader: 0, total, totalQty, amountInclSur: true },
    pkl: groups.length ? { groups, unit, level: 'lot' } : null,
  };
}

/* =====================  BLAO / NEW STYLE  ===================== */
function readBlao(wb) {
  const wi = wb.worksheets.find((w) => /^inv$/i.test(String(w.name).trim()))
    || wb.worksheets.find((w) => /INV/i.test(String(w.name)));
  const wp = wb.worksheets.find((w) => /PACKING/i.test(String(w.name)));
  const rLab = findRow(wi, /No\s*&\s*Date of invoice/i, 20);
  let no = '', invDate = '';
  if (rLab) {
    const c0 = colBy(wi, rLab, /No\s*&\s*Date of invoice/i);
    for (let c = c0; c <= c0 + 8; c++) {
      const v = V(wi, rLab + 1, c), t = T(wi, rLab + 1, c);
      if (!t) continue;
      const d = parseDateCell(v instanceof Date || typeof v === 'number' ? v : t);
      if (d) { invDate = d; } else if (!no) { no = t; }
    }
  }
  const hr = findRow(wi, /PO\s*Number/i) || 23;
  const C = {
    po: colBy(wi, hr, /PO\s*Number/i), desc: colBy(wi, hr, /Description/i), color: colBy(wi, hr, /Colo/i),
    qty: colBy(wi, hr, /Quantity/i), price: colBy(wi, hr, /Unit\s*-?\s*Price/i), amt: colBy(wi, hr, /Amount/i),
  };
  const items = [];
  let surchargeHeader = 0, total = NaN, totalQty = NaN;
  for (let r = hr + 1; r <= wi.rowCount; r++) {
    const line = rowText(wi, r);
    if (/SURCHARGE/i.test(line)) { const s = N(wi, r, C.amt); if (!isNaN(s)) surchargeHeader += s; continue; }
    if (/TOTAL\s*:/i.test(line)) { total = N(wi, r, C.amt); if (isNaN(totalQty)) totalQty = N(wi, r, C.qty); continue; }
    const po = T(wi, r, C.po);
    const q = N(wi, r, C.qty);
    const px = N(wi, r, C.price);
    if (po && !isNaN(q) && !isNaN(px)) {
      const q2 = N(wi, r + 1, C.qty);
      const color = T(wi, r, C.color);
      items.push({
        poRaw: po, po: poSap(po), article: artOf(T(wi, r, C.desc)),
        colorText: color, colorCode: '',
        colorShort: color.replace(/[\r\n]+/g, ' ').trim(),
        desc: [T(wi, r, C.desc), color].filter(Boolean).join(' · '),
        qtyByUnit: isNaN(q2) ? { KG: q } : { KG: q, M: q2 }, unit: 'KG', qty: q,
        price: px, surcharge: 0, amount: N(wi, r, C.amt),
        code: color || T(wi, r, C.desc),
      });
    }
  }
  /* PACKING: tổng theo màu (Net-weight = số kg → dùng để tự kiểm đơn vị) */
  const groups = [];
  if (wp) {
    const pr = findRow(wp, /PO\s*Number/i) || 23;
    const P = {
      po: colBy(wp, pr, /PO\s*Number/i), desc: colBy(wp, pr, /Description/i), color: colBy(wp, pr, /Colo/i),
      qty: colBy(wp, pr, /Quantity/i), nw: colBy(wp, pr, /Net/i), gw: colBy(wp, pr, /Gross/i),
    };
    for (let r = pr + 1; r <= wp.rowCount; r++) {
      if (/TOTAL\s*:/i.test(rowText(wp, r))) break;
      const po = T(wp, r, P.po), q = N(wp, r, P.qty);
      if (!po || isNaN(q)) continue;
      const color = T(wp, r, P.color);
      groups.push({
        key: poSap(po) + '|' + AZ(color), po: poSap(po), poRaw: po,
        article: artOf(T(wp, r, P.desc)),
        color, lot: '', unit: 'KG', rolls: [], total: q,
        nw: N(wp, r, P.nw), rollCount: N(wp, r, P.qty - 1),
      });
    }
  }
  /* các sheet lô (Y92296…) cho chi tiết từng cây */
  const lots = [];
  for (const ws of wb.worksheets) {
    if (!/^Y\d{3,}/i.test(String(ws.name).trim())) continue;
    const rPo = findRow(ws, /order\s*No\./i, 12);
    const po = rPo ? T(ws, rPo, colBy(ws, rPo, /order\s*No\./i) + 1) || T(ws, rPo, 3) : '';
    const rCol = findRow(ws, /Color\s*&|Color.*No\./i, 12);
    let color = '';
    if (rCol) {
      const c0 = colBy(ws, rCol, /Color\s*&|Color.*No\./i);
      for (let c = c0 + 1; c <= c0 + 8; c++) { const v = T(ws, rCol, c); if (v) { color = v; break; } }
    }
    const hr2 = findRow(ws, /Roll\s*No\./i, 14);
    const cQ = hr2 ? colBy(ws, hr2, /Quantity\s*\(kg\)/i) : 0;
    const cR = hr2 ? colBy(ws, hr2, /Roll\s*No\./i) : 0;
    const g = { key: poSap(po) + '|' + AZ(color) + '|' + AZ(ws.name), po: poSap(po), poRaw: po, article: '', color, lot: String(ws.name).trim(), unit: 'KG', rolls: [], total: 0 };
    if (cQ) {
      for (let r = hr2 + 1; r <= ws.rowCount; r++) {
        const t1 = rowText(ws, r);
        if (/sub-?total|^\s*\|\s*Total/i.test(t1)) break;
        const q = N(ws, r, cQ);
        if (isNaN(q) || q <= 0) continue;
        g.rolls.push({ no: cR ? T(ws, r, cR) : String(g.rolls.length + 1), qty: q });
        g.total = r3(g.total + q);
      }
    }
    if (g.rolls.length) lots.push(g);
  }
  return {
    profile: 'BLAO', supplier: 'New Style Vietnam',
    inv: { no, invNo: no, invDate, items, currency: 'VND', unitDefault: 'KG', surchargeHeader, total, totalQty, amountInclSur: false },
    pkl: groups.length || lots.length ? { groups, lots, unit: 'KG', level: lots.length ? 'lot' : 'color' } : null,
  };
}

/* =====================  HYU / QUANZHOU HENGYU  ===================== */
function readHyu(wb) {
  const wi = wb.worksheets.find((w) => /^invoice$/i.test(String(w.name).trim()))
    || wb.worksheets.find((w) => /invoice/i.test(String(w.name)));
  const wd = wb.worksheets.find((w) => /DETAIL PACKING/i.test(rowText(w, 1) + rowText(w, 2)))
    || wb.worksheets.find((w) => /码单/.test(String(w.name)));
  const wp = wb.worksheets.find((w) => /^packing list$/i.test(String(w.name).trim()));
  const yh = yearHintOf(wb);

  const rNo = findRow(wi, /Inv\.?\s*No/i, 12);
  let no = '', invDate = '';
  if (rNo) {
    for (let c = 1; c <= 20; c++) {
      const t = T(wi, rNo, c);
      if (!t) continue;
      const m = t.match(/Inv\.?\s*No\.?\s*[:：]?\s*([A-Z0-9\-\/]+)/i);
      if (m && !no) no = m[1];
      if (/date/i.test(t)) invDate = parseDateCell(t.replace(/.*date\s*[:：]?/i, ''), yh) || invDate;
    }
  }
  const hr = findRow(wi, /Marks\s*&\s*Nos/i) || 12;
  const C = {
    po: colBy(wi, hr, /PO\s*No/i), desc: colBy(wi, hr, /Description/i), pkg: colBy(wi, hr, /Package/i),
    qty: colBy(wi, hr, /Quantity/i), price: colBy(wi, hr, /Unit\s*Price/i), amt: colBy(wi, hr, /TOTAL\s*VALUE|Amount/i),
  };
  let unit = unitKey(T(wi, hr + 1, C.qty) || 'YD');
  if (!['M', 'YD', 'KG', 'LBS'].includes(unit)) unit = 'YD';
  const items = [];
  let total = NaN, totalQty = NaN;
  for (let r = hr + 1; r <= wi.rowCount; r++) {
    if (/^total/i.test(T(wi, r, 1)) || /^total/i.test(T(wi, r, C.po))) { total = N(wi, r, C.amt); totalQty = N(wi, r, C.qty); break; }
    const q = N(wi, r, C.qty), p = N(wi, r, C.price);
    if (isNaN(q) || q <= 0 || isNaN(p)) continue;
    const po = T(wi, r, C.po) || no;
    const desc = T(wi, r, C.desc);
    const art = artOf(desc);
    items.push({
      poRaw: po, po: poSap(po), article: art, colorText: (desc.split(/[\r\n]/).map((x) => x.trim()).filter(Boolean).pop() || desc), colorCode: '',
      colorShort: (desc.split(/[\r\n]/).map((x) => x.trim()).filter(Boolean).pop() || '').slice(0, 40),
      desc, qtyByUnit: { [unit]: q }, unit, qty: q,
      price: p, surcharge: 0, amount: N(wi, r, C.amt),
      code: art || desc.slice(0, 24),
    });
  }
  /* 码单 — chi tiết từng cây: Color | Composition | Roll/No. | G.W | N.W | Yard | LOT# */
  const groups = [];
  if (wd) {
    const hr2 = findRow(wd, /Roll\s*\/?\s*No/i, 20) || 8;
    const P = {
      color: colBy(wd, hr2, /^Color/i) || 1, comp: colBy(wd, hr2, /Composition/i),
      roll: colBy(wd, hr2, /Roll/i), gw: colBy(wd, hr2, /G\.?W/i), nw: colBy(wd, hr2, /N\.?W/i),
      qty: colBy(wd, hr2, /Yard|Meter|Mtr/i), lot: colBy(wd, hr2, /LOT/i),
    };
    const rPo = findRow(wd, /PO\s*#/i, 12);
    let poD = '';
    if (rPo) { const t = rowText(wd, rPo).match(/PO\s*#\s*([A-Z0-9]+)/i); if (t) poD = t[1]; }
    let cc = '', cl = '', ca = '';
    for (let r = hr2 + 1; r <= wd.rowCount; r++) {
      const roll = T(wd, r, P.roll);
      if (/rolls?$/i.test(roll)) break;
      const q = N(wd, r, P.qty);
      if (isNaN(q) || q <= 0) continue;
      cc = T(wd, r, P.color) || cc; ca = (P.comp ? T(wd, r, P.comp) : '') || ca;
      cl = (P.lot ? T(wd, r, P.lot) : '') || cl;
      const key = poSap(poD) + '|' + AZ(cc) + '|' + AZ(cl);
      let g = groups.find((x) => x.key === key);
      if (!g) {
        g = { key, po: poSap(poD), poRaw: poD, article: artOf(ca), color: cc, lot: cl, unit, rolls: [], total: 0 };
        groups.push(g);
      }
      g.rolls.push({ no: roll || String(g.rolls.length + 1), qty: q, nw: N(wd, r, P.nw), gw: N(wd, r, P.gw) });
      g.total += q;
    }
  }
  if (!invDate && wp) {
    const r2 = findRow(wp, /Date/i, 12);
    if (r2) for (let c = 1; c <= 20; c++) { const d = parseDateCell(V(wp, r2, c), yh) || parseDateCell(T(wp, r2, c), yh); if (d) { invDate = d; break; } }
  }
  return {
    profile: 'HYU', supplier: 'Quanzhou Hengyu',
    inv: { no, invNo: no, invDate, items, currency: 'USD', unitDefault: unit, surchargeHeader: 0, total, totalQty, amountInclSur: false },
    pkl: groups.length ? { groups, unit, level: 'lot' } : null,
  };
}

/* =====================  YUBO (J&H Yubo) — packing list kèm đơn giá  =====================
   Chứng từ là một bảng PKL nội bộ: mỗi dòng là MỘT LÔ (Lot) của một mã vải + màu.
   Nhiều dòng cùng (SCAVI CODE + màu) sẽ gộp lại thành một dòng đối chiếu, vì trong
   file inbound chúng là cùng một dòng PO.  Ô PO ghi nhiều PO: "J&H0009000 / J&H0008500".  */
function readYubo(wb) {
  const ws = wb.worksheets.find((w) => /^PKL$/i.test(String(w.name).trim()))
    || wb.worksheets.find((w) => findRow(w, /SCAVI\s*CODE/i, 30));
  const hr = findRow(ws, /SCAVI\s*CODE/i, 30);
  const C = {
    name: colBy(ws, hr, /Item.*name/i), code: colBy(ws, hr, /SCAVI\s*CODE/i),
    po: colBy(ws, hr, /^PO$/i) || colBy(ws, hr, /\bPO\b/i), color: colBy(ws, hr, /Colo/i),
    lot: colBy(ws, hr, /^Lot$/i) || colBy(ws, hr, /\bLot\b/i),
    roll: colBy(ws, hr, /ROLL/i), qty: colBy(ws, hr, /C[âa]n n[ặa]ng|\(KG\)/i),
    gw: colBy(ws, hr, /Gross/i), price: colBy(ws, hr, /Đ[ơo]n gi[áa]|Price/i),
    amt: colBy(ws, hr, /AMOUNT/i),
  };
  /* cột bình luận của người kiểm (tên cột có chữ "comment") — chỉ dùng để nhắc, không lấy số */
  let cCmt = 0;
  for (let c = 1; c <= 24; c++) if (/comment/i.test(T(ws, hr, c))) { cCmt = c; break; }

  /* số & ngày hóa đơn: nhãn "HD:" và "Ngày xuất HD:" ở góc trên */
  let no = '', invDate = '';
  const rHd = findRow(ws, /\bHD\s*:/i, 12);
  if (rHd) {
    for (let r = 1; r <= Math.min(12, ws.rowCount); r++) {
      for (let c = 1; c <= 24; c++) {
        const t = T(ws, r, c);
        if (/Ng[àa]y xu[ấa]t\s*HD/i.test(t)) {
          for (let k = c + 1; k <= c + 4; k++) { const d = parseDateCell(V(ws, r, k)) || parseDateCell(T(ws, r, k)); if (d) { invDate = d; break; } }
        } else if (/^HD\s*:?$/i.test(t) && !no) {
          for (let k = c + 1; k <= c + 4; k++) { const v = T(ws, r, k); if (v) { no = v; break; } }
        }
      }
    }
  }
  /* PKL không có ô "HD:" (mẫu 10.2026): lấy ngày ở dòng "Ngày 6 tháng 10 Năm 2026" phía trên bảng để ghép với hoá đơn GTGT PDF */
  if (!invDate) {
    for (let r = 1; r <= Math.min(hr - 1, 12) && !invDate; r++) for (let c = 1; c <= 24; c++) { const d = parseDateCell(T(ws, r, c)); if (d) { invDate = d; break; } }
  }

  const map = new Map();
  let total = NaN, totalVat = NaN, totalQty = NaN;
  for (let r = hr + 1; r <= ws.rowCount; r++) {
    if (/T[ổo]ng c[ộo]ng/i.test(rowText(ws, r))) {
      total = N(ws, r, C.amt);
      totalVat = N(ws, r, C.amt + 1);
      totalQty = N(ws, r, C.qty);
      break;
    }
    const q = N(ws, r, C.qty), code = T(ws, r, C.code);
    if (!code || isNaN(q) || q <= 0) continue;
    const color = T(ws, r, C.color);
    const key = AZ(code) + '|' + AZ(color);
    let it = map.get(key);
    if (!it) {
      it = {
        poRaw: T(ws, r, C.po), poList: poListOf(T(ws, r, C.po)), po: poListOf(T(ws, r, C.po))[0] || poSap(T(ws, r, C.po)),
        article: code, colorText: color, colorShort: color.replace(/[\r\n]+/g, ' ').trim(), colorCode: '',
        desc: [T(ws, r, C.name), code, color].filter(Boolean).join(' · '),
        qtyByUnit: { KG: 0 }, unit: 'KG', qty: 0,
        price: N(ws, r, C.price), surcharge: 0, amount: 0,
        code, lots: [], warn: [], pklKey: key,
      };
      map.set(key, it);
    }
    it.qty = r3(it.qty + q);
    it.qtyByUnit.KG = it.qty;
    const a = N(ws, r, C.amt);
    if (!isNaN(a) && a > 0) it.amount = r3((it.amount || 0) + a);
    else {
      /* dòng không có thành tiền: hàng FOC (miễn phí, "FOC trong roll 8") — vẫn nhận hàng nên vẫn tính
         vào số lượng, nhưng KHÔNG tính tiền khi so thành tiền hóa đơn ↔ PO */
      const foc = /\bFOC\b|free\s*of\s*charge|mi[ễe]n\s*ph[íi]/i.test(rowText(ws, r));
      it.unpaidQty = r3((it.unpaidQty || 0) + q);
      it.unpaidNote = (it.unpaidNote ? it.unpaidNote + '; ' : '')
        + `lô ${C.lot ? T(ws, r, C.lot) : 'dòng ' + r} ${q} KG ${foc ? 'FOC (miễn phí)' : 'không ghi thành tiền'}`;
      if (!foc) it.warn.push(`dòng ${r}: ${q} KG không ghi thành tiền (không có chữ FOC) — kiểm lại`);
    }
    const px = N(ws, r, C.price);
    if (!isNaN(px) && isNaN(it.price)) it.price = px;
    const lot = C.lot ? T(ws, r, C.lot) : '';
    it.lots.push({ lot: lot || `dòng ${r}`, rolls: C.roll ? (N(ws, r, C.roll) || 0) : 0, qty: q, unit: 'KG' });
    if (cCmt) {
      const cm = N(ws, r, cCmt);
      if (!isNaN(cm) && Math.abs(cm - q) > 0.01) it.warn.push(`lô ${lot}: cột "${T(ws, hr, cCmt)}" ghi ${cm} ≠ ${q}`);
    }
  }
  const items = [...map.values()];
  const groups = items.map((it) => ({
    key: AZ(it.article) + '|' + AZ(it.colorText), po: '', poRaw: it.poRaw, article: it.article,
    color: it.colorText, lot: '', unit: 'KG', rolls: [], total: it.qty,
  }));
  const lots = [];
  items.forEach((it) => it.lots.forEach((L) => lots.push({
    key: AZ(it.article) + '|' + AZ(it.colorText) + '|' + AZ(L.lot), po: '', article: it.article,
    color: it.colorText, lot: L.lot, unit: 'KG',
    rolls: new Array(Math.max(1, Math.round(L.rolls || 1))).fill(0).map((x, i) => ({ no: String(i + 1), qty: 0 })),
    total: L.qty,
  })));
  return {
    profile: 'YUBO', supplier: 'J&H Yubo',
    inv: {
      no, invNo: no, invDate, items, currency: 'VND', unitDefault: 'KG',
      surchargeHeader: 0, total, totalVat, totalQty, amountInclSur: false, noInvoiceNo: !no,
      noSerial: !!no && /^\d{1,8}$/.test(String(no).trim()),
    },
    pkl: { groups, lots, unit: 'KG', level: 'lot' },
  };
}

/* =====================  CAPITAL TRICOT (Thái Lan) — chứng từ PDF  =====================
   Một file PDF: trang đầu INVOICE, các trang sau PACKING LIST ghi theo KIỆN (bale), mỗi dòng một cây.
   Hóa đơn (mỗi dòng một PO + design + màu + mã L/D):
     "PO.CAP0022700 N.295/207/73 Usable width 60"/Full width:62" (44GSM) MISTY ROSE 131/24I 309 YDS. 2.81 USD/YD USD 868.29"
     dòng kế tiếp cùng PO không ghi lại PO: "N.295L W SOFT Usable width 72"/… (43GSM) MISTY ROSE 298/24C 515 YDS. 1.75 USD/YD USD 901.25"
   Packing list (design / PO / L/D ghi một lần ở các dòng đầu kiện rồi GIỮ cho các cây kế tiếp, kể cả sang kiện / trang sau):
     "1 N.295/207/73 60" MISTY ROSE 1 80.00 X 1 5.40 5.87"  ·  "PO.CAP0022700 MISTY ROSE 2 100.00 X 1 6.75 7.22"
     "L/D 131/24I MISTY ROSE 3 129.00 X 1 8.66 9.13"         ·  "309.00 20.81 22.22" (dòng cộng kiện — bỏ)
   Khớp inbound: PO + design (Supplier Mat. No. "N.295/207/73 Solid") + màu (Color) ; L/D ↔ Lapdip Color.
   Nhóm packing list = PO | design | màu | L/D | kiện → nhiều kiện của cùng một dòng hóa đơn (kiện 5 + 6 = 300 + 304) cộng lại. */
const isCapitalText = (t) => /CAPITAL\s*TRICOT/i.test(t);
/* "N.295L w soft 72"" → "N.295L W SOFT" (bỏ khổ vải) */
const capDesign = (t) => String(t == null ? '' : t).toUpperCase().replace(/\d+(?:\.\d+)?\s*"/g, ' ').replace(/\s+/g, ' ').trim();
const capLd = (t) => String(t == null ? '' : t).toUpperCase().replace(/\s+/g, '');
function readCapital(lines) {
  const all = (lines || []).map((x) => String(x == null ? '' : x).replace(/\s+/g, ' ').trim());
  const t = all.join('\n');
  if (!isCapitalText(t) || !/INVOICE/i.test(t)) return null;
  /* số + ngày hóa đơn: "DELIVERY TO : CT-26-314T AUGUST 19, 2026" (dòng ngay dưới "INVOICE NO. DATE") */
  let no = '', invDate = '';
  for (let i = 0; i < all.length; i++) {
    if (!/INVOICE\s*NO\.?\s+DATE/i.test(all[i])) continue;
    for (let j = i + 1; j <= Math.min(all.length - 1, i + 3) && !no; j++) {
      const m = all[j].match(/\b([A-Z]{1,4}-\d{2}-\d{2,6}[A-Z]?)\b\s*(.*)$/);
      if (m) { no = m[1]; invDate = parseDateCell(m[2]) || ''; }
    }
    if (no) break;
  }
  if (!no) { const m = t.match(/\b([A-Z]{1,4}-\d{2}-\d{2,6}[A-Z]?)\b/); if (m) no = m[1]; }

  /* ---- dòng hàng trên hóa đơn ---- */
  const RE_ITEM = /^(?:PO\.?\s*([A-Z][A-Z&]{1,5}\s*-?\s*\d{5,})\s+)?(.*?)\s+([\d,]+(?:\.\d+)?)\s*(YDS?|YARDS?|MTRS?|M|KGS?|PCS)\.?\s+([\d,]*\.?\d+)\s*(?:USD|US\$)\s*\/\s*(?:YD|M|KG|PC)S?\.?\s+(?:USD|US\$)\s*([\d,]+(?:\.\d+)?)\s*$/i;
  const items = [];
  let lastPo = '', total = NaN, totalQty = NaN, inPkl = false;
  for (let i = 0; i < all.length; i++) {
    const L = all[i];
    if (/^PACKING\s*LIST\b/i.test(L)) { inPkl = true; break; }
    if (/^TOTAL\b/i.test(L) && items.length && isNaN(total)) {
      const mq = L.match(/([\d,]+(?:\.\d+)?)\s*(?:YDS?|YARDS?|MTRS?|M|KGS?|PCS)\b/i);
      const ma = L.match(/(?:USD|US\$)\s*([\d,]+(?:\.\d+)?)\s*$/i);
      if (mq) totalQty = Number(mq[1].replace(/,/g, ''));
      if (ma) total = Number(ma[1].replace(/,/g, ''));
      continue;
    }
    const m = L.match(RE_ITEM);
    if (!m) continue;
    if (m[1]) lastPo = m[1].replace(/\s+/g, '');
    if (!lastPo) continue;
    const unit = unitKey(m[4]);
    const qty = Number(m[3].replace(/,/g, ''));
    const price = Number(m[5].replace(/,/g, ''));
    const amount = Number(m[6].replace(/,/g, ''));
    /* phần mô tả: "<design> Usable width 60"/Full width:62" (44GSM) <MÀU> <L/D>" */
    let body = m[2];
    let design = body, tail = '';
    const w = body.search(/\b(?:USABLE|FULL)\s*WIDTH\b|\(\s*\d+\s*GSM\s*\)/i);
    if (w >= 0) {
      design = body.slice(0, w);
      const after = body.slice(w);
      const g = after.match(/\(\s*\d+\s*GSM\s*\)\s*(.*)$/i);
      tail = g ? g[1] : after.replace(/^.*?"\s*/g, '');
    }
    tail = tail.trim();
    let ld = '', color = tail;
    const ml = tail.match(/^(.*?)\s+(\d{2,5}\s*\/\s*\d{2}\s*[A-Z]?)$/);
    if (ml) { color = ml[1].trim(); ld = capLd(ml[2]); }
    else { const m2 = tail.match(/^(.*?)\s+(\S*\d\S*)$/); if (m2 && /^[A-Z ]+$/.test(m2[1])) { color = m2[1].trim(); ld = capLd(m2[2]); } }
    design = capDesign(design);
    const po = poSap(lastPo);
    items.push({
      poRaw: lastPo, po, article: design, colorText: color, colorCode: ld, colorShort: color,
      desc: [design, color, ld ? 'L/D ' + ld : '', body].filter(Boolean).join(' · '),
      qtyByUnit: { [unit]: qty }, unit, qty, price, surcharge: 0, amount,
      code: design || color, lapdip: ld,
      pklKey: po + '|' + AZ(design) + '|' + AZ(color) + '|' + AZ(ld),
    });
  }
  if (!items.length) return null;
  const unit = items[0].unit || 'YD';

  /* ---- packing list theo kiện ----
     PO / L/D / design của một kiện có thể ghi ở dòng thứ 2-3 (sau cây số 1) → gom hết các cây của kiện rồi mới
     chốt thuộc tính; kiện không ghi gì thì kế thừa kiện trước (kiện 6 tiếp kiện 5, kiện 9 tiếp design của kiện 8).
     Kiện kết thúc ở dòng cộng kiện ("309.00 20.81 22.22") hoặc GRAND TOTAL; sang trang vẫn là cùng kiện. */
  const groups = [];
  if (inPkl) {
    const RE_ROLL = /^(.*?)\s*(\d{1,3})\s+([\d,]+\.\d{1,2})\s+X\s*(\d+)\s+([\d.]+)\s+([\d.]+)\s*$/;
    let prev = { po: '', ld: '', design: '' };
    let cur = null;
    const open = () => { if (!cur) cur = { bale: '', po: '', ld: '', design: '', rolls: [] }; return cur; };
    const flush = () => {
      if (!cur) return;
      const po = cur.po || prev.po, ld = cur.ld || (cur.po ? '' : prev.ld), design = cur.design || prev.design;
      for (const r of cur.rolls) {
        const itemKey = po + '|' + AZ(design) + '|' + AZ(r.color) + '|' + AZ(ld);
        const key = itemKey + '|' + AZ(cur.bale);
        let g = groups.find((x) => x.key === key);
        if (!g) {
          g = { key, itemKey, po, poRaw: po, article: design, color: r.color, ld, lot: cur.bale ? 'Kiện ' + cur.bale : '(không ghi kiện)', unit, rolls: [], total: 0 };
          groups.push(g);
        }
        g.rolls.push({ no: r.no, qty: r.qty, nw: r.nw, gw: r.gw });
        g.total = r3(g.total + r.qty);
      }
      if (cur.rolls.length) prev = { po, ld, design };
      cur = null;
    };
    let started = false;
    for (const L of all) {
      if (!started) { started = /^PACKING\s*LIST\b/i.test(L); continue; }
      if (/^GRAND\s*TOTAL/i.test(L)) { flush(); break; }
      if (/^(?:BALE|NO\.)\b.*(?:DESIGN|YDS)/i.test(L)) continue;         // tiêu đề bảng
      if (/^[\d.,\s]+$/.test(L) && /\./.test(L)) { flush(); continue; }   // dòng cộng kiện
      const m = L.match(RE_ROLL);
      let head = m ? m[1] : L;
      const mb = head.match(/^(\d{1,2})\b\s*(.*)$/);
      let bale = '';
      if (mb && (m || /PO\.|L\s*\/\s*D|^[A-Z]\.?\d/i.test(mb[2]))) { bale = mb[1]; head = mb[2]; }
      let color = '';
      if (m) {
        const mc = head.match(/((?:\b[A-Z]{2,}\b\s*)+)$/);
        if (mc) { color = mc[1].trim(); head = head.slice(0, head.length - mc[0].length).trim(); }
      }
      let mp, po = '', ld = '';
      if ((mp = head.match(/PO\.?\s*([A-Z][A-Z&]{1,5}\s*-?\s*\d{5,})/i))) { po = poSap(mp[1].replace(/\s+/g, '')); head = head.replace(mp[0], ' '); }
      if ((mp = head.match(/L\s*\/\s*D\s*[:.]?\s*(\d{2,5}\s*\/\s*\d{2}\s*[A-Z]?)\b/i))) { ld = capLd(mp[1]); head = head.replace(mp[0], ' '); }
      const d = capDesign(head.replace(/\bX\b/g, ' '));
      const design = d && /^[A-Z]\.?\d/i.test(d) ? d : '';
      if (!m && !bale && !po && !ld && !design) continue;                 // dòng địa chỉ / tiêu đề trang
      const c = open();
      if (bale) c.bale = bale;
      if (po) c.po = po;
      if (ld) c.ld = ld;
      if (design) c.design = design;
      if (!m) continue;
      const q = Number(m[3].replace(/,/g, ''));
      if (!q || isNaN(q) || !color) continue;
      c.rolls.push({ no: m[2], qty: q, nw: Number(m[5]), gw: Number(m[6]), color });
    }
    flush();
  }
  return {
    profile: 'CAPITAL', supplier: 'Capital Tricot',
    inv: { no, invNo: no, invDate, items, currency: 'USD', unitDefault: unit, surchargeHeader: 0, total, totalQty, amountInclSur: true },
    pkl: groups.length ? { groups, unit, level: 'lot' } : null,
  };
}

/* =====================  CARVICO (Ý) — hóa đơn FATTURA/INVOICE PDF + packing list PDF rời  =====================
   Hóa đơn: không ghi PO SAP (chỉ "Order 579"), mỗi nhóm hàng mở đầu bằng dòng
     "00851853 000825 160070 SYDNEY ECO [1]"  → số packing list 851853 · mã article 000825 · tên "SYDNEY ECO"
   rồi mỗi màu một dòng: "WIDTH 160CM G/M2 170 WONDERLAND 1E 03261 MT 632,50 5,45 3.447,13 N1"
     → màu WONDERLAND · mã màu 03261 · MT (mét) · 632,50 · 5,45 · 3.447,13 (số kiểu Ý: chấm nghìn, phẩy lẻ).
   Số/ngày: dòng dưới "NR.DOCUMENTO / DOCUMENT No": "6014 00E BANK TRANSFER 90 DAYS 28493 22/07/26".
   Inbound: Material Description "825 SYDNEY ECO Solid Fab PE100, 170GM2 160CM BLACK #9164" → article "825 SYDNEY ECO",
   Color "BLACK #9164" / "WONDERLAND 03261" / "MYSTIC BLUE # 6063" (mã màu bỏ 0 đầu, có khi giữ) → so màu theo TÊN,
   mã màu để trong desc. Hai dòng hóa đơn cùng màu (BLACK ở 2 packing list) cộng vào cùng một dòng inbound.
   Packing list (file riêng "PKL_851853_-_INV_28493.pdf"): mỗi màu một mục
     "872470 000825 160070 SYDNEY ECO 003261 WONDERLAND" rồi từng cây "074178902 1E 741789 CM016902M 70,20 20,50 [lỗi]"
     (741789 = số lô/batch, chỉ ghi ở cây đầu của lô), "Tot Rolls / batch :" và "Tot Rolls / colour:" là dòng cộng.
   Khóa ghép hóa đơn ↔ packing list: số PKL | article | mã màu (bỏ 0 đầu). */
const isCarvicoText = (t) => /CARVICO/i.test(t);
const itNum = (s) => { const t = String(s == null ? '' : s).trim().replace(/\./g, '').replace(',', '.'); const n = Number(t); return t && !isNaN(n) ? n : NaN; };
const carvArticle = (code, name) => (String(Number(code) || code) + ' ' + String(name || '').replace(/\[\d\]\s*$/, '').trim()).trim();
const stripZero = (s) => String(s == null ? '' : s).trim().replace(/^0+(?=\d)/, '');
function readCarvico(lines) {
  const all = (lines || []).map((x) => String(x == null ? '' : x).replace(/\s+/g, ' ').trim());
  const t = all.join('\n');
  if (!isCarvicoText(t) || !/FATTURA|NR\.?\s*DOCUMENTO/i.test(t)) return null;
  let no = '', invDate = '';
  for (let i = 0; i < all.length && !no; i++) {
    if (!/DOCUMENT\s*No/i.test(all[i])) continue;
    for (let j = i + 1; j <= Math.min(all.length - 1, i + 2); j++) {
      const m = all[j].match(/\b(\d{4,7})\s+(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s*$/);
      if (m) { no = m[1]; const y = m[4].length === 2 ? '20' + m[4] : m[4]; invDate = dmy(m[2], m[3], y); break; }
    }
  }
  const RE_HEAD = /^(\d{6,9})\s+(\d{6})\s+(\d{6})\s+([A-Z][A-Z0-9 .\/-]*?)\s*(?:\[\d\])?$/;
  const RE_ITEM = /^(.*?)\s*\b(\S{1,3})\s+(\d{4,6})\s+(MT|MTS|M|KG|KGS|YD|YDS|PCS)\s+([\d.]+,\d+)\s+([\d.]+,\d+)\s+([\d.]+,\d+)(?:\s+\S{1,3})?$/i;
  const items = [];
  let pklNo = '', artCode = '', artName = '', total = NaN, totalQty = NaN;
  for (const L of all) {
    let m;
    if ((m = L.match(RE_HEAD))) { pklNo = stripZero(m[1]); artCode = m[2]; artName = m[4]; continue; }
    if (/^TOTAL\b/i.test(L)) { const mq = L.match(/\b(MT|MTS|M|KG|KGS|YD|YDS|PCS)\s+([\d.]+,\d+)/i); if (mq && isNaN(totalQty)) totalQty = itNum(mq[2]); continue; }
    if ((m = L.match(/^\d{8}\s+[\d.]+,\d+\s+[\d.]+,\d+\s+\d+\s+([\d.]+,\d+)$/))) { if (isNaN(total)) total = itNum(m[1]); continue; }   // bảng mã HS: … PACK NO AMOUNT USD
    if (!pklNo) continue;
    m = L.match(RE_ITEM);
    if (!m) continue;
    const color = m[1].replace(/^.*?G\/M2\s*\d+\s*/i, '').replace(/^WIDTH\s*\d+\s*CM\s*/i, '').trim();
    if (!color || !/[A-Z]/i.test(color)) continue;
    const unit = unitKey(m[4]);
    const qty = itNum(m[5]), price = itNum(m[6]), amount = itNum(m[7]);
    if (isNaN(qty) || qty <= 0) continue;
    const article = carvArticle(artCode, artName);
    const codeS = stripZero(m[3]);
    items.push({
      poRaw: '', po: '', article, colorText: color, colorCode: m[3], colorShort: color,
      desc: [article, color + ' ' + codeS, 'PKL ' + pklNo].join(' · '),
      qtyByUnit: { [unit]: qty }, unit, qty, price, surcharge: 0, amount,
      code: article + ' · ' + color, pklNo, colorKey: codeS,
      pklKey: pklNo + '|' + AZ(article) + '|' + codeS,
    });
  }
  if (!items.length) return null;
  return {
    profile: 'CARVICO', supplier: 'Carvico', noPo: true,
    inv: { no, invNo: no, invDate, items, currency: 'USD', unitDefault: items[0].unit, surchargeHeader: 0, total, totalQty, amountInclSur: true, hasSapPo: false },
    pkl: null,
  };
}
/* packing list rời của Carvico → dạng gen.pkl (kind 'genpkl'), gắn vào hóa đơn Carvico cùng số PKL */
function readCarvicoPkl(lines, fname) {
  const all = (lines || []).map((x) => String(x == null ? '' : x).replace(/\s+/g, ' ').trim());
  const t = all.join('\n');
  if (!isCarvicoText(t) || !/PACKING\s*LIST/i.test(t) || /FATTURA|NR\.?\s*DOCUMENTO/i.test(t)) return null;
  let pklNo = '', date = '';
  for (const L of all) {
    const m = L.match(/^(\d{5,8})\s+(\d{1,2})\/(\d{1,2})\/(\d{2,4})\b/);
    if (m) { pklNo = stripZero(m[1]); date = dmy(m[2], m[3], m[4].length === 2 ? '20' + m[4] : m[4]); break; }
  }
  const RE_SEC = /^(\d{6,9})\s+(\d{6})\s+(\d{6})\s+([A-Z][A-Z0-9 .\/-]*?)\s+(\d{5,7})\s+([A-Z][A-Z ]*[A-Z])$/;
  const RE_ROLL = /^(\d{8,10})\s+(\S{1,3})\s+(?:(\d{6})\s+)?([A-Z]{2}\d{5,7}[A-Z]?)\s+([\d.]+,\d+)\s+([\d.]+,\d+)(?:\s+\d+)?$/;
  const groups = [];
  let sec = null, batch = '';
  for (const L of all) {
    let m;
    if ((m = L.match(RE_SEC))) {
      /* tiêu đề mục lặp lại ở đầu trang mới (cùng màu, chưa gặp "Tot Rolls / colour") → vẫn là lô đang đọc */
      const same = sec && sec.artCode === m[2] && sec.colorCode === m[5];
      sec = { confirm: m[1], artCode: m[2], artName: m[4], colorCode: m[5], color: m[6].trim() };
      if (!same) batch = '';
      continue;
    }
    if (!sec) continue;
    if (/^Tot\s*Rolls\s*\/\s*colou?r/i.test(L)) { sec = null; continue; }
    m = L.match(RE_ROLL);
    if (!m) continue;
    if (m[3]) batch = m[3];
    const q = itNum(m[5]);
    if (isNaN(q) || q <= 0) continue;
    const article = carvArticle(sec.artCode, sec.artName);
    const codeS = stripZero(sec.colorCode);
    const itemKey = pklNo + '|' + AZ(article) + '|' + codeS;
    const key = itemKey + '|' + AZ(batch);
    let g = groups.find((x) => x.key === key);
    if (!g) {
      g = { key, itemKey, po: '', poRaw: '', article, color: sec.color, colorCode: sec.colorCode, lot: batch ? 'Lô ' + batch : '(không ghi lô)', unit: 'M', rolls: [], total: 0, pklNo };
      groups.push(g);
    }
    g.rolls.push({ no: m[1], qty: q, nw: itNum(m[6]), lotCode: m[4] });
    g.total = r3(g.total + q);
  }
  if (!groups.length) return null;
  return {
    profile: 'CARVICO', supplier: 'Carvico', role: 'pkl', file: fname || '', pklNo,
    inv: { no: pklNo, invNo: '', invDate: date, items: [], currency: 'USD', unitDefault: 'M', total: NaN },
    pkl: { groups, unit: 'M', level: 'lot', soft: false },
    pklOnly: true,
  };
}

/* =====================  SUZHOU CELEB (Trung Quốc) — hóa đơn Excel + packing list Excel rời  =====================
   Hóa đơn (sheet " INVOICE"): "INVOICE NO : CELEB260807-3" · "DATE: 7th,Agu,2026" (tháng viết sai chính tả) ·
     bảng PO NO. | DESCRIPTION OF GOODS | Color | QUANTITY (Y) | UNIT PRICE | AMOUNT:
     "CEL0010900 | RC031 Solid Fab REC PE30 PE62 SP8, 105G | Jet Black 19-0303 | 693 | 2.41 | 1670.13"
     dòng "surcharge | 1 | 150 | 150" đứng ngay dưới dòng hàng → phụ phí của dòng đó (inbound: cột Surcharge Item = 150).
   Packing list (mỗi file một màu, sheet 发货码单): Lot No. | Roll No. | PO Number | Color | Name | Comp | Q'ty(Y) | N.W | G.W
     màu viết khác nhau: "True Red（19-1664 TCX）" / "Jet Black （19-0303 TCX）" / "Pink-a-boo 13-2801 TCX".
   Inbound: Supplier Ref "RC031 Solid", Color "JET BLACK 19-0303 TCX", Lapdip "ML221111-04D1"; PO CEL0010900 ghi sẵn trên hóa đơn.
   Khóa ghép hóa đơn ↔ packing list: PO | article (RC031) | màu rút gọn (bỏ ngoặc, TCX, dấu): JETBLACK190303. */
const celebColorKey = (t) => AZ(String(t == null ? '' : t).toUpperCase().replace(/[（）()]/g, ' ').replace(/\bTCX\b/g, ' '));
const celebArticle = (t) => (String(t == null ? '' : t).trim().match(/^[A-Z]{1,4}\d{2,5}[A-Z0-9-]*/i) || [''])[0].toUpperCase();
function readCeleb(wb) {
  const wi = wb.worksheets.find((w) => /invoice/i.test(String(w.name))) || wb.worksheets[0];
  let no = '', invDate = '';
  for (let r = 1; r <= Math.min(20, wi.rowCount); r++) for (let c = 1; c <= 12; c++) {
    const t = T(wi, r, c);
    let m;
    if (!no && (m = t.match(/INVOICE\s*NO\.?\s*:?\s*([A-Z0-9][A-Z0-9\/-]{3,})/i))) no = m[1];
    if (!invDate && (m = t.match(/\bDATE\s*:?\s*(.+)$/i))) {
      const u = m[1].toUpperCase().replace(/\s+/g, ' ').trim();
      const d = u.match(/(\d{1,2})\s*(?:ST|ND|RD|TH)?\s*[,.\s]+\s*([A-Z]{3,9})\s*[,.\s]+\s*(\d{4})/);   // 7th,Agu,2026
      if (d) {
        const FIX = { AGU: 'AUG', SEPT: 'SEP', JUL: 'JUL' };
        const mon = FIX[d[2]] || d[2];
        const mi = MONTHS.findIndex((x) => mon.startsWith(x) || x.startsWith(mon.slice(0, 3)));
        if (mi >= 0) invDate = dmy(d[1], mi + 1, d[3]);
      }
      if (!invDate) invDate = parseDateCell(m[1]) || '';
    }
  }
  const hr = findRow(wi, /PO\s*NO\..*DESCRIPTION/i) || findRow(wi, /DESCRIPTION OF GOODS/i);
  if (!hr) return null;
  const C = { po: colBy(wi, hr, /^PO\s*NO/i) || 1, desc: colBy(wi, hr, /DESCRIPTION/i) || 2, color: colBy(wi, hr, /^COLOU?R/i) || 3,
    qty: colBy(wi, hr, /QUANTITY/i) || 4, price: colBy(wi, hr, /UNIT\s*PRICE/i) || 5, amt: colBy(wi, hr, /AMOUNT/i) || 6 };
  const unit = unitKey((rowText(wi, hr) + rowText(wi, hr + 1)).match(/\(\s*(Y|YD|YDS|M|MT|KG)\s*\)/i) ? (rowText(wi, hr) + rowText(wi, hr + 1)).match(/\(\s*(Y|YD|YDS|M|MT|KG)\s*\)/i)[1] : 'YD');
  const items = [];
  let total = NaN, totalQty = NaN, nSur = 0;
  for (let r = hr + 1; r <= wi.rowCount; r++) {
    const po = T(wi, r, C.po), desc = T(wi, r, C.desc);
    if (/^TOTAL/i.test(po) || /^TOTAL/i.test(desc)) { total = N(wi, r, C.amt); totalQty = N(wi, r, C.qty); break; }
    if (/surcharge|phụ phí|extra charge/i.test(desc) && items.length) {
      const a = N(wi, r, C.amt);
      if (!isNaN(a)) { const last = items[items.length - 1]; last.surcharge = r3((last.surcharge || 0) + a); last.amount = r3(last.amount + a); nSur++; }
      continue;
    }
    const qty = N(wi, r, C.qty);
    if (!po || isNaN(qty) || qty <= 0) continue;
    const color = T(wi, r, C.color).replace(/\s+/g, ' ').trim();
    const article = celebArticle(desc);
    const poS = poSap(po);
    items.push({
      poRaw: po, po: poS, article, colorText: color, colorCode: (color.match(/\d{2}-\d{4}/) || [''])[0], colorShort: color,
      desc: [article, color, desc.replace(/\s+/g, ' ')].filter(Boolean).join(' · '),
      qtyByUnit: { [unit]: qty }, unit, qty, price: N(wi, r, C.price), surcharge: 0, amount: N(wi, r, C.amt),
      code: article + ' · ' + color, pklKey: poS + '|' + AZ(article) + '|' + celebColorKey(color),
    });
  }
  if (!items.length) return null;
  /* dòng TOTAL cộng cả "số lượng" 1 của mỗi dòng phụ phí → trừ ra để so với tổng Invoice Quantity */
  if (!isNaN(totalQty) && nSur) totalQty = r3(totalQty - nSur);
  return {
    profile: 'CELEB', supplier: 'Suzhou Celeb',
    inv: { no, invNo: no, invDate, items, currency: 'USD', unitDefault: unit, surchargeHeader: 0, total, totalQty, amountInclSur: true },
    pkl: null,
  };
}
function readCelebPkl(wb, fname) {
  let head = '';
  for (const w of wb.worksheets) for (let r = 1; r <= Math.min(10, w.rowCount); r++) head += rowText(w, r);
  if (!/CELEB/i.test(head)) return null;
  const groups = [];
  let unit = 'YD';
  for (const ws of wb.worksheets) {
    const hr = findRow(ws, /LOT\s*NO\..*ROLL\s*NO\..*PO\s*NUMBER/i, 15);
    if (!hr) continue;
    const C = { lot: colBy(ws, hr, /^LOT\s*NO/i), roll: colBy(ws, hr, /^ROLL\s*NO/i), po: colBy(ws, hr, /^PO\s*NUMBER/i), color: colBy(ws, hr, /^COLOU?R/i),
      name: colBy(ws, hr, /^NAME/i), qty: colBy(ws, hr, /Q'?TY/i), nw: colBy(ws, hr, /NET\s*WEIGHT/i), gw: colBy(ws, hr, /GROSS\s*WEIGHT/i) };
    if (!C.po || !C.qty) continue;
    const um = T(ws, hr, C.qty).match(/\(\s*(Y|YD|YDS|M|MT|KG)\s*\)/i);
    if (um) unit = unitKey(um[1]);
    let lot = '', po = '', color = '', name = '';
    for (let r = hr + 1; r <= ws.rowCount; r++) {
      if (/^TOTAL/i.test(T(ws, r, C.po)) || /^TOTAL/i.test(T(ws, r, 1))) break;
      const q = N(ws, r, C.qty);
      if (isNaN(q) || q <= 0) continue;
      lot = T(ws, r, C.lot) || lot; po = T(ws, r, C.po) || po; color = T(ws, r, C.color).replace(/\s+/g, ' ').trim() || color; name = (C.name ? T(ws, r, C.name) : '') || name;
      if (!po) continue;
      const article = celebArticle(name) || name.toUpperCase();
      const itemKey = poSap(po) + '|' + AZ(article) + '|' + celebColorKey(color);
      const key = itemKey + '|' + AZ(lot);
      let g = groups.find((x) => x.key === key);
      if (!g) {
        g = { key, itemKey, po: poSap(po), poRaw: po, article, color: color.replace(/[（）()]/g, ' ').replace(/\s+/g, ' ').trim(), lot: lot ? 'Lô ' + lot : '(không ghi lô)', unit, rolls: [], total: 0 };
        groups.push(g);
      }
      g.rolls.push({ no: C.roll ? T(ws, r, C.roll) : String(g.rolls.length + 1), qty: q, nw: C.nw ? N(ws, r, C.nw) : NaN, gw: C.gw ? N(ws, r, C.gw) : NaN });
      g.total = r3(g.total + q);
    }
  }
  if (!groups.length) return null;
  return {
    profile: 'CELEB', supplier: 'Suzhou Celeb', role: 'pkl', file: fname || '',
    inv: { no: '', invNo: '', invDate: '', items: [], currency: 'USD', unitDefault: unit, total: NaN },
    pkl: { groups, unit, level: 'lot', soft: false },
    pklOnly: true,
  };
}

/* =====================  CHEUNG HING (Hồng Kông) — phụ liệu (hangtag / sticker / label), hóa đơn Excel + packing list Excel  =====================
   Hóa đơn: "NO:20260908" · "Date :8/9/2026" (ngày/tháng/năm) · bảng Item | PO | Description | <Style-Mã> | Quantity | đơn vị | Unit Price | Total
     dòng tiêu đề nhóm (HANGTAG / STICKER / LABEL) chỉ có cột C; dòng hàng: "1 | CH10034200 | Tag REC PAPER100, L63 XW63mm | Panache-PAN3ST | 4379 | pc | 0.076 | 332.804"
     đơn vị pc / set / doz (inbound: PC / SET / DZ); dòng "BANK CHARGE | 1 | transation | 50 | 50" → phụ phí chung (surchargeHeader).
   Mã hàng = phần sau dấu "-" của cột Style: Panache-PAN3ST → PAN3ST · Envy-7283 → 7283 · Sculptresse-SCLPST12/SCLPST_10 → SCLPST12/SCLPST_10
     khớp cột Specification ("PAN3ST", "white blossom- 7283", "SCLPST12/SCLPST_10") hoặc Supplier Ref ("PANMLR PANACHE") của inbound.
   Cùng mã ở hai đơn vị (Evangeline-11434: sticker pc + label doz) → lọc theo đơn vị. Sticker/label có nhiều dòng inbound theo SIZE:
     tổng PO các size = SL hóa đơn → chia theo size (KHỚP (chia theo size)); khác → CẦN KIỂM TAY, giữ dòng trong file INB để điền tay.
   Packing list: Box # | PO# | Description | Style | <mã> | Item | Quantity | đơn vị; dòng không ghi thùng/PO/mô tả thì kế thừa dòng trên. */
const isCheungHingText = (t) => /CHEUNG\s*HING/i.test(t);
const chSplitStyle = (style) => {
  const t = String(style == null ? '' : style).replace(/\s+/g, ' ').trim();
  const m = t.match(/^(.*?)\s*-\s*(.+)$/);
  return m ? { brand: m[1].trim(), code: m[2].trim() } : { brand: '', code: t };
};
const CH_PO_RE = /^[A-Z]{2,4}\d{7,8}$/i;
function readCheungHing(wb) {
  const wi = wb.worksheets.find((w) => findRow(w, /ITEM.*\bPO\b.*DESCRIPTION.*QUANTITY/i, 30)) || wb.worksheets[0];
  const hr = findRow(wi, /ITEM.*\bPO\b.*DESCRIPTION.*QUANTITY/i, 30);
  if (!hr) return null;
  let no = '', invDate = '';
  for (let r = 1; r < hr; r++) for (let c = 1; c <= 14; c++) {
    const t = T(wi, r, c);
    let m;
    if (!no && (m = t.match(/\bNO\s*[:.]?\s*([A-Z0-9][A-Z0-9\/-]{4,})/i))) no = m[1];
    if (!invDate && (m = t.match(/\bDATE\s*[:.]?\s*(.+)$/i))) invDate = parseDateCell(m[1].trim()) || '';
  }
  const C = { no: colBy(wi, hr, /^ITEM/i) || 1, po: colBy(wi, hr, /^PO$/i) || 2, desc: colBy(wi, hr, /DESCRIPTION/i) || 3,
    qty: colBy(wi, hr, /QUANTITY/i) || 5, price: colBy(wi, hr, /UNIT\s*PRICE/i) || 7, amt: colBy(wi, hr, /TOTAL/i) || 8 };
  C.style = C.desc + 1; C.unit = C.qty + 1;
  const items = [];
  let cat = '', total = NaN, surchargeHeader = 0;
  for (let r = hr + 1; r <= wi.rowCount; r++) {
    const po = T(wi, r, C.po), desc = T(wi, r, C.desc);
    const qty = N(wi, r, C.qty), amt = N(wi, r, C.amt);
    if (/^(HANGTAG|STICKER|LABEL|TAG)S?$/i.test(desc) && !po) { cat = desc.toUpperCase(); continue; }
    if (/BANK\s*CHARGE|HANDLING|COURIER|FREIGHT/i.test(desc) && !isNaN(amt)) { surchargeHeader = r3(surchargeHeader + amt); continue; }
    if (!po && !desc && !isNaN(amt) && isNaN(qty)) { total = amt; continue; }        // dòng tổng chỉ có cột Total
    if (/^REMARK/i.test(po) || /^REMARK/i.test(desc)) break;
    if (!CH_PO_RE.test(po) || isNaN(qty) || qty <= 0) continue;
    const style = T(wi, r, C.style).replace(/\s+/g, ' ').trim();
    const { brand, code } = chSplitStyle(style);
    const unit = unitKey(T(wi, r, C.unit) || 'pc');
    const poS = poSap(po);
    items.push({
      poRaw: po, po: poS, article: code, brand, colorText: '', colorCode: '', colorShort: '',
      desc: [cat, desc.replace(/\s+/g, ' '), style].filter(Boolean).join(' · '),
      qtyByUnit: { [unit]: qty }, unit, qty, price: N(wi, r, C.price), surcharge: 0, amount: amt,
      code: style || code, pklKey: poS + '|' + AZ(code) + '|' + unit,
    });
  }
  if (!items.length) return null;
  return {
    profile: 'CHEUNGHING', supplier: 'Cheung Hing',
    /* amountInclSur=false: thành tiền dòng = SL × đơn giá; phí ngân hàng (surchargeHeader) SAP ghi vào Surcharge Item của một dòng PO bất kỳ → so ở mức tổng */
    inv: { no, invNo: no, invDate, items, currency: 'USD', unitDefault: 'PCS', surchargeHeader, total, totalQty: NaN, amountInclSur: false },
    pkl: null,
  };
}
function readCheungHingPkl(wb, fname) {
  let head = '';
  for (const w of wb.worksheets) for (let r = 1; r <= Math.min(8, w.rowCount); r++) head += rowText(w, r);
  if (!isCheungHingText(head)) return null;
  const groups = [];
  for (const ws of wb.worksheets) {
    const hr = findRow(ws, /BOX\s*#.*PO\s*#.*QUANTITY/i, 30);
    if (!hr) continue;
    const C = { box: colBy(ws, hr, /^BOX/i) || 1, po: colBy(ws, hr, /^PO\s*#/i) || 2, desc: colBy(ws, hr, /DESCRIPTION/i) || 4,
      style: colBy(ws, hr, /^STYLE/i) || 5, item: colBy(ws, hr, /^ITEM/i) || 7, qty: colBy(ws, hr, /^QUANTITY/i) || 8 };
    C.code = C.style + 1; C.unit = C.qty + 1;
    let box = '', po = '', desc = '', style = '', code = '', item = '';
    for (let r = hr + 1; r <= ws.rowCount; r++) {
      const tPo = T(ws, r, C.po), tCode = T(ws, r, C.code), tItem = T(ws, r, C.item), tBox = T(ws, r, C.box);
      if (/^REMARK/i.test(tPo) || /^REMARK/i.test(T(ws, r, C.desc - 1))) break;
      if (!tPo && !tCode && !tItem) continue;                       // dòng cộng nhóm (chỉ có số) / ghi chú
      const q = N(ws, r, C.qty);
      if (tBox) box = tBox;
      if (tPo && CH_PO_RE.test(tPo)) po = tPo;
      if (T(ws, r, C.desc)) desc = T(ws, r, C.desc).replace(/\s+/g, ' ');
      if (T(ws, r, C.style)) style = T(ws, r, C.style).trim();
      if (tCode) code = tCode.replace(/\s+/g, '');
      if (tItem) item = tItem.toUpperCase();
      if (isNaN(q) || q <= 0 || !po || !code) continue;
      const unit = unitKey(T(ws, r, C.unit) || 'pc');
      const itemKey = poSap(po) + '|' + AZ(code) + '|' + unit;
      const key = itemKey + '|' + AZ(box);
      let g = groups.find((x) => x.key === key);
      if (!g) {
        g = { key, itemKey, po: poSap(po), poRaw: po, article: code, color: '', item, style, lot: box ? 'Thùng ' + box : '(không ghi thùng)', unit, rolls: [], total: 0 };
        groups.push(g);
      }
      g.rolls.push({ no: String(g.rolls.length + 1), qty: q });
      g.total = r3(g.total + q);
    }
  }
  if (!groups.length) return null;
  return {
    profile: 'CHEUNGHING', supplier: 'Cheung Hing', role: 'pkl', file: fname || '',
    inv: { no: '', invNo: '', invDate: '', items: [], currency: 'USD', unitDefault: 'PCS', total: NaN },
    pkl: { groups, unit: 'PCS', level: 'lot', soft: false },
    pklOnly: true,
  };
}

function readFab(wb) {
  const p = fabProfile(wb);
  if (p === 'TECHWORK') return readTechwork(wb);
  if (p === 'BLAO') return readBlao(wb);
  if (p === 'HYU') return readHyu(wb);
  if (p === 'YUBO') return readYubo(wb);
  if (p === 'CELEB') return readCeleb(wb);
  if (p === 'CHEUNGHING') return readCheungHing(wb);
  return null;
}

/* =====================================================================
   Đối chiếu 1 hóa đơn vải với file inbound
   Trả về đúng dạng {lines, VAL} như analyze() của trimming để dùng lại
   toàn bộ phần báo cáo.
   ===================================================================== */
function analyzeFab(inv, pkl, rows, opts) {
  opts = opts || {};
  const EPS = 0.01;
  const cur = inv.currency || 'USD';
  const pTol = cur === 'VND' ? 0.5 : 1e-6;      // sai số cho phép của đơn giá
  const aTol = cur === 'VND' ? 1 : 0.02;        // sai số cho phép của thành tiền
  const lines = [];
  const lotSeen = [];
  const consumed = new Map();   // dòng inbound -> số đã phân bổ trong hóa đơn này

  /* bộ đọc chung: ghép packing list THẬN TRỌNG — chỉ nhận nhóm packing list có đủ dấu hiệu
     (PO/màu/mã/size/lô) của dòng chứng từ; nhiều dòng chứng từ cùng trỏ về một nhóm
     (hoá đơn tách 6 + 106, packing list ghi 112) thì so tổng của các dòng đó. */
  const pklPre = new Map();
  if (pkl && pkl.groups && inv.gen) {
    const docHasPo = inv.items.some((x) => x.po);
    const sizeOfDoc = (it) => {
      if (it.size) return it.size;
      const up = String([it.article, it.desc, it.ctx].filter(Boolean).join(' · ')).toUpperCase();
      const m = up.match(/\b(\d{2,3}(?:\.\d)?)\s*CM\b/) || up.match(/\b(\d{2}[A-K]{1,2})\b/);
      if (m) return m[1] + (/CM/.test(m[0]) ? 'CM' : '');
      for (const p0 of up.split(/\s·\s/)) { const m2 = p0.replace(/[\s·,\/-]+$/, '').match(/(?:^|[\s\/·-])(XXS|XS|S|M|L|XL|XXL|XXXL|2XL|3XL|4XL|5XL)$/); if (m2) return m2[1]; }
      return '';
    };
    const sums = {};
    for (const it of inv.items) {
      const txt = [it.colorText, it.colorCode, it.desc, it.ctx, it.article].filter(Boolean).join(' ');
      const tAZ = AZ(txt);
      const ds = sizeOfDoc(it);
      const scored = pkl.groups.map((g) => {
        if (g.po && it.po && docHasPo && !poSame(g.po, it.po)) return { g, s: -99 };
        let sc = 0;
        const gc = AZ(g.color);
        if (gc && gc.length >= 3 && (tAZ.includes(gc) || colorScore(it.colorText || txt, g.color, txt) >= 6)) sc += 3;
        else if (gc && colorScore(it.colorText || txt, g.color, txt) >= 4) sc += 2;
        else if (gc) sc -= 2;
        if (g.article && artHit(g.article, txt)) sc += 2;
        else if (g.article && (!it.article || !artHit(it.article, g.article))) sc -= 3;
        if (g.size && ds) sc += sizeSame(ds, g.size) ? 2 : -6;
        if (g.lot && it.lot) sc += (AZ(it.lot) === AZ(g.lot)) ? 2 : -4;
        else if (g.lot && AZ(g.lot).length >= 4 && new RegExp('(^|[^A-Z0-9])' + AZ(g.lot) + '([^A-Z0-9]|$)').test(String(txt).toUpperCase().replace(/[^A-Z0-9]+/g, ' '))) sc += 1;
        if (!gc && !g.article && !g.size) sc += (pkl.groups.length === 1 ? 3 : 0);
        return { g, s: sc };
      }).filter((x) => x.s >= 3);
      const best = scored.length ? Math.max(...scored.map((x) => x.s)) : 0;
      const gs = scored.filter((x) => x.s === best).map((x) => x.g);
      const k = gs.map((g) => g.key).sort().join('||');
      pklPre.set(it, { gs, k });
      if (gs.length) sums[k] = r3((sums[k] || 0) + it.qty);
    }
    pklPre.sums = sums;
  }

  for (const it of inv.items) {
    /* ---- 1. lọc theo PO (bù 0 cho đủ 7 chữ số) ---- */
    let usedPo = it.po, poNote = '';
    const poWanted = (it.poList && it.poList.length) ? it.poList : [usedPo];
    let poRows = rows ? rows.filter((r) => poWanted.some((p) => poSame(r.poV, p))) : [];
    if (rows && !poRows.length && opts.poIdx) {
      const cands = [it.poRaw, it.po].filter(Boolean).map((x) => String(x).toUpperCase());
      for (const t of cands) {
        const alt = (opts.poIdx.scax && opts.poIdx.scax.get(t)) || (opts.poIdx.sap && opts.poIdx.sap.get(t));
        if (alt) {
          const rr = rows.filter((r) => poSame(r.poV, alt));
          if (rr.length) { poRows = rr; usedPo = alt; poNote = `PO tra qua file PO SCAF-SCAX: ${it.poRaw} → ${alt}. `; break; }
        }
      }
    }
    /* bộ đọc chung: chứng từ không ghi PO SAP (hoặc ghi mã riêng của chủ hàng) */
    /* …hoặc mẫu riêng mà chứng từ hoàn toàn không ghi PO (Carvico) */
    if (rows && !poRows.length && (it.gen || !String(it.poRaw || it.po || '').trim())) {
      const distinct = [...new Set(rows.map((r) => poSap(r.poV)))];
      const rawPo = String(it.poRaw || it.po || '').trim();
      if (!rawPo && inv.hasSapPo) {
        poNote += 'Dòng này không ghi PO (các dòng khác có) — không tự gán. ';
      } else if (!rawPo) {
        poRows = rows; usedPo = distinct.length === 1 ? distinct[0] : '';
        poNote += distinct.length === 1 ? `Chứng từ không ghi PO — inbound chỉ có PO ${usedPo}. ` : `Chứng từ không ghi PO — dò trên toàn bộ ${distinct.length} PO của inbound. `;
      } else if (distinct.length === 1 && !/^(?:[A-Z][A-Z&]{2}\d{7}|CH\d{8})$/.test(rawPo.toUpperCase())) {
        poRows = rows; usedPo = distinct[0];
        poNote += `PO trên chứng từ "${rawPo}" không phải mã SAP — inbound chỉ có PO ${usedPo} nên dùng mã này. `;
      }
    }
    if (rows && poRows.length) {
      if (poWanted.length > 1) poNote += `Ô PO ghi ${poWanted.length} mã (${poWanted.join(', ')}) — đã dò trong tất cả. `;
      else if (poSap(it.poRaw) !== String(it.poRaw).toUpperCase().trim()) {
        poNote += `PO trên hóa đơn ghi "${it.poRaw}" — đã bù 0 thành ${poSap(it.poRaw)}. `;
      }
    }

    /* ---- 2. lọc theo mã article rồi chấm điểm màu ---- */
    const invText = [it.colorText, it.colorCode, it.desc, it.ctx].filter(Boolean).join(' ');
    const ctxAZ = AZ(invText);
    /* mã Material SAP ghi ngay trên chứng từ → lấy thẳng dòng đó (kể cả khi PO ghi khác) */
    let matBy = false;
    const mats = (it.mats && it.mats.length ? it.mats : (it.material ? [it.material] : [])).map(AZ);
    if (rows && mats.length) {
      const src = poRows.length ? poRows : rows;
      let byMat = src.filter((r) => mats.includes(AZ(r.material)));
      /* mã Material chỉ xuất hiện ở dòng mô tả chung (KNE ghi 3 mã trên một dòng) → không dùng khi trỏ tới nhiều dòng */
      if (it.matsFromCtx && (byMat.length > 1 || mats.length > 1)) byMat = [];
      if (byMat.length) { if (!poRows.length) { poRows = byMat; usedPo = byMat[0].poV; poNote += 'Dò theo mã Material ghi trên chứng từ. '; } matBy = true; }
    }
    let pool = matBy ? poRows.filter((r) => mats.includes(AZ(r.material)))
      : poRows.filter((r) => artHit(it.article, r.material, r.desc, r.spec, r.supRef) || (it.gen && refHit(r, invText)));
    /* chứng từ ghi đơn vị (pc / doz / set): cùng mã nhưng khác đơn vị là mặt hàng khác (Cheung Hing: sticker tính cái, nhãn tính tá) */
    if (pool.length > 1 && it.unit) {
      const sameU = pool.filter((r) => unitKey(r.unit) === unitKey(it.unit));
      if (sameU.length && sameU.length < pool.length) pool = sameU;
    }
    const artOK = pool.length > 0;
    if (!pool.length) pool = poRows;
    /* size ghi trên chứng từ: giữ đúng size (nếu inbound có cột size) */
    let sizeBy = false, sizeMiss = false;
    let docSize = it.size || '';
    if (it.gen && pool.some((r) => r.size)) {
      /* ứng viên size: cột size trên chứng từ, rồi size nằm trong mô tả ("PAN38 36.9CM" · "PD2659-1-75B" · "Bra cup M") */
      const cands = [];
      if (it.size) cands.push(it.size);
      const up = String(invText).toUpperCase();
      const m = up.match(/\b(\d{2,3}(?:\.\d)?)\s*CM\b/) || up.match(/\b(\d{2}[A-K]{1,2})\b/);
      if (m) cands.push(m[1] + (/CM/.test(m[0]) ? 'CM' : ''));
      const parts = String((it.article || '') + ' · ' + (it.desc || '')).toUpperCase().split(/\s·\s/);
      for (const p0 of parts) {
        const t = p0.replace(/[\s·,\/-]+$/, '');
        const m2 = t.match(/(?:^|[\s\/·-])(XXS|XS|S|M|L|XL|XXL|XXXL|2XL|3XL|4XL|5XL|F|OS)$/);
        if (m2) { cands.push(m2[1]); break; }
      }
      let picked = '';
      for (const c of cands) { const ps = pool.filter((r) => sizeSame(c, r.size)); if (ps.length) { pool = ps; sizeBy = true; picked = c; break; } }
      docSize = picked || cands[0] || '';
      if (!picked && cands.length && artOK) { pool = []; sizeMiss = true; }
    } else if (docSize && pool.some((r) => r.size)) {
      const ps = pool.filter((r) => sizeSame(docSize, r.size));
      if (ps.length) { pool = ps; sizeBy = true; }
    }
    const scored = pool.map((r) => {
      let s = colorScore(it.colorText, r.color, invText) * 3;
      const lt = WORDS(r.lapdip)[0];
      if (lt && lt.length >= 6 && (AZ(invText).includes(lt) || AZ(it.colorCode).includes(lt))) s += 2;
      if (!isNaN(r.price) && !isNaN(it.price) && Math.abs(r.price - it.price) <= pTol) s += 1;
      if (it.gen) {
        /* mã trong Specification / Lapdip có trên chứng từ (+2 mỗi loại) */
        if (codeTokens(r.spec).some((k) => ctxAZ.includes(k))) s += 2;
        if (codeTokens(r.lapdip).some((k) => k.length >= 5 && ctxAZ.includes(k))) s += 2;
        if (matBy) s += 3;
      }
      return { r, s };
    });
    const best = scored.length ? Math.max(...scored.map((x) => x.s)) : -1;
    let hit = best > 0 ? scored.filter((x) => x.s === best).map((x) => x.r) : [];
    if (!hit.length && pool.length === 1 && artOK) hit = [pool[0]];   // 1 dòng duy nhất trong PO + đúng article
    /* còn nhiều dòng cùng điểm → thử tách bằng số lượng: chỉ giữ dòng mà SL hóa đơn
       nằm trong khoảng PO cho phép, rồi ưu tiên dòng có Quantity đúng bằng SL hóa đơn */
    let pickedByQty = false;
    const sameMat = hit.length > 1 && new Set(hit.map((r) => AZ(r.material))).size === 1;
    if (hit.length > 1 && !sameMat) {
      const qAny = it.qtyByUnit[unitKey(hit[0].unit)] != null ? it.qtyByUnit[unitKey(hit[0].unit)] : it.qty;
      if (!isNaN(qAny)) {
        const exact = hit.filter((r) => Math.abs((isNaN(r.qty) ? -1 : r.qty) - qAny) <= EPS);
        const fits = hit.filter((r) => r.unltd || qAny <= (isNaN(r.overTol) ? (isNaN(r.qty) ? 0 : r.qty) : r.overTol) - (isNaN(r.deliv) ? 0 : r.deliv) + EPS);
        if (exact.length === 1) { hit = exact; pickedByQty = true; }
        else if (fits.length === 1) { hit = fits; pickedByQty = true; }
      }
    }
    /* nhiều dòng cùng điểm, khác Material (các size của cùng mã) mà tổng còn nhận của
       chúng đúng bằng SL chứng từ → chứng từ gộp các size: chia đúng theo PO từng size */
    let splitSize = false;
    if (hit.length > 1 && !sameMat && !pickedByQty) {
      const need = hit.reduce((a, r) => a + Math.max(0, (isNaN(r.qty) ? 0 : r.qty) - (isNaN(r.deliv) ? 0 : r.deliv)), 0);
      const qAny = it.qtyByUnit[unitKey(hit[0].unit)] != null ? it.qtyByUnit[unitKey(hit[0].unit)] : it.qty;
      if (!isNaN(qAny) && need > 0 && Math.abs(need - qAny) <= EPS) splitSize = true;
    }
    const alt = pool.filter((r) => !hit.includes(r));
    /* nhiều dòng khác Material mà không tách được → KHÔNG tự điền (tránh ghi bừa theo FIFO) */
    const ambiguousMat = hit.length > 1 && !sameMat && !splitSize;

    /* ---- 3. số lượng theo đúng đơn vị của inbound ---- */
    const units = [...new Set(hit.map((r) => unitKey(r.unit)).filter(Boolean))];
    const inbUnit = units.length === 1 ? units[0] : '';
    const avail = Object.keys(it.qtyByUnit || {});
    let q = NaN, unitBad = false;
    let unitNote = '';
    if (!hit.length) q = it.qty;
    else if (inbUnit && it.qtyByUnit[inbUnit] != null) q = it.qtyByUnit[inbUnit];
    else if (!inbUnit && avail.length === 1) q = it.qtyByUnit[avail[0]];
    else if (it.gen && !avail.length) { q = it.qty; unitNote = `Chứng từ không ghi đơn vị — coi như ${inbUnit || 'cùng đơn vị inbound'}. `; }
    else { q = NaN; unitBad = true; }

    /* ---- 4. phân bổ số lượng vào (các) dòng inbound ----
       Cùng một mã vải + màu có thể nằm trong nhiều PO với số Material y hệt nhau
       (Yubo). Không có cách nào đọc ra PO đúng, nên chia theo PO cũ trước (FIFO):
       lấp đầy phần còn nhận được của PO số nhỏ nhất, thừa thì tràn sang PO kế tiếp. */
    /* PO đánh dấu Unltd Overdelivery = X → SAP cho giao vượt không giới hạn, không có mức dung sai */
    const roomOf = (r) => (r.unltd ? Infinity : (isNaN(r.overTol) ? (isNaN(r.qty) ? 0 : r.qty) : r.overTol)
      - (isNaN(r.deliv) ? 0 : r.deliv) - (consumed.get(r) || 0));
    const alloc = [];
    let overflow = 0;
    if (hit.length && !unitBad && !isNaN(q) && !ambiguousMat) {
      if (splitSize) {
        hit.forEach((r) => { const take = r3(Math.max(0, (isNaN(r.qty) ? 0 : r.qty) - (isNaN(r.deliv) ? 0 : r.deliv))); if (take > EPS) alloc.push({ r, qty: take }); });
      } else if (hit.length === 1) {
        const rm = roomOf(hit[0]);
        alloc.push({ r: hit[0], qty: q });
        if (q > rm + EPS) overflow = r3(q - rm);
      } else {
        const order = hit.slice().sort((x, y) => String(x.poV).localeCompare(String(y.poV)));
        /* dòng PO đã giao đủ (Delivered ≥ Quantity) coi như đã xong — chỉ dùng phần dung sai
           còn lại của nó khi các dòng còn mở không đủ chỗ                                  */
        const isOpen = (r) => (isNaN(r.deliv) ? 0 : r.deliv) < (isNaN(r.qty) ? 0 : r.qty) - EPS;
        const seq = order.filter(isOpen).concat(order.filter((r) => !isOpen(r)));
        let left = q;
        for (const r of seq) {
          if (left <= EPS) break;
          const rm = roomOf(r);
          if (rm <= EPS) continue;
          const take = r3(Math.min(left, rm));
          alloc.push({ r, qty: take });
          left = r3(left - take);
        }
        if (left > EPS) {
          if (alloc.length) alloc[alloc.length - 1].qty = r3(alloc[alloc.length - 1].qty + left);
          else alloc.push({ r: seq[0], qty: q });
          overflow = left;
        }
      }
      alloc.forEach((x) => consumed.set(x.r, r3((consumed.get(x.r) || 0) + x.qty)));
    }
    const split = alloc.length > 1;

    /* ---- 5. giá trị, dung sai ---- */
    /* các số của PO tính trên đúng những dòng đã phân bổ, không phải mọi dòng ứng viên */
    const used = alloc.length ? alloc.map((x) => x.r) : hit;
    const inbPrices = [...new Set(used.map((h) => h.price).filter((v) => !isNaN(v)))].sort((a, b) => a - b);
    const surSum = used.reduce((a, h) => a + (isNaN(h.sur) ? 0 : h.sur), 0);
    const qtyPo = r3(used.reduce((a, h) => a + (isNaN(h.qty) ? 0 : h.qty), 0));
    const delivered = r3(used.reduce((a, h) => a + (isNaN(h.deliv) ? 0 : h.deliv), 0));
    const overTol = r3(used.reduce((a, h) => a + (isNaN(h.overTol) ? (isNaN(h.qty) ? 0 : h.qty) : h.overTol), 0));
    const hadQty = r3(used.reduce((a, h) => a + (isNaN(h.invQty) ? 0 : h.invQty), 0));
    const unitPrice = inbPrices.length === 1 ? inbPrices[0] : (isNaN(it.price) ? 0 : it.price);
    /* hàng FOC / dòng không ghi thành tiền: có nhận (tính vào SL) nhưng không tính tiền */
    const unpaid = !isNaN(q) && it.unpaidQty > 0 && it.unpaidQty < q + EPS ? it.unpaidQty : 0;
    const inbAmount = (isNaN(q) ? 0 : q - unpaid) * unitPrice + (inv.amountInclSur ? surSum : 0);
    const priceBad = used.length > 0 && (inbPrices.length > 1
      || (!isNaN(it.price) && inbPrices.length === 1 && Math.abs(inbPrices[0] - it.price) > pTol));
    const amtBad = alloc.length > 0 && !unitBad && !isNaN(it.amount) && !isNaN(q)
      && Math.abs(inbAmount - it.amount) > aTol;
    const unltd = used.some((h) => h.unltd);
    const room = unltd ? Infinity : r3(overTol - delivered);
    const overBad = overflow > EPS;
    /* Delivered Qty đã đúng bằng SL hóa đơn → hàng đã nhập kho trước khi đối chiếu hóa đơn (chỉ ghi chú) */
    const grDone = used.length === 1 && !isNaN(q) && delivered > EPS && Math.abs(delivered - q) <= EPS;

    /* ---- 5. packing list: lô / cây ---- */
    const gsOf = (arr) => {
      /* chứng từ nào tự sinh nhóm packing list từ chính dòng hàng (Yubo) thì ghép
         bằng khóa chính xác — dò mờ theo màu sẽ gộp lẫn các màu "matching" với nhau */
      if (it.pklKey) {
        const ex = (arr || []).filter((g) => g.key && (g.key === it.pklKey || g.key.indexOf(it.pklKey + '|') === 0));
        if (ex.length) return ex;
      }
      return (arr || []).filter((g) => (!g.po || poWanted.some((p) => poSame(g.po, p)))
        && (!g.article || !it.article || artHit(g.article, it.article) || artHit(it.article, g.article))
        && (!g.size || !docSize || sizeSame(docSize, g.size))
        && (colorScore(it.colorText || (it.gen ? invText : ''), g.color, invText) >= 3 || (it.gen && !g.color && (arr || []).length === 1)));
    };
    let pklTotal = 0, lots = [], pklSize = {}, pklSumDoc = NaN;
    if (pkl && pklPre.has(it)) {
      const pre = pklPre.get(it);
      pre.gs.forEach((g) => {
        pklTotal = r3(pklTotal + g.total);
        if (g.lot) { lots.push({ lot: g.lot, rolls: g.rolls.length, qty: r3(g.total), unit: g.unit || inbUnit }); pklSize[g.lot] = r3((pklSize[g.lot] || 0) + g.total); }
      });
      pklSumDoc = pre.gs.length ? pklPre.sums[pre.k] : NaN;
    } else if (pkl) {
      const gs = gsOf(pkl.groups);
      gs.forEach((g) => { pklTotal = r3(pklTotal + g.total); });
      const src = (pkl.lots && pkl.lots.length) ? gsOf(pkl.lots) : gs.filter((g) => g.lot);
      src.forEach((g) => {
        lots.push({ lot: g.lot || '(không ghi lô)', rolls: g.rolls.length, qty: r3(g.total), unit: g.unit || inbUnit });
        pklSize[g.lot || '(không ghi lô)'] = r3((pklSize[g.lot || '(không ghi lô)'] || 0) + g.total);
      });
      if (!gs.length && src.length) src.forEach((g) => { pklTotal = r3(pklTotal + g.total); });
    }
    const pklDiff = !!pkl && pklTotal > 0 && !isNaN(q) && (isNaN(pklSumDoc) ? Math.abs(pklTotal - q) > EPS : Math.abs(pklTotal - pklSumDoc) > EPS);
    /* packing list đọc từ PDF chỉ để tham khảo (khó tách cuộn/lô chính xác) — không hạ trạng thái */
    const pklBad = pklDiff && !(pkl && pkl.soft);
    const pklSoftNote = pklDiff && pkl && pkl.soft ? ` Packing list (PDF) đọc được ${pklTotal} ≠ ${isNaN(pklSumDoc) ? q : pklSumDoc} — chỉ để tham khảo, kiểm tra tay nếu cần.` : '';
    const diffs = pklBad ? lots.map((L) => ({ size: L.lot, inb: '', pkl: L.qty, diff: '' })) : [];
    if (lots.length) lotSeen.push({ it, lots });

    /* ---- 6. kết luận ---- */
    let status, note = poNote;
    if (!rows) {
      status = 'CHƯA CÓ INBOUND';
      note += pklTotal ? (Math.abs(pklTotal - it.qty) < EPS ? `Chưa có file inbound — packing list khớp hóa đơn (${it.qty})`
        : `Chưa có file inbound — packing list ${pklTotal} ≠ hóa đơn ${it.qty}`) : 'Chưa có file inbound';
    } else if (!poRows.length) {
      status = 'LỖI';
      note += `Không có PO ${it.po} trong file inbound` + (opts.hasPoFile ? '.' : ' — nếu PO này thuộc hệ thống cũ, hãy thả thêm file PO SCAF-SCAX để em tra chuyển đổi.');
    } else if (!hit.length && sizeMiss) {
      status = 'THIẾU DÒNG';
      note += `PO ${usedPo} có mã "${it.article}" nhưng không có size ${docSize} (các size trong PO: ${[...new Set(poRows.map((r) => r.size).filter(Boolean))].slice(0, 12).join(', ')})`;
    } else if (!hit.length) {
      status = 'THIẾU DÒNG';
      note += `PO ${usedPo} có ${poRows.length} dòng nhưng không dòng nào khớp article "${it.article}" + màu "${String(it.colorText).split(/[\r\n]/)[0]}"`;
    } else if (hit.length > 1 && !sameMat && !splitSize) {
      status = 'CẦN KIỂM TAY';
      note += `${hit.length} dòng inbound cùng điểm khớp nhưng KHÁC mã Material (${[...new Set(hit.map((h) => h.material || h.color))].join(' · ')}) — em không tự điền để tránh sai, anh chọn tay.`;
    } else if (unitBad) {
      status = 'SAI ĐƠN VỊ';
      note += `Inbound tính bằng ${inbUnit || '(không rõ)'} nhưng hóa đơn chỉ có ${avail.join(' / ') || '(không rõ)'} — không so trực tiếp được.`;
    } else if (priceBad || amtBad) {
      status = 'LỆCH GIÁ TRỊ';
      note += (priceBad ? `Đơn giá hóa đơn ${it.price} ≠ đơn giá PO ${inbPrices.join(' / ')}. ` : '')
        + (amtBad ? `Thành tiền hóa đơn ${it.amount} ≠ tính theo PO ${Math.round(inbAmount * 1000) / 1000} (lệch ${Math.round((inbAmount - it.amount) * 1000) / 1000}). ` : '')
        + 'CẦN KIỂM TRA LẠI HÓA ĐƠN.';
    } else if (overBad) {
      status = 'VƯỢT DUNG SAI';
      note += `SL hóa đơn ${q} ${inbUnit} vượt mức cho phép của PO (PO ${qtyPo}, tối đa ${overTol}, đã giao ${delivered} → còn ${room}) — thừa ${overflow}, SAP sẽ báo lỗi khi import.`;
    } else if (pklBad) {
      status = 'LỆCH PKL';
      note += `Hóa đơn ${!isNaN(pklSumDoc) && Math.abs(pklSumDoc - q) > EPS ? pklSumDoc + ' (gộp các dòng cùng nhóm)' : q} ${inbUnit} nhưng packing list ${pklTotal}` + (lots.length ? ` (${lots.length} lô: ` + lots.map((L) => `${L.lot} ${L.qty}`).join('; ') + ')' : '');
    } else if (splitSize) {
      status = 'KHỚP (chia theo size)';
      note += `Chứng từ ghi gộp ${q} ${inbUnit} cho ${alloc.length} dòng (size) — tổng PO đúng bằng nên điền theo từng dòng: `
        + alloc.map((x) => `${x.r.size || x.r.material} ${x.qty}`).join(' + ') + '.';
    } else if (split) {
      status = 'KHỚP (chia nhiều PO)';
      note += `Mã + màu này có ở ${hit.length} PO với cùng số Material — đã chia theo PO cũ trước: `
        + alloc.map((x) => `${x.r.poV} ${x.qty}`).join(' + ') + `. Kiểm lại nếu thứ tự PO khác.`;
    } else if (q > qtyPo + EPS && unltd) {
      status = 'KHỚP (trong dung sai)';
      note += `Giao ${q} ${inbUnit} vượt PO ${qtyPo} — PO cho phép giao vượt không giới hạn (Unltd Overdelivery = X).`;
    } else if (q > qtyPo + EPS) {
      status = 'KHỚP (trong dung sai)';
      note += `Giao ${q} ${inbUnit} vượt PO ${qtyPo} nhưng còn trong dung sai (tối đa ${overTol}).`;
    } else if (delivered > EPS && q > qtyPo - delivered + EPS) {
      /* PO đã giao một phần: đợt này vượt phần còn lại nhưng cộng dồn vẫn trong dung sai */
      status = 'KHỚP (trong dung sai)';
      note += `Giao ${q} ${inbUnit}, PO ${qtyPo} đã giao ${delivered} chỉ còn ${Math.round((qtyPo - delivered) * 1000) / 1000} — vượt ${Math.round((q - qtyPo + delivered) * 1000) / 1000} nhưng cộng dồn còn trong dung sai (tối đa ${overTol}).`;
    } else if (q < qtyPo - EPS) {
      status = 'KHỚP (giao thiếu)';
      note += `Giao ${q}/${qtyPo} ${inbUnit} — còn lại ${Math.round((qtyPo - delivered - q) * 1000) / 1000}.`;
    } else {
      status = 'KHỚP';
    }
    /* CẦN KIỂM TAY: giữ nguyên các dòng ứng viên trong file INB (có số/ngày HĐ, Invoice Quantity để nguyên) để buyer điền tay */
    if (status === 'CẦN KIỂM TAY' && hit.length) {
      hit.forEach((r) => { r.keepHand = r.keepHand || it; });
      note += ` Đã giữ ${hit.length} dòng này trong file INB (ghi sẵn số/ngày hóa đơn) — điền Invoice Quantity từng dòng bằng tay` + (!isNaN(q) ? ` cho đủ ${q}.` : '.');
    }
    if (grDone && String(status).indexOf('KHỚP') === 0) note += ` Delivered Qty ${delivered} đã bằng SL hóa đơn — hàng đã nhập kho trước, chỉ cần ghi số/ngày hóa đơn.`;
    if (alloc.length && !unitBad && !isNaN(q) && Math.abs(hadQty - q) > EPS) {
      note += ` Đã ghi Invoice Quantity = ` + (split ? alloc.map((x) => `${x.qty} (${x.r.poV})`).join(' + ') : String(q))
        + (hadQty ? ` (file có sẵn ${hadQty})` : '') + '.';
    }
    if (unpaid) note += ` Gồm ${unpaid} ${inbUnit || 'KG'} không tính tiền (${it.unpaidNote}) — SL nhập vẫn tính phần này, thành tiền so trên ${r3(q - unpaid)} ${inbUnit || 'KG'}.`;
    if (pickedByQty) note += ' Nhiều dòng cùng điểm màu — đã tách bằng số lượng.';
    if (unitNote) note += ' ' + unitNote;
    if (pklSoftNote) note += pklSoftNote;
    if ((it.warn || []).length) note += ' ⚠ ' + it.warn.join('; ') + '.';
    if (alt.length && hit.length === 1 && status.indexOf('KHỚP') === 0) {
      note += ` (PO còn ${alt.length} dòng màu khác: ${alt.map((x) => x.color).slice(0, 3).join(' · ')})`;
    }

    /* ---- 7. ghi vào dòng inbound ---- */
    const ok = alloc.length > 0 && !unitBad && !isNaN(q);
    if (ok) {
      alloc.forEach((x) => {
        x.r.matched = x.r.matched || it;
        x.r.eff = r3((x.r.eff || 0) + x.qty);
        x.r.setQty = r3((x.r.setQty || 0) + x.qty);
      });
    }

    const cs = String((hit.length === 1 && hit[0].color) ? hit[0].color : (it.colorShort || '')).replace(/\s+/g, ' ').trim().slice(0, 30);
    it.code = [it.article, cs].filter(Boolean).join(' · ') || it.code || '';
    it.po = usedPo; it.poVia = poNote.indexOf('SCAF-SCAX') >= 0 ? 'ScaX' : 'ScaF'; it.poScax = '';
    it.qty = isNaN(q) ? it.qty : q;
    it.unitUsed = inbUnit;

    lines.push({
      it, sapPo: poRows.length ? (alloc.length ? [...new Set(alloc.map((x) => x.r.poV))].join(' + ') : usedPo) : '', hit, base: isNaN(q) ? 0 : q,
      sumQty: qtyPo, sumInvQty: hadQty, pklCodes: lots.map((L) => L.lot), pklSize, pklTotal,
      bySize: pklBad ? {} : pklSize, diffs, poRows: [], status, note: note.trim(),
      inbPrices, inbAmount: ok ? Math.round(inbAmount * 1e6) / 1e6 : 0, priceBad, amtBad,
      hasInb: !!rows, hasPkl: !!pkl, useInv: ok,
      matchBy: (matBy ? 'mã Material' : (artOK ? 'PO + article + màu' : 'PO + màu')) + (sizeBy ? ' + size' : '') + (pickedByQty ? ' + số lượng' : ''),
      ambiguous: hit.length > 1, usedVar: '', pickedByQty, altRows: alt.length,
      isFab: true, lots, qtyPo, overTol, delivered, unitUsed: inbUnit,
      alloc: alloc.map((x) => ({ po: x.r.poV, material: x.r.material, qty: x.qty })), split: split || splitSize, overflow,
      surInb: alloc.length ? surSum : 0,
    });
  }

  const rd = (n) => (isNaN(n) || n == null ? n : Math.round(n * 1e6) / 1e6);
  const itemsTotal = rd(lines.reduce((a, l) => a + (isNaN(l.it.amount) ? 0 : l.it.amount), 0));
  const invTotal = isNaN(inv.total) ? rd(itemsTotal + (inv.surchargeHeader || 0)) : inv.total;
  const cmp = lines.filter((l) => l.hasInb && l.useInv);
  const inbTotal = rd(cmp.reduce((a, l) => a + l.inbAmount, 0));
  const invCmpTotal = rd(cmp.reduce((a, l) => a + (isNaN(l.it.amount) ? 0 : l.it.amount), 0));
  const totalDiff = rd(inbTotal - invCmpTotal);
  const valueLines = lines.filter((l) => l.priceBad || l.amtBad);
  const okSt = (s) => s.indexOf('KHỚP') === 0 || s === 'CHƯA CÓ INBOUND';
  const otherLines = lines.filter((l) => !(l.priceBad || l.amtBad) && !okSt(l.status));
  const amtOf = (l) => l.inbAmount - (isNaN(l.it.amount) ? 0 : l.it.amount);
  const hasAmt = lines.some((l) => !isNaN(l.it.amount));
  const VAL = {
    inbTotal, itemsTotal, invTotal, invCmpTotal, totalDiff, noInvoiceNo: !!inv.noInvoiceNo, noFrom: inv.noFrom || '',
    valueBad: valueLines.length > 0,
    /* phụ phí ghi trong inbound (Surcharge Item của các dòng đã ghép) — đối chiếu với phụ phí dòng riêng trên hóa đơn khi hóa đơn không cộng phụ phí vào dòng hàng */
    surInb: rd(lines.reduce((a, l) => a + (l.surInb || 0), 0)),
    surMatch: !inv.amountInclSur && (inv.surchargeHeader || 0) > 0 && Math.abs(rd(lines.reduce((a, l) => a + (l.surInb || 0), 0)) - inv.surchargeHeader) <= aTol,
    totalBad: hasAmt && cmp.length > 0 && Math.abs(totalDiff) > aTol * Math.max(1, cmp.length),
    noAmounts: !hasAmt,
    valueLines, otherLines, pendingLines: [], cmpCount: cmp.length,
    pdfTotal: inv.pdfTotal, pdfDiff: (inv.pdfTotal == null || isNaN(inv.pdfTotal)) ? NaN : rd(itemsTotal - inv.pdfTotal),
    pdfBad: !(inv.pdfTotal == null || isNaN(inv.pdfTotal)) && Math.abs(itemsTotal - inv.pdfTotal) > aTol * Math.max(1, lines.length),
    noSerial: !!inv.noSerial,
    docQty: inv.totalQty, wroteQty: rd(lines.reduce((a, l) => a + (l.alloc || []).reduce((x, y) => x + y.qty, 0), 0)),
    qtyBad: !(inv.totalQty == null || isNaN(inv.totalQty))
      && Math.abs(rd(lines.reduce((a, l) => a + (l.alloc || []).reduce((x, y) => x + y.qty, 0), 0)) - inv.totalQty) > 0.01,
    unitLabel: inv.unitDefault || '',
    valueDiff: valueLines.reduce((a, l) => a + amtOf(l), 0),
    otherDiff: otherLines.reduce((a, l) => a + amtOf(l), 0),
    hasInb: !!rows, isFab: true, surchargeHeader: inv.surchargeHeader || 0, currency: cur,
  };
  return { lines, VAL };
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { AZ, WORDS, colorCore, poSap, poSame, poListOf, unitKey, colorScore, shareWords, artHit, fabProfile, readFab, parseDateCell, analyzeFab, artOf, refKeys, refHit, codeTokens };
}
