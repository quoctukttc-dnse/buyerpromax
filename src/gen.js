/* =====================================================================
   GEN — bộ đọc CHUNG cho chủ hàng chưa có mẫu riêng (v11)
   Nhận: Excel (ExcelJS workbook) hoặc PDF (danh sách dòng chữ có toạ độ)
   Trả: { role:'inv'|'pkl', inv:{invNo, invDate, items[]}, pkl:{groups[]} }
   đúng cấu trúc của fab.js để dùng lại analyzeFab + toàn bộ báo cáo.
   Nguyên tắc: không đoán cột cứng; tìm dòng tiêu đề bằng từ khoá, dòng hàng
   là dòng có SỐ LƯỢNG; mã PO SAP / mã Material / màu / size lấy từ chữ trên
   dòng (và các dòng "dính" phía trên trong cùng khối).
   ===================================================================== */

const GEN_UNIT = {
  M: 'M', MT: 'M', MTR: 'M', MTRS: 'M', MTS: 'M', METER: 'M', METERS: 'M', METRE: 'M', METRES: 'M', MET: 'M',
  Y: 'YD', YD: 'YD', YDS: 'YD', YARD: 'YD', YARDS: 'YD', YARDAGE: 'YD',
  PC: 'PCS', PCS: 'PCS', PIECE: 'PCS', PIECES: 'PCS', EA: 'PCS',
  PR: 'PAA', PRS: 'PAA', PAIR: 'PAA', PAIRS: 'PAA', PAA: 'PAA', '對': 'PAA', '对': 'PAA',
  SET: 'SET', SETS: 'SET', DZ: 'DZ', DOZ: 'DZ', DOZEN: 'DZ', KG: 'KG', KGS: 'KG',
};
const GEN_UNIT_RE = /^(M|MT|MTRS?|MTS|METERS?|METRES?|MET|Y|YDS?|YARDS?|YARDAGE|PCS?|PIECES?|EA|PRS?|PAIRS?|PAA|SETS?|DZ|DOZ|DOZEN|KGS?)\.?$/i;
const genUnit = (s) => GEN_UNIT[String(s == null ? '' : s).toUpperCase().replace(/[^A-Z對对]/g, '')] || '';

/* số: "3.386,30" → 3386.3 · "1,673.0" → 1673 · "632,50" → 632.5 · "7.606" (VN) → 7606 */
function genNum(raw, vn) {
  if (raw == null || raw === '') return NaN;
  if (typeof raw === 'number') return raw;
  if (typeof raw === 'object') {
    if (raw.result != null) return genNum(raw.result, vn);
    if (Array.isArray(raw.richText)) return genNum(raw.richText.map((x) => x.text || '').join(''), vn);
    return NaN;
  }
  let s = String(raw).trim().replace(/^[^\d\-+.,]*/, '').replace(/[^\d.,\-]+$/, '');
  if (!/\d/.test(s)) return NaN;
  const dot = s.lastIndexOf('.'), com = s.lastIndexOf(',');
  if (dot >= 0 && com >= 0) {
    s = dot > com ? s.replace(/,/g, '') : s.replace(/\./g, '').replace(',', '.');
  } else if (com >= 0) {
    /* chỉ có dấu phẩy: "1,673" / "3,000,000" → phần nghìn · "632,50" / "770,00" (VN) → thập phân */
    const parts = s.split(',');
    const tail = parts[parts.length - 1].length;
    if (parts.length > 2 || (tail === 3 && !vn)) s = s.replace(/,/g, '');
    else s = s.replace(',', '.');
  } else if (dot >= 0) {
    const tail = s.length - dot - 1;
    if ((s.match(/\./g) || []).length > 1) s = s.replace(/\./g, '');
    else if (vn && tail === 3) s = s.replace('.', '');
  }
  const n = parseFloat(s);
  return isNaN(n) ? NaN : n;
}

const GEN_PO_RE = /\b(?:P\.?O\.?\s*[:#.\-]?\s*)?([A-Z][A-Z&]{1,5}\d{7})\b/g;
/* mã PO SAP đứng riêng hoặc sau "PO:", "PO.", "PO#", "P/O No:" */
function genPos(text) {
  const t = String(text == null ? '' : text).toUpperCase();
  const out = [];
  let m;
  const re = /\b([A-Z][A-Z&]{2}\d{7}|[A-Z]{2}\d{8})\b/g;
  while ((m = re.exec(t))) if (!out.includes(m[1])) out.push(m[1]);
  return out;
}
/* mã Material SAP: 8 ký tự chữ/số + 4 hoặc 7 chữ số — chỉ tin khi có trong inbound */
const GEN_MAT_RE = /\b([A-Z][A-Z0-9]{7}\d{4}(?:\d{3})?)\b/g;
function genMats(text) {
  const t = String(text == null ? '' : text).toUpperCase();
  const out = []; let m;
  const re = new RegExp(GEN_MAT_RE.source, 'g');
  while ((m = re.exec(t))) if (!out.includes(m[1])) out.push(m[1]);
  return out;
}

/* size: "36.9CM" ↔ "0369" ↔ "0309-30.9CM" · "12DD" · "75B" · "M" */
function sizeKeyG(s) {
  let k = AZ(s).replace(/CM$/, '');
  k = k.replace(/^0+(?=\d)/, '');
  return k;
}
function sizeSame(doc, inb) {
  const a = sizeKeyG(doc), b = sizeKeyG(inb);
  if (!a || !b) return false;
  if (a === b) return true;
  /* inbound "0309-30.9CM" → "309309": chứa size chứng từ ở đầu */
  if (b.length > a.length && b.startsWith(a) && /\d/.test(a) && a.length >= 2) return true;
  if (a.length > b.length && a.startsWith(b) && /\d/.test(b) && b.length >= 2) return true;
  return false;
}

/* size ghi trong chữ của dòng chứng từ: "PAN38 36.9CM" · "BRA-CUP 75B" · "Bra cup M" · "… / L" */
function docSizeOf(it) {
  if (it.size) return it.size;
  const up = String([it.article, it.desc, it.ctx].filter(Boolean).join(' · ')).toUpperCase();
  const m = up.match(/\b(\d{2,3}(?:\.\d)?)\s*CM\b/) || up.match(/\b(\d{2}[A-K]{1,2})\b/);
  if (m) return m[1] + (/CM/.test(m[0]) ? 'CM' : '');
  for (const p0 of String((it.article || '') + ' · ' + (it.desc || '')).toUpperCase().split(/\s·\s/)) {
    const m2 = p0.replace(/[\s·,\/-]+$/, '').match(/(?:^|[\s\/·-])(XXS|XS|S|M|L|XL|XXL|XXXL|2XL|3XL|4XL|5XL)$/);
    if (m2) return m2[1];
  }
  return '';
}

/* ---------- tiêu đề cột (Excel và PDF dùng chung) ---------- */
const H_RE = {
  po: /^(P\.?\s*\/?\s*O\.?|PO)\s*[-#:.]?\s*(#|NO|NUMBER|NUM|\.|:)?\.?\s*[:#]?\s*$|^(CUST\.?\s*PO|CUSTOMER\s*PO|PURCHAS(E|ING)\s*(ORDER|NO)|ORDER\s*NO|P\/O\s*NO|PO\s*NUMBER|PO\s*-?\s*NO\.?|PO#|订单号|訂購單號)/i,
  qty: /QTY|QUANTITY|Q'TY|QUANTIT|QAUNTIT|数量|數量|TTL\s*\/?\s*M|^TTL$|^\(?(M|MTS?|MTRS?|METERS?|METRES?|Y|YDS?|YARDS?|PCS?|PRS|PAIRS?)\.?\)?$|^YARDAGE|LENGTH|DELIVERY\s*QTY|DELIVER$|SETTLEMENT|TOTAL\s*-?\s*NO\.?\s*OF\s*-?\s*UNITS|^UNITS?\s*\(|SỐ LƯỢNG/i,
  qtyNot: /CTN|CARTON|ROLL|PKG|PCKG|PACKAGE|BOX|N\.?\s*W|G\.?\s*W|WEIGHT|CBM|PRICE|AMOUNT|VOLUME|MEAS|\/CTN|PER\s*CTN|ORDER\s*QTY|Q\.?TY\s*ORDER|箱数|PCS\/|单件|BALE/i,
  unit: /^UNITS?$|^UOM$|^U\/M$|^UNIT\s*$|^单位|^單位/i,
  price: /PRICE|U\/P|UNIT\s*PR|單價|单价/i,
  amount: /AMOUNT|^TOTAL\s*\(?(USD|US\$)|VALUE|總价|总价|金额/i,
  sur: /SURCHARGE|MCQ|MOQ\s*CHARGE/i,
  color: /^COLOU?R(\s*NAME)?$|^CLR$|^SHADE$|COLOU?R\s*\/\s*|^CUST\.?\s*COLOU?R|COLOU?R\s*NAME|颜色|顏色|^MÀU/i,
  ccode: /COLOU?R\s*(CODE|NO)|^COLOR\s*NO|色号|色號|DYE\s*LOT|LOT\s*NO|^LOT$|^BATCH|缸号/i,
  size: /^SIZE$|^SIZES?\b|^DIMENSION|尺碼|尺码|^SPEC\.?$/i,
  art: /^ITEM(\s*NO\.?)?$|^ITEM\s*(NO|NUMBER|CODE|REF)|ARTICLE|^ART\.?\s*(NO|-NO)?\.?$|^MODEL|^STYLE|^DESCRIPTION|DESCRIPTION\s*OF|^DESCRIP|^NAME\s*OF|^MATERIAL|^DESIGN|^QUALITY|^SUPPLIER\s*ARTICLE|^SUPP\.?\s*REF|^CUST\.?\s*(NO|REF)|^CUSTOMER\s*REF|^OUR\s*QUALITY|货号|貨號|品名|^ITEM\s*NAME|^MÃ HÀNG|^TÊN HÀNG/i,
  desc: /DESCRIPTION|^DESC|名称|^COMP/i,
  lot: /LOT\s*NO|^LOT$|BATCH|DYE\s*LOT|缸号|^LOT\b/i,
  roll: /^ROLL|ROLL\s*(NO|#)|卷号|^PEZZA/i,
  ctn: /^CTN|CARTON|^C\/NO|^BOX|^PACK(AGE|ING)?\s*(NO|#)|箱|^BALE|^CASE|PCKGS|PKGS|NO\.?\s*OF\s*-?\s*(PACK|CTN|CARTON|BOX|BALE)/i,
  mat: /MATERIAL\s*CODE|CUSTOMER\s*REF|CUST\.?\s*NO|YOUR\s*ITEM|SCAVI\s*CODE|客戶貨號/i,
};
const H_ANY = /PO|QTY|QUANTITY|Q'TY|COLOU?R|CLR|ITEM|ARTICLE|DESCRIPTION|MODEL|STYLE|UNIT|PRICE|AMOUNT|SIZE|LOT|ROLL|CTN|CARTON|MATERIAL|数量|數量|货号|颜色|订单|訂購|TTL|METER|YARD|N\.?W|G\.?W|WEIGHT|BATCH|SHADE|MARKS/i;

/* chấm điểm một dòng tiêu đề: cần ≥ 2 cột nhận ra, trong đó phải có cột số lượng */
function headerScore(cells) {
  let n = 0, hasQty = false;
  for (const c of cells) {
    const t = String(c || '').trim();
    if (!t || t.length > 60) continue;
    if (H_RE.qty.test(t) && !H_RE.qtyNot.test(t)) { hasQty = true; n++; continue; }
    if (H_RE.po.test(t) || H_RE.color.test(t) || H_RE.art.test(t) || H_RE.price.test(t) || H_RE.amount.test(t) || H_RE.size.test(t) || H_RE.unit.test(t) || H_RE.lot.test(t) || H_RE.roll.test(t) || H_RE.ctn.test(t)) n++;
  }
  return hasQty ? n : 0;
}

/* đơn vị ghi ngay trong tiêu đề số lượng: "QTY/M", "Q'ty(Y)", "LENGTH(YARDS)", "Quantity(PRS)" */
function unitFromHeader(t) {
  const s = String(t == null ? '' : t).toUpperCase();
  let m = s.match(/[\/(（]\s*(M|MTS|MTRS?|METERS?|METRES?|Y|YDS?|YARDS?|PCS?|PRS|PAIRS?|SETS?|KGS?|DZ|DOZ|對|对)\s*[)）]?\s*$/);
  if (m) return genUnit(m[1]);
  m = s.match(/\b(METERS?|METRES?|YARDS?|YDS|PCS|PAIRS?|PRS|KGS)\b/);
  if (m) return genUnit(m[1]);
  return '';
}

/* ô tiêu đề: ô gộp ngang chỉ tính ở ô đầu (ô chủ) để mỗi cột tiêu đề chỉ nhận một lần */
function TH(ws, r, c) {
  try {
    const cell = ws.getCell(r, c);
    if (cell.isMerged && cell.master && (Number(cell.master.row) !== Number(r) || Number(cell.master.col) !== Number(c))) return '';
    return cellText(cell.value).trim();
  } catch (e) { return ''; }
}

/* số trong ô: công thức đơn giản cùng dòng (=E54*F54, =F54) được tính lại từ ô gốc,
   vì kết quả lưu trong file đôi khi cũ hơn số đã sửa tay */
function cellNum(ws, r, c, vn) {
  const raw = V(ws, r, c);
  if (raw && typeof raw === 'object' && raw.formula) {
    const f = String(raw.formula).replace(/^=/, '').replace(/\$/g, '').trim().toUpperCase();
    let m = f.match(/^([A-Z]{1,2})(\d+)\s*([*+\-\/])\s*([A-Z]{1,2})(\d+)$/);
    const colN = (L) => L.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
    if (m && Number(m[2]) === r && Number(m[5]) === r) {
      const a = genNum(V(ws, r, colN(m[1])), vn), b = genNum(V(ws, r, colN(m[4])), vn);
      if (!isNaN(a) && !isNaN(b)) return m[3] === '*' ? a * b : m[3] === '+' ? a + b : m[3] === '-' ? a - b : (b ? a / b : NaN);
    }
    m = f.match(/^([A-Z]{1,2})(\d+)$/);
    if (m && Number(m[2]) === r) return genNum(V(ws, r, colN(m[1])), vn);
  }
  return genNum(raw, vn);
}

/* giá / thành tiền: ô ghi "USD" còn số nằm ở ô kế bên (ô gộp lệch) */
function numNear(ws, r, c, vn) {
  const n = cellNum(ws, r, c, vn);
  if (!isNaN(n)) return n;
  const t = T(ws, r, c);
  if (t && !/^(USD|US\$|\$|VND|RMB|EUR|CNY|HKD|JPY|THB)\.?$/i.test(t)) return NaN;
  for (const dc of [1, -1, 2]) { const n2 = cellNum(ws, r, c + dc, vn); if (!isNaN(n2)) return n2; }
  return NaN;
}

/* ---------- Excel: tìm bảng trong mỗi sheet ---------- */
const GEN_STOP = /^\*?\s*(GRAND\s*)?TOTAL\b|^TTL\b|^SUB\s*-?\s*TOTAL|^TOTAL\s*[:：]|合計|合计|总计|總計|^Tổng\s*cộng|^Tot(ale)?\b|^Σ/i;
const GEN_END = /BANK\s*(DETAIL|INFO|NAME)|BENEFICIAR|SWIFT|AMOUNT\s*IN\s*WORDS|SAY\s+(TOTAL\s+)?(U\.?S|US\s*D|DOLLAR)|\bIBAN\b|A\/C\s*(NO|NAME)|ACCOUNT\s*NO/i;

function genSheetText(ws, maxRows) {
  let t = '';
  const lim = Math.min(maxRows || 40, ws.rowCount);
  for (let r = 1; r <= lim; r++) for (let c = 1; c <= 30; c++) t += ' ' + T(ws, r, c);
  return t;
}

/* tìm "INVOICE NO" / "DATE" trong các ô (nhãn + giá trị cùng ô, ô bên phải, hoặc ô dưới) */
function genLabelValue(ws, labelRe, valueOk, maxRows) {
  const lim = Math.min(maxRows || 45, ws.rowCount);
  const labelLike = (v) => (/^[A-Z .&()\/]+[:：]?$/i.test(v) && /INVOICE|DATE|NO\b|TERM|PAYMENT|REV|PO\b|REF|TOTAL/i.test(v))
    || (labelRe.test(v) && !String(v).replace(labelRe, '').replace(/[\s:：.#\-]+/g, '').match(/[A-Z0-9]{4,}/i))
    || /^(INVOICE|DATE)\b.*[:：]\s*$/i.test(v.replace(/\s+/g, ' '));
  const okv = (v) => v && !labelLike(v) && valueOk(v);
  for (let r = 1; r <= lim; r++) {
    for (let c = 1; c <= 30; c++) {
      const raw = V(ws, r, c);
      const t = cellText(raw).replace(/\s+/g, ' ').trim();
      if (!t || !labelRe.test(t)) continue;
      /* cùng ô: "INVOICE NO : CELEB260807-3" · "NO:20260917" · "Invoice No. IV260929" */
      const m = t.match(labelRe);
      let rest = t.slice(m.index + m[0].length).replace(/^[\s:：.#\-]+/, '').trim();
      rest = rest.replace(/\s+(DATE|PAYMENT|TERM|REV|ETD|PO\s*NO|DATED).*$/i, '').trim();
      if (okv(rest)) return { value: rest, raw: rest, r, c };
      const cands = [];
      for (let cc = c + 1; cc <= Math.min(c + 8, 30) && cands.length < 2; cc++) {
        const rv = V(ws, r, cc); const tv = cellText(rv).trim();
        if (tv) cands.push({ tv, rv, r, c: cc });
      }
      const bv = V(ws, r + 1, c); const tb = cellText(bv).trim();
      if (tb) cands.push({ tv: tb, rv: bv, r: r + 1, c });
      /* nhãn ghép "No. & date": số ở ô dưới, ngày ở ô bên phải của ô dưới */
      for (let cc = c + 1, n = 0; cc <= Math.min(c + 6, 30) && n < 2; cc++) {
        const rv = V(ws, r + 1, cc); const tv = cellText(rv).trim();
        if (tv) { cands.push({ tv, rv, r: r + 1, c: cc }); n++; }
      }
      for (const k of cands) if (okv(k.tv) || (k.rv instanceof Date && !labelLike(k.tv))) return { value: k.tv, raw: k.rv, r: k.r, c: k.c };
    }
  }
  return null;
}

const INV_NO_RE = /(?:COMMERCIAL\s+)?INVOICE\s*(?:NO|NUMBER|#|№)\.?|INV\.?\s*(?:NO|#)\.?|(?:^|\b)NO\.?\s*&\s*DATE\s*OF\s*(?:ORDER|INVOICE)\s*(?:\(INV\))?|CI\s*NO\.?|SỐ\s*HÓA\s*ĐƠN|HÓA\s*ĐƠN\s*SỐ/i;
const INV_NO_RE2 = /^S\.?O\.?\s*(?:NO|NUMBER|#)\.?|^NO\s*[.:：]\s*(?=[A-Z0-9][A-Z0-9\-\/]{3,}\s*$)|^NO\.?\s*[:：]?\s*$|PACKING\s*NO\.?|DELIVERY\s*NO/i;
const DATE_RE = /\bDATE\b|DATED|NGÀY|日期/i;
const invNoOk = (v) => /[A-Z0-9]{4,}/i.test(v) && !/^(AS\s+BELOW|DETAIL)/i.test(v) && v.length <= 32 && !/,/.test(v) && /^[A-Z0-9][A-Z0-9\-\/().#_ ]*$/i.test(v) && !/^(TEL|FAX|ADD)/i.test(v) && (v.match(/ /g) || []).length <= 2 && !/^(?:[A-Z][A-Z&]{2}\d{7}|[A-Z]{2}\d{8})[A-Z]{0,4}$/i.test(v);

function readGenXlsx(wb, opts) {
  opts = opts || {};
  const yearHint = yearHintOf(wb);
  const docs = [];
  for (const ws of wb.worksheets) {
    if (!ws.rowCount) continue;
    const name = String(ws.name || '').toUpperCase();
    const top = genSheetText(ws, 12).toUpperCase();
    /* dòng tiêu đề: điểm cao nhất trong 60 dòng đầu */
    let hr = 0, hs = 0;
    for (let r = 1; r <= Math.min(60, ws.rowCount); r++) {
      const cells = []; for (let c = 1; c <= 30; c++) cells.push(TH(ws, r, c));
      const s = headerScore(cells);
      /* cùng điểm → lấy dòng sau (tiêu đề con nằm sát dữ liệu hơn) */
      if (s > hs || (s === hs && s >= 2 && r - hr <= 12)) { hs = s; hr = r; }
      if (hs >= 5) break;
    }
    if (!hr || hs < 2) continue;
    const C = {};
    const used = new Set();
    const hdr = (c) => (TH(ws, hr, c) + ' ' + TH(ws, hr + 1, c)).replace(/\s+/g, ' ').trim();
    const take = (key, re, not) => {
      for (let c = 1; c <= 30; c++) {
        if (used.has(c)) continue;
        const t = TH(ws, hr, c).replace(/\s+/g, ' ').trim();
        if (!t) continue;
        if (re.test(t) && !(not && not.test(t))) { C[key] = c; used.add(c); return c; }
      }
      return 0;
    };
    take('qty', H_RE.qty, H_RE.qtyNot);
    take('po', H_RE.po);
    /* hai cột cùng tên "PO No" (mã nội bộ của chủ hàng + mã SAP) → chọn cột có mã SAP */
    if (C.po) {
      const sapCount = (c) => { let n = 0; for (let r = hr + 1; r <= Math.min(hr + 40, ws.rowCount); r++) if (genPos(T(ws, r, c)).length) n++; return n; };
      let bestC = C.po, bestN = sapCount(C.po);
      for (let c = 1; c <= 30; c++) {
        if (c === C.po || used.has(c)) continue;
        const t = TH(ws, hr, c).replace(/\s+/g, ' ').trim();
        if (t && H_RE.po.test(t)) { const n = sapCount(c); if (n > bestN) { bestN = n; bestC = c; } }
      }
      if (bestC !== C.po) { used.delete(C.po); C.poAlt = C.po; C.po = bestC; used.add(bestC); }
    }
    take('mat', H_RE.mat);
    take('ccode', H_RE.ccode);
    take('color', H_RE.color);
    take('size', H_RE.size);
    take('unit', H_RE.unit);
    take('price', H_RE.price);
    take('sur', H_RE.sur);
    take('amount', H_RE.amount);
    take('lot', H_RE.lot);
    take('roll', H_RE.roll);
    take('ctn', H_RE.ctn);
    /* cột mã hàng: có thể nhiều cột (ITEM + DESCRIPTION + STYLE…) */
    C.arts = [];
    for (let c = 1; c <= 30; c++) {
      if (used.has(c)) continue;
      const t = TH(ws, hr, c).replace(/\s+/g, ' ').trim();
      if (t && H_RE.art.test(t)) { C.arts.push(c); used.add(c); }
    }
    if (!C.qty) continue;
    /* cột "Item"/"No." chỉ đánh số thứ tự → không phải mã hàng */
    C.arts = C.arts.filter((c) => {
      let n = 0, ints = 0;
      for (let r = hr + 1; r <= Math.min(hr + 40, ws.rowCount); r++) { const t = T(ws, r, c); if (!t) continue; n++; if (/^\d{1,3}(\.0)?$/.test(t)) ints++; }
      return !(n >= 2 && ints / n >= 0.8);
    });
    const known = new Set([C.qty, C.po, C.mat, C.ccode, C.color, C.size, C.unit, C.price, C.amount, C.sur, C.lot, C.roll, C.ctn].filter(Boolean).concat(C.arts));
    let unit = unitFromHeader(TH(ws, hr, C.qty)) || unitFromHeader(TH(ws, hr + 1, C.qty)) || unitFromHeader(hdr(C.qty));
    const vn = /SỐ LƯỢNG|ĐƠN GIÁ/i.test(top);
    /* cột số lượng gần như trống nhưng cột kề bên lại có số ở hầu hết dòng (Qty/CTN khi mỗi dòng = 1 thùng) → dùng cột kề */
    {
      const cnt = (c) => { let n = 0; for (let r = hr + 1; r <= Math.min(hr + 80, ws.rowCount); r++) { const v = genNum(V(ws, r, c), vn); if (!isNaN(v) && v > 0) n++; } return n; };
      const own = cnt(C.qty);
      let bestC = 0, bestN = own;
      for (const dc of [-1, 1, -2, 2]) {
        const c2 = C.qty + dc; if (c2 < 1 || c2 > 30 || known.has(c2)) continue;
        const n2 = cnt(c2);
        if (n2 > bestN * 2 && n2 >= 3) { bestN = n2; bestC = c2; }
      }
      if (bestC) { known.delete(C.qty); C.qtyAlt = C.qty; C.qty = bestC; known.add(bestC); }
    }
    const isPkl = /PACK|PKL|^PL\b|装箱|碼單|码单|DPL/.test(name) || (!C.price && !C.amount);
    const isInv = !isPkl && (!!C.price || !!C.amount || /INV|发票|發票|CI$/.test(name));
    const items = [];
    let sticky = { po: '', art: '', color: '', size: '', lot: '' };
    let blank = 0, skippedQty = 0;
    for (let r = hr + 1; r <= ws.rowCount; r++) {
      const cells = []; for (let c = 1; c <= 30; c++) cells.push(T(ws, r, c));
      const rowTxt = cells.join(' | ');
      if (!rowTxt.replace(/\|/g, '').trim()) { if (++blank > 25) break; continue; }
      blank = 0;
      const first = cells.find((x) => x) || '';
      if (GEN_END.test(rowTxt)) break;
      if (first.length > 160 || /^(NOTE|REMARKS?|GHI CHÚ)\s*[:：]/i.test(first)) continue;
      if (GEN_STOP.test(first) || cells.slice(0, 8).some((x) => GEN_STOP.test(x))) {
        /* dòng tổng: nếu chưa có dòng hàng nào thì bỏ qua, có rồi vẫn đi tiếp (bảng có tổng phụ) */
        continue;
      }
      let qRaw = V(ws, r, C.qty);
      let qShift = 0, qCol = C.qty;
      if ((qRaw == null || qRaw === '') && C.qty > 2) {
        /* ô gộp lệch cột: số lượng nằm ở ô bên cạnh, có đơn vị ngay sau (AIM: I=3792 | J=YDS, tiêu đề ở K) */
        for (const dc of [-2, -1, 1, 2]) {
          const c2 = C.qty + dc; if (c2 < 1 || c2 > 30 || known.has(c2)) continue;
          const v2 = V(ws, r, c2); const n2 = genNum(v2, vn);
          if (!isNaN(n2) && n2 > 0 && genUnit(cells[c2])) { qRaw = v2; qShift = dc; qCol = c2; break; }
        }
      }
      const q = Math.round(cellNum(ws, r, qCol, vn) * 1000) / 1000;
      if (/GROSS\s*WEIGHT|NET\s*WEIGHT|\bROLLS?\b|MEASUREMENT|\bCBM\b|CARTONS?\b|PACKAGES?\b|\bBALES?\b|\bCTNS?\b/i.test(cells.filter((t, i) => i + 1 !== C.qty).join(' ')) && !(C.color && cells[C.color - 1]) && !(C.po && cells[C.po - 1])) continue;
      /* ô chữ không thuộc cột nào đã biết → thêm vào mô tả (cột không có tiêu đề) */
      const extra = cells.map((t, i) => (known.has(i + 1) || !t || /^[\d.,\s%$-]+$/.test(t) || /^\d+(\.\d+)?\s*[X×*]\s*\d+/i.test(t) || genUnit(t) || /^(USD|US\$|VND|RMB|EUR)$/i.test(t) ? '' : t)).filter(Boolean).join(' ');
      /* đơn vị ghi ở ô ngay sau số lượng khi không có cột đơn vị */
      const unitNext = !C.unit && C.qty + qShift < 30 ? genUnit(cells[C.qty + qShift]) : '';
      /* kéo xuống giá trị ô trống từ dòng trên (ô gộp / để trống cho gọn) */
      const poTxt = C.po ? cells[C.po - 1] : '';
      const artTxt = C.arts.map((c) => cells[c - 1]).filter(Boolean).join(' · ');
      const colTxt = C.color ? cells[C.color - 1] : '';
      const sizeTxt = C.size ? cells[C.size - 1] : '';
      const lotTxt = C.lot ? cells[C.lot - 1] : (C.ccode && H_RE.lot.test(hdr(C.ccode)) ? cells[C.ccode - 1] : '');
      if (poTxt) sticky.po = poTxt;
      else if (!C.po) { const pr = genPos(rowTxt); if (pr.length) sticky.po = pr.join(' / '); }
      if (artTxt) sticky.art = artTxt;
      if (colTxt) sticky.color = colTxt;
      else if (!C.color && extra && !/^[\d.,\s%$-]+$/.test(extra)) sticky.color = extra;
      if (sizeTxt) sticky.size = sizeTxt;
      if (lotTxt) sticky.lot = lotTxt;
      if (isNaN(q) || q <= 0) continue;
      /* dòng chỉ có mỗi con số lẻ loi (tổng phụ) → bỏ; nhưng dòng cuộn chỉ gồm số (STT cuộn + số lượng + cân nặng) thì giữ */
      const ownText = cells.filter((t, i) => t && i + 1 !== C.qty && !/^[\d.,\s%$-]+$/.test(t));
      const nCells = cells.filter((t) => t).length;
      if (!ownText.length && (nCells < 3 || !(C.roll || C.lot))) { skippedQty += q; continue; }
      if (!sticky.po && !sticky.art && !sticky.color && !colTxt && !artTxt && !extra) continue;
      /* bảng có cột PO + cột giá: dòng không PO, không giá là tiêu đề nhóm / tổng phụ */
      if (C.po && C.price && !poTxt && isNaN(genNum(V(ws, r, C.price), vn)) && !colTxt) { skippedQty += q; continue; }
      if (/SURCHARGE|BANK\s*CHARGE|FREIGHT|DISCOUNT|DYEING|LEFTOVER|MOQ|MCQ|TRANSATION|TRANSACTION/i.test(artTxt + ' ' + first + ' ' + colTxt) && !colTxt) { skippedQty += q; continue; }
      const pos = genPos(sticky.po + ' ' + (C.po ? '' : rowTxt));
      const mats = genMats(rowTxt + ' ' + sticky.art + ' ' + (C.mat ? cells[C.mat - 1] : ''));
      const u = (C.unit ? genUnit(cells[C.unit - 1]) : '') || unitNext || unit;
      const ccode = C.ccode ? cells[C.ccode - 1] : '';
      const colorText = (colTxt || sticky.color).replace(/[\r\n]+/g, ' ').trim();
      const art = (artTxt || sticky.art).replace(/[\r\n]+/g, ' ').replace(/^[A-Z .\/]{2,20}[:：]\s*(·\s*)?/i, '').trim();
      items.push({
        poRaw: sticky.po, po: pos[0] || '', poList: pos.length > 1 ? pos : null,
        poSapLike: pos.length > 0,
        article: art.split(/\s·\s/)[0], desc: [art, colorText, ccode, sticky.lot, extra].filter(Boolean).join(' · '),
        colorText, colorCode: ccode, colorShort: colorText,
        size: sizeTxt || sticky.size, lot: lotTxt || sticky.lot, mats, material: mats[0] || '',
        qty: q, unit: u, qtyByUnit: u ? { [u]: q } : {},
        price: C.price ? numNear(ws, r, C.price, vn) : NaN,
        amount: C.amount ? numNear(ws, r, C.amount, vn) : NaN, surcharge: C.sur ? (genNum(V(ws, r, C.sur), vn) || 0) : 0,
        roll: C.roll ? cells[C.roll - 1] : '', ctn: C.ctn ? cells[C.ctn - 1] : '',
        code: art || colorText, ctx: [sticky.po, art, colorText, ccode, sizeTxt || sticky.size, extra].filter(Boolean).join(' · '), sheet: ws.name, row: r, gen: true,
      });
    }
    if (!items.length) continue;
    /* số hoá đơn + ngày: tìm trong sheet (ưu tiên phía trên bảng) */
    const no = genLabelValue(ws, INV_NO_RE, invNoOk, hr + 2) || genLabelValue(ws, INV_NO_RE2, invNoOk, hr + 2);
    const dt = genLabelValue(ws, DATE_RE, (v) => !!genDate(v, yearHint), hr + 2);
    const total = (() => {
      const last = items[items.length - 1].row;
      for (let r = last + 1; r <= Math.min(last + 12, ws.rowCount); r++) {
        const f = (T(ws, r, 1) + ' ' + T(ws, r, 2) + ' ' + T(ws, r, 3) + ' ' + T(ws, r, 4) + ' ' + T(ws, r, 5) + ' ' + T(ws, r, 6)).trim();
        if (GEN_STOP.test(f) && !/WEIGHT|ROLLS|MEASURE|CBM|PACKAGE|CARTON/i.test(f)) return { amount: C.amount ? genNum(V(ws, r, C.amount), vn) : NaN, qty: genNum(V(ws, r, C.qty), vn) };
      }
      return { amount: NaN, qty: NaN };
    })();
    const hasSapPo = items.some((x) => x.poSapLike);
    const sumQ = items.reduce((a, b) => a + b.qty, 0);
    if (!isNaN(total.qty) && Math.abs(total.qty - sumQ - skippedQty) < 0.01) total.qty = Math.round(sumQ * 1000) / 1000;
    docs.push({
      sheet: ws.name, role: isInv ? 'inv' : (isPkl ? 'pkl' : 'inv'), items, unit, vn, hasSapPo,
      invNo: no ? String(no.value).trim() : '', invDate: dt ? (dt.raw instanceof Date || typeof dt.raw === 'number' ? parseDateCell(dt.raw, yearHint) : genDate(dt.value, yearHint)) : '',
      total: total.amount, totalQty: total.qty, hasPrice: !!C.price,
    });
  }
  return docs;
}

/* ---------- gộp các bảng trong 1 workbook thành chứng từ ---------- */
function genFromDocs(docs, fname, dir, text) {
  if (!docs.length) return null;
  const invD = docs.filter((d) => d.role === 'inv');
  const pklD = docs.filter((d) => d.role === 'pkl');
  /* nhiều bảng hoá đơn (FUJIANHONG: Page 2 = packing, Page 3 = invoice có giá) → lấy bảng có giá */
  let inv = invD.find((d) => d.hasPrice) || invD[0] || null;
  let no = (inv && inv.invNo) || docs.map((d) => d.invNo).find(Boolean) || '', noFrom = '';
  if (!no) { const f = invNoFromName(fname, dir, text); if (f) { no = f.no; noFrom = f.from; } }
  const invDate = (inv && inv.invDate) || docs.map((d) => d.invDate).find(Boolean) || '';
  const groups = [];
  for (const d of pklD) {
    for (const it of d.items) {
      const key = (it.po || AZ(it.poRaw)) + '|' + AZ(it.article) + '|' + AZ(it.colorText) + '|' + AZ(it.lot) + '|' + AZ(it.size);
      let g = groups.find((x) => x.key === key);
      if (!g) { g = { key, po: it.po, poRaw: it.poRaw, article: it.article, color: it.colorText, size: it.size, lot: it.lot, unit: it.unit, rolls: [], total: 0 }; groups.push(g); }
      g.rolls.push({ no: it.roll || it.ctn || String(g.rolls.length + 1), qty: it.qty });
      g.total = Math.round((g.total + it.qty) * 1000) / 1000;
    }
  }
  if (!inv && !groups.length) return null;
  let items = inv ? inv.items : groups.map((g) => ({
    poRaw: g.poRaw || g.po, po: g.po, poList: null, poSapLike: !!g.po, article: g.article, desc: [g.article, g.color, g.size, g.lot].filter(Boolean).join(' · '),
    colorText: g.color, colorCode: '', colorShort: g.color, size: g.size, lot: g.lot, mats: genMats(g.article + ' ' + g.color), material: '',
    qty: g.total, unit: g.unit, qtyByUnit: g.unit ? { [g.unit]: g.total } : {}, price: NaN, amount: NaN, surcharge: 0,
    code: [g.article, g.color].filter(Boolean).join(' · '), ctx: [g.po, g.article, g.color, g.size, g.lot].filter(Boolean).join(' · '), gen: true, fromPkl: true,
  }));
  /* hoá đơn ghi gộp theo mã hàng (không size / không PO) còn packing list ghi từng PO + size + màu
     (SUNPO): nếu tổng packing list của mã đó đúng bằng tổng hoá đơn → tách dòng hoá đơn theo packing list */
  if (inv && groups.length && groups.some((g) => g.size || g.po)) {
    const byArt = new Map();
    for (const it of items) {
      if (docSizeOf(it) || !it.article) continue;
      const k = AZ(it.article);
      if (!byArt.has(k)) byArt.set(k, []);
      byArt.get(k).push(it);
    }
    const out = [];
    const replaced = new Set();
    for (const [k, its] of byArt) {
      const gs = groups.filter((g) => g.article && (AZ(g.article).includes(k) || k.includes(AZ(g.article))) && (g.size || g.po));
      if (!gs.length) continue;
      /* chỉ tách khi packing list cho thêm thông tin: size, hoặc PO mà hoá đơn không có */
      if (!gs.some((g) => g.size) && its.some((x) => x.po)) continue;
      const sumI = its.reduce((a, b) => a + b.qty, 0), sumG = gs.reduce((a, b) => a + b.total, 0);
      if (Math.abs(sumI - sumG) > 0.011) continue;
      const prices = [...new Set(its.map((x) => x.price).filter((v) => !isNaN(v)))];
      const price = prices.length === 1 ? prices[0] : NaN;
      /* gộp các thùng cùng PO + size + màu */
      const agg = [];
      for (const g of gs) {
        const key = (g.po || '') + '|' + AZ(g.size) + '|' + AZ(g.color);
        let a = agg.find((x) => x.key === key);
        if (!a) { a = { key, g, qty: 0 }; agg.push(a); }
        a.qty = Math.round((a.qty + g.total) * 1000) / 1000;
      }
      const base = its[0];
      for (const a of agg) {
        out.push(Object.assign({}, base, {
          poRaw: a.g.poRaw || base.poRaw, po: a.g.po || base.po, poList: null, poSapLike: !!(a.g.po || base.poSapLike),
          size: a.g.size || '', colorText: a.g.color || base.colorText, colorShort: a.g.color || base.colorShort,
          lot: a.g.lot || '', qty: a.qty, qtyByUnit: (a.g.unit || base.unit) ? { [a.g.unit || base.unit]: a.qty } : {}, unit: a.g.unit || base.unit,
          price, amount: isNaN(price) ? NaN : Math.round(price * a.qty * 10000) / 10000,
          desc: [base.article, a.g.color, a.g.size].filter(Boolean).join(' · '), ctx: [a.g.po, base.article, a.g.color, a.g.size].filter(Boolean).join(' · '),
          fromPkl: true, warn: ['tách theo packing list (hoá đơn ghi gộp ' + base.article + ')'],
        }));
      }
      its.forEach((x) => replaced.add(x));
    }
    if (replaced.size) items = items.filter((x) => !replaced.has(x)).concat(out);
  }
  const out = {
    profile: 'GEN', supplier: '', role: inv ? 'inv' : 'pkl', file: fname || '',
    inv: {
      no, invNo: no, invDate, items, currency: 'USD', unitDefault: inv ? inv.unit : (pklD[0] && pklD[0].unit) || '',
      explodedByPkl: items.some((x) => x.fromPkl),
      surchargeHeader: 0, total: inv ? inv.total : NaN, totalQty: inv ? inv.totalQty : NaN, amountInclSur: items.some((x) => x.surcharge > 0),
      noInvoiceNo: !no, noFrom, gen: true, hasSapPo: inv ? inv.hasSapPo : docs.some((d) => d.hasSapPo),
    },
    pkl: groups.length ? { groups, unit: (pklD[0] && pklD[0].unit) || '', level: 'lot' } : null,
    pklOnly: !inv,
  };
  if (inv && inv.vn) out.inv.currency = 'VND';
  return out;
}

/* chứng từ đã đọc + packing list rời vừa gắn thêm → chạy lại bước tách dòng theo packing list */
function genFromDocsRefresh(fab) {
  if (!fab || !fab.inv || !fab.pkl || !fab.pkl.groups || !fab.pkl.groups.length) return null;
  if (fab.inv.explodedByPkl) return null;
  const items0 = fab.inv.items.filter((x) => !x.fromPkl);
  if (!items0.length) return null;
  const invDoc = { role: 'inv', items: items0, unit: fab.inv.unitDefault, vn: fab.inv.currency === 'VND', invNo: fab.inv.invNo, invDate: fab.inv.invDate, total: fab.inv.total, totalQty: fab.inv.totalQty, hasPrice: true, hasSapPo: fab.inv.hasSapPo };
  const pklDocs = fab.pkl.groups.map((g) => ({
    role: 'pkl', unit: g.unit, vn: false, invNo: '', invDate: '', total: NaN, totalQty: NaN, hasPrice: false,
    items: g.rolls.map((r) => ({ po: g.po, poRaw: g.poRaw, article: g.article, colorText: g.color, lot: g.lot, size: g.size, qty: r.qty, unit: g.unit, roll: r.no, poSapLike: !!g.po })),
  }));
  const re = genFromDocs([invDoc].concat(pklDocs), fab.file, '', '');
  if (!re || !re.inv.items.some((x) => x.fromPkl)) return null;
  return Object.assign({}, fab.inv, { items: re.inv.items, explodedByPkl: true });
}

function readGenWb(wb, fname, dir) {
  const docs = readGenXlsx(wb);
  let text = '';
  for (const ws of wb.worksheets) text += genSheetText(ws, 60);
  return genFromDocs(docs, fname, dir, text);
}

if (typeof module !== 'undefined' && module.exports) {
  Object.assign(module.exports, { genNum, genPos, genMats, genUnit, sizeSame, sizeKeyG, readGenXlsx, readGenWb, genFromDocs, genFromDocsRefresh, headerScore, unitFromHeader });
}

/* =====================================================================
   PDF — đọc chung
   pages = [[{x, y, w, s}, …], …] (text items của pdf.js, toạ độ gốc trái-dưới)
   ===================================================================== */
/* toạ độ của một text item pdf.js — trang xoay 90°/270° (PIONEER) thì đổi trục để dòng chữ nằm ngang */
function pdfItemXY(it) {
  const t = it.transform || [1, 0, 0, 1, 0, 0];
  const w = it.width || 0;
  if (Math.abs(t[0]) < 1e-6 && Math.abs(t[1]) > 1e-6) {
    return t[1] > 0 ? { x: t[5], y: -t[4], w, s: it.str } : { x: -t[5] - w, y: t[4], w, s: it.str };
  }
  return { x: t[4], y: t[5], w, s: it.str };
}

function genLinesOf(items) {
  const its = items.slice().sort((a, b) => (b.y - a.y) || (a.x - b.x));
  const lines = []; let cur = [], curY = null;
  const flush = () => {
    if (!cur.length) return;
    cur.sort((a, b) => a.x - b.x);
    const cells = cur.map((c) => ({ x: c.x, x1: c.x + (c.w || 0), s: String(c.s).replace(/\s+/g, ' ').trim() })).filter((c) => c.s);
    lines.push({ y: curY, cells, text: cells.map((c) => c.s).join(' ') });
    cur = [];
  };
  for (const it of its) {
    if (curY === null || Math.abs(it.y - curY) <= 3) { cur.push(it); if (curY === null) curY = it.y; }
    else { flush(); cur = [it]; curY = it.y; }
  }
  flush();
  return lines;
}

/* ngày trên PDF: "24-Sep-26" · "22/07/26" · "AUGUST 19, 2026" · "2026/09/10" · "dated Septenber 04th, 2026" */
const MONTH_FIX = { AGU: 'AUG', SEPT: 'SEP', SEPTENBER: 'SEP', OKT: 'OCT', MAI: 'MAY', JUNE: 'JUN', JULY: 'JUL' };
const monthIdx = (name) => { const k = String(name).toUpperCase(); const f = MONTH_FIX[k] || k; return MONTHS.findIndex((x) => f.startsWith(x)); };
function genDate(s, yearHint) {
  const t = String(s == null ? '' : s).trim().replace(/(\d)(ST|ND|RD|TH)\b/gi, '$1');
  let m = t.toUpperCase().match(/(\d{1,2})\s*[-\s\/.,]\s*([A-Z]{3,9})\s*[-\s\/.,]\s*(\d{2,4})\b/);
  if (m) {
    const mi = monthIdx(m[2]);
    if (mi >= 0) return dmy(m[1], mi + 1, m[3].length === 2 ? '20' + m[3] : m[3]);
  }
  m = t.toUpperCase().match(/([A-Z]{3,9})\.?\s*[-\/.,]?\s*(\d{1,2})\s*[-\/.,]?\s*(\d{4})\b/);
  if (m) { const mi = monthIdx(m[1]); if (mi >= 0) return dmy(m[2], mi + 1, m[3]); }
  m = t.match(/\b(\d{1,2})\s*[\/\-.]\s*(\d{1,2})\s*[\/\-.]\s*(\d{2})\b(?!\d)/);
  if (m && !/\d{4}/.test(t)) return dmy(m[1], m[2], '20' + m[3]);
  return parseDateCell(t, yearHint);
}

const RE_TOTAL_LINE = /^(GRAND\s*-?\s*)?TOTAL|^SUB\s*-?\s*TOTAL|\bTOTAL\s*(:|QTY|QUANTITY|AMOUNT|NET|GROSS|ROLLS|BAGS|PACKAGES|CARTONS|CBM|M3|WEIGHT)|^TOT(ALE)?\b|TOT\s+ROLLS|^SUMMARY|^SAY\b|^TTL\b|^合计|^合計|^总计|^Tổng\s*cộng/i;
const RE_SKIP_LINE = /DIFETTO|FAULT|TARIFF|ORIGIN|\b(DAYS?|PAYMENTS?|COMPLAINTS?|WITHIN|REMARKS?|SHIPPING\s*MARKS?|INTEREST|CHARGES?)\b|CUSTOMS?\b|\b(N\.?\s*W\.?|G\.?\s*W\.?|NET\s*WEIGHT|GROSS\s*WEIGHT|CBM|VOLUME|MEASUREMENT|DIMENSION)\b|^C\/NO|^CARTON|^ROLLS?\s*:|^BANK|SWIFT|^A\/C|ACCOUNT\s*NO|BENEFICIAR|^TEL|^FAX|^E-?MAIL|TAX\s*(ID|CODE)/i;
const RE_NUM_TOK = /^[-+]?\(?\d[\d.,]*\)?$/;
const numTok = (s) => RE_NUM_TOK.test(String(s).replace(/[^\d.,\-+()]/g, '') ? String(s).replace(/[^\d.,\-+()]/g, '') : '') && /\d/.test(s);

/* tách một ô PDF thành từ, giữ vị trí x ước lượng theo tỉ lệ chiều dài */
function wordsOf(line) {
  const out = [];
  for (const c of line.cells) {
    const parts = c.s.split(/\s+/).filter(Boolean);
    const len = Math.max(1, c.s.length);
    let pos = 0;
    for (const p of parts) {
      const i = c.s.indexOf(p, pos);
      const x0 = c.x + (c.x1 - c.x) * (i / len), x1 = c.x + (c.x1 - c.x) * ((i + p.length) / len);
      out.push({ s: p, x: x0, x1 });
      pos = i + p.length;
    }
  }
  return out;
}

/* số lượng trên một dòng:
   1) số đứng NGAY TRƯỚC hoặc NGAY SAU một từ đơn vị  ("309 YDS." · "MT 632,50" · "1,673.0 meters")
   2) nếu không: số nằm dưới cột tiêu đề số lượng (qx0..qx1)                                   */
function qtyOnLine(line, qcol, vn) {
  const w = wordsOf(line);
  /* số "giống số lượng": không có 0 đứng đầu (03261 là mã màu), không quá 7 chữ số liền (074178902 là số cuộn) */
  const pureNum = (s) => /^\(?[-+]?\d[\d.,]*\)?$/.test(s) && /\d/.test(s) && !/^0\d/.test(s) && !/^\d{8,}$/.test(s) && !/^\d{5,}$/.test(s.replace(/[.,]/g, '').length > 7 ? s : '');
  const inCol = (i) => {
    if (!qcol) return false;
    const ov = Math.min(w[i].x1, qcol.x1 + 6) - Math.max(w[i].x, qcol.x0 - 6);
    const d = Math.abs((w[i].x + w[i].x1) / 2 - (qcol.x0 + qcol.x1) / 2);
    return ov > 0 || d < 25;
  };
  const cands = [];
  for (let i = 0; i < w.length; i++) {
    const u = genUnit(w[i].s.replace(/[.()]/g, ''));
    if (!u || !GEN_UNIT_RE.test(w[i].s.replace(/[()]/g, ''))) continue;
    if (i > 0 && pureNum(w[i - 1].s)) { const q = genNum(w[i - 1].s, vn); if (!isNaN(q) && q > 0) cands.push({ q, unit: u, idx: i - 1, w, col: inCol(i - 1) }); }
    if (i + 1 < w.length && pureNum(w[i + 1].s)) { const q = genNum(w[i + 1].s, vn); if (!isNaN(q) && q > 0) cands.push({ q, unit: u, idx: i + 1, w, after: true, col: inCol(i + 1) }); }
  }
  if (cands.length) {
    const c = cands.find((x) => x.col) || cands[0];
    return c;
  }
  if (qcol) {
    let best = null;
    for (let i = 0; i < w.length; i++) {
      if (!pureNum(w[i].s) || !inCol(i)) continue;
      const d = Math.abs((w[i].x + w[i].x1) / 2 - (qcol.x0 + qcol.x1) / 2);
      if (!best || d < best.d) best = { i, d };
    }
    if (best) { const q = genNum(w[best.i].s, vn); if (!isNaN(q) && q > 0) return { q, unit: '', idx: best.i, w, byCol: true }; }
  }
  return null;
}

/* tiêu đề bảng trên PDF: dòng (hoặc 2 dòng liền nhau) có nhiều từ khoá cột; lấy vị trí x của cột số lượng */
function pdfHeader(lines, from) {
  let best = null;
  for (let i = from; i < Math.min(lines.length, from + 60); i++) {
    const cells = lines[i].cells.map((c) => c.s);
    let s = headerScore(cells);
    if (s < 2 && i + 1 < lines.length) s = Math.max(s, headerScore(cells.concat(lines[i + 1].cells.map((c) => c.s)).map((x) => x)) - 1);
    if (s >= 2 && (!best || s > best.s)) best = { i, s };
    if (best && best.s >= 4) break;
  }
  if (!best) {
    for (let i = from; i < Math.min(lines.length, from + 60); i++) {
      const cs = lines[i].cells;
      if (cs.length <= 6 && cs.some((c) => /^(METERS?|METRES?|QUANTITY|QTY|YARDS?|PCS|PAIRS?|PRS)$/i.test(c.s.trim()))) { best = { i, s: 1 }; break; }
    }
    if (!best) return null;
  }
  const L = lines[best.i];
  let qc = L.cells.find((c) => H_RE.qty.test(c.s) && !H_RE.qtyNot.test(c.s));
  let nxt = null;
  if (!qc && best.i + 1 < lines.length) { qc = lines[best.i + 1].cells.find((c) => H_RE.qty.test(c.s) && !H_RE.qtyNot.test(c.s)); if (qc) nxt = best.i + 1; }
  const unitTxt = qc ? qc.s + ' ' + (lines[best.i + 1] ? lines[best.i + 1].cells.filter((c) => Math.abs(c.x - qc.x) < 40).map((c) => c.s).join(' ') : '') : '';
  const both = L.cells.concat(lines[best.i + 1] ? lines[best.i + 1].cells : []);
  const colOf = (re, not) => { const c = both.find((x) => re.test(x.s) && !(not && not.test(x.s))); return c ? { x0: c.x, x1: c.x1 } : null; };
  const cols = { size: colOf(H_RE.size), color: colOf(H_RE.color, H_RE.ccode), po: colOf(H_RE.po), price: colOf(H_RE.price) };
  return { i: nxt != null ? nxt : best.i, qcol: qc ? { x0: qc.x, x1: qc.x1 } : null, unit: unitFromHeader(unitTxt), cols };
}

/* ô nằm dưới một cột tiêu đề (theo x) */
function cellUnder(line, col, maxLen) {
  if (!col) return '';
  const mid = (col.x0 + col.x1) / 2;
  const hit = line.cells.filter((c) => (c.x <= col.x1 + 8 && c.x1 >= col.x0 - 8) || Math.abs((c.x + c.x1) / 2 - mid) < 20)
    .sort((a, b) => Math.abs((a.x + a.x1) / 2 - mid) - Math.abs((b.x + b.x1) / 2 - mid));
  const v = hit.length ? hit[0].s.trim() : '';
  return maxLen && v.length > maxLen ? '' : v;
}

/* vai trò của một trang */
function pageRole(lines) {
  const heads = lines.slice(0, 25).map((l) => l.text.toUpperCase());
  const titleIs = (re) => heads.some((t) => re.test(t));
  if (titleIs(/HÓA ĐƠN GIÁ TRỊ GIA TĂNG|\(VAT INVOICE\)/)) return 'vat';
  if (titleIs(/PROFORMA\s+INVOICE/)) return 'proforma';
  const pkl = titleIs(/^(DETAILED\s+)?PACKING\s*(\/\s*WEIGHT\s*)?LISTS?\b|^DELIVERY\s+NOTE|PHIẾU ĐÓNG HÀNG|^PACKING LIST\s/);
  const inv = titleIs(/^(COMMERCIAL\s+)?INVOICE\s*$|^FATTURA\b|^INVOICE\s+(NO|NUMBER)|^INVOICE\s{2,}/) || heads.some((t) => /\bINVOICE\b/.test(t) && !/INVOICE\s*(NO|REFERENCE|NUMBER|#|:)/.test(t) && !/PACKING/.test(t));
  if (pkl && !inv) return 'pkl';
  if (inv && !pkl) return 'inv';
  if (inv && pkl) return heads.findIndex((t) => /PACKING\s*LIST|DELIVERY\s+NOTE/.test(t)) < heads.findIndex((t) => /INVOICE|FATTURA/.test(t)) ? 'pkl' : 'inv';
  return '';
}

/* nhãn ở dòng tiêu đề, giá trị ở dòng ngay dưới, cùng cột (CAPITAL: "INVOICE NO.  DATE" / "CT-26-314T  AUGUST 19, 2026") */
function valueBelow(lines, labelRe, ok) {
  for (let i = 0; i < Math.min(lines.length, 80); i++) {
    const lc = lines[i].cells.find((c) => labelRe.test(c.s));
    if (!lc) continue;
    const rest = lc.s.replace(labelRe, '').replace(/^[\s:：.#-]+/, '').trim();
    if (rest && ok(rest)) return rest;
    for (let j = i + 1; j <= Math.min(i + 2, lines.length - 1); j++) {
      const near = lines[j].cells.filter((c) => Math.abs(c.x - lc.x) < 45 || (c.x >= lc.x - 5 && c.x <= lc.x1 + 5)).sort((a, b) => Math.abs(a.x - lc.x) - Math.abs(b.x - lc.x));
      for (const c of near) { const v = c.s.trim(); if (ok(v)) return v; }
    }
  }
  return '';
}

/* số hoá đơn + ngày trên PDF */
function pdfInvNo(lines) {
  const txt = lines.map((l) => l.text).join('\n');
  const tries = [
    /INVOICE\s*(?:NO|NUMBER|#|№)\.?\s*(?:&\s*DATE)?\s*[:：]?\s*\n?\s*([A-Z0-9][A-Z0-9\-\/().]{3,})/i,
    /\bINVOICE\s*[:：]\s*([A-Z0-9][A-Z0-9\-\/().]{3,})/i,
    /\bINV\.?\s*(?:NO|#)\.?\s*[:：]?\s*([A-Z0-9][A-Z0-9\-\/().]{3,})/i,
    /CI\s*NO\.?\s*[:：]?\s*([A-Z0-9][A-Z0-9\-\/]{3,})/i,
    /COMMERCIAL\s*INVOICE\s*#\s*[:：]?\s*([A-Z0-9][A-Z0-9\-\/]{3,})/i,
    /(?:NR\.?\s*DOCUMENTO|DOCUMENT\s*No)[^\n]*\n[^\n]*?(\d{5,})\s*$/im,
    /(?:DELIVERY\s*NO|PACKING\s*LIST)\s*[:：]?\s+(\d{5,})/i,
  ];
  for (const re of tries) {
    const m = txt.match(re);
    if (m && /\d/.test(m[1]) && !/^(DATE|DATED|AND|NO|REFERENCE|NUMBER)$/i.test(m[1])) return m[1].replace(/[.:,]+$/, '');
  }
  const okNo = (v) => /^[A-Z0-9][A-Z0-9\-\/().]{3,}$/i.test(v) && !/^(DATE|DATED|INVOICE|NO)$/i.test(v) && /\d/.test(v);
  return valueBelow(lines, /^(COMMERCIAL\s+)?INVOICE\s*(NO|NUMBER|#)\.?\s*:?$|^NR\.?\s*DOCUMENTO|DOCUMENT\s*No\.?$|^INVOICE\s*:?$/i, okNo) || '';
}
function pdfInvDate(lines, yearHint) {
  const txt = lines.map((l) => l.text).join('\n');
  const ms = [
    /DATED?\s*[:：]?\s*((?:\d{1,2}[-\s\/.]*)?[A-Z]{3,9}\.?\s*\d{1,2}(?:ST|ND|RD|TH)?,?\s*\d{2,4})/i,
    /DATED?\s*[:：]?\s*(\d{1,2}\s*[-\/.]\s*[A-Z]{3,9}\s*[-\/.]\s*\d{2,4})/i,
    /DATED?\s*[:：]?\s*(\d{4}\s*[-\/.]\s*\d{1,2}\s*[-\/.]\s*\d{1,2})/i,
    /DATED?\s*[:：]?\s*(\d{1,2}\s*[-\/.]\s*\d{1,2}\s*[-\/.]\s*\d{2,4})/i,
    /\b(\d{1,2}\s*\/\s*\d{1,2}\s*\/\s*\d{2,4})\b/,
    /\b(\d{4}\s*[-\/]\s*\d{1,2}\s*[-\/]\s*\d{1,2})\b/,
    /\b(\d{1,2}-[A-Z]{3}-\d{4})\b/i,
    /\b([A-Z]{3,9}\.?\s+\d{1,2}(?:ST|ND|RD|TH)?,\s*\d{4})\b/i,
  ];
  for (const re of ms) { const m = txt.match(re); if (m) { const d = genDate(m[1], yearHint); if (d) return d; } }
  const v = valueBelow(lines, /^DATE\s*:?$|^DATA\s+DOCUMENTO|^DOCUMENT\s+DATE/i, (x) => !!genDate(x, yearHint));
  return v ? genDate(v, yearHint) : '';
}

/* số hoá đơn suy từ tên file / tên thư mục khi trong chứng từ không có nhãn:
   "MF2615049.PDF" (mã có trong chữ của file) · thư mục "invoice 100972 - PO BRU0001900" */
function invNoFromName(fname, dir, text) {
  const base = String(fname || '').replace(/\.[^.]+$/, '');
  const T = String(text || '').toUpperCase();
  const toks = base.toUpperCase().split(/[\s_,()\[\]#+]+/).filter((t) => /\d/.test(t) && /^[A-Z0-9][A-Z0-9\-\/.]{3,}$/.test(t) && !/^(PO|IB|INB)/.test(t) && !/^(20\d{6}|\d{1,2}[-.]\d{1,2}[-.]\d{2,4})$/.test(t) && !/^(?:[A-Z]{3}\d{7}|[A-Z]{2}\d{8})$/.test(t));
  for (const t of toks) if (T.includes(t)) return { no: t, from: 'tên file' };
  const d = String(dir || '').toUpperCase().match(/\bINV(?:OICE)?\s*[:#._-]?\s*([A-Z0-9][A-Z0-9\-\/.]{3,})/);
  if (d && /\d/.test(d[1]) && !/^(NO|NUMBER)$/.test(d[1])) return { no: d[1].replace(/[.:,]+$/, ''), from: 'tên thư mục' };
  return null;
}

/* ---------- hoá đơn GTGT Việt Nam: dòng hàng bắt đầu bằng STT ---------- */
const VN_UNIT = /^(M|MÉT|MET|CÁI|CAI|PCS|PC|CHIẾC|ĐÔI|DOI|CẶP|KG|KGS|YDS?|YARD|CUỘN|BỘ|SET|TẤM|CÂY|HỘP)$/i;
function readVatPdf(lines, fname) {
  const idx = [];
  for (let i = 0; i < lines.length; i++) {
    const w = wordsOf(lines[i]).map((x) => x.s);
    if (!/^\d{1,3}$/.test(w[0] || '')) continue;
    const ui = w.findIndex((t, k) => k > 0 && VN_UNIT.test(t));
    if (ui < 1) continue;
    const nums = w.slice(ui + 1).filter((t) => /^\d[\d.,]*$/.test(t));
    if (nums.length < 2) continue;
    idx.push({ i, w, ui, nums });
  }
  const items = [];
  for (let k = 0; k < idx.length; k++) {
    const { i, w, ui, nums } = idx[k];
    const prev = k ? idx[k - 1].i : -1, next = k + 1 < idx.length ? idx[k + 1].i : lines.length;
    const ctxLines = [];
    const sttLines = new Set(idx.map((x) => x.i));
    for (let j = Math.max(prev + 1, i - 3); j < Math.min(next, i + 4); j++) {
      if (j === i) continue;
      const t = lines[j].text;
      if (/Tổng|Thuế suất|Cộng tiền|Thành tiền trước|Số tiền viết|Ký bởi|Tra cứu|Mã của cơ quan|STT|\(No\.?\)|\(Unit\)|Đơn vị/i.test(t)) continue;
      const d = Math.abs(j - i);
      /* dòng cách 2–3 dòng mà lại sát dòng STT khác → thuộc mặt hàng kia */
      if (d >= 2 && (sttLines.has(j - 1) || sttLines.has(j + 1))) continue;
      ctxLines.push({ t, d });
    }
    ctxLines.sort((a, b) => a.d - b.d);
    const ctx = [w.join(' ')].concat(ctxLines.map((x) => x.t)).join(' \n ');
    const unit = genUnit(w[ui]) || (/^(CÁI|CAI|CHIẾC)$/i.test(w[ui]) ? 'PCS' : (/^(ĐÔI|DOI|CẶP)$/i.test(w[ui]) ? 'PAA' : AZ(w[ui])));
    const q = genNum(nums[0], true), price = genNum(nums[1], true), amount = nums.length > 2 ? genNum(nums[2], true) : NaN;
    const near = [w.join(' ')].concat(ctxLines.filter((x) => x.d <= 1).map((x) => x.t)).join(' ');
    const pos = genPos(near).length ? genPos(near) : genPos(ctx);
    const matsNear = genMats(near);
    const mats = matsNear.length ? matsNear : genMats(ctx);
    const descLine = ctxLines.filter((x) => x.d <= 1).map((x) => x.t).join(' ');
    items.push({
      poRaw: pos[0] || '', po: pos[0] || '', poList: pos.length > 1 ? pos : null, poSapLike: pos.length > 0,
      article: '', desc: ctx, colorText: ctx, colorCode: '', colorShort: descLine.slice(0, 60),
      size: '', lot: '', mats, matsFromCtx: !matsNear.length, material: mats[0] || '', qty: q, unit, qtyByUnit: unit ? { [unit]: q } : {},
      price, amount, surcharge: 0, code: descLine.slice(0, 40), ctx, vat: true, stt: w[0], gen: true,
    });
  }
  const txt = lines.map((l) => l.text).join('\n');
  const serial = (txt.match(/K[ýy]\s*hi[ệe]u[^:]*:\s*([A-Z0-9]+)/i) || [])[1] || '';
  const no = (txt.match(/S[ốo]\s*\(\s*No\.?\s*\)\s*:?\s*(\d+)/i) || txt.match(/\bNo\.?\s*\)\s*:?\s*(\d+)/i) || txt.match(/S[ốo]\s*:\s*(\d{3,})/i) || [])[1] || '';
  const md = txt.replace(/\s+/g, ' ').match(/Ng[àa]y\s*\(?\s*(?:date)?\s*\)?\s*(\d{1,2})\s*th[áa]ng\s*\(?\s*(?:month)?\s*\)?\s*(\d{1,2})\s*n[ăa]m\s*\(?\s*(?:year)?\s*\)?\s*(\d{4})/i);
  const invDate = md ? dmy(md[1], md[2], md[3]) : '';
  const tot = txt.match(/T[ổo]ng\s*c[ộo]ng\s*ti[ềe]n\s*ch[ưu]a\s*thu[ếe][^\d]*([\d.,]+)/i) || txt.match(/Thuế suất 8%[^\n]*?([\d.,]+)\s+[\d.,]+\s+[\d.,]+/i) || txt.match(/T[ổo]ng\s*c[ộo]ng\s*\(Total\)\s*:?\s*([\d.,]+)/i);
  const invNo = serial && no ? `${serial}#${no.padStart(8, '0')}` : (no || '');
  return {
    profile: 'GEN', supplier: '', role: 'inv', file: fname || '', vat: true,
    inv: { no: invNo, invNo, invDate, items, currency: 'VND', unitDefault: items[0] ? items[0].unit : 'M', surchargeHeader: 0,
      total: tot ? genNum(tot[1], true) : NaN, totalQty: NaN, amountInclSur: false, noInvoiceNo: !invNo, gen: true, serial, number: no, hasSapPo: items.some((x) => x.poSapLike) },
    pkl: null, pklOnly: false,
  };
}

/* ---------- PDF chung: khối theo PO, dòng hàng = dòng có số lượng ---------- */
function readGenPdf(pages, fname, dir) {
  const yearHint = new Date().getUTCFullYear();
  const perPage = pages.map((p) => genLinesOf(p));
  const roles = perPage.map((ls) => pageRole(ls));
  const allLines = perPage.flat();
  if (roles.some((r) => r === 'vat')) {
    const d = readVatPdf(perPage.filter((_, i) => roles[i] === 'vat' || roles[i] === '').flat(), fname);
    return d;
  }
  const hasInv = roles.includes('inv'), hasPkl = roles.includes('pkl');
  const role = hasInv ? 'inv' : (hasPkl ? 'pkl' : (roles.includes('proforma') ? 'proforma' : 'inv'));
  if (role === 'proforma') return { profile: 'GEN', role: 'proforma', file: fname, inv: { items: [], invNo: pdfInvNo(allLines) }, pkl: null };
  const vn = false;
  const readPages = (want) => {
    const items = [];
    let docUnit = '';
    let block = null;
    const blocks = [];
    for (let p = 0; p < perPage.length; p++) {
      if (!want(roles[p], p)) continue;
      const lines = perPage[p];
      const hd = pdfHeader(lines, 0);
      const start = hd ? hd.i + 1 : 0;
      if (hd && hd.unit && !docUnit) docUnit = hd.unit;
      /* chữ phía trên bảng (mã hàng, "Your item: MFKSJQSL0127001") làm ngữ cảnh cho khối đầu */
      const pre = [];
      for (let i = Math.max(0, start - 12); i < start; i++) { const t = lines[i].text; if (/[A-Z]{2,}\d{3,}|\d{3,}[A-Z]{2,}|\b[A-Z0-9]+[\/-][A-Z0-9]+\b/i.test(t) && !/^(TEL|FAX|TAX|E-?MAIL|ADDRESS|ADD)\b/i.test(t)) pre.push(t); }
      /* sang trang: giữ PO + ngữ cảnh đầu khối của trang trước (bảng cuộn kéo dài nhiều trang) */
      if (!block) { block = { po: '', pos: [], ctx: pre, items: [] }; blocks.push(block); }
      else { block = { po: block.po, pos: block.pos, ctx: pre, items: [], head: (block.head || []).concat(block.items.length ? [] : block.ctx) }; blocks.push(block); }
      for (let i = start; i < lines.length; i++) {
        const L = lines[i];
        const t = L.text;
        if (!t.trim()) continue;
        if (RE_TOTAL_LINE.test(t.trim())) { block = { po: block.po, pos: block.pos, ctx: [], items: [], afterTotal: true, head: block.head }; blocks.push(block); continue; }
        if (/^(PACKING\s*LIST|INVOICE|FATTURA|COMMERCIAL\s*INVOICE)\s*$/i.test(t.trim())) continue;
        const pos = genPos(t);
        const q = RE_SKIP_LINE.test(t) && !pos.length ? null : qtyOnLine(L, hd ? hd.qcol : null, vn);
        if (pos.length) {
          /* dòng có PO → mở khối mới (giữ lại các dòng ngay trước nó chưa thuộc dòng hàng nào).
             Dòng hàng vừa đứng ngay trên mà chưa có PO (ô PO canh giữa theo chiều dọc) → kéo sang khối mới */
          const carry = block.items.length ? [] : block.ctx;
          const nb = { po: pos[0], pos, ctx: carry.slice(), items: [], poLine: t };
          if (!block.po && block.items.length) {
            const last = block.items[block.items.length - 1];
            if (i - last.line <= 2) { block.items.pop(); last.po = last.poRaw = pos[0]; last.poList = pos.length > 1 ? pos : null; last.poSapLike = true; nb.items.push(last); }
          }
          block = nb;
          blocks.push(block);
        }
        if (q && !(q.unit === 'KG' && !/KG/.test(docUnit || ''))) {
          const w = q.w;
          const after = w.slice(q.idx + 1).map((x) => x.s).filter((s) => /^\(?[-+]?\d[\d.,]*\)?$/.test(s) && /\d/.test(s));
          let price = NaN, amount = NaN;
          if (after.length >= 2) {
            price = genNum(after[after.length - 2], vn); amount = genNum(after[after.length - 1], vn);
            if (isNaN(price) || isNaN(amount) || Math.abs(price * q.q - amount) > Math.max(1, Math.abs(amount) * 0.03)) { price = NaN; amount = NaN; }
          }
          const cols = hd ? hd.cols : null;
          const poCell = cols ? cellUnder(L, cols.po, 30) : '';
          const szCell = cols ? cellUnder(L, cols.size, 8) : '';
          const colorCell = cols ? cellUnder(L, cols.color, 60) : '';
          if (!block.items.length) block.head = (block.head || []).concat(block.ctx);
          const it = {
            poRaw: block.po, po: block.po, poList: block.pos.length > 1 ? block.pos : null, poSapLike: !!block.po,
            own: t, before: (block.head || []).concat(block.items.length ? block.ctx : []), after: [], qty: q.q, unit: q.unit || docUnit || '', price, amount, surcharge: 0,
            line: i, page: p, idxInBlock: block.items.length, byCol: !!q.byCol,
            sizeCell: /^[A-Z0-9.\/-]{1,8}$/i.test(szCell) && !/^\d+[.,]\d+$/.test(szCell) && szCell !== String(q.w[q.idx].s) ? szCell : '',
            poCell,
            colorCell: colorCell && !/^\(?[-+]?\d[\d.,]*\)?$/.test(colorCell) ? colorCell : '',
          };
          /* dòng chỉ có số (size + số lượng + giá) → thừa hưởng màu/mô tả của dòng hàng cùng khối */
          const bare = !/[A-Z]{3,}/i.test(t.replace(/US\$|USD|VND|RMB|EUR|\b(M|MT|MTRS?|YDS?|PCS?|PRS?|PAIRS?|SETS?|KGS?)\b/gi, ''));
          if (bare) {
            const prev = block.items.length ? block.items[block.items.length - 1] : null;
            if (prev) { it.inherit = prev.own + ' ' + (prev.inherit || ''); it.colorCell = it.colorCell || prev.colorCell; }
            else it.inheritNext = true;
          } else if (block.items.length && block.items[0].inheritNext) {
            const f = block.items[0]; f.inherit = t; f.colorCell = f.colorCell || it.colorCell; f.inheritNext = false;
          }
          /* ô dưới cột PO ghi mã riêng của chủ hàng (PT.WINNER1-0667) → không lấy PO của khối */
          if (it.poCell && /[A-Z]/i.test(it.poCell) && /\d/.test(it.poCell) && !genPos(it.poCell).length && !/^(HS|NO)\b/i.test(it.poCell)) { it.po = ''; it.poRaw = it.poCell; it.poList = null; it.poSapLike = false; }
          block.items.push(it); items.push(it);
          block.ctx = [];
        } else {
          block.ctx.push(t);
          if (block.items.length) block.items[block.items.length - 1].after.push(t);
        }
      }
      /* khối chỉ có 1 dòng hàng: toàn bộ chữ trong khối thuộc dòng hàng đó (DERUN ghi màu dưới số lượng) */
      for (const b of blocks) {
        if (b.items.length === 1) b.items[0].afterOK = true;
        else b.items.forEach((it, k) => { it.afterOK = (k === b.items.length - 1) && it.after.length <= 2; });
      }
    }
    return { items, docUnit };
  };
  const { items: raw, docUnit } = readPages((r) => (hasInv ? r === 'inv' : r !== 'proforma'));
  const unitVotes = {};
  raw.forEach((it) => { if (it.unit) unitVotes[it.unit] = (unitVotes[it.unit] || 0) + 1; });
  const unitMain = docUnit || Object.keys(unitVotes).sort((a, b) => unitVotes[b] - unitVotes[a])[0] || '';
  const items = raw.filter((it) => !(it.unit === 'KG' && unitMain !== 'KG')).map((it) => {
    const ctx = [].concat(it.before, [it.own], it.inherit ? [it.inherit] : [], it.afterOK ? it.after : []).join(' \n ');
    const matsOwn = genMats(it.own);
    const mats = matsOwn.length ? matsOwn : genMats(ctx);
    const own = it.own.replace(/\b(?:[A-Z][A-Z&]{2}\d{7}|[A-Z]{2}\d{8})\b/g, ' ');
    const unit = it.unit || unitMain;
    return {
      poRaw: it.poRaw, po: it.po, poList: it.poList, poSapLike: it.poSapLike,
      article: '', desc: ctx, colorText: it.colorCell || ctx, colorCode: '', colorShort: (it.colorCell || own).replace(/\s+/g, ' ').trim().slice(0, 60),
      size: it.sizeCell || '', lot: '', mats, matsFromCtx: !matsOwn.length, material: mats[0] || '', qty: it.qty, unit, qtyByUnit: unit ? { [unit]: it.qty } : {},
      price: it.price, amount: it.amount, surcharge: 0, code: own.replace(/\s+/g, ' ').trim().slice(0, 40), ctx, own: it.own, page: it.page, gen: true,
    };
  });
  /* chỉ có packing list (hoá đơn là bản scan): gộp các dòng cuộn cùng PO + mã + màu + size thành một dòng */
  if (!hasInv && hasPkl && items.length > 1) {
    const agg = [];
    for (const it of items) {
      const k = [it.po, it.mats.join('/'), AZ(it.colorCell || ''), it.size].join('|');
      let g = agg.find((x) => x._k === k && (it.colorCell || !x.colorCell || AZ(x.colorText) === AZ(it.colorText)));
      if (!g) { g = Object.assign({}, it, { _k: k, rolls: 0 }); agg.push(g); }
      else { g.qty = Math.round((g.qty + it.qty) * 1000) / 1000; g.qtyByUnit = it.unit ? { [it.unit]: g.qty } : {}; g.amount = NaN; g.price = NaN; }
      g.rolls++;
    }
    if (agg.length < items.length) { items.length = 0; agg.forEach((g) => { g.colorShort = (g.colorShort || '') + (g.rolls > 1 ? ` (${g.rolls} cuộn)` : ''); g.code = g.colorShort.slice(0, 40); items.push(g); }); }
  }
  let no = pdfInvNo(allLines), noFrom = '';
  /* chỉ có packing list: số trên packing list không phải số hoá đơn → ưu tiên số ghi ở tên thư mục/tên file */
  if (!no || (!hasInv && hasPkl)) { const f = invNoFromName(fname, dir, allLines.map((l) => l.text).join('\n')); if (f && (f.from === 'tên thư mục' || !no)) { no = f.no; noFrom = f.from + (!hasInv ? ' (chứng từ chỉ có packing list)' : ''); } }
  const invDate = pdfInvDate(allLines, yearHint);
  /* tổng: dòng TOTAL có số + đơn vị */
  let totalQty = NaN;
  for (const L of allLines) {
    if (!/^(GRAND\s*)?TOTAL\b|^TOT(ALE)?\b|^TTL\b/i.test(L.text.trim())) continue;
    if (/BATCH|ROLLS\s*\/|COLOUR|BALE/i.test(L.text) && !/TOTAL\s+ITEM/i.test(L.text)) continue;
    const q = qtyOnLine(L, null, vn);
    if (q && q.q > 0 && (!q.unit || q.unit === unitMain)) {
      /* dòng tổng có cả số cuộn lẫn số mét ("N. rolls 34 MT 2088,74") → lấy số lớn nhất kèm đơn vị */
      const w = wordsOf(L); let best = q.q;
      for (let i = 1; i < w.length; i++) if (genUnit(w[i].s) === (q.unit || unitMain) && /^\d[\d.,]*$/.test(w[i - 1].s)) best = Math.max(best, genNum(w[i - 1].s, vn) || 0);
      for (let i = 0; i + 1 < w.length; i++) if (genUnit(w[i].s) === (q.unit || unitMain) && /^\d[\d.,]*$/.test(w[i + 1].s)) best = Math.max(best, genNum(w[i + 1].s, vn) || 0);
      totalQty = best; break;
    }
  }
  let groups = null;
  if (hasPkl) {
    const pk = readPages((r) => r === 'pkl').items.filter((it) => !(it.unit === 'KG' && unitMain !== 'KG'));
    groups = [];
    for (const it of pk) {
      const ctx = [].concat(it.before, [it.own], it.afterOK ? it.after : []).join(' ');
      const key = (it.po || '') + '|' + AZ(it.own.replace(/[\d.,]+/g, ' ')).slice(0, 40);
      const artTok = (it.own.replace(/\b(?:[A-Z][A-Z&]{2}\d{7}|[A-Z]{2}\d{8})\b/g, ' ').match(/\b[A-Z0-9][A-Z0-9\/.#-]*\d[A-Z0-9\/.#-]*\b/gi) || []).find((x) => /[A-Z]/i.test(x) && x.length >= 5 && !/^(HS|NO)/i.test(x)) || '';
      let g = groups.find((x) => x.key === key);
      if (!g) { g = { key, po: it.po, poRaw: it.poRaw, article: artTok, color: it.colorCell || ctx, lot: '', unit: it.unit || unitMain, rolls: [], total: 0, pdf: true }; groups.push(g); }
      g.rolls.push({ no: String(g.rolls.length + 1), qty: it.qty });
      g.total = Math.round((g.total + it.qty) * 1000) / 1000;
    }
  }
  const out = {
    profile: 'GEN', supplier: '', role, file: fname || '',
    inv: { no, invNo: no, invDate, items, currency: 'USD', unitDefault: unitMain, surchargeHeader: 0, total: NaN, totalQty, amountInclSur: false, noInvoiceNo: !no, noFrom, gen: true, hasSapPo: items.some((x) => x.poSapLike) },
    pkl: groups && groups.length ? { groups, unit: unitMain, level: 'lot', soft: true } : null,
    pklOnly: !hasInv && hasPkl,
  };
  return out;
}

if (typeof module !== 'undefined' && module.exports) {
  Object.assign(module.exports, { readGenPdf, readVatPdf, genLinesOf, genDate, pdfInvNo, pdfInvDate, pageRole, pdfItemXY });
}
