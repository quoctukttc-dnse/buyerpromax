/* =====================================================================================================
   ĐIỀN ORDER FORM CỦA SUPPLIER TỪ PO SCAVI (trang order.html — tách riêng, không chạm luồng đối chiếu inbound)
   - Đọc PO SCAVI (Excel "PURCHASE ORDER": No. | Material Code | Description | Supp. Ref. | Size | Color | Spec. | Qty. | Unit | Price | ETD | ETA)
   - Nhận diện form của supplier theo "hồ sơ mẫu" (PROFILES), ghép từng dòng form với dòng PO và điền Số lượng.
   - Ghi THẲNG vào file xlsx gốc: chỉ sửa các ô số lượng (+ vài ô phần đầu nếu trống) trong XML của sheet, mọi phần khác
     (định dạng, dropdown, công thức, control, hình) giữ nguyên từng byte. Không dùng thư viện ngoài: xlsx = zip → giải/nén
     bằng DecompressionStream / CompressionStream của trình duyệt (Chrome/Edge 103+, Safari 16.4+).
   Phần lõi (OrderCore) không đụng DOM để có thể chạy kiểm thử bằng Node.
   ===================================================================================================== */
(function (root) {
  'use strict';

  /* ---------------------------------------- tiện ích ---------------------------------------- */
  const AZ = (s) => String(s == null ? '' : s).toUpperCase().replace(/[^A-Z0-9]/g, '');
  const colIdx = (L) => { let n = 0; for (const ch of String(L).toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
  const colLetter = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const splitRef = (ref) => { const m = String(ref).match(/^([A-Z]+)(\d+)$/); return m ? { col: colIdx(m[1]), row: +m[2], L: m[1] } : null; };
  const xmlUnesc = (s) => String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (m, d) => String.fromCharCode(+d)).replace(/&#x([0-9a-f]+);/gi, (m, h) => String.fromCharCode(parseInt(h, 16))).replace(/&amp;/g, '&');
  const xmlEsc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const td = new TextDecoder('utf-8'), te = new TextEncoder();
  const fmtN = (n) => (n == null || isNaN(n)) ? '' : Number(n).toLocaleString('vi-VN');
  /* ngày Excel (số serial) → dd/mm/yyyy */
  const serialToDMY = (v) => { const n = Number(v); if (!isFinite(n) || n < 20000 || n > 80000) return ''; const d = new Date(Math.round((n - 25569) * 86400 * 1000)); return `${String(d.getUTCDate()).padStart(2, '0')}/${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()}`; };
  const anyDate = (v) => { if (v == null || v === '') return ''; if (typeof v === 'number') return serialToDMY(v); const s = String(v).trim(); const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return `${m[3]}/${m[2]}/${m[1]}`; return s; };

  /* ---------------------------------------- CRC32 ---------------------------------------- */
  const CRC_T = (() => { const t = new Int32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c; } return t; })();
  const crc32 = (u8) => { let c = -1; for (let i = 0; i < u8.length; i++) c = CRC_T[(c ^ u8[i]) & 0xFF] ^ (c >>> 8); return (c ^ -1) >>> 0; };

  /* ---------------------------------------- ZIP đọc / ghi ---------------------------------------- */
  const hasStreams = typeof DecompressionStream !== 'undefined' && typeof CompressionStream !== 'undefined';
  async function inflateRaw(u8) {
    const ds = new DecompressionStream('deflate-raw');
    const buf = await new Response(new Blob([u8]).stream().pipeThrough(ds)).arrayBuffer();
    return new Uint8Array(buf);
  }
  async function deflateRaw(u8) {
    const cs = new CompressionStream('deflate-raw');
    const buf = await new Response(new Blob([u8]).stream().pipeThrough(cs)).arrayBuffer();
    return new Uint8Array(buf);
  }
  /* đọc zip: trả { entries: [{name, method, crc, csize, usize, time, date, extAttr, comp(Uint8Array) }] } */
  function zipRead(buf) {
    const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 70000); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw new Error('Không phải file zip/xlsx hợp lệ');
    const n = dv.getUint16(eocd + 10, true), cdOff = dv.getUint32(eocd + 16, true);
    const entries = [];
    let p = cdOff;
    for (let i = 0; i < n; i++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('Zip: hỏng thư mục trung tâm');
      const method = dv.getUint16(p + 10, true), time = dv.getUint16(p + 12, true), date = dv.getUint16(p + 14, true);
      const crc = dv.getUint32(p + 16, true), csize = dv.getUint32(p + 20, true), usize = dv.getUint32(p + 24, true);
      const nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
      const extAttr = dv.getUint32(p + 38, true), lho = dv.getUint32(p + 42, true);
      const name = td.decode(u8.subarray(p + 46, p + 46 + nl));
      const lnl = dv.getUint16(lho + 26, true), lel = dv.getUint16(lho + 28, true);
      const dataOff = lho + 30 + lnl + lel;
      entries.push({ name, method, crc, csize, usize, time, date, extAttr, comp: u8.subarray(dataOff, dataOff + csize) });
      p += 46 + nl + el + cl;
    }
    return { entries };
  }
  async function zipEntryText(e) {
    if (e.method === 0) return td.decode(e.comp);
    if (e.method !== 8) throw new Error('Zip: phương pháp nén không hỗ trợ (' + e.method + ')');
    return td.decode(await inflateRaw(e.comp));
  }
  /* ghi zip: entries gốc giữ nguyên dữ liệu nén; entry trong `changed` (name → string) được nén lại */
  async function zipWrite(zip, changed) {
    const parts = [], central = [];
    let off = 0;
    const u16 = (v) => [v & 255, (v >>> 8) & 255], u32 = (v) => [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255];
    for (const e of zip.entries) {
      let method = e.method, crc = e.crc, csize = e.csize, usize = e.usize, data = e.comp;
      if (changed && Object.prototype.hasOwnProperty.call(changed, e.name)) {
        const raw = te.encode(changed[e.name]);
        usize = raw.length; crc = crc32(raw);
        if (hasStreams) { data = await deflateRaw(raw); method = 8; } else { data = raw; method = 0; }
        csize = data.length;
      }
      const nameB = te.encode(e.name);
      const lh = new Uint8Array([...u32(0x04034b50), ...u16(20), ...u16(0), ...u16(method), ...u16(e.time), ...u16(e.date), ...u32(crc), ...u32(csize), ...u32(usize), ...u16(nameB.length), ...u16(0)]);
      parts.push(lh, nameB, data);
      const ch = new Uint8Array([...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0), ...u16(method), ...u16(e.time), ...u16(e.date), ...u32(crc), ...u32(csize), ...u32(usize), ...u16(nameB.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(e.extAttr), ...u32(off)]);
      central.push(ch, nameB);
      off += lh.length + nameB.length + data.length;
    }
    const cdLen = central.reduce((a, b) => a + b.length, 0);
    const eocd = new Uint8Array([...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(zip.entries.length), ...u16(zip.entries.length), ...u32(cdLen), ...u32(off), ...u16(0)]);
    const total = off + cdLen + eocd.length;
    const out = new Uint8Array(total);
    let q = 0;
    for (const p2 of parts.concat(central, [eocd])) { out.set(p2, q); q += p2.length; }
    return out;
  }

  /* ---------------------------------------- XLSX: đọc nhẹ ---------------------------------------- */
  function parseSharedStrings(xml) {
    const out = [];
    if (!xml) return out;
    const re = /<si>([\s\S]*?)<\/si>|<si\/>/g;
    let m;
    while ((m = re.exec(xml))) {
      if (m[1] == null) { out.push(''); continue; }
      const ts = []; const rt = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<t(?:\s[^>]*)?\/>/g; let k;
      while ((k = rt.exec(m[1]))) ts.push(k[1] == null ? '' : xmlUnesc(k[1]));
      out.push(ts.join(''));
    }
    return out;
  }
  /* ô của một sheet: Map ref → { v (giá trị đã giải), raw, t, s, f } */
  function parseSheetCells(xml, sst) {
    const cells = new Map();
    const re = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;   // thuộc tính có thể đứng trước r (PO SCAVI: <c t="s" r="A18" s="22">)
    let m;
    while ((m = re.exec(xml))) {
      const attrs = m[1] || '', inner = m[2] || '';
      const ref = (attrs.match(/\br="([A-Z]+\d+)"/) || [])[1];
      if (!ref) continue;
      const t = (attrs.match(/\bt="([^"]*)"/) || [])[1] || '';
      const s = (attrs.match(/\bs="([^"]*)"/) || [])[1] || '';
      const f = /<f[\s>]/.test(inner) || /<f\/>/.test(inner);
      let v = null;
      const vm = inner.match(/<v>([\s\S]*?)<\/v>/);
      if (t === 's') { v = vm ? (sst[+vm[1]] != null ? sst[+vm[1]] : '') : ''; }
      else if (t === 'inlineStr') { const ts = []; const rt = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g; let k; while ((k = rt.exec(inner))) ts.push(xmlUnesc(k[1])); v = ts.join(''); }
      else if (t === 'str') { v = vm ? xmlUnesc(vm[1]) : ''; }
      else if (t === 'b') { v = vm ? vm[1] === '1' : null; }
      else if (t === 'e') { v = vm ? xmlUnesc(vm[1]) : null; }
      else { v = vm ? Number(vm[1]) : null; if (vm && isNaN(v)) v = xmlUnesc(vm[1]); }
      cells.set(ref, { v, t, s, f });
    }
    return cells;
  }
  /* mở workbook: { zip, sheets: [{name, path, xml, cells}], sst, workbookXml } */
  async function openBook(buf) {
    const zip = zipRead(buf);
    const byName = Object.fromEntries(zip.entries.map((e) => [e.name, e]));
    const need = (n) => { const e = byName[n]; if (!e) throw new Error('xlsx thiếu ' + n); return zipEntryText(e); };
    const wbXml = await need('xl/workbook.xml');
    const rels = await need('xl/_rels/workbook.xml.rels');
    const sstE = byName['xl/sharedStrings.xml'];
    const sst = parseSharedStrings(sstE ? await zipEntryText(sstE) : '');
    const relMap = {};
    for (const m of rels.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
      const a = m[1]; const id = (a.match(/\bId="([^"]*)"/) || [])[1]; const tg = (a.match(/\bTarget="([^"]*)"/) || [])[1];
      if (id && tg) relMap[id] = tg.replace(/^\//, '').replace(/^(?!xl\/)/, 'xl/');
    }
    const sheets = [];
    for (const m of wbXml.matchAll(/<sheet\b([^>]*)\/?>/g)) {
      const a = m[1]; const name = xmlUnesc((a.match(/\bname="([^"]*)"/) || [])[1] || ''); const rid = (a.match(/\br:id="([^"]*)"/) || a.match(/\b\w+:id="([^"]*)"/) || [])[1];
      const path = relMap[rid]; if (!path || !byName[path]) continue;
      const xml = await zipEntryText(byName[path]);
      sheets.push({ name, path, xml, cells: parseSheetCells(xml, sst) });
    }
    return { zip, sheets, sst, workbookXml: wbXml };
  }
  const cellV = (sh, ref) => { const c = sh.cells.get(ref); return c ? c.v : null; };
  const cellS = (sh, ref) => { const v = cellV(sh, ref); return v == null ? '' : String(v).trim(); };

  /* ---------------------------------------- XLSX: ghi ô (sửa chuỗi XML) ---------------------------------------- */
  function findElem(xml, from, openRe) {
    openRe.lastIndex = from;
    const m = openRe.exec(xml);
    if (!m) return null;
    const start = m.index;
    const gt = xml.indexOf('>', start);
    const selfClose = xml[gt - 1] === '/';
    let end;
    if (selfClose) end = gt + 1;
    else { const tag = m[0].slice(1).match(/^[a-zA-Z]+/)[0]; end = xml.indexOf('</' + tag + '>', gt) + tag.length + 3; }
    return { start, end, openTag: xml.slice(start, gt + 1), selfClose, inner: selfClose ? '' : xml.slice(gt + 1, end - (m[0].slice(1).match(/^[a-zA-Z]+/)[0].length + 3)) };
  }
  /* đặt giá trị một ô: val = { num } | { text } ; giữ style s của ô cũ (hoặc style do caller gợi ý) */
  function setCell(xml, ref, val, styleHint) {
    const R = splitRef(ref); if (!R) throw new Error('ref lỗi ' + ref);
    const sdOpen = xml.indexOf('<sheetData');
    if (sdOpen < 0) throw new Error('sheet không có sheetData');
    if (/<sheetData\s*\/>/.test(xml.slice(sdOpen, sdOpen + 20))) xml = xml.replace(/<sheetData\s*\/>/, '<sheetData></sheetData>');
    const sdEnd = xml.indexOf('</sheetData>');
    const rowRe = new RegExp('<row\\b[^>]*?\\br="' + R.row + '"(?=[\\s>/])', 'g');
    let row = findElem(xml, sdOpen, rowRe);
    if (row && row.start > sdEnd) row = null;
    const mk = (s) => {
      const sa = s ? ` s="${s}"` : '';
      if ('num' in val) return `<c r="${ref}"${sa}><v>${val.num}</v></c>`;
      return `<c r="${ref}"${sa} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(val.text)}</t></is></c>`;
    };
    if (!row) {
      /* chèn dòng mới đúng thứ tự */
      let insAt = sdEnd;
      const rr = /<row\b[^>]*?\br="(\d+)"/g; rr.lastIndex = sdOpen; let m;
      while ((m = rr.exec(xml)) && m.index < sdEnd) { if (+m[1] > R.row) { insAt = m.index; break; } }
      const frag = `<row r="${R.row}">${mk(styleHint || '')}</row>`;
      return xml.slice(0, insAt) + frag + xml.slice(insAt);
    }
    const cellRe = new RegExp('<c\\b[^>]*?\\br="' + ref + '"(?=[\\s>/])', 'g');
    const cell = findElem(xml, row.start, cellRe);
    if (cell && cell.start < row.end) {
      const s = (cell.openTag.match(/\bs="([^"]*)"/) || [])[1] || styleHint || '';
      if (/<f[\s>\/]/.test(cell.inner)) return xml;   // ô công thức → không ghi đè
      return xml.slice(0, cell.start) + mk(s) + xml.slice(cell.end);
    }
    /* chèn ô mới đúng thứ tự cột trong dòng */
    let rowXml = xml.slice(row.start, row.end);
    if (row.selfClose) rowXml = rowXml.replace(/\/>$/, '>') + '</row>';
    const openEnd = rowXml.indexOf('>') + 1;
    let insAt = rowXml.lastIndexOf('</row>');
    const cr = /<c\b[^>]*?\br="([A-Z]+)(\d+)"/g; cr.lastIndex = openEnd; let m;
    while ((m = cr.exec(rowXml))) { if (colIdx(m[1]) > R.col) { insAt = m.index; break; } }
    rowXml = rowXml.slice(0, insAt) + mk(styleHint || '') + rowXml.slice(insAt);
    return xml.slice(0, row.start) + rowXml + xml.slice(row.end);
  }
  /* Excel tính lại công thức khi mở (ô Total = SUM đang giữ giá trị cũ) */
  function forceRecalc(wbXml) {
    if (/<calcPr\b[^>]*fullCalcOnLoad=/.test(wbXml)) return wbXml;
    if (/<calcPr\b/.test(wbXml)) return wbXml.replace(/<calcPr\b/, '<calcPr fullCalcOnLoad="1"');
    return wbXml.replace(/<\/workbook>/, '<calcPr fullCalcOnLoad="1"/></workbook>');
  }

  /* ---------------------------------------- PO SCAVI ---------------------------------------- */
  /* trả { po, orderDate, supplier, buyer, buyerAddr, buyerTax, items: [{no, material, desc, supRef, size, color, spec, qty, unit, price, etd, eta}], sheet } */
  function readScaviPo(book) {
    for (const sh of book.sheets) {
      /* tìm dòng tiêu đề bảng hàng */
      const byRow = new Map();
      for (const [ref, c] of sh.cells) { const R = splitRef(ref); if (!byRow.has(R.row)) byRow.set(R.row, []); byRow.get(R.row).push({ ref, col: R.col, L: R.L, v: c.v }); }
      let hdr = null;
      for (const [r, cs] of byRow) {
        const txt = cs.map((c) => AZ(c.v));
        if (txt.includes('MATERIALCODE') && txt.some((t) => /^(QTY|QUANTITY)/.test(t))) { hdr = { r, cs }; break; }
      }
      if (!hdr) continue;
      const col = (re) => { const c = hdr.cs.find((x) => re.test(String(x.v == null ? '' : x.v).replace(/\s+/g, ' ').trim())); return c ? c.L : ''; };
      const C = { no: col(/^No\.?$/i), material: col(/^Material\s*Code$/i), desc: col(/^Description$/i), supRef: col(/^Supp\.?\s*Ref/i), size: col(/^Size$/i), color: col(/^Colou?r$/i), spec: col(/^Spec/i), qty: col(/^(Qty|Quantity)/i), unit: col(/^Unit$/i), price: col(/^Price$/i), etd: col(/^ETD$/i), eta: col(/^ETA$/i),
        orderNo: col(/^Order\s*No/i), reference: col(/^Reference$/i), story: col(/^Story$/i) };   // PO phụ liệu Corele: thêm Order No | Reference | Story, cột "Quantity"
      const items = [];
      const rows = [...byRow.keys()].filter((r) => r > hdr.r).sort((a, b) => a - b);
      for (const r of rows) {
        const g = (L) => (L ? cellV(sh, L + r) : null);
        const no = g(C.no), mat = g(C.material), qty = g(C.qty);
        if (mat == null && no == null) { if (items.length && /^[A-Za-z]/.test(String(cellS(sh, 'A' + r)))) break; continue; }
        if (!mat || !(Number(qty) >= 0) || !/^\d+$/.test(String(no == null ? '' : no).trim())) continue;
        const S = (L) => String(g(L) == null ? '' : g(L)).trim();
        items.push({ no: +no, material: String(mat).trim(), desc: S(C.desc), supRef: S(C.supRef), size: S(C.size), color: S(C.color), spec: S(C.spec), qty: Number(qty), unit: S(C.unit), price: Number(g(C.price)), etd: anyDate(g(C.etd)), eta: anyDate(g(C.eta)), orderNo: S(C.orderNo), reference: S(C.reference), story: S(C.story), row: r });
      }
      if (!items.length) continue;
      /* phần đầu: nhãn → ô có giá trị kế bên phải cùng dòng */
      const labelVal = (re, nth = 1) => {
        let seen = 0;
        for (const r of [...byRow.keys()].filter((x) => x < hdr.r).sort((a, b) => a - b)) {
          const cs = byRow.get(r).sort((a, b) => a.col - b.col);
          for (let i = 0; i < cs.length; i++) {
            if (re.test(String(cs[i].v == null ? '' : cs[i].v).trim())) {
              seen++;
              if (seen < nth) continue;
              const nx = cs.slice(i + 1).find((c) => c.v != null && String(c.v).trim() !== '');
              return nx ? nx.v : '';
            }
          }
        }
        return '';
      };
      /* "PO No:" + ô kế bên (Avery) hoặc gộp một ô "PO No: AVV0155600" (PO phụ liệu Corele) */
      let po = String(labelVal(/^PO\s*No\.?\s*:?$/i) || '').trim();
      if (!po) for (const [, c] of sh.cells) { const m = String(c.v == null ? '' : c.v).match(/PO\s*No\.?\s*:\s*([A-Z]{2,3}\d{7,8})/i); if (m) { po = m[1].toUpperCase(); break; } }
      return {
        po, orderDate: anyDate(labelVal(/^Order\s*Date\s*:?$/i)), supplier: String(labelVal(/^Supplier\s*:?$/i) || '').trim(),
        buyer: String(labelVal(/^Buyer\s*\/?\s*Invoice\s*To\s*:?$/i) || '').trim(), buyerAddr: String(labelVal(/^Address\s*:?$/i, 2) || labelVal(/^Address\s*:?$/i) || '').trim(),
        buyerTax: String(labelVal(/^Tax\s*Code\s*:?$/i, 2) || '').trim(), planNo: String(labelVal(/^Plan\s*No\s*:?$/i) || '').trim(),
        items, sheet: sh.name, etdMin: items.map((x) => x.etd).filter(Boolean).sort((a, b) => a.split('/').reverse().join('').localeCompare(b.split('/').reverse().join('')))[0] || '',
      };
    }
    return null;
  }
  /* Spec của PO Avery: "522866-BKY-10P" (ID) · "522866-P25T-LE-SW-US-BKY-10P" (care label) · "522867-BLACK/EGRET TOSSED FLORAL -10L-PO 234911" (sticker) */
  function parseSpec(spec, itemCode) {
    const s = String(spec || '').trim();
    const style = (s.match(/^(\d{6})\b/) || [])[1] || '';
    let rest = s.replace(/^\d{6}-?/, '');
    if (itemCode) { const i = rest.toUpperCase().indexOf(itemCode.toUpperCase()); if (i >= 0) rest = rest.slice(i + itemCode.length).replace(/^-/, ''); }
    /* tách theo "-" nhưng KHÔNG cắt dấu gạch nằm trong ngoặc: "US-AW3-L (14-16)" → [US, AW3, L (14-16)] (PO AVV0156900) */
    const parts = []; let cur = '', depth = 0;
    for (const ch of rest) { if (ch === '(') depth++; if (ch === ')') depth = Math.max(0, depth - 1); if (ch === '-' && depth === 0) { parts.push(cur); cur = ''; } else cur += ch; }
    parts.push(cur);
    const ps = parts.map((x) => x.trim()).filter(Boolean);
    return { style, parts: ps, size: ps[ps.length - 1] || '', color: ps.length >= 2 ? ps[ps.length - 2] : '', coo: ps.length >= 3 ? ps[0] : '' };
  }

  /* ---------------------------------------- HỒ SƠ MẪU FORM ---------------------------------------- */
  /* Mỗi hồ sơ: id, name, detect(book) → ctx | null ; plan(book, ctx, po, opts) → { writes:[{sheet, ref, num|text, what}], report, styles } */
  const PROFILES = [];

  /* Avery Dennison — Lands' End STYLE ID Label (ID LABEL): sheet REGULAR, dòng từ tiêu đề "UPC_BARCODE_IN*" + 1:
     B = Lands' End Style, F = màu (Optional info 1), G = size (Optional info 2), cột "Quantity" của dòng tiêu đề; PO: Supp. Ref "ID", Spec style-màu-size */
  PROFILES.push({
    id: 'AVERY_ID_LABEL', name: 'Avery Dennison — Lands\' End STYLE ID Label',
    detect(book) {
      const cands = [];
      for (const sh of book.sheets) {
        let hdrRow = 0, qtyCol = '';
        for (const [ref, c] of sh.cells) { if (/^UPC_BARCODE_IN/i.test(String(c.v || ''))) { hdrRow = splitRef(ref).row; break; } }
        if (!hdrRow) continue;
        for (const [ref, c] of sh.cells) { const R = splitRef(ref); if (R.row === hdrRow && /^Quantity/i.test(String(c.v || '').trim())) qtyCol = R.L; }
        if (!qtyCol) continue;
        let n = 0; for (let r = hdrRow + 1; r <= hdrRow + 2000; r++) { if (cellS(sh, 'B' + r)) n++; else if (n && !cellS(sh, 'A' + r) && !cellS(sh, qtyCol + r)) break; }
        cands.push({ sheet: sh.name, hdrRow, qtyCol, rows: n });
      }
      if (!cands.length) return null;
      const pick = cands.find((c) => /^REGULAR$/i.test(c.sheet)) || cands.sort((a, b) => b.rows - a.rows)[0];
      return Object.assign({ all: cands.map((c) => c.sheet) }, pick);
    },
    plan(book, ctx, po, opts) {
      const sh = book.sheets.find((s) => s.name === ctx.sheet);
      const poRows = po.items.filter((it) => AZ(it.supRef) === 'ID' || /STYLE\s*ID\s*LABEL/i.test(it.desc));
      const byKey = new Map();
      for (const it of poRows) { const p = parseSpec(it.spec); const k = [p.style, AZ(p.color), AZ(p.size)].join('|'); byKey.set(k, (byKey.get(k) || 0) + it.qty); }
      const used = new Set(), writes = [], zero = [], formRows = [];
      for (let r = ctx.hdrRow + 1; r <= ctx.hdrRow + 5000; r++) {
        const style = cellS(sh, 'B' + r), color = cellS(sh, 'F' + r), size = cellS(sh, 'G' + r);
        if (!style) { if (!cellS(sh, 'A' + r) && !cellS(sh, ctx.qtyCol + r)) { if (formRows.length) break; continue; } continue; }
        if (!/^\d{5,}$/.test(style)) continue;
        const k = [style, AZ(color), AZ(size)].join('|');
        formRows.push(r);
        if (byKey.has(k)) { writes.push({ sheet: ctx.sheet, ref: ctx.qtyCol + r, num: byKey.get(k), what: `${style}-${color}-${size}` }); used.add(k); }
        else if (opts.zeroMissing) { writes.push({ sheet: ctx.sheet, ref: ctx.qtyCol + r, num: 0, what: `${style}-${color}-${size} (không có trên PO)` }); zero.push(`${style}-${color}-${size}`); }
      }
      const notInForm = [...byKey.keys()].filter((k) => !used.has(k)).map((k) => k.replace(/\|/g, '-'));
      const sumForm = writes.reduce((a, w) => a + (w.num || 0), 0), sumPo = poRows.reduce((a, it) => a + it.qty, 0);
      /* phần đầu form: chỉ điền ô đang trống */
      const head = [];
      if (opts.fillHeader) {
        const put = (labelRef, valRef, text, what) => { if (text && cellS(sh, valRef) === '') { writes.push({ sheet: ctx.sheet, ref: valRef, text, what }); head.push(`${what}: ${text}`); } else if (text && cellS(sh, valRef) !== '') head.push(`${what}: giữ "${cellS(sh, valRef)}" (form đã có)`); };
        const findLabel = (re) => { for (const [ref, c] of sh.cells) { const R = splitRef(ref); if (R.row < ctx.hdrRow && re.test(String(c.v || '').trim())) return R; } return null; };
        const right = (R, n = 1) => colLetter(R.col + n) + R.row;
        const L1 = findLabel(/^PO#?\s*:?$/i); if (L1) put(L1, right(L1), po.po, 'PO#');
        const L2 = findLabel(/^Order\s*Date\s*:?$/i); if (L2) put(L2, right(L2), po.orderDate, 'Order Date');
        const L3 = findLabel(/^Request\s*Ship\s*Date\s*:?$/i); if (L3) put(L3, right(L3), po.etdMin, 'Request Ship Date');
        const L4 = findLabel(/^BILL\s*TO\s*:?$/i); if (L4) put(L4, right(L4), po.buyer, 'Bill to');
        const L5 = findLabel(/^ADDRESS\s*:?$/i); if (L5) put(L5, right(L5), po.buyerAddr, 'Bill to address');
      }
      return { writes, report: { formRows: formRows.length, filled: writes.filter((w) => 'num' in w && w.num > 0).length, zero: zero.length, zeroList: zero, notInForm, sumForm, sumPo, poRows: poRows.length, head, sheet: ctx.sheet, note: ctx.all.length > 1 ? `File có ${ctx.all.length} sheet mẫu (${ctx.all.join(', ')}) — chỉ điền sheet ${ctx.sheet}.` : '' } };
    },
  });

  /* Avery Dennison — Lands' End care / size label order form (P25T-LE-SW): sheet "Print Size-COO-Content" (hoặc sheet có Customer Item# + US Size + Quantity):
     G6 = Item code; dòng size: I = US Size, L = Quantity (từ dòng tiêu đề "US Size"/"Quantity" + 1 tới khi hết size).
     PO: Supp. Ref chứa item code (C.LAB P25T-LE-SW(33*45mm)), Spec style-item-COO-màu-size; một form = một style (lấy từ tên file / ô LE style# / người dùng chọn),
     cộng dồn các màu (lọc theo màu + COO ghi trong tên file nếu có). */
  PROFILES.push({
    id: 'AVERY_SIZE_FORM', name: 'Avery Dennison — Lands\' End size / care label order form',
    detect(book) {
      for (const sh of book.sheets) {
        let itemRef = '', hdr = null;
        for (const [ref, c] of sh.cells) { if (/^Customer\s*Item#?$/i.test(String(c.v || '').trim())) itemRef = ref; }
        if (!itemRef) continue;
        for (const [ref, c] of sh.cells) { if (/^US\s*Size$/i.test(String(c.v || '').trim())) { const R = splitRef(ref); hdr = hdr && hdr.row < R.row ? hdr : R; } }
        if (!hdr) continue;
        let qtyCol = '';
        for (const [ref, c] of sh.cells) { const R = splitRef(ref); if (R.row === hdr.row && /^Quantity$/i.test(String(c.v || '').trim())) qtyCol = R.L; }
        if (!qtyCol) continue;
        const IR = splitRef(itemRef);
        const itemCode = cellS(sh, IR.L + (IR.row + 1)) || cellS(sh, colLetter(IR.col) + (IR.row + 1));
        const sizes = [];
        for (let r = hdr.row + 1; r <= hdr.row + 60; r++) { const sz = cellS(sh, hdr.L + r); if (!sz) break; sizes.push({ r, size: sz, gender: cellS(sh, 'G' + r) }); }
        if (!sizes.length) continue;
        const styleCell = (() => { for (const [ref, c] of sh.cells) { if (/^LE\s*style#?$/i.test(String(c.v || '').trim())) { const R = splitRef(ref); return cellS(sh, colLetter(R.col + 1) + R.row) || cellS(sh, R.L + (R.row + 1)); } } return ''; })();
        return { sheet: sh.name, itemCode, sizeCol: hdr.L, qtyCol, sizes, styleCell };
      }
      return null;
    },
    options(ctx, po, fname) {
      /* style ứng viên trong PO cho item code này; chọn sẵn theo tên file / ô LE style# */
      const code = AZ(ctx.itemCode);
      const rows = po.items.filter((it) => code && (AZ(it.supRef).includes(code) || AZ(it.spec).includes(code)));
      const styles = [...new Set(rows.map((it) => parseSpec(it.spec, ctx.itemCode).style).filter(Boolean))].sort();
      const fn = String(fname || '').toUpperCase();
      const fromName = styles.filter((s) => fn.includes(s));
      const cellStyle = (String(ctx.styleCell || '').match(/\b\d{6}\b/g) || []).find((x) => styles.includes(x)) || '';   // ô LE style# có thể ghi "154489 - 555328 COLOR AW3 JYT"
      const pre = cellStyle && styles.includes(cellStyle) ? [cellStyle] : (fromName.length ? fromName : (styles.length === 1 ? styles : []));
      const colors = [...new Set(rows.map((it) => parseSpec(it.spec, ctx.itemCode).color).filter(Boolean))].sort();
      const coos = [...new Set(rows.map((it) => parseSpec(it.spec, ctx.itemCode).coo).filter(Boolean))].sort();
      const fnToks = fn.split(/[^A-Z0-9]+/).filter(Boolean);
      return { styles, pre, colors, colorsPre: colors.filter((c) => fnToks.includes(c.toUpperCase())), coos, cooPre: coos.filter((c) => fnToks.includes(c.toUpperCase())), poRows: rows.length };
    },
    plan(book, ctx, po, opts) {
      const sh = book.sheets.find((s) => s.name === ctx.sheet);
      const code = AZ(ctx.itemCode);
      let rows = po.items.filter((it) => code && (AZ(it.supRef).includes(code) || AZ(it.spec).includes(code)));
      const sel = opts.styles || [];
      if (sel.length) rows = rows.filter((it) => sel.includes(parseSpec(it.spec, ctx.itemCode).style));
      if (opts.colors && opts.colors.length) rows = rows.filter((it) => opts.colors.includes(parseSpec(it.spec, ctx.itemCode).color));
      if (opts.coos && opts.coos.length) rows = rows.filter((it) => opts.coos.includes(parseSpec(it.spec, ctx.itemCode).coo));
      /* khóa size: chữ + số sau khi bỏ ký tự khác ("L (14-16)" ↔ "L 14-16" → L1416; "10P" ↔ "10P"); dự phòng: cột Size của PO (L) ↔ chữ đầu của size form */
      const bySize = new Map(), byLetter = new Map();
      for (const it of rows) {
        const k = AZ(parseSpec(it.spec, ctx.itemCode).size || it.size);
        bySize.set(k, (bySize.get(k) || 0) + it.qty);
        const L = AZ(it.size); if (L) byLetter.set(L, (byLetter.get(L) || 0) + it.qty);
      }
      const writes = [], used = new Set(), zero = [];
      for (const s of ctx.sizes) {
        const k = AZ(s.size);
        const letter = AZ(String(s.size).trim().split(/[\s(]/)[0]);
        let v = null, hit = '';
        if (bySize.has(k)) { v = bySize.get(k); hit = k; }
        else if (letter && byLetter.has(letter) && [...bySize.keys()].filter((x) => x.startsWith(letter) && !/^[A-Z]*\d/.test(x.slice(letter.length))).length === 1) { hit = [...bySize.keys()].find((x) => x.startsWith(letter)); v = bySize.get(hit); }
        if (v != null) { writes.push({ sheet: ctx.sheet, ref: ctx.qtyCol + s.r, num: v, what: `size ${s.size}` }); used.add(hit); }
        else if (opts.zeroMissing) { writes.push({ sheet: ctx.sheet, ref: ctx.qtyCol + s.r, num: 0, what: `size ${s.size} (không có trên PO)` }); zero.push(s.size); }
      }
      const notInForm = [...bySize.keys()].filter((k) => !used.has(k));
      const head = [];
      if (opts.fillHeader) {
        const findLabel = (re) => { for (const [ref, c] of sh.cells) { const R = splitRef(ref); if (R.row < 40 && re.test(String(c.v || '').trim())) return R; } return null; };
        const put = (R, text, what) => { if (!R || !text) return; const ref = colLetter(R.col + 1) + R.row; if (cellS(sh, ref) === '') { writes.push({ sheet: ctx.sheet, ref, text, what }); head.push(`${what}: ${text}`); } else head.push(`${what}: giữ "${cellS(sh, ref)}"`); };
        put(findLabel(/^PO#?\s*:?$/i), po.po, 'PO#');
        put(findLabel(/^Order\s*Date\s*:?$/i), po.orderDate, 'Order Date');
        put(findLabel(/^Request\s*Ship\s*Date\s*:?$/i), po.etdMin, 'Request Ship Date');
        put(findLabel(/^Bill\s*to\s*Company\s*Name\s*:?$/i), po.buyer, 'Bill to Company');
        put(findLabel(/^Bill\s*to\s*Address\s*:?$/i), po.buyerAddr, 'Bill to Address');
        const LS = findLabel(/^LE\s*style#?$/i); if (LS && sel.length === 1) put(LS, sel[0], 'LE style#');
      }
      return { writes, report: { formRows: ctx.sizes.length, filled: writes.filter((w) => 'num' in w && w.num > 0).length, zero: zero.length, zeroList: zero, notInForm, sumForm: writes.reduce((a, w) => a + (w.num || 0), 0), sumPo: rows.reduce((a, it) => a + it.qty, 0), poRows: rows.length, head, sheet: ctx.sheet, note: `Item ${ctx.itemCode}${sel.length ? ' · style ' + sel.join(', ') : ' · chưa chọn style'}${opts.colors && opts.colors.length ? ' · màu ' + opts.colors.join('/') : ''}${opts.coos && opts.coos.length ? ' · COO ' + opts.coos.join('/') : ''}` } };
    },
  });

  /* Avery — Corele V RFID sticker (file "Corele RFID sticker 34x54mm-AVV0155600 …"): sheet có dòng 1 "Quantity* | CATEGORY* | SIZE* | COLOR* | STYLE* |
     BARCODE* | PRODUCTION YEAR* | PO* | RETAIL PRICE* | BRAND*", dòng dữ liệu từ dòng có STYLE; cột Quantity là cột A.
     PO: Spec "02169-6_4 Toasted almond/2 Black#4 Da/2 Đen" = style _ màu tiếng Anh # màu tiếng Việt; Size = cột Size; Order No (261022187) = cột PO* của form.
     Khóa: style | size | màu tiếng Việt (bỏ dấu, bỏ khoảng trắng) [+ Order No khi form ghi]. */
  const foldVN = (t) => String(t == null ? '' : t).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[đĐ]/g, 'D');
  const colorKey = (t) => AZ(foldVN(t));
  PROFILES.push({
    id: 'AVERY_CORELE_RFID', name: 'Avery Dennison — Corele V RFID sticker',
    detect(book) {
      for (const sh of book.sheets) {
        let hdrRow = 0; const cols = {};
        for (const [ref, c] of sh.cells) {
          const R = splitRef(ref); if (R.row > 3) continue;
          const t = String(c.v || '').trim().toUpperCase().replace(/\*$/, '').trim();
          if (/^QUANTITY$/.test(t)) { cols.qty = R.L; hdrRow = hdrRow || R.row; }
          if (/^SIZE$/.test(t)) cols.size = R.L; if (/^COLOU?R$/.test(t)) cols.color = R.L; if (/^STYLE$/.test(t)) cols.style = R.L; if (/^PO$/.test(t)) cols.po = R.L; if (/^BARCODE$/.test(t)) cols.barcode = R.L;
        }
        if (!cols.qty || !cols.size || !cols.color || !cols.style) continue;
        let first = 0, n = 0;
        for (let r = hdrRow + 1; r <= hdrRow + 3000; r++) { if (cellS(sh, cols.style + r)) { if (!first) first = r; n++; } else if (n) break; }
        if (!n) continue;
        return { sheet: sh.name, hdrRow, cols, first, rows: n };
      }
      return null;
    },
    plan(book, ctx, po, opts) {
      const sh = book.sheets.find((s) => s.name === ctx.sheet);
      const { cols } = ctx;
      const poRows = po.items.filter((it) => /RFID|STICKER/i.test(it.supRef + ' ' + it.desc));
      const byKey = new Map(), byKeyNoPo = new Map();
      for (const it of poRows) {
        const sp = String(it.spec || ''); const style = sp.split('_')[0].trim(); const vn = sp.includes('#') ? sp.slice(sp.lastIndexOf('#') + 1) : sp.slice(style.length + 1);
        const k = [AZ(style), AZ(it.size), colorKey(vn)].join('|');
        byKeyNoPo.set(k, (byKeyNoPo.get(k) || 0) + it.qty);
        const kp = k + '|' + AZ(it.orderNo); byKey.set(kp, (byKey.get(kp) || 0) + it.qty);
      }
      const used = new Set(), writes = [], zero = [], formRows = [];
      for (let r = ctx.first; r <= ctx.first + 3000; r++) {
        const style = cellS(sh, cols.style + r); if (!style) { if (formRows.length) break; continue; }
        const size = cellS(sh, cols.size + r), color = cellS(sh, cols.color + r), pno = cols.po ? cellS(sh, cols.po + r) : '';
        formRows.push(r);
        const k = [AZ(style), AZ(size), colorKey(color)].join('|');
        const kp = k + '|' + AZ(pno);
        let v = null, hit = '';
        if (pno && byKey.has(kp)) { v = byKey.get(kp); hit = kp; } else if (byKeyNoPo.has(k)) { v = byKeyNoPo.get(k); hit = k; }
        else {
          /* Spec của PO bị cắt ngắn (70 ký tự: "#2 Da/ 2 Đen/ 2 Trắng" ↔ form "2Da/2Đen/2Trắng kem") → so tiền tố màu, chỉ nhận khi duy nhất */
          const pre = [AZ(style), AZ(size)].join('|') + '|', ck = colorKey(color);
          const cands = [...byKeyNoPo.keys()].filter((x) => x.startsWith(pre) && !used.has(x)).filter((x) => { const c = x.slice(pre.length); return c.length >= 4 && (ck.startsWith(c) || c.startsWith(ck)); });
          if (cands.length === 1) { v = byKeyNoPo.get(cands[0]); hit = cands[0]; }
        }
        if (v != null) { writes.push({ sheet: ctx.sheet, ref: cols.qty + r, num: v, what: `${style} · ${size} · ${color}${pno ? ' · PO ' + pno : ''}` }); used.add(hit); used.add(k); }
        else if (opts.zeroMissing) { writes.push({ sheet: ctx.sheet, ref: cols.qty + r, num: 0, what: `${style} · ${size} · ${color} (không có trên PO)` }); zero.push(`${style} ${size} ${color}`); }
      }
      const notInForm = [...byKeyNoPo.keys()].filter((k) => !used.has(k)).map((k) => k.replace(/\|/g, ' · '));
      return { writes, report: { formRows: formRows.length, filled: writes.filter((w) => 'num' in w && w.num > 0).length, zero: zero.length, zeroList: zero, notInForm, sumForm: writes.reduce((a, w) => a + (w.num || 0), 0), sumPo: poRows.reduce((a, it) => a + it.qty, 0), poRows: poRows.length, head: [], sheet: ctx.sheet, note: 'khóa style · size · màu tiếng Việt (sau "#" của Spec) · Order No' } };
    },
  });

  /* Avery — Lacoste care label order form (GBY): sheet "Order form" (hoặc sheet đầu có tiêu đề "Item Code*" + "Style No*" + "quantity"):
     dòng tiêu đề có "Item Code*", "Style No*", cột size = "Accessories(配件尺码）" (XXS…3XL) hoặc SizeK/SizeC, cột "quantity".
     PO: Supp. Ref bắt đầu bằng Item Code (FGUW154x35ADVN), Spec "CARE LABEL FGUW AH795/FGUW" (+ "… ADD", "… ADD1" = trang phụ cùng số lượng → không cộng),
     Size = cột Size của PO. Khóa: item code · style · size. */
  PROFILES.push({
    id: 'AVERY_LACOSTE_CARE', name: 'Avery Dennison — Lacoste care label order form',
    detect(book) {
      const sheets = book.sheets.slice().sort((a, b) => (/^Order\s*form$/i.test(b.name) ? 1 : 0) - (/^Order\s*form$/i.test(a.name) ? 1 : 0));
      for (const sh of sheets) {
        let hdrRow = 0; const cols = {};
        for (const [ref, c] of sh.cells) {
          const R = splitRef(ref); if (R.row > 60) continue;
          const t = String(c.v || '').replace(/\s+/g, ' ').trim();
          if (/^Item\s*Code\*?$/i.test(t)) { cols.item = R.L; hdrRow = R.row; }
          if (/^Style\s*No\*?$/i.test(t)) cols.style = R.L;
          if (/^quantity$/i.test(t)) cols.qty = R.L;
          if (/^Accessories/i.test(t)) cols.sizeAcc = R.L;
          if (/^SizeK$/i.test(t)) cols.sizeK = R.L; if (/^SizeC$/i.test(t)) cols.sizeC = R.L;
          if (/^V\s*code$/i.test(t)) cols.vcode = R.L;
        }
        if (!hdrRow || !cols.item || !cols.style || !cols.qty) continue;
        const sameRow = (L) => L && [...sh.cells.keys()].some((ref) => ref === L + hdrRow);
        if (!sameRow(cols.style) || !sameRow(cols.qty)) continue;
        let n = 0; for (let r = hdrRow + 1; r <= hdrRow + 2000; r++) { if (cellS(sh, cols.item + r) && cellS(sh, cols.style + r)) n++; else if (n) break; }
        if (!n) continue;
        return { sheet: sh.name, hdrRow, cols, rows: n };
      }
      return null;
    },
    plan(book, ctx, po, opts) {
      const sh = book.sheets.find((s) => s.name === ctx.sheet);
      const { cols } = ctx;
      const formRows = [];
      for (let r = ctx.hdrRow + 1; r <= ctx.hdrRow + 2000; r++) {
        const item = cellS(sh, cols.item + r), style = cellS(sh, cols.style + r);
        if (!item || !style) { if (formRows.length) break; continue; }
        const size = (cols.sizeAcc && cellS(sh, cols.sizeAcc + r)) || (cols.sizeC && cellS(sh, cols.sizeC + r)) || (cols.sizeK && cellS(sh, cols.sizeK + r)) || '';
        formRows.push({ r, item, style, size });
      }
      const items = [...new Set(formRows.map((x) => AZ(x.item)))];
      const poRows = po.items.filter((it) => items.some((k) => k && (AZ(it.supRef).startsWith(k) || AZ(it.spec).includes(k))));
      /* nhóm PO theo item · style · size; dòng gốc = Spec không có "ADD" sau item code */
      const groups = new Map();
      for (const it of poRows) {
        const item = items.find((k) => AZ(it.supRef).startsWith(k) || AZ(it.spec).includes(k));
        const styleM = String(it.spec || '').match(/\b([A-Z]{2}\d{3,4}[A-Z0-9]*)\b(?=\s*\/)/) || String(it.spec || '').match(/\b([A-Z]{2}\d{3,4})\b/);
        const style = styleM ? styleM[1] : '';
        const k = [item, AZ(style), AZ(it.size)].join('|');
        const g = groups.get(k) || { base: null, extra: [], all: [] };
        const isAdd = /\bADD\d*\s*$/i.test(String(it.spec || ''));
        if (!isAdd) g.base = (g.base || 0) + it.qty; else g.extra.push(it);
        g.all.push(it); groups.set(k, g);
      }
      const writes = [], used = new Set(), zero = [];
      for (const f of formRows) {
        const k = [AZ(f.item), AZ(f.style), AZ(f.size)].join('|');
        const g = groups.get(k);
        if (g) { const v = g.base != null ? g.base : Math.max(...g.all.map((x) => x.qty)); writes.push({ sheet: ctx.sheet, ref: cols.qty + f.r, num: v, what: `${f.item} · ${f.style} · ${f.size}${g.extra.length ? ` (bỏ ${g.extra.length} dòng ADD)` : ''}` }); used.add(k); }
        else if (opts.zeroMissing) { writes.push({ sheet: ctx.sheet, ref: cols.qty + f.r, num: 0, what: `${f.item} · ${f.style} · ${f.size} (không có trên PO)` }); zero.push(`${f.style} ${f.size}`); }
      }
      const notInForm = [...groups.keys()].filter((k) => !used.has(k)).map((k) => k.replace(/\|/g, ' · '));
      const extraN = [...groups.values()].reduce((a, g) => a + g.extra.length, 0);
      const head = [];
      if (opts.fillHeader) {
        const findLabel = (re) => { for (const [ref, c] of sh.cells) { const R = splitRef(ref); if (R.row < ctx.hdrRow && R.col <= 4 && re.test(String(c.v || '').replace(/\s+/g, ' ').trim())) return R; } return null; };
        const put = (R, text, what) => { if (!R || !text) return; const ref = colLetter(R.col + 1) + R.row; if (cellS(sh, ref) === '') { writes.push({ sheet: ctx.sheet, ref, text, what }); head.push(`${what}: ${text}`); } else head.push(`${what}: giữ "${cellS(sh, ref)}"`); };
        put(findLabel(/^Order\s*Date\s*:?$/i), po.orderDate, 'Order Date');
        put(findLabel(/Required\s*Ship\s*Date/i), po.etdMin, 'Required Ship Date');
        put(findLabel(/^Company\s*Name\s*:?$/i), po.buyer, 'Bill-to Company');
        put(findLabel(/^Address\s*:?$/i), po.buyerAddr, 'Bill-to Address');
      }
      const baseSum = [...groups.values()].reduce((a, g) => a + (g.base || 0), 0);
      return { writes, report: { formRows: formRows.length, filled: writes.filter((w) => 'num' in w && w.num > 0).length, zero: zero.length, zeroList: zero, notInForm, sumForm: writes.reduce((a, w) => a + (w.num || 0), 0), sumPo: baseSum, poRows: poRows.length, head, sheet: ctx.sheet, note: `item ${items.join(', ')} · ${extraN ? extraN + ' dòng PO "ADD/ADD1" (trang phụ) không cộng vào form · ' : ''}Σ PO tính theo dòng gốc` } };
    },
  });

  /* ===== r-pac (Haddad / Lacoste): OF do Packaging điền sẵn mọi cột trừ số lượng; buyer VLOOKUP số lượng từ PO theo Item / Region / Color / Size.
     PO SCAVI (RPV…): Spec "P6107H Hangtag/7933/34/02S" = item / ref Haddad / mã thị trường / mã màu, cột Size "10A/S". Một PO gồm nhiều ref;
     ref của OF lấy từ tên file ("…20260730_7933_334" → 7933) hoặc người dùng chọn. */
  /* bảng thị trường OF → mã trong Spec PO (buyer xác nhận 12.11.2026): Panama 34 · Brazil H6 · Croatia BDS · SEA MAP INDO · USA_DTC US;
     Europe HBE · Canada CNH (theo dữ liệu PO RPV0033400); thị trường khác: tên trên PO và OF như nhau */
  const RPAC_REGION = { PANAMA: '34', BRAZIL: 'H6', CROATIA: 'BDS', SEA: 'MAPINDO', USADTC: 'US', EUROPE: 'HBE', CANADA: 'CNH' };
  const poRefOf = (it) => { const p = String(it.spec || '').split('/'); return p.length >= 4 ? p[1].trim() : ''; };
  const sizeKey = (t) => AZ(String(t || '').split('/')[0]).replace(/^0+(?=\d)/, '');
  function rpacRefOptions(ctx, po, fname) {
    const refs = [...new Set(po.items.map(poRefOf).filter((r) => /^\d{3,5}$/.test(r)))].sort();
    const toks = String(fname || '').split(/[^A-Za-z0-9]+/);
    const pre = refs.filter((r) => toks.includes(r));
    return { styles: refs, pre: pre.length ? pre : (refs.length === 1 ? refs : []), colors: [], colorsPre: [], coos: [], cooPre: [], poRows: po.items.filter((it) => poRefOf(it)).length, label: 'Ref Haddad' };
  }
  /* r-pac — Lacoste SMS order file (sheet "Data_…"): Vendor ID | PO# | PO Date | Region | Item code | SEASON | Production Date | Order quantity | Style | Color Code | … | US Size */
  PROFILES.push({
    id: 'RPAC_SMS', name: 'r-pac — Lacoste SMS order file (Haddad)',
    detect(book) {
      for (const sh of book.sheets) {
        let hdrRow = 0; const cols = {};
        for (const [ref, c] of sh.cells) {
          const R = splitRef(ref); if (R.row > 12) continue;
          const t = String(c.v || '').replace(/\s+/g, ' ').trim();
          if (/^Order quantity$/i.test(t)) { cols.qty = R.L; hdrRow = R.row; }
          if (/^Item code$/i.test(t)) cols.item = R.L; if (/^Region$/i.test(t)) cols.region = R.L; if (/^Color Code$/i.test(t)) cols.color = R.L;
          if (/^US Size$/i.test(t)) cols.size = R.L; if (/^PO#$/i.test(t)) cols.hpo = R.L; if (/^Vendor ID$/i.test(t)) cols.vendor = R.L; if (/^Style$/i.test(t)) cols.style = R.L;
        }
        if (!hdrRow || !cols.item || !cols.color || !cols.size) continue;
        let n = 0; for (let r = hdrRow + 1; r <= hdrRow + 5000; r++) { const it = cellS(sh, cols.item + r); if (!it) { if (n) break; continue; } if (/^EXAMPLE$/i.test(cellS(sh, (cols.vendor || 'A') + r))) continue; n++; }
        if (!n) continue;
        return { sheet: sh.name, hdrRow, cols, rows: n };
      }
      return null;
    },
    options: rpacRefOptions,
    plan(book, ctx, po, opts) {
      const sh = book.sheets.find((s) => s.name === ctx.sheet); const { cols } = ctx;
      const refs = opts.styles || [];
      const formItems = new Set(); for (let r = ctx.hdrRow + 1; r <= ctx.hdrRow + 5000; r++) { const it = cellS(sh, cols.item + r); if (!it) { if (formItems.size) break; continue; } formItems.add(it.toUpperCase()); }
      let pool = po.items.filter((it) => poRefOf(it) && (!refs.length || refs.includes(poRefOf(it))) && [...formItems].some((k) => String(it.spec).toUpperCase().startsWith(k)));
      const used = new Set(), writes = [], amb = [], miss = [], regionSeen = {};
      let formRows = 0;
      for (let r = ctx.hdrRow + 1; r <= ctx.hdrRow + 5000; r++) {
        const item = cellS(sh, cols.item + r); if (!item) { if (formRows) break; continue; }
        if (/^EXAMPLE$/i.test(cellS(sh, (cols.vendor || 'A') + r))) continue;
        formRows++;
        const region = cols.region ? cellS(sh, cols.region + r) : '', color = cellS(sh, cols.color + r), size = cellS(sh, cols.size + r);
        let cands = pool.filter((it) => !used.has(it) && String(it.spec).toUpperCase().startsWith(item.toUpperCase()) && AZ(String(it.spec).split('/')[3]) === AZ(color) && sizeKey(it.size) === sizeKey(size));
        const code = RPAC_REGION[AZ(region)] || '';
        if (code) { const c1 = cands.filter((it) => AZ(String(it.spec).split('/')[2]) === code); if (c1.length || cands.length > 1) cands = c1; }
        else if (region) {
          const same = cands.filter((it) => AZ(String(it.spec).split('/')[2]) === AZ(region));   // thị trường khác: tên trên PO và OF như nhau
          if (same.length) cands = same;
          else { const known = new Set(Object.values(RPAC_REGION)); const c2 = cands.filter((it) => !known.has(AZ(String(it.spec).split('/')[2]))); if (c2.length) cands = c2; }
        }
        const what = `${item} · ${region} · ${color} · ${size}`;
        if (cands.length === 1 || (cands.length > 1 && new Set(cands.map((x) => x.qty)).size === 1)) {
          const pick = cands[0]; used.add(pick);
          const mk = String(pick.spec).split('/')[2];
          regionSeen[region] = regionSeen[region] || new Set(); regionSeen[region].add(mk);
          writes.push({ sheet: ctx.sheet, ref: cols.qty + r, num: pick.qty, what: what + (cands.length > 1 ? ` (${cands.length} dòng PO cùng SL, thị trường ${cands.map((x) => String(x.spec).split('/')[2]).join('/')})` : '') });
          if (cands.length > 1) amb.push(what);
        } else if (cands.length > 1) { amb.push(what + ` — ${cands.length} dòng PO khác SL (${cands.map((x) => String(x.spec).split('/')[2] + ':' + x.qty).join(', ')}) → để trống`); }
        else { miss.push(what); if (opts.zeroMissing) writes.push({ sheet: ctx.sheet, ref: cols.qty + r, num: 0, what: what + ' (không có trên PO)' }); }
      }
      const notInForm = pool.filter((it) => !used.has(it)).map((it) => `${it.spec} ${it.size} = ${it.qty}`);
      const regNote = Object.entries(regionSeen).map(([k, v]) => `${k} → ${[...v].join('/')}`).join(', ');
      return { writes, report: { formRows, filled: writes.filter((w) => w.num > 0).length, zero: miss.length, zeroList: miss, notInForm, sumForm: writes.reduce((a, w) => a + (w.num || 0), 0), sumPo: pool.reduce((a, it) => a + it.qty, 0), poRows: pool.length, head: [], sheet: ctx.sheet,
        note: `ref ${refs.join(', ') || '(tất cả)'} · khóa item · màu · size · thị trường${regNote ? ' (' + regNote + ')' : ''}${amb.length ? ` · ⚠ ${amb.length} dòng nhiều ứng viên: ${amb.slice(0, 4).join('; ')}${amb.length > 4 ? '…' : ''}` : ''}` } };
    },
  });
  /* r-pac — HADDAD Korean Hang Tag Order Form (sheet "Order Form …"): ID | HADDAD PO# | PO Date | Vendor ID | … | Ticket Item# | Lacoste Style | Season 1 | Season 2 |
     Korean colour | Korean size | QTY (PCS). Màu Korea → màu Lacoste qua sheet danh mục (cột "LACOSTE KOREA STYLE # wo…" + "KOREA COLOR #" → "COLOR").
     PO: Spec "Hangtag FC7648/7867/KOREA/S3C", Size "10A/S" (so phần trước "/" với Korean size, bỏ số 0 đầu). */
  PROFILES.push({
    id: 'RPAC_KHT', name: 'r-pac — HADDAD Korean Hang Tag Order Form',
    detect(book) {
      for (const sh of book.sheets) {
        let hdrRow = 0; const cols = {};
        for (const [ref, c] of sh.cells) {
          const R = splitRef(ref); if (R.row > 6) continue;
          const t = String(c.v || '').replace(/\s+/g, ' ').trim();
          if (/^QTY\s*\(PCS\)$/i.test(t)) { cols.qty = R.L; hdrRow = R.row; }
          if (/^Ticket Item#?$/i.test(t)) cols.item = R.L; if (/^Lacoste\s*Style$/i.test(t)) cols.style = R.L; if (/^Korean colou?r$/i.test(t)) cols.kcol = R.L;
          if (/^Korean size$/i.test(t)) cols.size = R.L; if (/^HADDAD PO#?$/i.test(t)) cols.hpo = R.L;
        }
        if (!hdrRow || !cols.item || !cols.kcol || !cols.size) continue;
        let n = 0; for (let r = hdrRow + 1; r <= hdrRow + 3000; r++) { const it = cellS(sh, cols.item + r); if (!it) { if (n) break; continue; } if (/^EXAMPLE$/i.test(cellS(sh, (cols.hpo || 'B') + r))) continue; n++; }
        if (!n) continue;
        /* bảng màu Korea ↔ Lacoste */
        const cmap = new Map();
        for (const s2 of book.sheets) {
          let h = 0; const cc = {};
          for (const [ref, c] of s2.cells) { const R = splitRef(ref); if (R.row > 3) continue; const t = String(c.v || '').trim(); if (/^COLOR$/i.test(t)) { cc.color = R.L; h = R.row; } if (/^KOREA COLOR/i.test(t)) cc.kcol = R.L; if (/^LACOSTE KOREA STYLE # wo/i.test(t)) cc.style = R.L; }
          if (!h || !cc.color || !cc.kcol) continue;
          for (let r = h + 1; r <= h + 5000; r++) { const k = cellS(s2, cc.kcol + r), col = cellS(s2, cc.color + r); if (!k && !col) { if (r > h + 50) break; continue; } if (!k || !col) continue; const st = cc.style ? cellS(s2, cc.style + r) : ''; cmap.set(AZ(st) + '|' + AZ(k), col); if (!cmap.has('|' + AZ(k))) cmap.set('|' + AZ(k), col); }
        }
        return { sheet: sh.name, hdrRow, cols, rows: n, cmap: [...cmap] };
      }
      return null;
    },
    options: rpacRefOptions,
    plan(book, ctx, po, opts) {
      const sh = book.sheets.find((s) => s.name === ctx.sheet); const { cols } = ctx; const cmap = new Map(ctx.cmap || []);
      const refs = opts.styles || [];
      const formItems = new Set(); for (let r = ctx.hdrRow + 1; r <= ctx.hdrRow + 3000; r++) { const it = cellS(sh, cols.item + r); if (!it) { if (formItems.size) break; continue; } formItems.add(AZ(it)); }
      const pool = po.items.filter((it) => poRefOf(it) && (!refs.length || refs.includes(poRefOf(it))) && [...formItems].some((k) => AZ(String(it.spec).split('/')[0]).includes(k)));
      const used = new Set(), writes = [], miss = [], amb = []; let formRows = 0;
      for (let r = ctx.hdrRow + 1; r <= ctx.hdrRow + 3000; r++) {
        const item = cellS(sh, cols.item + r); if (!item) { if (formRows) break; continue; }
        if (/^EXAMPLE$/i.test(cellS(sh, (cols.hpo || 'B') + r))) continue;
        formRows++;
        const style = cols.style ? cellS(sh, cols.style + r) : '', kcol = cellS(sh, cols.kcol + r), size = cellS(sh, cols.size + r);
        const lcol = cmap.get(AZ(style) + '|' + AZ(kcol)) || cmap.get('|' + AZ(kcol)) || kcol;
        const cands = pool.filter((it) => !used.has(it) && AZ(String(it.spec).split('/')[0]).includes(AZ(item)) && AZ(String(it.spec).split('/')[3]) === AZ(lcol) && sizeKey(it.size) === sizeKey(size));
        const what = `${item} · ${style} · ${kcol}${lcol !== kcol ? '→' + lcol : ''} · ${size}`;
        if (cands.length === 1 || (cands.length > 1 && new Set(cands.map((x) => x.qty)).size === 1)) { used.add(cands[0]); writes.push({ sheet: ctx.sheet, ref: cols.qty + r, num: cands[0].qty, what }); if (cands.length > 1) amb.push(what); }
        else if (cands.length > 1) amb.push(what + ' — nhiều dòng PO khác SL → để trống');
        else { miss.push(what); if (opts.zeroMissing) writes.push({ sheet: ctx.sheet, ref: cols.qty + r, num: 0, what: what + ' (không có trên PO)' }); }
      }
      const notInForm = pool.filter((it) => !used.has(it)).map((it) => `${it.spec} ${it.size} = ${it.qty}`);
      return { writes, report: { formRows, filled: writes.filter((w) => w.num > 0).length, zero: miss.length, zeroList: miss, notInForm, sumForm: writes.reduce((a, w) => a + (w.num || 0), 0), sumPo: pool.reduce((a, it) => a + it.qty, 0), poRows: pool.length, head: [], sheet: ctx.sheet, note: `ref ${refs.join(', ') || '(tất cả)'} · màu Korea → Lacoste qua sheet danh mục (${cmap.size} mã)${amb.length ? ' · ⚠ ' + amb.join('; ') : ''}` } };
    },
  });

  /* Silver Printing — EMPORIO ARMANI ORDER FORM: form chuẩn chung (catalogue ITEM NO. | DESCRIPTION | PCS | UNIT PRICE, hai cột trái A/B/C và phải G/I/J).
     PO SCAVI (SLV…): mã Armani sau dấu "#" trong Spec / Supp. Ref ("PK11330 Z8004 Y0041#GA010XX1484CB"); gộp theo mã. Mã có sẵn trên form → ghi PCS vào
     dòng đó; mã chưa có → ghi ITEM NO · DESCRIPTION · PCS vào các dòng trống dưới bảng (trái trước, phải sau), như buyer copy từ PO dán vào.
     Phần đầu (DATE / PO# / DELY DATE) của form mẫu là của PO cũ → ghi đè bằng PO mới. */
  PROFILES.push({
    id: 'SILVER_ARMANI', name: 'Silver Printing — Emporio Armani order form',
    detect(book) {
      const cands = [];
      for (const sh of book.sheets) {
        let hdrRow = 0; let armani = false; const pairs = [];
        for (const [ref, c] of sh.cells) {
          const R = splitRef(ref); const t = String(c.v || '').replace(/\s+/g, ' ').trim();
          if (/ARMANI ORDER FORM/i.test(t)) armani = true;
          if (/^ITEM NO\.?$/i.test(t) && R.row <= 40) { hdrRow = hdrRow || R.row; pairs.push(R); }
        }
        if (!armani || !hdrRow) continue;
        /* mỗi cặp: cột ITEM NO, cột DESCRIPTION (kế bên có chữ), cột PCS (ô "PCS" cùng dòng tiêu đề gần nhất bên phải) */
        const row = pairs.filter((R) => R.row === hdrRow).sort((a, b) => a.col - b.col);
        const pcsCols = []; for (const [ref, c] of sh.cells) { const R = splitRef(ref); if (R.row === hdrRow && /^PCS$/i.test(String(c.v || '').trim())) pcsCols.push(R.col); }
        const cols = row.map((R) => { const pcs = pcsCols.filter((c) => c > R.col).sort((a, b) => a - b)[0]; const descC = [...sh.cells.entries()].filter(([k, c]) => c.v != null && String(c.v).trim() !== '').map(([k]) => splitRef(k)).filter((x) => x.row === hdrRow && x.col > R.col && x.col < (pcs || 99)).sort((a, b) => a.col - b.col)[0]; return pcs ? { item: R.L, desc: descC ? descC.L : colLetter(R.col + 1), pcs: colLetter(pcs) } : null; }).filter(Boolean);
        if (!cols.length) continue;
        let footer = 0, last = hdrRow; const items = new Map();
        for (let r = hdrRow + 1; r <= hdrRow + 600; r++) {
          const a = cellS(sh, 'A' + r);
          if (/^\d\)\s/.test(a) || /PRODUCTION LEADTIME|TRACKING NUMBER/i.test(a)) { footer = r; break; }
          for (const cp of cols) { const v = cellS(sh, cp.item + r); const m = v.match(/GA\d{3}[A-Z]{2}\d{4}[A-Z]{2}/); if (m) { items.set(m[0], { r, cp }); last = r; } }
        }
        cands.push({ sheet: sh.name, hdrRow, cols, footer: footer || hdrRow + 600, last, items: [...items], n: items.size });
      }
      if (!cands.length) return null;
      return Object.assign({ all: cands.map((c) => c.sheet), cands }, cands[0]);
    },
    options(ctx, po, fname) { return { styles: ctx.all, pre: [ctx.sheet], colors: [], colorsPre: [], coos: [], cooPre: [], poRows: po.items.length, label: 'Sheet điền' }; },
    plan(book, ctx0, po, opts) {
      const pick = (opts.styles || [])[0]; const ctx = (ctx0.cands || []).find((c) => c.sheet === pick) || ctx0;
      const sh = book.sheets.find((s) => s.name === ctx.sheet);
      const agg = new Map();
      for (const it of po.items) {
        const m = (String(it.spec || '') + ' ' + String(it.supRef || '')).match(/#\s*(GA\d{3}[A-Z]{2}\d{4}[A-Z]{2})/);
        const code = m ? m[1] : (String(it.supRef || '').trim() || '?');
        const a = agg.get(code) || { qty: 0, n: 0, desc: new Set(), sup: new Set() };
        a.qty += it.qty; a.n++; if (it.desc) a.desc.add(it.desc); if (it.supRef) a.sup.add(it.supRef); agg.set(code, a);
      }
      const existing = new Map(ctx.items);
      const writes = [], placedExisting = [], placedNew = [], noRoom = [];
      /* dòng trống: dưới dòng hàng cuối tới trước footer, trái (cặp 1) rồi phải (cặp 2) */
      const slots = [];
      for (const cp of ctx.cols) for (let r = ctx.hdrRow + 1; r < ctx.footer; r++) { if (!cellS(sh, cp.item + r) && !cellS(sh, cp.desc + r) && !cellS(sh, cp.pcs + r)) slots.push({ r, cp }); }
      slots.sort((a, b) => (ctx.cols.indexOf(a.cp) - ctx.cols.indexOf(b.cp)) || (a.r - b.r));
      let si = 0;
      for (const [code, a] of [...agg].sort()) {
        const descTxt = [...a.desc][0] ? ([...a.desc][0] + ([...a.sup][0] ? ' / ' + [...a.sup][0] : '')) : [...a.sup][0] || '';
        if (existing.has(code)) { const { r, cp } = existing.get(code); writes.push({ sheet: ctx.sheet, ref: cp.pcs + r, num: a.qty, what: `${code} (dòng có sẵn ${r})` }); placedExisting.push(code); continue; }
        const sl = slots[si++];
        if (!sl) { noRoom.push(`${code} = ${fmtN(a.qty)}`); continue; }
        writes.push({ sheet: ctx.sheet, ref: sl.cp.item + sl.r, text: code, what: `${code} → dòng trống ${sl.r}` });
        writes.push({ sheet: ctx.sheet, ref: sl.cp.desc + sl.r, text: descTxt, what: `mô tả ${code}` });
        writes.push({ sheet: ctx.sheet, ref: sl.cp.pcs + sl.r, num: a.qty, what: `PCS ${code}` });
        placedNew.push(code);
      }
      const head = [];
      if (opts.fillHeader) {
        const findLabel = (re) => { for (const [ref, c] of sh.cells) { const R = splitRef(ref); if (R.row < ctx.hdrRow && re.test(String(c.v || '').replace(/\s+/g, ' ').trim())) return R; } return null; };
        const over = (R, text, what) => { if (!R || !text) return; const ref = colLetter(R.col + 1) + R.row; const old = cellS(sh, ref); writes.push({ sheet: ctx.sheet, ref, text, what }); head.push(`${what}: ${text}${old && old !== text ? ` (thay "${old}")` : ''}`); };
        const dmy = (d) => `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}.${d.getFullYear()}`;
        const today = new Date(), d21 = new Date(today.getTime() + 21 * 86400000);
        over(findLabel(/^PO#?\s*:?$/i), po.po, 'PO#');
        over(findLabel(/^DATE\s*:?$/i), dmy(today), 'DATE (hôm nay)');
        over(findLabel(/^DELY\s*DATE\s*:?$/i), po.etdMin ? po.etdMin.replace(/\//g, '.') : dmy(d21), po.etdMin ? 'DELY DATE (ETD PO)' : 'DELY DATE (hôm nay + 3 tuần, sửa nếu cần)');
      }
      const sumForm = writes.filter((w) => 'num' in w).reduce((a, w) => a + w.num, 0), sumPo = po.items.reduce((a, it) => a + it.qty, 0);
      return { writes, report: { formRows: existing.size, filled: placedExisting.length + placedNew.length, zero: 0, zeroList: [], notInForm: noRoom, sumForm, sumPo, poRows: po.items.length, head, sheet: ctx.sheet,
        note: `${agg.size} mã trên PO: ${placedExisting.length} có sẵn trên form, ${placedNew.length} ghi vào dòng trống (dòng ${slots[0] ? slots[0].r : '-'}…), ${noRoom.length} hết chỗ${(ctx0.all || []).length > 1 ? ` · sheet khác: ${ctx0.all.filter((x) => x !== ctx.sheet).join(', ')}` : ''}` } };
    },
  });

  /* Avery Dennison Hong Kong — Adore Me order form (PFL care label / HTL label / paper tag): mỗi sheet một Item Code (ô dưới nhãn "Item Code"),
     dòng tiêu đề "PO# | *Size | *RN# | *Style No | … | Quantity (pcs)" (paper tag: "*Size | *Item description | *UPC | Quantity (pcs) | … | STYLE ADOREME | COLOR").
     PO (AV6…): Supp. Ref = Item Code; style ở cột Reference (phần trước "_": "SW-14113-BIK_SCA057-PA-BIK-X" → SW-14113-BIK) hoặc Spec ("SW-15787-TPI/ Dot 134"),
     hoặc Color ("GREY-SW-09993-TPI"); Size = cột Size. Khóa: item code · style · size. Sheet có cột PO# ghi PO khác (AV60003300…) → bỏ qua. */
  const AM_ALIAS = { AMTAG2601: ['AMMAIN001'] };
  PROFILES.push({
    id: 'AVERY_HK_ADOREME', name: 'Avery Dennison HK — Adore Me order form',
    detect(book) {
      const sheets = [];
      for (const sh of book.sheets) {
        let itemRef = null, hdrRow = 0; const cols = {};
        for (const [ref, c] of sh.cells) {
          const R = splitRef(ref); if (R.row > 40) continue;
          const t = String(c.v || '').replace(/\s+/g, ' ').trim();
          if (/^Item Code$/i.test(t) && !itemRef) itemRef = R;
          if (/^Quantity\s*\(pcs\)$/i.test(t)) { cols.qty = R.L; hdrRow = R.row; }
        }
        if (!itemRef || !hdrRow) continue;
        for (const [ref, c] of sh.cells) {
          const R = splitRef(ref); if (R.row !== hdrRow) continue;
          const t = String(c.v || '').replace(/\s+/g, ' ').trim();
          if (/^\*?Size\b/i.test(t)) cols.size = R.L; if (/^\*?Style No/i.test(t) || /^STYLE ADOREME/i.test(t)) cols.style = R.L; if (/^PO#?$/i.test(t)) cols.po = R.L; if (/^\*?Item description/i.test(t)) cols.sku = R.L;
          if (/^COLOU?R$/i.test(t)) cols.color = R.L;
        }
        if (!cols.size || !cols.qty) continue;
        const itemCode = cellS(sh, itemRef.L + (itemRef.row + 1));
        if (!itemCode) continue;
        let n = 0; const pos = new Set();
        for (let r = hdrRow + 1; r <= hdrRow + 3000; r++) { const sz = cellS(sh, cols.size + r); if (!sz) { if (n) break; continue; } n++; if (cols.po) { const p = cellS(sh, cols.po + r); if (p) pos.add(p.toUpperCase()); } }
        if (!n) continue;
        sheets.push({ sheet: sh.name, itemCode, hdrRow, cols, rows: n, pos: [...pos] });
      }
      if (!sheets.length) return null;
      return { sheet: sheets.map((x) => x.sheet).join(' + '), sheets, rows: sheets.reduce((a, x) => a + x.rows, 0) };
    },
    plan(book, ctx, po, opts) {
      /* style Adore Me dạng SW-14110-TPI / PA-15906-HIP nằm ở Reference ("SW-14113-BIK_SCA057-PA-BIK-X", PO mới: "SCA057-BR-BAL-SW-14110-TPI"), Spec hoặc Color */
      const STYLE_RE = /\b[A-Z]{2}-\d{4,6}-[A-Z]{2,4}\b/;
      const styleOf = (it) => { for (const t of [it.reference, it.spec, it.color]) { const m = String(t || '').toUpperCase().match(STYLE_RE); if (m) return m[0]; } const r = String(it.reference || '').trim(); if (r) return r.split('_')[0].trim(); return String(it.spec || '').split('/')[0].trim(); };
      const formStyle = (t) => { const m = String(t || '').toUpperCase().match(STYLE_RE); return m ? m[0] : String(t || '').trim(); };
      const writes = [], notes = [], skipped = [], zeroAll = [], notInFormAll = [];
      let formRows = 0, filled = 0, sumForm = 0, sumPo = 0, poRowsN = 0;
      for (const S of ctx.sheets) {
        const sh = book.sheets.find((s) => s.name === S.sheet);
        const other = S.pos.filter((p) => p !== String(po.po).toUpperCase());
        if (other.length && !S.pos.includes(String(po.po).toUpperCase())) { skipped.push(`${S.sheet} (PO ${other.join(', ')})`); continue; }
        let rows = po.items.filter((it) => AZ(it.supRef) === AZ(S.itemCode));
        if (!rows.length) rows = po.items.filter((it) => (AM_ALIAS[AZ(S.itemCode)] || []).includes(AZ(it.supRef)));   // mã tương đương chỉ dùng khi PO không có đúng Item Code
        poRowsN += rows.length; sumPo += rows.reduce((a, it) => a + it.qty, 0);
        if (!rows.length) { notes.push(`${S.sheet}: PO không có Supp. Ref ${S.itemCode}`); continue; }
        const hasSize = (it) => it.size && !/^UNKNOWN$/i.test(it.size);
        if (rows.every((it) => !hasSize(it))) { notes.push(`${S.sheet}: PO chỉ có ${rows.length} dòng tổng ${fmtN(rows.reduce((a, it) => a + it.qty, 0))} cho ${S.itemCode} (${rows.map((it) => it.supRef).filter((v, i, a) => a.indexOf(v) === i).join(', ')}), không chia theo size → không điền ${S.rows} dòng form`); continue; }
        const used = new Set(); const zero = [];
        for (let r = S.hdrRow + 1; r <= S.hdrRow + 3000; r++) {
          const size = cellS(sh, S.cols.size + r); if (!size) { if (formRows) break; continue; }
          formRows++;
          let style = S.cols.style ? formStyle(cellS(sh, S.cols.style + r)) : '';
          if (!style && S.cols.sku) style = formStyle(cellS(sh, S.cols.sku + r)) === cellS(sh, S.cols.sku + r) ? '' : formStyle(cellS(sh, S.cols.sku + r));
          let cands = rows.filter((it) => !used.has(it) && AZ(it.size) === AZ(size) && (!style || AZ(styleOf(it)) === AZ(style) || AZ(styleOf(it)).startsWith(AZ(style))));
          /* paper tag: cùng style + size nhưng khác màu (BEGONIA PINK / Meadow Green) → so cột COLOR của form với phần 2 của Spec ("SW-15808-TPI/BEGONIA PINK/Dot 134") */
          const color = S.cols.color ? cellS(sh, S.cols.color + r) : '';
          const colorOf = (it) => { const p = String(it.spec || '').split('/'); return p.length >= 3 ? p[1].trim() : String(it.color || '').replace(/-.*$/, ''); };
          if (cands.length > 1 && color) { const c2 = cands.filter((it) => AZ(colorOf(it)) === AZ(color)); if (c2.length) cands = c2; }
          const what = `${S.sheet} · ${style || '?'}${color ? ' · ' + color : ''} · ${size}`;
          if (cands.length > 1 && new Set(cands.map((x) => x.qty)).size > 1) { zero.push(`${style} ${size}: ${cands.length} dòng PO khác màu (${cands.map((it) => colorOf(it) + ':' + it.qty).join(', ')}) — form không ghi màu`); continue; }
          if (cands.length) { used.add(cands[0]); writes.push({ sheet: S.sheet, ref: S.cols.qty + r, num: cands[0].qty, what }); filled++; sumForm += cands[0].qty; }
          else { zero.push(`${style} ${size}`); if (opts.zeroMissing) writes.push({ sheet: S.sheet, ref: S.cols.qty + r, num: 0, what: what + ' (không có trên PO)' }); }
        }
        zeroAll.push(...zero.map((z) => S.sheet + ': ' + z));
        notInFormAll.push(...rows.filter((it) => !used.has(it)).map((it) => `${S.sheet}: ${styleOf(it)} ${it.size} = ${fmtN(it.qty)}`));
      }
      return { writes, report: { formRows, filled, zero: zeroAll.length, zeroList: zeroAll, notInForm: notInFormAll, sumForm, sumPo, poRows: poRowsN, head: [], sheet: ctx.sheets.map((x) => `${x.sheet} (${x.itemCode})`).join(', '),
        note: [skipped.length ? `bỏ qua sheet của PO khác: ${skipped.join('; ')}` : '', ...notes].filter(Boolean).join(' · ') } };
    },
  });

  function detectForm(book) {
    for (const p of PROFILES) { const ctx = p.detect(book); if (ctx) return { profile: p, ctx }; }
    return null;
  }
  /* áp dụng writes vào book → Uint8Array xlsx mới */
  async function applyWrites(book, writes) {
    const changed = {};
    const bySheet = new Map();
    for (const w of writes) { if (!bySheet.has(w.sheet)) bySheet.set(w.sheet, []); bySheet.get(w.sheet).push(w); }
    for (const [name, ws] of bySheet) {
      const sh = book.sheets.find((s) => s.name === name);
      let xml = sh.xml;
      /* style gợi ý cho ô mới: style của ô cùng cột ở dòng trên có dữ liệu */
      for (const w of ws) {
        const R = splitRef(w.ref);
        let hint = '';
        for (let r = R.row - 1; r >= Math.max(1, R.row - 30) && !hint; r--) { const c = sh.cells.get(R.L + r); if (c && c.s) hint = c.s; }
        xml = setCell(xml, w.ref, 'num' in w ? { num: w.num } : { text: w.text }, hint);
      }
      changed[sh.path] = xml;
    }
    changed['xl/workbook.xml'] = forceRecalc(book.workbookXml);
    return zipWrite(book.zip, changed);
  }

  const OrderCore = { AZ, colIdx, colLetter, zipRead, zipWrite, zipEntryText, openBook, parseSheetCells, parseSharedStrings, setCell, forceRecalc, readScaviPo, parseSpec, PROFILES, detectForm, applyWrites, cellS, cellV, fmtN, hasStreams, serialToDMY };
  if (typeof module !== 'undefined' && module.exports) module.exports = OrderCore;
  root.OrderCore = OrderCore;

  /* ---------------------------------------- GIAO DIỆN ---------------------------------------- */
  if (typeof document === 'undefined') return;
  const $ = (s) => document.querySelector(s);
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const ST = { pos: [], forms: [], outputs: [] };

  async function addFiles(list) {
    const files = [...list].filter((f) => /\.xlsx$/i.test(f.name));
    const skipped = [...list].filter((f) => !/\.xlsx$/i.test(f.name)).map((f) => f.name);
    for (const f of files) {
      try {
        const buf = new Uint8Array(await f.arrayBuffer());
        const book = await openBook(buf);
        const po = readScaviPo(book);
        if (po && po.items.length) { ST.pos = ST.pos.filter((p) => p.file.name !== f.name); ST.pos.push({ file: f, po }); continue; }
        const det = detectForm(book);
        if (det) { ST.forms = ST.forms.filter((x) => x.file.name !== f.name); ST.forms.push({ file: f, book, profile: det.profile, ctx: det.ctx, sel: null }); continue; }
        skipped.push(f.name + ' (không nhận ra PO SCAVI hay form mẫu)');
      } catch (e) { skipped.push(f.name + ' (' + (e && e.message ? e.message : e) + ')'); }
    }
    $('#detect').textContent = skipped.length ? 'Bỏ qua: ' + skipped.join(' · ') : '';
    render();
  }
  function poFor(form) {
    if (ST.pos.length === 1) return ST.pos[0];
    const fn = form.file.name.toUpperCase();
    return ST.pos.find((p) => fn.includes(String(p.po.po).toUpperCase())) || ST.pos[0] || null;
  }
  function render() {
    let h = '';
    if (ST.pos.length) {
      h += '<table><thead><tr><th>PO SCAVI</th><th>Supplier</th><th>Order date</th><th>ETD sớm nhất</th><th class="n">Dòng</th><th class="n">Tổng SL</th><th>Loại hàng (Supp. Ref)</th></tr></thead><tbody>';
      for (const p of ST.pos) {
        const refs = {}; p.po.items.forEach((it) => { refs[it.supRef || '?'] = (refs[it.supRef || '?'] || 0) + it.qty; });
        h += `<tr><td><b>${esc(p.po.po)}</b><span class="fn">${esc(p.file.name)}</span></td><td>${esc(p.po.supplier)}</td><td>${esc(p.po.orderDate)}</td><td>${esc(p.po.etdMin)}</td><td class="n">${p.po.items.length}</td><td class="n">${fmtN(p.po.items.reduce((a, b) => a + b.qty, 0))}</td><td class="small">${Object.entries(refs).map(([k, v]) => `${esc(k)}: ${fmtN(v)}`).join('<br>')}</td></tr>`;
      }
      h += '</tbody></table>';
    } else h += '<div class="hint">Chưa có file PO SCAVI (Excel "PURCHASE ORDER" có bảng Material Code | Supp. Ref. | Spec. | Qty.).</div>';
    if (ST.forms.length) {
      h += '<table style="margin-top:12px"><thead><tr><th>Form supplier</th><th>Hồ sơ mẫu</th><th>Sheet điền</th><th>Ghép với PO</th><th>Tùy chọn</th></tr></thead><tbody>';
      ST.forms.forEach((f, i) => {
        const p = poFor(f);
        let opt = '';
        if (f.profile.options && p) {
          const o = f.profile.options(f.ctx, p.po, f.file.name);
          if (!f.sel) f.sel = { styles: o.pre.slice(), colors: o.colorsPre.slice(), coos: o.cooPre.slice() };
          opt += `<div class="small">Item <b>${esc(f.ctx.itemCode)}</b> — ${o.poRows} dòng PO</div>`;
          opt += `<div class="small">${esc(o.label || 'Style')}: ${o.styles.map((s) => `<label class="chk"><input type="checkbox" data-f="${i}" data-k="styles" value="${esc(s)}" ${f.sel.styles.includes(s) ? 'checked' : ''}> ${esc(s)}</label>`).join(' ') || '(không có trong PO)'}</div>`;
          if (o.colors.length > 1) opt += `<div class="small">Màu (bỏ trống = tất cả): ${o.colors.map((s) => `<label class="chk"><input type="checkbox" data-f="${i}" data-k="colors" value="${esc(s)}" ${f.sel.colors.includes(s) ? 'checked' : ''}> ${esc(s)}</label>`).join(' ')}</div>`;
          if (o.coos.length > 1) opt += `<div class="small">COO: ${o.coos.map((s) => `<label class="chk"><input type="checkbox" data-f="${i}" data-k="coos" value="${esc(s)}" ${f.sel.coos.includes(s) ? 'checked' : ''}> ${esc(s)}</label>`).join(' ')}</div>`;
        } else if (f.ctx.rows != null) opt = `<span class="small">${f.ctx.rows} dòng SKU có sẵn</span>`;
        h += `<tr><td><b>${esc(f.file.name)}</b></td><td>${esc(f.profile.name)}</td><td>${esc(f.ctx.sheet)}</td><td>${p ? esc(p.po.po) : '<span class="bad">chưa có PO</span>'}</td><td>${opt}</td></tr>`;
      });
      h += '</tbody></table>';
    } else h += '<div class="hint" style="margin-top:8px">Chưa có form mẫu của supplier (hiện hỗ trợ: Avery Lands\' End STYLE ID Label · Avery size / care label order form).</div>';
    $('#files').innerHTML = h;
    $('#files').querySelectorAll('input[type=checkbox]').forEach((cb) => cb.addEventListener('change', (e) => {
      const f = ST.forms[+e.target.dataset.f], k = e.target.dataset.k;
      const vals = [...$('#files').querySelectorAll(`input[data-f="${e.target.dataset.f}"][data-k="${k}"]:checked`)].map((x) => x.value);
      f.sel[k] = vals;
    }));
    $('#run').disabled = !(ST.pos.length && ST.forms.length);
  }
  async function run() {
    $('#run').disabled = true; $('#dls').innerHTML = ''; $('#report').innerHTML = '';
    ST.outputs.forEach((o) => URL.revokeObjectURL(o.url)); ST.outputs = [];
    const opts0 = { zeroMissing: $('#zeroMissing').checked, fillHeader: $('#fillHeader').checked };
    let h = '';
    for (const f of ST.forms) {
      const p = poFor(f);
      if (!p) { h += `<div class="card warn"><b>${esc(f.file.name)}</b>: chưa có PO để ghép.</div>`; continue; }
      try {
        const opts = Object.assign({}, opts0, f.sel || {});
        const { writes, report } = f.profile.plan(f.book, f.ctx, p.po, opts);
        const out = await applyWrites(f.book, writes);
        const base = f.file.name.replace(/\.xlsx$/i, '');
        const fname = (base.toUpperCase().includes(String(p.po.po).toUpperCase()) ? base : `${base} - ${p.po.po}`) + '.xlsx';
        const url = URL.createObjectURL(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
        ST.outputs.push({ url, fname });
        const ok = Math.abs(report.sumForm - report.sumPo) < 0.5 && !report.notInForm.length;
        h += `<div class="card ${ok ? 'ok' : 'warn'}"><div><b>${esc(fname)}</b> <span class="tag ${ok ? 'good' : 'bad'}">${ok ? 'Σ form = Σ PO' : 'cần xem lại'}</span></div>
          <div class="small">${esc(f.profile.name)} · sheet <b>${esc(report.sheet)}</b> · PO ${esc(p.po.po)}${report.note ? ' · ' + esc(report.note) : ''}</div>
          <div class="lines"><span>Dòng form: <b>${report.formRows}</b></span><span>Điền SL &gt; 0: <b>${report.filled}</b></span><span>Ghi 0: <b>${report.zero}</b></span><span>Dòng PO dùng: <b>${report.poRows}</b></span><span>Σ form: <b>${fmtN(report.sumForm)}</b></span><span>Σ PO: <b>${fmtN(report.sumPo)}</b></span></div>
          ${report.notInForm.length ? `<div class="li bad">⚠ ${report.notInForm.length} mã trên PO không có dòng trên form: ${esc(report.notInForm.slice(0, 20).join(', '))}${report.notInForm.length > 20 ? '…' : ''}</div>` : ''}
          ${report.zeroList.length ? `<div class="li">Dòng form không có trên PO (ghi 0): ${esc(report.zeroList.slice(0, 15).join(', '))}${report.zeroList.length > 15 ? ` … (+${report.zeroList.length - 15})` : ''}</div>` : ''}
          ${report.head.length ? `<div class="li">Phần đầu form: ${esc(report.head.join(' · '))}</div>` : ''}
          <details><summary class="small">Các ô đã ghi (${writes.length})</summary><div class="small mono">${writes.map((w) => `${esc(w.ref)} = ${'num' in w ? fmtN(w.num) : esc(w.text)} — ${esc(w.what || '')}`).join('<br>')}</div></details></div>`;
      } catch (e) { h += `<div class="card warn"><b>${esc(f.file.name)}</b>: lỗi — ${esc(e && e.message ? e.message : e)}</div>`; }
    }
    $('#report').innerHTML = h || '<p class="sub">Không có gì để điền.</p>';
    $('#dls').innerHTML = ST.outputs.map((o) => `<a class="dl" href="${o.url}" download="${esc(o.fname)}">⬇ ${esc(o.fname)}</a>`).join('');
    $('#run').disabled = false;
  }
  const drop = $('#drop'), inp = drop.querySelector('input');
  drop.addEventListener('click', () => inp.click());
  inp.addEventListener('change', () => { addFiles(inp.files); inp.value = ''; });
  ['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('drag'); }));
  ['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('drag'); }));
  drop.addEventListener('drop', (e) => addFiles(e.dataTransfer.files));
  $('#run').addEventListener('click', run);
  $('#clear').addEventListener('click', () => { ST.pos = []; ST.forms = []; ST.outputs.forEach((o) => URL.revokeObjectURL(o.url)); ST.outputs = []; $('#dls').innerHTML = ''; $('#report').innerHTML = ''; $('#detect').textContent = ''; render(); });
  if (!hasStreams) $('#detect').textContent = 'Trình duyệt không hỗ trợ nén/giải nén (DecompressionStream) — hãy dùng Chrome / Edge mới.';
  render();
})(typeof window !== 'undefined' ? window : globalThis);
