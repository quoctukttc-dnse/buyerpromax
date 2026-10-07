import { chromium } from 'playwright';
import fs from 'fs'; import path from 'path';
const dir = process.argv[2];
const B = await chromium.launch({ args:['--headless=new','--no-sandbox'], headless:false, executablePath:'/opt/pw-browsers/chromium' });
const pg = await B.newPage();
await pg.addInitScript(() => { window.OCR_CFG = { script: 'http://127.0.0.1:8899/ocr/tesseract.min.js', workerPath: 'http://127.0.0.1:8899/ocr/worker.min.js', corePath: 'http://127.0.0.1:8899/ocr/core/', langPath: 'http://127.0.0.1:8899/ocr/lang' }; });
await pg.goto('http://127.0.0.1:8899/index.html'); await pg.waitForSelector('#run');
const all=[]; (function walk(d){for(const f of fs.readdirSync(d,{withFileTypes:true})){const p=path.join(d,f.name); if(f.isDirectory()) walk(p); else if(!f.name.startsWith('.')) all.push(p);} })(dir);
await pg.setInputFiles('#dropall input', all);
try { await pg.waitForFunction(() => /Hoàn tất|Chưa có hóa đơn|Lỗi:/.test(document.querySelector('#log').innerText), null, { timeout: 400000 }); } catch (e) {}
console.log(JSON.stringify(await pg.evaluate(() => ({
  items: STATE.items.map((x) => { const c = CACHE.get(x.file); return { name: x.file.name, kind: x.kind, ckind: c && c.kind, headerOnly: c && c.headerOnly, ocr: c && c.ocr, invNo: x.inv && x.inv.invNo, date: x.inv && x.inv.invDate, n: x.inv ? x.inv.items.length : -1, fabOwner: !!x.fabOwner, pdfInv: !!x.pdfInv, pklOnly: x.fab && x.fab.pklOnly }; }),
  dup: (STATE.dupOcr || []).map((d) => d.o.file.name + ' → ' + d.twin.file.name), headerOnly: (STATE.headerOnly || []).map((x) => x.file.name),
})), null, 1));
await B.close();
