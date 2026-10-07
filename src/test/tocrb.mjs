import { chromium } from 'playwright';
import fs from 'fs'; import path from 'path';
const dir = process.argv[2];
const B = await chromium.launch({ args:['--headless=new','--no-sandbox'], headless:false, executablePath:'/opt/pw-browsers/chromium' });
const pg = await B.newPage();
await pg.addInitScript(() => { window.OCR_CFG = { script: 'http://127.0.0.1:8899/ocr/tesseract.min.js', workerPath: 'http://127.0.0.1:8899/ocr/worker.min.js', corePath: 'http://127.0.0.1:8899/ocr/core/', langPath: 'http://127.0.0.1:8899/ocr/lang' }; });
const errs=[]; pg.on('pageerror', e=>errs.push('PAGEERROR: '+e.message));
pg.on('console', m=>{ if(m.type()==='error' && !/404|favicon/.test(m.text())) errs.push('CONSOLE: '+m.text()); });
await pg.goto('http://127.0.0.1:8899/index.html'); await pg.waitForSelector('#run');
const all=[]; (function walk(d){for(const f of fs.readdirSync(d,{withFileTypes:true})){const p=path.join(d,f.name); if(f.isDirectory()) walk(p); else if(!f.name.startsWith('.')) all.push(p);} })(dir);
const t0 = Date.now();
await pg.setInputFiles('#dropall input', all);
try { await pg.waitForFunction(() => /Hoàn tất|Chưa có hóa đơn|Lỗi:/.test(document.querySelector('#log').innerText), null, { timeout: 400000 }); } catch (e) { console.log('TIMEOUT'); }
console.log('thời gian', Math.round((Date.now()-t0)/1000), 's');
console.log(await pg.textContent('#count'));
console.log('detect:', await pg.textContent('#detect'));
const res = await pg.evaluate(()=>window.__RESULT__);
for (const a of (res||[])) {
  const st={}; a.lines.forEach(l=>st[l.status]=(st[l.status]||0)+1);
  console.log('##', a.invNo||'(chưa có số HĐ)', a.invDate, '| inb', a.inb?'có':'chưa', '| dòng', a.lines.length, '| tổng', a.invTotal, '/', a.inbTotal, '|', JSON.stringify(st));
  for (const l of a.lines) console.log('     ', l.status, '|', l.code, '|', l.sapPo, '| HĐ', l.invQty, 'inb', l.inbQty, '|', (l.note||'').slice(0,140));
}
console.log(await pg.evaluate(()=>document.querySelector('#log').innerText.split('\n').filter(x=>/OCR/.test(x)).join('\n')));
if (errs.length) console.log('!!! LỖI:\n'+errs.join('\n'));
await B.close();
