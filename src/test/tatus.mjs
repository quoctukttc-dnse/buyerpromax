import { chromium } from 'playwright';
import fs from 'fs';
import path from 'path';
const B = await chromium.launch({ args:['--headless=new','--no-sandbox'], headless:false, executablePath:'/opt/pw-browsers/chromium' });
const pg = await B.newPage();
const errs=[]; pg.on('pageerror', e=>errs.push('PAGEERROR: '+e.message));
await pg.goto('http://127.0.0.1:8899/index.html');
await pg.waitForSelector('#run');
const all=[];
(function walk(d){ for(const f of fs.readdirSync(d,{withFileTypes:true})){ const p=path.join(d,f.name); if(f.isDirectory()) walk(p); else all.push(p);} })('/tmp/atus');
await pg.setInputFiles('#dropall input', all);
await pg.waitForTimeout(40000);
console.log(await pg.textContent('#count'));
const res = await pg.evaluate(()=>window.__RESULT__);
let nBad=0, nLine=0;
for (const a of (res||[])) {
  const bad=a.lines.filter(l=>l.status.indexOf('KHỚP')!==0 && l.status!=='CHƯA CÓ INBOUND');
  nLine+=a.lines.length; nBad+=bad.length;
  console.log(a.invNo, a.invDate, '| inb', a.inb?'có':'chưa', '| dòng', a.lines.length, '| lệch', bad.length, '| saiGT', a.valueBad, '| tổng', a.invTotal, '/', a.inbTotal);
  for (const l of bad) console.log('    ', l.status, l.code, l.sapPo, 'HĐ', l.invQty, 'inb', l.inbQty, 'pkl', l.pklQty);
}
console.log('TỔNG:', (res||[]).length, 'hóa đơn,', nLine, 'dòng,', nBad, 'dòng lệch');
if (errs.length) console.log('!!!', errs.join('\n'));
await B.close();
