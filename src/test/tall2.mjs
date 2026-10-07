import { chromium } from 'playwright';
import fs from 'fs'; import path from 'path';
const dir = process.argv[2];
const B = await chromium.launch({ args:['--headless=new','--no-sandbox'], headless:false, executablePath:'/opt/pw-browsers/chromium' });
const pg = await B.newPage();
const errs=[]; pg.on('pageerror', e=>errs.push('PAGEERROR: '+e.message));
pg.on('console', m=>{ if(m.type()==='error' && !/404|favicon/.test(m.text())) errs.push('CONSOLE: '+m.text()); });
await pg.goto('http://127.0.0.1:8899/index.html'); await pg.waitForSelector('#run');
const all=[]; (function walk(d){for(const f of fs.readdirSync(d,{withFileTypes:true})){const p=path.join(d,f.name); if(f.isDirectory()) walk(p); else if(!f.name.startsWith('.')) all.push(p);} })(dir);
await pg.setInputFiles('#dropall input', all);
try { await pg.waitForFunction(() => /Hoàn tất|Chưa có hóa đơn|Lỗi:/.test(document.querySelector('#log').innerText), null, { timeout: 90000 }); } catch (e) {}
await pg.waitForTimeout(500);
console.log(await pg.textContent('#count'));
console.log('detect:', await pg.textContent('#detect'));
const res = await pg.evaluate(()=>window.__RESULT__);
for (const a of (res||[])) {
  const st={}; a.lines.forEach(l=>st[l.status]=(st[l.status]||0)+1);
  console.log('##', a.invNo||'(chưa có số HĐ)', a.invDate, '| inb', a.inb?'có':'chưa', '| dòng', a.lines.length, '| tổng', a.invTotal, '/', a.inbTotal, '|', JSON.stringify(st));
  if (process.argv.includes('-v')) for (const l of a.lines) console.log('     ', l.status, '|', l.code, '|', l.sapPo, '| HĐ', l.invQty, 'inb', l.inbQty, 'pkl', l.pklQty);
}
console.log('FILE:', await pg.$$eval('#dls a', as=>as.map(a=>a.download).join(', ')));
if (errs.length) console.log('!!! LỖI:\n'+errs.join('\n'));
await B.close();
