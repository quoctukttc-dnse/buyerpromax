import { chromium } from 'playwright';
import fs from 'fs'; import path from 'path';
const dir = process.argv[2] || '/tmp/t3';
const B = await chromium.launch({ args:['--headless=new','--no-sandbox'], headless:false, executablePath:'/opt/pw-browsers/chromium' });
const pg = await B.newPage();
const errs=[]; pg.on('pageerror', e=>errs.push('PAGEERROR: '+e.message));
pg.on('console', m=>{ if(m.type()==='error' && !/404|favicon/.test(m.text())) errs.push('CONSOLE: '+m.text()); });
await pg.goto('http://127.0.0.1:8899/index.html');
await pg.waitForSelector('#run');
const all=[]; (function walk(d){for(const f of fs.readdirSync(d,{withFileTypes:true})){const p=path.join(d,f.name); if(f.isDirectory()) walk(p); else all.push(p);} })(dir);
await pg.setInputFiles('#dropall input', all);
await pg.waitForTimeout(45000);
console.log(await pg.textContent('#count'));
console.log('detect:', await pg.textContent('#detect'));
const rows = await pg.$$eval('table.files tbody tr', rs => rs.map(r => [...r.querySelectorAll('td')].map(t=>t.innerText.replace(/\n/g,' | ')).join('  ||  ')));
rows.forEach(r=>console.log('  ', r));
const res = await pg.evaluate(()=>window.__RESULT__);
for (const a of (res||[])) {
  const bad=a.lines.filter(l=>l.status.indexOf('KHỚP')!==0 && l.status!=='CHƯA CÓ INBOUND');
  const st={}; a.lines.forEach(l=>st[l.status]=(st[l.status]||0)+1);
  console.log('##', a.invNo||'(chưa có số HĐ)', a.invDate, '| inb', a.inb?'có':'chưa', '| dòng', a.lines.length, '| tổng', a.invTotal, '/', a.inbTotal, '|', JSON.stringify(st));
  for(const l of bad) console.log('     ', l.status, '|', l.code, '|', l.sapPo, '| HĐ', l.invQty, 'inb', l.inbQty, 'pkl', l.pklQty);
}
console.log('FILE:', await pg.$$eval('#dls a', as=>as.map(a=>a.download).join(', ')));
if (errs.length) console.log('!!! LỖI:\n'+errs.join('\n'));
await B.close();
