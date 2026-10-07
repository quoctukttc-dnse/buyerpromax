import { chromium } from 'playwright';
import fs from 'fs'; import path from 'path';
const B=await chromium.launch({args:['--headless=new','--no-sandbox'],headless:false,executablePath:'/opt/pw-browsers/chromium'});
const errs=[];
async function run(label, dirs, clickPo){
  const pg=await B.newPage(); pg.on('pageerror',e=>errs.push(label+': '+e.message));
  await pg.goto('http://127.0.0.1:8899/index.html'); await pg.waitForSelector('#run');
  const all=[]; for(const d of dirs){ (function walk(x){for(const f of fs.readdirSync(x,{withFileTypes:true})){const p=path.join(x,f.name); if(f.isDirectory()) walk(p); else all.push(p);}})(d); }
  await pg.setInputFiles('#dropall input', all);
  await pg.waitForTimeout(dirs.length>1?45000:12000);
  console.log('###', label, '|', await pg.textContent('#count'));
  if (clickPo){ await pg.click('#runpo'); await pg.waitForTimeout(4000);
    const rows=await pg.$$eval('#report table tbody tr', rs=>rs.map(r=>[...r.querySelectorAll('td')].map(t=>t.innerText).join(' | ')));
    rows.forEach(r=>console.log('   PO:', r));
    console.log('   file:', await pg.$$eval('#dls a', as=>as.map(a=>a.download).join(', ')));
  } else {
    const res=await pg.evaluate(()=>window.__RESULT__);
    for(const a of (res||[])){ const bad=a.lines.filter(l=>l.status.indexOf('KHỚP')!==0&&l.status!=='CHƯA CÓ INBOUND');
      console.log('  ', a.invNo, '| inb', a.inb?'có':'chưa', '| dòng', a.lines.length, '| lệch', bad.length);
      for(const l of bad) console.log('       ', l.status, l.code, l.sapPo); }
    console.log('   file:', await pg.$$eval('#dls a', as=>as.map(a=>a.download).join(', ')));
  }
  await pg.close();
}
await run('CHỈ CHỨNG TỪ VẢI, CHƯA CÓ INBOUND → xuất danh sách PO', ['/tmp/vainoinb'], true);
await run('TRỘN trimming + vải', ['/tmp/atus','/tmp/vai'], false);
if(errs.length) console.log('!!! LỖI:', errs.join('\n'));
await B.close();
