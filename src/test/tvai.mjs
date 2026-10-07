import { chromium } from 'playwright';
import fs from 'fs';
const B = await chromium.launch({ args: ['--headless=new', '--no-sandbox'], headless: false,
  executablePath: '/opt/pw-browsers/chromium' });
const pg = await B.newPage();
const errs = [];
pg.on('console', m => { if (m.type()==='error') errs.push(m.text()); });
pg.on('pageerror', e => errs.push('PAGEERROR: ' + e.message));
await pg.goto('http://127.0.0.1:8899/index.html');
await pg.waitForFunction(() => typeof window.__ready__ !== 'undefined' || document.querySelector('#run'), null, {timeout:15000});
const dir = '/tmp/vai/';
const files = fs.readdirSync(dir).map(f => dir + f);
await pg.setInputFiles('#dropall input', files);
await pg.waitForTimeout(9000);
console.log('--- bảng file ---');
console.log(await pg.textContent('#count'));
console.log(await pg.textContent('#detect'));
const rows = await pg.$$eval('table.files tbody tr', rs => rs.map(r => [...r.querySelectorAll('td')].map(t => t.innerText.replace(/\n/g,' | ')).join('  ||  ')));
rows.forEach(r => console.log(' ', r));
console.log('--- log ---');
console.log(await pg.textContent('#log'));
const res = await pg.evaluate(() => window.__RESULT__);
console.log('--- kết quả ---');
for (const a of (res||[])) {
  console.log(a.invNo, a.invDate, 'pkl', a.pkl, 'inb', a.inb, 'saiGT', a.valueBad, 'tổng', a.invTotal, '/', a.inbTotal);
  for (const l of a.lines) console.log('   ', l.status.padEnd(22), l.sapPo, '|', l.code, '| HĐ', l.invQty, '| inb', l.inbQty, '| pkl', l.pklQty);
}
console.log('--- file xuất ---');
console.log(await pg.$$eval('#dls a', as => as.map(a => a.download).join('\n')));
if (errs.length) console.log('!!! LỖI CONSOLE:\n' + errs.join('\n'));
await B.close();
