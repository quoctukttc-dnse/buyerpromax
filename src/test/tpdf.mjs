import { chromium } from 'playwright';
const B = await chromium.launch({ args:['--headless=new','--no-sandbox'], headless:false, executablePath:'/opt/pw-browsers/chromium' });
const pg = await B.newPage();
await pg.goto('http://127.0.0.1:8899/index.html');
await pg.waitForSelector('#run');
await pg.setInputFiles('#dropall input', ['/tmp/t3/HD 614.pdf']);
await pg.waitForTimeout(6000);
const out = await pg.evaluate(async () => {
  const f = [...STATE.items].find(x=>/HD 614/.test(x.file.name));
  const c = await classify(f.file);
  return { kind: c.kind, lines: c.lines, inv: parseInvoice(c.lines) };
});
console.log('kind', out.kind);
out.lines.forEach((l,i)=>console.log(String(i).padStart(3), JSON.stringify(l)));
console.log('--- parse ---');
console.log(JSON.stringify({serial:out.inv.serial,no:out.inv.no,invNo:out.inv.invNo,invDate:out.inv.invDate,total:out.inv.total}, null, 1));
out.inv.items.forEach(it=>console.log('  code=',JSON.stringify(it.code),'po=',JSON.stringify(it.po),'qty',it.qty,'px',it.price,'amt',it.amount));
await B.close();
