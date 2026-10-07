import { chromium } from 'playwright';
import fs from 'fs';
const B = await chromium.launch({ args:['--headless=new','--no-sandbox'], headless:false, executablePath:'/opt/pw-browsers/chromium' });
const ctx = await B.newContext({ acceptDownloads: true });
const pg = await ctx.newPage();
await pg.goto('http://127.0.0.1:8899/index.html');
await pg.waitForSelector('#run');
const dir=process.argv[2]||'/tmp/t3/';
await pg.setInputFiles('#dropall input', fs.readdirSync(dir).map(f=>dir+f));
await pg.waitForTimeout(45000);
fs.rmSync('/tmp/out',{recursive:true,force:true}); fs.mkdirSync('/tmp/out',{recursive:true});
const links = await pg.$$('#dls a');
for (const a of links) {
  const [dl] = await Promise.all([pg.waitForEvent('download'), a.click()]);
  await dl.saveAs('/tmp/out/' + dl.suggestedFilename());
}
console.log(fs.readdirSync('/tmp/out').join('\n'));
await B.close();
