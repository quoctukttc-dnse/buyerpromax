import fs from 'fs';
import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const ExcelJS = require('exceljs');
const pdfjs = require('/home/claude/build/node_modules/pdfjs-dist/legacy/build/pdf.js');
const src = fs.readFileSync('/home/claude/build/fab.js','utf8')+'\n'+fs.readFileSync('/home/claude/build/gen.js','utf8')+'\nreturn {genLinesOf,pdfHeader,qtyOnLine,genPos,pageRole,pdfItemXY};';
const lib = new Function('ExcelJS','module',src)(ExcelJS,{exports:{}});
const data = new Uint8Array(fs.readFileSync(process.argv[2]));
const pdf = await pdfjs.getDocument({data,isEvalSupported:false,useSystemFonts:true,verbosity:0}).promise;
for (let p=1;p<=Math.min(pdf.numPages, +(process.argv[3]||1));p++){
  const page=await pdf.getPage(p); const tc=await page.getTextContent();
  const lines=lib.genLinesOf(tc.items.filter(i=>i.str&&i.str.trim()).map(i=>lib.pdfItemXY(i)));
  const hd=lib.pdfHeader(lines,0); console.log('PAGE',p,'role',lib.pageRole(lines),'header',JSON.stringify(hd));
  lines.forEach((L,i)=>{const q=lib.qtyOnLine(L,hd?hd.qcol:null,false); console.log(String(i).padStart(3), Math.round(L.y), (q?`Q=${q.q}${q.unit||''}${q.byCol?'(col)':''}`:'').padEnd(14), lib.genPos(L.text).join(','), '|', L.cells.map(c=>Math.round(c.x)+':'+c.s).join(' ¦ ').slice(0,170));});
}
