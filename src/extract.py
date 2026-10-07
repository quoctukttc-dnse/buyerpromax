#!/usr/bin/env python3
"""Dựng lại cây nguồn từ site/index.html đã build (dùng khi mất máy dựng).
   Chạy: python3 extract.py ../index.html   → tạo vendor/, app-all.js, fab.js, index-src.html"""
import base64, pathlib, re, sys

src = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else 'index.html')
HERE = pathlib.Path(__file__).parent
h = src.read_text(encoding='utf8')


def grab(i):
    m = re.search(r'<script type="text/plain" id="%s">(.*?)</script>' % i, h, re.S)
    return m.group(1)


(HERE / 'vendor').mkdir(exist_ok=True)
for i, fn in [('lib-exceljs', 'vendor/exceljs.min.js'), ('lib-pdfjs', 'vendor/pdf.min.js'),
              ('lib-pdfworker', 'vendor/pdf.worker.min.js'), ('lib-sheetjs', 'vendor/xlsx.core.min.js')]:
    (HERE / fn).write_bytes(base64.b64decode(grab(i).strip()))
app = grab('app-src')
parts = re.split(r'/\* ==== (\S+) ==== \*/\n', app)
if len(parts) > 1:
    for name, body in zip(parts[1::2], parts[2::2]):
        (HERE / name).write_text(body.rstrip('\n') + '\n', encoding='utf8')
else:
    (HERE / 'app-all.js').write_text(app, encoding='utf8')
shell = h.replace(app, '@@APP@@')
for i in ['lib-exceljs', 'lib-pdfjs', 'lib-pdfworker', 'lib-sheetjs']:
    shell = shell.replace(grab(i), '@@%s@@' % i)
(HERE / 'index-src.html').write_text(shell, encoding='utf8')
print('OK — đã dựng lại nguồn. Chạy: python3 build.py')
