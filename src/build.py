#!/usr/bin/env python3
"""Dựng site/index.html: nhúng vendor (base64) + app (plain) vào index-src.html."""
import base64, os, pathlib

HERE = pathlib.Path(__file__).parent
shell = (HERE / 'index-src.html').read_text(encoding='utf8')

VEND = {
    'lib-exceljs': 'vendor/exceljs.min.js',
    'lib-pdfjs': 'vendor/pdf.min.js',
    'lib-pdfworker': 'vendor/pdf.worker.min.js',
    'lib-sheetjs': 'vendor/xlsx.core.min.js',
}
for key, fn in VEND.items():
    b64 = base64.b64encode((HERE / fn).read_bytes()).decode()
    shell = shell.replace('@@%s@@' % key, b64)

parts = []
for fn in ['fab.js', 'gen.js', 'app-all.js']:
    p = HERE / fn
    if p.exists():
        parts.append('/* ==== %s ==== */\n' % fn + p.read_text(encoding='utf8'))
app = '\n'.join(parts)
shell = shell.replace('@@APP@@', app)

out = HERE / 'site' / 'index.html'
out.parent.mkdir(exist_ok=True)
out.write_text(shell, encoding='utf8')
print('OK', out, len(shell), 'ký tự')

# ---- order.html: trang điền order form supplier từ PO SCAVI (không dùng thư viện ngoài) ----
osrc = HERE / 'order-src.html'
if osrc.exists() and (HERE / 'order.js').exists():
    oshell = osrc.read_text(encoding='utf8')
    ver = re_ver = None
    import re
    m = re.search(r'class="ver">(v[\d.]+ · [\d.]+)', shell)
    oshell = oshell.replace('@@ORDER_VER@@', (m.group(1) if m else '') + ' · điền order form')
    oshell = oshell.replace('@@ORDER_APP@@', (HERE / 'order.js').read_text(encoding='utf8'))
    (HERE / 'site' / 'order.html').write_text(oshell, encoding='utf8')
    print('OK', HERE / 'site' / 'order.html', len(oshell), 'ký tự')
