# -*- coding: utf-8 -*-
"""Chuyển các mẫu trong mau/ sang cú pháp docxtemplater cho bản online (web/mau_data.js)."""
import base64
import io
import json
import re
import zipfile
from pathlib import Path

HERE = Path(__file__).parent
OUT = HERE.parent / 'web' / 'mau_data.js'


def convert(xml):
    loops = re.findall(r'\{%p for dong in (\w+) %\}', xml)
    for var in loops:
        xml = xml.replace(f'{{%p for dong in {var} %}}', f'{{{{#{var}}}}}', 1)
        xml = xml.replace('{{ dong }}', '{{.}}', 1)
        xml = xml.replace('{%p endfor %}', f'{{{{/{var}}}}}', 1)
    xml = re.sub(r'\{\{ (\w+) \}\}', r'{{\1}}', xml)
    assert '{%' not in xml and '{{ ' not in xml
    return xml


def pack(path):
    src = zipfile.ZipFile(path)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            content = src.read(item.filename)
            if item.filename == 'word/document.xml':
                content = convert(content.decode('utf-8')).encode('utf-8')
            dst.writestr(item, content)
    return base64.b64encode(buf.getvalue()).decode()


data = {}  # mau1, ... và bản tô đỏ chữ được điền mau1_do, ...
for folder, suffix in (('mau', ''), ('mau_do', '_do')):
    for f in sorted((HERE / folder).glob('mau*.docx')):
        data[f.stem + suffix] = pack(f)

OUT.parent.mkdir(exist_ok=True)
nhan_xet = json.loads((HERE / 'mau_nhan_xet.json').read_text(encoding='utf-8'))
OUT.write_text('window.MAU_DATA = ' + json.dumps(data) + ';\n'
               + 'window.MAU_NX = ' + json.dumps(nhan_xet, ensure_ascii=False) + ';\n'
               + 'window.MAU_EXCEL = "' + base64.b64encode((HERE.parent / 'Mau nhap thong tin.xlsx').read_bytes()).decode()
               + '";\n', encoding='utf-8')
print('Đã tạo', OUT, OUT.stat().st_size // 1024, 'KB')

# Bản cho GitHub Pages (thư mục docs/): thêm khung HTML mà claude.ai tự bọc khi đăng
DOCS = HERE.parent / 'docs'
DOCS.mkdir(exist_ok=True)
page = (OUT.parent / 'index.html').read_text(encoding='utf-8').replace('<meta charset="utf-8">\n', '', 1)
(DOCS / 'index.html').write_text(
    '<!doctype html>\n<html lang="vi">\n<head>\n<meta charset="utf-8">\n'
    '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">\n'
    '<style>:root{padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}'
    'body{margin:0}img{max-width:100%}[hidden]{display:none!important}</style>\n</head>\n<body>\n'
    + page + '\n</body>\n</html>\n', encoding='utf-8')
(DOCS / 'mau_data.js').write_bytes(OUT.read_bytes())
print('Đã tạo', DOCS / 'index.html')
