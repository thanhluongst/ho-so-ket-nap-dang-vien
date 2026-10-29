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


data = {}
for f in sorted((HERE / 'mau').glob('mau*.docx')):
    src = zipfile.ZipFile(f)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            content = src.read(item.filename)
            if item.filename == 'word/document.xml':
                content = convert(content.decode('utf-8')).encode('utf-8')
            dst.writestr(item, content)
    data[f.stem] = base64.b64encode(buf.getvalue()).decode()

OUT.parent.mkdir(exist_ok=True)
OUT.write_text('window.MAU_DATA = ' + json.dumps(data) + ';\n', encoding='utf-8')
print('Đã tạo', OUT, OUT.stat().st_size // 1024, 'KB')
