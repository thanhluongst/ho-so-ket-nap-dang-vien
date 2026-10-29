# -*- coding: utf-8 -*-
"""Tạo mẫu có trường điền (thư mục mau/) từ các mẫu gốc .docx (thư mục goc/).

Chỉ cần chạy lại khi thay đổi mẫu gốc. Mỗi chỗ trống "……" trong một đoạn
được thay lần lượt bằng một trường {{ ... }}; chữ và định dạng còn lại giữ nguyên.
"""
import copy
import re
from pathlib import Path

import docx
from docx.oxml.ns import qn

HERE = Path(__file__).parent
DOT = re.compile(r'[.…]*(?:…|\.\.)[.…]*')
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

D = lambda p: [f'{p}_d', f'{p}_m', f'{p}_y']   # ngày, tháng, năm
F = lambda name: f'{{{{ {name} }}}}'

# Mỗi mục: (regex nhận diện đoạn, hành động)
#   list  -> các trường điền lần lượt vào từng chỗ trống của đoạn
#   ('next', list) -> điền vào đoạn toàn dấu chấm ngay sau đoạn đó
#   ('append', var) -> thêm trường vào cuối đoạn
#   ('del',)       -> xóa đoạn toàn dấu chấm ngay sau đoạn đó
#   ('block', var) -> thay khối dòng chấm phía sau bằng vòng lặp các đoạn văn của var
#   ('sign', var)  -> thêm họ tên người ký dưới dòng "(Ký, ghi rõ ...)"
SPEC = {
    'mau1': [
        (r'Kính gửi: Chi ủy', ['chi bộ ' + F('chi_bo')]),
        (r'^Đảng ủy', ['Trường ' + F('truong')]),
        (r'^Tôi là', ['ho_ten', *D('ns')]),
        (r'^Nơi sinh', ['noi_sinh']),
        (r'^Quê quán', ['que_quan']),
        (r'^Dân tộc', ['dan_toc', 'ton_giao']),
        (r'^Trình độ học vấn', ['hoc_van']),
        (r'^Nơi cư trú', ['cu_tru']),
        (r'^Nghề nghiệp', ['nghe_nghiep']),
        (r'^Đơn vị công tác', ['don_vi']),
        (r'^Chức vụ chính quyền', ['chuc_vu']),
        (r'^Chức vụ chính quyền', ('del',)),
        (r'^Vào Đoàn', [*D('doan'), 'doan_noi']),
        (r'^Được xét là cảm tình', D('ctd')),
        (r'^Được xét là cảm tình', ('append', 'ctd_chi_bo')),
        (r'^Được xét là cảm tình', ('del',)),
        (r'tôi đã nhận thức được', ('block', 'nhan_thuc')),
        (r'ngày.*tháng.*năm', ['dia_danh', *D('don')]),
        (r'Ký, ghi rõ', ('sign', 'ho_ten')),
    ],
    'mau3': [
        (r'Kính gửi: Chi bộ', ['chi_bo']),
        (r'^Tôi là', ['gt_ten', *D('gt_ns')]),
        (r'^Vào Đảng ngày', [*D('gt_vd'), *D('gt_ct')]),
        (r'^Chức vụ trong Đảng', ['gt_cv_dang']),
        (r'^Chức vụ\s+chính quyền', ['gt_cv_cq']),
        (r'^Quê quán', ['gt_que']),
        (r'^Đang sinh hoạt', ['gt_chi_bo']),
        (r'phân công giáo dục', [*D('pc'), 'ho_ten']),
        (r'^Những đặc điểm chính', ['']),
        (r'^Những đặc điểm chính', ('block', 'ly_lich')),
        (r'^\(Về chấp hành', ('block', 'uu_kd')),
        (r'^Đối chiếu', ['ho_ten']),
        (r'ngày.*tháng.*năm', ['dia_danh', *D('gt_ky')]),
        (r'Ký, ghi rõ', ('sign', 'gt_ten')),
    ],
    'mau5': [
        (r'^ĐẢNG BỘ TRƯỜNG', ['TRUONG']),
        (r'^CHI BỘ', ['CHI_BO']),
        (r'ngày.*tháng.*năm', ['dia_danh', *D('m5')]),
        (r'^Căn cứ ý kiến', ['ho_ten']),
        (r'^Tên của các tổ chức', [F('ten_doan_the') + ', tổng số có: ' + F('so_dt') + ' đồng chí.']),
        (r'^Tên của các tổ chức', ('del',)),
        (r'^[.…].*tổng số có:', ('drop',)),
        (r'^Tên chi uỷ nơi cư trú', ['ten_ct', 'so_ct']),
        (r'^\(Về chấp hành', ('block', 'uu_kd')),
        (r'^Số đồng chí đại diện', ['ho_ten', 'dt_tan', 'dt_tong', 'dt_pt', 'dt_khong', 'dt_kpt', 'dt_lydo']),
        (r'^Số đồng chí trong chi ủy', ['ho_ten', 'ct_tan', 'ct_tong', 'ct_pt', 'ct_khong', 'ct_kpt', 'ct_lydo']),
        (r'^Số đồng chí trong chi ủy', ('del',)),
        (r'Ký, ghi rõ', ('sign', 'bt_chi_bo')),
    ],
    'mau6': [
        (r'^ĐẢNG BỘ TRƯỜNG', ['TRUONG']),
        (r'^CHI BỘ', ['CHI_BO']),
        (r'ngày.*tháng.*năm', ['dia_danh', *D('cb')]),
        (r'^Ngày .*Chi bộ', [*D('cb'), 'chi_bo', 'ho_ten']),
        (r'^Tổng số đảng viên', ['tong_dv', 'dv_ct', 'dv_db']),
        (r'^Có mặt', ['cm_tong', 'cm_ct', 'cm_db']),
        (r'^Vắng mặt', ['vm_tong', 'vm_ct', 'vm_db']),
        (r'^Lý do vắng', ['cb_lydo_vang']),
        (r'^Chủ trì', ['cb_chu_tri']),
        (r'^Chức vụ', ['cb_chuc_vu']),
        (r'^Thư ký', ['cb_thu_ky']),
        (r'^Sau khi nghe', ['ho_ten']),
        (r'^Về lý lịch', ('block', 'ly_lich')),
        (r'^\(Về lịch sử', ('block', 'uu_kd')),
        (r'^Đối chiếu', ['ho_ten', 'cb_tan', 'cb_pt', 'cb_khong', 'cb_kpt', 'cb_lydo_khong']),
        (r'^Đối chiếu', ('del',)),
        (r'^Chi bộ đề nghị', ['ho_ten']),
        (r'Ký, ghi rõ', ('sign', 'bt_chi_bo')),
    ],
    'mau8': [
        (r'^ĐẢNG ỦY', ['DANG_UY']),
        (r'ngày.*tháng.*năm', ['dia_danh', *D('du')]),
        (r'^Ngày .*Ban Chấp hành', [*D('du'), 'dang_bo', 'ho_ten']),
        (r'^Tổng số ủy viên', ['uv_tong', 'uv_cm', 'uv_vm']),
        (r'^Lý do vắng', ['du_lydo_vang']),
        (r'^Chủ trì', ['du_chu_tri']),
        (r'^Chức vụ', ['du_chuc_vu']),
        (r'^Thư ký', ['du_thu_ky']),
        (r'^Sau khi xem xét', ['so_nq_cb', *D('cb'), 'chi_bo', 'ho_ten']),
        (r'^Về lý lịch', ('block', 'ly_lich')),
        (r'^\(Về lịch sử', ('block', 'uu_kd')),
        (r'^Đối chiếu', ['ho_ten']),
        (r'^- Đồng ý', ['du_dy', 'du_pt']),
        (r'^- Không đồng ý', ['du_kdy', 'du_kpt', 'du_lydo_khong']),
        (r'^Đề nghị Ban Thường vụ', ['ho_ten']),
        (r'Lưu: Đảng bộ Trường', ['truong']),
        (r'Ký, đóng dấu', ('sign', 'bt_dang_uy')),
    ],
}
# Sửa lỗi chữ nhỏ trong mẫu gốc
FIXES = {'mau8': [('Của chi bộ', 'của chi bộ'), ('đồng chí (đạt {{ du_kpt', 'đồng chí (chiếm {{ du_kpt')]}
# Ô "Số      -NQ/.." trên tiêu đề nghị quyết
SO_NQ = {'mau6': ('-NQ/CB', 'so_nq_cb'), 'mau8': ('-NQ/ĐU', 'so_nq_du')}


def ptext(p):
    return ''.join(t.text or '' for t in p.iter(qn('w:t')))


def is_dotted(p):
    t = ptext(p)
    return bool(t.strip()) and re.fullmatch(r'[\s.…]+', t) is not None


def set_text(p, new):
    """Ghi toàn bộ chữ của đoạn vào w:t đầu tiên (giữ định dạng của nó), xóa các w:t khác."""
    ts = list(p.iter(qn('w:t')))
    ts[0].text = new
    ts[0].set(XML_SPACE, 'preserve')
    for t in ts[1:]:
        t.text = ''


def fill(p, fields):
    """Thay lần lượt từng chỗ trống bằng trường; chữ quanh chỗ trống giữ nguyên run của nó."""
    ts = list(p.iter(qn('w:t')))
    full = ''.join(t.text or '' for t in ts)
    spans = [m.span() for m in DOT.finditer(full)]
    assert len(spans) == len(fields), f'{full!r}: {len(spans)} chỗ trống, {len(fields)} trường'
    owner = []  # ký tự thứ i thuộc w:t nào
    for k, t in enumerate(ts):
        owner += [k] * len(t.text or '')
    chars = [list(t.text or '') for t in ts]
    offs = [sum(len(t.text or '') for t in ts[:k]) for k in range(len(ts))]
    for (a, b), f in reversed(list(zip(spans, fields))):
        val = f if (not f or '{{' in f) else F(f)
        if a > 0 and (full[a - 1].isalnum() or full[a - 1] in ':'):
            val = ' ' + val
        if b < len(full) and full[b].isalnum():
            val += ' '
        for i in range(a, b):
            chars[owner[i]][i - offs[owner[i]]] = ''
        k = owner[a]
        chars[k][a - offs[k]] = val
    for t, cs in zip(ts, chars):
        t.text = ''.join(cs)
        t.set(XML_SPACE, 'preserve')
        if '{{' in t.text:
            plain(t.getparent())
    # dọn các dấu "…. . đồng chí", "% )" còn sót
    t = ptext(p)
    t2 = re.sub(r'\}\} \. ', '}} ', t).replace('% )', '%)')
    if t2 != t:
        _rewrite_keep_runs(p, t, t2)


def _rewrite_keep_runs(p, old, new):
    for t in p.iter(qn('w:t')):
        s = t.text or ''
        s2 = re.sub(r'\}\} \. ', '}} ', s).replace('% )', '%)')
        if s2 != s:
            t.text = s2
    if ptext(p) != new:  # dấu nằm vắt qua hai run → gộp
        set_text(p, new)


def plain(run):
    """Bỏ màu chữ / tô sáng của run (dòng chấm trong mẫu gốc in màu xám)."""
    rpr = run.find(qn('w:rPr'))
    if rpr is not None:
        for tag in ('w:color', 'w:highlight', 'w:shd'):
            for e in rpr.findall(qn(tag)):
                rpr.remove(e)


def new_para(src_p, text, bold=False):
    p = copy.deepcopy(src_p)
    for r in list(p):
        if r.tag != qn('w:pPr'):
            p.remove(r)
    r = copy.deepcopy(src_p.find('.//' + qn('w:r')))
    for c in list(r):
        if c.tag != qn('w:rPr'):
            r.remove(c)
    rpr = r.find(qn('w:rPr'))
    if rpr is not None:
        for tag in ('w:b', 'w:i', 'w:bCs', 'w:iCs'):
            for e in rpr.findall(qn(tag)):
                rpr.remove(e)
        for tag in ('w:u', 'w:color', 'w:highlight', 'w:shd'):
            for e in rpr.findall(qn(tag)):
                rpr.remove(e)
        if bold:
            rpr.insert(0, rpr.makeelement(qn('w:b'), {}))
    t = r.makeelement(qn('w:t'), {})
    t.text = text
    t.set(XML_SPACE, 'preserve')
    r.append(t)
    p.append(r)
    return p


def next_dotted(p):
    n = p.getnext()
    while n is not None and n.tag == qn('w:p') and not ptext(n).strip():
        n = n.getnext()
    assert n is not None and is_dotted(n), f'không thấy dòng chấm sau {ptext(p)!r}'
    return n


def build(name):
    d = docx.Document(HERE / 'goc' / f'{name}.docx')
    paras = list(d.element.body.iter(qn('w:p')))
    body_ref = None  # đoạn văn thân bài làm mẫu thụt lề cho đoạn trong vòng lặp
    for p in paras:
        if ptext(p).startswith(('Tôi là', 'Căn cứ', 'Tổng số')):
            body_ref = p
            break

    for pattern, action in SPEC[name]:
        for p in paras:
            if is_dotted(p) or p.getparent() is None:
                continue
            t = ptext(p).strip()
            if isinstance(action, list) and not DOT.search(t):
                continue  # đoạn đã điền rồi
            if re.search(pattern, t):
                break
        else:
            raise SystemExit(f'{name}: không tìm thấy đoạn {pattern!r}')

        if isinstance(action, list):
            fill(p, action)
        elif action[0] == 'next':
            fill(next_dotted(p), action[1])
        elif action[0] == 'append':
            t = list(p.iter(qn('w:t')))[-1]
            t.text = (t.text or '').rstrip() + ' ' + F(action[1])
            t.set(XML_SPACE, 'preserve')
            plain(t.getparent())
        elif action[0] == 'del':
            n = next_dotted(p)
            n.getparent().remove(n)
        elif action[0] == 'drop':
            p.getparent().remove(p)
        elif action[0] == 'block':
            var = action[1]
            first = next_dotted(p)
            block = [first]
            n = first.getnext()
            while n is not None and n.tag == qn('w:p') and (is_dotted(n) or not ptext(n).strip()):
                block.append(n)
                n = n.getnext()
            while not ptext(block[-1]).strip():  # để lại dòng trống cuối khối
                block.pop()
            content = new_para(p, F('dong'))
            ppr_ref = (p if p.getparent().tag == qn('w:body') else body_ref).find(qn('w:pPr'))
            old = content.find(qn('w:pPr'))
            if old is not None:
                content.remove(old)
            if ppr_ref is not None:
                content.insert(0, copy.deepcopy(ppr_ref))
            first.addprevious(new_para(first, f'{{%p for dong in {var} %}}'))
            first.addprevious(content)
            first.addprevious(new_para(first, '{%p endfor %}'))
            for b in block:
                b.getparent().remove(b)
        elif action[0] == 'sign':
            anchor = p
            for _ in range(4):
                e = new_para(p, '')
                anchor.addnext(e)
                anchor = e
            anchor.addnext(new_para(p, F(action[1]), bold=True))

    if name in SO_NQ:
        suffix, var = SO_NQ[name]
        for p in paras:
            t = ptext(p)
            if re.fullmatch(r'Số\s+' + suffix, t):
                set_text(p, f'Số {F(var)}{suffix}')
                break
        else:
            raise SystemExit(f'{name}: không thấy ô số nghị quyết')

    for old, new in FIXES.get(name, []):
        for p in paras:
            if old in ptext(p):
                for t in p.iter(qn('w:t')):
                    if old in (t.text or ''):
                        t.text = t.text.replace(old, new)
                        break
                else:
                    set_text(p, ptext(p).replace(old, new))

    out = HERE / 'mau' / f'{name}.docx'
    out.parent.mkdir(exist_ok=True)
    d.save(out)
    print('Đã tạo', out)

    # Bản tô đỏ chữ được điền (mau_do/) để dễ nhận biết khi kiểm tra
    to_do(d)
    out = HERE / 'mau_do' / f'{name}.docx'
    out.parent.mkdir(exist_ok=True)
    d.save(out)
    print('Đã tạo', out)


MAU_DO = 'FF0000'


def set_red(run):
    rpr = run.find(qn('w:rPr'))
    if rpr is None:
        rpr = run.makeelement(qn('w:rPr'), {})
        run.insert(0, rpr)
    for e in rpr.findall(qn('w:color')):
        rpr.remove(e)
    color = rpr.makeelement(qn('w:color'), {qn('w:val'): MAU_DO})
    # w:color phải đứng trước w:sz, w:u... theo thứ tự của Word
    after = [qn(t) for t in ('w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps', 'w:smallCaps',
                              'w:strike', 'w:dstrike', 'w:outline', 'w:shadow', 'w:emboss', 'w:imprint',
                              'w:noProof', 'w:snapToGrid', 'w:vanish', 'w:webHidden')]
    idx = 0
    for k, c in enumerate(rpr):
        if c.tag in after:
            idx = k + 1
    rpr.insert(idx, color)


def to_do(d):
    """Tách từng trường {{ ... }} ra run riêng và tô đỏ, chữ cố định của mẫu giữ nguyên màu."""
    for t in list(d.element.body.iter(qn('w:t'))):
        text = t.text or ''
        if '{{' not in text:
            continue
        run = t.getparent()
        others = [c for c in run if c.tag not in (qn('w:rPr'), qn('w:t'))]
        if others or len(run.findall(qn('w:t'))) > 1:
            set_red(run)
            continue
        for piece in re.split(r'(\{\{.*?\}\})', text):
            if not piece:
                continue
            r = copy.deepcopy(run)
            r.find(qn('w:t')).text = piece
            if piece.startswith('{{'):
                set_red(r)
            run.addprevious(r)
        run.getparent().remove(run)


if __name__ == '__main__':
    for n in SPEC:
        build(n)
