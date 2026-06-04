# -*- coding: utf-8 -*-
"""
Công văn chuẩn theo Nghị định 30/2020/NĐ-CP và Phụ lục I
"""
from docx import Document
from docx.shared import Pt, Mm, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

doc = Document()
sec = doc.sections[0]
sec.page_width    = Mm(210); sec.page_height   = Mm(297)
sec.left_margin   = Mm(30);  sec.right_margin  = Mm(15)
sec.top_margin    = Mm(20);  sec.bottom_margin = Mm(20)

# ── Helpers ────────────────────────────────────────────────────
TNR = 'Times New Roman'

def set_run(run, size=14, bold=False, italic=False, underline=False, caps=False):
    f = run.font
    f.name = TNR; f.size = Pt(size)
    f.bold = bold; f.italic = italic; f.underline = underline
    if caps: f.all_caps = True
    run._element.rPr.rFonts.set(qn('w:eastAsia'), TNR)

def fmt_para(p, align=WD_ALIGN_PARAGRAPH.LEFT, sb=0, sa=0, ls=1.15, fi=None):
    pf = p.paragraph_format
    pf.alignment = align
    pf.space_before = Pt(sb); pf.space_after = Pt(sa)
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = ls
    if fi is not None: pf.first_line_indent = Mm(fi)

def add_run(p, text, size=14, bold=False, italic=False, underline=False, caps=False):
    r = p.add_run(text); set_run(r, size, bold, italic, underline, caps); return r

def new_para(text='', align=WD_ALIGN_PARAGRAPH.LEFT, size=14,
             bold=False, italic=False, underline=False, caps=False,
             sb=0, sa=0, ls=1.15, fi=None):
    p = doc.add_paragraph(); fmt_para(p, align, sb, sa, ls, fi)
    if text: add_run(p, text, size, bold, italic, underline, caps)
    return p

def no_border(tbl):
    tblPr = tbl._tbl.tblPr
    b = OxmlElement('w:tblBorders')
    for n in ['top','left','bottom','right','insideH','insideV']:
        e = OxmlElement(f'w:{n}'); e.set(qn('w:val'), 'none'); b.append(e)
    tblPr.append(b)

def add_bottom_border(paragraph, color='000000', size=6):
    """Thêm đường kẻ dưới đoạn văn (thay cho gạch chân)"""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), str(size))
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), color)
    pBdr.append(bottom); pPr.append(pBdr)

def cell_para(cell, idx=0):
    return cell.paragraphs[idx] if idx < len(cell.paragraphs) else cell.add_paragraph()

# ══════════════════════════════════════════════════════════════
# PHẦN 1 — QUỐC HIỆU & TÊN CƠ QUAN
# ══════════════════════════════════════════════════════════════
# Bảng 2 cột, không viền: cột trái = cơ quan, cột phải = quốc hiệu
hdr = doc.add_table(rows=2, cols=2)
hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
no_border(hdr)

COL_L = Mm(87)   # cột trái
COL_R = Mm(95)   # cột phải

for row in hdr.rows:
    row.cells[0].width = COL_L
    row.cells[1].width = COL_R

# Hàng 1: UBND | CỘNG HOÀ
c_ubnd = hdr.cell(0,0).paragraphs[0]
fmt_para(c_ubnd, WD_ALIGN_PARAGRAPH.CENTER, sb=0, sa=0, ls=1.0)
add_run(c_ubnd, 'UBND TỈNH QUẢNG TRỊ', size=12, bold=False, caps=True)

c_ch = hdr.cell(0,1).paragraphs[0]
fmt_para(c_ch, WD_ALIGN_PARAGRAPH.CENTER, sb=0, sa=0, ls=1.0)
add_run(c_ch, 'CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM', size=12, bold=True, caps=True)

# Hàng 2: SỞ XÂY DỰNG (gạch dưới) | Độc lập (gạch dưới)
c_sxd = hdr.cell(1,0).paragraphs[0]
fmt_para(c_sxd, WD_ALIGN_PARAGRAPH.CENTER, sb=0, sa=3, ls=1.0)
add_run(c_sxd, 'SỞ XÂY DỰNG', size=14, bold=True, caps=True)
add_bottom_border(c_sxd, size=6)

c_dl = hdr.cell(1,1).paragraphs[0]
fmt_para(c_dl, WD_ALIGN_PARAGRAPH.CENTER, sb=0, sa=3, ls=1.0)
add_run(c_dl, 'Độc lập - Tự do - Hạnh phúc', size=13, bold=True)
add_bottom_border(c_dl, size=6)

# ══════════════════════════════════════════════════════════════
# PHẦN 2 — SỐ KÝ HIỆU & ĐỊA DANH NGÀY THÁNG
# ══════════════════════════════════════════════════════════════
so_ngay = doc.add_table(rows=1, cols=2)
so_ngay.alignment = WD_TABLE_ALIGNMENT.CENTER
no_border(so_ngay)
so_ngay.cell(0,0).width = COL_L
so_ngay.cell(0,1).width = COL_R

c_so = so_ngay.cell(0,0).paragraphs[0]
fmt_para(c_so, WD_ALIGN_PARAGRAPH.CENTER, sb=4, sa=0, ls=1.0)
add_run(c_so, 'Số:        /SXD-HTKT', size=13)

c_ngay = so_ngay.cell(0,1).paragraphs[0]
fmt_para(c_ngay, WD_ALIGN_PARAGRAPH.CENTER, sb=4, sa=0, ls=1.0)
add_run(c_ngay, 'Quảng Trị, ngày      tháng      năm 2026', size=13, italic=True)

# ══════════════════════════════════════════════════════════════
# PHẦN 3 — TRÍCH YẾU
# ══════════════════════════════════════════════════════════════
new_para('', sb=4, sa=0)

# V/v — căn giữa, nghiêng, 13pt (chuẩn công văn)
for line in [
    'V/v thông báo kết quả thẩm định và yêu cầu bổ sung, hoàn chỉnh hồ sơ',
    'điều chỉnh Báo cáo kinh tế - kỹ thuật dự án Tạo quỹ đất khu dân cư',
    'phía Đông đường Hà Huy Tập, tổ dân phố 6, phường Đông Sơn, tỉnh Quảng Trị',
]:
    new_para(line, WD_ALIGN_PARAGRAPH.CENTER, size=13, italic=True, sb=0, sa=0, ls=1.0)

new_para('', sb=4, sa=0)

# ══════════════════════════════════════════════════════════════
# PHẦN 4 — KÍNH GỬI
# ══════════════════════════════════════════════════════════════
p_kg = new_para(ls=1.15, sb=0, sa=2)
fmt_para(p_kg, WD_ALIGN_PARAGRAPH.LEFT)
add_run(p_kg, 'Kính gửi:', size=14, bold=True, italic=True)
add_run(p_kg, ' Trung tâm Phát triển quỹ đất tỉnh Quảng Trị.', size=14, italic=True)

new_para('', sb=2, sa=0)

# ══════════════════════════════════════════════════════════════
# HELPERS NỘI DUNG
# ══════════════════════════════════════════════════════════════
def body(text, fi=7):
    p = new_para(text, WD_ALIGN_PARAGRAPH.JUSTIFY, size=14, sb=0, sa=2, ls=1.15, fi=fi)
    return p

def body_mixed(parts, fi=7):
    """parts = list of (text, bold, italic, underline)"""
    p = doc.add_paragraph()
    fmt_para(p, WD_ALIGN_PARAGRAPH.JUSTIFY, sb=0, sa=2, ls=1.15, fi=fi)
    for txt, bd, it, ul in parts:
        add_run(p, txt, 14, bd, it, ul)
    return p

def h1(text):
    new_para(text, WD_ALIGN_PARAGRAPH.CENTER, size=14, bold=True,
             caps=True, sb=6, sa=2, ls=1.15)

def h2(text, num):
    p = new_para(ls=1.15, sb=4, sa=2)
    fmt_para(p, WD_ALIGN_PARAGRAPH.LEFT)
    add_run(p, f'{num}. {text}', size=14, bold=True, underline=True)

def item(label, text):
    p = new_para(ls=1.15, sb=0, sa=2, fi=7)
    fmt_para(p, WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_run(p, label + ' ', size=14, bold=True)
    add_run(p, text, size=14)

def yeu_cau(text):
    p = new_para(ls=1.15, sb=0, sa=4, fi=7)
    fmt_para(p, WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_run(p, 'Yêu cầu: ', size=14, bold=True, italic=True)
    add_run(p, text, size=14)

# ══════════════════════════════════════════════════════════════
# NỘI DUNG VĂN BẢN
# ══════════════════════════════════════════════════════════════
body('Sở Xây dựng tỉnh Quảng Trị nhận được hồ sơ kèm theo Tờ trình số 109/TTr-PTQĐ '
     'của Trung tâm Phát triển quỹ đất tỉnh Quảng Trị về việc đề nghị thẩm định điều '
     'chỉnh Báo cáo kinh tế - kỹ thuật đầu tư xây dựng dự án Tạo quỹ đất khu dân cư '
     'phía Đông đường Hà Huy Tập, tổ dân phố 6, phường Đông Sơn, tỉnh Quảng Trị '
     '(sau đây gọi là dự án).')
body('Sau khi nghiên cứu hồ sơ, Sở Xây dựng có ý kiến như sau:')

# ── I ─────────────────────────────────────────────────────────
h1('I. NHẬN XÉT CHUNG')
body('Hồ sơ điều chỉnh Báo cáo kinh tế - kỹ thuật do Công ty TNHH Tư vấn Thiết kế '
     'Phú Sơn lập, trình bày nội dung bổ sung các hạng mục hạ tầng kỹ thuật giai đoạn '
     '2 (vỉa hè, thoát nước mặt, cấp điện, chiếu sáng) với tổng mức đầu tư bổ sung '
     '1.739.000.000 đồng, nâng tổng mức đầu tư điều chỉnh lên 16.689.000.000 đồng. '
     'Bố cục hồ sơ cơ bản đủ thành phần theo quy định. Tuy nhiên, hồ sơ còn một số nội '
     'dung cần bổ sung, hoàn chỉnh trước khi Sở Xây dựng tiếp tục thẩm định, cụ thể như sau:')

# ── II ────────────────────────────────────────────────────────
h1('II. CÁC NỘI DUNG YÊU CẦU BỔ SUNG, HOÀN CHỈNH')

# 1
h2('Về căn cứ pháp lý lập dự toán', '1')
body('Thuyết minh dự toán xây dựng (Sheet TM, file TMDT-HA HUY TAP.xlsx) trích dẫn '
     'các văn bản quy phạm pháp luật đã hết hiệu lực, cụ thể:')

items_1 = [
    ('a)', 'Nghị định số 32/2015/NĐ-CP ngày 25/3/2015 đã hết hiệu lực; '
           'đề nghị thay thế bằng Nghị định số 10/2021/NĐ-CP ngày 09/02/2021 '
           'của Chính phủ về quản lý chi phí đầu tư xây dựng.'),
    ('b)', 'Thông tư số 06/2016/TT-BXD ngày 10/3/2016 và Thông tư số 05/2016/TT-BXD '
           'ngày 10/3/2016 của Bộ Xây dựng đã hết hiệu lực; đề nghị thay thế bằng '
           'Thông tư số 11/2021/TT-BXD ngày 31/8/2021.'),
    ('c)', 'Thông tư số 01/2017/TT-BXD ngày 06/02/2017 về chi phí khảo sát đã hết '
           'hiệu lực; đề nghị thay thế bằng Thông tư số 11/2021/TT-BXD ngày 31/8/2021.'),
    ('d)', 'Thông tư số 09/2016/TT-BTC ngày 18/01/2016 về quyết toán dự án đã hết '
           'hiệu lực; đề nghị thay thế bằng Thông tư số 96/2021/TT-BTC ngày 11/11/2021.'),
    ('đ)', 'Thông tư số 150/2014/TT-BTC ngày 10/10/2014 về phí thẩm duyệt thiết kế PCCC '
           'đã hết hiệu lực từ ngày 01/01/2017; đề nghị thay thế bằng '
           'Thông tư số 258/2016/TT-BTC ngày 11/11/2016.'),
    ('e)', 'Nghị định số 63/2014/NĐ-CP ngày 26/6/2014 về lựa chọn nhà thầu đã hết '
           'hiệu lực; đề nghị thay thế bằng Nghị định số 24/2024/NĐ-CP ngày 27/02/2024.'),
    ('g)', 'Quyết định số 79/2017/QĐ-BXD ngày 15/02/2017 về định mức chi phí quản lý dự '
           'án đã hết hiệu lực; đề nghị áp dụng Thông tư số 11/2021/TT-BXD và các văn '
           'bản hướng dẫn hiện hành.'),
    ('h)', 'Quyết định số 06/2016/QĐ-UBND ngày 29/4/2016 của UBND tỉnh Quảng Bình về '
           'biểu cước vận chuyển không còn áp dụng trên địa bàn tỉnh Quảng Trị; đề nghị '
           'cập nhật theo biểu cước hiện hành của tỉnh Quảng Trị.'),
    ('i)', 'Các công bố định mức dự toán xây dựng số 1776/BXD-VP (2007), 1129/BXD-VP '
           '(2009), 588/BXD-VP (2014) và định mức vật tư 1784/BXD-VP (2007) đã hết hiệu '
           'lực; đề nghị áp dụng Thông tư số 12/2021/TT-BXD ngày 31/8/2021 của Bộ Xây dựng.'),
]
for lbl, txt in items_1: item(lbl, txt)
yeu_cau('Trung tâm Phát triển quỹ đất tỉnh Quảng Trị chỉ đạo đơn vị tư vấn rà soát, '
        'cập nhật toàn bộ căn cứ pháp lý trong thuyết minh dự toán, đảm bảo các văn bản '
        'trích dẫn còn hiệu lực tại thời điểm lập hồ sơ.')

# 2
h2('Về đơn giá vật liệu xây dựng và nhân công', '2')
body('Dự toán xây dựng được lập trên cơ sở giá vật liệu xây dựng tháng 4/2019 theo Thông '
     'báo số 1358/CB-LN ngày 03/5/2019 của liên ngành tỉnh Quảng Bình (Sheet TM, dòng cơ '
     'sở lập dự toán). Đơn giá vật liệu từ thời điểm hơn 06 năm trước, lại thuộc địa bàn '
     'tỉnh Quảng Bình, không phản ánh mặt bằng giá tại thời điểm lập hồ sơ trên địa bàn '
     'tỉnh Quảng Trị.')
yeu_cau('Trung tâm Phát triển quỹ đất tỉnh Quảng Trị chỉ đạo đơn vị tư vấn cập nhật đơn '
        'giá vật liệu xây dựng theo Thông báo giá vật liệu xây dựng mới nhất do Sở Xây '
        'dựng (hoặc liên Sở) tỉnh Quảng Trị công bố tại thời điểm lập dự toán; cập nhật '
        'đơn giá nhân công theo quy định hiện hành.')

# 3
h2('Về sai sót ngày ban hành trong Thuyết minh BCKTKT', '3')
body_mixed([
    ('Tại Mục I Thuyết minh Báo cáo kinh tế - kỹ thuật, Nghị định số 35/2023/NĐ-CP được '
     'ghi ngày "09/02/2023". Tuy nhiên, ngày ban hành thực tế của Nghị định này là ', False, False, False),
    ('20/06/2023', True, False, False),
    ('. Đề nghị đơn vị tư vấn chỉnh sửa lại ngày ban hành cho chính xác.', False, False, False),
])

# 4
h2('Về quy chuẩn kỹ thuật quốc gia về nước thải sinh hoạt', '4')
body_mixed([
    ('Thuyết minh kỹ thuật áp dụng QCVN 14:2008/BTNMT về nước thải sinh hoạt. '
     'Quy chuẩn này đã hết hiệu lực kể từ ngày ', False, False, False),
    ('01/9/2025', True, False, False),
    (', được thay thế bởi ', False, False, False),
    ('QCVN 14:2025/BTNMT', True, False, False),
    (' ban hành kèm theo Thông tư số 05/2025/TT-BTNMT ngày 28/02/2025 của Bộ Tài nguyên '
     'và Môi trường. Đề nghị đơn vị tư vấn cập nhật, áp dụng QCVN 14:2025/BTNMT '
     'trong hồ sơ thiết kế.', False, False, False),
])

# 5
h2('Về tiêu chuẩn kỹ thuật trong Chỉ dẫn kỹ thuật quản lý chất lượng (Mục V Thuyết minh)', '5')
body('Bảng danh mục tiêu chuẩn thí nghiệm và kiểm tra chất lượng tại Mục V Thuyết minh '
     'có nhiều tiêu chuẩn đã hết hiệu lực hoặc ghi sai số hiệu, cụ thể:')

# Bảng tiêu chuẩn
tbl = doc.add_table(rows=1, cols=3)
tbl.style = 'Table Grid'
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl.columns[0].width = Mm(12)
tbl.columns[1].width = Mm(55)
tbl.columns[2].width = Mm(103)

for cell, txt in zip(tbl.rows[0].cells,
                     ['STT', 'Tiêu chuẩn trong hồ sơ', 'Yêu cầu cập nhật / xử lý']):
    p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(txt); set_run(r, 12, True)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    p.paragraph_format.line_spacing = 1.1

rows_data = [
    ('4',  'TCVN 6735:2000',         'Hết hiệu lực → Thay bằng TCVN 6735:2018'),
    ('5',  'TCVN 9385:2011',         'Sai số hiệu → Sửa thành TCVN 9385:2012'),
    ('6',  'TCVN 3118:1993',         'Hết hiệu lực → Thay bằng TCVN 3118:2022'),
    ('7',  'TCVN 3121:2003',         'Nhiều phần đã thay → Cập nhật theo TCVN 3121:2022'),
    ('8',  'TCVN 197:2002',          'Hết hiệu lực → Thay bằng TCVN 197-1:2014'),
    ('9',  'TCVN 198:1985',          'Hết hiệu lực → Thay bằng TCVN 198:2008'),
    ('10', 'TCVN 1651-1:2008',       'Hết hiệu lực → Thay bằng TCVN 1651-1:2018'),
    ('11', 'TCVN 1651-2:2008',       'Hết hiệu lực → Thay bằng TCVN 1651-2:2018'),
    ('14', 'TCVN 6260:2009',         'Hết hiệu lực → Thay bằng TCVN 6260:2020'),
    ('15', 'TCVN 4030:2003',         'Hết hiệu lực → Thay bằng TCVN 13605:2023'),
    ('16', 'TCVN 4787:2001',         'Hết hiệu lực → Thay bằng TCVN 4787:2009'),
    ('17', 'TCVN 6016:1995',         'Hết hiệu lực → Thay bằng TCVN 6016:2011'),
    ('18', 'TCVN 6017:1995',         'Hết hiệu lực → Thay bằng TCVN 6017:2015'),
    ('19', 'TCVN 4201:1995',         'Hết hiệu lực → Thay bằng TCVN 4201:2012'),
    ('20', 'TCVN 4198:1995',         'Hết hiệu lực → Thay bằng TCVN 4198:2014'),
    ('24', 'TCVN 8867:2011',         'Hết hiệu lực → Thay bằng TCVN 8867:2025'),
    ('3',  'TCXDVN 239:2006',        'Cần cập nhật → Áp dụng TCVN 14524:2025 hoặc TCVN 12252:2020'),
    ('21', '22 TCN 346-06',          'Tiêu chuẩn ngành cũ → Làm rõ hoặc chuyển sang TCVN tương ứng'),
    ('22', '22 TCN 333-06',          'Tiêu chuẩn ngành cũ → Xác định tiêu chuẩn đầm nén hiện hành'),
    ('12', 'TCVN 7572:2006',         'Còn áp dụng → Ghi rõ từng phần (TCVN 7572-1, 7572-2...) theo phép thử'),
    ('25', 'TCVN 8860:2011',         'Còn áp dụng → Ghi rõ từng phần (TCVN 8860-1...) theo phép thử'),
]

for stt, tc, xu_ly in rows_data:
    row = tbl.add_row().cells
    for cell, txt, align in zip(row, [stt, tc, xu_ly],
                                 [WD_ALIGN_PARAGRAPH.CENTER,
                                  WD_ALIGN_PARAGRAPH.LEFT,
                                  WD_ALIGN_PARAGRAPH.LEFT]):
        p = cell.paragraphs[0]; p.alignment = align
        r = p.add_run(txt); set_run(r, 11)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        p.paragraph_format.line_spacing = 1.1

new_para('', sb=2, sa=0)
yeu_cau('Đơn vị tư vấn rà soát, cập nhật toàn bộ danh mục tiêu chuẩn tại Mục V Thuyết '
        'minh; thay thế các tiêu chuẩn đã hết hiệu lực bằng phiên bản hiện hành; ghi rõ '
        'số hiệu đúng và từng phần cụ thể đối với bộ tiêu chuẩn nhiều phần.')

# 6
h2('Về kết cấu vỉa hè', '6')
body('Hồ sơ thiết kế quy định kết cấu lớp bê tông xi măng M150 dày 10 cm cho hạng mục '
     'vỉa hè. Đề nghị đơn vị tư vấn kiểm tra, rà soát lại cường độ bê tông vỉa hè; '
     'trường hợp giữ nguyên M150 thì phải có giải trình kỹ thuật cụ thể và căn cứ '
     'tiêu chuẩn áp dụng kèm theo.')

# 7
h2('Về tên đơn vị hành chính trong hồ sơ', '7')
body_mixed([
    ('Thuyết minh và một số bản vẽ kỹ thuật còn sử dụng địa danh cũ "Phường Bắc Nghĩa, '
     'thành phố Đồng Hới, tỉnh Quảng Bình" trong khi theo Tờ trình số 109/TTr-PTQĐ, '
     'địa điểm xây dựng hiện là ', False, False, False),
    ('"Phường Đông Sơn, tỉnh Quảng Trị"', True, False, False),
    ('. Đề nghị đơn vị tư vấn cập nhật thống nhất tên đơn vị hành chính '
     'trong toàn bộ hồ sơ.', False, False, False),
])

# ── III ───────────────────────────────────────────────────────
h1('III. THỜI HẠN BỔ SUNG HỒ SƠ')
body_mixed([
    ('Đề nghị Trung tâm Phát triển quỹ đất tỉnh Quảng Trị chỉ đạo đơn vị tư vấn '
     'hoàn chỉnh các nội dung nêu trên và nộp lại hồ sơ về Sở Xây dựng trong thời hạn ',
     False, False, False),
    ('15 ngày làm việc', True, False, False),
    (' kể từ ngày nhận được văn bản này. Trường hợp cần thêm thời gian, Trung tâm Phát '
     'triển quỹ đất tỉnh Quảng Trị có văn bản đề nghị gia hạn gửi Sở Xây dựng '
     'trước khi hết hạn.', False, False, False),
])
body('Sở Xây dựng tỉnh Quảng Trị thông báo để Trung tâm Phát triển quỹ đất tỉnh '
     'Quảng Trị biết, phối hợp thực hiện./.', fi=7)

new_para('', sb=6, sa=0)

# ══════════════════════════════════════════════════════════════
# PHẦN CUỐI — NƠI NHẬN & KÝ TÊN
# ══════════════════════════════════════════════════════════════
ky = doc.add_table(rows=1, cols=2)
ky.alignment = WD_TABLE_ALIGNMENT.CENTER
no_border(ky)
ky.cell(0,0).width = COL_L
ky.cell(0,1).width = COL_R

# Nơi nhận
nn = ky.cell(0,0)
p_nn = nn.paragraphs[0]; fmt_para(p_nn, WD_ALIGN_PARAGRAPH.LEFT, sb=0, sa=0, ls=1.0)
add_run(p_nn, 'Nơi nhận:', size=12, bold=True, italic=True)
for txt in ['- Như trên;', '- UBND tỉnh (b/c);', '- Lưu: VT, HTKT.']:
    pn = nn.add_paragraph(txt)
    fmt_para(pn, WD_ALIGN_PARAGRAPH.LEFT, sb=0, sa=0, ls=1.1)
    for r in pn.runs: set_run(r, 12)

# Ký tên
kt = ky.cell(0,1)
for txt, bold in [('KT. GIÁM ĐỐC', True), ('PHÓ GIÁM ĐỐC', True)]:
    p = kt.paragraphs[0] if txt == 'KT. GIÁM ĐỐC' else kt.add_paragraph()
    fmt_para(p, WD_ALIGN_PARAGRAPH.CENTER, sb=0, sa=0, ls=1.1)
    add_run(p, txt, size=13, bold=bold)

for _ in range(4):
    pe = kt.add_paragraph()
    fmt_para(pe, WD_ALIGN_PARAGRAPH.CENTER, sb=0, sa=0, ls=1.2)

p_ten = kt.add_paragraph()
fmt_para(p_ten, WD_ALIGN_PARAGRAPH.CENTER, sb=0, sa=0, ls=1.0)
add_run(p_ten, 'Nguyễn Xuân Hoàng', size=14, bold=True)

# ══════════════════════════════════════════════════════════════
out = r'D:\Dropbox\Dropbox\Luu_du_lieu\Binh 2026\Tham dinh\Đường Hà Huy Tập TDP6\CV_yeu_cau_bo_sung_HA_HUY_TAP_v4.docx'
doc.save(out)
import sys; sys.stdout.buffer.write(('OK: ' + out + '\n').encode('utf-8'))
