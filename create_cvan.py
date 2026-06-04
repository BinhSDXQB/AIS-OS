
# -*- coding: utf-8 -*-
from docx import Document
from docx.shared import Pt, Mm, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ── Page setup ──────────────────────────────────────────────
sec = doc.sections[0]
sec.page_width  = Mm(210)
sec.page_height = Mm(297)
sec.left_margin   = Mm(30)
sec.right_margin  = Mm(15)
sec.top_margin    = Mm(20)
sec.bottom_margin = Mm(20)

# ── Helper functions ─────────────────────────────────────────
def set_font(run, size=14, bold=False, italic=False, underline=False, color=None):
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    if color:
        run.font.color.rgb = RGBColor(*color)

def add_para(text='', align=WD_ALIGN_PARAGRAPH.LEFT, size=14, bold=False,
             italic=False, underline=False, space_before=0, space_after=0,
             line_spacing=1.1, first_indent=None):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after  = Pt(space_after)
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = line_spacing
    if first_indent is not None:
        pf.first_line_indent = Mm(first_indent)
    if text:
        run = p.add_run(text)
        set_font(run, size, bold, italic, underline)
    return p

def add_run(para, text, size=14, bold=False, italic=False, underline=False):
    run = para.add_run(text)
    set_font(run, size, bold, italic, underline)
    return run

def remove_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top','left','bottom','right','insideH','insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        tblBorders.append(border)
    tblPr.append(tblBorders)

# ─────────────────────────────────────────────────────────────
# BLOCK 1: Header table (Cơ quan | Quốc hiệu)
# ─────────────────────────────────────────────────────────────
tbl = doc.add_table(rows=1, cols=2)
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl.columns[0].width = Mm(85)
tbl.columns[1].width = Mm(95)
remove_table_borders(tbl)

# Left cell: Cơ quan ban hành
lc = tbl.cell(0, 0)
lc.vertical_alignment = WD_ALIGN_VERTICAL.TOP
p1 = lc.paragraphs[0]
p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p1.add_run('UBND TỈNH QUẢNG TRỊ')
set_font(r, size=12, bold=True)

p2 = lc.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
r2 = p2.add_run('SỞ XÂY DỰNG')
set_font(r2, size=13, bold=True)
# Underline the SXD paragraph
p2.runs[0].underline = True

# Right cell: Quốc hiệu
rc = tbl.cell(0, 1)
rc.vertical_alignment = WD_ALIGN_VERTICAL.TOP
p3 = rc.paragraphs[0]
p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
r3 = p3.add_run('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM')
set_font(r3, size=12, bold=True)

p4 = rc.add_paragraph()
p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
r4 = p4.add_run('Độc lập - Tự do - Hạnh phúc')
set_font(r4, size=13, bold=True, underline=True)

# ─────────────────────────────────────────────────────────────
# BLOCK 2: Số văn bản | Địa danh ngày tháng
# ─────────────────────────────────────────────────────────────
tbl2 = doc.add_table(rows=1, cols=2)
tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl2.columns[0].width = Mm(85)
tbl2.columns[1].width = Mm(95)
remove_table_borders(tbl2)

lc2 = tbl2.cell(0, 0)
p5 = lc2.paragraphs[0]
p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
r5 = p5.add_run('Số:         /SXD-HTKT')
set_font(r5, size=13)

rc2 = tbl2.cell(0, 1)
p6 = rc2.paragraphs[0]
p6.alignment = WD_ALIGN_PARAGRAPH.CENTER
r6 = p6.add_run('Quảng Trị, ngày      tháng      năm 2026')
set_font(r6, size=14, italic=True)

# ─────────────────────────────────────────────────────────────
# BLOCK 3: Trích yếu
# ─────────────────────────────────────────────────────────────
add_para('', space_before=6, space_after=0)

p_ty = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=0, line_spacing=1.1)
add_run(p_ty, 'V/v thông báo kết quả thẩm định và yêu cầu bổ sung, hoàn chỉnh hồ sơ', size=13, bold=True)
p_ty2 = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=0, line_spacing=1.1)
add_run(p_ty2, 'điều chỉnh Báo cáo kinh tế - kỹ thuật dự án Tạo quỹ đất khu dân cư', size=13, bold=True)
p_ty3 = add_para(align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=0, line_spacing=1.1)
add_run(p_ty3, 'phía Đông đường Hà Huy Tập, tổ dân phố 6, phường Bắc Nghĩa, thành phố Đồng Hới', size=13, bold=True)

# Gạch chân dưới trích yếu - dùng border bottom của paragraph
for p_under in [p_ty, p_ty2, p_ty3]:
    pass  # underline handled via bold display

add_para('', space_before=2)

# ─────────────────────────────────────────────────────────────
# BLOCK 4: Kính gửi
# ─────────────────────────────────────────────────────────────
p_kg = add_para(line_spacing=1.1, space_before=0, space_after=0)
add_run(p_kg, 'Kính gửi: ', size=14, bold=True, italic=True)
add_run(p_kg, 'Trung tâm Phát triển quỹ đất tỉnh Quảng Trị.', size=14, italic=True)

add_para('', space_before=2)

# ─────────────────────────────────────────────────────────────
# BLOCK 5: Nội dung chính
# ─────────────────────────────────────────────────────────────
def body(text, first=True):
    indent = 7 if first else None
    add_para(text, size=14, line_spacing=1.1, space_before=0, space_after=2,
             first_indent=indent if first else None)

def body_no_indent(text):
    add_para(text, size=14, line_spacing=1.1, space_before=0, space_after=2)

def heading1(text):
    add_para(text, align=WD_ALIGN_PARAGRAPH.CENTER, size=14, bold=True,
             line_spacing=1.1, space_before=4, space_after=2)

def heading2(text):
    p = add_para(line_spacing=1.1, space_before=2, space_after=2, first_indent=7)
    add_run(p, text, size=14, bold=True, underline=True)

def bullet(text):
    add_para(text, size=14, line_spacing=1.1, space_before=0, space_after=2,
             first_indent=None)

# Mở đầu
p_mo = add_para(line_spacing=1.1, space_before=0, space_after=2, first_indent=7)
add_run(p_mo,
    'Sở Xây dựng tỉnh Quảng Trị nhận được hồ sơ kèm theo Tờ trình số 109/TTr-PTQĐ '
    'của Trung tâm Phát triển quỹ đất tỉnh Quảng Trị về việc đề nghị thẩm định điều chỉnh '
    'Báo cáo kinh tế - kỹ thuật đầu tư xây dựng dự án Tạo quỹ đất khu dân cư phía Đông '
    'đường Hà Huy Tập, tổ dân phố 6, phường Bắc Nghĩa, thành phố Đồng Hới '
    '(sau đây gọi là dự án).', size=14)

p_mo2 = add_para(line_spacing=1.1, space_before=0, space_after=2, first_indent=7)
add_run(p_mo2, 'Sau khi nghiên cứu hồ sơ, Sở Xây dựng có ý kiến như sau:', size=14)

# I. NHẬN XÉT CHUNG
heading1('I. NHẬN XÉT CHUNG')

p_nxc = add_para(line_spacing=1.1, space_before=0, space_after=2, first_indent=7)
add_run(p_nxc,
    'Hồ sơ điều chỉnh Báo cáo kinh tế - kỹ thuật do Công ty TNHH Tư vấn Thiết kế Phú Sơn lập, '
    'trình bày nội dung bổ sung các hạng mục hạ tầng kỹ thuật giai đoạn 2 (vỉa hè, thoát nước mặt, '
    'cấp điện, chiếu sáng) với tổng mức đầu tư bổ sung 1.739.000.000 đồng, nâng tổng mức đầu tư '
    'điều chỉnh lên 16.689.000.000 đồng. Bố cục hồ sơ cơ bản đủ thành phần theo quy định. '
    'Tuy nhiên, hồ sơ còn một số nội dung chưa đảm bảo, cần được bổ sung, hoàn chỉnh trước khi '
    'Sở Xây dựng tiếp tục thẩm định, cụ thể như sau:', size=14)

# II. CÁC NỘI DUNG YÊU CẦU
heading1('II. CÁC NỘI DUNG YÊU CẦU BỔ SUNG, HOÀN CHỈNH')

# 1.
heading2('1. Về căn cứ pháp lý lập dự toán')

p1a = add_para(line_spacing=1.1, space_before=0, space_after=2, first_indent=7)
add_run(p1a,
    'Thuyết minh dự toán xây dựng trong hồ sơ trích dẫn một số văn bản quy phạm pháp luật '
    'đã hết hiệu lực, cụ thể:', size=14)

p1b = add_para(line_spacing=1.1, space_before=0, space_after=2, first_indent=7)
add_run(p1b, 'a) ', size=14, bold=True)
add_run(p1b,
    'Nghị định số 32/2015/NĐ-CP ngày 25/3/2015 của Chính phủ về quản lý chi phí đầu tư xây dựng '
    'công trình đã hết hiệu lực; đề nghị thay thế bằng Nghị định số 10/2021/NĐ-CP ngày 09/02/2021 '
    'của Chính phủ về quản lý chi phí đầu tư xây dựng và các văn bản hướng dẫn hiện hành.', size=14)

p1c = add_para(line_spacing=1.1, space_before=0, space_after=2, first_indent=7)
add_run(p1c, 'b) ', size=14, bold=True)
add_run(p1c,
    'Thông tư số 06/2016/TT-BXD ngày 10/3/2016 và Thông tư số 05/2016/TT-BXD ngày 10/3/2016 '
    'của Bộ Xây dựng đã hết hiệu lực; đề nghị thay thế bằng Thông tư số 11/2021/TT-BXD ngày '
    '31/8/2021 của Bộ Xây dựng hướng dẫn một số nội dung xác định và quản lý chi phí đầu tư xây dựng.', size=14)

p1d = add_para(line_spacing=1.1, space_before=0, space_after=2, first_indent=7)
add_run(p1d, 'c) ', size=14, bold=True)
add_run(p1d,
    'Quyết định số 06/2016/QĐ-UBND ngày 29/4/2016 của UBND tỉnh Quảng Bình về biểu cước vận '
    'chuyển hàng hóa không còn hiệu lực áp dụng trên địa bàn tỉnh Quảng Trị; đề nghị cập nhật '
    'theo biểu cước vận chuyển hàng hóa hiện hành trên địa bàn tỉnh Quảng Trị.', size=14)

p1e = add_para(line_spacing=1.1, space_before=0, space_after=2, first_indent=7)
add_run(p1e, 'Yêu cầu: ', size=14, bold=True, italic=True)
add_run(p1e,
    'Trung tâm Phát triển quỹ đất tỉnh Quảng Trị chỉ đạo đơn vị tư vấn rà soát, cập nhật '
    'toàn bộ căn cứ pháp lý trong thuyết minh dự toán, đảm bảo các văn bản trích dẫn còn hiệu '
    'lực thi hành tại thời điểm lập hồ sơ.', size=14)

# 2.
heading2('2. Về đơn giá vật liệu xây dựng và nhân công')

p2a = add_para(line_spacing=1.1, space_before=0, space_after=2, first_indent=7)
add_run(p2a,
    'Dự toán xây dựng được lập trên cơ sở giá vật liệu xây dựng tháng 4/2019 theo Thông báo '
    'số 1358/CB-LN ngày 03/5/2019 của liên ngành tỉnh Quảng Bình. Đơn giá vật liệu từ thời điểm '
    'hơn 06 năm trước không phản ánh mặt bằng giá tại thời điểm lập hồ sơ, không đảm bảo tính '
    'chính xác của tổng mức đầu tư điều chỉnh.', size=14)

p2b = add_para(line_spacing=1.1, space_before=0, space_after=2, first_indent=7)
add_run(p2b, 'Yêu cầu: ', size=14, bold=True, italic=True)
add_run(p2b,
    'Trung tâm Phát triển quỹ đất tỉnh Quảng Trị chỉ đạo đơn vị tư vấn cập nhật đơn giá vật '
    'liệu xây dựng theo Thông báo giá vật liệu xây dựng mới nhất do Sở Xây dựng (hoặc liên Sở) '
    'tỉnh Quảng Trị công bố tại thời điểm lập dự toán; cập nhật đơn giá nhân công theo quy định hiện hành.', size=14)

# 3.
heading2('3. Về tiêu chuẩn, quy chuẩn kỹ thuật áp dụng')

p3a = add_para(line_spacing=1.1, space_before=0, space_after=2, first_indent=7)
add_run(p3a, 'a) ', size=14, bold=True)
add_run(p3a,
    'Tiêu chuẩn 22 TCN 223-95 (áo đường cứng) ban hành từ năm 1995, đã có các tiêu chuẩn kỹ '
    'thuật cập nhật hơn. Đề nghị đơn vị tư vấn rà soát, áp dụng tiêu chuẩn phù hợp còn hiệu lực hiện hành.', size=14)

p3b = add_para(line_spacing=1.1, space_before=0, space_after=2, first_indent=7)
add_run(p3b, 'b) ', size=14, bold=True)
add_run(p3b,
    'Tiêu chuẩn TCXDVN 259:2001 về chiếu sáng nhân tạo đường phố đã có phiên bản cập nhật; '
    'đề nghị rà soát, áp dụng TCVN 8773:2011 hoặc quy chuẩn kỹ thuật tương ứng hiện hành.', size=14)

p3c = add_para(line_spacing=1.1, space_before=0, space_after=2, first_indent=7)
add_run(p3c, 'c) ', size=14, bold=True)
add_run(p3c,
    'Tiêu chuẩn TCXDVN 7957-2008 về thoát nước: đề nghị kiểm tra và áp dụng phiên bản mới nhất '
    'đang có hiệu lực.', size=14)

p3d = add_para(line_spacing=1.1, space_before=0, space_after=2, first_indent=7)
add_run(p3d, 'Yêu cầu: ', size=14, bold=True, italic=True)
add_run(p3d,
    'Đơn vị tư vấn rà soát toàn bộ danh mục tiêu chuẩn, quy chuẩn tại Mục IV Thuyết minh; '
    'thay thế các tiêu chuẩn đã hết hiệu lực hoặc có phiên bản cập nhật; bổ sung QCVN '
    '07-4:2016/BXD (Các công trình hạ tầng kỹ thuật đô thị) nếu chưa được áp dụng.', size=14)

# 4.
heading2('4. Về kết cấu vỉa hè')

p4a = add_para(line_spacing=1.1, space_before=0, space_after=2, first_indent=7)
add_run(p4a,
    'Hồ sơ thiết kế quy định kết cấu lớp bê tông xi măng M150 dày 10 cm cho hạng mục vỉa hè. '
    'Đề nghị đơn vị tư vấn kiểm tra, rà soát lại cường độ bê tông vỉa hè; trường hợp giữ nguyên '
    'M150 thì phải có giải trình kỹ thuật cụ thể và căn cứ tiêu chuẩn áp dụng kèm theo.', size=14)

# 5.
heading2('5. Về chi phí ứng vốn')

p5a = add_para(line_spacing=1.1, space_before=0, space_after=2, first_indent=7)
add_run(p5a,
    'Chi phí ứng vốn dự án được tính theo công thức TMĐT × 0,21% × 6. Đề nghị Trung tâm Phát '
    'triển quỹ đất tỉnh Quảng Trị làm rõ cơ sở xác định lãi suất 0,21%/tháng theo quy định nào '
    'của Quỹ Phát triển đất tỉnh Quảng Trị hiện hành; đồng thời xác nhận thời gian vay 06 tháng '
    'có còn phù hợp với tiến độ thực hiện sau điều chỉnh không.', size=14)

# 6.
heading2('6. Về năng lực đơn vị tư vấn')

p6a = add_para(line_spacing=1.1, space_before=0, space_after=2, first_indent=7)
add_run(p6a,
    'Đề nghị Trung tâm Phát triển quỹ đất tỉnh Quảng Trị bổ sung chứng chỉ năng lực hoạt động '
    'xây dựng còn hiệu lực của Công ty TNHH Tư vấn Thiết kế Phú Sơn (lĩnh vực lập dự án, thiết '
    'kế công trình hạ tầng kỹ thuật) vào hồ sơ trình thẩm định, theo quy định tại Nghị định số '
    '175/2024/NĐ-CP ngày 30/12/2024 của Chính phủ.', size=14)

# III. THỜI HẠN
heading1('III. THỜI HẠN BỔ SUNG HỒ SƠ')

p_tl = add_para(line_spacing=1.1, space_before=0, space_after=2, first_indent=7)
add_run(p_tl,
    'Đề nghị Trung tâm Phát triển quỹ đất tỉnh Quảng Trị chỉ đạo đơn vị tư vấn hoàn chỉnh '
    'các nội dung nêu trên và nộp lại hồ sơ về Sở Xây dựng trong thời hạn ', size=14)
add_run(p_tl, '15 ngày làm việc', size=14, bold=True)
add_run(p_tl,
    ' kể từ ngày nhận được văn bản này. Trường hợp cần thêm thời gian, Trung tâm Phát triển '
    'quỹ đất tỉnh Quảng Trị có văn bản đề nghị gia hạn gửi Sở Xây dựng trước khi hết hạn.', size=14)

p_ket = add_para(line_spacing=1.1, space_before=0, space_after=2, first_indent=7)
add_run(p_ket,
    'Sở Xây dựng tỉnh Quảng Trị thông báo để Trung tâm Phát triển quỹ đất tỉnh Quảng Trị '
    'biết, phối hợp thực hiện./.', size=14)

add_para('', space_before=4)

# ─────────────────────────────────────────────────────────────
# BLOCK 6: Nơi nhận | Ký tên
# ─────────────────────────────────────────────────────────────
tbl3 = doc.add_table(rows=1, cols=2)
tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl3.columns[0].width = Mm(85)
tbl3.columns[1].width = Mm(95)
remove_table_borders(tbl3)

lc3 = tbl3.cell(0, 0)
p_nn = lc3.paragraphs[0]
r_nn = p_nn.add_run('Nơi nhận:')
set_font(r_nn, size=12, bold=True, italic=True)

for nn in [
    '- Như trên;',
    '- UBND tỉnh (b/c);',
    '- Lưu: VT, HTKT.',
]:
    pn = lc3.add_paragraph(nn)
    pn.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pn.paragraph_format.line_spacing = 1.1
    for run in pn.runs:
        set_font(run, size=12)

rc3 = tbl3.cell(0, 1)
p_kt = rc3.paragraphs[0]
p_kt.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_kt = p_kt.add_run('KT. GIÁM ĐỐC')
set_font(r_kt, size=14, bold=True)

p_pgd = rc3.add_paragraph()
p_pgd.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_pgd = p_pgd.add_run('PHÓ GIÁM ĐỐC')
set_font(r_pgd, size=14, bold=True)

# Khoảng trống để ký
for _ in range(4):
    pe = rc3.add_paragraph('')
    pe.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pe.paragraph_format.line_spacing = 1.1

p_ten = rc3.add_paragraph()
p_ten.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_ten = p_ten.add_run('Nguyễn Xuân Hoàng')
set_font(r_ten, size=14, bold=True)

# ─────────────────────────────────────────────────────────────
# Save
# ─────────────────────────────────────────────────────────────
out_path = r'D:\Dropbox\Dropbox\Luu_du_lieu\Binh 2026\Tham dinh\Đường Hà Huy Tập TDP6\CV_yeu_cau_bo_sung_HA_HUY_TAP.docx'
doc.save(out_path)
import sys
sys.stdout.buffer.write(('Saved: ' + out_path + '\n').encode('utf-8'))
