from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

OUTPUT = '苗栗泰安兩天一夜每日時間軸.docx'


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), fill)
    tc_pr.append(shd)


def set_cell_margins(cell, top=110, start=130, bottom=110, end=130):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in('w:tcMar')
    if tc_mar is None:
        tc_mar = OxmlElement('w:tcMar')
        tc_pr.append(tc_mar)
    for margin, value in [('top', top), ('start', start), ('bottom', bottom), ('end', end)]:
        node = tc_mar.find(qn(f'w:{margin}'))
        if node is None:
            node = OxmlElement(f'w:{margin}')
            tc_mar.append(node)
        node.set(qn('w:w'), str(value))
        node.set(qn('w:type'), 'dxa')


def set_cell_border(cell, color='D9D9D9'):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in('w:tcBorders')
    if borders is None:
        borders = OxmlElement('w:tcBorders')
        tc_pr.append(borders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        element = borders.find(qn(f'w:{edge}'))
        if element is None:
            element = OxmlElement(f'w:{edge}')
            borders.append(element)
        element.set(qn('w:val'), 'single')
        element.set(qn('w:sz'), '6')
        element.set(qn('w:color'), color)


def add_run(paragraph, text, bold=False, size=10.5, color=None):
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.name = 'PingFang TC'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang TC')
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor(*color)
    return run


def add_timeline_table(document, rows):
    table = document.add_table(rows=1, cols=3)
    table.autofit = False
    table.columns[0].width = Cm(3.1)
    table.columns[1].width = Cm(5.2)
    table.columns[2].width = Cm(8.0)
    headers = ['時間', '行程', '說明']
    for i, text in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.width = table.columns[i].width
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_shading(cell, '19493E')
        set_cell_margins(cell)
        set_cell_border(cell)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run(p, text, bold=True, size=10.5, color=(255, 255, 255))
    for row_index, (time, activity, detail) in enumerate(rows):
        cells = table.add_row().cells
        for i, value in enumerate((time, activity, detail)):
            cells[i].width = table.columns[i].width
            cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if row_index % 2 == 1:
                set_cell_shading(cells[i], 'F3F7F5')
            set_cell_margins(cells[i])
            set_cell_border(cells[i])
            p = cells[i].paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i == 0 else WD_ALIGN_PARAGRAPH.LEFT
            add_run(p, value, bold=(i == 1), size=10.2)
    document.add_paragraph().paragraph_format.space_after = Pt(4)


doc = Document()
section = doc.sections[0]
section.top_margin = Cm(1.7)
section.bottom_margin = Cm(1.7)
section.left_margin = Cm(1.8)
section.right_margin = Cm(1.8)

styles = doc.styles
styles['Normal'].font.name = 'PingFang TC'
styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang TC')
styles['Normal'].font.size = Pt(10.5)
styles['Normal'].paragraph_format.space_after = Pt(6)
for name, size in [('Title', 22), ('Heading 1', 16), ('Heading 2', 12)]:
    style = styles[name]
    style.font.name = 'PingFang TC'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang TC')
    style.font.size = Pt(size)
    style.font.color.rgb = RGBColor(0, 0, 0)

title = doc.add_paragraph(style='Title')
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(title, '苗栗泰安兩天一夜 每日時間軸', bold=True, size=22)
subtitle = doc.add_paragraph()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(subtitle, '2026 年 12 月 7 日至 8 日  11 人家庭旅行', size=11, color=(75, 95, 88))

intro = doc.add_paragraph()
intro.paragraph_format.space_before = Pt(10)
add_run(intro, '行程摘要：', bold=True)
add_run(intro, '第一天於苗栗高鐵站集合後前往泰安觀止，於 12:30–14:00 享用午餐並於下午泡湯休憩；第二天造訪烏嘎彥竹林，在荷塘居午餐後返回苗栗高鐵站。山區路段請以當日導航與路況為準。')

day1 = [
    ('10:30–10:45', '苗栗高鐵站集合', '英哲和筠婷父母集合；實際集合方式與時間請於出發前確認。'),
    ('10:45–12:30', '前往泰安觀止溫泉會館', '預留行車、停車與集合緩衝，請依當日導航調整。'),
    ('12:30–14:00', '午餐 泰安觀止溫泉會館', '苗栗縣泰安鄉錦水村圓墩58號。'),
    ('14:00–18:00', '泡湯、下午茶與自由時間', '飯店 14:00 開放泡湯；請攜帶泳帽、泳衣並享用下午茶。'),
    ('18:00–19:30', '飯店晚餐', '於泰安觀止享用晚餐。'),
]
day2 = [
    ('08:00–09:00', '泰安觀止早餐', '享用早餐，確認當日天氣與山區路況。'),
    ('09:00–09:30', '整理行李 退房', '確認行李與車輛分組後出發。'),
    ('09:30–10:15', '前往烏嘎彥竹林', '山路狹窄且有坡度，請依路況安全行駛。'),
    ('10:15–11:00', '烏嘎彥竹林散步 拍照', '短段參觀，以安全與舒適為優先。'),
    ('11:00–12:00', '前往荷塘居', '由竹林前往公館；請依當日導航調整抵達時間。'),
    ('12:00–13:30', '午餐 荷塘居', '苗栗縣公館鄉館南村 13 鄰 332 號。'),
    ('13:45–14:30', '前往苗栗高鐵站', '用餐後前往苗栗高鐵站。'),
    ('14:30–15:00', '苗栗高鐵站集合 解散', '清點人員及行李後，各自安全返程。'),
]

doc.add_heading('第一天 12 月 7 日 星期一', level=1)
p = doc.add_paragraph()
add_run(p, '泰安觀止午餐  泰安觀止溫泉休憩', bold=True, size=11)
add_timeline_table(doc, day1)

doc.add_heading('第二天 12 月 8 日 星期二', level=1)
p = doc.add_paragraph()
add_run(p, '泰安早餐  烏嘎彥竹林  荷塘居午餐', bold=True, size=11)
add_timeline_table(doc, day2)

doc.add_heading('重要提醒', level=2)
for text in [
    '泰安觀止 14:00 起開放泡湯，請自備泳衣。',
    '烏嘎彥竹林山路狹窄且有坡度，天候不佳或路況不適合時，應以安全為優先調整行程。',
    '用餐座位、飯店設施與集合細節，請於出發前再次確認。',
]:
    p = doc.add_paragraph(style='Normal')
    p.style = styles['Normal']
    p.paragraph_format.left_indent = Cm(0.5)
    add_run(p, '• ' + text)

footer = section.footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(footer, '苗栗泰安兩天一夜 2026', size=8.5, color=(100, 120, 113))

doc.save(OUTPUT)
print(OUTPUT)
