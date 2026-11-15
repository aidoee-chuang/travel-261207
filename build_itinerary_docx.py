from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

OUTPUT = '苗栗泰安兩天一夜每日時間軸.docx'
UPDATED_AT = '2026.11.15 19:14'

DAYS = [
    ('第一天', '12/7 星期一', '', [
        ('12:30–14:00', '午餐 苗栗大湖福菜園客家菜餐廳', '苗栗縣大湖鄉法雲寺5-5號', '提供客家料理，適合團體共餐。'),
        ('14:00–14:30', '前往泰安觀止溫泉會館', '苗栗縣泰安鄉錦水村圓墩58號', '坐落在苗栗泰安的汶水溪畔的溫泉旅館。住一晚，可以從下午的溪谷光影一路看到隔天清晨的山景。'),
        ('14:30–18:00', '泡湯、下午茶', '', '泡湯需自備泳帽、泳衣；飯店提供拖鞋、大毛巾與沐浴用品。\n露天風呂：14:00–22:00\n下午茶：14:00–16:30'),
        ('18:00–19:30', '晚餐', '', ''),
    ]),
    ('第二天', '12/8 星期二', '', [
        ('07:30–10:00', '自行早餐及自由活動', '', ''),
        ('10:00–10:30', '前往烏嘎彥竹林', '苗栗縣泰安鄉大興村烏嘎彥竹林', '位在苗栗縣泰安鄉大興村。小徑兩旁的竹子向上交疊，形成一段約 100 公尺的天然隧道。晴天時，陽光從竹葉縫隙灑進來，走在裡面就能看到深淺不同的綠色。'),
        ('10:30–11:30', '烏嘎彥竹林散步', '', ''),
        ('11:30–12:00', '前往荷塘居', '苗栗縣公館鄉館南村 13 鄰 332 號', '荷塘居結合田園景觀與美食，提供套餐與桌菜；桌菜請先透過臉書預約。'),
        ('12:00–13:30', '午餐 荷塘居', '', ''),
    ]),
]

def shade(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), fill)
    tc_pr.append(shd)

def set_cell_text(cell, text, bold=False, color=None):
    cell.text = ''
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(text)
    run.bold = bold
    run.font.name = 'PingFang TC'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang TC')
    run.font.size = Pt(10)
    if color:
        run.font.color.rgb = RGBColor(*color)

doc = Document()
section = doc.sections[0]
section.top_margin = Inches(0.6)
section.bottom_margin = Inches(0.6)
section.left_margin = Inches(0.65)
section.right_margin = Inches(0.65)

styles = doc.styles
styles['Normal'].font.name = 'PingFang TC'
styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang TC')
styles['Normal'].font.size = Pt(10.5)

title = doc.add_paragraph(style='Title')
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run('苗栗泰安兩天一夜行程')
run.font.name = 'PingFang TC'
run._element.rPr.rFonts.set(qn('w:eastAsia'), 'PingFang TC')
run.font.color.rgb = RGBColor(0, 0, 0)

intro = doc.add_paragraph('12/7–12/8｜福菜園、泰安觀止、烏嘎彥竹林與荷塘居')
intro.alignment = WD_ALIGN_PARAGRAPH.CENTER
intro.paragraph_format.space_after = Pt(2)

updated = doc.add_paragraph(f'更新時間：{UPDATED_AT}')
updated.alignment = WD_ALIGN_PARAGRAPH.CENTER
updated.paragraph_format.space_after = Pt(12)

for day_index, (day, date, subtitle, items) in enumerate(DAYS):
    heading = doc.add_heading(f'{day}｜{date}', level=1)
    heading.paragraph_format.space_before = Pt(10)
    heading.paragraph_format.space_after = Pt(2)
    if subtitle:
        doc.add_paragraph(subtitle).paragraph_format.space_after = Pt(8)

    if day_index == 0:
        h = doc.add_heading('泰安觀止住宿資訊', level=2)
        h.paragraph_format.space_before = Pt(6)
        lodging = [
            '戶外泡湯區：早場 07:00–12:00，最後入場 11:00；午場 14:00–22:00，最後入場 21:00。',
            '泳裝區：包含戶外冷水池、溫泉池、蒸氣室，需著泳帽、泳衣。',
            '裸湯區：包含男、女裸湯露天池區、蒸氣室、烤箱。',
            '房客不限入場次數。',
            '房型：觀水日式兩人房（一大床）× 3；觀水歐式兩人房（兩小床）× 1；觀山四人房（兩大床）× 1。',
        ]
        for text in lodging:
            doc.add_paragraph(text, style='List Bullet')

    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    table.autofit = False
    widths = [Inches(1.0), Inches(2.0), Inches(3.9)]
    for cell, label, width in zip(table.rows[0].cells, ['時間', '行程', '資訊'], widths):
        cell.width = width
        shade(cell, '174F43')
        set_cell_text(cell, label, True, (255, 255, 255))
    for time, title, address, description in items:
        cells = table.add_row().cells
        set_cell_text(cells[0], time, True, (23, 79, 67))
        set_cell_text(cells[1], title, True)
        detail = (f'地址：{address}\n' if address else '') + description
        set_cell_text(cells[2], detail)
    doc.add_paragraph().paragraph_format.space_after = Pt(3)

doc.save(OUTPUT)
print(OUTPUT)
