import docx
from docx.shared import Inches, Pt, RGBColor, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
import os

FOREST_HEX = "294B40"
FOREST_RGB = RGBColor(41, 75, 64)
SAGE_HEX = "89957B"
SAGE_RGB = RGBColor(137, 149, 123)
BEIGE_HEX = "EAE6DB"
BORDER_HEX = "D5DEC9"
IVORY_HEX = "FAF9F4"
MUTED_HEX = "6E8575"
MUTED_RGB = RGBColor(110, 133, 117)
DARK_TEXT_HEX = "2A3831"
DARK_TEXT_RGB = RGBColor(42, 56, 49)

def set_para_rtl(p, align=WD_ALIGN_PARAGRAPH.RIGHT):
    p.alignment = align
    pPr = p._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)

def add_persian_run(p, text, font_name="Vazirmatn", size_pt=10, bold=False, color_rgb=DARK_TEXT_RGB, italic=False):
    r = p.add_run(text)
    r.font.name = font_name
    r.font.size = Pt(size_pt)
    r.bold = bold
    r.italic = italic
    r.font.color.rgb = color_rgb
    
    rPr = r._r.get_or_add_rPr()
    rtl = OxmlElement('w:rtl')
    rtl.set(qn('w:val'), '1')
    rPr.append(rtl)
    
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)
    rPr.append(rFonts)
    return r

def set_table_rtl(tbl):
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = tbl._tbl.tblPr
    bidi = OxmlElement('w:bidiVisual')
    tblPr.append(bidi)

def set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=140, bottom=140, left=180, right=180):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def set_cell_border(cell, **kwargs):
    """
    kwargs: top, bottom, left, right
    values: dict(val='single', sz='4', color='D5DEC9') or None
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        edge_data = kwargs.get(edge)
        if edge_data:
            el = OxmlElement(f'w:{edge}')
            el.set(qn('w:val'), edge_data.get('val', 'single'))
            el.set(qn('w:sz'), str(edge_data.get('sz', 4)))
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), edge_data.get('color', BORDER_HEX))
            tcBorders.append(el)
        else:
            el = OxmlElement(f'w:{edge}')
            el.set(qn('w:val'), 'none')
            tcBorders.append(el)
    tcPr.append(tcBorders)

def create_section_header(doc, number_str, title_str):
    p = doc.add_paragraph()
    set_para_rtl(p, WD_ALIGN_PARAGRAPH.RIGHT)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    
    add_persian_run(p, f"  {number_str}  ", font_name="Vazirmatn", size_pt=9.5, bold=True, color_rgb=RGBColor(250, 249, 244))
    r_pill = p.runs[-1]
    # Highlight or background for pill
    rPr = r_pill._r.get_or_add_rPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{FOREST_HEX}"/>')
    rPr.append(shd)
    
    add_persian_run(p, f"  {title_str}", font_name="Vazirmatn", size_pt=12.5, bold=True, color_rgb=FOREST_RGB)

def add_writing_box(doc, label_text, lines=3, fill_hex=IVORY_HEX):
    p_label = doc.add_paragraph()
    set_para_rtl(p_label, WD_ALIGN_PARAGRAPH.RIGHT)
    p_label.paragraph_format.space_before = Pt(4)
    p_label.paragraph_format.space_after = Pt(2)
    p_label.paragraph_format.keep_with_next = True
    
    add_persian_run(p_label, "◆  ", font_name="Vazirmatn", size_pt=8, bold=True, color_rgb=SAGE_RGB)
    add_persian_run(p_label, label_text, font_name="Vazirmatn", size_pt=9.5, bold=True, color_rgb=FOREST_RGB)
    
    tbl = doc.add_table(rows=1, cols=1)
    set_table_rtl(tbl)
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Mm(174)
    
    set_cell_shading(cell, fill_hex)
    set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
    set_cell_border(cell,
        top=dict(val='single', sz=4, color=BORDER_HEX),
        bottom=dict(val='single', sz=4, color=BORDER_HEX),
        left=dict(val='single', sz=4, color=BORDER_HEX),
        right=dict(val='single', sz=4, color=BORDER_HEX)
    )
    
    p_in = cell.paragraphs[0]
    set_para_rtl(p_in, WD_ALIGN_PARAGRAPH.RIGHT)
    p_in.paragraph_format.space_before = Pt(0)
    p_in.paragraph_format.space_after = Pt(0)
    p_in.paragraph_format.line_spacing = 1.6
    
    # Empty lines for typing
    for i in range(lines - 1):
        add_persian_run(p_in, "\n", size_pt=9)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(2)

def build_docx():
    doc = docx.Document()
    
    # Page setup: A4
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Mm(14)
        section.bottom_margin = Mm(14)
        section.left_margin = Mm(18)
        section.right_margin = Mm(18)
        section.header_distance = Mm(8)
        section.footer_distance = Mm(8)
        
        # Footer
        footer = section.footer
        p_f = footer.paragraphs[0]
        set_para_rtl(p_f, WD_ALIGN_PARAGRAPH.RIGHT)
        add_persian_run(p_f, "مرکز مشاوره، روانشناسی و سلامت روان  |  گزارش هفتگی گروه‌ها  |  محرمانه", size_pt=8, color_rgb=MUTED_RGB)
    
    # ==================== PAGE 1 ====================
    # Header Image / Logo
    p_logo = doc.add_paragraph()
    set_para_rtl(p_logo, WD_ALIGN_PARAGRAPH.CENTER)
    p_logo.paragraph_format.space_before = Pt(0)
    p_logo.paragraph_format.space_after = Pt(2)
    
    logo_path = "assets/logo_clean_transparent.png"
    if os.path.exists(logo_path):
        run_logo = p_logo.add_run()
        run_logo.add_picture(logo_path, width=Mm(72))
    
    # Subtitle
    p_sub = doc.add_paragraph()
    set_para_rtl(p_sub, WD_ALIGN_PARAGRAPH.CENTER)
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(1)
    add_persian_run(p_sub, "مرکز مشاوره، روانشناسی و سلامت روان", font_name="Vazirmatn", size_pt=10.5, bold=False, color_rgb=SAGE_RGB)
    
    # Main Title
    p_title = doc.add_paragraph()
    set_para_rtl(p_title, WD_ALIGN_PARAGRAPH.CENTER)
    p_title.paragraph_format.space_before = Pt(1)
    p_title.paragraph_format.space_after = Pt(2)
    add_persian_run(p_title, "گزارش هفتگی گروه‌ها", font_name="Vazirmatn", size_pt=18.5, bold=True, color_rgb=FOREST_RGB)
    
    # Form Subtitle
    p_desc = doc.add_paragraph()
    set_para_rtl(p_desc, WD_ALIGN_PARAGRAPH.CENTER)
    p_desc.paragraph_format.space_before = Pt(0)
    p_desc.paragraph_format.space_after = Pt(3)
    add_persian_run(p_desc, "فرم اختصاصی ارزیابی هفتگی فعالیت‌ها، پایش سلامت تیمی و برنامه‌ریزی یکپارچه", font_name="Vazirmatn", size_pt=8.5, color_rgb=MUTED_RGB)
    
    # Divider Image
    div_path = "assets/botanical_divider_transparent.png"
    if os.path.exists(div_path):
        p_div = doc.add_paragraph()
        set_para_rtl(p_div, WD_ALIGN_PARAGRAPH.CENTER)
        p_div.paragraph_format.space_before = Pt(0)
        p_div.paragraph_format.space_after = Pt(6)
        r_div = p_div.add_run()
        r_div.add_picture(div_path, width=Mm(95))
    
    # Section 1: مشخصات گروه
    create_section_header(doc, "۱", "مشخصات گروه")
    
    tbl_meta = doc.add_table(rows=2, cols=4)
    set_table_rtl(tbl_meta)
    tbl_meta.autofit = False
    
    meta_structure = [
        [("نام گروه:", Mm(28)), (".......................................................", Mm(59)),
         ("نام سرگروه:", Mm(28)), (".......................................................", Mm(59))],
        [("بازه زمانی گزارش:", Mm(28)), ("از ................... تا ...................", Mm(59)),
         ("تاریخ ارائه:", Mm(28)), ("۱۴۰... / ...... / ......", Mm(59))]
    ]
    
    for r_idx, row_data in enumerate(meta_structure):
        for c_idx, (text, width) in enumerate(row_data):
            cell = tbl_meta.cell(r_idx, c_idx)
            cell.width = width
            set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
            p = cell.paragraphs[0]
            set_para_rtl(p, WD_ALIGN_PARAGRAPH.RIGHT)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            
            if c_idx in (0, 2):
                set_cell_shading(cell, BEIGE_HEX)
                add_persian_run(p, text, font_name="Vazirmatn", size_pt=9, bold=True, color_rgb=FOREST_RGB)
                set_cell_border(cell,
                    top=dict(val='single', sz=4, color=BORDER_HEX),
                    bottom=dict(val='single', sz=4, color=BORDER_HEX),
                    left=dict(val='single', sz=4, color=BORDER_HEX),
                    right=dict(val='single', sz=4, color=BORDER_HEX)
                )
            else:
                set_cell_shading(cell, IVORY_HEX)
                add_persian_run(p, text, font_name="Vazirmatn", size_pt=9, color_rgb=MUTED_RGB)
                set_cell_border(cell,
                    top=dict(val='single', sz=4, color=BORDER_HEX),
                    bottom=dict(val='single', sz=4, color=BORDER_HEX),
                    left=dict(val='single', sz=4, color=BORDER_HEX),
                    right=dict(val='single', sz=4, color=BORDER_HEX)
                )
    
    # Section 2: گزارش عملکرد سرگروه
    create_section_header(doc, "۲", "گزارش عملکرد سرگروه")
    
    add_writing_box(doc, "اقدامات و فعالیت‌های انجام‌شده:", lines=5)
    add_writing_box(doc, "برنامه‌ها و ایده‌های پیشنهادی:", lines=4)
    add_writing_box(doc, "دستاوردها و نتایج:", lines=3)
    
    # ==================== PAGE 2 ====================
    doc.add_page_break()
    
    # Continuation Header
    p_ch2 = doc.add_paragraph()
    set_para_rtl(p_ch2, WD_ALIGN_PARAGRAPH.RIGHT)
    p_ch2.paragraph_format.space_before = Pt(0)
    p_ch2.paragraph_format.space_after = Pt(4)
    add_persian_run(p_ch2, "گزارش هفتگی گروه‌ها  |  بخش ۳: گزارش اعضای گروه", font_name="Vazirmatn", size_pt=11.5, bold=True, color_rgb=FOREST_RGB)
    
    # Section 3: گزارش اعضای گروه
    create_section_header(doc, "۳", "گزارش اعضای گروه")
    
    # Helper hint banner
    p_hint = doc.add_paragraph()
    set_para_rtl(p_hint, WD_ALIGN_PARAGRAPH.RIGHT)
    p_hint.paragraph_format.space_before = Pt(1)
    p_hint.paragraph_format.space_after = Pt(4)
    add_persian_run(p_hint, "راهنما: ", font_name="Vazirmatn", size_pt=8.5, bold=True, color_rgb=FOREST_RGB)
    add_persian_run(p_hint, "این بخش به ازای هر یک از اعضای گروه تکمیل می‌گردد. جهت ثبت اعضای بیشتر، جدول زیر را انتخاب نموده و کپی (Copy/Paste) فرمایید.", font_name="Vazirmatn", size_pt=8.5, color_rgb=MUTED_RGB)
    
    def add_member_block(member_number):
        tbl_m = doc.add_table(rows=4, cols=2)
        set_table_rtl(tbl_m)
        tbl_m.autofit = False
        
        # Header row: Name
        cell_m_lbl = tbl_m.cell(0, 0)
        cell_m_lbl.width = Mm(36)
        set_cell_shading(cell_m_lbl, BEIGE_HEX)
        set_cell_margins(cell_m_lbl, top=120, bottom=120, left=140, right=140)
        set_cell_border(cell_m_lbl,
            top=dict(val='single', sz=6, color=FOREST_HEX),
            bottom=dict(val='single', sz=4, color=BORDER_HEX),
            left=dict(val='single', sz=4, color=BORDER_HEX),
            right=dict(val='single', sz=6, color=FOREST_HEX)
        )
        p0 = cell_m_lbl.paragraphs[0]
        set_para_rtl(p0, WD_ALIGN_PARAGRAPH.RIGHT)
        add_persian_run(p0, f"نام عضو (شماره {member_number}):", font_name="Vazirmatn", size_pt=9.5, bold=True, color_rgb=FOREST_RGB)
        
        cell_m_val = tbl_m.cell(0, 1)
        cell_m_val.width = Mm(138)
        set_cell_shading(cell_m_val, IVORY_HEX)
        set_cell_margins(cell_m_val, top=120, bottom=120, left=140, right=140)
        set_cell_border(cell_m_val,
            top=dict(val='single', sz=6, color=FOREST_HEX),
            bottom=dict(val='single', sz=4, color=BORDER_HEX),
            left=dict(val='single', sz=4, color=BORDER_HEX),
            right=dict(val='single', sz=4, color=BORDER_HEX)
        )
        p1 = cell_m_val.paragraphs[0]
        set_para_rtl(p1, WD_ALIGN_PARAGRAPH.RIGHT)
        add_persian_run(p1, "....................................................................................", size_pt=9.5, color_rgb=MUTED_RGB)
        
        # Rows: Fields
        fields = [
            ("شرح فعالیت‌ها و اقدامات انجام‌شده:", 4),
            ("نیازها و درخواست‌های حمایتی:", 3),
            ("پیشنهادها و برنامه‌های پیش‌رو:", 3)
        ]
        
        for idx, (f_label, lines_cnt) in enumerate(fields, start=1):
            c_lbl = tbl_m.cell(idx, 0)
            c_lbl.width = Mm(36)
            set_cell_shading(c_lbl, BEIGE_HEX)
            set_cell_margins(c_lbl, top=100, bottom=100, left=120, right=120)
            set_cell_border(c_lbl,
                top=dict(val='single', sz=4, color=BORDER_HEX),
                bottom=dict(val='single', sz=4, color=BORDER_HEX),
                left=dict(val='single', sz=4, color=BORDER_HEX),
                right=dict(val='single', sz=6, color=FOREST_HEX)
            )
            p_l = c_lbl.paragraphs[0]
            set_para_rtl(p_l, WD_ALIGN_PARAGRAPH.RIGHT)
            add_persian_run(p_l, f_label, font_name="Vazirmatn", size_pt=8.5, bold=True, color_rgb=FOREST_RGB)
            
            c_val = tbl_m.cell(idx, 1)
            c_val.width = Mm(138)
            set_cell_shading(c_val, IVORY_HEX)
            set_cell_margins(c_val, top=100, bottom=100, left=120, right=120)
            set_cell_border(c_val,
                top=dict(val='single', sz=4, color=BORDER_HEX),
                bottom=dict(val='single', sz=4, color=BORDER_HEX),
                left=dict(val='single', sz=4, color=BORDER_HEX),
                right=dict(val='single', sz=4, color=BORDER_HEX)
            )
            p_v = c_val.paragraphs[0]
            set_para_rtl(p_v, WD_ALIGN_PARAGRAPH.RIGHT)
            p_v.paragraph_format.line_spacing = 1.6
            for _ in range(lines_cnt - 1):
                add_persian_run(p_v, "\n", size_pt=8.5)
        
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_before = Pt(2)
        p_sp.paragraph_format.space_after = Pt(2)
    
    add_member_block("۱")
    add_member_block("۲")
    
    # ==================== PAGE 3 ====================
    doc.add_page_break()
    
    # Continuation Header
    p_ch3 = doc.add_paragraph()
    set_para_rtl(p_ch3, WD_ALIGN_PARAGRAPH.RIGHT)
    p_ch3.paragraph_format.space_before = Pt(0)
    p_ch3.paragraph_format.space_after = Pt(4)
    add_persian_run(p_ch3, "گزارش هفتگی گروه‌ها  |  بخش ۴ و ۵: چالش‌ها، برنامه آتی و تأییدیه", font_name="Vazirmatn", size_pt=11.5, bold=True, color_rgb=FOREST_RGB)
    
    # Section 4: چالش‌ها و موانع
    create_section_header(doc, "۴", "چالش‌ها و موانع")
    
    add_writing_box(doc, "مشکلات و موانع موجود:", lines=3)
    add_writing_box(doc, "موارد نیازمند پیگیری:", lines=3)
    add_writing_box(doc, "راهکارهای پیشنهادی:", lines=3)
    
    # Section 5: برنامه هفته آینده
    create_section_header(doc, "۵", "برنامه هفته آینده")
    
    tbl_plan = doc.add_table(rows=6, cols=4)
    set_table_rtl(tbl_plan)
    tbl_plan.autofit = False
    
    headers = [("ردیف", Mm(14)), ("اقدامات پیشنهادی", Mm(84)), ("مسئول انجام", Mm(42)), ("زمان‌بندی", Mm(34))]
    for c_idx, (h_text, width) in enumerate(headers):
        cell = tbl_plan.cell(0, c_idx)
        cell.width = width
        set_cell_shading(cell, BEIGE_HEX)
        set_cell_margins(cell, top=120, bottom=120, left=100, right=100)
        set_cell_border(cell,
            top=dict(val='single', sz=6, color=FOREST_HEX),
            bottom=dict(val='single', sz=6, color=FOREST_HEX),
            left=dict(val='single', sz=4, color=BORDER_HEX),
            right=dict(val='single', sz=4, color=BORDER_HEX)
        )
        p = cell.paragraphs[0]
        set_para_rtl(p, WD_ALIGN_PARAGRAPH.CENTER if c_idx in (0, 3) else WD_ALIGN_PARAGRAPH.RIGHT)
        add_persian_run(p, h_text, font_name="Vazirmatn", size_pt=9, bold=True, color_rgb=FOREST_RGB)
    
    row_numbers = ["۱", "۲", "۳", "۴", "۵"]
    for r_idx in range(1, 6):
        num_str = row_numbers[r_idx - 1]
        for c_idx, (_, width) in enumerate(headers):
            cell = tbl_plan.cell(r_idx, c_idx)
            cell.width = width
            set_cell_shading(cell, IVORY_HEX if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
            set_cell_border(cell,
                top=dict(val='single', sz=4, color=BORDER_HEX),
                bottom=dict(val='single', sz=4, color=BORDER_HEX),
                left=dict(val='single', sz=4, color=BORDER_HEX),
                right=dict(val='single', sz=4, color=BORDER_HEX)
            )
            p = cell.paragraphs[0]
            set_para_rtl(p, WD_ALIGN_PARAGRAPH.CENTER if c_idx in (0, 3) else WD_ALIGN_PARAGRAPH.RIGHT)
            if c_idx == 0:
                add_persian_run(p, num_str, font_name="Vazirmatn", size_pt=9, bold=True, color_rgb=FOREST_RGB)
    
    # Signatures Table
    p_sign_h = doc.add_paragraph()
    p_sign_h.paragraph_format.space_before = Pt(8)
    p_sign_h.paragraph_format.space_after = Pt(2)
    
    tbl_sign = doc.add_table(rows=1, cols=2)
    set_table_rtl(tbl_sign)
    tbl_sign.autofit = False
    
    # Leader Sign Box
    cell_s1 = tbl_sign.cell(0, 0)
    cell_s1.width = Mm(86)
    set_cell_shading(cell_s1, IVORY_HEX)
    set_cell_margins(cell_s1, top=140, bottom=140, left=160, right=160)
    set_cell_border(cell_s1,
        top=dict(val='single', sz=4, color=BORDER_HEX),
        bottom=dict(val='single', sz=4, color=BORDER_HEX),
        left=dict(val='single', sz=4, color=BORDER_HEX),
        right=dict(val='single', sz=4, color=BORDER_HEX)
    )
    p_s1 = cell_s1.paragraphs[0]
    set_para_rtl(p_s1, WD_ALIGN_PARAGRAPH.RIGHT)
    add_persian_run(p_s1, "امضا و تأیید سرگروه\n", font_name="Vazirmatn", size_pt=9.5, bold=True, color_rgb=FOREST_RGB)
    add_persian_run(p_s1, "صحت مطالب و گزارش‌های ثبت‌شده فوق مورد تأیید است.\n\n", font_name="Vazirmatn", size_pt=8, color_rgb=MUTED_RGB)
    add_persian_run(p_s1, "ملاحظات تکمیلی سرگروه: ........................................................\n\n", font_name="Vazirmatn", size_pt=8, color_rgb=MUTED_RGB)
    add_persian_run(p_s1, "نام و نام خانوادگی: .............................  امضا و تاریخ: .............................", font_name="Vazirmatn", size_pt=8, color_rgb=DARK_TEXT_RGB)
    
    # Supervisor Sign Box
    cell_s2 = tbl_sign.cell(0, 1)
    cell_s2.width = Mm(86)
    set_cell_shading(cell_s2, IVORY_HEX)
    set_cell_margins(cell_s2, top=140, bottom=140, left=160, right=160)
    set_cell_border(cell_s2,
        top=dict(val='single', sz=4, color=BORDER_HEX),
        bottom=dict(val='single', sz=4, color=BORDER_HEX),
        left=dict(val='single', sz=4, color=BORDER_HEX),
        right=dict(val='single', sz=4, color=BORDER_HEX)
    )
    p_s2 = cell_s2.paragraphs[0]
    set_para_rtl(p_s2, WD_ALIGN_PARAGRAPH.RIGHT)
    add_persian_run(p_s2, "تأییدیه و بازخورد سوپروایزر / مدیریت مرکز\n", font_name="Vazirmatn", size_pt=9.5, bold=True, color_rgb=FOREST_RGB)
    add_persian_run(p_s2, "ملاحظات نظارتی، ارزیابی کیفی و تأمین منابع مورد تأیید است.\n\n", font_name="Vazirmatn", size_pt=8, color_rgb=MUTED_RGB)
    add_persian_run(p_s2, "دستور اقدامات نظارتی: ................................................................\n\n", font_name="Vazirmatn", size_pt=8, color_rgb=MUTED_RGB)
    add_persian_run(p_s2, "نام مسئول: .............................  امضا و مهر مرکز: .............................", font_name="Vazirmatn", size_pt=8, color_rgb=DARK_TEXT_RGB)
    
    output_docx = "weekly_report_psychology.docx"
    doc.save(output_docx)
    print(f"Saved {output_docx} successfully!")

if __name__ == "__main__":
    build_docx()
