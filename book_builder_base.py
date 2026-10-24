# -*- coding: utf-8 -*-
"""
build_book_100p.py: Master Generator for 100-Page Telecom Job Preparation Guide Book.
Comprehensive, mathematically rigorous, publication-grade, fully indexed.
"""
import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def create_telecom_book_100p():
    doc = docx.Document()

    # Page Margins (1 inch all sides)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Color Palette
    COLOR_PRIMARY = RGBColor(10, 37, 64)       # Deep Navy
    COLOR_SECONDARY = RGBColor(32, 96, 149)   # Slate Blue
    COLOR_TEXT = RGBColor(45, 55, 72)          # Charcoal Dark
    COLOR_MUTED = RGBColor(100, 110, 125)      # Muted Gray
    COLOR_LINK = RGBColor(0, 85, 165)          # Hyperlink Blue

    bm_counter = [1]

    def set_font(run, font_name="Calibri", size_pt=11, bold=False, italic=False, color=COLOR_TEXT):
        run.font.name = font_name
        run.font.size = Pt(size_pt)
        run.bold = bold
        run.italic = italic
        run.font.color.rgb = color

    def attach_bookmark(paragraph, bookmark_name):
        bm_id = bm_counter[0]
        bm_counter[0] += 1
        p = paragraph._p
        start = parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="{bm_id}" w:name="{bookmark_name}"/>')
        end = parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="{bm_id}"/>')
        p.insert(0, start)
        p.append(end)

    def add_h1(text, bookmark_name=None):
        h = doc.add_heading(level=1)
        h.paragraph_format.space_before = Pt(18)
        h.paragraph_format.space_after = Pt(8)
        h.paragraph_format.keep_with_next = True
        run = h.add_run(text)
        set_font(run, size_pt=16, bold=True, color=COLOR_PRIMARY)
        pPr = h._p.get_or_add_pPr()
        pPr.append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="0"/>'))
        if bookmark_name:
            attach_bookmark(h, bookmark_name)
        return h

    def add_h2(text, bookmark_name=None):
        h = doc.add_heading(level=2)
        h.paragraph_format.space_before = Pt(14)
        h.paragraph_format.space_after = Pt(6)
        h.paragraph_format.keep_with_next = True
        run = h.add_run(text)
        set_font(run, size_pt=13, bold=True, color=COLOR_SECONDARY)
        pPr = h._p.get_or_add_pPr()
        pPr.append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="1"/>'))
        if bookmark_name:
            attach_bookmark(h, bookmark_name)
        return h

    def add_h3(text, bookmark_name=None):
        h = doc.add_heading(level=3)
        h.paragraph_format.space_before = Pt(10)
        h.paragraph_format.space_after = Pt(4)
        h.paragraph_format.keep_with_next = True
        run = h.add_run(text)
        set_font(run, size_pt=11.5, bold=True, color=COLOR_PRIMARY)
        pPr = h._p.get_or_add_pPr()
        pPr.append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="2"/>'))
        if bookmark_name:
            attach_bookmark(h, bookmark_name)
        return h

    def add_body(text, space_after=6, bold=False, italic=False):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(text)
        set_font(run, size_pt=10.5, bold=bold, italic=italic, color=COLOR_TEXT)
        return p

    def add_bullet(bold_prefix, text):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        r1 = p.add_run(bold_prefix + " ")
        set_font(r1, size_pt=10.5, bold=True, color=COLOR_PRIMARY)
        r2 = p.add_run(text)
        set_font(r2, size_pt=10.5, color=COLOR_TEXT)
        return p

    def add_callout(title, text):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)
        tcPr = cell._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F4F7FA"/>'))
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>\n'
            f'  <w:top w:val="none"/>\n'
            f'  <w:left w:val="single" w:sz="24" w:space="0" w:color="0A2540"/>\n'
            f'  <w:bottom w:val="none"/>\n'
            f'  <w:right w:val="none"/>\n'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        r_title = p.add_run("💡 " + title + "\n")
        set_font(r_title, size_pt=10.5, bold=True, color=COLOR_PRIMARY)
        r_body = p.add_run(text)
        set_font(r_body, size_pt=10, italic=False, color=COLOR_TEXT)
        empty_p = doc.add_paragraph()
        empty_p.paragraph_format.space_before = Pt(0)
        empty_p.paragraph_format.space_after = Pt(4)

    def add_formula(title, formula_text, breakdown="", intuition=""):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)
        tcPr = cell._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="EDF2F7"/>'))
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>\n'
            f'  <w:top w:val="single" w:sz="8" w:space="0" w:color="CBD5E0"/>\n'
            f'  <w:left w:val="single" w:sz="20" w:space="0" w:color="206095"/>\n'
            f'  <w:bottom w:val="single" w:sz="8" w:space="0" w:color="CBD5E0"/>\n'
            f'  <w:right w:val="single" w:sz="8" w:space="0" w:color="CBD5E0"/>\n'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        r1 = p.add_run("📐 সূত্র: " + title + "\n")
        set_font(r1, size_pt=11, bold=True, color=COLOR_SECONDARY)
        r2 = p.add_run(formula_text + "\n")
        set_font(r2, font_name="Consolas", size_pt=11.5, bold=True, color=COLOR_PRIMARY)
        if breakdown:
            r3 = p.add_run("🔍 সূত্রের প্যারামিটার পরিচিতি:\n" + breakdown + "\n")
            set_font(r3, size_pt=9.5, color=COLOR_TEXT)
        if intuition:
            r4 = p.add_run("💡 সহজবোধ্য ব্যাখ্যা ও তাৎপর্য:\n" + intuition)
            set_font(r4, size_pt=9.5, italic=True, color=COLOR_PRIMARY)
        empty_p = doc.add_paragraph()
        empty_p.paragraph_format.space_before = Pt(0)
        empty_p.paragraph_format.space_after = Pt(4)

    def add_math_solution(prob_title, given_data, formula_used, step_by_step, tips=""):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)
        tcPr = cell._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F8FAFC"/>'))
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>\n'
            f'  <w:top w:val="single" w:sz="10" w:space="0" w:color="0284C7"/>\n'
            f'  <w:left w:val="single" w:sz="20" w:space="0" w:color="0284C7"/>\n'
            f'  <w:bottom w:val="single" w:sz="10" w:space="0" w:color="0284C7"/>\n'
            f'  <w:right w:val="single" w:sz="10" w:space="0" w:color="0284C7"/>\n'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)

        r_head = p.add_run("📝 গাণিতিক সমাধান: " + prob_title + "\n")
        set_font(r_head, size_pt=10.5, bold=True, color=COLOR_PRIMARY)

        r_given = p.add_run("📋 দেওয়া উপাত্ত (Given Data):\n" + given_data + "\n\n")
        set_font(r_given, size_pt=9.5, color=COLOR_TEXT)

        r_form = p.add_run("📐 প্রযোজ্য সূত্র (Applicable Formula):\n" + formula_used + "\n\n")
        set_font(r_form, font_name="Consolas", size_pt=10, bold=True, color=COLOR_SECONDARY)

        r_steps = p.add_run("🔍 পর্যায়ক্রমিক গণনা (Step-by-Step Calculation):\n" + step_by_step)
        set_font(r_steps, size_pt=9.5, color=COLOR_TEXT)

        if tips:
            r_tips = p.add_run("\n\n⚠️ পরীক্ষার টিপস ও সাধারণ ভুল এড়ানোর উপায়:\n" + tips)
            set_font(r_tips, size_pt=9, bold=True, italic=True, color=RGBColor(180, 50, 50))

        empty_p = doc.add_paragraph()
        empty_p.paragraph_format.space_before = Pt(0)
        empty_p.paragraph_format.space_after = Pt(4)

    def add_img(img_path, caption, width_in=5.8):
        if os.path.exists(img_path):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run()
            run.add_picture(img_path, width=Inches(width_in))
            cp = doc.add_paragraph()
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cp.paragraph_format.space_before = Pt(0)
            cp.paragraph_format.space_after = Pt(8)
            c_run = cp.add_run(f"চিত্র: {caption}")
            set_font(c_run, size_pt=9.5, bold=True, italic=True, color=COLOR_MUTED)

    def add_back_to_toc():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(12)
        hyperlink = parse_xml(
            f'<w:hyperlink {nsdecls("w")} w:anchor="bookmark_toc" w:history="1">'
            f'<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:color w:val="0055A5"/><w:sz w:val="18"/><w:u w:val="single"/></w:rPr>'
            f'<w:t>⬆️ সূচীপত্রে ফিরে যান (Back to TOC)</w:t></w:r></w:hyperlink>'
        )
        p._p.append(hyperlink)

    def add_toc_link(p, display_text, bookmark_target, is_bold=False, level=1):
        if level == 1:
            indent = ""
            color_code = "0A2540"
            size_val = "22"
        elif level == 2:
            indent = "    • "
            color_code = "206095"
            size_val = "20"
        else:
            indent = "        - "
            color_code = "4A5568"
            size_val = "19"

        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        if indent:
            r_ind = p.add_run(indent)
            set_font(r_ind, size_pt=10, color=COLOR_MUTED)
        b_tag = '<w:b/>' if is_bold else ''
        safe_text = display_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        hyperlink = parse_xml(
            f'<w:hyperlink {nsdecls("w")} w:anchor="{bookmark_target}" w:history="1">'
            f'<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:color w:val="{color_code}"/>{b_tag}<w:sz w:val="{size_val}"/><w:u w:val="none"/></w:rPr>'
            f'<w:t>{safe_text}</w:t></w:r></w:hyperlink>'
        )
        p._p.append(hyperlink)

    print("Building 100-page expanded book...")
    return doc, {
        "set_font": set_font, "attach_bookmark": attach_bookmark,
        "add_h1": add_h1, "add_h2": add_h2, "add_h3": add_h3,
        "add_body": add_body, "add_bullet": add_bullet,
        "add_callout": add_callout, "add_formula": add_formula,
        "add_math_solution": add_math_solution, "add_img": add_img,
        "add_back_to_toc": add_back_to_toc, "add_toc_link": add_toc_link,
        "COLOR_PRIMARY": COLOR_PRIMARY, "COLOR_SECONDARY": COLOR_SECONDARY,
        "COLOR_TEXT": COLOR_TEXT, "COLOR_MUTED": COLOR_MUTED, "COLOR_LINK": COLOR_LINK
    }

if __name__ == "__main__":
    doc, helpers = create_telecom_book_100p()
    print("Base initialized successfully.")
