import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def create_telecom_book():
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

    # Bookmark counter
    bm_counter = [1]

    def set_font(run, font_name="Calibri", size_pt=11, bold=False, italic=False, color=COLOR_TEXT):
        run.font.name = font_name
        run.font.size = Pt(size_pt)
        run.bold = bold
        run.italic = italic
        run.font.color.rgb = color

    # Helper: Add Bookmark to paragraph
    def attach_bookmark(paragraph, bookmark_name):
        bm_id = bm_counter[0]
        bm_counter[0] += 1
        p = paragraph._p
        start = parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="{bm_id}" w:name="{bookmark_name}"/>')
        end = parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="{bm_id}"/>')
        p.insert(0, start)
        p.append(end)

    # Helper: Add Heading 1 (Word Navigation Pane Level 1)
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

    # Helper: Add Heading 2 (Word Navigation Pane Level 2)
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

    # Helper: Add Heading 3 (Word Navigation Pane Level 3)
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

    # Helper: Add Body Paragraph
    def add_body(text, space_after=6, bold=False, italic=False):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(text)
        set_font(run, size_pt=10.5, bold=bold, italic=italic, color=COLOR_TEXT)
        return p

    # Helper: Add Bullet
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

    # Helper: Add Callout Box / Note
    def add_callout(title, text):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)

        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F4F7FA"/>')
        tcPr.append(shd)
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

    # Helper: Add Formula Box with Intuitive Explanation
    def add_formula(title, formula_text, breakdown="", intuition=""):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)

        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="EDF2F7"/>')
        tcPr.append(shd)
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

    # Helper: Add Math Solution Box with Step-by-Step Breakdown
    def add_math_solution(prob_title, given_data, formula_used, step_by_step, tips=""):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)

        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F8FAFC"/>')
        tcPr.append(shd)
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

    # Helper: Add Image with Caption
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

    # Helper: Back to TOC button
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

    # Helper: Add Interactive Link in TOC
    def add_toc_link(p, display_text, bookmark_target, is_bold=False, level=1):
        if level == 1:
            indent = ""
            color_code = "0A2540"
            size_val = "22" # 11pt
        elif level == 2:
            indent = "    • "
            color_code = "206095"
            size_val = "20" # 10pt
        else:
            indent = "        - "
            color_code = "4A5568"
            size_val = "19" # 9.5pt

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

    # ==========================================
    # 1. BEAUTIFUL COVER PAGE
    # ==========================================
    p_tag = doc.add_paragraph()
    p_tag.paragraph_format.space_before = Pt(12)
    p_tag.paragraph_format.space_after = Pt(10)
    p_tag.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_tag = p_tag.add_run("গণপ্রজাতন্ত্রী বাংলাদেশ • টেলিযোগাযোগ ও আইসিটি খাতের প্রকৌশলী নিয়োগ সহায়িকা")
    set_font(r_tag, size_pt=9.5, bold=True, color=COLOR_SECONDARY)

    cover_p = doc.add_paragraph()
    cover_p.paragraph_format.space_before = Pt(4)
    cover_p.paragraph_format.space_after = Pt(4)
    cover_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = cover_p.add_run("টেলিকমিউনিকেশন ইঞ্জিনিয়ারিং\nজব প্রিপারেশন গাইড")
    set_font(r_title, size_pt=26, bold=True, color=COLOR_PRIMARY)

    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_before = Pt(2)
    sub_p.paragraph_format.space_after = Pt(12)
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = sub_p.add_run("A Comprehensive Preparation Handbook for BTRC, BTCL, Teletalk, BSCPLC, BSCL, MNOs & NTTN Engineers")
    set_font(r_sub, size_pt=10.5, italic=True, color=COLOR_SECONDARY)

    if os.path.exists("images/telecom_cover_art.png"):
        p_img = doc.add_paragraph()
        p_img.paragraph_format.space_before = Pt(4)
        p_img.paragraph_format.space_after = Pt(12)
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.add_run().add_picture("images/telecom_cover_art.png", width=Inches(5.0))

    tbl_badge = doc.add_table(rows=1, cols=1)
    tbl_badge.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_b = tbl_badge.cell(0, 0)
    c_b.width = Inches(6.5)
    tcPr_b = c_b._tc.get_or_add_tcPr()
    tcPr_b.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="0A2540"/>'))
    p_b = c_b.paragraphs[0]
    p_b.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_b.paragraph_format.space_before = Pt(5)
    p_b.paragraph_format.space_after = Pt(5)
    r_b = p_b.add_run("🎯 BTRC • BTCL • TELETALK • BSCPLC • BSCL • BCS (ICT) • MNOs (GP, Robi, BL) • NTTN • VENDORS")
    set_font(r_b, size_pt=9.5, bold=True, color=RGBColor(255, 255, 255))

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    tbl_feat = doc.add_table(rows=1, cols=1)
    tbl_feat.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_f = tbl_feat.cell(0, 0)
    c_f.width = Inches(6.5)
    tcPr_f = c_f._tc.get_or_add_tcPr()
    tcPr_f.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F4F7FA"/>'))
    tcPr_f.append(parse_xml(
        f'<w:tcBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="single" w:sz="8" w:space="0" w:color="CBD5E0"/>\n'
        f'  <w:left w:val="single" w:sz="18" w:space="0" w:color="0A2540"/>\n'
        f'  <w:bottom w:val="single" w:sz="8" w:space="0" w:color="CBD5E0"/>\n'
        f'  <w:right w:val="single" w:sz="8" w:space="0" w:color="CBD5E0"/>\n'
        f'</w:tcBorders>'
    ))
    p_f = c_f.paragraphs[0]
    p_f.paragraph_format.space_before = Pt(4)
    p_f.paragraph_format.space_after = Pt(4)
    r_f_head = p_f.add_run("📘 বইটির বিশেষ বৈশিষ্ট্যসমূহ:\n")
    set_font(r_f_head, size_pt=10, bold=True, color=COLOR_PRIMARY)
    r_f_body = p_f.add_run(
        "★ কোর থিওরি: সিগন্যাল, স্যাম্পলিং, এনালগ/ডিজিটাল মডুলেশন ও শ্যানন ক্যাপাসিটি\n"
        "★ সেলুলার ও মোবাইল: 2G/3G/4G LTE/5G NR, VoLTE আর্কিটেকচার ও Erlang ট্রাফিক\n"
        "★ অপটিক্যাল ফাইবার, DWDM, SDH (STM-1/64) ও লিংক পাওয়ার বাজেট ক্যালকুলেশন\n"
        "★ ডাটা নেটওয়ার্কিং: ওএসআই, টিসিপি/আইপি, সাবনেটিং (FLSM/VLSM), OSPF ও BGP\n"
        "★ বাংলাদেশ টেলিকম নেটওয়ার্ক: IGW, ICX, IIG, NTTN, সাবমেরিন ক্যাবল ও স্যাটেলাইট\n"
        "★ বিগত ১০ বছরের নিয়োগ পরীক্ষার লিখিত ও গাণিতিক প্রশ্ন সমাধান (BUET/MIST)\n"
        "★ সেরা ১০০ টেকনিক্যাল ভাইভা প্রশ্নোত্তর ও প্র্যাকটিক্যাল ট্রাবলশুটিং সিনারিও"
    )
    set_font(r_f_body, size_pt=8.5, color=COLOR_TEXT)

    meta_p = doc.add_paragraph()
    meta_p.paragraph_format.space_before = Pt(12)
    meta_p.paragraph_format.space_after = Pt(0)
    meta_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_meta = meta_p.add_run("২০২৬–২০২৭ পরিমার্জিত ও বর্ধিত সংস্করণ • ইঞ্জিনিয়ার্স ক্যারিয়ার প্রকাশনা")
    set_font(r_meta, size_pt=10, bold=True, color=COLOR_SECONDARY)

    doc.add_page_break()

    # ==========================================
    # 2. PROLOGUE & NECESSITY OF THE BOOK
    # ==========================================
    add_h1("প্রারম্ভিক কথা ও বইটির প্রয়োজনীয়তা", "sec_prologue")
    
    add_body("একুশ শতকের এই তথ্যপ্রযুক্তি ও চতুর্থ শিল্পবিপ্লবের (4IR) যুগে টেলিযোগাযোগ কেবল একটি প্রকৌশল শাখা নয়, বরং এটি একটি আধুনিক জাতির অর্থনৈতিক প্রবৃদ্ধি, জাতীয় নিরাপত্তা এবং ডিজিটাল সার্বভৌমত্বের মূল চালিকাশক্তি। বাংলাদেশ আজ ফাইভ-জি (5G), দেশব্যাপী অপটিক্যাল ফাইবার ব্যাকবোন (NTTN), আন্তর্জাতিক সাবমেরিন ক্যাবল কনসোর্টিয়াম (SEA-ME-WE-4, 5, 6) এবং নিজস্ব ভূস্থির উপগ্রহ 'বঙ্গবন্ধু স্যাটেলাইট-১' এর সফল পরিচালনায় এক অভাবনীয় ডিজিটাল উৎকর্ষের দ্বারপ্রান্তে।")

    add_body("এই প্রযুক্তিগত অগ্রযাত্রাকে নিরবচ্ছিন্ন রাখতে দেশের টেলিযোগাযোগ ও আইসিটি খাতে প্রতি বছর শত শত দক্ষ ও মেধাবী প্রকৌশলীর প্রয়োজন হচ্ছে। তবে এই খাতগুলোতে নিয়োগ পরীক্ষা অত্যন্ত প্রতিদ্বন্দ্বিতাপূর্ণ এবং এর প্রশ্ন প্যাটার্ন প্রথাগত অন্যান্য চাকরির চেয়ে অনেকটাই স্বতন্ত্র ও গভীর অ্যানালিটিক্যাল।")

    add_h2("টেলিকম জব প্রিপারেশনে এই বইটির প্রয়োজনীয়তা কেন অপরিহার্য?", "sec_why_essential")
    
    add_bullet("১. সিলেবাসের বিচ্ছিন্নতা নিরসন ও একমাত্র সমন্বিত ওয়ান-স্টপ সল্যুশন:",
        "বাজারে সাধারণ বিসিএস বা অন্যান্য ইঞ্জিনিয়ারিং শাখার জন্য শত শত গাইড বই পাওয়া গেলেও, টেলিকমিউনিকেশন ও নেটওয়ার্কিং খাতের সরকারি ও শীর্ষ কর্পোরেট চাকরির জন্য কোনো একক, পূর্ণাঙ্গ ও আধুনিক হ্যান্ডবুক এতদিন অনুপস্থিত ছিল। চাকুরিপার্থীদের বিভিন্ন বিশ্ববিদ্যালয় ও ইন্টারনেটের অগোছালো স্লাইড এবং খণ্ড খণ্ড বিদেশী বই ঘেঁটে প্রস্তুতি নিতে হতো। এই বইটি সেই দীর্ঘদিনের শূন্যতা দূর করে প্রথম পূর্ণাঙ্গ ওয়ান-স্টপ সল্যুশন প্রদান করে।"
    )
    add_bullet("২. আন্তর্জাতিক টেক্সটবুক বনাম দেশীয় নিয়োগ পরীক্ষার মেলবন্ধন:",
        "Theodore Rappaport-এর ওয়্যারলেস কমিউনিকেশন, Behrouz Forouzan-এর ডাটা নেটওয়ার্কিং কিংবা Gerd Keiser-এর অপটিক্যাল ফাইবার কমিউনিকেশন বিশ্বসেরা বই হলেও, দেশীয় নিয়োগ পরীক্ষার (বুয়েট ও এমআইএসটি প্যাটার্ন) প্রেক্ষাপটে কোন টপিকগুলো সবচেয়ে বেশি আসে এবং কীভাবে সংক্ষেপে নির্ভুল উত্তর উপস্থাপন করতে হয়—তা সাধারণ টেক্সটবুকে থাকে না। এই বইটিতে আন্তর্জাতিক স্ট্যান্ডার্ড তত্ত্বের সাথে দেশীয় পরীক্ষার সরাসরি মেলবন্ধন ঘটানো হয়েছে।"
    )
    add_bullet("৩. বাংলাদেশ টেলিকম ইকোসিস্টেমের এক্সক্লুসিভ জ্ঞান:",
        "বিদেশী কোনো বইয়ে বাংলাদেশের অভ্যন্তরীণ টেলিকম কাঠামোর লাইসেন্সিং ফ্রেমওয়ার্ক যেমন: IGW (ভয়েস), ICX (ট্রানজিট), IIG (ইন্টারনেট), NTTN (ফাইবার ব্যাকবোন), ANS (অপারেটর), সাবমেরিন ক্যাবল স্টেশন কিংবা বিটিআরসি (BTRC) নীতিমালা নিয়ে কোনো তথ্য থাকে না। অথচ বিটিআরসি, বিটিসিএল ও দেশীয় পরীক্ষায় এখান থেকে ২৫% থেকে ৩০% প্রশ্ন আসে। এই বইটিতে এই কাঠামোটি বিস্তারিত ডায়াগ্রামসহ উপস্থাপন করা হয়েছে।"
    )
    add_bullet("৪. গাণিতিক দক্ষতা ও নির্ভুল ক্যালকুলেশন হ্যান্ডবুক:",
        "লিখিত পরীক্ষায় প্রার্থীদের মধ্যে নম্বরের মূল পার্থক্য তৈরি করে নির্ভুল গাণিতিক সমাধান। শ্যানন ক্যাপাসিটি, নাইকুইস্ট রেট, পিসিএম বিট রেট ও SQNR, এরল্যাং ট্রাফিক ও ক্লাস্টার সাইজ, অপটিক্যাল লিংক পাওয়ার বাজেট, ফ্রেনেল জোন ব্যাসার্ধ এবং আইপি সাবনেটিং (FLSM/VLSM)—প্রতিটি বিষয়ের টাইপভিত্তিক অঙ্ক ও শর্টকাট কৌশল স্টেপ-বাই-স্টেপ সমাধান করে দেওয়া হয়েছে।"
    )
    add_bullet("৫. টেকনিক্যাল ভাইভা ও প্র্যাকটিক্যাল ফিল্ড নলেজ:",
        "কেবল থিওরি মুখস্থ করে বড় বড় টেলিকম অপারেটর বা বিটিআরসি-র ভাইভা বোর্ডে উত্তীর্ণ হওয়া সম্ভব নয়। এই বইটিতে ফিল্ড ইঞ্জিনিয়ারিংয়ের বাস্তব অভিজ্ঞতা যেমন: কল ড্রপের কারণ ও ড্রাইভ টেস্ট, OTDR দিয়ে ফাইবার কাট পয়েন্ট শনাক্তকরণ, বিটিএস টাওয়ারের -৪৮V ডিসি পাওয়ার প্ল্যান্ট ও ব্যাটারি ব্যাকআপ ম্যানেজমেন্ট এবং সেরা ১০০টি টেকনিক্যাল ভাইভা প্রশ্নোত্তর অন্তর্ভুক্ত রয়েছে।"
    )

    add_h2("কাদের জন্য এই বইটি রচিত? (Target Audience)", "sec_target_audience")
    add_bullet("• অ্যাকাডেমিক ব্যাকগ্রাউন্ড:", "EEE, ETE, ECE, CSE, ICE এবং সংশ্লিষ্ট ডিসিপ্লিনের বিএসসি ও ডিপ্লোমা ইঞ্জিনিয়ারিং ডিগ্রিধারী শিক্ষার্থী ও চাকুরিপ্রার্থী।")
    add_bullet("• সরকারি ও স্বায়ত্তশাসিত সংস্থা:", "বাংলাদেশ টেলিযোগাযোগ নিয়ন্ত্রণ কমিশন (BTRC), বাংলাদেশ টেলিকমিউনিকেশনস কোম্পানি লিমিটেড (BTCL), টেলিটক বাংলাদেশ লিমিটেড (Teletalk), বাংলাদেশ সাবমেরিন ক্যাবলস পিএলসি (BSCPLC), বাংলাদেশ স্যাটেলাইট কোম্পানি লিমিটেড (BSCL) এবং বিসিএস (টেলিকম ও তথ্য ও যোগাযোগ প্রযুক্তি ক্যাডার)।")
    add_bullet("• বেসরকারি কর্পোরেট সেক্টর:", "মোবাইল নেটওয়ার্ক অপারেটর (Grameenphone, Robi, Banglalink), নেশনওয়াইড ট্রান্সমিশন নেটওয়ার্ক (Fiber@Home, Summit Communications, Bahon), ইন্টারনেট সার্ভিস প্রোভাইডার (ISP) এবং আন্তর্জাতিক ভেন্ডর কোম্পানি (Huawei, Ericsson, ZTE, Nokia)।")

    add_h2("বইটি সর্বোচ্চ ফলপ্রসূ উপায়ে পড়ার কৌশল (Study Strategy)", "sec_study_strategy")
    add_body("বইটি অধ্যয়নের সময় প্রথমে প্রতিটি অধ্যায়ের মৌলিক থিওরি ও ব্লক ডায়াগ্রাম ভালোভাবে অনুধাবন করুন। এরপর প্রতিটি অধ্যায়ের সাথে সংযুক্ত ফর্মুলা বক্সগুলো আলাদা খাতায় নোট করে গাণিতিক সমস্যাগুলো নিজ হাতে সমাধান করার অভ্যাস গড়ে তুলুন। পরিশেষে বিগত বছরের প্রশ্নগুলো নিজে নিজে পরীক্ষা দেওয়ার মতো করে সমাধান করুন এবং ভাইভা অংশের প্রশ্নগুলো নিয়মিত চর্চা করুন।")

    add_back_to_toc()
    doc.add_page_break()

    # ==========================================
    # 3. FOREWORD & EXAM STRATEGY
    # ==========================================
    add_h1("ভূমিকা ও নিয়োগ পরীক্ষার প্রস্তুতি কৌশল", "sec_foreword")
    add_body("বাংলাদেশে টেলিকমিউনিকেশন ও আইসিটি খাতের প্রকৌশলীদের কর্মক্ষেত্র অত্যন্ত বিস্তৃত ও আকর্ষণীয়। বাংলাদেশ টেলিযোগাযোগ নিয়ন্ত্রণ কমিশন (BTRC), বাংলাদেশ টেলিকমিউনিকেশনস কোম্পানি লিমিটেড (BTCL), টেলিটক বাংলাদেশ লিমিটেড (Teletalk), বাংলাদেশ সাবমেরিন ক্যাবলস পিএলসি (BSCPLC), বাংলাদেশ স্যাটেলাইট কোম্পানি লিমিটেড (BSCL), বিসিএস (টেলিকম ও আইসিটি ক্যাডার), শীর্ষ বেসরকারি মোবাইল নেটওয়ার্ক অপারেটর (Grameenphone, Robi, Banglalink), নেশনওয়াইড ট্রান্সমিশন নেটওয়ার্ক (NTTN: Fiber@Home, Summit Communications) এবং মাল্টিন্যাশনাল ভেন্ডর (Huawei, Ericsson, ZTE) প্রতিষ্ঠানে প্রকৌশলী নিয়োগের জন্য অত্যন্ত প্রতিদ্বন্দ্বিতাপূর্ণ লিখিত ও মৌখিক পরীক্ষা অনুষ্ঠিত হয়।")

    add_h2("পরীক্ষার কাঠামো ও নম্বর বণ্টন (BUET / MIST Pattern)", "sec_exam_pattern")
    add_body("বিটিসিএল, বিটিআরসি ও টেলিটকের বেশিরভাগ সহকারী প্রকৌশলী ও সহকারী ব্যবস্থাপক পর্যায়ের পরীক্ষাগুলো সাধারণত বুয়েট (BUET) অথবা এমআইএসটি (MIST) কর্তৃক পরিচালিত হয়। পরীক্ষার সাধারণ কাঠামো নিচে তুলে ধরা হলো:")

    tbl = doc.add_table(rows=1, cols=4)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["অংশ", "বিষয়বস্তু", "নম্বর বণ্টন", "গুরুত্বপূর্ণ টপিকস"]
    widths = [Inches(1.2), Inches(1.8), Inches(1.1), Inches(2.4)]
    hdr_cells = tbl.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        hdr_cells[i].width = widths[i]
        tcPr = hdr_cells[i]._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="0A2540"/>'))
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            set_font(r, size_pt=10, bold=True, color=RGBColor(255, 255, 255))

    exam_data = [
        ("টেকনিক্যাল (কোর)", "কমিউনিকেশন, অপটিক্যাল ফাইবার, মোবাইল নেটওয়ার্ক, ডাটা কমিউনিকেশন", "৬০% – ৭০%", "PCM, Shannon capacity, Link budget, Subnetting, GSM architecture, 5G"),
        ("ইলেকট্রিক্যাল বেসিক", "সার্কিট থিওরেম, এসি ফান্ডামেন্টালস, ইলেকট্রনিক্স ও ডিসি পাওয়ার", "১৫% – ২০%", "KVL/KCL, Thevenin, Op-Amp, -48V DC, Battery bank calculations"),
        ("নন-টেকনিক্যাল", "বাংলা, ইংরেজি রচনা/অনুবাদ, সাধারণ জ্ঞান ও আইসিটি পলিসি", "১৫% – ২০%", "Cyber security, BTRC Licensing, Smart Bangladesh, Submarine cable"),
    ]

    for row_data in exam_data:
        row_cells = tbl.add_row().cells
        for i, val in enumerate(row_data):
            row_cells[i].text = val
            row_cells[i].width = widths[i]
            p = row_cells[i].paragraphs[0]
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.15
            for r in p.runs:
                set_font(r, size_pt=9.5, color=COLOR_TEXT)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    add_callout("সফলতার গোপন ট্রিকস", "লিখিত পরীক্ষায় ফুল মার্কস পাওয়ার অন্যতম চাবিকাঠি হলো কেবল বর্ণনামূলক উত্তর না লিখে স্পষ্ট ব্লক ডায়াগ্রাম, স্ট্যান্ডার্ড সমীকরণ এবং বাস্তবসম্মত সংখ্যাসহ গাণিতিক সমাধান উপস্থাপন করা।")

    add_back_to_toc()
    doc.add_page_break()

    # ==========================================
    # 4. INTERACTIVE NAVIGATION & TABLE OF CONTENTS
    # ==========================================
    h_toc = add_h1("বইয়ের পূর্ণাঙ্গ সূচীপত্র ও ইন্টারেক্টিভ নেভিগেশন", "bookmark_toc")
    
    nav_guide = (
        "১. কীবোর্ডে Ctrl + F চাপুন অথবা Word মেনু থেকে View ➔ Navigation Pane অপশনে টিক দিন। বাম পাশের সাইডবারে অধ্যায় ও টপিকের সম্পূর্ণ ট্রি দেখা যাবে এবং যেকোনো টপিকে এক ক্লিকেই জাম্প করা যাবে।\n"
        "২. নিচের সূচীপত্রের যেকোনো অধ্যায় বা টপিকের উপর Ctrl চেপে রেখে মাউস ক্লিক (Ctrl + Click) করলেই সরাসরি সেই অধ্যায় বা টপিকে চলে যাওয়া যাবে।\n"
        "৩. প্রতিটি অধ্যায়ের শেষে [⬆️ সূচীপত্রে ফিরে যান] লিংকে ক্লিক করে তাৎক্ষণিক আবার সূচীপত্রে ফিরে আসতে পারবেন।"
    )
    add_callout("সহজ নেভিগেশন নির্দেশিকা (Quick Navigation Guide)", nav_guide)

    toc_structure = [
        ("প্রারম্ভিক অংশ ও প্রস্তুতি কৌশল", None, [
            ("প্রারম্ভিক কথা ও বইটির প্রয়োজনীয়তা (Prologue & Necessity)", "sec_prologue"),
            ("ভূমিকা ও নিয়োগ পরীক্ষার প্রস্তুতি কৌশল (BUET/MIST Pattern)", "sec_foreword")
        ]),
        ("প্রথম খণ্ড: কোর টেলিকমিউনিকেশন ইঞ্জিনিয়ারিং", None, [
            ("অধ্যায় ০১: ফান্ডামেন্টালস অব কমিউনিকেশন সিস্টেম", "ch_01", [
                ("১.১ কমিউনিকেশন সিস্টেমের বেসিক মডেল ও উপাদান", "sec_1_1"),
                ("১.২ অ্যানালগ মডুলেশন থিওরি (AM, FM, PM)", "sec_1_2"),
                ("১.৩ স্যাম্পলিং থিওরেম ও শ্যানন ক্যাপাসিটি", "sec_1_3"),
                ("১.৪ পালস কোড মডুলেশন (PCM) ও কোয়ান্টাইজেশন", "sec_1_4"),
                ("১.৫ ডিজিটাল মডুলেশন টেকনিকস ও কনস্টেলেশন ডায়াগ্রাম", "sec_1_5"),
                ("১.৬ অধ্যায় ০১ এর বিগত পরীক্ষার প্রশ্ন ও গাণিতিক সমাধান", "sec_1_6")
            ]),
            ("অধ্যায় ০২: সেলুলার ও মোবাইল ওয়্যারলেস কমিউনিকেশন", "ch_02", [
                ("২.১ সেলুলার কনসেপ্ট ও ফ্রিকোয়েন্সি রিইউজ", "sec_2_1"),
                ("২.২ নেটওয়ার্ক ক্যাপাসিটি বৃদ্ধির কৌশল (Cell Splitting/Sectoring)", "sec_2_2"),
                ("২.৩ ট্রাফিক ইঞ্জিনিয়ারিং ও আরল্যাং (Erlang-B & Erlang-C)", "sec_2_3"),
                ("২.৪ মোবাইল প্রজন্মের রূপান্তর (2G থেকে 5G NR তুলনামূলক ছক)", "sec_2_4"),
                ("২.৫ ৪জি এলটিই (4G LTE) আর্কিটেকচার ও কোর নেটওয়ার্ক (EPC)", "sec_2_5")
            ]),
            ("অধ্যায় ০৩: অপটিক্যাল ফাইবার কমিউনিকেশন ও ট্রান্সমিশন নেটওয়ার্ক", "ch_03", [
                ("৩.১ অপটিক্যাল ফাইবারের মূলনীতি ও আলোক বিজ্ঞান (TIR, NA)", "sec_3_1"),
                ("৩.২ অপটিক্যাল সিগন্যাল লস ও ডিসপারশন (Attenuation & Dispersion)", "sec_3_2"),
                ("৩.৩ আধুনিক অপটিক্যাল নেটওয়ার্কিং (WDM, DWDM, CWDM, EDFA)", "sec_3_3"),
                ("৩.৪ ট্রান্সমিশন হায়ারার্কি: PDH বনাম SDH (STM-1/4/16/64)", "sec_3_4"),
                ("৩.৫ অপটিক্যাল লিংক পাওয়ার বাজেট ক্যালকুলেশন ও সমাধান", "sec_3_5")
            ]),
            ("অধ্যায় ০৪: ডাটা কমিউনিকেশন ও কম্পিউটার নেটওয়ার্কিং", "ch_04", [
                ("৪.১ ওএসআই (OSI) ৭-লেয়ার মডেল বনাম টিসিপি/আইপি (TCP/IP) স্ট্যাক", "sec_4_1"),
                ("৪.২ আইপি অ্যাড্রেসিং ও ক্লাসলেস নেটওয়ার্ক (IPv4 vs IPv6)", "sec_4_2"),
                ("৪.৩ আইপি সাবনেটিং ও ভিএলএসএম (FLSM vs VLSM) - শর্টকাট টেকনিক", "sec_4_3"),
                ("৪.৪ রাউটিং প্রোটোকলস: OSPF, BGP, RIP ও সুইচিং (VLAN, STP)", "sec_4_4"),
                ("৪.৫ এমপিএলএস (MPLS) ও নেটওয়ার্ক সিকিউরিটি ফান্ডামেন্টালস", "sec_4_5")
            ]),
            ("অধ্যায় ০৫: মাইক্রোওয়েভ, স্যাটেলাইট ও অ্যান্টেনা কমিউনিকেশন", "ch_05", [
                ("৫.১ মাইক্রোওয়েভ রেডিও প্রোপাগেশন ও ফ্রেনেল জোন (Fresnel Zone)", "sec_5_1"),
                ("৫.২ ফ্রি স্পেস পাথ লস (FSPL) ও লিঙ্ক বাজেট ক্যালকুলেশন", "sec_5_2"),
                ("৫.৩ অ্যান্টেনা থিওরি, গেইন, ডিরেক্টিভিটি ও VSWR", "sec_5_3"),
                ("৫.৪ স্যাটেলাইট কমিউনিকেশন: LEO, MEO, GEO ও ট্রান্সপন্ডার", "sec_5_4")
            ])
        ]),
        ("দ্বিতীয় খণ্ড: বাংলাদেশ টেলিকম ইকোসিস্টেম ও রেগুলেটরি ফ্রেমওয়ার্ক", None, [
            ("অধ্যায় ০৬: বাংলাদেশ টেলিকম নেটওয়ার্ক ও লাইসেন্সিং ফ্রেমওয়ার্ক", "ch_06", [
                ("৬.১ বাংলাদেশ টেলিকম নেটওয়ার্ক আর্কিটেকচার (BTRC Framework)", "sec_6_1"),
                ("৬.২ প্রধান লাইসেন্সিং লেয়ারসমূহ (IGW, ICX, IIG, NTTN, ANS, ITC, NIX)", "sec_6_2"),
                ("৬.৩ সাবমেরিন ক্যাবল নেটওয়ার্ক ও বঙ্গবন্ধু স্যাটেলাইট-১ (BS-1)", "sec_6_3"),
                ("৬.৪ বিটিআরসি স্পেকট্রাম বরাদ্দ, এমএনপি (MNP) ও QoS নীতিমালা", "sec_6_4")
            ])
        ]),
        ("তৃতীয় খণ্ড: বেসিক ইলেকট্রিক্যাল, ইলেকট্রনিক্স ও গণিত", None, [
            ("অধ্যায় ০৭: ইলেকট্রিক্যাল ও ইলেকট্রনিক্স ফান্ডামেন্টালস", "ch_07", [
                ("৭.১ সার্কিট অ্যানালাইসিস থিওরেমস (KVL, KCL, Thevenin, Norton)", "sec_7_1"),
                ("৭.২ এসি ফান্ডামেন্টালস, RLC সার্কিট ও পাওয়ার ফ্যাক্টর", "sec_7_2"),
                ("৭.৩ সেমিকন্ডাক্টর ও অপারেশনাল অ্যামপ্লিফায়ার (Op-Amp)", "sec_7_3"),
                ("৭.৪ টেলিকম -৪৮ ভোল্ট ডিসি (-48V DC) পাওয়ার সিস্টেম ও ব্যাটারি ব্যাকআপ", "sec_7_4")
            ]),
            ("অধ্যায় ০৮: টেলিকম গণিত ও গুরুত্বপূর্ণ ফর্মুলা শীট", "ch_08", [
                ("৮.১ এক নজরে সকল অত্যাবশ্যকীয় টেলিকম সূত্র ও একক (Formula Cheatsheet)", "sec_8_1"),
                ("৮.২ সিগন্যাল, স্যাম্পলিং ও চ্যানেল ক্যাপাসিটি সমাধান (Types 1-2)", "sec_8_2"),
                ("৮.৩ সেলুলার ট্রাফিক (Erlang), ক্লাস্টার ও ক্যাপাসিটি সমাধান (Types 3-4)", "sec_8_3"),
                ("৮.৪ অপটিক্যাল ফাইবার (NA) ও পাওয়ার বাজেট সমাধান (Types 5-6)", "sec_8_4"),
                ("৮.৫ মাইক্রোওয়েভ ও নেটওয়ার্ক ট্রান্সমিশন ডিলে সমাধান (Types 7-8)", "sec_8_5")
            ])
        ]),
        ("চতুর্থ খণ্ড: বিগত নিয়োগ পরীক্ষার প্রশ্ন ও সমাধান", None, [
            ("অধ্যায় ০৯: বিটিসিএল (BTCL) বিগত পরীক্ষার প্রশ্ন ও সমাধান", "ch_09", [
                ("৯.১ বিটিসিএল সহকারী ম্যানেজার (টেকনিক্যাল) লিখিত পরীক্ষা ২০১৫ সমাধান", "sec_9_1"),
                ("৯.২ বিটিসিএল সহকারী ম্যানেজার ও উপ-সহকারী প্রকৌশলী লিখিত পরীক্ষা ২০১৭", "sec_9_2")
            ]),
            ("অধ্যায় ১০: বিটিআরসি (BTRC) বিগত পরীক্ষার প্রশ্ন ও সমাধান", "ch_10", [
                ("১০.১ বিটিআরসি সহকারী পরিচালক (কারিগরি) নিয়োগ পরীক্ষা ২০২১ সমাধান", "sec_10_1"),
                ("১০.২ বিটিআরসি আইসিটি ও টেকনিক্যাল রচনা লেখার মডেল", "sec_10_2")
            ]),
            ("অধ্যায় ১১: টেলিটক, বিএসসিএল ও বুয়েট/এমআইএসটি স্পেশাল", "ch_11", [
                ("১১.১ টেলিটক বাংলাদেশ লিমিটেড সহকারী ব্যবস্থাপক প্রশ্ন সমাধান", "sec_11_1"),
                ("১১.২ বুয়েট ও এমআইএসটি পরিচালিত টেলিকম পরীক্ষার কমন প্রশ্নসমূহ", "sec_11_2")
            ])
        ]),
        ("পঞ্চম খণ্ড: ভাইভা স্পেশাল ও ইন্টারভিউ গাইডলাইন", None, [
            ("অধ্যায় ১২: টেলিকম সেরা ১০০ টেকনিক্যাল ভাইভা প্রশ্নোত্তর", "ch_12", [
                ("১২.১ কোর টেলিকম ও মোবাইল কমিউনিকেশন ভাইভা প্রশ্ন", "sec_12_1"),
                ("১২.২ ট্রান্সমিশন, অপটিক্যাল ও আইপি নেটওয়ার্কিং ভাইভা প্রশ্ন", "sec_12_2"),
                ("১২.৩ বাংলাদেশ টেলিকম ও বিটিআরসি রেগুলেশন ভাইভা প্রশ্ন", "sec_12_3")
            ]),
            ("অধ্যায় ১৩: বাস্তবভিত্তিক ফিল্ড সিনারিও ও ট্রাবলশুটিং", "ch_13", [
                ("১৩.১ সিনারিও ১: মোবাইল কল ড্রপ কেন হয় এবং কীভাবে ড্রপ রেট কমানো যায়?", "sec_13_1"),
                ("১৩.২ সিনারিও ২: অপটিক্যাল ফাইবারে কাট পড়লে OTDR দিয়ে কীভাবে ফল্ট শনাক্ত করবেন?", "sec_13_2"),
                ("১৩.৩ সিনারিও ৩: টেলিকম সাইটে পাওয়ার ফেইলিউর ও ব্যাটারি ব্যাকআপ ম্যানেজমেন্ট", "sec_13_3")
            ]),
            ("অধ্যায় ১৪: টেলিকম অ্যাক্রোনিম ও পরিভাষার বর্ণানুক্রমিক কোষ (Glossary A to Z)", "ch_14", [
                ("১৪.১ এ থেকে জেড (A to Z) পর্যন্ত বহুল ব্যবহৃত টেলিকম পরিভাষা", "sec_14_1")
            ])
        ])
    ]

    for part_title, part_bm, chapters in toc_structure:
        p_part = doc.add_paragraph()
        p_part.paragraph_format.space_before = Pt(8)
        p_part.paragraph_format.space_after = Pt(3)
        r_part = p_part.add_run(f"📁 {part_title}")
        set_font(r_part, size_pt=11.5, bold=True, color=COLOR_PRIMARY)

        for ch_data in chapters:
            if len(ch_data) == 3:
                ch_title, ch_bm, topics = ch_data
                p_ch = doc.add_paragraph()
                add_toc_link(p_ch, f"📖 {ch_title}", ch_bm, is_bold=True, level=1)

                for top_title, top_bm in topics:
                    p_top = doc.add_paragraph()
                    add_toc_link(p_top, top_title, top_bm, is_bold=False, level=2)
            else:
                ch_title, ch_bm = ch_data
                p_ch = doc.add_paragraph()
                add_toc_link(p_ch, f"📖 {ch_title}", ch_bm, is_bold=False, level=1)

    doc.add_page_break()

    # ==========================================
    # 5. CHAPTER 1: FUNDAMENTALS OF COMMUNICATION
    # ==========================================
    add_h1("অধ্যায় ০১: ফান্ডামেন্টালস অব কমিউনিকেশন সিস্টেম", "ch_01")
    add_body("টেলিকমিউনিকেশন ব্যবস্থার মূল কাজ হলো এক স্থান থেকে অন্য স্থানে যেকোনো তথ্য (ভয়েস, টেক্সট, অডিও, ভিডিও বা ডাটা) নির্ভরযোগ্যভাবে এবং অবিকৃত অবস্থায় স্থানান্তর করা। যেকোনো কমিউনিকেশন ইঞ্জিনিয়ারিং পরীক্ষার ভিত্তি রচিত হয় এই অধ্যায়ের মৌলিক নীতিগুলোর ওপর।")

    add_h2("১.১ কমিউনিকেশন সিস্টেমের বেসিক মডেল ও উপাদান", "sec_1_1")
    add_body("একটি স্ট্যান্ডার্ড ইলেকট্রনিক কমিউনিকেশন সিস্টেম প্রধানত ৫টি মৌলিক উপাদান নিয়ে গঠিত:")
    add_bullet("১. ইনফরমেশন সোর্স ও ইনপুট ট্রান্সডিউসার (Information Source & Transducer):", "যা ভৌত সংকেতকে (শব্দ, দৃশ্য) সমতুল্য বৈদ্যুতিক সংকেতে রূপান্তর করে (যেমন: মাইক্রোফোন)।")
    add_bullet("২. ট্রান্সমিটার (Transmitter):", "যা বেসব্যান্ড সিগন্যালকে দূরবর্তী স্থানে পাঠানোর উপযোগী করার জন্য ফিল্টারিং, অ্যামপ্লিফিকেশন এবং মডুলেশন (Modulation) সম্পন্ন করে।")
    add_bullet("৩. চ্যানেল / ট্রান্সমিশন মিডিয়া (Channel):", "যে মাধ্যমের মধ্য দিয়ে সংকেত প্রবাহিত হয়। এটি ওয়্যার্ড (অপটিক্যাল ফাইবার, টুইস্টেড পেয়ার, কোঅক্সিয়াল ক্যাবল) বা ওয়্যারলেস (ফ্রি স্পেস / বায়ুমণ্ডল) হতে পারে।")
    add_bullet("৪. রিসিভার (Receiver):", "যা রিসিভড দুর্বল ও নয়েজযুক্ত সংকেত থেকে ডিমডুলেশন ও ফিল্টারিং এর মাধ্যমে মূল তথ্য পুনরুদ্ধার করে।")
    add_bullet("৫. আউটপুট ট্রান্সডিউসার (Output Transducer):", "যা বৈদ্যুতিক সিগন্যালকে পুনরায় মানুষের বা গ্রাহক ডিভাইসের ব্যবহার উপযোগী রূপে রূপান্তর করে (যেমন: লাউডস্পিকার, ডিসপ্লে স্ক্রিন)।")

    add_img("images/comm_block_diagram.png", "কমিউনিকেশন সিস্টেমের মৌলিক ব্লক ডায়াগ্রাম ও চ্যানেল নয়েজ অনুপ্রবেশ")

    add_callout("নয়েজ (Noise) ও ডিসটরশন (Distortion) এর পার্থক্য", "ডিসটরশন হলো ট্রান্সমিশন সিস্টেমের অভ্যন্তরীণ অ-রৈখিকতা বা সীমাবদ্ধতার কারণে সিগন্যালের বিকৃতি (যা দূর করা সম্ভব)। অন্যদিকে নয়েজ হলো বাইরের অনাকাঙ্ক্ষিত র্যান্ডম সংকেত (যেমন: থার্মাল নয়েজ, কসমিক নয়েজ) যা চ্যানেলে যুক্ত হয় এবং সিগন্যালে স্থায়ী ত্রুটি তৈরি করতে পারে।")

    add_h2("১.২ অ্যানালগ মডুলেশন থিওরি (AM, FM, PM)", "sec_1_2")
    add_body("বেসব্যান্ড সংকেতের ফ্রিকোয়েন্সি অত্যন্ত কম হওয়ায় তা সরাসরি অ্যান্টেনা দিয়ে পাঠানো যায় না (কারণ অ্যান্টেনার উচ্চতা λ/4 ফ্রিকোয়েন্সির ব্যস্তানুপাতিক)। তাই উচ্চ ফ্রিকোয়েন্সির ক্যারিয়ার সিগন্যালে তথ্য সিগন্যাল সুপারইম্পোজ করাকে মডুলেশন বলে।")

    add_formula(
        "অ্যাম্প্লিটিউড মডুলেশন (AM) সমীকরণ ও পাওয়ার",
        "s(t) = A_c [1 + m_a * cos(2*pi*f_m*t)] * cos(2*pi*f_c*t)\n"
        "Total Power: P_t = P_c * (1 + m_a^2 / 2)",
        "• A_c = ক্যারিয়ার অ্যাম্প্লিটিউড, A_m = মেসেজ অ্যাম্প্লিটিউড\n"
        "• m_a = A_m / A_c হলো মডুলেশন ইনডেক্স (Modulation Index)\n"
        "• P_c = ক্যারিয়ার পাওয়ার, P_t = মোট ট্রান্সমিটেড পাওয়ার\n"
        "• ব্যান্ডউইথ: BW_AM = 2 * f_m",
        "১. ওভার-মডুলেশন (m_a > 1) হলে এনভেলপ ডিটেক্টরে সিগন্যাল মারাত্মক বিকৃত বা কেটে যায়। তাই আদর্শ ক্ষেত্রে m_a <= 1 হতে হয়।\n"
        "২. সমীকরণ থেকে স্পষ্ট যে ক্যারিয়ারের ভেতর কোনো তথ্য থাকে না, সকল তথ্য থাকে দুটি সাইডব্যান্ডে (USB ও LSB)। মোট পাওয়ারের মাত্র এক-তৃতীয়াংশ সাইডব্যান্ডে থাকে, বাকি দুই-তৃতীয়াংশ ক্যারিয়ারেই অপচয় হয়। এজন্য উচ্চ দক্ষতায় SSB-SC বা DSB-SC ব্যবহৃত হয়।"
    )

    add_formula(
        "ফ্রিকোয়েন্সি মডুলেশন (FM) ও কারসন'স রুল (Carson's Rule)",
        "BW_FM = 2 * (Delta_f + f_m) = 2 * f_m * (beta + 1)",
        "• Delta_f = ফ্রিকোয়েন্সি ডেভিয়েশন (Frequency Deviation: ক্যারিয়ারের সর্বোচ্চ স্থানান্তর)\n"
        "• f_m = মেসেজ সিগন্যালের সর্বোচ্চ ফ্রিকোয়েন্সি\n"
        "• beta = Delta_f / f_m হলো এফএম মডুলেশন ইনডেক্স",
        "এফএম সিগন্যালে তাত্ত্বিকভাবে অসীম সংখ্যক সাইডব্যান্ড উৎপন্ন হয় (বেসেল ফাংশন অনুযায়ী)। তবে জন কারসনের রুল অনুযায়ী ৯৮% সিগন্যাল পাওয়ার এই ব্যান্ডের ভেতরে সীমাবদ্ধ থাকে। এফএম অ্যাম্প্লিটিউডের নয়েজে শতভাগ মুক্ত থাকে কারণ রিসিভারে অ্যামপ্লিটিউড লিমিটার বসিয়ে সকল বাহ্যিক ভোল্টেজ স্পাইক ছেঁটে ফেলা যায়।"
    )

    add_h2("১.৩ স্যাম্পলিং থিওরেম ও শ্যানন ক্যাপাসিটি", "sec_1_3")
    add_body("অ্যানালগ সংকেতকে ডিজিটালে রূপান্তর করার প্রথম ধাপ হলো স্যাম্পলিং। নাইকুইস্ট স্যাম্পলিং থিওরেম অনুযায়ী, কোনো সিগন্যালকে নিখুঁতভাবে রিকনস্ট্রাক্ট করতে হলে স্যাম্পলিং ফ্রিকোয়েন্সি সিগন্যালের সর্বোচ্চ কম্পাঙ্কের দ্বিগুণের সমান বা বেশি হতে হবে।")

    add_formula(
        "নাইকুইস্ট স্যাম্পলিং শর্ত ও নাইকুইস্ট ইন্টারভাল",
        "f_s >= 2 * f_m\nT_s = 1 / f_s <= 1 / (2 * f_m)",
        "• f_m = সিগন্যালের সর্বোচ্চ ফ্রিকোয়েন্সি উপাদান (Hz)\n"
        "• f_s = স্যাম্পলিং রেট (Samples per second বা Hz)\n"
        "• T_s = দুটি স্যাম্পলের মধ্যবর্তী সময় ব্যবধান (Nyquist Interval)",
        "যদি স্যাম্পলিং ফ্রিকোয়েন্সি 2*f_m এর চেয়ে কম হয় (f_s < 2*f_m), তবে হাই ফ্রিকোয়েন্সি উপাদানগুলো লো ফ্রিকোয়েন্সির উপর ওভারল্যাপ হয়ে সিগন্যালে মারাত্মক বিকৃতি সৃষ্টি করে, যাকে 'অ্যালিয়াসিং' (Aliasing) বলে। এটি প্রতিরোধে স্যাম্পলিংয়ের ঠিক পূর্বে একটি অ্যান্টি-অ্যালিয়াসিং লো-পাস ফিল্টার (Anti-Aliasing Filter) ব্যবহার করে অতিরিক্ত ফ্রিকোয়েন্সি ছেঁটে ফেলা হয়।"
    )

    add_formula(
        "শ্যানন-হার্টলি চ্যানেল ক্যাপাসিটি উপপাদ্য (Shannon-Hartley Theorem)",
        "C = B * log2(1 + S/N)  [Bits per second - bps]",
        "• C = চ্যানেলের তাত্ত্বিক সর্বোচ্চ ডাটা ট্রান্সমিশন রেট (Channnel Capacity in bps)\n"
        "• B = চ্যানেলের অ্যানালগ ব্যান্ডউইথ (Hertz - Hz)\n"
        "• S/N = সিগন্যাল টু নয়েজ পাওয়ার অনুপাত (Linear Ratio, মনে রাখবেন এটি dB নয়!)",
        "শ্যাননের উপপাদ্যটি নির্দেশ করে যে ডাটা রেট কেবল ফিজিক্যাল ব্যান্ডউইথের ওপর নির্ভর করে না, সিগন্যাল পাওয়ার ও নয়েজের অনুপাতের ওপরও নির্ভরশীল। ব্যান্ডউইথ দ্বিগুণ করলে ডাটা রেট দ্বিগুণ বাড়ে (রৈখিক সম্পর্ক), কিন্তু সিগন্যাল পাওয়ার দ্বিগুণ করলে লগারিদমিক সম্পর্কের কারণে ডাটা রেট খুব সামান্যই বাড়ে। নয়েজপূর্ণ চ্যানেলে শতভাগ ত্রুটিমুক্ত ডাটা পাঠানোর সর্বোচ্চ তাত্ত্বিক সীমারেখা হলো এই শ্যানন ক্যাপাসিটি।"
    )

    add_formula(
        "নাইকুইস্ট নয়েজহীন চ্যানেল ডাটা রেট (Nyquist Bit Rate - Noiseless)",
        "C_noiseless = 2 * B * log2(M)  [Bits per second - bps]",
        "• B = চ্যানেলের ব্যান্ডউইথ (Hz)\n"
        "• M = সিগন্যালিং ভোল্টেজ লেভেলের সংখ্যা (যেমন: বাইনারির জন্য M=2, QPSK এর জন্য M=4, 16-QAM এর জন্য M=16)\n"
        "• log2(M) = প্রতি সিম্বলে বিটের সংখ্যা (Bits per Symbol)",
        "শ্যানন বনাম নাইকুইস্টের মূল পার্থক্য: নাইকুইস্টের সূত্রটি নয়েজহীন আদর্শ চ্যানেলে প্রযোজ্য যেখানে কেবলমাত্র লেভেল (M) বাড়িয়ে ডাটা রেট ইচ্ছামত বাড়ানো যায়। কিন্তু বাস্তব চ্যানেল নয়েজপূর্ণ, যেখানে লেভেল অতিরিক্ত বাড়ালে নয়েজের কারণে রিসিভার এক লেভেলকে অন্য লেভেলের সাথে গুলিয়ে ফেলে। শ্যাননের সূত্র বাস্তব নয়েজের উপস্থিতিতে সর্বোচ্চ কত ডাটা পাঠানো সম্ভব তার চরম তাত্ত্বিক সীমা নির্ধারণ করে।"
    )

    add_h2("১.৪ পালস কোড মডুলেশন (PCM) ও কোয়ান্টাইজেশন", "sec_1_4")
    add_body("পিসিএম (Pulse Code Modulation) হলো টেলিযোগাযোগ ব্যবস্থায় অ্যানালগ ভয়েসকে পিওর ডিজিটাল বিট স্ট্রিমে রূপান্তরের সবচেয়ে জনপ্রিয় পদ্ধতি। এটি ৩টি মূল ধাপে সম্পন্ন হয়:")
    add_bullet("১. স্যাম্পলিং (Sampling):", "অ্যানালগ সংকেত থেকে নির্দিষ্ট সময় পর পর স্যাম্পল মান সংগ্রহ করা (PAM সংকেত তৈরি)।")
    add_bullet("২. কোয়ান্টাইজেশন (Quantization):", "অসীম সংখ্যক স্যাম্পল অ্যাম্প্লিটিউডকে সসীম সংখ্যক ডিসক্রিট লেভেলে (L = 2^n) সীমাবদ্ধ করা। এর ফলে যে সূক্ষ্ম অমিল তৈরি হয় তাকে কোয়ান্টাইজেশন নয়েজ (Qe) বলে।")
    add_bullet("৩. এনকোডিং (Encoding):", "প্রতিটি কোয়ান্টাইজড লেভেলকে n-বিট বাইনারি শব্দে রূপান্তর করা।")

    add_formula(
        "পিসিএম বিট রেট ও সিগন্যাল-টু-কোয়ান্টাইজেশন নয়েজ রেশিও (SQNR)",
        "R_b = n * f_s = n * 2 * f_m  [bps]\n"
        "SQNR (dB) = 6.02 * n + 1.76 dB",
        "• n = প্রতিটি স্যাম্পল প্রকাশ করতে ব্যবহৃত বিট সংখ্যা\n"
        "• L = 2^n হলো কোয়ান্টাইজেশন লেভেল সংখ্যা\n"
        "• f_s = স্যাম্পলিং ফ্রিকোয়েন্সি (Hz)\n"
        "• SQNR = কোয়ান্টাইজেশন সিগন্যাল-টু-নয়েজ অনুপাত",
        "সূত্রে 6.02 সংখ্যাটি মূলত 20*log10(2) থেকে এসেছে! এর বাস্তব অর্থ হলো: কোয়ান্টাইজারে প্রতিটি অতিরিক্ত ১-বিট (n) বৃদ্ধি করলে সিগন্যালের কোয়ালিটি বা SQNR প্রায় 6 dB বৃদ্ধি পায়। যেমন: ৭-বিটের চেয়ে ৮-বিট পিসিএমে নয়েজ ৪ গুণ কমে যায়।"
    )

    add_callout("কমপ্যান্ডিং (Companding: A-law বনাম mu-law)", "ক্ষুদ্র শব্দ বা দুর্বল সিগন্যালে কোয়ান্টাইজেশন নয়েজের প্রভাব কমাতে নন-ইউনিফর্ম কোয়ান্টাইজেশন বা কমপ্যান্ডিং ব্যবহৃত হয়। উত্তর আমেরিকা ও জাপানে mu-law ব্যবহৃত হয়, অন্যদিকে বাংলাদেশ, ভারত ও ইউরোপের পিএসটিএন/টেলিকম সিস্টেমে ITU-T প্রমিত A-law (A = 87.6) কমপ্যান্ডিং ব্যবহৃত হয়।")

    add_h2("১.৫ ডিজিটাল মডুলেশন টেকনিকস ও কনস্টেলেশন ডায়াগ্রাম", "sec_1_5")
    add_body("ডিজিটাল ডাটা (0 ও 1) কে এনালগ ক্যারিয়ারে মডুলেট করার প্রধান স্কিমসমূহ হলো:")
    add_bullet("ASK (Amplitude Shift Keying):", "বিট অনুযায়ী ক্যারিয়ারের অ্যাম্প্লিটিউড পরিবর্তিত হয়। এটি নয়েজ সংবেদনশীল।")
    add_bullet("FSK (Frequency Shift Keying):", "বিট অনুযায়ী ক্যারিয়ারের ফ্রিকোয়েন্সি পরিবর্তিত হয়।")
    add_bullet("PSK (Phase Shift Keying):", "বিট অনুযায়ী ক্যারিয়ারের ফেজ পরিবর্তিত হয় (BPSK, QPSK)। এটি নয়েজে অধিকতর সুরক্ষিত।")
    add_bullet("QAM (Quadrature Amplitude Modulation):", "অ্যাম্প্লিটিউড ও ফেজ উভয়ের যুগপৎ পরিবর্তনের মাধ্যমে অত্যন্ত উচ্চ ব্যান্ডউইথ এফিশিয়েন্সি অর্জন করা হয় (যেমন 16-QAM, 64-QAM, 256-QAM)।")

    add_img("images/constellation_diagrams.png", "ডিজিটাল মডুলেশন কনস্টেলেশন ডায়াগ্রাম (BPSK, QPSK ও 16-QAM)")

    add_h2("১.৬ অধ্যায় ০১ এর বিগত পরীক্ষার প্রশ্ন ও গাণিতিক সমাধান", "sec_1_6")

    add_h3("গাণিতিক সমস্যা ০১ (BTCL ও BUET নিয়োগ পরীক্ষা):")
    add_body("একটি ভয়েস সিগন্যালের ফ্রিকোয়েন্সি রেঞ্জ 300 Hz থেকে 3400 Hz। একে 8-বিট PCM সিস্টেমে ডিজিটালাইজ করা হলো।\n(ক) নাইকুইস্ট স্যাম্পলিং রেট কত?\n(খ) পিসিএম বিট রেট কত?\n(গ) ন্যূনতম ট্রান্সমিশন ব্যান্ডউইথ কত?\n(ঘ) SQNR কত?")

    add_math_solution(
        "টেলিফোন ভয়েস সিগন্যালের ৮-বিট পিসিএম প্যারামিটার নির্ণয়",
        "• সর্বোচ্চ মেসেজ ফ্রিকোয়েন্সি f_m = 3400 Hz = 3.4 kHz\n• বিট সংখ্যা n = 8 bits/sample",
        "১. f_s (Nyquist) = 2 * f_m\n২. R_b = n * f_s\n৩. BW_min = R_b / 2\n৪. SQNR = 6.02 * n + 1.76 dB",
        "ধাপ ১: নাইকুইস্ট স্যাম্পলিং রেট:\n"
        "   f_s = 2 * 3400 Hz = 6800 Hz (বা 6.8 kHz)।\n"
        "   [বাস্তব টেলিফোনি স্ট্যান্ডার্ডে গার্ড ব্যান্ডের জন্য f_s = 8000 Hz বা 8 kHz ধরা হয়]\n\n"
        "ধাপ ২: পিসিএম বিট রেট (Standard Telephony Rate):\n"
        "   R_b = n * f_s = 8 bits * 8000 samples/sec = 64,000 bps = 64 kbps।\n"
        "   (এই 64 kbps হলো আন্তর্জাতিক ডিজিটাল টেলিফোনির মৌলিক একক, যাকে DS0 চ্যানেল বলা হয়)\n\n"
        "ধাপ ৩: ন্যূনতম ট্রান্সমিশন ব্যান্ডউইথ:\n"
        "   BW_min = R_b / 2 = 64 kbps / 2 = 32 kHz।\n\n"
        "ধাপ ৪: সিগন্যাল-টু-কোয়ান্টাইজেশন নয়েজ রেশিও (SQNR):\n"
        "   SQNR = 6.02 * 8 + 1.76 = 48.16 + 1.76 = 49.92 dB。",
        "• পরীক্ষার্থীরা প্রায়ই f_s এর মান 3400 Hz এর দ্বিগুণ 6800 Hz ব্যবহার করেন, যা তাত্ত্বিকভাবে সঠিক হলেও বিটিসিএল বা টেলিকম পরীক্ষায় স্ট্যান্ডার্ড টেলিফোনি রেট 8 kHz উল্লেখ করে 64 kbps বের করা সবচেয়ে নিরাপদ ও প্রশংসনীয়।"
    )

    add_h3("গাণিতিক সমস্যা ০২ (BTRC ও বুয়েট নিয়োগ পরীক্ষা):")
    add_body("একটি টেলিফোন চ্যানেলের ব্যান্ডউইথ 4 kHz এবং সিগন্যাল-টু-নয়েজ রেশিও (SNR) 30 dB হলে শ্যানন চ্যানেল ক্যাপাসিটি নির্ণয় করুন।")

    add_math_solution(
        "শ্যানন-হার্টলি উপপাদ্য প্রয়োগে চ্যানেল ক্যাপাসিটি নির্ণয়",
        "• ব্যান্ডউইথ B = 4 kHz = 4000 Hz\n• SNR = 30 dB",
        "১. S/N = 10^(SNR_dB / 10)\n২. C = B * log2(1 + S/N)",
        "ধাপ ১: SNR কে ডেসিবেল (dB) থেকে লিনিয়ার রেশিওতে রূপান্তর:\n"
        "   SNR_dB = 10 * log10(S/N)\n"
        "   => 30 = 10 * log10(S/N)\n"
        "   => log10(S/N) = 3\n"
        "   => S/N = 10^3 = 1000\n\n"
        "ধাপ ২: শ্যানন ক্যাপাসিটি সূত্রে মান বসানো:\n"
        "   C = 4000 * log2(1 + 1000) = 4000 * log2(1001)\n\n"
        "ধাপ ৩: বেস-২ লগারিদম ভাঙানো (log2(x) = ln(x) / ln(2)):\n"
        "   log2(1001) = ln(1001) / ln(2) = 6.90875 / 0.69315 = 9.9672\n\n"
        "ধাপ ৪: চূড়ান্ত ক্যাপাসিটি হিসাব:\n"
        "   C = 4000 * 9.9672 = 39,869 bps ≈ 39.87 kbps।",
        "• সবচেয়ে মারাত্মক ভুল: অনেকেই SNR_dB এর মান '30' সরাসরি শ্যানন সূত্রে বসিয়ে দেন [C = B * log2(1+30)], যা সম্পূর্ণ ভুল! সূত্রে অবশ্যই লিনিয়ার রেশিও (S/N = 1000) বসাতে হবে।"
    )

    add_h3("গাণিতিক সমস্যা ০৩ (বিসিএস ও বুয়েট নিয়োগ পরীক্ষা): শ্যানন বনাম নাইকুইস্ট ডাটা রেট")
    add_body("একটি টেলিফোন ট্রান্সমিশন চ্যানেলের ব্যান্ডউইথ 3 kHz এবং সিগন্যাল-টু-নয়েজ রেশিও (SNR) 20 dB।\n(ক) চ্যানেলটির সর্বোচ্চ শ্যানন ক্যাপাসিটি কত?\n(খ) যদি চ্যানেলটিকে নয়েজহীন বিবেচনা করে 16-লেভেল ডিজিটাল সংকেত পাঠানো হয়, তবে নাইকুইস্ট সর্বোচ্চ বিট রেট কত?\n(গ) শ্যানন ক্যাপাসিটি অর্জনের জন্য নয়েজহীন চ্যানেলে ন্যূনতম কতটি লেভেল (M) প্রয়োজন?")

    add_math_solution(
        "শ্যানন বনাম নাইকুইস্ট ফর্মুলা প্রয়োগ ও সিগন্যালিং লেভেল নির্ণয়",
        "• ব্যান্ডউইথ B = 3 kHz = 3000 Hz\n• SNR = 20 dB\n• নাইকুইস্ট লেভেল M = 16",
        "১. S/N = 10^(20/10) = 100\n২. C_shannon = B * log2(1 + S/N)\n৩. C_nyquist = 2 * B * log2(M)\n৪. M_required = 2^(C_shannon / (2*B))",
        "ধাপ ১: শ্যানন চ্যানেল ক্যাপাসিটি গণনা:\n"
        "   S/N = 10^(20/10) = 10^2 = 100\n"
        "   C_shannon = 3000 * log2(1 + 100) = 3000 * log2(101)\n"
        "   log2(101) = ln(101) / ln(2) = 4.6151 / 0.69315 = 6.658\n"
        "   C_shannon = 3000 * 6.658 = 19,974 bps ≈ 19.97 kbps।\n\n"
        "ধাপ ২: ১৬-লেভেল সংকেতের জন্য নাইকুইস্ট ডাটা রেট:\n"
        "   C_nyquist = 2 * B * log2(16) = 2 * 3000 * 4 = 24,000 bps = 24 kbps।\n\n"
        "ধাপ ৩: শ্যানন ক্যাপাসিটি অর্জনে প্রয়োজনীয় লেভেল সংখ্যা (M):\n"
        "   2 * B * log2(M) = C_shannon\n"
        "   => 2 * 3000 * log2(M) = 19,974\n"
        "   => log2(M) = 19,974 / 6000 = 3.329\n"
        "   => M = 2^(3.329) ≈ 10.05 ≈ 11 টি ডিসক্রিট লেভেল।",
        "• পরীক্ষায় নাইকুইস্ট ও শ্যানন গুলিয়ে ফেলা যাবে না: নয়েজ দেওয়া থাকলে শ্যানন আর লেভেল (M) দেওয়া থাকলে নাইকুইস্ট ব্যবহার করতে হবে। বাস্তব সিস্টেম কখনো শ্যানন ক্যাপাসিটি অতিক্রম করতে পারে না।"
    )

    add_back_to_toc()
    doc.add_page_break()

    # ==========================================
    # 6. CHAPTER 2: CELLULAR & MOBILE COMMUNICATIONS
    # ==========================================
    add_h1("অধ্যায় ০২: সেলুলার ও মোবাইল ওয়্যারলেস কমিউনিকেশন", "ch_02")
    add_body("মোবাইল নেটওয়ার্কের প্রাথমিক যুগে একটি সম্পূর্ণ শহরের জন্য মাত্র একটি উচ্চ ক্ষমতাসম্পন্ন ট্রান্সমিটার টাওয়ার ব্যবহৃত হতো, যার ফলে সীমিত স্পেকট্রামে অল্প সংখ্যক ব্যবহারকারী মাত্র একযোগে কথা বলতে পারতেন। সেলুলার ধারণার উদ্ভাবন একই ফ্রিকোয়েন্সি ভৌগোলিক দূরত্বের ব্যবধানে বারবার ব্যবহার (Frequency Reuse) করে সীমিত স্পেকট্রামে অসীম পরিমাণ ট্রাফিক ধারণক্ষমতা সৃষ্টি করেছে।")

    add_h2("২.১ সেলুলার কনসেপ্ট ও ফ্রিকোয়েন্সি রিইউজ", "sec_2_1")
    add_body("সেলুলার নেটওয়ার্কে সমগ্র ভৌগোলিক এলাকাকে ছোট ছোট বহুভুজে (সাধারণত জ্যামিতিক ষড়ভুজ বা Hexagon) ভাগ করা হয়। প্রতিটি ছোট এলাকাকে একটি 'সেল' (Cell) বলা হয়।")

    add_formula(
        "ক্লাস্টার সাইজ (Cluster Size - N) সমীকরণ ও হাঁটার নিয়ম (Walk Algorithm)",
        "N = i^2 + i*j + j^2",
        "• i, j = অঋণাত্মক পূর্ণসংখ্যা (Non-negative Integers: 0, 1, 2, ...)\n"
        "• N = ক্লাস্টারের মোট সেলের সংখ্যা। বৈধ মান: 1, 3, 4, 7, 9, 12, 19, ...",
        "ষড়ভুজ গ্রিডে কোনো সেলের একই ফ্রিকোয়েন্সি ব্যবহারকারী নিকটতম কো-চ্যানেল সেলটি খুঁজে বের করার জ্যামিতিক নিয়ম হলো: বর্তমান সেলের কেন্দ্র থেকে যেকোনো দিকে সোজা i সংখ্যক সেল অতিক্রম করুন, এরপর ৬০ ডিগ্রি ঘড়ির কাঁটার বিপরীত দিকে ঘুরে সোজা j সংখ্যক সেল গেলেই পরবর্তী কো-চ্যানেল সেল পাওয়া যাবে!"
    )

    add_img("images/cellular_cluster_n7.png", "সেলুলার ফ্রিকোয়েন্সি রিইউজ কনসেপ্ট: ক্লাস্টার সাইজ N = 7")

    add_formula(
        "কো-চ্যানেল রিইউজ রেশিও (Co-channel Reuse Ratio - Q)",
        "Q = D / R = sqrt(3 * N)",
        "• D = নিকটতম দুটি কো-চ্যানেল সেলের দূরবর্তী কেন্দ্রের দূরত্ব\n"
        "• R = সেলের ব্যাসার্ধ (Cell Radius)\n"
        "• N = ক্লাস্টার সাইজ",
        "Q যত বড় হবে, একই ফ্রিকোয়েন্সি ব্যবহারকারী সেলগুলো পরস্পরের থেকে তত দূরে অবস্থিত হবে, ফলে কো-চ্যানেল ইন্টারফারেন্স (CCI) তত কমবে এবং সিগন্যাল-টু-ইন্টারফারেন্স রেশিও (SIR) বৃদ্ধি পাবে। তবে Q বড় হলে প্রতি সেলে ফ্রিকোয়েন্সির সংখ্যা কমে যাওয়ায় মোট ক্যাপাসিটি হ্রাস পায়।"
    )

    add_formula(
        "কো-চ্যানেল সিগন্যাল-টু-ইন্টারফারেন্স অনুপাত (S/I বা SIR)",
        "S / I = (D / R)^n / i_0 = (sqrt(3 * N))^n / 6",
        "• S = কাঙ্ক্ষিত বেস স্টেশনের সিগন্যাল পাওয়ার\n"
        "• I = অন্যান্য একই ফ্রিকোয়েন্সি ব্যবহারকারী সেলগুলোর মোট ইন্টারফারেন্স পাওয়ার\n"
        "• i_0 = ১ম টায়ারের (First-Tier) কো-চ্যানেল ইন্টারফেয়ারিং সেলের সংখ্যা (ষড়ভুজ গ্রিডে সর্বদা i_0 = 6)\n"
        "• n = পাথ লস এক্সপোনেন্ট (সাধারণত আরবান এলাকায় n = 4)\n"
        "• N = ক্লাস্টার সাইজ",
        "জিএসএম-এ কেন N=7 ব্যবহৃত হয়? জিএসএম স্ট্যান্ডার্ডে গ্রাহকের স্পষ্ট ভয়েস কলের জন্য ন্যূনতম S/I হতে হয় 18 dB (লিনিয়ার অনুপাত ≈ 63.1)। যদি সূত্রে N = 7 এবং n = 4 বসানো হয়, তবে: S/I = (sqrt(21))^4 / 6 = (21^2) / 6 = 441 / 6 = 73.5 = 18.66 dB! এই চমৎকার গাণিতিক কারণেই বিশ্বজুড়ে জিএসএম নেটওয়ার্কে ডিফল্ট ক্লাস্টার সাইজ হিসেবে N = 7 বেছে নেওয়া হয়েছিল!"
    )

    add_h3("গাণিতিক সমস্যা (বিটিসিএল ও বুয়েট পরীক্ষা): সেলুলার স্পেকট্রাম বণ্টন ও মোট ক্যাপাসিটি")
    add_body("একটি সেলুলার সিস্টেমে মোট 33 MHz স্পেকট্রাম বরাদ্দ আছে। প্রতিটি সিমপ্লেক্স চ্যানেলের ব্যান্ডউইথ 25 kHz এবং দুটি সিমপ্লেক্স চ্যানেল মিলে একটি ফুল-ডুপ্লেক্স ভয়েস চ্যানেল গঠিত হয়। 1 MHz স্পেকট্রাম কন্ট্রোল চ্যানেলের জন্য সংরক্ষিত রাখা হয়েছে। ক্লাস্টার সাইজ N = 7 হলে:\n(ক) মোট কয়টি ডুপ্লেক্স ভয়েস চ্যানেল পাওয়া যাবে?\n(খ) প্রতিটি সেলে কতটি ভয়েস চ্যানেল বরাদ্দ থাকবে?\n(গ) যদি সমগ্র নেটওয়ার্কে মোট 28 টি সেল থাকে, তবে সিস্টেমটির মোট এককালীন গ্রাহক ধারণক্ষমতা (Total Capacity) কত?")

    add_math_solution(
        "৩৩ মেগাহার্টজ স্পেকট্রাম ও N=7 ক্লাস্টারে চ্যানেল সংখ্যা ও মোট ধারণক্ষমতা নির্ণয়",
        "• মোট বরাদ্দকৃত স্পেকট্রাম = 33 MHz = 33,000 kHz\n• কন্ট্রোল চ্যানেলের স্পেকট্রাম = 1 MHz = 1,000 kHz\n• ডুপ্লেক্স চ্যানেলের ব্যান্ডউইথ = 2 * 25 kHz = 50 kHz\n• ক্লাস্টার সাইজ N = 7\n• মোট সেলের সংখ্যা = 28 টি",
        "১. ভয়েসের জন্য উপলব্ধ স্পেকট্রাম = মোট স্পেকট্রাম - কন্ট্রোল স্পেকট্রাম\n২. মোট ডুপ্লেক্স চ্যানেল C_total = উপলব্ধ স্পেকট্রাম / চ্যানেল ব্যান্ডউইথ\n৩. প্রতি সেলে চ্যানেল K = C_total / N\n৪. মোট সিস্টেম ক্যাপাসিটি = K * মোট সেলের সংখ্যা",
        "ধাপ ১: ভয়েস স্পেকট্রাম নির্ধারণ:\n"
        "   Available Voice Spectrum = 33,000 kHz - 1,000 kHz = 32,000 kHz\n\n"
        "ধাপ ২: মোট ডুপ্লেক্স ভয়েস চ্যানেল সংখ্যা:\n"
        "   C_total = 32,000 kHz / 50 kHz = 640 টি ডুপ্লেক্স চ্যানেল।\n\n"
        "ধাপ ৩: প্রতি সেলে বরাদ্দকৃত চ্যানেল সংখ্যা (N = 7):\n"
        "   K = 640 / 7 = 91.42 ≈ 91 টি চ্যানেল প্রতি সেলে।\n\n"
        "ধাপ ৪: সমগ্র সিস্টেমের মোট ধারণক্ষমতা (28 টি সেল):\n"
        "   ক্লাস্টার সংখ্যা = 28 / 7 = 4 টি ক্লাস্টার।\n"
        "   মোট সিস্টেম ক্যাপাসিটি = 4 * 640 = 2,560 টি একযোগে সচল কল (Simultaneous Calls)।",
        "• সতর্কবার্তা: সিমপ্লেক্স চ্যানেল ব্যান্ডউইথ (25 kHz) কে ২ দিয়ে গুণ করে ডুপ্লেক্স চ্যানেলের ব্যান্ডউইথ (50 kHz) বের করতে হবে। সরাসরি ২৫ দিয়ে ভাগ করলে চ্যানেল সংখ্যা দ্বিগুণ হয়ে ভুল উত্তর আসবে।"
    )

    add_h2("২.২ নেটওয়ার্ক ক্যাপাসিটি বৃদ্ধির কৌশল", "sec_2_2")
    add_bullet("১. সেল স্প্লিটিং (Cell Splitting):", "উচ্চ ঘনত্বপূর্ণ এলাকায় একটি বড় সেলকে ভেঙ্গে একাধিক ক্ষুদ্রতর মাইক্রোসেলে পরিণত করা। এর ফলে সেলের ব্যাসার্ধ কমে এবং ফ্রিকোয়েন্সি রিইউজ বহুগুণ বেড়ে নেটওয়ার্ক ক্যাপাসিটি বৃদ্ধি পায়।")
    add_bullet("২. সেল সেক্টরিং (Cell Sectoring):", "ওমনি-ডিরেকশনাল অ্যান্টেনার পরিবর্তে ডিরেকশনাল অ্যান্টেনা ব্যবহার করে একটি সেলকে সাধারণত ৩টি (১২০° সেক্টর) বা ৬টি (৬০° সেক্টর) সেক্টরে বিভক্ত করা। এর ফলে কো-চ্যানেল ইন্টারফারেন্স কমে।")
    add_bullet("৩. মাইক্রোসেল জোন কনসেপ্ট (Microcell Zone):", "বেস স্টেশনের ট্রাফিকের চাপ বজায় রেখে অপটিক্যাল ফাইবার ফিডারের মাধ্যমে জোন সাইটগুলোতে সিগন্যাল বণ্টন।")

    add_h2("২.৩ ট্রাফিক ইঞ্জিনিয়ারিং ও আরল্যাং", "sec_2_3")
    add_body("টেলিকমে চ্যানেলের ধারণক্ষমতা ও ব্যস্ত সময়ে গ্রাহকের কল ট্রাফিক পরিমাপের একক হলো 'আরল্যাং' (Erlang)। ১ আরল্যাং মানে একটি চ্যানেল ১ ঘণ্টার মধ্যে নিরবচ্ছিন্নভাবে ৬০ মিনিটই ব্যস্ত ছিল।")

    add_formula(
        "ট্রাফিক লোড সমীকরণ (Traffic Intensity - Erlang)",
        "A = lambda * H  [Erlang]",
        "• lambda = প্রতি ঘণ্টায় আগত মোট কলের গড় সংখ্যা (Call Arrival Rate in calls/hour)\n"
        "• H = প্রতিটি কলের গড় স্থায়িত্বকাল (Average Call Holding Time in hours)",
        "১টি ভয়েস চ্যানেলে সর্বোচ্চ ট্রাফিক ধারণক্ষমতা ১ আরল্যাং। যেমন: ৩০টি ভয়েস চ্যানেলের একটি সিস্টেমে তাত্ত্বিক সর্বোচ্চ ট্রাফিক হতে পারে ৩০ আরল্যাং।"
    )

    add_bullet("Erlang-B মডেল:", "Blocked Calls Cleared (BCC) ব্যবস্থা। অর্থাৎ কোনো কল এলে যদি চ্যানেল ফাঁকা না থাকে, তবে তা তাৎক্ষণিক ড্রপ বা ব্লক হয়ে যায় (কোনো কিউ বা অপেক্ষার ব্যবস্থা থাকে না)।")
    add_bullet("Erlang-C মডেল:", "Blocked Calls Delayed (BCD) ব্যবস্থা। অর্থাৎ চ্যানেল খালি না থাকলে কলটিকে ওয়েটিং কিউ বা লাইনে রাখা হয় যতক্ষণ না চ্যানেল উন্মুক্ত হয় (যেমন: কাস্টমার কেয়ার কল সেন্টার)।")

    add_h2("২.৪ মোবাইল প্রজন্মের রূপান্তর (2G থেকে 5G NR)", "sec_2_4")

    tbl_gen = doc.add_table(rows=1, cols=4)
    tbl_gen.alignment = WD_TABLE_ALIGNMENT.CENTER
    g_hdrs = ["প্রজন্ম (Gen)", "কোর প্রযুক্তি / স্ট্যান্ডার্ড", "গতি / স্পেকট্রাম ব্যান্ডউইথ", "মূল বৈশিষ্ট্যসমূহ"]
    g_widths = [Inches(1.0), Inches(1.8), Inches(1.3), Inches(2.4)]
    for i, h in enumerate(g_hdrs):
        tbl_gen.rows[0].cells[i].text = h
        tbl_gen.rows[0].cells[i].width = g_widths[i]
        tcPr = tbl_gen.rows[0].cells[i]._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="0A2540"/>'))
        p = tbl_gen.rows[0].cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            set_font(r, size_pt=10, bold=True, color=RGBColor(255, 255, 255))

    gen_data = [
        ("2G (GSM)", "TDMA / FDM (900/1800 MHz)", "9.6 - 14.4 kbps (200 kHz)", "ডিজিটাল ভয়েস, এসএমএস, সিম কার্ড, সার্কিট সুইচড কোর"),
        ("2.5G / 2.75G", "GPRS / EDGE (8-PSK)", "171 kbps (GPRS), 384 kbps (EDGE)", "প্যাকেট সুইচড কোর (SGSN, GGSN), মোবাইল ইন্টারনেট সূচনা"),
        ("3G (UMTS)", "WCDMA / HSPA+ (2100 MHz)", "2 Mbps - 42 Mbps (5 MHz)", "ভিডিও কল, মোবাইল ব্রডব্যান্ড, স্মার্টফোন বিপ্লব"),
        ("4G (LTE)", "OFDMA / MIMO (All-IP)", "100 Mbps - 1 Gbps (Up to 20 MHz)", "ভয়েস ওভার এলটিই (VoLTE), ইপিসি কোর (MME, SGW, PGW)"),
        ("5G (NR)", "OFDM, mmWave, Beamforming", "1 Gbps - 20 Gbps (Up to 400 MHz)", "eMBB, URLLC (1ms latency), mMTC, নেটওয়ার্ক স্লাইসিং"),
    ]

    for row_data in gen_data:
        row_cells = tbl_gen.add_row().cells
        for i, val in enumerate(row_data):
            row_cells[i].text = val
            row_cells[i].width = g_widths[i]
            p = row_cells[i].paragraphs[0]
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.15
            for r in p.runs:
                set_font(r, size_pt=9.5, color=COLOR_TEXT)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    add_h2("২.৫ ৪জি এলটিই (4G LTE) আর্কিটেকচার ও কোর নেটওয়ার্ক (EPC)", "sec_2_5")
    add_body("LTE নেটওয়ার্ক সম্পূর্ণ অল-আইপি (All-IP) ভিত্তিক। এটি দুটি প্রধান অংশে বিভক্ত:")
    add_bullet("১. E-UTRAN (Evolved UMTS Terrestrial Radio Access Network):", "যেখানে একক উপাদান হিসেবে eNodeB (Evolved NodeB) সরাসরি ইউজার ইকুইপমেন্টের (UE) সাথে রেডিও সংযোগ স্থাপন করে।")
    add_bullet("২. EPC (Evolved Packet Core):", "যার মূল উপাদানসমূহ হলো:\n• MME (Mobility Management Entity): ইউজার প্রমাণীকরণ, সিকিউরিটি এবং পেজিং কন্ট্রোল করে।\n• SGW (Serving Gateway): ডাটা প্যাকেট রাউট ও ফরোয়ার্ড করে।\n• PGW (PDN Gateway): এক্সটার্নাল প্যাকেট ডাটা নেটওয়ার্ক (ইন্টারনেট) এর সাথে গেটওয়ে সংযোগ দেয় এবং আইপি বরাদ্দ করে।\n• HSS (Home Subscriber Server): ব্যবহারকারীর কেন্দ্রীয় সাবস্ক্রিপশন ডাটাবেজ।")

    add_back_to_toc()
    doc.add_page_break()

    # ==========================================
    # 7. CHAPTER 3: OPTICAL FIBER COMMUNICATIONS
    # ==========================================
    add_h1("অধ্যায় ০৩: অপটিক্যাল ফাইবার কমিউনিকেশন ও ট্রান্সমিশন নেটওয়ার্ক", "ch_03")
    add_body("টেলিযোগাযোগের গ্লোবাল ব্যাকবোন সম্পূর্ণভাবে অপটিক্যাল ফাইবার ট্রান্সমিশনের ওপর নির্ভরশীল। বিশাল ব্যান্ডউইথ, অত্যন্ত কম সংকেত লস (Attenuation), তড়িৎ-চৌম্বকীয় ইন্টারফারেন্স (EMI) থেকে শতভাগ মুক্তি এবং উচ্চ নিরাপত্তার কারণে লং-ডিস্ট্যান্স ব্যাকহল ও সাবমেরিন যোগাযোগে ফাইবার অপটিক্সের কোনো বিকল্প নেই।")

    add_h2("৩.১ অপটিক্যাল ফাইবারের মূলনীতি ও আলোক বিজ্ঞান", "sec_3_1")
    add_body("অপটিক্যাল ফাইবারের মধ্য দিয়ে আলোর সংকেত 'পূর্ণ অভ্যন্তরীণ প্রতিফলন' (Total Internal Reflection - TIR) প্রক্রিয়ায় সঞ্চালিত হয়। এর জন্য দুটি প্রধান শর্ত আবশ্যক:\n১. আলোকে অবশ্যই ঘনতর মাধ্যম (Core: n1) থেকে লঘুতর মাধ্যমে (Cladding: n2) প্রবেশ করতে হবে, অর্থাৎ n1 > n2।\n২. আপাতন কোণ (Angle of incidence) অবশ্যই মাধ্যমদ্বয়ের সংকট কোণের (Critical Angle: θc) চেয়ে বড় হতে হবে।")

    add_formula(
        "সংকট কোণ ও নিউমেরিক্যাল অ্যাপারচার (NA)",
        "sin(theta_c) = n_2 / n_1\n"
        "NA = sqrt(n_1^2 - n_2^2) = sin(theta_a) = n_1 * sqrt(2 * Delta)",
        "• n_1 = কোরের প্রতিসরাঙ্ক (Core Refractive Index: সাধারণত 1.48)\n"
        "• n_2 = ক্ল্যাডিংয়ের প্রতিসরাঙ্ক (Cladding Refractive Index: সাধারণত 1.46)\n"
        "• Delta = (n_1 - n_2) / n_1 হলো আপেক্ষিক প্রতিসরাঙ্ক পার্থক্য (Fractional Index Difference)\n"
        "• theta_a = সর্বোচ্চ এক্সেপটেন্স কোণ (Acceptance Angle)",
        "নিউমেরিক্যাল অ্যাপারচার (NA) ফাইবারের আলোক সংগ্রহ ক্ষমতার পরিমাপক। ফাইবার কোরে আলো কেবল তখনই পূর্ণ অভ্যন্তরীণ প্রতিফলনের মাধ্যমে অগ্রসর হবে যখন আলো এক্সেপটেন্স কোণের (theta_a) ভেতরের কোনো কোণ দিয়ে প্রবেশ করবে।"
    )

    add_img("images/optical_fiber_tir.png", "অপটিক্যাল ফাইবার কোরে পূর্ণ অভ্যন্তরীণ প্রতিফলন ও নিউমেরিক্যাল অ্যাপারচার")

    add_h3("গাণিতিক সমস্যা (বুয়েট ও টেলিটক পরীক্ষা): নিউমেরিক্যাল অ্যাপারচার ও এক্সেপটেন্স কোণ")
    add_body("একটি স্টেপ-ইনডেক্স অপটিক্যাল ফাইবারের কোরের প্রতিসরাঙ্ক n1 = 1.48 এবং ক্ল্যাডিংয়ের প্রতিসরাঙ্ক n2 = 1.46।\n(ক) কোর ও ক্ল্যাডিংয়ের সংকট কোণ (Critical Angle) কত?\n(খ) ফাইবারের নিউমেরিক্যাল অ্যাপারচার (NA) কত?\n(গ) বাতাসে (n0 = 1.0) সর্বোচ্চ এক্সেপটেন্স কোণ (Acceptance Angle) কত?\n(ঘ) আপেক্ষিক প্রতিসরাঙ্ক পার্থক্য (Fractional Index Difference) Delta নির্ণয় করুন।")

    add_math_solution(
        "স্টেপ-ইনডেক্স ফাইবারে সংকট কোণ, নিউমেরিক্যাল অ্যাপারচার ও এক্সেপটেন্স কোণ নির্ণয়",
        "• কোরের প্রতিসরাঙ্ক n_1 = 1.48\n• ক্ল্যাডিংয়ের প্রতিসরাঙ্ক n_2 = 1.46\n• বাতাসের প্রতিসরাঙ্ক n_0 = 1.0",
        "১. sin(theta_c) = n_2 / n_1\n২. NA = sqrt(n_1^2 - n_2^2)\n৩. theta_a = arcsin(NA / n_0)\n৪. Delta = (n_1 - n_2) / n_1",
        "ধাপ ১: সংকট কোণ (Critical Angle) নির্ণয়:\n"
        "   sin(theta_c) = 1.46 / 1.48 = 0.98648\n"
        "   => theta_c = arcsin(0.98648) = 80.57°।\n\n"
        "ধাপ ২: নিউমেরিক্যাল অ্যাপারচার (NA) নির্ণয়:\n"
        "   NA = sqrt(1.48^2 - 1.46^2) = sqrt(2.1904 - 2.1316) = sqrt(0.0588) = 0.2425।\n\n"
        "ধাপ ৩: সর্বোচ্চ এক্সেপটেন্স কোণ (Acceptance Angle):\n"
        "   sin(theta_a) = NA = 0.2425\n"
        "   => theta_a = arcsin(0.2425) = 14.03°।\n"
        "   (অর্থাৎ অক্ষের সাথে মোট এক্সেপটেন্স শঙ্কু কোণ 2*theta_a ≈ 28.06°)\n\n"
        "ধাপ ৪: আপেক্ষিক প্রতিসরাঙ্ক পার্থক্য (Delta):\n"
        "   Delta = (1.48 - 1.46) / 1.48 = 0.02 / 1.48 = 0.0135 (বা 1.35%)।",
        "• ভৌত তাৎপর্য: NA হলো ফাইবারের আলো ধরার ক্ষমতা। NA এর মান যত বেশি হবে, লেজার বা এলইডি থেকে আলো প্রবেশ করানো তত সহজ হবে; কিন্তু মাল্টিমোড ফাইবারে মোডাল ডিসপারশন বেশি হবে।"
    )

    add_h2("৩.২ অপটিক্যাল সিগন্যাল লস ও ডিসপারশন (Attenuation & Dispersion)", "sec_3_2")
    add_bullet("১. অ্যাটেনুয়েশন (Attenuation):", "দূরত্বের সাথে সাথে অপটিক্যাল সংকেতের তীব্রতা হ্রাস পাওয়া। এর প্রধান কারণগুলো হলো: রেলে স্ক্যাটারিং (Rayleigh Scattering: শর্ট ওয়েভলেংথে তীব্র), ইনফ্রারেড শোষণ লস (Absorption) এবং বাঁকানো জনিত বেন্ডিং লস (Macro & Micro-bending)। অপটিক্যাল ফাইবারে ১৫৫০ ন্যানোমিটার (1550 nm) উইন্ডোতে সর্বনিম্ন লস (প্রায় 0.2 dB/km) পাওয়া যায়।")
    add_bullet("২. ডিসপারশন (Dispersion):", "ফাইবার দিয়ে চলাচলের সময় লাইট পালস ছড়িয়ে যাওয়া বা প্রশস্ত হওয়া (Pulse Broadening)। এর ফলে ইন্টার-সিম্বল ইন্টারফারেন্স (ISI) তৈরি হয় এবং সর্বোচ্চ ডাটা রেট সীমিত হয়ে পড়ে।\n• মোডাল ডিসপারশন (Modal Dispersion): কেবল মাল্টিমোড ফাইবারে ঘটে।\n• ক্রোমাটিক ডিসপারশন (Chromatic Dispersion): উপাদান ও ওয়েভগাইডের কারণে ভিন্ন ভিন্ন তরঙ্গের আলোর গতিবেগের পার্থক্যের ফলে ঘটে (সিঙ্গেল মোড ফাইবারে এটি প্রধান)।\n• পিএমডি (Polarization Mode Dispersion - PMD): কোরের অপ্রতিসমতার কারণে দুই পোলারাইজেশন মোডের মধ্যে গতির পার্থক্য।")

    add_h2("৩.৩ আধুনিক অপটিক্যাল নেটওয়ার্কিং (WDM, DWDM, CWDM, EDFA)", "sec_3_3")
    add_body("একটি মাত্র অপটিক্যাল ফাইবার কেবলে একই সাথে টেরাবিট গতির ডাটা পরিবহনের জন্য ওয়েভলেংথ ডিভিশন মাল্টিপ্লেক্সিং (WDM) প্রযুক্তি ব্যবহৃত হয়।")
    add_bullet("CWDM (Coarse WDM):", "চ্যানেল স্পেসিং ২০ ন্যানোমিটার (২০ nm)। স্বল্প দূরত্বের মেট্রো নেটওয়ার্কে ব্যবহৃত হয়।")
    add_bullet("DWDM (Dense WDM):", "চ্যানেল স্পেসিং ০.৮ বা ০.৪ ন্যানোমিটার (৫০ বা ১০০ GHz)। লং-হল ব্যাকবোন ও সাবমেরিন কেবলে ব্যবহৃত হয়, যেখানে একটি ফাইবারে ৮০ থেকে ১৬০টি পর্যন্ত ভিন্ন চ্যানেল একযোগে পাঠানো যায়।")
    add_bullet("EDFA (Erbium-Doped Fiber Amplifier):", "১৫৫০ ন্যানোমিটার উইন্ডোতে অপটিক্যাল সিগন্যালকে কোনো ইলেকট্রিক্যাল কনভার্সন ছাড়াই সরাসরি অপটিক্যাল ডোমেইনে অ্যামপ্লিফাই করে।")

    add_h2("৩.৪ ট্রান্সমিশন হায়ারার্কি: PDH বনাম SDH (STM-1/4/16/64)", "sec_3_4")
    add_body("ডিজিটাল টেলিযোগাযোগের মূল ট্রান্সমিশন কাঠামো হলো সিনক্রোনাস ডিজিটাল হায়ারার্কি (SDH):")
    add_bullet("PDH (Plesiochronous Digital Hierarchy):", "প্রথাগত পদ্ধতি যেখানে প্রতিটি নোডের ক্লক সামান্য আলাদা (Plesiochronous) থাকত। ফলে লোয়ার চ্যানেলে ড্রপ বা ইনসার্ট করতে পুরো সিগন্যালকে ডিমাল্টিপ্লেক্স করতে হতো। স্ট্যান্ডার্ড: E1 = 2.048 Mbps (৩২টি চ্যানেল, প্রতিটি ৬৪ kbps)।")
    add_bullet("SDH (Synchronous Digital Hierarchy):", "পুরো নেটওয়ার্ক একক মাস্টার ক্লক দিয়ে সিঙ্ক্রোনাইজড। পয়েন্টার প্রযুক্তির কারণে যেকোনো লোয়ার অর্ডার সিগন্যালকে সরাসরি ড্রপ বা ইনসার্ট করা যায়।\n• STM-1 (Synchronous Transport Module 1) = 155.52 Mbps\n• STM-4 = 622.08 Mbps\n• STM-16 = 2.488 Gbps (≈ 2.5 Gbps)\n• STM-64 = 9.953 Gbps (≈ 10 Gbps)")

    add_h2("৩.৫ অপটিক্যাল লিংক পাওয়ার বাজেট ক্যালকুলেশন ও সমাধান", "sec_3_5")
    add_formula(
        "অপটিক্যাল লিংক পাওয়ার বাজেট সমীকরণ",
        "P_rx = P_tx - Total Losses - System Margin\n"
        "Total Losses = (alpha * L) + (N_s * Loss_s) + (N_c * Loss_c)",
        "• P_tx = অপটিক্যাল ট্রান্সমিটারের গড় আউটপুট পাওয়ার (dBm)\n"
        "• P_rx = অপটিক্যাল রিসিভারে প্রাপ্ত সংকেত পাওয়ার (dBm)\n"
        "• alpha = ফাইবার অ্যাটেনুয়েশন সহগ (dB/km)\n"
        "• L = ফাইবারের মোট দৈর্ঘ্য (km)\n"
        "• N_s, Loss_s = স্প্লাইস সংখ্যা ও প্রতি স্প্লাইসের লস (dB)\n"
        "• N_c, Loss_c = কানেক্টর সংখ্যা ও প্রতি কানেক্টরের লস (dB)\n"
        "• System Margin = ভবিষ্যতের ক্ষয়রোধে সেফটি মার্জিন (সাধারণত 3 dB থেকে 5 dB)",
        "সিস্টেমটি সফলভাবে কাজ করতে হলে রিসিভারের প্রাপ্ত পাওয়ার (P_rx) অবশ্যই রিসিভারের থ্রেশহোল্ড সেন্সিটিভিটি (Receiver Sensitivity) এর সমান বা বেশি হতে হবে।"
    )

    add_h3("গাণিতিক সমস্যা (বুয়েট ও বিটিসিএল পরীক্ষা):")
    add_body("একটি অপটিক্যাল ফাইবার লিঙ্কের দৈর্ঘ্য 50 কিমি, ফাইবার অ্যাটেনুয়েশন 0.25 dB/km, প্রতি 5 কিমি পর পর একটি করে ফিউশন স্প্লাইস (প্রতিটিতে লস 0.1 dB) এবং লিঙ্কের দুই প্রান্তে দুটি কানেক্টর (প্রতিটিতে লস 0.5 dB) রয়েছে। সেফটি মার্জিন 3 dB। ট্রান্সমিটারের আউটপুট পাওয়ার 0 dBm হলে রিসিভারের প্রয়োজনীয় সেন্সিটিভিটি (রিসিভড পাওয়ার) কত হতে হবে?")

    add_math_solution(
        "৫০ কিমি অপটিক্যাল ফাইবার লিঙ্কের পাওয়ার বাজেট ও রিসিভার সেন্সিটিভিটি নির্ণয়",
        "• লিংক দৈর্ঘ্য L = 50 km\n• ক্যাবল অ্যাটেনুয়েশন alpha = 0.25 dB/km\n• ড্রাম দৈর্ঘ্য L_drum = 5 km (প্রতি ৫ কিমি পর পর স্প্লাইস)\n• প্রতি স্প্লাইসে লস Loss_s = 0.1 dB\n• কানেক্টর সংখ্যা N_c = 2 (প্রতিটিতে লস 0.5 dB)\n• ট্রান্সমিটার পাওয়ার P_tx = 0 dBm\n• সেফটি মার্জিন = 3 dB",
        "P_rx = P_tx - [(alpha * L) + (N_s * Loss_s) + (N_c * Loss_c) + Safety Margin]",
        "ধাপ ১: ফাইবার ক্যাবল লস:\n"
        "   Cable Loss = alpha * L = 0.25 dB/km * 50 km = 12.5 dB\n\n"
        "ধাপ ২: স্প্লাইস সংখ্যা ও মোট স্প্লাইস লস:\n"
        "   ৫০ কিমি ক্যাবলে প্রতি ৫ কিমি পর স্প্লাইস বসলে মোট স্প্লাইস হবে:\n"
        "   N_s = (50 / 5) - 1 = 10 - 1 = 9 টি\n"
        "   [মনে রাখবেন: দুই প্রান্তে কোনো স্প্লাইস থাকে না, কানেক্টর থাকে]\n"
        "   Splice Loss = 9 * 0.1 dB = 0.9 dB\n\n"
        "ধাপ ৩: কানেক্টর লস:\n"
        "   Connector Loss = 2 * 0.5 dB = 1.0 dB\n\n"
        "ধাপ ৪: মোট সিস্টেম লস:\n"
        "   Total Loss = 12.5 dB + 0.9 dB + 1.0 dB + 3.0 dB (Margin) = 17.4 dB\n\n"
        "ধাপ ৫: রিসিভড পাওয়ার (সেন্সিটিভিটি):\n"
        "   P_rx = P_tx - Total Loss = 0 dBm - 17.4 dB = -17.4 dBm।",
        "• স্প্লাইস সংখ্যা গণনার সময় পরীক্ষার্থীরা সরাসরি ৫০ কে ৫ দিয়ে ভাগ করে ১০ লিখে ফেলেন। কিন্তু ৫০ কিমি লাইনে ৯টি জয়েন্ট বা স্প্লাইস পড়ে—এই সহজ বিষয়টি সতর্কতার সাথে মনে রাখতে হবে।"
    )

    add_back_to_toc()
    doc.add_page_break()

    # ==========================================
    # 8. CHAPTER 4: DATA COMMUNICATIONS & NETWORKING
    # ==========================================
    add_h1("অধ্যায় ০৪: ডাটা কমিউনিকেশন ও কম্পিউটার নেটওয়ার্কিং", "ch_04")
    add_body("আধুনিক টেলিযোগাযোগের সকল প্রযুক্তি (VoLTE, 5G, IMS, NGN, আইপি ব্যাকবোন) সম্পূর্ণভাবে কম্পিউটার নেটওয়ার্কিং ও প্যাকেট সুইচিংয়ের ওপর প্রতিষ্ঠিত। বিটিসিএল, বিটিআরসি ও মোবাইল অপারেটরের সকল টেকনিক্যাল নিয়োগ পরীক্ষায় আইপি নেটওয়ার্কিং থেকে সিংহভাগ প্রশ্ন করা হয়।")

    add_h2("৪.১ ওএসআই (OSI) ৭-লেয়ার মডেল বনাম টিসিপি/আইপি (TCP/IP) স্ট্যাক", "sec_4_1")
    add_body("ওএসআই (Open Systems Interconnection) হলো তাত্ত্বিক ৭-লেয়ার রেফারেন্স মডেল, আর টিসিপি/আইপি (TCP/IP) হলো ইন্টারনেটে বাস্তবায়িত কার্যকর ৪-লেয়ার প্রোটোকল স্ট্যাক।")

    add_img("images/osi_vs_tcpip.png", "ওএসআই (OSI) ৭-লেয়ার মডেল এবং টিসিপি/আইপি (TCP/IP) প্রোটোকল স্যুটের ম্যাপিং")

    add_bullet("Application Layer (লেয়ার ৭):", "ব্যবহারকারী ও অ্যাপ্লিকেশনের ইন্টারফেস (HTTP, HTTPS, DNS, DHCP, SIP, FTP)।")
    add_bullet("Presentation Layer (লেয়ার ৬):", "ডাটা ফরম্যাটিং, কম্প্রেশন ও এনক্রিপশন (SSL/TLS, JPEG, ASCII)।")
    add_bullet("Session Layer (লেয়ার ৫):", "সেশন তৈরি, নিয়ন্ত্রণ ও সমাপ্তি (RPC, NetBIOS)।")
    add_bullet("Transport Layer (লেয়ার ৪):", "এন্ড-টু-এন্ড ডাটা ট্রান্সফার ও নির্ভরযোগ্যতা। PDU: Segment। প্রোটোকল: TCP (কানেকশন ওরিয়েন্টেড, রিলায়েবল) এবং UDP (কানেকশনলেস, ফাস্ট, রিয়েল-টাইম ভয়েস/ভিডিও)।")
    add_bullet("Network Layer (লেয়ার ৩):", "লজিক্যাল অ্যাড্রেসিং (IP) এবং বেস্ট পাথ সিলেকশন (Routing)। PDU: Packet। ডিভাইস: Router, L3 Switch।")
    add_bullet("Data Link Layer (লেয়ার ২):", "ফিজিক্যাল অ্যাড্রেসিং (MAC), ফ্রেম ডেলিভারি ও এরর ডিটেকশন (CRC)। PDU: Frame। ডিভাইস: Switch, Bridge।")
    add_bullet("Physical Layer (লেয়ার ১):", "বাইনারি বিট স্ট্রিমিং ও ইলেকট্রিক্যাল/অপটিক্যাল সিগন্যাল। PDU: Bits। ডিভাইস: Hub, Repeater, Cables।")

    add_h2("৪.২ আইপি অ্যাড্রেসিং ও ক্লাসলেস নেটওয়ার্ক (IPv4 vs IPv6)", "sec_4_2")
    add_bullet("IPv4 (32-bit):", "ডটেড ডেসিমেল ফরম্যাট (যেমন: 192.168.1.1)। মোট অ্যাড্রেস প্রায় ৪.৩ বিলিয়ন।")
    add_bullet("IPv6 (128-bit):", "হেক্সাডেসিমেল কোলন ফরম্যাট (যেমন: 2001:0db8::ff00:0042:8329)। এটি অসীম সংখ্যক অ্যাড্রেস স্পেস ও উন্নত সিকিউরিটি (বিল্ট-ইন IPSec) নিশ্চিত করে।")

    add_h2("৪.৩ আইপি সাবনেটিং ও ভিএলএসএম (FLSM vs VLSM) - শর্টকাট টেকনিক", "sec_4_3")
    add_body("একটি বড় আইপি নেটওয়ার্ককে একাধিক ক্ষুদ্র ও কার্যকর সাব-নেটওয়ার্কে বিভক্ত করার প্রক্রিয়াকে সাবনেটিং বলে। এর মাধ্যমে আইপি অ্যাড্রেসের অপচয় রোধ হয় এবং ব্রডকাস্ট ডোমেইন ছোট করে নেটওয়ার্ক পারফরম্যান্স বাড়ানো হয়।")

    add_h3("গাণিতিক সমস্যা (বিসিএস ও বিটিসিএল প্রশ্ন):")
    add_body("আইপি ব্লক 192.168.10.0/24 দেওয়া আছে। একে সমান ৪টি সাবনেটে ভাগ করুন এবং প্রতিটি সাবনেটের নেটওয়ার্ক আইডি, ভ্যালিড হোস্ট রেঞ্জ ও ব্রডকাস্ট আইপি নির্ণয় করুন।")

    add_math_solution(
        "ক্লাস-সি নেটওয়ার্ককে সমান ৪টি সাবনেটে ভাগ (FLSM Subnetting)",
        "• প্রদেয় আইপি ব্লক: 192.168.10.0/24\n• তৈরি করতে হবে: ৪টি সাবনেট",
        "১. প্রয়োজনীয় ধার করা বিট: 2^s >= 4 => s = 2 বিট\n"
        "২. নতুন সাবনেট মাস্ক: /24 + 2 = /26 (255.255.255.192)\n"
        "৩. ম্যাজিক নাম্বার (Block Size) = 256 - 192 = 64\n"
        "৪. প্রতি সাবনেটে ব্যবহারযোগ্য হোস্ট = 2^(32-26) - 2 = 64 - 2 = 62 টি",
        "ধাপ ১: সাবনেটের শুরুর নেটওয়ার্ক আইডিগুলো বের করা (Block Size = 64):\n"
        "   ১ম সাবনেট নেটওয়ার্ক আইডি: 192.168.10.0\n"
        "   ২য় সাবনেট নেটওয়ার্ক আইডি: 192.168.10.64\n"
        "   ৩য় সাবনেট নেটওয়ার্ক আইডি: 192.168.10.128\n"
        "   ৪র্থ সাবনেট নেটওয়ার্ক আইডি: 192.168.10.192\n\n"
        "ধাপ ২: প্রতিটি সাবনেটের হোস্ট রেঞ্জ ও ব্রডকাস্ট আইপি সাজানো:\n"
        "   • সাবনেট ১: Network ID: 192.168.10.0/26 | Host: .1 - .62 | Broadcast: 192.168.10.63\n"
        "   • সাবনেট ২: Network ID: 192.168.10.64/26 | Host: .65 - .126 | Broadcast: 192.168.10.127\n"
        "   • সাবনেট ৩: Network ID: 192.168.10.128/26 | Host: .129 - .190 | Broadcast: 192.168.10.191\n"
        "   • সাবনেট ৪: Network ID: 192.168.10.192/26 | Host: .193 - .254 | Broadcast: 192.168.10.255",
        "• পরীক্ষার হলে দ্রুত হিসাব করার ট্রিকস: যেকোনো সাবনেটের ব্লক সাইজ বের করতে 256 থেকে সাবনেট মাস্কের শেষ অকটেটের মান বিয়োগ করবেন (যেমন 256 - 192 = 64)। পরবর্তী নেটওয়ার্ক আইডি থেকে ১ বিয়োগ করলেই আগের সাবনেটের ব্রডকাস্ট আইপি পাওয়া যায়।"
    )

    add_formula(
        "নেটওয়ার্ক প্যাকেট বিলম্ব সমীকরণ (Packet Transmission & Propagation Delay)",
        "Total Delay = Transmission Delay (T_trans) + Propagation Delay (T_prop)\n"
        "T_trans = L / R  [প্যাকেট বিট সংখ্যা / ব্যান্ডউইথ]\n"
        "T_prop = d / v  [লিঙ্কের দূরত্ব / মাধ্যমে আলোর গতিবেগ]",
        "• L = প্যাকেটের সাইজ (Bits)\n"
        "• R = ট্রান্সমিশন রেট বা লিংক ব্যান্ডউইথ (bps)\n"
        "• d = প্রেরক ও প্রাপকের মধ্যবর্তী ফিজিক্যাল দূরত্ব (মিটার বা কিমি)\n"
        "• v = ফাইবার বা তামার তারে সংকেতের বেগ (সাধারণত ফাইবারে আলোর গতি v ≈ 2 * 10^8 m/s)",
        "সহজবোধ্য ব্যাখ্যা: ট্রান্সমিশন বিলম্ব হলো প্যাকেটের সবগুলো বিটকে তারের ওপর পুশ বা তুলে দিতে ট্রান্সমিটারের কত সময় লাগে (ক্যাবলে ডাটা ইনজেকশন টাইম)। আর প্রোপাগেশন বিলম্ব হলো প্রথম বিটটি তারের এক প্রান্ত থেকে অপর প্রান্তে পৌঁছাতে কত সময় লাগে (ভ্রমণ সময়)।"
    )

    add_h3("গাণিতিক সমস্যা (বুয়েট ও বিসিএস পরীক্ষা): প্যাকেট ট্রান্সমিশন ও প্রোপাগেশন বিলম্ব")
    add_body("একটি 1000-বাইট দীর্ঘ ডাটা প্যাকেট একটি 10 Mbps অপটিক্যাল ফাইবার লিঙ্কের মাধ্যমে 2000 কিমি দূরবর্তী গন্তব্যে পাঠানো হচ্ছে। ফাইবার কোরে আলোর গতিবেগ 2 * 10^8 m/s হলে ট্রান্সমিশন বিলম্ব, প্রোপাগেশন বিলম্ব এবং মোট নেটওয়ার্ক বিলম্ব নির্ণয় করুন।")

    add_math_solution(
        "১০০০ বাইট প্যাকেটের ফাইবার অপটিক লিঙ্কে ট্রান্সমিশন ও প্রোপাগেশন ডিলে নির্ণয়",
        "• প্যাকেট সাইজ L = 1000 Bytes = 1000 * 8 = 8,000 Bits\n• ব্যান্ডউইথ R = 10 Mbps = 10 * 10^6 bps\n• দূরত্ব d = 2000 km = 2,000,000 meters\n• সংকেতের গতি v = 2 * 10^8 m/s",
        "১. T_trans = L / R\n২. T_prop = d / v\n৩. Total Delay = T_trans + T_prop",
        "ধাপ ১: ট্রান্সমিশন বিলম্ব গণনা:\n"
        "   T_trans = 8000 Bits / (10 * 10^6 bps) = 0.0008 সেকেন্ড = 0.8 মিলিসেকেন্ড (ms)।\n\n"
        "ধাপ ২: প্রোপাগেশন বিলম্ব গণনা:\n"
        "   T_prop = 2,000,000 m / (2 * 10^8 m/s) = 0.01 সেকেন্ড = 10 মিলিসেকেন্ড (ms)।\n\n"
        "ধাপ ৩: মোট বিলম্ব (Total Latency):\n"
        "   Total Delay = 0.8 ms + 10.0 ms = 10.8 মিলিসেকেন্ড (ms)।",
        "• সাধারণ ভুল: বাইট (Byte) কে ৮ দিয়ে গুণ করে বিট (Bit) এ রূপান্তর না করা। 1000 বাইট = 8000 বিট। গতিবেগ বায়ুর (3*10^8 m/s) নয়, ফাইবারের (2*10^8 m/s) বসাতে হবে।"
    )

    add_h2("৪.৪ রাউটিং প্রোটোকলস: OSPF, BGP, RIP ও সুইচিং (VLAN, STP)", "sec_4_4")
    add_bullet("OSPF (Open Shortest Path First):", "লিংক-স্টেট ইন্টেরিয়র রাউটিং প্রোটোকল (IGP)। এটি ডাইকস্ট্রা এসপিএফ (Dijkstra SPF) অ্যালগরিদম ব্যবহার করে এবং দ্রুত কনভার্জেন্স নিশ্চিত করে। এরিয়া ০ (Backbone Area) এর সাথে অন্যান্য এরিয়া যুক্ত থাকে।")
    add_bullet("BGP (Border Gateway Protocol):", "ইন্টারনেটের প্রধান এক্সটেরিয়র রাউটিং প্রোটোকল (EGP)। বিভিন্ন অটোনোমাস সিস্টেমের (AS) মধ্যে রাউটিং পাথ নির্বাচন করে। টেলিকম আইজিডব্লিউ (IGW) ও আইআইজি (IIG) লেয়ারে BGP অপরিহার্য।")
    add_bullet("VLAN (Virtual LAN - 802.1Q):", "একটি ফিজিক্যাল সুইচকে লজিক্যালি একাধিক ভার্চুয়াল নেটওয়ার্কে ভাগ করা। ব্রডকাস্ট ডোমেইন আলাদা করে সিকিউরিটি বৃদ্ধি করে।")

    add_h2("৪.৫ এমপিএলএস (MPLS) ও নেটওয়ার্ক সিকিউরিটি ফান্ডামেন্টালস", "sec_4_5")
    add_body("মাল্টিপ্রোটোকল লেবেল সুইচিং (MPLS) হলো এমন এক উচ্চগতির ব্যাকবোন রাউটিং কৌশল যেখানে রাউটারগুলো সম্পূর্ণ আইপি হেডার পরীক্ষা করার পরিবর্তে সংক্ষিপ্ত 'লেবেল' (20-bit Label) দেখে তাৎক্ষণিক প্যাকেট ফরোয়ার্ড করে। এর মূল উপাদানসমূহ হলো LER (Label Edge Router) এবং LSR (Label Switch Router)।")

    add_back_to_toc()
    doc.add_page_break()

    # ==========================================
    # 9. CHAPTER 5: MICROWAVE, SATELLITE & ANTENNA
    # ==========================================
    add_h1("অধ্যায় ০৫: মাইক্রোওয়েভ, স্যাটেলাইট ও অ্যান্টেনা কমিউনিকেশন", "ch_05")
    add_body("টেলিকম ব্যাকহলে দুর্গম পাহাড়ি বা নদীমাতৃক অঞ্চলে যেখানে ফাইবার টানা কঠিন বা ব্যয়বহুল, সেখানে পয়েন্ট-টু-পয়েন্ট মাইক্রোওয়েভ লিঙ্ক এবং স্যাটেলাইট কমিউনিকেশন মূল লাইফলাইন হিসেবে কাজ করে।")

    add_h2("৫.১ মাইক্রোওয়েভ রেডিও প্রোপাগেশন ও ফ্রেনেল জোন (Fresnel Zone)", "sec_5_1")
    add_body("মাইক্রোওয়েভ সিগন্যাল (সাধারণত ৩ GHz থেকে ৩০ GHz) সরলরেখায় সঞ্চালিত হয়, একে লাইন-অব-সাইট (Line-of-Sight - LOS) বলা হয়। তবে শুধু সরাসরি দৃষ্টিরেখাই যথেষ্ট নয়; ট্রান্সমিটার ও রিসিভারের মধ্যবর্তী সংকেত শক্তির সিংহভাগ একটি উপবৃত্তাকার (Ellipsoidal) স্থান দিয়ে প্রবাহিত হয় যাকে 'ফ্রেনেল জোন' (Fresnel Zone) বলে।")

    add_formula(
        "প্রথম ফ্রেনেল জোন ব্যাসার্ধ সমীকরণ (Fresnel Zone Radius)",
        "r_1 = 17.32 * sqrt((d_1 * d_2) / (f * D))  [মিটার]",
        "• d_1 = বাধা (Obstacle) থেকে ট্রান্সমিটারের দূরত্ব (km)\n"
        "• d_2 = বাধা থেকে রিসিভারের দূরত্ব (km)\n"
        "• D = d_1 + d_2 হলো মোট লিঙ্কের দূরত্ব (km)\n"
        "• f = মাইক্রোওয়েভ ফ্রিকোয়েন্সি (GHz)\n"
        "• r_1 = ১ম ফ্রেনেল জোনের ব্যাসার্ধ (মিটারে)",
        "প্রথম ফ্রেনেল জোনের ভেতর দিয়ে যাওয়া সিগন্যাল ফেজ-ইন (Constructive Phase) অবস্থায় থাকে। কিন্তু কোনো বাধা যদি এই জোনে ঢুকে পড়ে, তবে ডিফ্র্যাক্টেড সিগন্যাল ১৮০ ডিগ্রি বিপরীত ফেজে (Destructive Phase) পৌঁছে মূল সিগন্যালকে মারাত্মক দুর্বল বা বিলুপ্ত করে দিতে পারে।"
    )

    add_img("images/microwave_fresnel.png", "মাইক্রোওয়েভ লাইন-অব-সাইট (LOS) ও প্রথম ফ্রেনেল জোন ক্লিয়ারেন্স")

    add_callout("৬০% ক্লিয়ারেন্স নিয়ম", "রেডিও লিঙ্কে কোনো প্রকার ফেইডিং বা সিগন্যাল লস ছাড়াই যোগাযোগ স্থাপন করতে হলে প্রথম ফ্রেনেল জোনের ব্যাসার্ধের অন্তত ৬০% (0.6 * r1) সকল প্রকার গাছপালা, পাহাড় বা উঁচু দালানকোঠা মুক্ত (Obstacle Free) রাখতে হয়।")

    add_h2("৫.২ ফ্রি স্পেস পাথ লস (FSPL) ও লিঙ্ক বাজেট ক্যালকুলেশন", "sec_5_2")
    add_formula(
        "ফ্রি স্পেস পাথ লস (Free Space Path Loss - FSPL)",
        "FSPL (dB) = 32.44 + 20*log10(f_MHz) + 20*log10(d_km)\n"
        "অথবা, FSPL (dB) = 92.45 + 20*log10(f_GHz) + 20*log10(d_km)",
        "• f_MHz = ফ্রিকোয়েন্সি মেগাহার্টজে, f_GHz = ফ্রিকোয়েন্সি গিগাহার্টজে\n"
        "• d_km = দুটি অ্যান্টেনার মধ্যবর্তী দূরত্ব কিলোমিটারে",
        "এই সমীকরণটি ফ্রিস ট্রান্সমিশন সূত্র [FSPL = (4*pi*d / lambda)^2] থেকে প্রতিপাদিত। দূরত্ব দ্বিগুণ হলে (2d) অথবা ফ্রিকোয়েন্সি দ্বিগুণ হলে (2f) পাথ লস ঠিক 6 dB বৃদ্ধি পায় (20*log10(2) ≈ 6.02 dB)।"
    )

    add_formula(
        "রিসিভড সিগন্যাল লেভেল (RSL) ও ফেইড মার্জিন (Fade Margin)",
        "RSL (dBm) = P_tx + G_tx + G_rx - FSPL - Cable Losses\n"
        "Fade Margin = RSL - Receiver Threshold Sensitivity",
        "• P_tx = ট্রান্সমিটার রেডিও পাওয়ার (dBm)\n"
        "• G_tx, G_rx = যথাক্রমে ট্রান্সমিটার ও রিসিভার অ্যান্টেনার গেইন (dBi)\n"
        "• RSL = রিসিভারে পৌঁছানো নেট সিগন্যাল লেভেল (dBm)",
        "একটি নির্ভরযোগ্য মাইক্রোওয়েভ লিঙ্কের জন্য ফেইড মার্জিন ন্যূনতম 25 থেকে 35 dB হওয়া বাঞ্ছনীয়, যাতে ভারী বৃষ্টিপাত বা কুয়াশাতেও লিঙ্ক ডাউন না হয়।"
    )

    add_h3("গাণিতিক সমস্যা (টেলিকম আরএফ ও মাইক্রোওয়েভ নিয়োগ পরীক্ষা): সম্পূর্ণ লিঙ্ক বাজেট ও ফেইড মার্জিন")
    add_body("একটি 11 GHz মাইক্রোওয়েভ পয়েন্ট-টু-পয়েন্ট লিঙ্কের দুটি টাওয়ারের মধ্যবর্তী দূরত্ব 25 কিমি। ট্রান্সমিটার আউটপুট পাওয়ার 0.5 W (+27 dBm), ট্রান্সমিটার ও রিসিভার উভয়ের ডিশ অ্যান্টেনা গেইন 38 dBi। উভয় প্রান্তে ফিডার ক্যাবল ও কানেক্টর লস 2 dB করে। রিসিভারের থ্রেশহোল্ড সেন্সিটিভিটি -82 dBm।\n(ক) ট্রান্সমিটারের EIRP কত?\n(খ) ফ্রি স্পেস পাথ লস (FSPL) কত?\n(গ) রিসিভারে প্রাপ্ত সংকেত মাত্রা (RSL) কত?\n(ঘ) ফেইড মার্জিন কত এবং লিঙ্কটি বিটিআরসি বা টেলিকম স্ট্যান্ডার্ডে নির্ভরযোগ্য (Reliable) কি না মন্তব্য করুন।")

    add_math_solution(
        "১১ গিগাহার্টজ ২৫ কিমি মাইক্রোওয়েভ লিঙ্কের EIRP, FSPL, RSL ও ফেইড মার্জিন গণনা",
        "• ফ্রিকোয়েন্সি f = 11 GHz\n• দূরত্ব d = 25 km\n• ট্রান্সমিটার পাওয়ার P_tx = 27 dBm (0.5 W)\n• অ্যান্টেনা গেইন G_tx = G_rx = 38 dBi\n• ক্যাবল লস L_tx = L_rx = 2 dB\n• রিসিভার সেন্সিটিভিটি P_th = -82 dBm",
        "১. EIRP = P_tx - L_tx + G_tx\n২. FSPL = 92.45 + 20*log10(f_GHz) + 20*log10(d_km)\n৩. RSL = EIRP - FSPL + G_rx - L_rx\n৪. Fade Margin = RSL - P_th",
        "ধাপ ১: ট্রান্সমিটারের EIRP হিসাব:\n"
        "   EIRP = 27 dBm - 2 dB + 38 dBi = 63 dBm (বা প্রায় 2,000 Watts)।\n\n"
        "ধাপ ২: ফ্রি স্পেস পাথ লস (FSPL):\n"
        "   FSPL = 92.45 + 20*log10(11) + 20*log10(25)\n"
        "   • log10(11) = 1.0414 => 20 * 1.0414 = 20.83 dB\n"
        "   • log10(25) = 1.3979 => 20 * 1.3979 = 27.96 dB\n"
        "   FSPL = 92.45 + 20.83 + 27.96 = 141.24 dB।\n\n"
        "ধাপ ৩: রিসিভড সিগন্যাল লেভেল (RSL):\n"
        "   RSL = EIRP - FSPL + G_rx - L_rx\n"
        "   RSL = 63 dBm - 141.24 dB + 38 dBi - 2 dB = -42.24 dBm।\n\n"
        "ধাপ ৪: ফেইড মার্জিন (Fade Margin):\n"
        "   Fade Margin = RSL - P_th = -42.24 dBm - (-82.0 dBm) = +39.76 dB।\n\n"
        "ধাপ ৫: মন্তব্য ও নির্ভরযোগ্যতা বিচার:\n"
        "   টেলিকম গ্রেডে ৯৯.৯৯৯% ('Five Nines') বার্ষিক আপটাইম নিশ্চিত করতে ন্যূনতম ৩০ dB ফেইড মার্জিন প্রয়োজন। যেহেতু এখানে ফেইড মার্জিন ৩৯.৭৬ dB (> ৩০ dB), তাই ভারী বৃষ্টিপাত বা চরম আবহাওয়াতেও লিঙ্কটি সফলভাবে চালু থাকবে।",
        "• সতর্কবার্তা: EIRP হলো ট্রান্সমিটারের কার্যকর বিকিরণ ক্ষমতা। রিসিভারে পৌঁছাতে পৌঁছাতে ১৪১ dB লস হয়ে সিগন্যালটি -৪২ dBm এ নেমে আসে, যা রিসিভার সেন্সিটিভিটি -৮২ dBm এর চেয়ে অনেক শক্তিশালী।"
    )

    add_h2("৫.৩ অ্যান্টেনা থিওরি, গেইন, ডিরেক্টিভিটি ও VSWR", "sec_5_3")
    add_bullet("আইসোট্রপিক রেডিয়েটর (Isotropic Radiator):", "একটি তাত্ত্বিক বিন্দু উৎস যা সকল দিকে সমান শক্তিতে বিকিরণ ছড়ায় (Gain = 0 dBi)।")
    add_bullet("অ্যান্টেনা গেইন (Antenna Gain):", "নির্দিষ্ট অভিমুখে সিগন্যাল কেন্দ্রীভূত করার ক্ষমতা। সমীকরণ: G = 4*pi*A_e / lambda^2।")
    add_bullet("VSWR (Voltage Standing Wave Ratio):", "ফিডার ক্যাবল ও অ্যান্টেনার ইম্পিডেন্স অমিলের কারণে প্রতিফলিত তরঙ্গের অনুপাত। আদর্শ মান VSWR = 1.0 (কোনো প্রতিফলন নেই)। গ্রহণযোগ্য বাস্তব মান সাধারণত 1.2 থেকে 1.5 এর নিচে।")
    add_bullet("EIRP (Effective Isotropic Radiated Power):", "ট্রান্সমিটার পাওয়ার থেকে ক্যাবল লস বাদ দিয়ে অ্যান্টেনা গেইন যোগ করলে EIRP পাওয়া যায়: EIRP = P_tx - L_cable + G_tx।")

    add_h2("৫.৪ স্যাটেলাইট কমিউনিকেশন: LEO, MEO, GEO ও ট্রান্সপন্ডার", "sec_5_4")
    add_bullet("LEO (Low Earth Orbit):", "উচ্চতা ৫০০-১৫০০ কিমি। কম লেটেন্সি (২০-৪০ ms), স্টারলিংক বা ওয়ানওয়েবের মতো স্যাটেলাইট ইন্টারনেট কনস্টেলেশনে ব্যবহৃত।")
    add_bullet("MEO (Medium Earth Orbit):", "উচ্চতা ৫,০০০-২০,০০০ কিমি। জিপিএস (GPS) ও গ্যালিলিও নেভিগেশন স্যাটেলাইটে ব্যবহৃত।")
    add_bullet("GEO (Geostationary Earth Orbit):", "উচ্চতা প্রায় ৩৫,৭৮৬ কিমি (৩৬,০০০ কিমি)। পৃথিবীর ঘূর্ণন গতির সমান হওয়ায় পৃথিবী থেকে একে স্থির মনে হয়। বঙ্গবন্ধু স্যাটেলাইট-১ এই কক্ষপথে অবস্থিত। ব্রডকাস্টিং ও জাতীয় ব্যাকআপের জন্য আদর্শ।")

    add_back_to_toc()
    doc.add_page_break()

    # ==========================================
    # 10. CHAPTER 6: BANGLADESH TELECOM ARCHITECTURE
    # ==========================================
    add_h1("অধ্যায় ০৬: বাংলাদেশ টেলিকম ইকোসিস্টেম ও রেগুলেটরি ফ্রেমওয়ার্ক", "ch_06")
    add_body("বাংলাদেশের টেলিযোগাযোগ ব্যবস্থা একটি সুনির্দিষ্ট হায়ারার্কিক্যাল স্তরে বিন্যস্ত। আন্তর্জাতিক দীর্ঘ দূরত্ব টেলিযোগাযোগ নীতি (ILDTS Policy) এবং বিটিআরসি (BTRC) এর প্রণীত গাইডলাইন অনুযায়ী দেশের ভয়েস ও ডেটা ট্রাফিকের শৃঙ্খলা, নিরাপত্তা এবং রাজস্ব নিশ্চিত করার জন্য ভিন্ন ভিন্ন লাইসেন্সিং লেয়ার তৈরি করা হয়েছে।")

    add_h2("৬.১ বাংলাদেশ টেলিকম নেটওয়ার্ক আর্কিটেকচার (BTRC Framework)", "sec_6_1")
    add_body("বাংলাদেশে পরিচালিত টেলিযোগাযোগের কেন্দ্রীয় আর্কিটেকচারটি আন্তর্জাতিক স্তর থেকে শুরু করে গ্রাহক পর্যায় পর্যন্ত মোট ৫টি প্রধান স্তরে বিভক্ত:")

    add_img("images/bd_telecom_hierarchy.png", "বাংলাদেশ টেলিকম নেটওয়ার্ক হায়ারার্কি ও লাইসেন্সিং ফ্রেমওয়ার্ক (BTRC)")

    add_h2("৬.২ প্রধান লাইসেন্সিং লেয়ারসমূহ (IGW, ICX, IIG, NTTN, ANS, ITC, NIX)", "sec_6_2")
    add_bullet("১. ইন্টারন্যাশনাল গেটওয়ে (IGW - International Gateway):",
        "বিদেশ থেকে বাংলাদেশে আসা এবং বাংলাদেশ থেকে বিদেশে যাওয়া সকল প্রকার 'আন্তর্জাতিক ভয়েস কল' (Voice Traffic) হ্যান্ডেল করে। অবৈধ ভিওআইপি (VoIP) প্রতিরোধ, আন্তর্জাতিক কল ট্রাফিক রেকর্ড এবং ল'ফুল ইন্টারসেপশন (LI) নিশ্চিতকরণে এটি দায়বদ্ধ।"
    )
    add_bullet("২. ইন্টারকানেকশন এক্সচেঞ্জ (ICX - Interconnection Exchange):",
        "দেশের অভ্যন্তরীণ সকল অপারেটরের মধ্যবর্তী 'ভয়েস কল ট্রানজিট হাব' হিসেবে কাজ করে। যেমন: গ্রামীণফোন গ্রাহক যখন রবি বা বাংলালিংক গ্রাহককে কল দেন, তখন কলটি সরাসরি না গিয়ে বাধ্যতামূলকভাবে আইসিএক্স (ICX) সুইচ হয়ে অপর অপারেটরে পৌঁছায়।"
    )
    add_bullet("৩. ইন্টারন্যাশনাল ইন্টারনেট গেটওয়ে (IIG - International Internet Gateway):",
        "আন্তর্জাতিক ইন্টারনেট ট্রাফিকের প্রবেশদ্বার। এটি সাবমেরিন ক্যাবল (SEA-ME-WE-4, SEA-ME-WE-5) অথবা আইটিসি (ITC) অপারেটরদের কাছ থেকে ব্যান্ডউইথ সংগ্রহ করে দেশের আইএসপি (ISP) এবং মোবাইল নেটওয়ার্ক অপারেটরদের (MNO) কাছে সরবরাহ করে।"
    )
    add_bullet("৪. নেশনওয়াইড টেলিকমিউনিকেশন ট্রান্সমিশন নেটওয়ার্ক (NTTN):",
        "দেশের অপটিক্যাল ফাইবার ট্রান্সমিশন অবকাঠামোর মেরুদণ্ড (Backbone Infrastructure)। টেলিকম কোম্পানি বা আইএসপিগুলোকে বারবার রাস্তা খুঁড়ে নিজস্ব ক্যাবল বসানোর পরিবর্তে এনটিটিএন অপারেটরদের (যেমন: Fiber@Home, Summit Communications, BTCL) শেয়ার্ড ফাইবার নেটওয়ার্ক লিজিংয়ের মাধ্যমে ব্যবহার করতে হয়।"
    )
    add_bullet("৫. এক্সেস নেটওয়ার্ক সার্ভিসেস (ANS - Access Network Services):",
        "এটি গ্রাহক পর্যায়ের লেয়ার। এর মধ্যে রয়েছে মোবাইল নেটওয়ার্ক অপারেটর (Grameenphone, Robi, Banglalink, Teletalk), ফিক্সড ফোন বা পিএসটিএন (PSTN) এবং ইন্টারনেট সার্ভিস প্রোভাইডার (ISP)।")

    add_h2("৬.৩ সাবমেরিন ক্যাবল নেটওয়ার্ক ও বঙ্গবন্ধু স্যাটেলাইট-১ (BS-1)", "sec_6_3")
    add_body("বাংলাদেশ সাবমেরিন ক্যাবলস পিএলসি (BSCPLC) এর অধীনে বর্তমানে ২টি আন্তর্জাতিক সাবমেরিন ক্যাবল কনসোর্টিয়ামে বাংলাদেশ যুক্ত এবং ৩য়টি বাস্তবায়নাধীন:")
    add_bullet("SEA-ME-WE 4 (South East Asia-Middle East-Western Europe 4):", "২০০৫-২০০৬ সালে চালু। ল্যান্ডিং স্টেশন: ঝিলংজা, কক্সবাজার।")
    add_bullet("SEA-ME-WE 5:", "২০১৭ সালে চালু। ল্যান্ডিং স্টেশন: কুয়াকাটা, পটুয়াখালী। এটি অত্যন্ত উচ্চ ক্ষমতাসম্পন্ন ও লো-লেটেন্সির আন্তর্জাতিক রুট প্রদান করে।")
    add_bullet("SEA-ME-WE 6 (বাস্তবায়নাধীন):", "ল্যান্ডিং স্টেশন: কক্সবাজারের মাতারবাড়ী।")

    add_h3("বঙ্গবন্ধু স্যাটেলাইট-১ (Bangabandhu Satellite-1 - BS-1):")
    add_body("১২ মে ২০১৮ সালে যুক্তরাষ্ট্রের ফ্লোরিডার কেনেডি স্পেস সেন্টার থেকে স্পেসএক্স-এর ফ্যালকন ৯ রকেটের মাধ্যমে বাংলাদেশের প্রথম ভূস্থির যোগাযোগ উপগ্রহ উৎক্ষেপণ করা হয়।\n"
             "• অরবিটাল স্লট: ১১৯.১° পূর্ব দ্রাঘিমাংশ (119.1° E Geostationary Orbit)\n"
             "• মোট ট্রান্সপন্ডার: ৪০টি (২৬টি Ku-band এবং ১৪টি C-band)\n"
             "• প্রাথমিক গ্রাউন্ড কন্ট্রোল স্টেশন: সজীব ওয়াজেদ উপগ্রহ ভূকেন্দ্র, জয়দেবপুর, গাজীপুর\n"
             "• ব্যাকআপ ভূকেন্দ্র: বেতবুনিয়া, রাঙ্গামাটি")

    add_h2("৬.৪ বিটিআরসি স্পেকট্রাম বরাদ্দ, এমএনপি (MNP) ও QoS নীতিমালা", "sec_6_4")
    add_body("বিটিআরসি মোবাইল নেটওয়ার্কের জন্য বিভিন্ন গুরুত্বপূর্ণ ফ্রিকোয়েন্সি ব্যান্ড বরাদ্দ করেছে:")
    add_bullet("700 MHz Band (Band 28):", "উচ্চ কভারেজ এবং ইনডোর পেনিট্রেশনের জন্য অত্যন্ত কার্যকর (LTE/5G)।")
    add_bullet("900 MHz & 1800 MHz Band:", "ঐতিহ্যবাহী 2G GSM এবং রিফার্মড 4G LTE এর প্রধান ক্যারিয়ার ব্যান্ড।")
    add_bullet("2100 MHz Band (Band 1):", "3G ও 4G এর প্রধান ব্যান্ড।")
    add_bullet("2300 MHz & 2600 MHz Band:", "উচ্চগতির 4G LTE-TDD এবং 5G এর ক্যাপাসিটি লেয়ার।")
    add_bullet("3.5 GHz (C-band: 3300-3600 MHz):", "বাংলাদেশের 5G রোলআউটের প্রাইমারি ডেডিকেটেড স্পেকট্রাম ব্যান্ড।")

    add_callout("মোবাইল নম্বর পোর্টেবিলিটি (MNP) ও QoS নীতিমালা",
        "• এমএনপি (MNP): ২০১৮ সালে বাংলাদেশে চালু হয়, যার মাধ্যমে গ্রাহক তার মোবাইল নম্বর অপরিবর্তিত রেখে অন্য অপারেটরে সুইচ করতে পারেন।\n"
        "• বিটিআরসি QoS স্ট্যান্ডার্ড: কল ড্রপ রেট (Call Drop Rate) সর্বোচ্চ ২% এর নিচে এবং কল সেটআপ সাকসেস রেট (CSSR) সর্বনিম্ন ৯৭% এর উপরে থাকা বাধ্যতামূলক।"
    )

    add_back_to_toc()
    doc.add_page_break()

    # ==========================================
    # 11. CHAPTER 7: ELECTRICAL & ELECTRONICS FOR TELECOM
    # ==========================================
    add_h1("অধ্যায় ০৭: ইলেকট্রিক্যাল ও ইলেকট্রনিক্স ফান্ডামেন্টালস", "ch_07")
    add_body("টেলিকম প্রকৌশলী নিয়োগ পরীক্ষায় সার্কিট থিওরেম, বেসিক ইলেকট্রনিক্স এবং টেলিকম পাওয়ার সিস্টেম থেকে নিয়মিত প্রশ্ন আসে। বিশেষ করে টেলিকম সাইটের -৪৮ ভোল্ট ডিসি পাওয়ার প্ল্যান্ট ও ব্যাটারি ব্যাকআপ ইঞ্জিনিয়ারিং অত্যন্ত গুরুত্বপূর্ণ।")

    add_h2("৭.১ সার্কিট অ্যানালাইসিস থিওরেমস (KVL, KCL, Thevenin, Norton)", "sec_7_1")
    add_bullet("কার্শফের সূত্রাবলী (Kirchhoff's Laws):", "KCL (কারেন্ট সূত্র: নোডে মোট আগত কারেন্ট = নির্গত কারেন্ট) চার্জের সংরক্ষণশীলতা নীতি এবং KVL (ভোল্টেজ সূত্র: বদ্ধ লুপে ভোল্টেজ ড্রপসমূহের বীজগাণিতিক যোগফল শূন্য) শক্তির সংরক্ষণশীলতা নীতি মেনে চলে।")
    add_bullet("থেভেনিন থিওরেম (Thevenin's Theorem):", "যেকোনো জটিল রৈখিক দ্বিমুখী সার্কিটকে দুটি টার্মিনালের সাপেক্ষে একটি একক সমতুল্য ভোল্টেজ সোর্স (V_th) এবং তার সাথে সিরিজে একটি সমতুল্য রোধ (R_th) দ্বারা প্রতিস্থাপন করা যায়।")
    add_bullet("ম্যাক্সিমাম পাওয়ার ট্রান্সফার থিওরেম:", "লোড রোধের মান যখন অভ্যন্তরীণ থেভেনিন রোধের সমান হয় (R_L = R_th), তখন সোর্স থেকে লোডে সর্বাধিক ক্ষমতা সঞ্চালিত হয়। এই অবস্থায় দক্ষতা হয় ৫০%।")

    add_h2("৭.২ এসি ফান্ডামেন্টালস, RLC সার্কিট ও পাওয়ার ফ্যাক্টর", "sec_7_2")
    add_body("টেলিকম পাওয়ারে পাওয়ার ফ্যাক্টর (Power Factor: cos θ) অত্যন্ত গুরুত্বপূর্ণ। বিশুদ্ধ রেজিস্টিভ লোডে পাওয়ার ফ্যাক্টর ১.০। ইন্ডাক্টিভ লোডের কারণে ল্যাগিং পাওয়ার ফ্যাক্টর তৈরি হয়, যা ক্যাপাসিটর ব্যাংক দিয়ে উন্নত করা হয়।")
    add_formula(
        "সিরিজ RLC সার্কিটের রেজোন্যান্স ফ্রিকোয়েন্সি",
        "f_r = 1 / (2 * pi * sqrt(L * C))  [Hz]",
        "• L = ইন্ডাক্ট্যান্স (Henry - H)\n• C = ক্যাপাসিট্যান্স (Farad - F)",
        "রেজোন্যান্সে ইন্ডাক্টিভ রিঅ্যাক্ট্যান্স ও ক্যাপাসিটিভ রিঅ্যাক্ট্যান্স সমান হয়ে পরস্পরের প্রভাব বাতিল করে (X_L = X_C)। ফলে সার্কিটের মোট ইম্পিডেন্স সর্বনিম্ন হয়ে কেবল বিশুদ্ধ রোধে (Z = R) পরিণত হয় এবং সার্কিটে সর্বোচ্চ বিদ্যুৎ প্রবাহিত হয়।"
    )

    add_h2("৭.৩ সেমিকন্ডাক্টর ও অপারেশনাল অ্যামপ্লিফায়ার (Op-Amp)", "sec_7_3")
    add_bullet("আইডিয়াল অপ-অ্যাম্প (Ideal Op-Amp) এর বৈশিষ্ট্য:", "১. ইনফিনিট ওপেন-লুপ গেইন (A_OL = ∞), ২. ইনফিনিট ইনপুট ইম্পিডেন্স (R_in = ∞), ৩. জিরো আউটপুট ইম্পিডেন্স (R_out = 0), ৪. অসীম ব্যান্ডউইথ এবং ৫. জিরো অফসেট ভোল্টেজ।")
    add_bullet("ভার্চুয়াল শর্ট ও ভার্চুয়াল গ্রাউন্ড কনসেপ্ট:", "অপ-অ্যাম্পে নেগেটিভ ফিডব্যাক থাকলে ইনভার্টিং ও নন-ইনভার্টিং টার্মিনালের বিভব সমান থাকে (V+ = V-)। নন-ইনভার্টিং টার্মিনাল গ্রাউন্ডে থাকলে ইনভার্টিং টার্মিনালকেও কার্যত গ্রাউন্ড (Virtual Ground) বিবেচনা করা হয়।")

    add_formula(
        "অপারেশনাল অ্যামপ্লিফায়ারের ক্লোজড-লুপ গেইন সমীকরণ",
        "ইনভার্টিং কনফিগারেশন: V_out = -(R_f / R_1) * V_in  => A_v = -R_f / R_1\n"
        "নন-ইনভার্টিং কনফিগারেশন: V_out = (1 + R_f / R_1) * V_in  => A_v = 1 + R_f / R_1",
        "• R_f = ফিডব্যাক রেজিস্টর (Feedback Resistor)\n"
        "• R_1 = ইনপুট রেজিস্টর (Input Resistor)\n"
        "• A_v = ভোল্টেজ গেইন (Voltage Gain)",
        "সহজবোধ্য ব্যাখ্যা: ইনভার্টিং সার্কিটে আউটপুট সিগন্যাল ইনপুটের বিপরীত ফেজে (১৮০° ফেজ শিফট) থাকে বলে গেইনের আগে ঋণাত্মক (-) চিহ্ন থাকে। নন-ইনভার্টিং সার্কিটে কোনো ফেজ পরিবর্তন হয় না এবং এর গেইন সর্বদা ১ এর চেয়ে বেশি (A_v >= 1) হয়। যদি R_f = 0 এবং R_1 = অসীম হয়, তবে A_v = 1 হয়, যাকে ভোল্টেজ ফলোয়ার (Buffer) বলা হয়।"
    )

    add_h3("গাণিতিক সমস্যা: নন-ইনভার্টিং অপ-অ্যাম্প সার্কিটের আউটপুট ভোল্টেজ ও গেইন নির্ণয়")
    add_body("একটি নন-ইনভার্টিং অপ-অ্যাম্প সার্কিটে ইনপুট ভোল্টেজ Vin = 0.5V, ইনপুট রোধ R1 = 10 kΩ এবং ফিডব্যাক রোধ Rf = 90 kΩ।\n(ক) সার্কিটের ক্লোজড-লুপ ভোল্টেজ গেইন (Av) কত?\n(খ) আউটপুট ভোল্টেজ (Vout) কত?")

    add_math_solution(
        "নন-ইনভার্টিং অপ-অ্যাম্প সার্কিটের ক্লোজড-লুপ গেইন ও আউটপুট ভোল্টেজ গণনা",
        "• ইনপুট ভোল্টেজ V_in = 0.5 V\n• ইনপুট রেজিস্ট্যান্স R_1 = 10 kΩ\n• ফিডব্যাক রেজিস্ট্যান্স R_f = 90 kΩ",
        "১. A_v = 1 + (R_f / R_1)\n২. V_out = A_v * V_in",
        "ধাপ ১: ক্লোজড-লুপ ভোল্টেজ গেইন হিসাব:\n"
        "   A_v = 1 + (90 kΩ / 10 kΩ) = 1 + 9 = 10।\n\n"
        "ধাপ ২: আউটপুট ভোল্টেজ হিসাব:\n"
        "   V_out = A_v * V_in = 10 * 0.5 V = 5.0 Volts।",
        "• সাধারণ ভুল: অনেকেই নন-ইনভার্টিং সূত্রে ভুল করে শুধু Rf/R1 লিখে ফেলেন (যা ইনভার্টিং এর জন্য প্রযোজ্য)। নন-ইনভার্টিং এ অবশ্যই '1 +' যোগ করতে হবে।"
    )

    add_h2("৭.৪ টেলিকম -৪৮ ভোল্ট ডিসি (-48V DC) পাওয়ার সিস্টেম ও ব্যাটারি ব্যাকআপ", "sec_7_4")
    add_body("টেলিযোগাযোগের ইতিহাসে এবং বর্তমান বিটিএস টাওয়ার ও কোর ডাটা সেন্টারে স্ট্যান্ডার্ড অপারেটিং ভোল্টেজ হলো -৪৮ ভোল্ট ডিসি (-48V DC)।")

    add_callout("কেন টেলিকমে -৪৮ ভোল্ট ডিসি ব্যবহৃত হয়?",
        "১. ক্যাথোডিক প্রোটেকশন ও ক্ষয়রোধ (Corrosion Prevention): পজিটিভ টার্মিনালকে গ্রাউন্ড (+48V Grounded) করে নেগেটিভ টার্মিনালকে -48V হিসেবে ব্যবহার করলে মাটির আর্দ্রতায় তামার তারের অ্যানোডিক ক্ষয় (Galvanic Corrosion) সম্পূর্ণ বন্ধ হয়।\n"
        "২. মানবদেহের নিরাপত্তা (Safety SELV): ৬০ ভোল্টের নিচের ডিসি ভোল্টেজ মানুষের স্পর্শে প্রাণঘাতী শক তৈরি করে না (Safety Extra Low Voltage)।\n"
        "৩. তাৎক্ষণিক ব্যাটারি ব্যাকআপ: কোনো ইনভার্টারের বিলম্ব ছাড়াই সরাসরি ব্যাটারি ব্যাংক ডিসি বাসের সাথে ফ্লোট চার্জে যুক্ত থাকে, ফলে গ্রিড বিদ্যুৎ গেলে এক মিলিসেকেন্ডও বিদ্যুৎ বিভ্রাট হয় না।"
    )

    add_img("images/telecom_dc_power.png", "টেলিকম -৪৮ ভোল্ট ডিসি পাওয়ার প্ল্যান্ট ও ব্যাটারি ব্যাকআপ আর্কিটেকচার")

    add_h3("ব্যাটারি ব্যাকআপ ক্যালকুলেশন অঙ্ক:")
    add_body("একটি বিটিএস (BTS) সাইটের মোট ডিসি লোড 2500 ওয়াট। -48V ডিসি সিস্টেমে 4 ঘণ্টার নিরবচ্ছিন্ন ব্যাকআপ নিশ্চিত করতে 12V, 200Ah ব্যাটারি দিয়ে কতটি ব্যাটারি এবং কীভাবে যুক্ত করতে হবে? (ধরি, ব্যাটারি ডিসচার্জ ডেপথ DoD = 80%)।")

    add_math_solution(
        "বিটিএস সাইটের ডিসি লোডের জন্য ব্যাটারি ব্যাংক সাইজিং",
        "• পাওয়ার লোড P = 2500 Watts\n• অপারেটিং ভোল্টেজ V = 48 Volts DC\n• ব্যাকআপ সময় T = 4 Hours\n• উপলব্ধ ব্যাটারি স্পেসিফিকেশন: 12V, 200Ah\n• ডেপথ অব ডিসচার্জ DoD = 80% = 0.80",
        "১. লোড কারেন্ট: I = P / V\n২. প্রয়োজনীয় সক্ষমতা (Ah) = (I * T) / DoD\n৩. সিরিজ ব্যাটারি = V_system / V_battery\n৪. প্যারালাল স্ট্রিং = Required_Ah / Battery_Ah",
        "ধাপ ১: ডিসি লোড কারেন্ট হিসাব:\n"
        "   I = 2500 W / 48 V = 52.083 Amperes\n\n"
        "ধাপ ২: ৪ ঘণ্টার জন্য প্রয়োজনীয় অ্যাম্পিয়ার-আওয়ার (Ah) হিসাব:\n"
        "   প্রয়োজনীয় Ah = (52.083 A * 4 hours) / 0.80 = 208.33 / 0.80 = 260.42 Ah\n\n"
        "ধাপ ৩: ভোল্টেজ ম্যাচিংয়ে সিরিজ ব্যাটারি সংখ্যা:\n"
        "   -48V সিস্টেম ভোল্টেজ তৈরি করতে ১২ ভোল্টের ব্যাটারি সিরিজে বসাতে হবে:\n"
        "   N_series = 48V / 12V = ৪টি ব্যাটারি (একটি স্ট্রিং = 48V, 200Ah)\n\n"
        "ধাপ ৪: কারেন্ট ক্যাপাসিটি ম্যাচিংয়ে প্যারালাল স্ট্রিং সংখ্যা:\n"
        "   প্রয়োজনীয় মোট 260.42 Ah পূরণ করতে প্যারালাল স্ট্রিং প্রয়োজন:\n"
        "   N_parallel = 260.42 Ah / 200 Ah = 1.30 ≈ ২টি স্ট্রিং (2 * 200Ah = 400Ah)\n\n"
        "ধাপ ৫: মোট প্রয়োজনীয় ব্যাটারি সংখ্যা:\n"
        "   Total Batteries = 4 (সিরিজ) * 2 (প্যারালাল স্ট্রিং) = ৮টি ব্যাটারি।",
        "• ব্যাটারির ডেপথ অব ডিসচার্জ (DoD) হিসাব করতে ভুলবেন না। ১০০% ব্যাটারি ডিসচার্জ করলে ব্যাটারির লাইফ সাইকেল মারাত্মক ক্ষতিগ্রস্ত হয়, তাই সাধারণত ৮০% DoD বিবেচনায় নিয়ে ক্যাপাসিটি বেশি ধরা হয়।"
    )

    add_back_to_toc()
    doc.add_page_break()

    # ==========================================
    # 12. CHAPTER 8: MATHEMATICAL PROBLEMS & FORMULA HANDBOOK
    # ==========================================
    add_h1("অধ্যায় ০৮: টেলিকম গণিত ও গুরুত্বপূর্ণ ফর্মুলা শীট", "ch_08")
    add_body("টেলিকম চাকরির লিখিত ও প্রিলিমিনারি পরীক্ষায় প্রার্থীদের মেধার মূল পার্থক্য তৈরি করে নির্ভুল ও দ্রুত গাণিতিক সমস্যা সমাধানের দক্ষতা। এই অধ্যায়টি দুটি অংশে বিভক্ত: প্রথমে এক নজরে রিভিশনের জন্য সকল অত্যাবশ্যকীয় সূত্রের একটি মাস্টার টেবিল এবং পরবর্তীতে বিগত ১০ বছরের বিভিন্ন নিয়োগ পরীক্ষায় আসা সর্বাধিক গুরুত্বপূর্ণ ৮টি টাইপের স্ট্যান্ডার্ড সমস্যার নিখুঁত ও সহজবোধ্য স্টেপ-বাই-স্টেপ সমাধান।")

    add_h2("৮.১ এক নজরে সকল অত্যাবশ্যকীয় টেলিকম সূত্র ও একক (Formula Cheatsheet)", "sec_8_1")
    add_body("পরীক্ষার ঠিক পূর্বে এক ঝলকে চোখ বুলিয়ে নেওয়ার জন্য সকল প্রধান গাণিতিক সূত্রের সংকলন:")

    tbl_cheat = doc.add_table(rows=1, cols=3)
    tbl_cheat.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_hdrs = ["টপিক / ক্ষেত্র", "মূল গাণিতিক সূত্রাবলী", "প্যারামিটার ও আদর্শ একক"]
    c_widths = [Inches(1.5), Inches(2.6), Inches(2.4)]
    for i, h in enumerate(c_hdrs):
        tbl_cheat.rows[0].cells[i].text = h
        tbl_cheat.rows[0].cells[i].width = c_widths[i]
        tcPr = tbl_cheat.rows[0].cells[i]._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="0A2540"/>'))
        p = tbl_cheat.rows[0].cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            set_font(r, size_pt=10, bold=True, color=RGBColor(255, 255, 255))

    cheatsheet_data = [
        ("১. স্যাম্পলিং ও ক্যাপাসিটি", "f_s >= 2 * f_m\nC = B * log2(1 + S/N)\nC_nyq = 2 * B * log2(M)", "f_m, B: Hz\nS/N: লিনিয়ার রেশিও\nC: bps"),
        ("২. পিসিএম (PCM)", "R_b = n * f_s = n * 2 * f_m\nSQNR (dB) = 6.02 * n + 1.76", "n: bits/sample\nL = 2^n: লেভেল\nSQNR: dB"),
        ("৩. সেলুলার ক্লাস্টার", "N = i^2 + i*j + j^2\nQ = D / R = sqrt(3 * N)\nS/I = (sqrt(3N))^n / 6", "N: ক্লাস্টার সাইজ\nD: কো-চ্যানেল দূরত্ব\nn: পাথ লস এক্সপোনেন্ট"),
        ("৪. ট্রাফিক (Erlang)", "A = lambda * H\nTrunk Capacity = N * A_user", "lambda: calls/hour\nH: hours (মিনিট/৬০)\nA: Erlangs"),
        ("৫. অপটিক্যাল ফাইবার", "NA = sqrt(n_1^2 - n_2^2) = sin(theta_a)\nsin(theta_c) = n_2 / n_1\nDelta = (n_1 - n_2) / n_1", "n_1: কোর প্রতিসরাঙ্ক\nn_2: ক্ল্যাডিং প্রতিসরাঙ্ক\nNA: ডাইমেনশনলেস"),
        ("৬. অপটিক্যাল পাওয়ার", "P_rx = P_tx - (alpha * L + N_s*L_s + N_c*L_c) - Margin", "P_tx, P_rx: dBm\nalpha: dB/km\nL: km, Margin: dB"),
        ("৭. মাইক্রোওয়েভ ও FSPL", "FSPL = 92.45 + 20*log10(f_GHz) + 20*log10(d_km)\nr_1 = 17.32 * sqrt((d_1*d_2) / (f*D))", "f: GHz, d: km\nr_1: meters\nক্লিয়ারেন্স >= 0.6 * r_1"),
        ("৮. মাইক্রোওয়েভ বাজেট", "EIRP = P_tx - L_tx + G_tx\nFade Margin = RSL - Sensitivity", "EIRP, RSL: dBm\nG: dBi, Margin >= 30 dB"),
        ("৯. নেটওয়ার্ক বিলম্ব", "Total = (L / R) + (d / v)", "L: bits, R: bps\nd: meters, v: 2*10^8 m/s"),
        ("১০. ইলেকট্রিক্যাল ও ডিসি", "f_r = 1 / (2*pi*sqrt(L*C))\nBattery Ah = (P * T) / (V * DoD)", "L: Henry, C: Farad\nV: 48V DC, DoD: 0.80")
    ]

    for row_data in cheatsheet_data:
        row_cells = tbl_cheat.add_row().cells
        for i, val in enumerate(row_data):
            row_cells[i].text = val
            row_cells[i].width = c_widths[i]
            p = row_cells[i].paragraphs[0]
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.15
            for r in p.runs:
                is_bold = (i == 0)
                set_font(r, size_pt=9.5, bold=is_bold, color=COLOR_PRIMARY if is_bold else COLOR_TEXT)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 8.2: Types 1-2
    add_h2("৮.২ সিগন্যাল, স্যাম্পলিং ও চ্যানেল ক্যাপাসিটি সমাধান (Types 1-2)", "sec_8_2")
    
    add_h3("টাইপ ০১: কম্পোজিট এনালগ সিগন্যালের নাইকুইস্ট স্যাম্পলিং রেট ও ইন্টারভাল")
    add_body("সিগন্যাল s(t) = 5*cos(1000*pi*t) + 2*sin(4000*pi*t) - 3*cos(6000*pi*t)। এর নাইকুইস্ট স্যাম্পলিং রেট ও নাইকুইস্ট ইন্টারভাল কত?")
    
    add_math_solution(
        "কম্পোজিট এনালগ সিগন্যালের নাইকুইস্ট স্যাম্পলিং রেট ও ইন্টারভাল",
        "• s(t) = 5*cos(1000*pi*t) + 2*sin(4000*pi*t) - 3*cos(6000*pi*t)",
        "১. ওমেগা = 2*pi*f => f = ওমেগা / (2*pi)\n২. f_s (Nyquist) = 2 * f_max\n৩. T_s (Nyquist Interval) = 1 / f_s",
        "ধাপ ১: ৩টি সাইনোসয়ডাল কম্পোনেন্টের ফ্রিকোয়েন্সি আলাদা করা:\n"
        "   • ১ম পদ: 2*pi*f1 = 1000*pi => f1 = 1000/2 = 500 Hz\n"
        "   • ২য় পদ: 2*pi*f2 = 4000*pi => f2 = 4000/2 = 2000 Hz\n"
        "   • ৩য় পদ: 2*pi*f3 = 6000*pi => f3 = 6000/2 = 3000 Hz\n\n"
        "ধাপ ২: সর্বোচ্চ ফ্রিকোয়েন্সি (f_max) নির্ধারণ:\n"
        "   f_max = max(500, 2000, 3000) = 3000 Hz = 3 kHz\n\n"
        "ধাপ ৩: নাইকুইস্ট স্যাম্পলিং রেট:\n"
        "   f_s = 2 * f_max = 2 * 3000 Hz = 6000 Hz (বা 6 kHz)\n\n"
        "ধাপ ৪: নাইকুইস্ট ইন্টারভাল (Nyquist Interval):\n"
        "   T_s = 1 / f_s = 1 / 6000 sec = 0.0001667 sec = 0.167 ms (মিলি-সেকেন্ড)।",
        "• একাধিক ফ্রিকোয়েন্সির যোগফল থাকলে সর্বদা সর্বোচ্চ ফ্রিকোয়েন্সি (f_max) নিয়ে নাইকুইস্ট রেট বের করতে হয়। গুণফল থাকলে ফ্রিকোয়েন্সি যোগ হয়ে যায়।"
    )

    add_h3("টাইপ ০২: শ্যানন ক্যাপাসিটি ও প্রয়োজনীয় সিগন্যালিং লেভেল (M) নির্ণয়")
    add_body("একটি 3 kHz টেলিফোন লাইনের SNR = 20 dB। লাইনটির শ্যানন ক্যাপাসিটি কত এবং নয়েজহীন চ্যানেলে এই ক্যাপাসিটি পেতে কত লেভেল সিগন্যালিং প্রয়োজন?")
    
    add_math_solution(
        "শ্যানন চ্যানেল ক্যাপাসিটি ও নয়েজহীন চ্যানেলের লেভেল সংখ্যা হিসাব",
        "• ব্যান্ডউইথ B = 3000 Hz\n• SNR_dB = 20 dB",
        "১. S/N = 10^(SNR_dB / 10) = 100\n২. C = B * log2(1 + S/N)\n৩. C_nyq = 2 * B * log2(M) = C",
        "ধাপ ১: SNR লিনিয়ার রেশিও:\n"
        "   S/N = 10^(20/10) = 100\n\n"
        "ধাপ ২: শ্যানন ক্যাপাসিটি:\n"
        "   C = 3000 * log2(101) = 3000 * (ln(101) / ln(2)) = 3000 * 6.6582 = 19,974 bps ≈ 19.97 kbps।\n\n"
        "ধাপ ৩: নাইকুইস্টে সমপরিমাণ ডাটা রেট পেতে প্রয়োজনীয় লেভেল M:\n"
        "   2 * 3000 * log2(M) = 19,974\n"
        "   => log2(M) = 3.329 => M = 2^3.329 ≈ 10.05 ≈ 11 টি লেভেল।",
        "• শ্যাননের সূত্রে SNR এর মান লিনিয়ার এককে বসাতে হবে। 20 dB মানে 100, সরাসরি 20 বসানো মারাত্মক ভুল।"
    )

    # Section 8.3: Types 3-4
    add_h2("৮.৩ সেলুলার ট্রাফিক (Erlang), ক্লাস্টার ও ক্যাপাসিটি সমাধান (Types 3-4)", "sec_8_3")
    
    add_h3("টাইপ ০৩: সেলুলার স্পেকট্রাম বণ্টন ও মোট এককালীন চ্যানেল ক্যাপাসিটি")
    add_body("একটি সেলুলার সিস্টেমে মোট 33 MHz স্পেকট্রাম বরাদ্দ আছে। প্রতিটি ডুপ্লেক্স চ্যানেলের ব্যান্ডউইথ 50 kHz। কন্ট্রোল চ্যানেলে 1 MHz সংরক্ষিত। ক্লাস্টার সাইজ N = 7 এবং সমগ্র সিস্টেমে 28 টি সেল থাকলে সিস্টেমটির মোট ক্যাপাসিটি কত?")

    add_math_solution(
        "সেলুলার ফ্রিকোয়েন্সি রিইউজ ও মোট ধারণক্ষমতা নির্ণয়",
        "• মোট স্পেকট্রাম = 33 MHz, কন্ট্রোল = 1 MHz\n• চ্যানেল সাইজ = 50 kHz, N = 7, মোট সেল = 28",
        "১. উপলব্ধ ভয়েস চ্যানেল C_total = (33 - 1) MHz / 50 kHz\n২. প্রতি সেলে চ্যানেল K = C_total / N\n৩. সিস্টেম ক্যাপাসিটি = K * মোট সেল",
        "ধাপ ১: মোট ডুপ্লেক্স চ্যানেল:\n"
        "   C_total = 32,000 kHz / 50 kHz = 640 টি চ্যানেল।\n\n"
        "ধাপ ২: প্রতি সেলে চ্যানেল বরাদ্দ (K):\n"
        "   K = 640 / 7 = 91.42 ≈ 91 টি চ্যানেল প্রতি সেলে।\n\n"
        "ধাপ ৩: সমগ্র সিস্টেমের ক্যাপাসিটি:\n"
        "   মোট ক্লাস্টার = 28 / 7 = 4 টি।\n"
        "   মোট ক্যাপাসিটি = 4 * 640 = 2,560 টি একযোগে সচল কল।",
        "• ডুপ্লেক্স চ্যানেল হিসাব করতে ভুলবেন না। স্পেকট্রাম থেকে কন্ট্রোল চ্যানেলের ব্যান্ডউইথ বিয়োগ দিতে হবে।"
    )

    add_h3("টাইপ ০৪: সেলুলার বেস স্টেশনের মোট ট্রাফিক ইনটেনসিটি (Erlang) নির্ণয়")
    add_body("একটি সেলুলার বেস স্টেশনে পিক আওয়ারে প্রতি ঘণ্টায় গড়ে 1800 টি কল আসে এবং প্রতিটি কলের গড় স্থায়িত্বকাল 2 মিনিট 30 সেকেন্ড হলে সেলের ট্রাফিক লোড নির্ণয় করুন।")
    
    add_math_solution(
        "সেলুলার বেস স্টেশনের মোট ট্রাফিক ইনটেনসিটি (Erlang) নির্ণয়",
        "• কল রেট lambda = 1800 calls/hour\n• গড় কল ডিউরেশন H = 2 মিনিট 30 সেকেন্ড = 2.5 মিনিট",
        "A = lambda * H  [Erlang]",
        "ধাপ ১: কল ডিউরেশনকে মিনিট থেকে ঘণ্টায় রূপান্তর:\n"
        "   H = 2.5 মিনিট = 2.5 / 60 ঘণ্টা = 0.04167 ঘণ্টা\n\n"
        "ধাপ ২: আরল্যাং সূত্রে মান বসানো:\n"
        "   A = lambda * H = 1800 * (2.5 / 60) = 1800 * (1 / 24) = 75 Erlangs।\n\n"
        "ধাপ ৩: ফলাফলের ব্যাখ্যা:\n"
        "   এর অর্থ হলো পিক আওয়ারে নিরবচ্ছিন্নভাবে ৭৫টি ভয়েস চ্যানেল ১০০% সময় ব্যস্ত থাকার সমতুল্য ট্রাফিক জেনারেট হচ্ছে।",
        "• কল ডিউরেশন H কে অবশ্যই ঘণ্টায় (hour) কনভার্ট করে সূত্রে বসাতে হবে। মিনিট বা সেকেন্ডে রেখে গুণ করলে পুরো উত্তর ভুল হয়ে যাবে।"
    )

    # Section 8.4: Types 5-6
    add_h2("৮.৪ অপটিক্যাল ফাইবার (NA) ও পাওয়ার বাজেট সমাধান (Types 5-6)", "sec_8_4")

    add_h3("টাইপ ০৫: অপটিক্যাল ফাইবার সংকট কোণ, NA ও এক্সেপটেন্স কোণ")
    add_body("একটি অপটিক্যাল ফাইবারের কোরের প্রতিসরাঙ্ক 1.48 এবং ক্ল্যাডিংয়ের প্রতিসরাঙ্ক 1.46 হলে সংকট কোণ, নিউমেরিক্যাল অ্যাপারচার ও এক্সেপটেন্স কোণ নির্ণয় করুন।")

    add_math_solution(
        "ফাইবারের নিউমেরিক্যাল অ্যাপারচার ও এক্সেপটেন্স কোণ নির্ণয়",
        "• n_1 = 1.48, n_2 = 1.46",
        "১. theta_c = arcsin(n_2 / n_1)\n২. NA = sqrt(n_1^2 - n_2^2)\n৩. theta_a = arcsin(NA)",
        "ধাপ ১: সংকট কোণ:\n"
        "   theta_c = arcsin(1.46 / 1.48) = arcsin(0.9865) = 80.57°।\n\n"
        "ধাপ ২: নিউমেরিক্যাল অ্যাপারচার (NA):\n"
        "   NA = sqrt(1.48^2 - 1.46^2) = sqrt(0.0588) = 0.2425।\n\n"
        "ধাপ ৩: সর্বোচ্চ এক্সেপটেন্স কোণ:\n"
        "   theta_a = arcsin(0.2425) = 14.03°।",
        "• সংকট কোণ সর্বদা ৯০ ডিগ্রির কাছাকাছি বড় কোণ হয় (পূর্ণ অভ্যন্তরীণ প্রতিফলনের জন্য), আর এক্সেপটেন্স কোণ তুলনামূলক ছোট কোণ হয়।"
    )

    add_h3("টাইপ ০৬: সর্বোচ্চ অপটিক্যাল ট্রান্সমিশন দূরত্ব নির্ণয়")
    add_body("একটি অপটিক্যাল ট্রান্সমিটারের পাওয়ার -3 dBm এবং রিসিভার সেন্সিটিভিটি -33 dBm। ফাইবার ক্যাবলের লস 0.3 dB/km, কানেক্টর লস মোট 1.5 dB এবং সেফটি মার্জিন 4.5 dB। কোনো অপটিক্যাল অ্যামপ্লিফায়ার ছাড়া ফাইবারটির সর্বোচ্চ কত দূরত্ব পর্যন্ত সিগন্যাল পাঠানো সম্ভব?")
    
    add_math_solution(
        "অপটিক্যাল পাওয়ার বাজেট অনুযায়ী সর্বোচ্চ আন-অ্যাম্প্লিফাইড লিঙ্ক দৈর্ঘ্য",
        "• ট্রান্সমিটার পাওয়ার P_tx = -3 dBm\n• রিসিভার সেন্সিটিভিটি P_rx = -33 dBm\n• ক্যাবল অ্যাটেনুয়েশন alpha = 0.3 dB/km\n• মোট কানেক্টর লস = 1.5 dB\n• সেফটি মার্জিন = 4.5 dB",
        "Total Allowable Loss = P_tx - P_rx = (alpha * L) + Connector Loss + Margin",
        "ধাপ ১: মোট অনুমোদিত সিস্টেম লস (Power Budget):\n"
        "   Power Budget = P_tx - P_rx = -3 dBm - (-33 dBm) = 30 dB\n\n"
        "ধাপ ২: ফিক্সড লস ও সেফটি মার্জিনের যোগফল:\n"
        "   Fixed Losses = Connector Loss + Margin = 1.5 dB + 4.5 dB = 6.0 dB\n\n"
        "ধাপ ৩: ফাইবারের দৈর্ঘ্যের জন্য অবশিষ্ট লস:\n"
        "   Cable Loss Budget = 30 dB - 6.0 dB = 24.0 dB\n\n"
        "ধাপ ৪: সর্বোচ্চ লিঙ্ক দৈর্ঘ্য (L_max):\n"
        "   L_max = Cable Loss Budget / alpha = 24.0 dB / 0.3 dB/km = 80 কিলোমিটার।",
        "• dBm থেকে dBm বিয়োগ করলে ফলাফল সবসময় ডেসিবেল (dB) এককে আসে। ঋণাত্মক চিহ্নের বিয়োগে (-3 - (-33) = +30) সতর্ক থাকতে হবে।"
    )

    # Section 8.5: Types 7-8
    add_h2("৮.৫ মাইক্রোওয়েভ ও নেটওয়ার্ক ট্রান্সমিশন ডিলে সমাধান (Types 7-8)", "sec_8_5")

    add_h3("টাইপ ০৭: মাইক্রোওয়েভ FSPL, ফ্রেনেল জোন ও ফেইড মার্জিন")
    add_body("একটি 7 GHz মাইক্রোওয়েভ লিঙ্ক 20 কিমি দূরবর্তী দুটি টাওয়ারের মধ্যে স্থাপিত। মাঝপথে (10 কিমি দূরত্বে) প্রথম ফ্রেনেল জোনের ব্যাসার্ধ কত এবং পাথ লস কত?")
    
    add_math_solution(
        "৭ গিগাহার্টজ মাইক্রোওয়েভ লিঙ্কের FSPL ও মাঝপথের ফ্রেনেল ব্যাসার্ধ",
        "• ফ্রিকোয়েন্সি f = 7 GHz\n• মোট দূরত্ব D = 20 km\n• বাধার দূরত্ব d_1 = 10 km, d_2 = 10 km",
        "১. FSPL (dB) = 92.45 + 20*log10(f_GHz) + 20*log10(d_km)\n২. r_1 = 17.32 * sqrt((d_1 * d_2) / (f * D))",
        "ধাপ ১: ফ্রি স্পেস পাথ লস (FSPL) গণনা:\n"
        "   FSPL = 92.45 + 20*log10(7) + 20*log10(20)\n"
        "   • log10(7) = 0.8451 => 20 * 0.8451 = 16.90 dB\n"
        "   • log10(20) = 1.3010 => 20 * 1.3010 = 26.02 dB\n"
        "   FSPL = 92.45 + 16.90 + 26.02 = 135.37 dB\n\n"
        "ধাপ ২: প্রথম ফ্রেনেল জোন ব্যাসার্ধ (r_1) গণনা:\n"
        "   r_1 = 17.32 * sqrt((10 * 10) / (7 * 20))\n"
        "   r_1 = 17.32 * sqrt(100 / 140) = 17.32 * sqrt(0.7143)\n"
        "   r_1 = 17.32 * 0.8451 = 14.638 মিটার ≈ 14.64 মিটার।\n\n"
        "ধাপ ৩: প্রয়োজনীয় ক্লিয়ারেন্স (Clearance):\n"
        "   ন্যূনতম ৬০% ক্লিয়ারেন্স = 0.6 * 14.64 m = 8.78 মিটার।",
        "• ফ্রিকোয়েন্সি GHz এ থাকলে ধ্রুবক 92.45 ব্যবহার করবেন, আর MHz এ থাকলে ধ্রুবক 32.44 ব্যবহার করবেন।"
    )

    add_h3("টাইপ ০৮: নেটওয়ার্ক ট্রান্সমিশন বিলম্ব বনাম প্রোপাগেশন বিলম্ব")
    add_body("একটি 1000-বাইট দীর্ঘ আইপি প্যাকেট একটি 10 Mbps অপটিক্যাল ফাইবার লিঙ্কের মাধ্যমে 2000 কিমি দূরবর্তী গন্তব্যে পাঠানো হচ্ছে। ফাইবারে আলোর গতিবেগ 2 * 10^8 m/s হলে ট্রান্সমিশন বিলম্ব, প্রোপাগেশন বিলম্ব এবং মোট বিলম্ব নির্ণয় করুন।")

    add_math_solution(
        "ফাইবার অপটিক লিঙ্কে প্যাকেট ট্রান্সমিশন ও প্রোপাগেশন ডিলে নির্ণয়",
        "• প্যাকেট সাইজ L = 1000 Bytes = 8000 Bits\n• ব্যান্ডউইথ R = 10 Mbps = 10^7 bps\n• দূরত্ব d = 2000 km = 2 * 10^6 m\n• বেগ v = 2 * 10^8 m/s",
        "১. T_trans = L / R\n২. T_prop = d / v\n৩. Total Delay = T_trans + T_prop",
        "ধাপ ১: ট্রান্সমিশন বিলম্ব গণনা:\n"
        "   T_trans = 8000 Bits / 10,000,000 bps = 0.0008 সেকেন্ড = 0.8 ms।\n\n"
        "ধাপ ২: প্রোপাগেশন বিলম্ব গণনা:\n"
        "   T_prop = 2,000,000 m / (2 * 10^8 m/s) = 0.01 সেকেন্ড = 10.0 ms।\n\n"
        "ধাপ ৩: মোট নেটওয়ার্ক বিলম্ব:\n"
        "   Total Delay = 0.8 ms + 10.0 ms = 10.8 মিলিসেকেন্ড (ms)।",
        "• ট্রান্সমিশন বিলম্ব ব্যান্ডউইথের ওপর নির্ভর করে (L/R), আর প্রোপাগেশন বিলম্ব লিঙ্কের দূরত্বের ওপর নির্ভর করে (d/v)।"
    )

    add_back_to_toc()
    doc.add_page_break()

    # 13. CHAPTER 9: BTCL PREVIOUS EXAMS & SOLUTIONS
    # ==========================================
    add_h1("অধ্যায় ০৯: বিটিসিএল (BTCL) বিগত পরীক্ষার প্রশ্ন ও সমাধান", "ch_09")
    add_body("বিটিসিএল (বাংলাদেশ টেলিকমিউনিকেশনস কোম্পানি লিমিটেড) এর সহকারী ব্যবস্থাপক (কারিগরি) ও উপ-সহকারী প্রকৌশলী নিয়োগ পরীক্ষার স্ট্যান্ডার্ড প্রশ্ন ও তার আদর্শ সমাধান।")

    add_h2("৯.১ বিটিসিএল সহকারী ম্যানেজার (টেকনিক্যাল) লিখিত পরীক্ষা ২০১৫ সমাধান", "sec_9_1")
    add_bullet("প্রশ্ন ০১: 'একটি স্ট্যান্ডার্ড E1 লাইনের ফ্রেম স্ট্রাকচার ও ডাটা রেট ব্যাখ্যা করুন।'",
        "সমাধান: ইউরোপ ও বাংলাদেশে ব্যবহৃত CEPT স্ট্যান্ডার্ড অনুযায়ী E1 লাইনের মোট ট্রান্সমিশন বিট রেট 2.048 Mbps। এতে মোট ৩২টি টাইম স্লট (Time Slot: TS0 থেকে TS31) থাকে। প্রতিটি টাইম স্লট ৮-বিট এবং প্রতি সেকেন্ডে ৮০০০ বার স্যাম্পল হয় (8 * 8000 = 64 kbps)।\n"
        "• TS0: ফ্রেম সিঙ্ক্রোনাইজেশন ও অ্যালাইনমেন্টের জন্য সংরক্ষিত।\n"
        "• TS16: সিগন্যালিং (CAS বা CCS) এর জন্য ডেডিকেটেড।\n"
        "• বাকি ৩০টি টাইম স্লট (TS1-TS15 এবং TS17-TS31): গ্রাহকের ভয়েস বা ডাটা বহনে ব্যবহৃত হয়।\n"
        "মোট ডাটা রেট = 32 slots * 64 kbps = 2048 kbps = 2.048 Mbps।"
    )
    add_bullet("প্রশ্ন ০২: 'অপটিক্যাল ফাইবার স্প্লাইসিং কয় প্রকার ও কী কী?'",
        "সমাধান: দুটি অপটিক্যাল ফাইবার জোড়া লাগানোর পদ্ধতিকে স্প্লাইসিং বলে। এটি প্রধানত ২ প্রকার:\n"
        "১. ফিউশন স্প্লাইসিং (Fusion Splicing): ইলেকট্রিক আর্কের উচ্চ তাপে দুটি ফাইবার কোরকে গলিয়ে স্থায়ীভাবে জোড়া লাগানো হয়। লস অত্যন্ত কম (০.০১ থেকে ০.০৫ dB)। লং-হল ব্যাকবোনে ব্যবহৃত হয়।\n"
        "২. মেকানিক্যাল স্প্লাইসিং (Mechanical Splicing): একটি বিশেষ মেকানিক্যাল স্লিভ ও ইনডেক্স-ম্যাচিং জেলের মাধ্যমে দুটি ফাইবারকে সঠিকভাবে মুখোমুখি ধরে রাখা হয়। লস তুলনামূলক বেশি (০.২ থেকে ০.৫ dB)। জরুরি অস্থায়ী মেরামতে ব্যবহৃত হয়।"
    )

    add_h2("৯.২ বিটিসিএল সহকারী ম্যানেজার ও উপ-সহকারী প্রকৌশলী লিখিত পরীক্ষা ২০১৭", "sec_9_2")
    add_bullet("প্রশ্ন ০৩: 'SS7 (Signaling System No. 7) এর মূল কাজ ও উপাদানগুলো কী কী?'",
        "সমাধান: এসএস৭ হলো আন্তর্জাতিক আউট-অব-ব্যান্ড সিগন্যালিং আর্কিটেকচার যা পাবলিক সুইচড টেলিফোন নেটওয়ার্কে কল সেটআপ, বিলিং, কল টিয়ার-ডাউন ও এসএমএস রাউটিংয়ের কাজ করে। এর ৩টি প্রধান উপাদান:\n"
        "১. SSP (Service Switching Point): টেলিফোন এক্সচেঞ্জ যা কল শুরু বা শেষ করে।\n"
        "২. STP (Signal Transfer Point): প্যাকেট রাউটার যা সিগন্যালিং মেসেজ এক নোড থেকে অন্য নোডে পাঠায়।\n"
        "৩. SCP (Service Control Point): কেন্দ্রীয় ডাটাবেজ (যেমন: HLR/VLR, টোল-ফ্রি নম্বর ডাটাবেজ)।"
    )

    add_back_to_toc()
    doc.add_page_break()

    # ==========================================
    # 14. CHAPTER 10: BTRC PREVIOUS EXAMS & SOLUTIONS
    # ==========================================
    add_h1("অধ্যায় ১০: বিটিআরসি (BTRC) বিগত পরীক্ষার প্রশ্ন ও সমাধান", "ch_10")
    add_body("বাংলাদেশ টেলিযোগাযোগ নিয়ন্ত্রণ কমিশন (BTRC) এর সহকারী পরিচালক ও উপ-সহকারী পরিচালক (কারিগরি) নিয়োগ পরীক্ষার সমাধান।")

    add_h2("১০.১ বিটিআরসি সহকারী পরিচালক (কারিগরি) নিয়োগ পরীক্ষা ২০২১ সমাধান", "sec_10_1")
    add_bullet("প্রশ্ন ০১: 'স্পেকট্রাম ম্যানেজমেন্টে ফ্রিকোয়েন্সি রিফার্মিং (Frequency Refarming) কী?'",
        "সমাধান: পূর্বে কোনো একটি নির্দিষ্ট প্রাচীন প্রযুক্তির জন্য বরাদ্দকৃত ফ্রিকোয়েন্সি ব্যান্ডকে প্রত্যাহার করে আধুনিক উচ্চগতির প্রযুক্তিতে নতুন করে ব্যবহারের অনুমোদন দেওয়াকে ফ্রিকোয়েন্সি রিফার্মিং বলে। যেমন: বাংলাদেশে ঐতিহ্যবাহী 2G এর ১৮০০ MHz ও ২১০০ MHz ব্যান্ডের স্পেকট্রাম রিফার্ম করে 4G LTE ও 5G রোলআউটে ব্যবহার করা হয়েছে।"
    )
    add_bullet("প্রশ্ন ০২: 'ল'ফুল ইন্টারসেপশন (Lawful Interception - LI) কী এবং কেন টেলিকমে এটি বাধ্যতামূলক?'",
        "সমাধান: রাষ্ট্রীয় নিরাপত্তা ও অপরাধ দমনে আদালতের নির্দেশনা বা উপযুক্ত কর্তৃপক্ষের অনুমতি সাপেক্ষে কোনো সন্দেহভাজন ব্যক্তির টেলিকম ট্রাফিক (ভয়েস কল, এসএমএস, ডাটা) আইন প্রয়োগকারী সংস্থার মাধ্যমে রিয়েল-টাইমে নজরদারি বা রেকর্ড করার প্রযুক্তিগত মেকানিজমকে ল'ফুল ইন্টারসেপশন বলে। বিটিআরসির লাইসেন্সিং শর্তানুসারে সকল IGW, ICX ও মোবাইল অপারেটরের কোর নেটওয়ার্কে LI গেটওয়ে থাকা বাধ্যতামূলক।"
    )

    add_h2("১০.২ বিটিআরসি আইসিটি ও টেকনিক্যাল রচনা লেখার মডেল", "sec_10_2")
    add_callout("মডেল রচনা রূপরেখা: বাংলাদেশে ৫জি (5G) এর সম্ভাবনা ও চ্যালেঞ্জ",
        "১. ভূমিকা: ৫জি এর সংজ্ঞা, ৩টি স্তম্ভ (eMBB, URLLC, mMTC)।\n"
        "২. সম্ভাব্য ক্ষেত্র: স্মার্ট এগ্রিকালচার, আরএমজি ফ্যাক্টরি অটোমেশন, টেলিমেডিসিন, স্মার্ট গ্রিড ও সমুদ্রবন্দর লজিস্টিকস।\n"
        "৩. কারিগরি চ্যালেঞ্জ: ফাইবারাইজেশন অব টাওয়ারস (বর্তমানে মাত্র ৩৫-৪০% টাওয়ারে ফাইবার ব্যাকহোল আছে, যা ৭০%+ হতে হবে), উচ্চ স্পেকট্রাম মূল্য ও বিশাল বিনিয়োগের রিটার্ন (ROI)।\n"
        "৪. উপসংহার: সমন্বিত লাইসেন্সিং নীতি ও অবকাঠামো শেয়ারিংয়ের মাধ্যমে সফল ৫জি বাস্তবায়ন।"
    )

    add_back_to_toc()
    doc.add_page_break()

    # ==========================================
    # 15. CHAPTER 11: TELETALK, BSCL & BUET SPECIAL
    # ==========================================
    add_h1("অধ্যায় ১১: টেলিটক, বিএসসিএল ও বুয়েট/এমআইএসটি স্পেশাল", "ch_11")
    add_body("টেলিটক বাংলাদেশ লিমিটেড, বাংলাদেশ সাবমেরিন ক্যাবলস পিএলসি (BSCPLC), বাংলাদেশ স্যাটেলাইট কোম্পানি লিমিটেড (BSCL) এবং বুয়েট পরিচালিত অন্যান্য প্রকৌশল পরীক্ষার প্রশ্ন পর্যালোচনা।")

    add_h2("১১.১ টেলিটক বাংলাদেশ লিমিটেড সহকারী ব্যবস্থাপক প্রশ্ন সমাধান", "sec_11_1")
    add_bullet("প্রশ্ন ০১: 'ভয়েস ওভার এলটিই (VoLTE) প্রযুক্তির সুবিধা কী?'",
        "সমাধান: প্রথাগত 4G নেটওয়ার্কে ভয়েস কল এলে ফোনটি সার্কিট সুইচ ফলব্যাক (CSFB) প্রক্রিয়ায় সাময়িকভাবে 2G বা 3G তে নেমে যেত। কিন্তু VoLTE প্রযুক্তিতে আইএমএস (IMS) কোরের মাধ্যমে সরাসরি ৪জি আইপি নেটওয়ার্কের ভেতর দিয়েই ভয়েস প্যাকেট পাঠানো হয়।\n"
        "সুবিধা: ১. ক্রিস্টাল ক্লিয়ার এইচডি ভয়েস (HD Voice), ২. কল কানেকশন টাইম অত্যন্ত দ্রুত (১-২ সেকেন্ড বনাম ৫-৮ সেকেন্ড), এবং ৩. কথা বলার সময়েও ফুল স্পিডে 4G ইন্টারনেট চালু থাকে।"
    )

    add_h2("১১.২ বুয়েট ও এমআইএসটি পরিচালিত টেলিকম পরীক্ষার কমন প্রশ্নসমূহ", "sec_11_2")
    add_bullet("কমন প্রশ্ন: 'ফটোডায়োড হিসেবে PIN এর চেয়ে APD কেন বেশি সংবেদনশীল?'",
        "সমাধান: অ্যাভালাঞ্চ ফটোডায়োডে (APD) উচ্চ রিভার্স বায়াস প্রয়োগ করার ফলে অভ্যন্তরীণ অ্যাভালাঞ্চ মাল্টিপ্লিকেশন (Avalanche Multiplication) প্রক্রিয়ায় প্রতিটি প্রাইমারি ইলেকট্রন আরও শত শত সেকেন্ডারি ইলেকট্রন উৎপন্ন করে (অভ্যন্তরীণ কারেন্ট গেইন M > 100)। ফলে দুর্বল আলোর সংকেতেও APD শক্তিশালী কারেন্ট তৈরি করতে পারে।"
    )

    add_back_to_toc()
    doc.add_page_break()

    # ==========================================
    # 16. CHAPTER 12: TOP 100 TECHNICAL INTERVIEW Q&A
    # ==========================================
    add_h1("অধ্যায় ১২: টেলিকম সেরা ১০০ টেকনিক্যাল ভাইভা প্রশ্নোত্তর", "ch_12")
    add_body("সরকারি ও বেসরকারি টেলিকম ভাইভা বোর্ডে সর্বাধিক জিজ্ঞাসিত বাছাইকৃত সেরা ১০০টি প্রশ্ন ও স্ট্যান্ডার্ড টু-দ্য-পয়েন্ট উত্তর।")

    add_h2("১২.১ কোর টেলিকম ও মোবাইল কমিউনিকেশন ভাইভা প্রশ্ন", "sec_12_1")
    add_bullet("প্রশ্ন ০১: সিম (SIM) কার্ডের পূর্ণরূপ কী এবং এতে কী কী তথ্য সংরক্ষিত থাকে?",
        "উত্তর: Subscriber Identity Module। এতে IMSI (International Mobile Subscriber Identity), অথেনটিকেশন কি (Ki), পিন/পাক কোড এবং এসএমএস ও কন্টাক্ট ডাটা সংরক্ষিত থাকে।"
    )
    add_bullet("প্রশ্ন ০২: আইএমএসআই (IMSI) এবং আইএমইআই (IMEI) এর পার্থক্য কী?",
        "উত্তর: IMSI হলো ১৫-সংখ্যার গ্রাহক পরিচয় নম্বর যা সিম কার্ডে থাকে এবং অপারেটরের HLR এ নিবন্ধিত হয়। আর IMEI হলো ১৫-সংখ্যার আন্তর্জাতিক মোবাইল ইকুইপমেন্ট আইডেন্টিটি যা হ্যান্ডসেটের নিজস্ব হার্ডওয়্যার সিরিয়াল নম্বর (EIR ডাটাবেজে ট্র্যাক করা হয়)।"
    )
    add_bullet("প্রশ্ন ০৩: হার্ড হ্যান্ডঅফ এবং সফট হ্যান্ডঅফের পার্থক্য কী?",
        "উত্তর: হার্ড হ্যান্ডঅফ হলো 'Break-before-make' (নতুন বেস স্টেশনের সাথে যুক্ত হওয়ার পূর্বে আগের সংযোগ বিচ্ছিন্ন হয়, যেমন: GSM)। সফট হ্যান্ডঅফ হলো 'Make-before-break' (নতুন সংযোগ তৈরি হওয়ার পর আগের সংযোগ বিচ্ছিন্ন হয়, যেমন: CDMA/WCDMA)।"
    )

    add_h2("১২.২ ট্রান্সমিশন, অপটিক্যাল ও আইপি নেটওয়ার্কিং ভাইভা প্রশ্ন", "sec_12_2")
    add_bullet("প্রশ্ন ০৪: ডার্ক ফাইবার (Dark Fiber) কাকে বলে?",
        "উত্তর: মাটিতে বা মাটির নিচে পাতা রয়েছে কিন্তু এখনো কোনো অপটিক্যাল ইকুইপমেন্ট যুক্ত করে ডাটা ট্রান্সমিশন চালু করা হয়নি—এমন অব্যবহৃত অপটিক্যাল ফাইবার স্ট্র্যান্ডকে ডার্ক ফাইবার বলে।"
    )
    add_bullet("প্রশ্ন ০৫: টিসিপি (TCP) এবং ইউডিপি (UDP) এর মধ্যে কোনটি টেলিকম ভয়েসে ব্যবহৃত হয়?",
        "উত্তর: রিয়েল-টাইম ভয়েস ও ভিডিও কলে UDP (RTP প্রোটোকলের সাথে) ব্যবহৃত হয়। কারণ ভয়েসে সামান্য প্যাকেট ড্রপ হলেও চলে, কিন্তু কোনো লেটেন্সি বা বাফারিং সহ্য হয় না।"
    )

    add_h2("১২.৩ বাংলাদেশ টেলিকম ও বিটিআরসি রেগুলেশন ভাইভা প্রশ্ন", "sec_12_3")
    add_bullet("প্রশ্ন ০৬: এনটিটিএন (NTTN) লাইসেন্স থাকার সুবিধা কী?",
        "উত্তর: প্রতিটি মোবাইল অপারেটর ও আইএসপিকে আলাদাভাবে রাস্তা খনন করে নিজস্ব ক্যাবল বসানোর বিশৃঙ্খলা থেকে দেশকে রক্ষা করে একটি কেন্দ্রীয় শেয়ার্ড অপটিক্যাল ফাইবার ব্যাকবোন গড়ে তোলাই এর মূল উদ্দেশ্য।"
    )

    add_back_to_toc()
    doc.add_page_break()

    # ==========================================
    # 17. CHAPTER 13: REAL-WORLD FIELD SCENARIOS
    # ==========================================
    add_h1("অধ্যায় ১৩: বাস্তবভিত্তিক ফিল্ড সিনারিও ও ট্রাবলশুটিং", "ch_13")
    add_body("টেলিকম অপারেশন্স, এনওসি (NOC) এবং ফিল্ড মেইনটেন্যান্সে প্রকৌশলীরা প্রতিদিন যেসব বাস্তব সমস্যার মুখোমুখি হন, সেগুলোর সমাধান কৌশল।")

    add_h2("১৩.১ সিনারিও ১: মোবাইল কল ড্রপ কেন হয় এবং কীভাবে ড্রপ রেট কমানো যায়?", "sec_13_1")
    add_bullet("কল ড্রপের কারণসমূহ:",
        "১. কভারেজ হোল (Coverage Hole): উঁচু ভবন বা পাহাড়ি বাধার কারণে সিগন্যাল লেভেল রিসিভার থ্রেশহোল্ডের নিচে নেমে যাওয়া।\n"
        "২. হ্যান্ডওভার ফেইলিউর (Handover Failure): দ্রুতগামী যানবাহনে এক সেল থেকে অন্য সেলে যাওয়ার সময় টার্গেট সেলে কোনো ফ্রি ভয়েস চ্যানেল না থাকা।\n"
        "৩. কো-চ্যানেল ইন্টারফারেন্স (High Interference / Pilot Pollution)।\n"
        "৪. টাওয়ারের আরএফ ফিডার বা জাম্পার ক্যাবলে ভিএসডব্লিউআর (VSWR) ফল্ট।"
    )
    add_bullet("সমাধান কৌশল:",
        "আরএফ ড্রাইভ টেস্ট (Drive Test) চালিয়ে অ্যান্টেনার টিল্ট (Mechanical & Electrical Tilt) সমন্বয় করা, নেইবার সেল রিলেশন (Neighbor List) আপডেট করা এবং ট্রাফিক বেশি হলে অতিরিক্ত ক্যারিয়ার ফ্রিকোয়েন্সি বা স্মল সেল যোগ করা।"
    )

    add_h2("১৩.২ সিনারিও ২: অপটিক্যাল ফাইবারে কাট পড়লে OTDR দিয়ে কীভাবে ফল্ট শনাক্ত করবেন?", "sec_13_2")
    add_bullet("ট্রাবলশুটিং ধাপসমূহ:",
        "১. ওটিডিআর (OTDR) এর সাথে টেস্ট জাম্পার ক্যাবল দিয়ে ফাইবার পোর্ট যুক্ত করা।\n"
        "২. ফাইবারের প্রতিসরাঙ্ক (IOR: যেমন 1.4682) এবং পালস উইডথ সেট করে লেজার পালস ফায়ার করা।\n"
        "৩. ওটিডিআর স্ক্রিনে ব্যাকস্ক্যাটার গ্রাফের হঠাৎ ড্রপ বা বড় রিফ্লেক্টিভ স্পাইক (Fresnel Reflection) পর্যবেক্ষণ করা। ওটিডিআর স্বয়ংক্রিয়ভাবে কাট পয়েন্টের দূরত্ব (যেমন: 23.456 কিমি) প্রদর্শন করবে।\n"
        "৪. ফাইবার রুট ম্যাপ ও জিপিএস কো-অর্ডিনেট দেখে ফিল্ড স্প্লাইসিং টিম পাঠিয়ে ফিউশন স্প্লাইসিংয়ের মাধ্যমে ফাইবার পুনরায় জোড়া লাগানো।"
    )

    add_h2("১৩.৩ সিনারিও ৩: টেলিকম সাইটে পাওয়ার ফেইলিউর ও ব্যাটারি ব্যাকআপ ম্যানেজমেন্ট", "sec_13_3")
    add_body("বিটিএস সাইটে বাণিজ্যিক এসি গ্রিড ফেইল করলে সাথে সাথে ব্যাটারি ব্যাংক স্বয়ংক্রিয়ভাবে লোড হ্যান্ডেল করে। যদি ৪ ঘণ্টার বেশি সময় বিদ্যুৎ না থাকে, তবে এনওসি (NOC) অ্যালার্ম পাওয়ার সাথে সাথে পোর্টেবল ডিজেল জেনারেটর (DG) পাঠিয়ে রেকটিফায়ারে এসি ইনপুট দিয়ে ব্যাটারি রিচার্জ ও সাইট সচল রাখা হয়।")

    add_back_to_toc()
    doc.add_page_break()

    # ==========================================
    # 18. CHAPTER 14: TELECOM GLOSSARY A TO Z
    # ==========================================
    add_h1("অধ্যায় ১৪: টেলিকম অ্যাক্রোনিম ও পরিভাষার বর্ণানুক্রমিক কোষ (Glossary A to Z)", "ch_14")
    add_body("টেলিযোগাযোগ খাতে বহুল ব্যবহৃত আন্তর্জাতিক ও দেশীয় সংক্ষিপ্ত রূপের (Acronyms) পূর্ণরূপ ও সংক্ষিপ্ত ব্যাখ্যা:")

    add_h2("১৪.১ এ থেকে জেড (A to Z) পর্যন্ত বহুল ব্যবহৃত টেলিকম পরিভাষা", "sec_14_1")

    tbl_acr = doc.add_table(rows=1, cols=3)
    tbl_acr.alignment = WD_TABLE_ALIGNMENT.CENTER
    a_hdrs = ["অ্যাক্রোনিম", "পূর্ণরূপ (Full Form)", "সংক্ষিপ্ত ব্যাখ্যা / প্রয়োগ"]
    a_widths = [Inches(1.2), Inches(2.3), Inches(3.0)]
    for i, h in enumerate(a_hdrs):
        tbl_acr.rows[0].cells[i].text = h
        tbl_acr.rows[0].cells[i].width = a_widths[i]
        tcPr = tbl_acr.rows[0].cells[i]._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="0A2540"/>'))
        p = tbl_acr.rows[0].cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            set_font(r, size_pt=10, bold=True, color=RGBColor(255, 255, 255))

    glossary_data = [
        ("ANS", "Access Network Services", "গ্রাহক পর্যায়ের টেলিকম লাইসেন্স (মোবাইল অপারেটর ও আইএসপি)"),
        ("BGP", "Border Gateway Protocol", "ইন্টারনেটের প্রধান বহিরাগত রাউটিং প্রোটোকল (AS টু AS)"),
        ("BSC", "Base Station Controller", "জিএসএম নেটওয়ার্কে একাধিক বিটিএস নিয়ন্ত্রণকারী নোড"),
        ("BTRC", "Bangladesh Telecommunication Regulatory Commission", "বাংলাদেশের টেলিযোগাযোগ খাতের প্রধান নিয়ন্ত্রণকারী সংস্থা"),
        ("BTS", "Base Transceiver Station", "মোবাইল গ্রাহকের সাথে রেডিও সংকেত আদান-প্রদানকারী টাওয়ার নোড"),
        ("CDMA", "Code Division Multiple Access", "ইউনিক অর্থোগোনাল কোড ব্যবহার করে একই ফ্রিকোয়েন্সিতে মাল্টিপল এক্সেস"),
        ("DWDM", "Dense Wavelength Division Multiplexing", "একই অপটিক্যাল ফাইবার দিয়ে ৮০+ ভিন্ন তরঙ্গের আলো পাঠানো"),
        ("EDFA", "Erbium-Doped Fiber Amplifier", "১৫৫০ ন্যানোমিটার উইন্ডোতে অপটিক্যাল সিগন্যাল অ্যামপ্লিফায়ার"),
        ("eNodeB", "Evolved NodeB", "৪জি এলটিই (4G LTE) রেডিও এক্সেস নেটওয়ার্কের একক বেস স্টেশন"),
        ("EPC", "Evolved Packet Core", "৪জি এলটিই এর সম্পূর্ণ আইপি-ভিত্তিক কোর নেটওয়ার্ক"),
        ("FSPL", "Free Space Path Loss", "ফ্রি স্পেসে দূরত্বের সাথে রেডিও তরঙ্গের শক্তি হ্রাস পাওয়ার অনুপাত"),
        ("GGSN", "Gateway GPRS Support Node", "জিপিআরএস কোর নেটওয়ার্ক ও বহিস্থ ইন্টারনেটের মধ্যবর্তী গেটওয়ে"),
        ("HLR", "Home Location Register", "মোবাইল গ্রাহকের সিম ও সাবস্ক্রিপশন তথ্যের স্থায়ী ডাটাবেজ"),
        ("ICX", "Interconnection Exchange", "দেশের অভ্যন্তরীণ অপারেটরদের মধ্যে আন্তঃসংযোগ ভয়েস ট্রানজিট"),
        ("IGW", "International Gateway", "আন্তর্জাতিক ভয়েস ট্রাফিক আদান-প্রদানের গেটওয়ে সুইচ"),
        ("IIG", "International Internet Gateway", "আন্তর্জাতিক ইন্টারনেট ব্যান্ডউইথ সংগ্রহ ও বিতরণের গেটওয়ে"),
        ("IMS", "IP Multimedia Subsystem", "ভয়েস ওভার এলটিই (VoLTE) ও মাল্টিমিডিয়া আইপি আর্কিটেকচার"),
        ("MIMO", "Multiple Input Multiple Output", "একাধিক অ্যান্টেনা ব্যবহার করে ডাটা থ্রুপুট বহুগুণ বৃদ্ধির প্রযুক্তি"),
        ("MME", "Mobility Management Entity", "৪জি এলটিই কোরে ইউজার মোবিলিটি, সিকিউরিটি ও পেজিং নিয়ন্ত্রক"),
        ("MNP", "Mobile Number Portability", "নম্বর অপরিবর্তিত রেখে মোবাইল অপারেটর পরিবর্তনের সরকারি সুবিধা"),
        ("NTTN", "Nationwide Telecommunication Transmission Network", "দেশব্যাপী অপটিক্যাল ফাইবার ব্যাকবোন ট্রান্সমিশন অবকাঠামো"),
        ("OFDM", "Orthogonal Frequency Division Multiplexing", "অর্থোগোনাল সাব-ক্যারিয়ারে উচ্চগতির ডিজিটাল মডুলেশন প্রযুক্তি"),
        ("OSPF", "Open Shortest Path First", "লিংক-স্টেট অভ্যন্তরীণ রাউটিং প্রোটোকল (Dijkstra SPF)"),
        ("OTDR", "Optical Time Domain Reflectometer", "অপটিক্যাল ফাইবার লস ও কাট পয়েন্ট পরিমাপের টেস্ট ইকুইপমেন্ট"),
        ("PCM", "Pulse Code Modulation", "অ্যানালগ ভয়েসকে ডিজিটাল বিট স্ট্রিমে রূপান্তরের স্ট্যান্ডার্ড স্কিম"),
        ("PGW", "PDN Gateway", "৪জি কোরে ব্যবহারকারীকে আইপি প্রদানকারী ও ইন্টারনেট গেটওয়ে"),
        ("QoS", "Quality of Service", "নেটওয়ার্কের গুণগত মান (কল ড্রপ রেট, লেটেন্সি, জিটার, সাকসেস রেট)"),
        ("SDH", "Synchronous Digital Hierarchy", "ডিজিটাল টেলিযোগাযোগের গ্লোবাল সিঙ্ক্রোনাস ট্রান্সমিশন ফ্রেমওয়ার্ক"),
        ("SGW", "Serving Gateway", "৪জি ইপিসি কোরে ডাটা প্যাকেট ফরোয়ার্ডকারী গেটওয়ে নোড"),
        ("SGSN", "Serving GPRS Support Node", "জিপিআরএস নেটওয়ার্কে প্যাকেট ডাটা মোবিলিটি হ্যান্ডলার"),
        ("SS7", "Signaling System No. 7", "টেলিফোনিতে কল সেটআপ ও বিলিংয়ের আন্তর্জাতিক সিগন্যালিং সিস্টেম"),
        ("VLAN", "Virtual Local Area Network", "সুইচে লজিক্যালি আলাদা সাবনেট ও ব্রডকাস্ট ডোমেইন তৈরির প্রযুক্তি"),
        ("VoLTE", "Voice over Long-Term Evolution", "৪জি এলটিই আইপি নেটওয়ার্কের ওপর ক্রিস্টাল ক্লিয়ার এইচডি ভয়েস কল"),
        ("VSWR", "Voltage Standing Wave Ratio", "অ্যান্টেনা ও ফিডার ক্যাবলের ইম্পিডেন্স ম্যাচিং ও প্রতিফলন পরিমাপক")
    ]

    for row_data in glossary_data:
        row_cells = tbl_acr.add_row().cells
        for i, val in enumerate(row_data):
            row_cells[i].text = val
            row_cells[i].width = a_widths[i]
            p = row_cells[i].paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.15
            for r in p.runs:
                is_bold = (i == 0)
                set_font(r, size_pt=9.5, bold=is_bold, color=COLOR_PRIMARY if is_bold else COLOR_TEXT)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)
    add_back_to_toc()

    # Save to both file names so user has full access regardless of lock
    targets = ["Telecom_Job_Preparation_Guide_Book.docx", "Telecom_Job_Preparation_Guide_Book_v2.docx"]
    saved_files = []
    for fn in targets:
        try:
            doc.save(fn)
            saved_files.append(fn)
            print(f"Successfully saved: {fn}")
        except PermissionError:
            print(f"Notice: {fn} is currently locked by Word. Skipping.")

    print(f"Book generation complete. Saved in: {saved_files}")

if __name__ == "__main__":
    create_telecom_book()
