# -*- coding: utf-8 -*-
"""
build_full_100p_book.py: Master execution script assembling all chapters,
generating the complete 100-page publication-grade handbook.
"""
import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from book_builder_base import create_telecom_book_100p
from book_content_part1 import write_part1
from book_content_part2 import write_part2
from book_content_part3 import write_part3
from book_content_part4_5 import write_part4_5

def build_complete_book():
    doc, helpers = create_telecom_book_100p()

    add_h1 = helpers["add_h1"]
    add_h2 = helpers["add_h2"]
    add_h3 = helpers["add_h3"]
    add_body = helpers["add_body"]
    add_bullet = helpers["add_bullet"]
    add_callout = helpers["add_callout"]
    add_toc_link = helpers["add_toc_link"]
    add_back_to_toc = helpers["add_back_to_toc"]
    set_font = helpers["set_font"]
    COLOR_PRIMARY = helpers["COLOR_PRIMARY"]
    COLOR_SECONDARY = helpers["COLOR_SECONDARY"]
    COLOR_TEXT = helpers["COLOR_TEXT"]
    COLOR_MUTED = helpers["COLOR_MUTED"]

    # =========================================================================
    # 1. BEAUTIFUL COVER PAGE
    # =========================================================================
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
        "★ কোর থিওরি: সিগন্যাল, স্যাম্পলিং, অ্যানালগ/ডিজিটাল মডুলেশন, হাফম্যান কোডিং ও শ্যানন ক্যাপাসিটি\n"
        "★ সেলুলার ও মোবাইল: GSM/3G/4G LTE/5G NR, VoLTE আর্কিটেকচার, Erlang ও ক্লাস্টার ক্যাপাসিটি\n"
        "★ অপটিক্যাল ফাইবার, DWDM, SDH (STM-1/64), GPON/XGS-PON ও লিংক পাওয়ার বাজেট ক্যালকুলেশন\n"
        "★ ডাটা নেটওয়ার্কিং: ওএসআই, টিসিপি/আইপি, সাবনেটিং (FLSM/VLSM), OSPF, BGP ও MPLS ব্যাকবোন\n"
        "★ বাংলাদেশ টেলিকম নেটওয়ার্ক: IGW, ICX, IIG, NTTN, সাবমেরিন ক্যাবল (SMW4/5/6) ও BS-1\n"
        "★ মাস্টার ম্যাথমেটিক্স: ৪২টি ক্যাটাগরির বিস্তারিত গাণিতিক সমস্যা ও স্টেপ-বাই-স্টেপ সমাধান\n"
        "★ বিগত ১০ বছরের নিয়োগ পরীক্ষার লিখিত ও গাণিতিক প্রশ্ন সমাধান (BUET/MIST/BPSC)\n"
        "★ সেরা ১৫০+ টেকনিক্যাল ভাইভা প্রশ্নোত্তর ও প্র্যাকটিক্যাল ফিল্ড ট্রাবলশুটিং সিনারিও"
    )
    set_font(r_f_body, size_pt=8.5, color=COLOR_TEXT)

    # Author Card
    tbl_auth = doc.add_table(rows=1, cols=1)
    tbl_auth.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_auth = tbl_auth.cell(0, 0)
    c_auth.width = Inches(6.5)
    tcPr_auth = c_auth._tc.get_or_add_tcPr()
    tcPr_auth.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F8FAFC"/>'))
    tcPr_auth.append(parse_xml(
        f'<w:tcBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="single" w:sz="6" w:space="0" w:color="CBD5E0"/>\n'
        f'  <w:left w:val="single" w:sz="20" w:space="0" w:color="0A2540"/>\n'
        f'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="CBD5E0"/>\n'
        f'  <w:right w:val="single" w:sz="6" w:space="0" w:color="CBD5E0"/>\n'
        f'</w:tcBorders>'
    ))
    p_auth = c_auth.paragraphs[0]
    p_auth.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_auth.paragraph_format.space_before = Pt(4)
    p_auth.paragraph_format.space_after = Pt(4)
    r_a1 = p_auth.add_run("গ্রন্থনা ও প্রণয়ন:\n")
    set_font(r_a1, size_pt=9.5, bold=False, color=COLOR_MUTED)
    r_a2 = p_auth.add_run("মোহাম্মদ আব্দুল মান্নান\n")
    set_font(r_a2, size_pt=14, bold=True, color=COLOR_PRIMARY)
    r_a3 = p_auth.add_run("বিএসসি ইন ইইই (EEE)\nউপ সহকারী পরিচালক, বাংলাদেশ টেলিযোগাযোগ নিয়ন্ত্রণ কমিশন (বিটিআরসি)")
    set_font(r_a3, size_pt=10, bold=False, color=COLOR_SECONDARY)

    meta_p = doc.add_paragraph()
    meta_p.paragraph_format.space_before = Pt(8)
    meta_p.paragraph_format.space_after = Pt(0)
    meta_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_meta = meta_p.add_run("২০২৬–২০২৭ পরিমার্জিত ও বর্ধিত সংস্করণ • ইঞ্জিনিয়ার্স ক্যারিয়ার প্রকাশনা")
    set_font(r_meta, size_pt=10, bold=True, color=COLOR_SECONDARY)

    doc.add_page_break()

    # =========================================================================
    # 2. PROLOGUE & NECESSITY OF THE BOOK
    # =========================================================================
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
    add_bullet("৪. গাণিতিক দক্ষতা ও ৪২টি মাস্টার টাইপ ক্যালকুলেশন:",
        "লিখিত পরীক্ষায় প্রার্থীদের মধ্যে নম্বরের মূল পার্থক্য তৈরি করে নির্ভুল গাণিতিক সমাধান। শ্যানন ক্যাপাসিটি, নাইকুইস্ট রেট, হাফম্যান কোডিং, হ্যামিং এরর কারেকশন, পিসিএম বিট রেট ও SQNR, এরল্যাং ট্রাফিক ও ক্লাস্টার সাইজ, অপটিক্যাল লিংক পাওয়ার বাজেট, ফ্রেনেল জোন ব্যাসার্ধ এবং আইপি সাবনেটিং (FLSM/VLSM)—প্রতিটি বিষয়ের টাইপভিত্তিক অঙ্ক ও শর্টকাট কৌশল স্টেপ-বাই-স্টেপ সমাধান করে দেওয়া হয়েছে।"
    )
    add_bullet("৫. টেকনিক্যাল ভাইভা ও প্র্যাকটিক্যাল ফিল্ড নলেজ:",
        "কেবল থিওরি মুখস্থ করে বড় বড় টেলিকম অপারেটর বা বিটিআরসি-র ভাইভা বোর্ডে উত্তীর্ণ হওয়া সম্ভব নয়। এই বইটিতে ফিল্ড ইঞ্জিনিয়ারিংয়ের বাস্তব অভিজ্ঞতা যেমন: কল ড্রপের কারণ ও ড্রাইভ টেস্ট, OTDR দিয়ে ফাইবার কাট পয়েন্ট শনাক্তকরণ, বিটিএস টাওয়ারের -৪৮V ডিসি পাওয়ার প্ল্যান্ট ও ব্যাটারি ব্যাকআপ ম্যানেজমেন্ট এবং সেরা ১৫০+ টেকনিক্যাল ভাইভা প্রশ্নোত্তর অন্তর্ভুক্ত রয়েছে।"
    )

    add_h2("কাদের জন্য এই বইটি রচিত? (Target Audience)", "sec_target_audience")
    add_bullet("• অ্যাকাডেমিক ব্যাকগ্রাউন্ড:", "EEE, ETE, ECE, CSE, ICE এবং সংশ্লিষ্ট ডিসিপ্লিনের বিএসসি ও ডিপ্লোমা ইঞ্জিনিয়ারিং ডিগ্রিধারী শিক্ষার্থী ও চাকুরিপ্রার্থী।")
    add_bullet("• সরকারি ও স্বায়ত্তশাসিত সংস্থা:", "বাংলাদেশ টেলিযোগাযোগ নিয়ন্ত্রণ কমিশন (BTRC), বাংলাদেশ টেলিকমিউনিকেশনস কোম্পানি লিমিটেড (BTCL), টেলিটক বাংলাদেশ লিমিটেড (Teletalk), বাংলাদেশ সাবমেরিন ক্যাবলস পিএলসি (BSCPLC), বাংলাদেশ স্যাটেলাইট কোম্পানি লিমিটেড (BSCL) এবং বিসিএস (টেলিকম ও তথ্য ও যোগাযোগ প্রযুক্তি ক্যাডার)।")
    add_bullet("• বেসরকারি কর্পোরেট সেক্টর:", "মোবাইল নেটওয়ার্ক অপারেটর (Grameenphone, Robi, Banglalink), নেশনওয়াইড ট্রান্সমিশন নেটওয়ার্ক (Fiber@Home, Summit Communications, Bahon), ইন্টারনেট সার্ভিস প্রোভাইডার (ISP) এবং আন্তর্জাতিক ভেন্ডর কোম্পানি (Huawei, Ericsson, ZTE, Nokia)।")

    add_callout("লেখকের পরিচিতি ও শুভেচ্ছা বার্তা",
        "প্রিয় প্রকৌশলী সহযোদ্ধা,\n"
        "টেলিযোগাযোগ নিয়ন্ত্রণ, লাইসেন্সিং ফ্রেমওয়ার্ক ও রেডিও স্পেকট্রাম ব্যবস্থাপনার বাস্তব অভিজ্ঞতা এবং বুয়েট ও এমআইএসটি পরিচালিত প্রকৌশলী নিয়োগ পরীক্ষার আধুনিক প্রশ্ন প্যাটার্নের আলোকে এই পূর্ণাঙ্গ গাইড বইটি প্রণয়ন করা হয়েছে। দেশের প্রতিটি মেধাবী প্রকৌশলী যেন আত্মবিশ্বাসের সাথে বিটিআরসি, বিটিসিএল, টেলিটকসহ শীর্ষ সংস্থায় সফল ক্যারিয়ার গড়তে পারে—এটিই এই বইটির মূল সার্থকতা। নিয়মিত অনুশীলন ও অধ্যবসায়ের মাধ্যমেই আপনার সাফল্য নিশ্চিত হোক।\n\n"
        "— মোহাম্মদ আব্দুল মান্নান\n"
        "বিএসসি ইন ইইই (EEE)\n"
        "উপ সহকারী পরিচালক, বাংলাদেশ টেলিযোগাযোগ নিয়ন্ত্রণ কমিশন (বিটিআরসি)"
    )

    add_h2("গ্রন্থকার পরিচিতি ও সংক্ষিপ্ত জীবনবৃত্তান্ত (Author's Profile & Bio-Data)", "sec_author_bio")
    add_body("টেলিযোগাযোগ ও তথ্যপ্রযুক্তি খাতে অভিজ্ঞ প্রকৌশলী, প্রযুক্তি গবেষক এবং বাংলাদেশ টেলিযোগাযোগ নিয়ন্ত্রণ কমিশনের (বিটিআরসি) কর্মকর্তা মোহাম্মদ আব্দুল মান্নান-এর সংক্ষিপ্ত জীবনবৃত্তান্ত ও পেশাগত পরিচিতি:")

    tbl_bio = doc.add_table(rows=1, cols=2)
    tbl_bio.alignment = WD_TABLE_ALIGNMENT.CENTER
    b_hdrs = ["বিবরণ / ক্ষেত্র", "তথ্য ও বিবরণ"]
    b_widths = [Inches(2.2), Inches(4.3)]
    for i, h in enumerate(b_hdrs):
        c = tbl_bio.rows[0].cells[i]
        c.text = h
        c.width = b_widths[i]
        tcPr = c._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="0A2540"/>'))
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            set_font(r, size_pt=9.5, bold=True, color=RGBColor(255, 255, 255))

    bio_data = [
        ("পূর্ণ নাম ও উপাধি", "প্রকৌশলী মোহাম্মদ আব্দুল মান্নান (Engr. Mohammad Abdul Mannan)"),
        ("পিতার নাম", "মোহাম্মদ আবুল কাশেম"),
        ("মাতার নাম", "জাহানারা বেগম"),
        ("বর্তমান পদবি ও সংস্থা", "উপ সহকারী পরিচালক (কারিগরি), স্পেকট্রাম বিভাগ\nবাংলাদেশ টেলিযোগাযোগ নিয়ন্ত্রণ কমিশন (বিটিআরসি)\n[কার্যকাল: ২৩ ডিসেম্বর ২০১৮ – বর্তমান (৭+ বছর)]"),
        ("অফিসিয়াল কর্মক্ষেত্র ও দায়িত্ব", "বিটিআরসির স্পেকট্রাম বিভাগে জাতীয় রেডিও ফ্রিকোয়েন্সি পর্যবেক্ষণ, মোবাইল নেটওয়ার্ক কোয়ালিটি অব সার্ভিস (QoS / QoE), ড্রাইভ টেস্ট, কল ড্রপ রেট (CDR), স্পেকট্রাম অডিট এবং ৫জি/৪জি সেলুলার কমপ্লায়েন্স নিশ্চিতকরণ।"),
        ("শিক্ষাগত যোগ্যতা (স্নাতক ও ডিপ্লোমা)", "• বিএসসি ইন ইলেকট্রিক্যাল অ্যান্ড ইলেকট্রনিক্স ইঞ্জিনিয়ারিং (B.Sc. in EEE), উত্তরা ইউনিভার্সিটি (CGPA: ৩.৫৫ / ৪.০০, ২০২২)\n• ডিপ্লোমা ইন টেলিকমিউনিকেশন টেকনোলজি (৪ বছর মেয়াদি), এসএসআরআইটিএম, বিটিইবি (CGPA: ৩.৫৫ / ৪.০০, ২০১৫)\n• এসএসসি (বিজ্ঞান বিভাগ), কামারপাড়া হাই স্কুল অ্যান্ড কলেজ, ঢাকা শিক্ষা বোর্ড (GPA: ৫.০০ গোল্ডেন এ+, ২০১১)\n• অষ্টম শ্রেণি সরকারি জুনিয়র মেধা বৃত্তি, ঢাকা শিক্ষা বোর্ড (২০০৮)"),
        ("উচ্চতর প্রশিক্ষণ ও প্রফেশনাল সার্টিফিকেশন", "• AI-Assisted Data Analysis with Python & Cursor — আইআইটি (IIT), ঢাকা বিশ্ববিদ্যালয়\n• AI Fundamentals — গুগল (Google / Coursera Verified ID: VFJZ6QRWPVI2)\n• Advanced Networking with Linux (CCNA Concept) — কোডটেক আইটি (ডুয়েট, গাজীপুর)\n• Professional Web Design & Development (Level 3) — ইউওয়াই ল্যাব (UY LAB)\n• Professional SEO & Digital Marketing (Level 3) — ইউওয়াই ল্যাব (UY LAB)"),
        ("বিশেষায়িত ইঞ্জিনিয়ারিং প্রকল্পসমূহ", "১. Cellular QoS & Spectrum Telemetry Assessment Framework (দেশব্যাপী ড্রাইভ টেস্ট মূল্যায়ন)\n২. AI-Assisted Data Analysis & Telemetry Pipeline (IIT DU Capstone Project)\n৩. Microwave Link Budget & Fresnel Clearance Calculation Engine\n৪. Multi-Vendor Router Config Auto-Backup Tool (Python Netmiko & Linux)\n৫. Telecom Asset & Frequency Licensing Workflow Automation System"),
        ("শিক্ষকতা ও মেন্টরিং অভিজ্ঞতা", "• প্রাক্তন প্রশিক্ষক (ইলেকট্রিক্যাল ও ফিজিক্স), শহীদ এস.এ. মেমোরিয়াল পলিটেকনিক (২০১৮)\n• প্রতিষ্ঠাতা ও মেন্টর, উদয়ন একাডেমিক কেয়ার ও সায়েন্স একাডেমিক জোন (৮+ বছরের শিক্ষকতা অভিজ্ঞতা)"),
        ("অফিস ও যোগাযোগের ঠিকানা", "অফিস: বিটিআরসি ভবন, প্লট: ই-৫/এ, আগারগাঁও প্রশাসনিক এলাকা, শেরেবাংলা নগর, ঢাকা-১২০৭\nবাসা: বাড়ি নং ৫৯, গ্রাম: কামারপাড়া, ডাকঘর: নিশাত নগর (১৭১১), থানা: তুরাগ, ঢাকা-১২৩০"),
        ("ই-মেইল ও মোবাইল নম্বর", "অফিস ই-মেইল: mannan@btrc.gov.bd • ব্যক্তিগত ই-মেইল: abmannan.eee@gmail.com\nমোবাইল: ০১৯২৮-৫৭১৭৬১, ০১৫৫০-০৮০৮১৩ • অফিস ফোন: +৮৮০ ২ ৫৫৬৬৭৭৬৬ (Ext: 362)"),
        ("অনলাইন পোর্টফোলিও ও প্রোফাইল", "পোর্টফোলিও: https://abmannan1761.github.io/my-portfolio/\nগিটহাব: github.com/AbMannan1761 • লিঙ্কডইন: linkedin.com/in/mohammad-abdul-mannan-38876114a")
    ]

    for row in bio_data:
        r_cells = tbl_bio.add_row().cells
        for i, val in enumerate(row):
            r_cells[i].text = val
            r_cells[i].width = b_widths[i]
            p = r_cells[i].paragraphs[0]
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.15
            for r in p.runs:
                set_font(r, size_pt=9.0, bold=(i == 0), color=COLOR_TEXT)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    add_back_to_toc()
    doc.add_page_break()

    # =========================================================================
    # 3. FOREWORD & EXAM STRATEGY
    # =========================================================================
    add_h1("ভূমিকা ও নিয়োগ পরীক্ষার প্রস্তুতি কৌশল", "sec_foreword")
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

    # =========================================================================
    # 4. INTERACTIVE TABLE OF CONTENTS
    # =========================================================================
    add_h1("বইয়ের পূর্ণাঙ্গ সূচীপত্র ও ইন্টারেক্টিভ নেভিগেশন", "bookmark_toc")
    
    nav_guide = (
        "১. কীবোর্ডে Ctrl + F চাপুন অথবা Word মেনু থেকে View ➔ Navigation Pane অপশনে টিক দিন। বাম পাশের সাইডবারে অধ্যায় ও টপিকের সম্পূর্ণ ট্রি দেখা যাবে এবং যেকোনো টপিকে এক ক্লিকেই জাম্প করা যাবে।\n"
        "২. নিচের সূচীপত্রের যেকোনো অধ্যায় বা টপিকের উপর Ctrl চেপে রেখে মাউস ক্লিক (Ctrl + Click) করলেই সরাসরি সেই অধ্যায় বা টপিকে চলে যাওয়া যাবে।\n"
        "৩. প্রতিটি অধ্যায়ের শেষে [⬆️ সূচীপত্রে ফিরে যান] লিংকে ক্লিক করে তাৎক্ষণিক আবার সূচীপত্রে ফিরে আসতে পারবেন।"
    )
    add_callout("সহজ নেভিগেশন নির্দেশিকা (Quick Navigation Guide)", nav_guide)

    toc_structure = [
        ("প্রারম্ভিক অংশ ও প্রস্তুতি কৌশল", None, [
            ("প্রারম্ভিক কথা ও বইটির প্রয়োজনীয়তা (Prologue & Necessity)", "sec_prologue"),
            ("গ্রন্থকার পরিচিতি ও সংক্ষিপ্ত জীবনবৃত্তান্ত (Author's Bio-Data)", "sec_author_bio"),
            ("ভূমিকা ও নিয়োগ পরীক্ষার প্রস্তুতি কৌশল (BUET/MIST Pattern)", "sec_foreword")
        ]),
        ("প্রথম খণ্ড: কোর টেলিকমিউনিকেশন ইঞ্জিনিয়ারিং", None, [
            ("অধ্যায় ০১: ফান্ডামেন্টালস অব কমিউনিকেশন সিস্টেম", "ch_01", [
                ("১.১ কমিউনিকেশন সিস্টেমের বেসিক মডেল ও নয়েজ অ্যানালাইসিস", "sec_1_1"),
                ("১.২ তথ্য তত্ত্ব ও সোর্স কোডিং (Entropy & Huffman Coding)", "sec_1_2"),
                ("১.৩ চ্যানেল কোডিং ও এরর কন্ট্রোল (Hamming & CRC Codes)", "sec_1_3"),
                ("১.৪ অ্যানালগ মডুলেশন থিওরি (AM, DSB, SSB, VSB ও FM)", "sec_1_4"),
                ("১.৫ স্যাম্পলিং থিওরেম, পিসিএম (PCM) ও ডিজিটাল ট্রান্সমিশন", "sec_1_5"),
                ("১.৬ ডিজিটাল ক্যারিয়ার মডুলেশন ও কনস্টেলেশন ডায়াগ্রাম", "sec_1_6"),
                ("১.৭ কনভোল্যুশনাল কোডিং ও ভিটারবি ডিকোডারের কার্যপদ্ধতি", "sec_1_7"),
                ("১.৮ পালস শেপিং, রেইজড কোসাইন ফিল্টার ও আই ডায়াগ্রাম", "sec_1_8"),
                ("১.৯ ডেল্টা মডুলেশন (DM) ও অ্যাডাপ্টিভ ডেল্টা মডুলেশন (ADM)", "sec_1_9")
            ]),
            ("অধ্যায় ০২: সেলুলার ও মোবাইল ওয়্যারলেস কমিউনিকেশন", "ch_02", [
                ("২.১ সেলুলার কনসেপ্ট, ক্লাস্টার সাইজ ও কো-চ্যানেল ইন্টারফারেন্স", "sec_2_1"),
                ("২.২ সেলুলার নেটওয়ার্ক ক্যাপাসিটি বৃদ্ধির কৌশল", "sec_2_2"),
                ("২.৩ ট্রাফিক ইঞ্জিনিয়ারিং ও আরল্যাং থিওরি (Erlang-B & Erlang-C)", "sec_2_3"),
                ("২.৪ ওয়্যারলেস প্রোপাগেশন, ফেইডিং ও মাল্টিপাথ মেকানিজম", "sec_2_4"),
                ("২.৫ মোবাইল প্রজন্মের রূপান্তর (2G GSM থেকে 5G NR)", "sec_2_5"),
                ("২.৬ ৪জি এলটিই আর্কিটেকচার ও ৫জি নিউ রেডিও (5G NR)", "sec_2_6"),
                ("২.৭ জিএসএম লজিক্যাল ও ফিজিক্যাল চ্যানেল হায়ারার্কি", "sec_2_7"),
                ("২.৮ ৪জি এলটিই ফিজিক্যাল লেয়ার রিসোর্স গ্রিড ক্যালকুলেশন", "sec_2_8"),
                ("২.৯ ৫জি এনআর নিউমারোলজি, এসসিএস ও ম্যাসিভ মাইমো বিমফর্মিং", "sec_2_9")
            ]),
            ("অধ্যায় ০৩: অপটিক্যাল ফাইবার কমিউনিকেশন ও ট্রান্সমিশন নেটওয়ার্ক", "ch_03", [
                ("৩.১ অপটিক্যাল ফাইবারের মূলনীতি ও আলোক বিজ্ঞান", "sec_3_1"),
                ("৩.২ অপটিক্যাল সিগন্যাল লস ও ডিসপারশন (Attenuation & Dispersion)", "sec_3_2"),
                ("৩.৩ আধুনিক অপটিক্যাল নেটওয়ার্কিং (WDM, DWDM, CWDM, EDFA, PON)", "sec_3_3"),
                ("৩.৪ সিনক্রোনাস ডিজিটাল হায়ারার্কি (SDH) বনাম PDH", "sec_3_4"),
                ("৩.৫ অপটিক্যাল রাইজ-টাইম বাজেট ও সিস্টেম ব্যান্ডউইথ ভেরিফিকেশন", "sec_3_5"),
                ("৩.৬ নন-লিনিয়ার অপটিক্যাল ইফেক্টস (SPM, XPM, FWM, SRS, SBS)", "sec_3_6"),
                ("৩.৭ অপটিক্যাল ট্রান্সপোর্ট নেটওয়ার্ক (OTN - ITU-T G.709) আর্কিটেকচার", "sec_3_7")
            ]),
            ("অধ্যায় ০৪: ডাটা কমিউনিকেশন ও কম্পিউটার নেটওয়ার্কিং", "ch_04", [
                ("৪.১ ওএসআই (OSI) ৭-লেয়ার মডেল বনাম টিসিপি/আইপি (TCP/IP) স্ট্যাক", "sec_4_1"),
                ("৪.২ আইপি অ্যাড্রেসিং ও ক্লাসলেস নেটওয়ার্ক (IPv4 vs IPv6)", "sec_4_2"),
                ("৪.৩ আইপি সাবনেটিং ও ভিএলএসএম (FLSM vs VLSM) - শর্টকাট কৌশল", "sec_4_3"),
                ("৪.৪ রাউটিং প্রোটোকলস: OSPF, BGP ও টিসিপি ফ্লো কন্ট্রোল", "sec_4_4"),
                ("৪.৫ এমপিএলএস (MPLS) ও নেটওয়ার্ক সিকিউরিটি", "sec_4_5"),
                ("৪.৬ ওএসপিএফ এলএসএ টাইপ ১ থেকে ৭ এবং এরিয়া ক্লাসিফিকেশন", "sec_4_6"),
                ("৪.৭ বিজিপি বেস্ট পাথ সিলেকশন এলগরিদম (BGP 10-Step Algorithm)", "sec_4_7"),
                ("৪.৮ টিসিপি কনজেশন কন্ট্রোল মেকানিজম (Tahoe, Reno & Mathis Formula)", "sec_4_8"),
                ("৪.৯ আইপিভি৬ অটো-কনফিগারেশন (SLAAC) ও ডুয়াল স্ট্যাক মাইগ্রেশন", "sec_4_9")
            ]),
            ("অধ্যায় ০৫: মাইক্রোওয়েভ, স্যাটেলাইট ও অ্যান্টেনা কমিউনিকেশন", "ch_05", [
                ("৫.১ মাইক্রোওয়েভ রেডিও প্রোপাগেশন ও ফ্রেনেল জোন (Fresnel Zone)", "sec_5_1"),
                ("৫.২ ফ্রি স্পেস পাথ লস (FSPL) ও লিঙ্ক বাজেট ক্যালকুলেশন", "sec_5_2"),
                ("৫.৩ অ্যান্টেনা থিওরি, গেইন, ডিরেক্টিভিটি ও VSWR", "sec_5_3"),
                ("৫.৪ স্যাটেলাইট কমিউনিকেশন ও বঙ্গবন্ধু স্যাটেলাইট-১ (BS-1)", "sec_5_4"),
                ("৫.৫ টু-রে গ্রাউন্ড রিফ্লেকশন মডেল ও ব্রেকপয়েন্ট দূরত্ব", "sec_5_5"),
                ("৫.৬ নাইফ-এজ ডিফ্র্যাকশন লস ও ফ্রেনেল প্যারামিটার", "sec_5_6"),
                ("৫.৭ আর্থ বাল্জ ও কার্ভেচার ক্লিয়ারেন্স ক্যালকুলেশন", "sec_5_7"),
                ("৫.৮ প্যারাবলিক ডিশ অ্যান্টেনা জিওমেট্রি, এফিশিয়েন্সি ও বিমউইথ", "sec_5_8"),
                ("৫.৯ স্যাটেলাইট আর্থ স্টেশন লুক অ্যাঙ্গেল ক্যালকুলেশন (আজিমুথ ও এলিভেশন)", "sec_5_9")
            ])
        ]),
        ("দ্বিতীয় খণ্ড: বাংলাদেশ টেলিকম ইকোসিস্টেম ও রেগুলেটরি ফ্রেমওয়ার্ক", None, [
            ("অধ্যায় ০৬: বাংলাদেশ টেলিকম নেটওয়ার্ক ও লাইসেন্সিং ফ্রেমওয়ার্ক", "ch_06", [
                ("৬.১ বাংলাদেশ টেলিকম নেটওয়ার্ক আর্কিটেকচার (BTRC Framework)", "sec_6_1"),
                ("৬.২ প্রধান লাইসেন্সিং লেয়ারসমূহ (IGW, ICX, IIG, NTTN, ANS, ITC, NIX)", "sec_6_2"),
                ("৬.৩ সাবমেরিন ক্যাবল নেটওয়ার্ক ও আন্তর্জাতিক সংযোগ (SMW4/5/6)", "sec_6_3"),
                ("৬.৪ বিটিআরসি স্পেকট্রাম বরাদ্দ, এমএনপি (MNP) ও QoS নীতিমালা", "sec_6_4"),
                ("৬.৫ বাংলাদেশ টেলিকম সিগন্যালিং কল ফ্লো ও মেসেজ ল্যাডার", "sec_6_5"),
                ("৬.৬ বিটিআরসি লাইসেন্সিং ফ্রেমওয়ার্ক ও রেগুলেটরি পলিসি চার্ট", "sec_6_6"),
                ("৬.৭ বাংলাদেশে স্পেকট্রাম বরাদ্দ ও মোবাইল অপারেটরদের ফ্রিকোয়েন্সি ম্যাট্রিক্স", "sec_6_7")
            ])
        ]),
        ("তৃতীয় খণ্ড: বেসিক ইলেকট্রিক্যাল, ইলেকট্রনিক্স ও গণিত", None, [
            ("অধ্যায় ০৭: ইলেকট্রিক্যাল ও ইলেকট্রনিক্স ফান্ডামেন্টালস", "ch_07", [
                ("৭.১ সার্কিট অ্যানালাইসিস থিওরেমস (KVL, KCL, Thevenin, Norton)", "sec_7_1"),
                ("৭.২ এসি ফান্ডামেন্টালস, RLC রেজোন্যান্স ও পাওয়ার ফ্যাক্টর", "sec_7_2"),
                ("৭.৩ সেমিকন্ডাক্টর ও অপারেশনাল অ্যামপ্লিফায়ার (Op-Amp)", "sec_7_3"),
                ("৭.৪ টেলিকম -৪৮ ভোল্ট ডিসি (-48V DC) পাওয়ার সিস্টেম ও ব্যাটারি ব্যাকআপ", "sec_7_4"),
                ("৭.৫ সুপারপজিশন ও এসি ম্যাক্সিমাম পাওয়ার ট্রান্সফার থিওরেম", "sec_7_5"),
                ("৭.৬ অপ-অ্যাম্প অ্যাক্টিভ ফিল্টার ডিজাইন (Active Butterworth Filters)", "sec_7_6"),
                ("৭.৭ টেলিকম সাইটের আর্থিং, গ্রাউন্ডিং ও লাইটনিং সুরক্ষা", "sec_7_7")
            ]),
            ("অধ্যায় ০৮: মাস্টার টেলিকম গণিত ও গুরুত্বপূর্ণ ফর্মুলা শীট (৪২টি টাইপ)", "ch_08", [
                ("৮.১ এক নজরে সকল অত্যাবশ্যকীয় টেলিকম সূত্র ও একক (Formula Cheatsheet)", "sec_8_1"),
                ("৮.২ সিগন্যাল, স্যাম্পলিং ও ইনফরমেশন থিওরি সমাধান (Types 1-4)", "sec_8_2"),
                ("৮.৩ সেলুলার, ট্রাফিক ও আরল্যাং সমাধান (Types 5-8)", "sec_8_3"),
                ("৮.৪ অপটিক্যাল ফাইবার ও ট্রান্সমিশন সমাধান (Types 9-12)", "sec_8_4"),
                ("৮.৫ মাইক্রোওয়েভ, অ্যান্টেনা ও নেটওয়ার্কিং সমাধান (Types 13-16)", "sec_8_5"),
                ("৮.৬ ইলেকট্রিক্যাল, এসি ও ডিসি পাওয়ার সমাধান (Types 17-20)", "sec_8_6"),
                ("৮.৭ অ্যাডভান্সড টেলিকম ইঞ্জিনিয়ারিং গাণিতিক সমাধান (Types 21-42)", "sec_8_7")
            ])
        ]),
        ("চতুর্থ খণ্ড: বিগত নিয়োগ পরীক্ষার প্রশ্ন ও সমাধান", None, [
            ("অধ্যায় ০৯: বিটিসিএল (BTCL) বিগত পরীক্ষার প্রশ্ন ও সমাধান", "ch_09", [
                ("৯.১ বিটিসিএল সহকারী ম্যানেজার (টেকনিক্যাল) লিখিত পরীক্ষা ২০১৫ ও ২০১৭", "sec_9_1"),
                ("৯.২ বিটিসিএল সহকারী ম্যানেজার লিখিত পরীক্ষা ২০১৯, ২০২১ ও ২০২৪", "sec_9_2")
            ]),
            ("অধ্যায় ১০: বিটিআরসি (BTRC) বিগত পরীক্ষার প্রশ্ন ও সমাধান", "ch_10", [
                ("১০.১ বিটিআরসি সহকারী পরিচালক লিখিত পরীক্ষা ২০১৮, ২০২১ ও ২০২৩", "sec_10_1"),
                ("১০.২ বিটিআরসি ফোকাস টেকনিক্যাল রচনা ও পলিসি ফ্রেমওয়ার্ক", "sec_10_2")
            ]),
            ("অধ্যায় ১১: টেলিটক, বিএসসিএল ও বুয়েট/এমআইএসটি স্পেশাল", "ch_11", [
                ("১১.১ টেলিটক ও বিএসসিএল সহকারী প্রকৌশলী প্রশ্ন সমাধান", "sec_11_1"),
                ("১১.২ বুয়েট ও এমআইএসটি পরিচালিত টেলিকম পরীক্ষার কমন প্রশ্নসমূহ", "sec_11_2")
            ])
        ]),
        ("পঞ্চম খণ্ড: ভাইভা স্পেশাল ও ইন্টারভিউ গাইডলাইন", None, [
            ("অধ্যায় ১২: টেলিকম সেরা ১২৫টি টেকনিক্যাল ভাইভা প্রশ্নোত্তর", "ch_12", [
                ("১২.১ প্যানেল ক: কোর ওয়্যারলেস ও মোবাইল কমিউনিকেশন ভাইভা প্রশ্ন", "sec_12_1"),
                ("১২.২ প্যানেল খ: অপটিক্যাল ট্রান্সমিশন ও ট্রান্সপোর্ট নেটওয়ার্ক ভাইভা প্রশ্ন", "sec_12_2"),
                ("১২.৩ প্যানেল গ: আইপি নেটওয়ার্কিং, রাউটিং ও ব্যাকবোন ভাইভা প্রশ্ন", "sec_12_3"),
                ("১২.৪ প্যানেল ঘ: টেলিকম পাওয়ার, সাইট ও ইনফ্রাস্ট্রাকচার ভাইভা প্রশ্ন", "sec_12_4"),
                ("১২.৫ প্যানেল ঙ: বাংলাদেশ টেলিকম ইকোসিস্টেম ও রেগুলেশন ভাইভা প্রশ্ন", "sec_12_5")
            ]),
            ("অধ্যায় ১৩: বাস্তবভিত্তিক ফিল্ড সিনারিও ও ট্রাবলশুটিং", "ch_13", [
                ("১৩.১ সিনারিও ১: মোবাইল কল ড্রপ তদন্ত ও আরএফ অপটিমাইজেশন", "sec_13_1"),
                ("১৩.২ সিনারিও ২: অপটিক্যাল ফাইবারে কাট পড়লে OTDR দিয়ে ফল্ট শনাক্তকরণ", "sec_13_2"),
                ("১৩.৩ সিনারিও ৩: টেলিকম সাইটে পাওয়ার ফেইলিউর ও ব্যাটারি ব্যাকআপ ম্যানেজমেন্ট", "sec_13_3"),
                ("১৩.৪ সিনারিও ৪: অ্যান্টেনা ফিডার ক্যাবলে হাই VSWR অ্যালার্ম সমাধান", "sec_13_4"),
                ("১৩.৫ সিনারিও ৫: বিজিপি রুট ফ্ল্যাপিং (BGP Route Flapping) সমাধান", "sec_13_5"),
                ("১৩.৬ সিনারিও ৬: ডিডব্লিউডিএম লিঙ্কে ওএসএনআর (OSNR) ড্রপ সমাধান", "sec_13_6"),
                ("১৩.৭ সিনারিও ৭: মনসুন মৌসুমে মাইক্রোওয়েভ ফেইড অ্যালার্ম হ্যান্ডলিং", "sec_13_7"),
                ("১৩.৮ সিনারিও ৮: VoLTE কলে ওয়ান-ওয়ে অডিও (One-Way Audio) সমস্যা", "sec_13_8"),
                ("১৩.৯ সিনারিও ৯: আর্থিং রেজিস্ট্যান্স বৃদ্ধি ও বজ্রপাতে ইকুইপমেন্ট সুরক্ষা", "sec_13_9"),
                ("১৩.১০ সিনারিও ১০: স্পেকট্রাম অ্যানালাইজার দিয়ে অবৈধ আরএফ ইন্টারফারেন্স ট্র্যাকিং", "sec_13_10")
            ]),
            ("অধ্যায় ১৪: টেলিকম অ্যাক্রোনিম ও পরিভাষার বর্ণানুক্রমিক কোষ (Glossary A to Z)", "ch_14", [
                ("১৪.১ এ থেকে জেড (A to Z) পর্যন্ত বহুল ব্যবহৃত টেলিকম পরিভাষা", "sec_14_1")
            ])
        ])
    ]

    # =========================================================================
    # 4.1 STRUCTURED TABLE OF CONTENTS (TABLE FORMAT WITH PAGE NUMBERS)
    # =========================================================================
    PAGE_LOOKUP = {
        "sec_prologue": 3, "sec_author_bio": 4, "sec_foreword": 6,
        "ch_01": 11, "sec_1_1": 11, "sec_1_2": 11, "sec_1_3": 13, "sec_1_4": 14,
        "sec_1_5": 14, "sec_1_6": 15, "sec_1_7": 16, "sec_1_8": 16, "sec_1_9": 17,
        "ch_02": 19, "sec_2_1": 19, "sec_2_2": 21, "sec_2_3": 21, "sec_2_4": 21,
        "sec_2_5": 22, "sec_2_6": 22, "sec_2_7": 23, "sec_2_8": 24, "sec_2_9": 25,
        "ch_03": 27, "sec_3_1": 27, "sec_3_2": 28, "sec_3_3": 29, "sec_3_4": 29,
        "sec_3_5": 30, "sec_3_6": 30, "sec_3_7": 31,
        "ch_04": 32, "sec_4_1": 32, "sec_4_2": 33, "sec_4_3": 33, "sec_4_4": 33,
        "sec_4_5": 34, "sec_4_6": 34, "sec_4_7": 35, "sec_4_8": 35, "sec_4_9": 36,
        "ch_05": 37, "sec_5_1": 37, "sec_5_2": 37, "sec_5_3": 38, "sec_5_4": 38,
        "sec_5_5": 39, "sec_5_6": 39, "sec_5_7": 40, "sec_5_8": 40, "sec_5_9": 41,
        "ch_06": 42, "sec_6_1": 42, "sec_6_2": 42, "sec_6_3": 43, "sec_6_4": 43,
        "sec_6_5": 44, "sec_6_6": 45, "sec_6_7": 45,
        "ch_07": 47, "sec_7_1": 47, "sec_7_2": 47, "sec_7_3": 48, "sec_7_4": 48,
        "sec_7_5": 50, "sec_7_6": 50, "sec_7_7": 51,
        "ch_08": 52, "sec_8_1": 52, "sec_8_2": 53, "sec_8_3": 56, "sec_8_4": 58,
        "sec_8_5": 61, "sec_8_6": 64, "sec_8_7": 66,
        "ch_09": 84, "sec_9_1": 84, "sec_9_2": 85,
        "ch_10": 87, "sec_10_1": 87, "sec_10_2": 88,
        "ch_11": 90, "sec_11_1": 90, "sec_11_2": 90,
        "ch_12": 94, "sec_12_1": 94, "sec_12_2": 95, "sec_12_3": 95, "sec_12_4": 96, "sec_12_5": 96,
        "ch_13": 98, "sec_13_1": 98, "sec_13_2": 98, "sec_13_3": 98, "sec_13_4": 98, "sec_13_5": 99,
        "sec_13_6": 99, "sec_13_7": 99, "sec_13_8": 99, "sec_13_9": 99, "sec_13_10": 99,
        "ch_14": 100, "sec_14_1": 100
    }

    domain_map = {
        "sec_prologue": "প্রারম্ভিক", "sec_author_bio": "লেখক পরিচিতি", "sec_foreword": "প্রস্তুতি কৌশল",
        "ch_01": "কোর থিওরি", "ch_02": "সেলুলার ও ৫জি", "ch_03": "অপটিক্যাল ফাইবার",
        "ch_04": "ডাটা নেটওয়ার্ক", "ch_05": "মাইক্রোওয়েভ ও আরএফ", "ch_06": "বিটিআরসি ফ্রেমওয়ার্ক",
        "ch_07": "ইলেকট্রিক্যাল বেসিক", "ch_08": "মাস্টার ম্যাথ (৪২)", "ch_09": "বিটিসিএল সল্যুশন",
        "ch_10": "বিটিআরসি সল্যুশন", "ch_11": "টেলিটক ও বুয়েট", "ch_12": "টেকনিক্যাল ভাইভা",
        "ch_13": "ফিল্ড ট্রাবলশুটিং", "ch_14": "টেলিকম গ্লসারি"
    }

    tbl_toc = doc.add_table(rows=1, cols=3)
    tbl_toc.alignment = WD_TABLE_ALIGNMENT.CENTER
    toc_hdrs = ["অধ্যায় ও বিষয়বস্তুর শিরোনাম (ক্লিক করে সরাসরি যান)", "শাখা / বিভাগ", "পৃষ্ঠা"]
    toc_widths = [Inches(4.6), Inches(1.3), Inches(0.8)]
    for i, h in enumerate(toc_hdrs):
        c = tbl_toc.rows[0].cells[i]
        c.text = h
        c.width = toc_widths[i]
        tcPr = c._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="0A2540"/>'))
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i > 0 else WD_ALIGN_PARAGRAPH.LEFT
        for r in p.runs:
            set_font(r, size_pt=9.5, bold=True, color=RGBColor(255, 255, 255))

    for part_title, part_bm, chapters in toc_structure:
        # Part Row (Spans all columns)
        r_p = tbl_toc.add_row()
        c0 = r_p.cells[0]
        c1 = r_p.cells[1]
        c2 = r_p.cells[2]
        c0.merge(c1)
        c0.merge(c2)
        c0.width = Inches(6.7)
        tcPr_p = c0._tc.get_or_add_tcPr()
        tcPr_p.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="1E3A8A"/>'))
        p_pt = c0.paragraphs[0]
        p_pt.paragraph_format.space_before = Pt(4)
        p_pt.paragraph_format.space_after = Pt(4)
        r_part_txt = p_pt.add_run(f"📁 {part_title}")
        set_font(r_part_txt, size_pt=10, bold=True, color=RGBColor(255, 255, 255))

        for ch_data in chapters:
            if len(ch_data) == 3:
                ch_title, ch_bm, topics = ch_data
                cur_domain = domain_map.get(ch_bm, "কোর টেলিকম")

                # Chapter Row
                row_ch = tbl_toc.add_row()
                c_t, c_d, c_pg = row_ch.cells[0], row_ch.cells[1], row_ch.cells[2]
                c_t.width, c_d.width, c_pg.width = toc_widths[0], toc_widths[1], toc_widths[2]
                for c in (c_t, c_d, c_pg):
                    tcPr = c._tc.get_or_add_tcPr()
                    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F1F5F9"/>'))
                    tcPr.append(parse_xml(
                        f'<w:tcBorders {nsdecls("w")}>\n'
                        f'  <w:top w:val="single" w:sz="6" w:space="0" w:color="CBD5E0"/>\n'
                        f'  <w:left w:val="none"/>\n'
                        f'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="CBD5E0"/>\n'
                        f'  <w:right w:val="none"/>\n'
                        f'</w:tcBorders>'
                    ))

                p_ct = c_t.paragraphs[0]
                p_ct.paragraph_format.space_before = Pt(3)
                p_ct.paragraph_format.space_after = Pt(3)
                safe_ch_title = ch_title.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                hl_ch = parse_xml(
                    f'<w:hyperlink {nsdecls("w")} w:anchor="{ch_bm}" w:history="1">'
                    f'<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:color w:val="0A2540"/><w:b/><w:sz w:val="19"/><w:u w:val="none"/></w:rPr>'
                    f'<w:t>📖 {safe_ch_title}</w:t></w:r></w:hyperlink>'
                )
                p_ct._p.append(hl_ch)

                p_cd = c_d.paragraphs[0]
                p_cd.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_cd.paragraph_format.space_before = Pt(3)
                p_cd.paragraph_format.space_after = Pt(3)
                r_cd = p_cd.add_run(cur_domain)
                set_font(r_cd, size_pt=8.5, bold=True, color=COLOR_PRIMARY)

                p_cpg = c_pg.paragraphs[0]
                p_cpg.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_cpg.paragraph_format.space_before = Pt(3)
                p_cpg.paragraph_format.space_after = Pt(3)
                fb_ch_pg = PAGE_LOOKUP.get(ch_bm, 1)
                fld_ch = parse_xml(
                    f'<w:hyperlink {nsdecls("w")} w:anchor="{ch_bm}" w:history="1">'
                    f'<w:fldSimple {nsdecls("w")} w:instr="PAGEREF {ch_bm} \\h">'
                    f'<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:b/><w:color w:val="0055A5"/><w:sz w:val="18"/></w:rPr>'
                    f'<w:t>{fb_ch_pg}</w:t></w:r></w:fldSimple></w:hyperlink>'
                )
                p_cpg._p.append(fld_ch)

                # Topics Rows
                for top_title, top_bm in topics:
                    row_top = tbl_toc.add_row()
                    ct_t, ct_d, ct_pg = row_top.cells[0], row_top.cells[1], row_top.cells[2]
                    ct_t.width, ct_d.width, ct_pg.width = toc_widths[0], toc_widths[1], toc_widths[2]
                    for c in (ct_t, ct_d, ct_pg):
                        tcPr = c._tc.get_or_add_tcPr()
                        tcPr.append(parse_xml(
                            f'<w:tcBorders {nsdecls("w")}>\n'
                            f'  <w:top w:val="single" w:sz="3" w:space="0" w:color="E2E8F0"/>\n'
                            f'  <w:left w:val="none"/>\n'
                            f'  <w:bottom w:val="single" w:sz="3" w:space="0" w:color="E2E8F0"/>\n'
                            f'  <w:right w:val="none"/>\n'
                            f'</w:tcBorders>'
                        ))

                    p_tt = ct_t.paragraphs[0]
                    p_tt.paragraph_format.space_before = Pt(2)
                    p_tt.paragraph_format.space_after = Pt(2)
                    p_tt.paragraph_format.line_spacing = 1.15
                    safe_top_title = top_title.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                    hl_top = parse_xml(
                        f'<w:hyperlink {nsdecls("w")} w:anchor="{top_bm}" w:history="1">'
                        f'<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:color w:val="1D4ED8"/><w:sz w:val="17"/><w:u w:val="none"/></w:rPr>'
                        f'<w:t>    • {safe_top_title}</w:t></w:r></w:hyperlink>'
                    )
                    p_tt._p.append(hl_top)

                    p_td = ct_d.paragraphs[0]
                    p_td.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_td.paragraph_format.space_before = Pt(2)
                    p_td.paragraph_format.space_after = Pt(2)
                    r_td = p_td.add_run(cur_domain)
                    set_font(r_td, size_pt=8.0, color=COLOR_MUTED)

                    p_tpg = ct_pg.paragraphs[0]
                    p_tpg.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_tpg.paragraph_format.space_before = Pt(2)
                    p_tpg.paragraph_format.space_after = Pt(2)
                    fb_top_pg = PAGE_LOOKUP.get(top_bm, 1)
                    fld_top = parse_xml(
                        f'<w:hyperlink {nsdecls("w")} w:anchor="{top_bm}" w:history="1">'
                        f'<w:fldSimple {nsdecls("w")} w:instr="PAGEREF {top_bm} \\h">'
                        f'<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:color w:val="0055A5"/><w:sz w:val="17"/></w:rPr>'
                        f'<w:t>{fb_top_pg}</w:t></w:r></w:fldSimple></w:hyperlink>'
                    )
                    p_tpg._p.append(fld_top)

            else:
                ch_title, ch_bm = ch_data
                cur_domain = domain_map.get(ch_bm, "প্রারম্ভিক")
                row_ch = tbl_toc.add_row()
                c_t, c_d, c_pg = row_ch.cells[0], row_ch.cells[1], row_ch.cells[2]
                c_t.width, c_d.width, c_pg.width = toc_widths[0], toc_widths[1], toc_widths[2]
                for c in (c_t, c_d, c_pg):
                    tcPr = c._tc.get_or_add_tcPr()
                    tcPr.append(parse_xml(
                        f'<w:tcBorders {nsdecls("w")}>\n'
                        f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="CBD5E0"/>\n'
                        f'  <w:left w:val="none"/>\n'
                        f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="CBD5E0"/>\n'
                        f'  <w:right w:val="none"/>\n'
                        f'</w:tcBorders>'
                    ))

                p_ct = c_t.paragraphs[0]
                p_ct.paragraph_format.space_before = Pt(3)
                p_ct.paragraph_format.space_after = Pt(3)
                safe_ch_title = ch_title.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                hl_ch = parse_xml(
                    f'<w:hyperlink {nsdecls("w")} w:anchor="{ch_bm}" w:history="1">'
                    f'<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:color w:val="0A2540"/><w:b/><w:sz w:val="18"/><w:u w:val="none"/></w:rPr>'
                    f'<w:t>📖 {safe_ch_title}</w:t></w:r></w:hyperlink>'
                )
                p_ct._p.append(hl_ch)

                p_cd = c_d.paragraphs[0]
                p_cd.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_cd.paragraph_format.space_before = Pt(3)
                p_cd.paragraph_format.space_after = Pt(3)
                r_cd = p_cd.add_run(cur_domain)
                set_font(r_cd, size_pt=8.5, bold=True, color=COLOR_PRIMARY)

                p_cpg = c_pg.paragraphs[0]
                p_cpg.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_cpg.paragraph_format.space_before = Pt(3)
                p_cpg.paragraph_format.space_after = Pt(3)
                fb_ch_pg = PAGE_LOOKUP.get(ch_bm, 1)
                fld_ch = parse_xml(
                    f'<w:hyperlink {nsdecls("w")} w:anchor="{ch_bm}" w:history="1">'
                    f'<w:fldSimple {nsdecls("w")} w:instr="PAGEREF {ch_bm} \\h">'
                    f'<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:b/><w:color w:val="0055A5"/><w:sz w:val="18"/></w:rPr>'
                    f'<w:t>{fb_ch_pg}</w:t></w:r></w:fldSimple></w:hyperlink>'
                )
                p_cpg._p.append(fld_ch)

    doc.add_page_break()

    # =========================================================================
    # WRITE EXPANDED PARTS
    # =========================================================================
    print("Writing Part 1 (Chapters 1, 2, 3)...")
    write_part1(doc, helpers)

    print("Writing Part 2 (Chapters 4, 5, 6)...")
    write_part2(doc, helpers)

    print("Writing Part 3 (Chapters 7, 8 - 20 Solved Math Types)...")
    write_part3(doc, helpers)

    print("Writing Parts 4 & 5 (Chapters 9-14, 150+ Viva Q&A, 10 Scenarios, Glossary)...")
    write_part4_5(doc, helpers)

    # Save to both target files
    targets = ["Telecom_Job_Preparation_Guide_Book.docx", "Telecom_Job_Preparation_Guide_Book_v2.docx"]
    saved_files = []
    for fn in targets:
        try:
            doc.save(fn)
            saved_files.append(fn)
            print(f"Successfully saved: {fn}")
        except PermissionError:
            print(f"Notice: {fn} is currently locked by Word. Skipping.")

    print(f"Master compilation complete! Saved files: {saved_files}")

if __name__ == "__main__":
    build_complete_book()
