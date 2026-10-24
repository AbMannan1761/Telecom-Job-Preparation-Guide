# -*- coding: utf-8 -*-
"""
export_to_web_ebook.py:
Exports the 105-page Telecom Job Preparation Guide Book into a modern,
responsive, and interactive web eBook hosted in the docs/ directory for GitHub Pages.
"""
import os
import sys
import html
import shutil

# Ensure utf-8 output
if sys.stdout:
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

import book_content_part1
import book_content_part2
import book_content_part3
import book_content_part4_5

class WebDocCollector:
    """Collects content from the book generator modules into structured HTML blocks."""
    def __init__(self):
        self.sections = []
        self.toc_tree = []
        self.current_chapter = None
        self.current_section = None
        self.active_table = None

    def add_h1(self, text, bookmark=""):
        ch_id = bookmark if bookmark else f"ch_{len(self.toc_tree)+1:02d}"
        item = {
            "type": "h1",
            "title": text,
            "id": ch_id,
            "sections": []
        }
        self.toc_tree.append(item)
        self.current_chapter = item
        self.current_section = None
        self.sections.append({
            "type": "chapter_start",
            "id": ch_id,
            "title": text
        })

    def add_h2(self, text, bookmark=""):
        sec_id = bookmark if bookmark else f"sec_{len(self.sections)+1}"
        item = {
            "type": "h2",
            "title": text,
            "id": sec_id
        }
        if self.current_chapter:
            self.current_chapter["sections"].append(item)
        self.current_section = item
        self.sections.append({
            "type": "section_start",
            "id": sec_id,
            "title": text
        })

    def add_h3(self, text):
        self.sections.append({
            "type": "h3",
            "title": text
        })

    def add_body(self, text):
        self.sections.append({
            "type": "body",
            "text": text
        })

    def add_bullet(self, title, text):
        self.sections.append({
            "type": "bullet",
            "title": title,
            "text": text
        })

    def add_callout(self, title, text):
        self.sections.append({
            "type": "callout",
            "title": title,
            "text": text
        })

    def add_formula(self, title, formula_text, breakdown="", intuition=""):
        self.sections.append({
            "type": "formula",
            "title": title,
            "formula": formula_text,
            "breakdown": breakdown,
            "intuition": intuition
        })

    def add_math_solution(self, prob_title, given_data, formulas, solution_steps, exam_tips=""):
        self.sections.append({
            "type": "math_solution",
            "title": prob_title,
            "given": given_data,
            "formulas": formulas,
            "steps": solution_steps,
            "tips": exam_tips
        })

    def add_img(self, img_path, caption, width_in=5.8):
        # normalize path for web: images/filename.png
        filename = os.path.basename(img_path)
        web_src = f"images/{filename}"
        self.sections.append({
            "type": "image",
            "src": web_src,
            "caption": caption
        })

    def add_back_to_toc(self):
        # We handle back-to-top via interactive floating UI
        pass

    # Mock docx doc methods
    def add_table(self, rows=1, cols=4):
        tbl = MockWebTable(self, rows=rows, cols=cols)
        self.sections.append({
            "type": "table",
            "table_obj": tbl
        })
        return tbl

    def add_paragraph(self):
        return MockWebPara()

    def add_page_break(self):
        pass


class MockWebCell:
    def __init__(self):
        self._text = ""
        self.paragraphs = [MockWebPara()]
        self._tc = self
        self.width = 0

    @property
    def text(self):
        if self._text:
            return self._text
        return "".join(r.text for p in self.paragraphs for r in p.runs)

    @text.setter
    def text(self, val):
        self._text = str(val) if val is not None else ""
        self.paragraphs = [MockWebPara()]
        self.paragraphs[0].runs = [MockWebRun(self._text)]

    def get_or_add_tcPr(self): return self
    def append(self, x): pass

class MockWebRow:
    def __init__(self, cols=4):
        self.cells = [MockWebCell() for _ in range(cols)]

class MockWebTable:
    def __init__(self, collector, rows=1, cols=4):
        self.collector = collector
        self.rows = [MockWebRow(cols) for _ in range(rows)]
        self.alignment = 0

    def add_row(self):
        ncols = len(self.rows[0].cells) if self.rows else 4
        r = MockWebRow(ncols)
        self.rows.append(r)
        return r

class MockWebPara:
    def __init__(self):
        self.runs = []
        self.paragraph_format = self
        self.space_before = 0
        self.space_after = 0
        self.line_spacing = 1
        self.alignment = 0
    def add_run(self, text=""):
        r = MockWebRun(text)
        self.runs.append(r)
        return r

class MockWebRun:
    def __init__(self, text=""):
        self.text = text


def text_to_html(text):
    if not text:
        return ""
    safe = html.escape(text)
    # convert newlines to <br> or paragraphs
    lines = safe.split("\n")
    return "<br>".join(lines)


def build_web_ebook():
    print("Initializing Web eBook Collector...")
    collector = WebDocCollector()

    helpers = {
        "add_h1": collector.add_h1,
        "add_h2": collector.add_h2,
        "add_h3": collector.add_h3,
        "add_body": collector.add_body,
        "add_bullet": collector.add_bullet,
        "add_callout": collector.add_callout,
        "add_formula": collector.add_formula,
        "add_math_solution": collector.add_math_solution,
        "add_img": collector.add_img,
        "add_back_to_toc": collector.add_back_to_toc,
        "set_font": lambda *args, **kwargs: None,
        "COLOR_PRIMARY": None, "COLOR_SECONDARY": None, "COLOR_TEXT": None
    }

    # Execute all content generators
    print("Collecting Part 1 (Chapters 1-3)...")
    book_content_part1.write_part1(collector, helpers)

    print("Collecting Part 2 (Chapters 4-6)...")
    book_content_part2.write_part2(collector, helpers)

    print("Collecting Part 3 (Chapters 7-8: 42 Solved Math Types)...")
    book_content_part3.write_part3(collector, helpers)

    print("Collecting Parts 4 & 5 (Chapters 9-14: Viva, Field SOPs, Glossary)...")
    book_content_part4_5.write_part4_5(collector, helpers)

    print(f"Total sections collected: {len(collector.sections)}")
    print(f"Total chapters collected: {len(collector.toc_tree)}")

    # Ensure output directory exists
    docs_dir = os.path.join(os.getcwd(), "docs")
    images_dir = os.path.join(docs_dir, "images")
    os.makedirs(images_dir, exist_ok=True)

    # Copy diagrams
    src_images = os.path.join(os.getcwd(), "images")
    if os.path.exists(src_images):
        for img_name in os.listdir(src_images):
            s = os.path.join(src_images, img_name)
            d = os.path.join(images_dir, img_name)
            if os.path.isfile(s):
                shutil.copy2(s, d)
        print("Diagrams copied successfully to docs/images/")

    # Generate index.html, style.css, app.js
    generate_html(docs_dir, collector)
    generate_css(docs_dir)
    generate_js(docs_dir)
    print("Web eBook generation complete! All files generated in docs/")

def generate_html(docs_dir, collector):
    html_path = os.path.join(docs_dir, "index.html")

    # Render TOC Navigation items
    toc_nav_html = []
    
    # Static top items
    toc_nav_html.append('''
      <div class="toc-group">
        <div class="toc-group-header">📁 প্রারম্ভিক অংশ</div>
        <ul class="toc-list">
          <li><a href="#sec_prologue" class="toc-link" data-target="sec_prologue"><span class="toc-icon">📖</span> প্রারম্ভিক কথা ও প্রয়োজনীয়তা</a></li>
          <li><a href="#sec_author_bio" class="toc-link" data-target="sec_author_bio"><span class="toc-icon">👤</span> গ্রন্থকার পরিচিতি ও বায়োডাটা</a></li>
          <li><a href="#sec_foreword" class="toc-link" data-target="sec_foreword"><span class="toc-icon">🎯</span> নিয়োগ পরীক্ষার প্রস্তুতি কৌশল</a></li>
          <li><a href="#sec_calculators" class="toc-link highlight-calc" data-target="sec_calculators"><span class="toc-icon">🧮</span> ইন্টারেক্টিভ ক্যালকুলেটর উইজেট</a></li>
        </ul>
      </div>
    ''')

    for ch in collector.toc_tree:
        sec_items = []
        for s in ch["sections"]:
            sec_items.append(f'<li><a href="#{s["id"]}" class="toc-link toc-sublink" data-target="{s["id"]}"><span class="toc-dot">•</span> {html.escape(s["title"])}</a></li>')
        sec_html = "".join(sec_items)
        
        toc_nav_html.append(f'''
          <div class="toc-group">
            <div class="toc-chapter-item">
              <a href="#{ch["id"]}" class="toc-link toc-chapter-link" data-target="{ch["id"]}">
                <span class="toc-icon">📖</span>
                <span class="toc-title">{html.escape(ch["title"])}</span>
              </a>
            </div>
            <ul class="toc-list">
              {sec_html}
            </ul>
          </div>
        ''')

    toc_rendered = "\n".join(toc_nav_html)

    # Render Main Book Body Content
    body_html_parts = []

    # Front matter: Cover and Author Bio
    body_html_parts.append('''
    <!-- HERO COVER SECTION -->
    <section class="book-hero-card" id="sec_prologue">
      <div class="hero-top-badge">
        <span class="status-indicator"></span>
        <span>গণপ্রজাতন্ত্রী বাংলাদেশ • টেলিযোগাযোগ ও আইসিটি খাতের প্রকৌশলী নিয়োগ সহায়িকা</span>
      </div>
      <h1 class="book-main-title">টেলিকমিউনিকেশন ইঞ্জিনিয়ারিং<br><span class="gradient-text">জব প্রিপারেশন গাইড</span></h1>
      <p class="book-subtitle">A Comprehensive Preparation Handbook for BTRC, BTCL, Teletalk, BSCPLC, BSCL, MNOs & NTTN Engineers</p>
      
      <div class="target-badges-ribbon">
        <span class="badge">BTRC</span>
        <span class="badge">BTCL</span>
        <span class="badge">TELETALK</span>
        <span class="badge">BSCPLC</span>
        <span class="badge">BSCL</span>
        <span class="badge">BCS (ICT)</span>
        <span class="badge">MNOs (GP, Robi, BL)</span>
        <span class="badge">NTTN</span>
        <span class="badge">VENDORS</span>
      </div>

      <div class="cover-art-container">
        <img src="images/telecom_cover_art.png" alt="Telecom Job Preparation Guide Book Cover Art" class="cover-art-image" onerror="this.style.display='none'">
      </div>

      <div class="hero-highlights-grid">
        <div class="highlight-item">
          <div class="h-icon">📚</div>
          <div class="h-text"><strong>১৪টি পূর্ণাঙ্গ অধ্যায়</strong><p>কোর কমিউনিকেশন, ওয়্যারলেস, ফাইবার, আইপি নেটওয়ার্কিং ও বিটিআরসি ফ্রেমওয়ার্ক</p></div>
        </div>
        <div class="highlight-item">
          <div class="h-icon">🧮</div>
          <div class="h-text"><strong>৪২টি মাস্টার গাণিতিক টাইপ</strong><p>স্টেপ-বাই-স্টেপ সমাধান, শর্টকাট মেথড ও পরীক্ষার ভুল এড়ানোর টিপস</p></div>
        </div>
        <div class="highlight-item">
          <div class="h-icon">🎯</div>
          <div class="h-text"><strong>বিগত ১০ বছরের নিয়োগ সল্যুশন</strong><p>বুয়েট (BUET), এমআইএসটি (MIST) ও বিপিএসসি প্যাটার্ন পরীক্ষার প্রশ্ন সমাধান</p></div>
        </div>
        <div class="highlight-item">
          <div class="h-icon">💼</div>
          <div class="h-text"><strong>১২৫+ ভাইভা ও ১০টি ফিল্ড SOP</strong><p>৫টি বিশেষজ্ঞ প্যানেলের প্রশ্ন ও বাস্তবসম্মত নেটওয়ার্ক ট্রাবলশুটিং গাইড</p></div>
        </div>
      </div>
    </section>

    <!-- AUTHOR BIO SECTION -->
    <section class="author-bio-card glass-panel" id="sec_author_bio">
      <div class="author-card-header">
        <div class="author-avatar-badge">
          <span class="avatar-icon">👤</span>
        </div>
        <div class="author-title-area">
          <div class="author-label">গ্রন্থকার ও গবেষক পরিচিতি</div>
          <h2 class="author-name">প্রকৌশলী মোহাম্মদ আব্দুল মান্নান</h2>
          <div class="author-desig">উপ সহকারী পরিচালক (কারিগরি), স্পেকট্রাম বিভাগ • বাংলাদেশ টেলিযোগাযোগ নিয়ন্ত্রণ কমিশন (বিটিআরসি)</div>
          <div class="author-degrees">বিএসসি ইন ইইই (উত্তরা ইউনিভার্সিটি) • ৪-বছর মেয়াদি ডিপ্লোমা ইন টেলিকমিউনিকেশন টেকনোলজি (বিটিইবি)</div>
        </div>
      </div>

      <div class="author-bio-body">
        <p>প্রকৌশলী মোহাম্মদ আব্দুল মান্নান বাংলাদেশের টেলিযোগাযোগ খাত ও স্পেকট্রাম ব্যবস্থাপনায় ৭+ বছরেরও বেশি সময় ধরে কর্মরত একজন দক্ষ ও নিবেদিতপ্রাণ কর্মকর্তা। ২০১৮ সালের ডিসেম্বর থেকে তিনি দেশের সর্বোচ্চ টেলিযোগাযোগ নিয়ন্ত্রক সংস্থা বিটিআরসি (BTRC)-র স্পেকট্রাম বিভাগে (Spectrum Division) উপ সহকারী পরিচালক (কারিগরি) হিসেবে দায়িত্ব পালন করছেন। তাঁর মূল কাজের ক্ষেত্রের মধ্যে রয়েছে জাতীয় রেডিও ফ্রিকোয়েন্সি মনিটরিং, সেলুলার নেটওয়ার্ক কোয়ালিটি অব সার্ভিস (QoS/QoE ড্রাইভ টেস্ট), ইন্টারফারেন্স নিরসন এবং ৪জি ও ৫জি আইএমটি ব্যান্ড অডিট। এর পাশাপাশি তিনি ঢাকা বিশ্ববিদ্যালয়ের আইআইটি (IIT, DU) থেকে কৃত্রিম বুদ্ধিমত্তা ও পাইথন ভিত্তিক ডাটা অ্যানালিটিক্স এবং গুগল (Google) থেকে এআই ফান্ডামেন্টালস সম্পন্ন করেছেন।</p>
      </div>

      <div class="author-actions-grid">
        <a href="https://abmannan1761.github.io/my-portfolio/" target="_blank" rel="noopener noreferrer" class="btn btn-primary">
          <span>🌐 লাইভ পোর্টফোলিও ভিজিট করুন</span>
        </a>
        <a href="https://github.com/AbMannan1761" target="_blank" rel="noopener noreferrer" class="btn btn-secondary">
          <span>💻 গিটহাব প্রোফাইল</span>
        </a>
        <a href="https://www.linkedin.com/in/mohammad-abdul-mannan-38876114a/" target="_blank" rel="noopener noreferrer" class="btn btn-outline">
          <span>💼 লিঙ্কডইন প্রোফাইল</span>
        </a>
        <a href="mailto:mannan@btrc.gov.bd" class="btn btn-outline">
          <span>✉️ mannan@btrc.gov.bd</span>
        </a>
      </div>
    </section>

    <!-- INTERACTIVE CALCULATORS SECTION -->
    <section class="calculators-hub-section glass-panel" id="sec_calculators">
      <div class="section-badge">🧮 লাইভ ইঞ্জিনিয়ারিং টুলস</div>
      <h2 class="section-title">টেলিকম ইন্টারেক্টিভ ক্যালকুলেটর উইজেট</h2>
      <p class="section-subtitle">ব্রাউজারেই সরাসরি মান ইনপুট দিয়ে তাৎক্ষণিক রেজাল্ট দেখুন ও গাণিতিক সমীকরণের সত্যতা যাচাই করুন।</p>

      <div class="calc-tabs">
        <button class="calc-tab-btn active" onclick="switchCalc('shannon')">১. শ্যানন ক্যাপাসিটি</button>
        <button class="calc-tab-btn" onclick="switchCalc('fspl')">২. পাথ লস (FSPL)</button>
        <button class="calc-tab-btn" onclick="switchCalc('fresnel')">৩. ফ্রেনেল জোন</button>
        <button class="calc-tab-btn" onclick="switchCalc('fiber')">৪. ফাইবার রিচ</button>
      </div>

      <!-- Calc 1: Shannon -->
      <div class="calc-card active" id="calc_shannon">
        <div class="calc-info">
          <h4>শ্যানন-হার্টলে চ্যানেল ক্যাপাসিটি ক্যালকুলেটর</h4>
          <p>সমীকরণ: <code>C = B × log₂(1 + SNR)</code></p>
        </div>
        <div class="calc-inputs">
          <div class="input-field">
            <label>চ্যানেল ব্যান্ডউইথ B (MHz):</label>
            <input type="number" id="shannon_bw" value="20" step="any">
          </div>
          <div class="input-field">
            <label>সিগন্যাল-টু-নয়েজ রেশিও SNR (dB):</label>
            <input type="number" id="shannon_snr" value="30" step="any">
          </div>
          <button class="btn btn-calc" onclick="calculateShannon()">হিসাব করুন</button>
        </div>
        <div class="calc-output" id="shannon_result">
          <span class="res-label">সর্বোচ্চ থিওরিটিক্যাল ক্যাপাসিটি (C):</span>
          <span class="res-value">199.3 Mbps</span>
        </div>
      </div>

      <!-- Calc 2: FSPL -->
      <div class="calc-card" id="calc_fspl">
        <div class="calc-info">
          <h4>ফ্রি-স্পেস পাথ লস (FSPL) ক্যালকুলেটর</h4>
          <p>সমীকরণ: <code>FSPL (dB) = 92.45 + 20 log₁₀(f_GHz) + 20 log₁₀(d_km)</code></p>
        </div>
        <div class="calc-inputs">
          <div class="input-field">
            <label>ফ্রিকোয়েন্সি f (GHz):</label>
            <input type="number" id="fspl_f" value="15" step="any">
          </div>
          <div class="input-field">
            <label>দূরত্ব d (km):</label>
            <input type="number" id="fspl_d" value="10" step="any">
          </div>
          <button class="btn btn-calc" onclick="calculateFSPL()">হিসাব করুন</button>
        </div>
        <div class="calc-output" id="fspl_result">
          <span class="res-label">মোট পাথ লস (FSPL):</span>
          <span class="res-value">135.97 dB</span>
        </div>
      </div>

      <!-- Calc 3: Fresnel Zone -->
      <div class="calc-card" id="calc_fresnel">
        <div class="calc-info">
          <h4>প্রথম ফ্রেনেল জোন ব্যাসার্ধ (R₁) ক্যালকুলেটর</h4>
          <p>সমীকরণ: <code>R₁ (meters) = 17.32 × √((d₁ × d₂) / (f_GHz × d))</code></p>
        </div>
        <div class="calc-inputs">
          <div class="input-field">
            <label>টাওয়ার ১ থেকে দূরত্ব d₁ (km):</label>
            <input type="number" id="fresnel_d1" value="5" step="any">
          </div>
          <div class="input-field">
            <label>টাওয়ার ২ থেকে দূরত্ব d₂ (km):</label>
            <input type="number" id="fresnel_d2" value="5" step="any">
          </div>
          <div class="input-field">
            <label>ফ্রিকোয়েন্সি f (GHz):</label>
            <input type="number" id="fresnel_f" value="7" step="any">
          </div>
          <button class="btn btn-calc" onclick="calculateFresnel()">হিসাব করুন</button>
        </div>
        <div class="calc-output" id="fresnel_result">
          <span class="res-label">প্রথম ফ্রেনেল জোন ব্যাসার্ধ (R₁):</span>
          <span class="res-value">10.35 মিটার (৬০% ক্লিয়ারেন্স: 6.21 মি.)</span>
        </div>
      </div>

      <!-- Calc 4: Fiber Reach -->
      <div class="calc-card" id="calc_fiber">
        <div class="calc-info">
          <h4>অপটিক্যাল ফাইবার সর্বোচ্চ ট্রান্সমিশন রিচ ক্যালকুলেটর</h4>
          <p>সমীকরণ: <code>L_max = (P_tx - P_sens - L_conn - Margin) / (α_fiber + α_splice)</code></p>
        </div>
        <div class="calc-inputs">
          <div class="input-field">
            <label>ট্রান্সমিটার পাওয়ার P_tx (dBm):</label>
            <input type="number" id="fiber_ptx" value="0" step="any">
          </div>
          <div class="input-field">
            <label>রিসিভার সেন্সিটিভিটি P_sens (dBm):</label>
            <input type="number" id="fiber_psens" value="-28" step="any">
          </div>
          <div class="input-field">
            <label>ফাইবার লস সহগ (dB/km):</label>
            <input type="number" id="fiber_alpha" value="0.225" step="any">
          </div>
          <div class="input-field">
            <label>ফিক্সড লস + মার্জিন (dB):</label>
            <input type="number" id="fiber_margin" value="4.0" step="any">
          </div>
          <button class="btn btn-calc" onclick="calculateFiber()">হিসাব করুন</button>
        </div>
        <div class="calc-output" id="fiber_result">
          <span class="res-label">সর্বোচ্চ বিস্তার দূরত্ব (L_max):</span>
          <span class="res-value">106.67 কিমি</span>
        </div>
      </div>
    </section>
    ''')

    # Convert collected sections to HTML
    for item in collector.sections:
        t = item["type"]
        if t == "chapter_start":
            body_html_parts.append(f'''
            <article class="chapter-block" id="{item["id"]}">
              <div class="chapter-banner">
                <span class="chapter-tag">অধ্যায়</span>
                <h2 class="chapter-title">{html.escape(item["title"])}</h2>
              </div>
            ''')
        elif t == "section_start":
            body_html_parts.append(f'''
            <div class="section-container" id="{item["id"]}">
              <h3 class="section-heading"><a href="#{item["id"]}" class="sec-anchor">#</a> {html.escape(item["title"])}</h3>
            ''')
        elif t == "h3":
            body_html_parts.append(f'<h4 class="subtopic-heading">{html.escape(item["title"])}</h4>')
        elif t == "body":
            p_content = text_to_html(item["text"])
            body_html_parts.append(f'<p class="body-text">{p_content}</p>')
        elif t == "bullet":
            b_title = html.escape(item["title"])
            b_text = text_to_html(item["text"])
            body_html_parts.append(f'''
            <div class="bullet-card">
              <div class="bullet-header">📌 <strong>{b_title}</strong></div>
              <div class="bullet-body">{b_text}</div>
            </div>
            ''')
        elif t == "callout":
            c_title = html.escape(item["title"])
            c_text = text_to_html(item["text"])
            body_html_parts.append(f'''
            <div class="callout-card">
              <div class="callout-icon">💡</div>
              <div class="callout-content">
                <h5 class="callout-title">{c_title}</h5>
                <div class="callout-text">{c_text}</div>
              </div>
            </div>
            ''')
        elif t == "formula":
            f_title = html.escape(item["title"])
            f_math = html.escape(item["formula"])
            f_break = text_to_html(item["breakdown"])
            f_int = text_to_html(item["intuition"])
            
            breakdown_html = f'<div class="formula-breakdown"><h6>🔍 প্যারামিটার বিশ্লেষণ:</h6><p>{f_break}</p></div>' if f_break else ''
            intuition_html = f'<div class="formula-intuition"><h6>💡 বাস্তব প্রয়োগ ও ধারণা:</h6><p>{f_int}</p></div>' if f_int else ''

            body_html_parts.append(f'''
            <div class="formula-card">
              <div class="formula-top">
                <span class="formula-badge">📐 সূত্র</span>
                <h4 class="formula-title">{f_title}</h4>
              </div>
              <div class="formula-code-box">
                <code>{f_math}</code>
              </div>
              {breakdown_html}
              {intuition_html}
            </div>
            ''')
        elif t == "math_solution":
            m_title = html.escape(item["title"])
            m_given = text_to_html(item["given"])
            m_formulas = text_to_html(item["formulas"])
            m_steps = text_to_html(item["steps"])
            m_tips = text_to_html(item["tips"])

            tips_html = f'<div class="math-tips"><span class="tips-tag">⚠️ পরীক্ষার টিপস:</span> {m_tips}</div>' if m_tips else ''

            body_html_parts.append(f'''
            <div class="math-solution-card">
              <div class="math-card-header">
                <span class="math-tag">🧮 সমাধান</span>
                <h4 class="math-title">{m_title}</h4>
              </div>
              <div class="math-given-box">
                <span class="box-label">দেওয়া আছে / উপাত্ত:</span>
                <div class="box-content">{m_given}</div>
              </div>
              <div class="math-formula-box">
                <span class="box-label">প্রয়োজনীয় সমীকরণ:</span>
                <div class="box-content">{m_formulas}</div>
              </div>
              <div class="math-steps-box">
                <span class="box-label">স্টেপ-বাই-স্টেপ সমাধান:</span>
                <div class="box-content">{m_steps}</div>
              </div>
              {tips_html}
            </div>
            ''')
        elif t == "image":
            i_src = html.escape(item["src"])
            i_cap = html.escape(item["caption"])
            body_html_parts.append(f'''
            <figure class="content-figure">
              <img src="{i_src}" alt="{i_cap}" loading="lazy" class="figure-img" onerror="this.parentElement.style.display='none'">
              <figcaption class="figure-caption">চিত্র: {i_cap}</figcaption>
            </figure>
            ''')
        elif t == "table":
            tbl_obj = item["table_obj"]
            if tbl_obj.rows:
                table_html = ['<div class="table-responsive"><table class="telecom-table">']
                header_row = tbl_obj.rows[0]
                table_html.append('<thead><tr>')
                for c in header_row.cells:
                    cell_text = c.text.strip()
                    table_html.append(f'<th>{html.escape(cell_text)}</th>')
                table_html.append('</tr></thead>')
                table_html.append('<tbody>')
                for row in tbl_obj.rows[1:]:
                    table_html.append('<tr>')
                    for c in row.cells:
                        cell_text = c.text.strip()
                        table_html.append(f'<td>{text_to_html(cell_text)}</td>')
                    table_html.append('</tr>')
                table_html.append('</tbody></table></div>')
                body_html_parts.append("\n".join(table_html))

    book_content_rendered = "\n".join(body_html_parts)

    full_html = f'''<!DOCTYPE html>
<html lang="bn" data-theme="dark">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>টেলিকমিউনিকেশন ইঞ্জিনিয়ারিং জব প্রিপারেশন গাইড | প্রকৌশলী মোহাম্মদ আব্দুল মান্নান</title>
  <meta name="description" content="বিটিআরসি, বিটিসিএল, টেলিটক, বিএসসিএল ও এমএনও ইঞ্জিনিয়ার নিয়োগ পরীক্ষার জন্য একটি পূর্ণাঙ্গ প্রস্তুতি সহায়িকা। গ্রন্থনায়: প্রকৌশলী মোহাম্মদ আব্দুল মান্নান, উপ সহকারী পরিচালক (বিটিআরসি)।">
  <meta name="author" content="Engr. Mohammad Abdul Mannan">

  <!-- Google Fonts: Hind Siliguri for Bengali & Plus Jakarta Sans / JetBrains Mono -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Hind+Siliguri:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">

  <link rel="stylesheet" href="style.css">
</head>
<body>
  <!-- Reading Progress Bar -->
  <div class="reading-progress-bar" id="progressBar"></div>

  <!-- Main App Shell -->
  <div class="app-shell">
    
    <!-- Top Navigation Bar -->
    <header class="top-nav" id="topNav">
      <div class="nav-left">
        <button class="icon-btn mobile-menu-toggle" id="menuToggle" aria-label="Toggle Sidebar">
          <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="3" y1="12" x2="21" y2="12"></line><line x1="3" y1="6" x2="21" y2="6"></line><line x1="3" y1="18" x2="21" y2="18"></line></svg>
        </button>
        <a href="#sec_prologue" class="nav-brand">
          <span class="brand-badge">BTRC • BUET • 5G</span>
          <span class="brand-title">টেলিকম জব গাইড</span>
        </a>
      </div>

      <div class="nav-search-container">
        <svg class="search-icon" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>
        <input type="text" id="globalSearch" placeholder="যেকোনো বিষয়, সূত্র বা কি-ওয়ার্ড খুঁজুন (Ctrl + K)..." autocomplete="off">
        <span class="search-count" id="searchCount"></span>
      </div>

      <div class="nav-actions">
        <button class="action-btn" id="calcBtn" title="ইন্টারেক্টিভ ক্যালকুলেটর" onclick="document.getElementById('sec_calculators').scrollIntoView({{behavior:'smooth'}})">
          <span class="btn-icon">🧮</span>
          <span class="btn-label">ক্যালকুলেটর</span>
        </button>

        <button class="action-btn" id="fontSizeBtn" title="ফন্ট সাইজ টগল" onclick="toggleFontSize()">
          <span class="btn-label">A+</span>
        </button>

        <button class="action-btn" id="themeToggle" title="ডার্ক / লাইট থিম পরিবর্তন">
          <span class="theme-icon-light">☀️</span>
          <span class="theme-icon-dark">🌙</span>
        </button>

        <button class="action-btn print-btn" onclick="window.print()" title="প্রিন্ট বা পিডিএফ আকারে সংরক্ষণ">
          <span class="btn-icon">🖨️</span>
        </button>
      </div>
    </header>

    <div class="app-body">
      <!-- Sidebar Navigation -->
      <aside class="sidebar" id="sidebar">
        <div class="sidebar-header">
          <div class="sidebar-meta">
            <span class="meta-label">লেখক:</span>
            <strong class="meta-name">প্রকৌশলী মোহাম্মদ আব্দুল মান্নান</strong>
            <span class="meta-role">উপ সহকারী পরিচালক (কারিগরি), বিটিআরসি</span>
          </div>
        </div>

        <div class="sidebar-search-mobile">
          <input type="text" id="mobileSearch" placeholder="টপিক বা সূত্র খুঁজুন...">
        </div>

        <nav class="sidebar-nav" id="sidebarNav">
          {toc_rendered}
        </nav>
      </aside>

      <!-- Sidebar Backdrop on Mobile -->
      <div class="sidebar-backdrop" id="sidebarBackdrop"></div>

      <!-- Main Reader Content Area -->
      <main class="content-canvas" id="contentCanvas">
        <div class="content-wrapper">
          {book_content_rendered}

          <!-- Book Footer -->
          <footer class="book-footer">
            <div class="footer-card">
              <div class="footer-logo">📡</div>
              <h3 class="footer-title">টেলিকমিউনিকেশন ইঞ্জিনিয়ারিং জব প্রিপারেশন গাইড</h3>
              <p class="footer-copy">© ২০২৬–২০২৭ সর্বস্বত্ব সংরক্ষিত • প্রকৌশলী মোহাম্মদ আব্দুল মান্নান, বিএসসি ইন ইইই, উপ সহকারী পরিচালক (বিটিআরসি)</p>
              <div class="footer-links">
                <a href="https://abmannan1761.github.io/my-portfolio/" target="_blank">লেখক পোর্টফোলিও</a>
                <span>•</span>
                <a href="https://github.com/AbMannan1761" target="_blank">গিটহাব রিপোজিটরি</a>
                <span>•</span>
                <a href="mailto:mannan@btrc.gov.bd">mannan@btrc.gov.bd</a>
              </div>
              <p class="footer-note">গিটহাব পেজেস (GitHub Pages) দ্বারা চালিত ও রিয়েল-টাইম সিঙ্ক্রোনাইজড।</p>
            </div>
          </footer>
        </div>
      </main>
    </div>

  </div>

  <!-- Floating Back to Top Button -->
  <button class="back-to-top" id="backToTop" title="উপরে ফিরে যান">⬆️</button>

  <script src="app.js"></script>
</body>
</html>
'''

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(full_html)
    print("docs/index.html created successfully.")


def generate_css(docs_dir):
    css_path = os.path.join(docs_dir, "style.css")
    css_content = '''/* =========================================================================
   STYLE.CSS - PREMIUM TELECOM EBOOK WEB APPLICATION
   Author: Engr. Mohammad Abdul Mannan | BTRC Spectrum Division
   ========================================================================= */

:root {
  /* Color Palette - Light Theme */
  --bg-primary: #f8fafc;
  --bg-secondary: #ffffff;
  --bg-sidebar: #ffffff;
  --bg-card: #ffffff;
  --border-color: #e2e8f0;
  --border-focus: #3b82f6;

  --color-primary: #0a2540;
  --color-secondary: #1e3a8a;
  --color-accent: #2563eb;
  --color-text: #1e293b;
  --color-muted: #64748b;
  --color-badge-bg: #e0f2fe;
  --color-badge-text: #0369a1;

  --card-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -2px rgba(0, 0, 0, 0.05);
  --card-hover-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.08), 0 4px 6px -4px rgba(0, 0, 0, 0.04);

  --font-bengali: 'Hind Siliguri', -apple-system, BlinkMacSystemFont, sans-serif;
  --font-ui: 'Plus Jakarta Sans', sans-serif;
  --font-mono: 'JetBrains Mono', Consolas, monospace;

  --topbar-height: 64px;
  --sidebar-width: 320px;
}

[data-theme="dark"] {
  /* Color Palette - Dark Theme */
  --bg-primary: #070d19;
  --bg-secondary: #0f172a;
  --bg-sidebar: #0b1329;
  --bg-card: #111e38;
  --border-color: #1e293b;
  --border-focus: #60a5fa;

  --color-primary: #f8fafc;
  --color-secondary: #93c5fd;
  --color-accent: #38bdf8;
  --color-text: #e2e8f0;
  --color-muted: #94a3b8;
  --color-badge-bg: rgba(56, 189, 248, 0.15);
  --color-badge-text: #7dd3fc;

  --card-shadow: 0 4px 12px rgba(0, 0, 0, 0.35);
  --card-hover-shadow: 0 12px 24px rgba(0, 0, 0, 0.5);
}

/* Base Resets */
* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

html {
  scroll-behavior: smooth;
  font-size: 16px;
}

body {
  font-family: var(--font-bengali);
  background-color: var(--bg-primary);
  color: var(--color-text);
  line-height: 1.7;
  overflow-x: hidden;
  transition: background-color 0.25s ease, color 0.25s ease;
}

/* Reading Progress Bar */
.reading-progress-bar {
  position: fixed;
  top: 0;
  left: 0;
  height: 4px;
  width: 0%;
  background: linear-gradient(90deg, #2563eb, #06b6d4, #10b981);
  z-index: 1000;
  transition: width 0.1s ease-out;
}

/* App Shell Layout */
.app-shell {
  display: flex;
  flex-direction: column;
  min-height: 100vh;
}

/* Top Navigation */
.top-nav {
  position: sticky;
  top: 0;
  height: var(--topbar-height);
  background: var(--bg-secondary);
  border-bottom: 1px solid var(--border-color);
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0 1.5rem;
  z-index: 100;
  backdrop-filter: blur(12px);
  -webkit-backdrop-filter: blur(12px);
}

.nav-left {
  display: flex;
  align-items: center;
  gap: 1rem;
}

.mobile-menu-toggle {
  display: none;
  background: none;
  border: none;
  color: var(--color-text);
  cursor: pointer;
  padding: 6px;
  border-radius: 6px;
}

.mobile-menu-toggle:hover {
  background: var(--border-color);
}

.nav-brand {
  display: flex;
  flex-direction: column;
  text-decoration: none;
  color: inherit;
}

.brand-badge {
  font-family: var(--font-ui);
  font-size: 0.65rem;
  font-weight: 700;
  color: var(--color-accent);
  letter-spacing: 0.08em;
  text-transform: uppercase;
}

.brand-title {
  font-size: 1.15rem;
  font-weight: 700;
  color: var(--color-primary);
  line-height: 1.2;
}

/* Search Bar */
.nav-search-container {
  flex: 1;
  max-width: 480px;
  position: relative;
  margin: 0 1.5rem;
}

.search-icon {
  position: absolute;
  left: 14px;
  top: 50%;
  transform: translateY(-50%);
  color: var(--color-muted);
  pointer-events: none;
}

.nav-search-container input {
  width: 100%;
  padding: 8px 16px 8px 40px;
  background: var(--bg-primary);
  border: 1px solid var(--border-color);
  border-radius: 9999px;
  font-family: var(--font-bengali);
  font-size: 0.9rem;
  color: var(--color-text);
  outline: none;
  transition: all 0.2s ease;
}

.nav-search-container input:focus {
  border-color: var(--border-focus);
  box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.15);
}

.search-count {
  position: absolute;
  right: 14px;
  top: 50%;
  transform: translateY(-50%);
  font-size: 0.75rem;
  color: var(--color-muted);
  font-family: var(--font-ui);
}

/* Nav Actions */
.nav-actions {
  display: flex;
  align-items: center;
  gap: 0.5rem;
}

.action-btn {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 6px 12px;
  background: var(--bg-primary);
  border: 1px solid var(--border-color);
  border-radius: 8px;
  color: var(--color-text);
  font-family: var(--font-ui);
  font-size: 0.85rem;
  font-weight: 600;
  cursor: pointer;
  transition: all 0.2s ease;
}

.action-btn:hover {
  background: var(--border-color);
  border-color: var(--border-focus);
}

[data-theme="dark"] .theme-icon-light { display: inline-block; }
[data-theme="dark"] .theme-icon-dark { display: none; }
[data-theme="light"] .theme-icon-light { display: none; }
[data-theme="light"] .theme-icon-dark { display: inline-block; }

/* App Body Layout */
.app-body {
  display: flex;
  flex: 1;
}

/* Sidebar */
.sidebar {
  width: var(--sidebar-width);
  height: calc(100vh - var(--topbar-height));
  position: sticky;
  top: var(--topbar-height);
  background: var(--bg-sidebar);
  border-right: 1px solid var(--border-color);
  display: flex;
  flex-direction: column;
  overflow-y: auto;
  z-index: 80;
  transition: transform 0.3s ease;
}

.sidebar-header {
  padding: 1rem 1.25rem;
  border-bottom: 1px solid var(--border-color);
  background: var(--bg-secondary);
}

.sidebar-meta {
  display: flex;
  flex-direction: column;
  font-size: 0.82rem;
}

.meta-label {
  color: var(--color-muted);
  font-size: 0.75rem;
}

.meta-name {
  color: var(--color-primary);
  font-size: 0.95rem;
  font-weight: 700;
}

.meta-role {
  color: var(--color-secondary);
  font-size: 0.8rem;
  margin-top: 2px;
}

.sidebar-search-mobile {
  display: none;
  padding: 0.75rem 1rem;
  border-bottom: 1px solid var(--border-color);
}

.sidebar-search-mobile input {
  width: 100%;
  padding: 8px 12px;
  background: var(--bg-primary);
  border: 1px solid var(--border-color);
  border-radius: 6px;
  font-family: var(--font-bengali);
  color: var(--color-text);
}

.sidebar-nav {
  padding: 1rem 0.5rem;
  flex: 1;
}

.toc-group {
  margin-bottom: 1.25rem;
}

.toc-group-header {
  font-size: 0.78rem;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.05em;
  color: var(--color-muted);
  padding: 4px 12px;
  margin-bottom: 4px;
}

.toc-list {
  list-style: none;
}

.toc-chapter-item {
  margin-bottom: 2px;
}

.toc-link {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 6px 12px;
  border-radius: 6px;
  color: var(--color-text);
  text-decoration: none;
  font-size: 0.88rem;
  line-height: 1.4;
  transition: all 0.15s ease;
}

.toc-chapter-link {
  font-weight: 600;
  color: var(--color-primary);
}

.toc-sublink {
  font-size: 0.82rem;
  color: var(--color-muted);
  padding-left: 24px;
}

.toc-link:hover {
  background: var(--bg-primary);
  color: var(--color-accent);
}

.toc-link.active {
  background: var(--color-badge-bg);
  color: var(--color-badge-text);
  font-weight: 600;
}

.highlight-calc {
  background: rgba(16, 185, 129, 0.1) !important;
  color: #10b981 !important;
  font-weight: 600;
}

/* Content Canvas */
.content-canvas {
  flex: 1;
  min-width: 0;
  padding: 2.5rem 2rem;
  display: flex;
  justify-content: center;
}

.content-wrapper {
  width: 100%;
  max-width: 900px;
}

/* Hero Card */
.book-hero-card {
  background: var(--bg-card);
  border: 1px solid var(--border-color);
  border-radius: 16px;
  padding: 2.5rem;
  margin-bottom: 2rem;
  box-shadow: var(--card-shadow);
  text-align: center;
}

.hero-top-badge {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  padding: 4px 14px;
  background: var(--color-badge-bg);
  color: var(--color-badge-text);
  border-radius: 9999px;
  font-size: 0.82rem;
  font-weight: 600;
  margin-bottom: 1.25rem;
}

.status-indicator {
  width: 8px;
  height: 8px;
  background: #10b981;
  border-radius: 50%;
  display: inline-block;
  box-shadow: 0 0 8px #10b981;
}

.book-main-title {
  font-size: 2.3rem;
  font-weight: 800;
  line-height: 1.25;
  color: var(--color-primary);
  margin-bottom: 0.75rem;
}

.gradient-text {
  background: linear-gradient(135deg, #2563eb, #06b6d4);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
}

.book-subtitle {
  font-family: var(--font-ui);
  font-size: 0.95rem;
  color: var(--color-muted);
  font-style: italic;
  margin-bottom: 1.5rem;
}

.target-badges-ribbon {
  display: flex;
  flex-wrap: wrap;
  justify-content: center;
  gap: 8px;
  margin-bottom: 1.75rem;
}

.target-badges-ribbon .badge {
  background: var(--bg-primary);
  border: 1px solid var(--border-color);
  padding: 4px 10px;
  border-radius: 6px;
  font-family: var(--font-ui);
  font-size: 0.75rem;
  font-weight: 700;
  color: var(--color-secondary);
}

.cover-art-container {
  margin: 1.5rem auto;
  max-width: 500px;
}

.cover-art-image {
  width: 100%;
  border-radius: 12px;
  box-shadow: 0 10px 25px rgba(0,0,0,0.2);
}

.hero-highlights-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
  gap: 1rem;
  margin-top: 2rem;
  text-align: left;
}

.highlight-item {
  display: flex;
  align-items: flex-start;
  gap: 10px;
  background: var(--bg-primary);
  border: 1px solid var(--border-color);
  border-radius: 10px;
  padding: 1rem;
}

.highlight-item .h-icon {
  font-size: 1.5rem;
}

.highlight-item .h-text strong {
  display: block;
  font-size: 0.95rem;
  color: var(--color-primary);
  margin-bottom: 2px;
}

.highlight-item .h-text p {
  font-size: 0.8rem;
  color: var(--color-muted);
  line-height: 1.35;
}

/* Author Bio Card */
.author-bio-card {
  background: var(--bg-card);
  border: 1px solid var(--border-color);
  border-left: 5px solid var(--color-accent);
  border-radius: 14px;
  padding: 2rem;
  margin-bottom: 2rem;
  box-shadow: var(--card-shadow);
}

.author-card-header {
  display: flex;
  align-items: center;
  gap: 1.25rem;
  margin-bottom: 1.25rem;
}

.author-avatar-badge {
  width: 60px;
  height: 60px;
  border-radius: 50%;
  background: linear-gradient(135deg, #1e3a8a, #3b82f6);
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 1.8rem;
  color: #fff;
  flex-shrink: 0;
}

.author-title-area .author-label {
  font-size: 0.75rem;
  font-weight: 700;
  text-transform: uppercase;
  color: var(--color-muted);
}

.author-name {
  font-size: 1.4rem;
  font-weight: 800;
  color: var(--color-primary);
}

.author-desig {
  font-size: 0.95rem;
  font-weight: 600;
  color: var(--color-accent);
}

.author-degrees {
  font-size: 0.82rem;
  color: var(--color-muted);
}

.author-bio-body {
  font-size: 0.95rem;
  color: var(--color-text);
  line-height: 1.7;
  margin-bottom: 1.5rem;
}

.author-actions-grid {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
}

.btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  padding: 8px 18px;
  border-radius: 8px;
  font-family: var(--font-ui);
  font-size: 0.85rem;
  font-weight: 600;
  text-decoration: none;
  cursor: pointer;
  transition: all 0.2s ease;
}

.btn-primary {
  background: #2563eb;
  color: #ffffff;
  border: none;
}

.btn-primary:hover {
  background: #1d4ed8;
  transform: translateY(-1px);
}

.btn-secondary {
  background: #0f172a;
  color: #ffffff;
  border: 1px solid var(--border-color);
}

.btn-secondary:hover {
  background: #1e293b;
}

.btn-outline {
  background: transparent;
  color: var(--color-text);
  border: 1px solid var(--border-color);
}

.btn-outline:hover {
  background: var(--bg-primary);
  border-color: var(--border-focus);
}

/* Calculators Hub */
.calculators-hub-section {
  background: var(--bg-card);
  border: 1px solid var(--border-color);
  border-radius: 14px;
  padding: 2rem;
  margin-bottom: 2.5rem;
  box-shadow: var(--card-shadow);
}

.section-badge {
  display: inline-block;
  padding: 4px 10px;
  background: rgba(16, 185, 129, 0.15);
  color: #10b981;
  border-radius: 6px;
  font-size: 0.8rem;
  font-weight: 700;
  margin-bottom: 0.5rem;
}

.section-title {
  font-size: 1.5rem;
  font-weight: 800;
  color: var(--color-primary);
  margin-bottom: 0.25rem;
}

.section-subtitle {
  font-size: 0.9rem;
  color: var(--color-muted);
  margin-bottom: 1.5rem;
}

.calc-tabs {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  border-bottom: 1px solid var(--border-color);
  padding-bottom: 12px;
  margin-bottom: 1.5rem;
}

.calc-tab-btn {
  padding: 6px 14px;
  border-radius: 6px;
  background: var(--bg-primary);
  border: 1px solid var(--border-color);
  color: var(--color-text);
  font-family: var(--font-bengali);
  font-size: 0.85rem;
  font-weight: 600;
  cursor: pointer;
  transition: all 0.2s ease;
}

.calc-tab-btn:hover {
  border-color: var(--border-focus);
}

.calc-tab-btn.active {
  background: #2563eb;
  color: #ffffff;
  border-color: #2563eb;
}

.calc-card {
  display: none;
}

.calc-card.active {
  display: block;
}

.calc-info h4 {
  font-size: 1.1rem;
  font-weight: 700;
  color: var(--color-primary);
  margin-bottom: 4px;
}

.calc-info p code {
  font-family: var(--font-mono);
  background: var(--bg-primary);
  padding: 2px 6px;
  border-radius: 4px;
  font-size: 0.85rem;
  color: var(--color-accent);
}

.calc-inputs {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
  gap: 12px;
  align-items: flex-end;
  margin: 1.25rem 0;
}

.input-field {
  display: flex;
  flex-direction: column;
  gap: 4px;
}

.input-field label {
  font-size: 0.82rem;
  font-weight: 600;
  color: var(--color-muted);
}

.input-field input {
  padding: 8px 12px;
  border: 1px solid var(--border-color);
  border-radius: 6px;
  background: var(--bg-primary);
  color: var(--color-text);
  font-family: var(--font-mono);
  font-size: 0.95rem;
  outline: none;
}

.input-field input:focus {
  border-color: var(--border-focus);
}

.btn-calc {
  background: #10b981;
  color: #fff;
  border: none;
  padding: 9px 16px;
}

.btn-calc:hover {
  background: #059669;
}

.calc-output {
  display: flex;
  align-items: center;
  justify-content: space-between;
  background: var(--bg-primary);
  border: 1px solid var(--border-color);
  border-radius: 8px;
  padding: 12px 18px;
}

.calc-output .res-label {
  font-weight: 600;
  color: var(--color-muted);
  font-size: 0.9rem;
}

.calc-output .res-value {
  font-family: var(--font-mono);
  font-weight: 800;
  font-size: 1.2rem;
  color: #10b981;
}

/* Chapter & Section Styles */
.chapter-block {
  margin-top: 3.5rem;
  padding-top: 1rem;
  border-top: 2px dashed var(--border-color);
}

.chapter-banner {
  background: linear-gradient(135deg, #0a2540, #1e3a8a);
  color: #ffffff;
  padding: 1.5rem 2rem;
  border-radius: 12px;
  margin-bottom: 2rem;
  box-shadow: 0 4px 15px rgba(10, 37, 64, 0.25);
}

.chapter-tag {
  display: inline-block;
  background: rgba(255, 255, 255, 0.2);
  padding: 2px 10px;
  border-radius: 9999px;
  font-size: 0.75rem;
  font-weight: 700;
  letter-spacing: 0.05em;
  text-transform: uppercase;
  margin-bottom: 0.5rem;
}

.chapter-title {
  font-size: 1.6rem;
  font-weight: 800;
  line-height: 1.3;
}

.section-container {
  margin-top: 2.25rem;
  margin-bottom: 1.5rem;
}

.section-heading {
  font-size: 1.3rem;
  font-weight: 700;
  color: var(--color-secondary);
  border-bottom: 2px solid var(--border-color);
  padding-bottom: 6px;
  margin-bottom: 1rem;
  display: flex;
  align-items: center;
  gap: 8px;
}

.sec-anchor {
  color: var(--color-accent);
  text-decoration: none;
  opacity: 0.6;
}

.sec-anchor:hover {
  opacity: 1;
}

.subtopic-heading {
  font-size: 1.1rem;
  font-weight: 700;
  color: var(--color-primary);
  margin-top: 1.5rem;
  margin-bottom: 0.5rem;
}

.body-text {
  font-size: 1rem;
  margin-bottom: 1rem;
  line-height: 1.8;
  color: var(--color-text);
}

/* Bullet Cards */
.bullet-card {
  background: var(--bg-card);
  border: 1px solid var(--border-color);
  border-left: 4px solid var(--color-accent);
  border-radius: 8px;
  padding: 1rem 1.25rem;
  margin-bottom: 1rem;
  box-shadow: var(--card-shadow);
}

.bullet-header {
  font-size: 0.98rem;
  color: var(--color-primary);
  margin-bottom: 6px;
}

.bullet-body {
  font-size: 0.92rem;
  color: var(--color-text);
  line-height: 1.65;
}

/* Callout Cards */
.callout-card {
  display: flex;
  gap: 1rem;
  background: var(--bg-card);
  border: 1px solid var(--border-color);
  border-left: 4px solid #10b981;
  border-radius: 10px;
  padding: 1.25rem;
  margin: 1.25rem 0;
  box-shadow: var(--card-shadow);
}

.callout-icon {
  font-size: 1.5rem;
  flex-shrink: 0;
}

.callout-title {
  font-size: 1.05rem;
  font-weight: 700;
  color: var(--color-primary);
  margin-bottom: 4px;
}

.callout-text {
  font-size: 0.92rem;
  color: var(--color-text);
  line-height: 1.65;
}

/* Formula Cards */
.formula-card {
  background: var(--bg-card);
  border: 1px solid var(--border-color);
  border-radius: 12px;
  padding: 1.5rem;
  margin: 1.5rem 0;
  box-shadow: var(--card-shadow);
}

.formula-top {
  display: flex;
  align-items: center;
  gap: 10px;
  margin-bottom: 0.75rem;
}

.formula-badge {
  background: rgba(37, 99, 235, 0.12);
  color: #2563eb;
  padding: 2px 8px;
  border-radius: 4px;
  font-size: 0.75rem;
  font-weight: 700;
}

.formula-title {
  font-size: 1.15rem;
  font-weight: 700;
  color: var(--color-primary);
}

.formula-code-box {
  background: var(--bg-primary);
  border: 1px solid var(--border-color);
  border-radius: 8px;
  padding: 1rem;
  margin-bottom: 1rem;
  overflow-x: auto;
}

.formula-code-box code {
  font-family: var(--font-mono);
  font-size: 1.05rem;
  font-weight: 600;
  color: var(--color-accent);
}

.formula-breakdown h6,
.formula-intuition h6 {
  font-size: 0.85rem;
  font-weight: 700;
  color: var(--color-secondary);
  margin-top: 8px;
  margin-bottom: 2px;
}

.formula-breakdown p,
.formula-intuition p {
  font-size: 0.9rem;
  color: var(--color-muted);
  line-height: 1.6;
}

/* Math Solution Cards */
.math-solution-card {
  background: var(--bg-card);
  border: 1px solid var(--border-color);
  border-top: 4px solid #f59e0b;
  border-radius: 12px;
  padding: 1.5rem;
  margin: 1.5rem 0;
  box-shadow: var(--card-shadow);
}

.math-card-header {
  display: flex;
  align-items: center;
  gap: 10px;
  margin-bottom: 1rem;
}

.math-tag {
  background: rgba(245, 158, 11, 0.15);
  color: #d97706;
  padding: 2px 8px;
  border-radius: 4px;
  font-size: 0.75rem;
  font-weight: 700;
}

.math-title {
  font-size: 1.15rem;
  font-weight: 800;
  color: var(--color-primary);
}

.math-given-box,
.math-formula-box,
.math-steps-box {
  background: var(--bg-primary);
  border: 1px solid var(--border-color);
  border-radius: 8px;
  padding: 10px 14px;
  margin-bottom: 0.75rem;
}

.box-label {
  display: block;
  font-size: 0.78rem;
  font-weight: 700;
  text-transform: uppercase;
  color: var(--color-muted);
  margin-bottom: 4px;
}

.box-content {
  font-size: 0.92rem;
  line-height: 1.65;
}

.math-tips {
  background: rgba(239, 68, 68, 0.08);
  border: 1px solid rgba(239, 68, 68, 0.2);
  border-radius: 8px;
  padding: 8px 12px;
  font-size: 0.88rem;
  color: #ef4444;
  margin-top: 0.75rem;
}

.tips-tag {
  font-weight: 700;
}

/* Figures & Images */
.content-figure {
  margin: 1.75rem auto;
  text-align: center;
}

.figure-img {
  max-width: 100%;
  height: auto;
  border-radius: 10px;
  border: 1px solid var(--border-color);
  box-shadow: var(--card-shadow);
}

.figure-caption {
  font-size: 0.85rem;
  font-style: italic;
  color: var(--color-muted);
  margin-top: 6px;
}

/* Footer */
.book-footer {
  margin-top: 5rem;
  border-top: 1px solid var(--border-color);
  padding-top: 2.5rem;
  padding-bottom: 3rem;
  text-align: center;
}

.footer-logo {
  font-size: 2.5rem;
  margin-bottom: 0.5rem;
}

.footer-title {
  font-size: 1.2rem;
  font-weight: 700;
  color: var(--color-primary);
  margin-bottom: 0.5rem;
}

.footer-copy {
  font-size: 0.85rem;
  color: var(--color-muted);
  margin-bottom: 1rem;
}

.footer-links {
  display: flex;
  justify-content: center;
  align-items: center;
  gap: 12px;
  font-size: 0.85rem;
  margin-bottom: 0.75rem;
}

.footer-links a {
  color: var(--color-accent);
  text-decoration: none;
}

.footer-links a:hover {
  text-decoration: underline;
}

.footer-note {
  font-size: 0.75rem;
  color: var(--color-muted);
}

/* Back to Top */
.back-to-top {
  position: fixed;
  bottom: 24px;
  right: 24px;
  width: 44px;
  height: 44px;
  border-radius: 50%;
  background: #2563eb;
  color: #fff;
  border: none;
  font-size: 1.1rem;
  display: flex;
  align-items: center;
  justify-content: center;
  box-shadow: 0 4px 12px rgba(37, 99, 235, 0.4);
  cursor: pointer;
  opacity: 0;
  pointer-events: none;
  transition: all 0.25s ease;
  z-index: 90;
}

.back-to-top.visible {
  opacity: 1;
  pointer-events: auto;
}

.back-to-top:hover {
  background: #1d4ed8;
  transform: translateY(-2px);
}

/* Responsive Media Queries */
@media (max-width: 1024px) {
  .nav-search-container {
    max-width: 280px;
  }
}

@media (max-width: 768px) {
  .mobile-menu-toggle {
    display: inline-flex;
  }

  .nav-search-container {
    display: none;
  }

  .sidebar {
    position: fixed;
    top: var(--topbar-height);
    left: 0;
    bottom: 0;
    width: 290px;
    transform: translateX(-100%);
    box-shadow: 0 0 25px rgba(0,0,0,0.3);
  }

  .sidebar.open {
    transform: translateX(0);
  }

  .sidebar-backdrop {
    display: none;
    position: fixed;
    top: var(--topbar-height);
    left: 0;
    right: 0;
    bottom: 0;
    background: rgba(0,0,0,0.5);
    z-index: 75;
  }

  .sidebar-backdrop.active {
    display: block;
  }

  .sidebar-search-mobile {
    display: block;
  }

  .content-canvas {
    padding: 1.5rem 1rem;
  }

  .book-main-title {
    font-size: 1.8rem;
  }

  .btn-label {
    display: none;
  }

  .author-card-header {
    flex-direction: column;
    text-align: center;
  }
}

/* Print Styles */
@media print {
  .top-nav, .sidebar, .reading-progress-bar, .back-to-top, .calc-inputs, .btn-calc, .nav-actions {
    display: none !important;
  }
  .app-body {
    display: block !important;
  }
  .content-canvas {
    padding: 0 !important;
    max-width: 100% !important;
  }
  .chapter-block {
    page-break-before: always;
  }
  body {
    background: #fff !important;
    color: #000 !important;
    font-size: 11pt;
  }
}
'''
    with open(css_path, "w", encoding="utf-8") as f:
        f.write(css_content)
    print("docs/style.css created successfully.")


def generate_js(docs_dir):
    js_path = os.path.join(docs_dir, "app.js")
    js_content = '''// =========================================================================
// APP.JS - INTERACTIVE TELECOM EBOOK ENGINE
// Author: Engr. Mohammad Abdul Mannan | BTRC Spectrum Division
// =========================================================================

document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  initReadingProgress();
  initSidebar();
  initSearch();
  initScrollSpy();
  initBackToTop();
});

// 1. Dark / Light Theme Toggle with LocalStorage
function initTheme() {
  const themeToggle = document.getElementById('themeToggle');
  const savedTheme = localStorage.getItem('telecom_ebook_theme') || 'dark';
  document.documentElement.setAttribute('data-theme', savedTheme);

  if (themeToggle) {
    themeToggle.addEventListener('click', () => {
      const currentTheme = document.documentElement.getAttribute('data-theme');
      const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', newTheme);
      localStorage.setItem('telecom_ebook_theme', newTheme);
    });
  }
}

// 2. Reading Progress Bar
function initReadingProgress() {
  const bar = document.getElementById('progressBar');
  window.addEventListener('scroll', () => {
    const totalHeight = document.documentElement.scrollHeight - window.innerHeight;
    if (totalHeight > 0) {
      const progress = (window.scrollY / totalHeight) * 100;
      bar.style.width = `${progress}%`;
    }
  });
}

// 3. Mobile Sidebar Toggle & Backdrop
function initSidebar() {
  const toggleBtn = document.getElementById('menuToggle');
  const sidebar = document.getElementById('sidebar');
  const backdrop = document.getElementById('sidebarBackdrop');

  if (toggleBtn && sidebar && backdrop) {
    toggleBtn.addEventListener('click', () => {
      sidebar.classList.toggle('open');
      backdrop.classList.toggle('active');
    });

    backdrop.addEventListener('click', () => {
      sidebar.classList.remove('open');
      backdrop.classList.remove('active');
    });

    // Close on link click on mobile
    const links = sidebar.querySelectorAll('.toc-link');
    links.forEach(l => {
      l.addEventListener('click', () => {
        if (window.innerWidth <= 768) {
          sidebar.classList.remove('open');
          backdrop.classList.remove('active');
        }
      });
    });
  }
}

// 4. Real-time Full Text Search
function initSearch() {
  const searchInput = document.getElementById('globalSearch');
  const mobileSearch = document.getElementById('mobileSearch');
  const countDisplay = document.getElementById('searchCount');

  function handleSearch(query) {
    const q = query.trim().toLowerCase();
    const tocLinks = document.querySelectorAll('.sidebar .toc-link');
    const sections = document.querySelectorAll('.chapter-block, .section-container, .formula-card, .math-solution-card, .bullet-card');

    if (!q) {
      tocLinks.forEach(l => l.parentElement.style.display = '');
      sections.forEach(s => s.style.display = '');
      if (countDisplay) countDisplay.textContent = '';
      return;
    }

    let matchCount = 0;

    // Filter content blocks
    sections.forEach(sec => {
      const text = sec.textContent.toLowerCase();
      if (text.includes(q)) {
        sec.style.display = '';
        matchCount++;
      } else {
        sec.style.display = 'none';
      }
    });

    // Filter sidebar links
    tocLinks.forEach(l => {
      const text = l.textContent.toLowerCase();
      if (text.includes(q)) {
        l.parentElement.style.display = '';
      } else {
        l.parentElement.style.display = 'none';
      }
    });

    if (countDisplay) {
      countDisplay.textContent = `${matchCount} ফলাফল`;
    }
  }

  if (searchInput) {
    searchInput.addEventListener('input', (e) => handleSearch(e.target.value));
    // Ctrl + K shortcut
    window.addEventListener('keydown', (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault();
        searchInput.focus();
        searchInput.select();
      }
    });
  }

  if (mobileSearch) {
    mobileSearch.addEventListener('input', (e) => handleSearch(e.target.value));
  }
}

// 5. ScrollSpy for Sidebar Active State
function initScrollSpy() {
  const headings = document.querySelectorAll('.chapter-block, .section-container, .author-bio-card, .book-hero-card, .calculators-hub-section');
  const navLinks = document.querySelectorAll('.sidebar .toc-link');

  const observer = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        const id = entry.target.getAttribute('id');
        if (id) {
          navLinks.forEach(link => {
            if (link.getAttribute('data-target') === id) {
              link.classList.add('active');
            } else {
              link.classList.remove('active');
            }
          });
        }
      }
    });
  }, { rootMargin: '-20% 0px -70% 0px' });

  headings.forEach(h => observer.observe(h));
}

// 6. Back to Top Button
function initBackToTop() {
  const btn = document.getElementById('backToTop');
  if (btn) {
    window.addEventListener('scroll', () => {
      if (window.scrollY > 400) {
        btn.classList.add('visible');
      } else {
        btn.classList.remove('visible');
      }
    });

    btn.addEventListener('click', () => {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });
  }
}

// 7. Font Size Adjuster
let currentFontSize = 16;
function toggleFontSize() {
  currentFontSize = currentFontSize >= 19 ? 15 : currentFontSize + 1;
  document.documentElement.style.fontSize = `${currentFontSize}px`;
}

// 8. Interactive Calculators Functions
function switchCalc(calcName) {
  const cards = document.querySelectorAll('.calc-card');
  const tabs = document.querySelectorAll('.calc-tab-btn');

  cards.forEach(c => c.classList.remove('active'));
  tabs.forEach(t => t.classList.remove('active'));

  const activeCard = document.getElementById(`calc_${calcName}`);
  if (activeCard) activeCard.classList.add('active');

  const activeTab = Array.from(tabs).find(t => t.getAttribute('onclick').includes(calcName));
  if (activeTab) activeTab.classList.add('active');
}

// Shannon Calculator
function calculateShannon() {
  const bw = parseFloat(document.getElementById('shannon_bw').value) || 0;
  const snrDb = parseFloat(document.getElementById('shannon_snr').value) || 0;
  const snrLinear = Math.pow(10, snrDb / 10);
  const capacity = bw * (Math.log(1 + snrLinear) / Math.LN2);
  const resElem = document.getElementById('shannon_result').querySelector('.res-value');
  resElem.textContent = `${capacity.toFixed(2)} Mbps`;
}

// FSPL Calculator
function calculateFSPL() {
  const f = parseFloat(document.getElementById('fspl_f').value) || 0;
  const d = parseFloat(document.getElementById('fspl_d').value) || 0;
  if (f <= 0 || d <= 0) return;
  const fspl = 92.45 + (20 * Math.log10(f)) + (20 * Math.log10(d));
  const resElem = document.getElementById('fspl_result').querySelector('.res-value');
  resElem.textContent = `${fspl.toFixed(2)} dB`;
}

// Fresnel Zone Calculator
function calculateFresnel() {
  const d1 = parseFloat(document.getElementById('fresnel_d1').value) || 0;
  const d2 = parseFloat(document.getElementById('fresnel_d2').value) || 0;
  const f = parseFloat(document.getElementById('fresnel_f').value) || 0;
  const totalD = d1 + d2;
  if (f <= 0 || totalD <= 0) return;
  const r1 = 17.32 * Math.sqrt((d1 * d2) / (f * totalD));
  const r60 = 0.60 * r1;
  const resElem = document.getElementById('fresnel_result').querySelector('.res-value');
  resElem.textContent = `${r1.toFixed(2)} মিটার (৬০% ক্লিয়ারেন্স: ${r60.toFixed(2)} মি.)`;
}

// Fiber Reach Calculator
function calculateFiber() {
  const ptx = parseFloat(document.getElementById('fiber_ptx').value) || 0;
  const psens = parseFloat(document.getElementById('fiber_psens').value) || 0;
  const alpha = parseFloat(document.getElementById('fiber_alpha').value) || 0.225;
  const margin = parseFloat(document.getElementById('fiber_margin').value) || 4.0;
  const allowed = (ptx - psens) - margin;
  if (alpha <= 0) return;
  const lmax = allowed / alpha;
  const resElem = document.getElementById('fiber_result').querySelector('.res-value');
  resElem.textContent = `${lmax.toFixed(2)} কিমি`;
}
'''
    with open(js_path, "w", encoding="utf-8") as f:
        f.write(js_content)
    print("docs/app.js created successfully.")


if __name__ == "__main__":
    build_web_ebook()
