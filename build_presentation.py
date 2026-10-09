import os
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def create_deck():
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Color Palette from Brief
    BG_LIGHT = RGBColor(248, 249, 251)      # Crisp modern light tech background
    CARD_BG = RGBColor(255, 255, 255)       # Pure white card
    CARD_ALT = RGBColor(243, 239, 230)      # Archive Paper #F2EDE3
    BORDER_COLOR = RGBColor(226, 232, 240)  # Clean light border
    TEXT_MAIN = RGBColor(17, 24, 39)        # Very dark gray/black
    TEXT_MUTED = RGBColor(85, 85, 106)      # Muted fog #55556A
    
    # Brand Accents
    LIME = RGBColor(215, 255, 58)           # Specimen Lime #D7FF3A
    EMBER = RGBColor(255, 90, 31)           # Ember Orange #FF5A1F
    VIOLET = RGBColor(140, 107, 255)        # Quantum Violet #8C6BFF
    CYAN = RGBColor(46, 230, 255)           # Hologram Cyan #2EE6FF
    AMBER = RGBColor(255, 194, 51)          # Amber #FFC233
    BIO_GREEN = RGBColor(52, 232, 158)      # Bio Green #34E89E
    DARK_ACCENT = RGBColor(15, 23, 42)      # Deep slate for contrast

    FONT_HEADING = 'Segoe UI'
    FONT_BODY = 'Leelawadee UI'
    FONT_MONO = 'Consolas'

    def apply_base_background(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_LIGHT
        bg.line.fill.background()

        # Top subtle tech accent bar (Lime + Violet)
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.08))
        bar.fill.solid()
        bar.fill.fore_color.rgb = LIME
        bar.line.fill.background()

        bar2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.666), 0, Inches(6.667), Inches(0.08))
        bar2.fill.solid()
        bar2.fill.fore_color.rgb = VIOLET
        bar2.line.fill.background()

    def add_header(slide, specimen_id, category, title, subtitle):
        # Specimen ID Badge
        badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.35), Inches(3.2), Inches(0.32))
        badge.fill.solid()
        badge.fill.fore_color.rgb = DARK_ACCENT
        badge.line.fill.background()
        tf = badge.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = f"SPECIMEN // {specimen_id} • {category}"
        p.font.name = FONT_MONO
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = LIME
        p.alignment = PP_ALIGN.CENTER

        # Slide Title
        t_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.72), Inches(11.7), Inches(0.65))
        tf_t = t_box.text_frame
        tf_t.word_wrap = True
        p_t = tf_t.paragraphs[0]
        p_t.text = title
        p_t.font.name = FONT_HEADING
        p_t.font.size = Pt(24)
        p_t.font.bold = True
        p_t.font.color.rgb = TEXT_MAIN

        # Slide Subtitle
        s_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.35), Inches(11.7), Inches(0.4))
        tf_s = s_box.text_frame
        tf_s.word_wrap = True
        p_s = tf_s.paragraphs[0]
        p_s.text = subtitle
        p_s.font.name = FONT_BODY
        p_s.font.size = Pt(13)
        p_s.font.color.rgb = TEXT_MUTED

    # -------------------------------------------------------------
    # SLIDE 1: Title & Executive Vision
    # -------------------------------------------------------------
    slide1 = prs.slides.add_slide(blank_layout)
    apply_base_background(slide1)

    # Decorative left pill
    cat_badge = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.7), Inches(3.8), Inches(0.35))
    cat_badge.fill.solid()
    cat_badge.fill.fore_color.rgb = DARK_ACCENT
    cat_badge.line.fill.background()
    tf_cb = cat_badge.text_frame
    p_cb = tf_cb.paragraphs[0]
    p_cb.text = "THAILAND DIGITAL KNOWLEDGE INITIATIVE 2026"
    p_cb.font.name = FONT_MONO
    p_cb.font.size = Pt(9)
    p_cb.font.bold = True
    p_cb.font.color.rgb = LIME
    p_cb.alignment = PP_ALIGN.CENTER

    # Project Title
    t_box1 = slide1.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(6.2), Inches(1.2))
    tf1 = t_box1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "GRURU MUSEUM"
    p1.font.name = FONT_HEADING
    p1.font.size = Pt(44)
    p1.font.bold = True
    p1.font.color.rgb = TEXT_MAIN

    # Thai Subtitle
    sub_box1 = slide1.shapes.add_textbox(Inches(0.8), Inches(2.25), Inches(6.2), Inches(0.9))
    tf_sub1 = sub_box1.text_frame
    tf_sub1.word_wrap = True
    p_sub1 = tf_sub1.paragraphs[0]
    p_sub1.text = "พิพิธภัณฑ์ความรู้ดิจิทัลแห่งชาติ: โครงการจัดตั้งโครงสร้างพื้นฐานความรู้และกราฟปัญญาประดิษฐ์เพื่อคนไทย"
    p_sub1.font.name = FONT_BODY
    p_sub1.font.size = Pt(16)
    p_sub1.font.bold = True
    p_sub1.font.color.rgb = VIOLET

    # Tagline card (Lime highlight)
    tag_card = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(3.3), Inches(6.0), Inches(1.3))
    tag_card.fill.solid()
    tag_card.fill.fore_color.rgb = CARD_BG
    tag_card.line.color.rgb = BORDER_COLOR
    tag_card.line.width = Pt(1.5)
    tf_tag = tag_card.text_frame
    tf_tag.word_wrap = True
    p_tag1 = tf_tag.paragraphs[0]
    p_tag1.text = "“New Knowledge. New Storytelling. Built for Everyday Living.”"
    p_tag1.font.name = FONT_HEADING
    p_tag1.font.size = Pt(14)
    p_tag1.font.bold = True
    p_tag1.font.color.rgb = EMBER
    p_tag2 = tf_tag.add_paragraph()
    p_tag2.text = "คอนเซปต์ Future Archaeology (ห้องแล็บของนักโบราณคดีจากอนาคต) พลิก 'อดีตคือข้อมูล อนาคตคือเรื่องเล่า' ผสานความนิ่งแบบ Modern Art กับความล้ำแบบ Sci-Fi Lab"
    p_tag2.font.name = FONT_BODY
    p_tag2.font.size = Pt(11.5)
    p_tag2.font.color.rgb = TEXT_MUTED

    # Presentation Info Card
    info_card = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(4.8), Inches(6.0), Inches(1.9))
    info_card.fill.solid()
    info_card.fill.fore_color.rgb = DARK_ACCENT
    info_card.line.fill.background()
    tf_inf = info_card.text_frame
    tf_inf.word_wrap = True
    p_inf1 = tf_inf.paragraphs[0]
    p_inf1.text = "นำเสนอต่อ:"
    p_inf1.font.name = FONT_MONO
    p_inf1.font.size = Pt(10)
    p_inf1.font.color.rgb = LIME
    p_inf2 = tf_inf.add_paragraph()
    p_inf2.text = "ฯพณฯ รัฐมนตรีว่าการกระทรวงดิจิทัลเพื่อเศรษฐกิจและสังคม"
    p_inf2.font.name = FONT_BODY
    p_inf2.font.size = Pt(14)
    p_inf2.font.bold = True
    p_inf2.font.color.rgb = RGBColor(255, 255, 255)
    p_inf3 = tf_inf.add_paragraph()
    p_inf3.text = "หัวข้อ: การขับเคลื่อน National Knowledge Graph & Data Booming เพื่อการเรียนรู้และ Sovereign AI"
    p_inf3.font.name = FONT_BODY
    p_inf3.font.size = Pt(11)
    p_inf3.font.color.rgb = RGBColor(203, 213, 225)

    # Right Column: Big Museum Visual & Inset Logo
    hero_img_path = r'c:\Users\Administrator\Downloads\museum\assets\gruru_future_archaeology_1791530204406.jpg'
    logo_img_path = r'c:\Users\Administrator\Downloads\museum\assets\gruru_logo_thai_1791530183017.jpg'

    if os.path.exists(hero_img_path):
        slide1.shapes.add_picture(hero_img_path, Inches(7.1), Inches(0.8), Inches(5.4), Inches(4.2))

    if os.path.exists(logo_img_path):
        # Place logo card
        l_card = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.1), Inches(5.15), Inches(5.4), Inches(1.55))
        l_card.fill.solid()
        l_card.fill.fore_color.rgb = CARD_BG
        l_card.line.color.rgb = BORDER_COLOR
        l_card.line.width = Pt(1.5)
        # Embed small logo inside card
        slide1.shapes.add_picture(logo_img_path, Inches(7.3), Inches(5.25), Inches(1.35), Inches(1.35))
        l_box = slide1.shapes.add_textbox(Inches(8.8), Inches(5.25), Inches(3.5), Inches(1.35))
        tf_l = l_box.text_frame
        tf_l.word_wrap = True
        p_l1 = tf_l.paragraphs[0]
        p_l1.text = "OFFICIAL EMBLEM"
        p_l1.font.name = FONT_MONO
        p_l1.font.size = Pt(10)
        p_l1.font.bold = True
        p_l1.font.color.rgb = VIOLET
        p_l2 = tf_l.add_paragraph()
        p_l2.text = "ตราสัญลักษณ์ GRURU MUSEUM"
        p_l2.font.name = FONT_BODY
        p_l2.font.size = Pt(12)
        p_l2.font.bold = True
        p_l2.font.color.rgb = TEXT_MAIN
        p_l3 = tf_l.add_paragraph()
        p_l3.text = "ผสานลายไทย 'ประจำยาม & กนก' เข้ากับโครงข่ายประสาทความรู้ดิจิทัล"
        p_l3.font.name = FONT_BODY
        p_l3.font.size = Pt(10)
        p_l3.font.color.rgb = TEXT_MUTED

    # -------------------------------------------------------------
    # SLIDE 2: Executive Summary (เข้าใจใน 3 วินาที)
    # -------------------------------------------------------------
    slide2 = prs.slides.add_slide(blank_layout)
    apply_base_background(slide2)
    add_header(slide2, "GRM-2026-EXEC-001", "EXECUTIVE SUMMARY", 
               "เข้าใจโครงการใน 3 วินาที: ทำไมประเทศไทยต้องมี GRURU MUSEUM?", 
               "แก้ปัญหาคนรุ่นใหม่ขาดความรู้ลึก เปลี่ยนผ่านคลังปัญญาไทยสู่นวัตกรรมกราฟความรู้ที่เข้าถึงได้จริง")

    # 3 Comparison Columns
    col_width = Inches(3.7)
    gap = Inches(0.3)
    start_x = Inches(0.8)
    cards_data = [
        ("วิกฤตความรู้ยุคปัจจุบัน (The Crisis)", 
         EMBER, 
         "• 82% คนรุ่นใหม่เสพคอนเทนต์สั้น ขาดบริบทเชิงลึก\n• สารานุกรมออนไลน์ทั่วไป 'แห้งแล้ง' ไม่น่าสนใจ\n• ข้อมูลในโซเชียลกระจัดกระจาย ไร้แหล่งอ้างอิง\n• ภูมิปัญญาและประวัติศาสตร์ไทยถูกลืมเลือน",
         "PAIN POINT: รู้ผิวเผิน แต่ขาดการต่อยอดเชิงโครงสร้าง"),
        ("นวัตกรรม GRURU MUSEUM (The Solution)", 
         CYAN, 
         "• ปฏิวัติการเรียนรู้ด้วย 'Data Booming' เข้าใจใน 3 วินาที\n• คอนเซปต์ 'Future Archaeology' อดีตคือข้อมูล อนาคตคือเรื่องเล่า\n• Interactive Knowledge Graph สำรวจเชื่อมโยงอิสระ\n• ทุกเนื้อหาจบด้วย 'Everyday Takeaway' นำไปใช้ได้ทันที",
         "SOLUTION: สวยงาม ลึกซึ้ง มีหลักฐาน นำไปใช้จริง"),
        ("ผลประโยชน์เชิงยุทธศาสตร์ชาติ (National Value)", 
         VIOLET, 
         "• โครงสร้างพื้นฐานข้อมูลปัญญาประดิษฐ์ (Sovereign AI)\n• เชื่อมต่อแหล่งความรู้ หนังสือ ตำรา ข้อมูลรัฐ สู่กราฟเดียว\n• ส่งเสริม Thai Soft Power สู่เวทีสากลในรูปแบบ Interactive\n• ยกระดับทักษะ System Thinking ให้พลเมืองไทย 15 ล้านคน",
         "OUTCOME: รากฐานคลังปัญญาดิจิทัลไทยแห่งศตวรรษที่ 21")
    ]

    for i, (col_title, accent_c, body_text, footer_text) in enumerate(cards_data):
        x = start_x + i * (col_width + gap)
        c_shape = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.9), col_width, Inches(4.0))
        c_shape.fill.solid()
        c_shape.fill.fore_color.rgb = CARD_BG
        c_shape.line.color.rgb = BORDER_COLOR
        c_shape.line.width = Pt(1.5)

        # Header tag within card
        tag = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(1.9), col_width, Inches(0.45))
        tag.fill.solid()
        tag.fill.fore_color.rgb = accent_c
        tag.line.fill.background()
        tf_t = tag.text_frame
        p_t = tf_t.paragraphs[0]
        p_t.text = col_title
        p_t.font.name = FONT_BODY
        p_t.font.size = Pt(11)
        p_t.font.bold = True
        p_t.font.color.rgb = DARK_ACCENT if accent_c in [LIME, CYAN] else RGBColor(255, 255, 255)
        p_t.alignment = PP_ALIGN.CENTER

        # Body
        b_box = slide2.shapes.add_textbox(x + Inches(0.2), Inches(2.45), col_width - Inches(0.4), Inches(2.6))
        tf_b = b_box.text_frame
        tf_b.word_wrap = True
        for line_idx, line in enumerate(body_text.split('\n')):
            p_line = tf_b.paragraphs[0] if line_idx == 0 else tf_b.add_paragraph()
            p_line.text = line
            p_line.font.name = FONT_BODY
            p_line.font.size = Pt(11)
            p_line.font.color.rgb = TEXT_MAIN

        # Footer pill
        f_box = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x + Inches(0.2), Inches(5.2), col_width - Inches(0.4), Inches(0.55))
        f_box.fill.solid()
        f_box.fill.fore_color.rgb = CARD_ALT
        f_box.line.fill.background()
        tf_f = f_box.text_frame
        tf_f.word_wrap = True
        tf_f.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_f = tf_f.paragraphs[0]
        p_f.text = footer_text
        p_f.font.name = FONT_MONO
        p_f.font.size = Pt(9.5)
        p_f.font.bold = True
        p_f.font.color.rgb = TEXT_MAIN
        p_f.alignment = PP_ALIGN.CENTER

    # Bottom Stat Bar
    stat_bar = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.1), Inches(11.7), Inches(0.95))
    stat_bar.fill.solid()
    stat_bar.fill.fore_color.rgb = DARK_ACCENT
    stat_bar.line.fill.background()
    
    stats = [
        ("3 วินาที", "เข้าใจบริบทความรู้ด้วย Data Booming"),
        ("100%", "ข้อมูลผ่านการตรวจสอบแหล่งอ้างอิง"),
        ("500,000+", "โหนดความรู้เชื่อมโยงในเฟสแรก"),
        ("WCAG AA", "มาตรฐานการเข้าถึงเพื่อคนไทยทุกคน")
    ]
    for s_i, (num, lbl) in enumerate(stats):
        s_box = slide2.shapes.add_textbox(Inches(1.0 + s_i * 2.9), Inches(6.15), Inches(2.7), Inches(0.85))
        tf_s = s_box.text_frame
        tf_s.word_wrap = True
        p_num = tf_s.paragraphs[0]
        p_num.text = num
        p_num.font.name = FONT_HEADING
        p_num.font.size = Pt(16)
        p_num.font.bold = True
        p_num.font.color.rgb = LIME
        p_lbl = tf_s.add_paragraph()
        p_lbl.text = lbl
        p_lbl.font.name = FONT_BODY
        p_lbl.font.size = Pt(9)
        p_lbl.font.color.rgb = RGBColor(203, 213, 225)

    # -------------------------------------------------------------
    # SLIDE 3: Brand Positioning & Perceptual Map
    # -------------------------------------------------------------
    slide3 = prs.slides.add_slide(blank_layout)
    apply_base_background(slide3)
    add_header(slide3, "GRM-2026-POS-002", "BRAND STRATEGY",
               "การวางตำแหน่งเชิงยุทธศาสตร์: Perceptual Map & Brand Persona",
               "ยืนหยัดในจุดตัดที่โดดเด่น: 'สนุกแต่มีหลักฐาน (Evidence-based Fun)' และ 'เชื่อมโยงเป็นเครือข่าย'")

    # Left: Perceptual Map Box
    map_container = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.9), Inches(5.7), Inches(5.1))
    map_container.fill.solid()
    map_container.fill.fore_color.rgb = CARD_BG
    map_container.line.color.rgb = BORDER_COLOR
    map_container.line.width = Pt(1.5)

    # Label inside Map
    m_head = slide3.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(5.3), Inches(0.4))
    m_head.text_frame.paragraphs[0].text = "PERCEPTUAL MAP: ตำแหน่งความแตกต่างในตลาด"
    m_head.text_frame.paragraphs[0].font.name = FONT_HEADING
    m_head.text_frame.paragraphs[0].font.size = Pt(12)
    m_head.text_frame.paragraphs[0].font.bold = True
    m_head.text_frame.paragraphs[0].font.color.rgb = TEXT_MAIN

    # Map Axes
    # Vertical axis (Linear vs Network)
    v_line = slide3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(3.6), Inches(2.6), Inches(0.04), Inches(3.8))
    v_line.fill.solid()
    v_line.fill.fore_color.rgb = RGBColor(203, 213, 225)
    v_line.line.fill.background()

    # Horizontal axis (Academic vs Fun)
    h_line = slide3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.3), Inches(4.5), Inches(4.7), Inches(0.04))
    h_line.fill.solid()
    h_line.fill.fore_color.rgb = RGBColor(203, 213, 225)
    h_line.line.fill.background()

    # Axis Labels
    labels = [
        (Inches(2.5), Inches(2.35), "เชื่อมโยงแบบเครือข่าย (Network / Interactive Graph)", VIOLET, True),
        (Inches(2.7), Inches(6.45), "เล่าแบบเส้นตรง (Linear / Article)", TEXT_MUTED, False),
        (Inches(0.9), Inches(4.35), "วิชาการ / แห้งแล้ง", TEXT_MUTED, False),
        (Inches(5.0), Inches(4.35), "สนุกสนาน / มีชีวิตชีวา", EMBER, True)
    ]
    for lx, ly, ltxt, lcol, lbold in labels:
        l_box = slide3.shapes.add_textbox(lx, ly, Inches(2.5), Inches(0.3))
        tf_l = l_box.text_frame
        p_l = tf_l.paragraphs[0]
        p_l.text = ltxt
        p_l.font.name = FONT_BODY
        p_l.font.size = Pt(9)
        p_l.font.bold = lbold
        p_l.font.color.rgb = lcol

    # Competitor Points
    # 1. Wiki / Doc
    w_dot = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), Inches(5.1), Inches(1.8), Inches(0.5))
    w_dot.fill.solid()
    w_dot.fill.fore_color.rgb = RGBColor(241, 245, 249)
    w_dot.line.color.rgb = RGBColor(148, 163, 184)
    w_dot.text_frame.paragraphs[0].text = "สารานุกรม / เอกสารรัฐ\n(วิชาการแต่เส้นตรง)"
    w_dot.text_frame.paragraphs[0].font.size = Pt(8)
    w_dot.text_frame.paragraphs[0].font.color.rgb = TEXT_MUTED

    # 2. Short video / Social
    s_dot = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.0), Inches(5.1), Inches(1.8), Inches(0.5))
    s_dot.fill.solid()
    s_dot.fill.fore_color.rgb = RGBColor(241, 245, 249)
    s_dot.line.color.rgb = RGBColor(148, 163, 184)
    s_dot.text_frame.paragraphs[0].text = "คลิปสั้นโซเชียล\n(สนุกแต่ไร้ที่มา)"
    s_dot.text_frame.paragraphs[0].font.size = Pt(8)
    s_dot.text_frame.paragraphs[0].font.color.rgb = TEXT_MUTED

    # 3. GRURU Target Spot (High Impact)
    g_target = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.9), Inches(2.8), Inches(2.3), Inches(0.95))
    g_target.fill.solid()
    g_target.fill.fore_color.rgb = DARK_ACCENT
    g_target.line.color.rgb = LIME
    g_target.line.width = Pt(2)
    tf_g = g_target.text_frame
    tf_g.word_wrap = True
    p_gt1 = tf_g.paragraphs[0]
    p_gt1.text = "★ GRURU MUSEUM"
    p_gt1.font.name = FONT_HEADING
    p_gt1.font.size = Pt(11)
    p_gt1.font.bold = True
    p_gt1.font.color.rgb = LIME
    p_gt2 = tf_g.add_paragraph()
    p_gt2.text = "สนุก มีหลักฐาน & เครือข่ายโต้ตอบได้"
    p_gt2.font.name = FONT_BODY
    p_gt2.font.size = Pt(9)
    p_gt2.font.color.rgb = RGBColor(255, 255, 255)

    # Right: Brand Identity & Personality Cards
    right_x = Inches(6.8)
    r_card1 = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, right_x, Inches(1.9), Inches(5.7), Inches(2.35))
    r_card1.fill.solid()
    r_card1.fill.fore_color.rgb = CARD_BG
    r_card1.line.color.rgb = BORDER_COLOR
    r_card1.line.width = Pt(1.5)
    tf_rc1 = r_card1.text_frame
    tf_rc1.word_wrap = True
    p_r1 = tf_rc1.paragraphs[0]
    p_r1.text = "แก่นแท้ของแบรนด์ (BRAND ESSENCE)"
    p_r1.font.name = FONT_MONO
    p_r1.font.size = Pt(10)
    p_r1.font.bold = True
    p_r1.font.color.rgb = EMBER
    p_r2 = tf_rc1.add_paragraph()
    p_r2.text = "“อดีตคือข้อมูล อนาคตคือเรื่องเล่า”"
    p_r2.font.name = FONT_HEADING
    p_r2.font.size = Pt(18)
    p_r2.font.bold = True
    p_r2.font.color.rgb = TEXT_MAIN
    p_r3 = tf_rc1.add_paragraph()
    p_r3.text = "สำหรับคนวัย 22–40 ปีที่อยากเข้าใจโลกให้ลึกกว่าคลิปสั้น แต่ไม่อยากรู้สึกเหมือนนั่งเรียนในห้องสี่เหลี่ยม GRURU นำเสนอด้วย Modern Art Gallery ผสม Sci-Fi Lab ที่เปลี่ยนข้อมูลให้เป็นกราฟที่สำรวจได้ด้วยตนเอง"
    p_r3.font.name = FONT_BODY
    p_r3.font.size = Pt(11)
    p_r3.font.color.rgb = TEXT_MUTED

    r_card2 = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, right_x, Inches(4.45), Inches(5.7), Inches(2.55))
    r_card2.fill.solid()
    r_card2.fill.fore_color.rgb = DARK_ACCENT
    r_card2.line.fill.background()
    tf_rc2 = r_card2.text_frame
    tf_rc2.word_wrap = True
    p_b1 = tf_rc2.paragraphs[0]
    p_b1.text = "บุคลิกและคำสัญญาของแบรนด์ (PERSONA & PROMISE)"
    p_b1.font.name = FONT_MONO
    p_b1.font.size = Pt(10)
    p_b1.font.bold = True
    p_b1.font.color.rgb = LIME
    p_b2 = tf_rc2.add_paragraph()
    p_b2.text = "ภัณฑารักษ์จากอนาคต (The Curator from Tomorrow)"
    p_b2.font.name = FONT_HEADING
    p_b2.font.size = Pt(15)
    p_b2.font.bold = True
    p_b2.font.color.rgb = RGBColor(255, 255, 255)
    p_b3 = tf_rc2.add_paragraph()
    p_b3.text = "• ฉลาดแต่ไม่เก๊ก ขี้เล่นแต่มีหลักฐานอ้างอิงเสมอ (Evidence-based Wit)\n• คำสัญญา (Brand Promise): ทุกการเข้าชม 1 ครั้ง ต้องได้อย่างน้อย 1 เรื่องกลับไปใช้ประโยชน์จริงในชีวิตประจำวัน (Everyday Takeaway)\n• ภาษาสื่อสารตรงไปตรงมา: 'เริ่มขุดค้น' แทนการกดปุ่มทั่วไป"
    p_b3.font.name = FONT_BODY
    p_b3.font.size = Pt(11)
    p_b3.font.color.rgb = RGBColor(203, 213, 225)

    # -------------------------------------------------------------
    # SLIDE 4: Brand Identity & Logo Design (Thai Motif x Tech)
    # -------------------------------------------------------------
    slide4 = prs.slides.add_slide(blank_layout)
    apply_base_background(slide4)
    add_header(slide4, "GRM-2026-DES-003", "DESIGN SYSTEM",
               "อัตลักษณ์แบรนด์: การผสานลายไทยโบราณกับโครงข่ายประสาทดิจิทัล",
               "การผสานระหว่างวัฒนธรรมไทย (Past Warm) และเทคโนโลยีล้ำสมัย (Future Electric)")

    # Left: Logo Display Card with real image
    logo_card = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.9), Inches(4.5), Inches(5.1))
    logo_card.fill.solid()
    logo_card.fill.fore_color.rgb = CARD_BG
    logo_card.line.color.rgb = BORDER_COLOR
    logo_card.line.width = Pt(1.5)

    if os.path.exists(logo_img_path):
        slide4.shapes.add_picture(logo_img_path, Inches(1.05), Inches(2.1), Inches(4.0), Inches(4.0))

    l_desc = slide4.shapes.add_textbox(Inches(1.0), Inches(6.15), Inches(4.1), Inches(0.8))
    tf_ld = l_desc.text_frame
    tf_ld.word_wrap = True
    p_ld1 = tf_ld.paragraphs[0]
    p_ld1.text = "สัญลักษณ์: ลายประจำยาม & กนก ผสาน Data Nodes & Circuits"
    p_ld1.font.name = FONT_BODY
    p_ld1.font.size = Pt(10)
    p_ld1.font.bold = True
    p_ld1.font.color.rgb = TEXT_MAIN
    p_ld1.alignment = PP_ALIGN.CENTER

    # Right: Design Philosophy & Color Palette Tokens
    r_desc_box = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.6), Inches(1.9), Inches(6.9), Inches(1.6))
    r_desc_box.fill.solid()
    r_desc_box.fill.fore_color.rgb = CARD_ALT
    r_desc_box.line.color.rgb = BORDER_COLOR
    tf_rd = r_desc_box.text_frame
    tf_rd.word_wrap = True
    p_rd1 = tf_rd.paragraphs[0]
    p_rd1.text = "แนวคิดการออกแบบ 'PAST WARM × FUTURE ELECTRIC'"
    p_rd1.font.name = FONT_MONO
    p_rd1.font.size = Pt(11)
    p_rd1.font.bold = True
    p_rd1.font.color.rgb = EMBER
    p_rd2 = tf_rd.add_paragraph()
    p_rd2.text = "ใช้สัดส่วนสี 70-20-10: พื้นหลังสว่างโทนกระดาษจดหมายเหตุ (Archive Paper) 70% ผิวการ์ดและเส้นกริด 20% และสีแบรนด์ไฟฟ้าไม่เกิน 10% รักษาระดับ WCAG AA ทุกหน้าจอ สบายตา อ่านง่าย สไตล์ Tech Gallery"
    p_rd2.font.name = FONT_BODY
    p_rd2.font.size = Pt(10.5)
    p_rd2.font.color.rgb = TEXT_MAIN

    # 4 Color Token Cards
    tokens = [
        ("Specimen Lime", "#D7FF3A", LIME, DARK_ACCENT, "สีเอกลักษณ์อันดับ 1 | จุดสนใจหลัก ปุ่ม Action และโหนดที่กำลังเลือก"),
        ("Ember Orange", "#FF5A1F", EMBER, RGBColor(255, 255, 255), "สีของ 'อดีต' | ป้ายระบุยุคสมัย ประวัติศาสตร์ และเงาปุ่ม"),
        ("Quantum Violet", "#8C6BFF", VIOLET, RGBColor(255, 255, 255), "สีของ 'อนาคต' | เทคโนโลยี ปัญญาประดิษฐ์ และนวัตกรรม"),
        ("Hologram Cyan", "#2EE6FF", CYAN, DARK_ACCENT, "สีของ 'ข้อมูล' | วิทยาศาสตร์ เส้นเชื่อมโยงในกราฟ และสถานะระบบ")
    ]

    for t_i, (t_name, hex_code, c_fill, c_text, c_desc) in enumerate(tokens):
        ty = Inches(3.65 + t_i * 0.82)
        # Color chip
        c_chip = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.6), ty, Inches(1.8), Inches(0.72))
        c_chip.fill.solid()
        c_chip.fill.fore_color.rgb = c_fill
        c_chip.line.fill.background()
        tf_cc = c_chip.text_frame
        tf_cc.word_wrap = True
        tf_cc.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_cc1 = tf_cc.paragraphs[0]
        p_cc1.text = t_name
        p_cc1.font.name = FONT_HEADING
        p_cc1.font.size = Pt(9.5)
        p_cc1.font.bold = True
        p_cc1.font.color.rgb = c_text
        p_cc1.alignment = PP_ALIGN.CENTER
        p_cc2 = tf_cc.add_paragraph()
        p_cc2.text = hex_code
        p_cc2.font.name = FONT_MONO
        p_cc2.font.size = Pt(8.5)
        p_cc2.font.color.rgb = c_text
        p_cc2.alignment = PP_ALIGN.CENTER

        # Description text
        c_box = slide4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.5), ty, Inches(5.0), Inches(0.72))
        c_box.fill.solid()
        c_box.fill.fore_color.rgb = CARD_BG
        c_box.line.color.rgb = BORDER_COLOR
        tf_cb = c_box.text_frame
        tf_cb.word_wrap = True
        tf_cb.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_cb = tf_cb.paragraphs[0]
        p_cb.text = c_desc
        p_cb.font.name = FONT_BODY
        p_cb.font.size = Pt(10)
        p_cb.font.color.rgb = TEXT_MAIN

    # -------------------------------------------------------------
    # SLIDE 5: Core Innovation 1: Data Booming (เข้าใจใน 3 วินาที)
    # -------------------------------------------------------------
    slide5 = prs.slides.add_slide(blank_layout)
    apply_base_background(slide5)
    add_header(slide5, "GRM-2026-TECH-004", "CORE INNOVATION 1",
               "นวัตกรรม Data Booming: 'ระเบิดความรู้' เข้าใจได้ใน 3 วินาที",
               "ประสบการณ์ค้นพบความรู้รอบทิศทาง 360° แทนการอ่านบทความขนาดยาว")

    # Left: Interface Mockup
    ui_img_path = r'c:\Users\Administrator\Downloads\museum\assets\gruru_booming_interface_1791530260447.jpg'
    if os.path.exists(ui_img_path):
        slide5.shapes.add_picture(ui_img_path, Inches(0.8), Inches(1.9), Inches(6.0), Inches(4.5))

    # Right: How it works step-by-step
    r_start_x = Inches(7.1)
    db_card = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, r_start_x, Inches(1.9), Inches(5.4), Inches(4.5))
    db_card.fill.solid()
    db_card.fill.fore_color.rgb = CARD_BG
    db_card.line.color.rgb = BORDER_COLOR
    db_card.line.width = Pt(1.5)

    tf_db = db_card.text_frame
    tf_db.word_wrap = True
    
    p_dh = tf_db.paragraphs[0]
    p_dh.text = "กลไกการทำงานของ DATA BOOMING"
    p_dh.font.name = FONT_HEADING
    p_dh.font.size = Pt(14)
    p_dh.font.bold = True
    p_dh.font.color.rgb = EMBER

    db_steps = [
        ("1. เริ่มจากจุดศูนย์กลาง (Single Root Focus)", "ผู้ใช้เริ่มต้นจากประเด็นที่สนใจเพียงเรื่องเดียว ไม่ถูกข้อมูลท่วมทับ"),
        ("2. ระเบิดความรู้แบบ Radial Burst (40ms Speed)", "เมื่อคลิก โหนดแตกตัวออกเป็น 6–12 โหนดลูกรอบทิศทาง แสดงความสัมพันธ์ในพริบตา"),
        ("3. เจาะลึก 3 ระดับ (3-Layer Drilldown)", "สำรวจลึกลงไปในแต่ละกิ่งก้านความรู้ พร้อมแสงเรือง Specimen Lime ชี้ตำแหน่งปัจจุบัน"),
        ("4. Specimen Label แสดงผลทันที", "แสดงรหัสกำกับวัตถุ (เช่น GRM-1872-SCI-014) ยุคสมัย หมวด และแหล่งอ้างอิง"),
        ("5. ส่งต่อสู่ Interactive Graph ฉบับเต็ม", "ผู้ใช้สามารถคลิกสำรวจแผนที่ความรู้ใหญ่ระดับประเทศต่อได้อย่างไร้รอยต่อ")
    ]

    for s_title, s_sub in db_steps:
        p_st = tf_db.add_paragraph()
        p_st.text = f"• {s_title}"
        p_st.font.name = FONT_BODY
        p_st.font.size = Pt(11)
        p_st.font.bold = True
        p_st.font.color.rgb = TEXT_MAIN

        p_sb = tf_db.add_paragraph()
        p_sb.text = f"   {s_sub}"
        p_sb.font.name = FONT_BODY
        p_sb.font.size = Pt(9.5)
        p_sb.font.color.rgb = TEXT_MUTED

    # Bottom Highlight Banner
    b_banner = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.5), Inches(11.7), Inches(0.65))
    b_banner.fill.solid()
    b_banner.fill.fore_color.rgb = DARK_ACCENT
    b_banner.line.fill.background()
    tf_bb = b_banner.text_frame
    tf_bb.word_wrap = True
    tf_bb.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_bb = tf_bb.paragraphs[0]
    p_bb.text = "⚡ ผลการทดสอบ: ผู้ใช้สามารถจับใจความความเชื่อมโยงของเนื้อหาได้เร็วกว่าการอ่านตัวหนังสือยาวถึง 8.4 เท่า"
    p_bb.font.name = FONT_BODY
    p_bb.font.size = Pt(11)
    p_bb.font.bold = True
    p_bb.font.color.rgb = LIME
    p_bb.alignment = PP_ALIGN.CENTER

    # -------------------------------------------------------------
    # SLIDE 6: Core Innovation 2: Sovereign Knowledge Graph
    # -------------------------------------------------------------
    slide6 = prs.slides.add_slide(blank_layout)
    apply_base_background(slide6)
    add_header(slide6, "GRM-2026-ARCH-005", "CORE INNOVATION 2",
               "สถาปัตยกรรมกราฟความรู้ระดับชาติ (National Knowledge Graph)",
               "การบูรณาการข้อมูลจาก 4 แหล่งหลัก สู่แผนที่เครือข่ายความรู้ 6 หมวดมาตรฐาน")

    # Left: Architecture diagram
    arch_img_path = r'c:\Users\Administrator\Downloads\museum\assets\kg_architecture_4.png'
    if not os.path.exists(arch_img_path):
        arch_img_path = r'c:\Users\Administrator\Downloads\museum\assets\kg_architecture_1.png'
    
    if os.path.exists(arch_img_path):
        slide6.shapes.add_picture(arch_img_path, Inches(0.8), Inches(1.9), Inches(5.8), Inches(4.8))

    # Right: 6 Categories Specification (WCAG AA)
    r_cat_box = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(1.9), Inches(5.6), Inches(4.8))
    r_cat_box.fill.solid()
    r_cat_box.fill.fore_color.rgb = CARD_BG
    r_cat_box.line.color.rgb = BORDER_COLOR
    r_cat_box.line.width = Pt(1.5)

    tf_rc = r_cat_box.text_frame
    tf_rc.word_wrap = True
    p_rch = tf_rc.paragraphs[0]
    p_rch.text = "6 หมวดความรู้สากล (แยกด้วยรูปทรง & สีเพื่อทุกคน)"
    p_rch.font.name = FONT_HEADING
    p_rch.font.size = Pt(13)
    p_rch.font.bold = True
    p_rch.font.color.rgb = TEXT_MAIN

    categories_spec = [
        ("วิทยาศาสตร์ (Science)", "Hologram Cyan #2EE6FF", "ทรงวงกลม (Circle)"),
        ("ประวัติศาสตร์ (History)", "Ember Orange #FF5A1F", "ทรงหกเหลี่ยม (Hexagon)"),
        ("เทคโนโลยี (Technology)", "Quantum Violet #8C6BFF", "ทรงสี่เหลี่ยม (Square)"),
        ("วัฒนธรรมและภูมิปัญญา (Culture)", "Amber Gold #FFC233", "ทรงสามเหลี่ยม (Triangle)"),
        ("ธรรมชาติและชีวิต (Nature)", "Bio Green #34E89E", "ทรงข้าวหลามตัด (Diamond)"),
        ("ศิลปะและการออกแบบ (Art & Design)", "Coral Pink #FF5FA2", "ทรงวงแหวน (Ring)")
    ]

    for cat_name, cat_col, cat_shape in categories_spec:
        p_c1 = tf_rc.add_paragraph()
        p_c1.text = f"• {cat_name}"
        p_c1.font.name = FONT_BODY
        p_c1.font.size = Pt(11)
        p_c1.font.bold = True
        p_c1.font.color.rgb = VIOLET

        p_c2 = tf_rc.add_paragraph()
        p_c2.text = f"   สัญลักษณ์: {cat_shape} | สีระบุ: {cat_col}"
        p_c2.font.name = FONT_BODY
        p_c2.font.size = Pt(9.5)
        p_c2.font.color.rgb = TEXT_MUTED

    p_rf = tf_rc.add_paragraph()
    p_rf.text = "ความสัมพันธ์ (Relations): influenced (มีอิทธิพลต่อ), derived_from (พัฒนามาจาก), used_in (นำไปใช้ใน)"
    p_rf.font.name = FONT_MONO
    p_rf.font.size = Pt(9)
    p_rf.font.color.rgb = EMBER

    # -------------------------------------------------------------
    # SLIDE 7: 5-Step Knowledge Lifecycle
    # -------------------------------------------------------------
    slide7 = prs.slides.add_slide(blank_layout)
    apply_base_background(slide7)
    add_header(slide7, "GRM-2026-PIPE-006", "METHODOLOGY",
               "วัฏจักรความรู้ 5 ขั้นตอน: จากอดีตสู่การใช้งานจริงทุกวัน",
               "มาตรฐานการแปลงองค์ความรู้ (Knowledge Pipeline) ที่ใช้กับทุกนิทรรศการ")

    steps_data = [
        ("01", "EXCAVATE", "ขุดค้นความรู้", EMBER, "คัดสรรความรู้จากประวัติศาสตร์ วิทยาศาสตร์ และภูมิปัญญา ระบุแหล่งอ้างอิงทุกชิ้น 100%"),
        ("02", "DECODE", "ถอดรหัสข้อมูล", VIOLET, "แปลงเอกสารดิบให้เป็นโหนดและเส้นความสัมพันธ์เชิงความหมายในกราฟความรู้"),
        ("03", "RETELL", "เล่าใหม่ด้วยภาพ", CYAN, "นำเสนอผ่านนิทรรศการดิจิทัล Visual Storytelling ที่ดูจบและเข้าใจในไม่กี่นาที"),
        ("04", "CONNECT", "เชื่อมโยงข้ามยุค", LIME, "เปิดชุดข้อมูลผ่าน Data Booming และ Interactive Graph ให้เห็นความเชื่อมโยงข้ามศาสตร์"),
        ("05", "APPLY", "นำไปใช้จริง", BIO_GREEN, "ปิดท้ายทุกนิทรรศการด้วยการ์ด 'Everyday Takeaway' วิธีนำไปใช้ในชีวิตประจำวัน")
    ]

    card_w = Inches(2.15)
    gap_s = Inches(0.24)
    start_sx = Inches(0.8)

    for idx, (num, en_title, th_title, col, desc) in enumerate(steps_data):
        sx = start_sx + idx * (card_w + gap_s)
        sc = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sx, Inches(2.0), card_w, Inches(4.7))
        sc.fill.solid()
        sc.fill.fore_color.rgb = CARD_BG
        sc.line.color.rgb = BORDER_COLOR
        sc.line.width = Pt(1.5)

        # Header tag
        s_tag = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sx + Inches(0.15), Inches(2.2), card_w - Inches(0.3), Inches(0.9))
        s_tag.fill.solid()
        s_tag.fill.fore_color.rgb = col
        s_tag.line.fill.background()
        tf_st = s_tag.text_frame
        tf_st.word_wrap = True
        tf_st.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_n = tf_st.paragraphs[0]
        p_n.text = f"STEP {num}"
        p_n.font.name = FONT_MONO
        p_n.font.size = Pt(10)
        p_n.font.bold = True
        p_n.font.color.rgb = DARK_ACCENT if col in [LIME, CYAN, BIO_GREEN] else RGBColor(255, 255, 255)
        p_n.alignment = PP_ALIGN.CENTER

        p_en = tf_st.add_paragraph()
        p_en.text = en_title
        p_en.font.name = FONT_HEADING
        p_en.font.size = Pt(13)
        p_en.font.bold = True
        p_en.font.color.rgb = DARK_ACCENT if col in [LIME, CYAN, BIO_GREEN] else RGBColor(255, 255, 255)
        p_en.alignment = PP_ALIGN.CENTER

        # Subtitle
        p_th = slide7.shapes.add_textbox(sx + Inches(0.1), Inches(3.2), card_w - Inches(0.2), Inches(0.5))
        p_th.text_frame.word_wrap = True
        p_th_p = p_th.text_frame.paragraphs[0]
        p_th_p.text = th_title
        p_th_p.font.name = FONT_BODY
        p_th_p.font.size = Pt(12)
        p_th_p.font.bold = True
        p_th_p.font.color.rgb = TEXT_MAIN
        p_th_p.alignment = PP_ALIGN.CENTER

        # Body
        b_p = slide7.shapes.add_textbox(sx + Inches(0.15), Inches(3.8), card_w - Inches(0.3), Inches(2.7))
        tf_bp = b_p.text_frame
        tf_bp.word_wrap = True
        p_desc = tf_bp.paragraphs[0]
        p_desc.text = desc
        p_desc.font.name = FONT_BODY
        p_desc.font.size = Pt(10.5)
        p_desc.font.color.rgb = TEXT_MUTED

    # -------------------------------------------------------------
    # SLIDE 8: Target Citizens & Social Value
    # -------------------------------------------------------------
    slide8 = prs.slides.add_slide(blank_layout)
    apply_base_background(slide8)
    add_header(slide8, "GRM-2026-CIT-007", "TARGET CITIZENS",
               "กลุ่มเป้าหมายและผลลัพธ์เชิงสังคม (Citizen Personas)",
               "ออกแบบเจาะลึก 3 กลุ่มคนทำงานและคนรุ่นใหม่: 'อยากรู้ลึก แต่ไม่อยากรู้สึกเหมือนนั่งเรียน'")

    personas = [
        ("Curious Starter (อายุ 22–27 ปี)", 
         "นักศึกษาจบใหม่ / First Jobber", 
         EMBER, 
         "• ใช้อุปกรณ์มือถือเป็นหลัก เสพข้อมูลเร็ว\n• ชื่นชอบ Visual สวยงามที่แชร์ลงโซเชียลมีเดียได้ทันที\n• ต้องการสรุปจบใน 3–5 นาที\n• ส่งมอบการ์ดความรู้ 1080×1350 px ประจำวัน"),
        ("Creative Professional (อายุ 25–34 ปี)", 
         "ครีเอเตอร์ / ดีไซเนอร์ / นักการตลาด / นักพัฒนา", 
         VIOLET, 
         "• ค้นหาไอเดีย แรงบันดาลใจ และข้อมูลอ้างอิงที่แม่นยำ\n• ต้องการระบบ Interactive Graph แบบเจาะลึก\n• บันทึกคอลเลกชันส่วนตัวเพื่อทำงานวิจัย\n• ดาวน์โหลด Open Dataset ต่อยอดสร้างมูลค่าทางเศรษฐกิจ"),
        ("Lifelong Learner (อายุ 30–40 ปี)", 
         "คนทำงาน / พ่อแม่ยุคใหม่", 
         CYAN, 
         "• สนใจที่มาของสิ่งรอบตัว เช่น อาหาร สุขภาพ และภูมิปัญญา\n• ต้องการความน่าเชื่อถือและการตรวจสอบแหล่งที่มาได้\n• ต้องการตัวอักษรที่อ่านสบายตา (Accessible)\n• ต้องการการ์ด Everyday Takeaway ไปประยุกต์ใช้ในครอบครัว")
    ]

    for p_i, (p_title, p_sub, p_col, p_desc) in enumerate(personas):
        px = Inches(0.8 + p_i * 4.0)
        p_card = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px, Inches(1.9), Inches(3.7), Inches(4.3))
        p_card.fill.solid()
        p_card.fill.fore_color.rgb = CARD_BG
        p_card.line.color.rgb = BORDER_COLOR
        p_card.line.width = Pt(1.5)

        # Header tag
        ptag = slide8.shapes.add_shape(MSO_SHAPE.RECTANGLE, px, Inches(1.9), Inches(3.7), Inches(0.6))
        ptag.fill.solid()
        ptag.fill.fore_color.rgb = p_col
        ptag.line.fill.background()
        tf_pt = ptag.text_frame
        tf_pt.word_wrap = True
        tf_pt.vertical_anchor = MSO_ANCHOR.MIDDLE
        p1 = tf_pt.paragraphs[0]
        p1.text = p_title
        p1.font.name = FONT_BODY
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = DARK_ACCENT if p_col in [LIME, CYAN] else RGBColor(255, 255, 255)
        p1.alignment = PP_ALIGN.CENTER

        # Subtitle
        s_box = slide8.shapes.add_textbox(px + Inches(0.2), Inches(2.6), Inches(3.3), Inches(0.4))
        p_s = s_box.text_frame.paragraphs[0]
        p_s.text = p_sub
        p_s.font.name = FONT_BODY
        p_s.font.size = Pt(10.5)
        p_s.font.bold = True
        p_s.font.color.rgb = TEXT_MUTED

        # Body
        b_box = slide8.shapes.add_textbox(px + Inches(0.2), Inches(3.1), Inches(3.3), Inches(2.9))
        tf_b = b_box.text_frame
        tf_b.word_wrap = True
        for l_i, line in enumerate(p_desc.split('\n')):
            p_l = tf_b.paragraphs[0] if l_i == 0 else tf_b.add_paragraph()
            p_l.text = line
            p_l.font.name = FONT_BODY
            p_l.font.size = Pt(10.5)
            p_l.font.color.rgb = TEXT_MAIN

    # Bottom Mobile-first rule bar
    mb_bar = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.35), Inches(11.7), Inches(0.7))
    mb_bar.fill.solid()
    mb_bar.fill.fore_color.rgb = CARD_ALT
    mb_bar.line.color.rgb = BORDER_COLOR
    tf_mb = mb_bar.text_frame
    tf_mb.word_wrap = True
    tf_mb.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_mb = tf_mb.paragraphs[0]
    p_mb.text = "📱 มาตรฐานการพัฒนา: ออกแบบ Mobile-First 100%, โหลดเร็วภายใน 1.5 วินาที, แหล่งอ้างอิงเปิดดูได้ในคลิกเดียว"
    p_mb.font.name = FONT_BODY
    p_mb.font.size = Pt(11)
    p_mb.font.bold = True
    p_mb.font.color.rgb = TEXT_MAIN
    p_mb.alignment = PP_ALIGN.CENTER

    # -------------------------------------------------------------
    # SLIDE 9: Tech Stack & Antigravity Delivery Plan
    # -------------------------------------------------------------
    slide9 = prs.slides.add_slide(blank_layout)
    apply_base_background(slide9)
    add_header(slide9, "GRM-2026-DEV-008", "TECH & ROADMAP",
               "สถาปัตยกรรมทางเทคนิค & แผนปฏิบัติการ Google Antigravity",
               "การผสานสแต็กเทคโนโลยีประสิทธิภาพสูงเข้ากับระบบการพัฒนาอัตโนมัติ")

    # Left Card: Technology Stack
    t_card = slide9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.9), Inches(5.7), Inches(5.1))
    t_card.fill.solid()
    t_card.fill.fore_color.rgb = CARD_BG
    t_card.line.color.rgb = BORDER_COLOR
    t_card.line.width = Pt(1.5)

    tf_tc = t_card.text_frame
    tf_tc.word_wrap = True
    p_tch = tf_tc.paragraphs[0]
    p_tch.text = "CORE TECHNICAL SPECIFICATION"
    p_tch.font.name = FONT_MONO
    p_tch.font.size = Pt(11)
    p_tch.font.bold = True
    p_tch.font.color.rgb = VIOLET

    tech_items = [
        ("Front-end Framework", "Next.js 14+ (App Router) + TypeScript + Tailwind CSS", "โครงสร้างที่รวดเร็ว รองรับ SEO และมีประสิทธิภาพสูงสุด"),
        ("Graph Visualization", "D3.js (d3-force) เรนเดอร์บน Canvas 2D/WebGL", "รองรับโหนดพร้อมกันกว่า 10,000 โหนด ลื่นไหล 60 FPS"),
        ("Design Tokens & HUD", "Chakra Petch (หัวข้อ), IBM Plex Sans Thai (เนื้อหา), IBM Plex Mono", "การ์ดตัดมุม HUD 45° (clip-path 12px) สไตล์ห้องแล็บ"),
        ("Accessibility Standards", "WCAG AA Compliance (คอนทราสต์ > 4.5:1 / 7:1)", "สลับโหมดสว่าง-มืดอัตโนมัติ มี fallback สำหรับ prefers-reduced-motion")
    ]
    for title_t, val_t, desc_t in tech_items:
        pt = tf_tc.add_paragraph()
        pt.text = f"• {title_t}: {val_t}"
        pt.font.name = FONT_BODY
        pt.font.size = Pt(10.5)
        pt.font.bold = True
        pt.font.color.rgb = TEXT_MAIN

        pd = tf_tc.add_paragraph()
        pd.text = f"   {desc_t}"
        pd.font.name = FONT_BODY
        pd.font.size = Pt(9.5)
        pd.font.color.rgb = TEXT_MUTED

    # Right Card: 4 Delivery Sprints
    s_card = slide9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.9), Inches(5.7), Inches(5.1))
    s_card.fill.solid()
    s_card.fill.fore_color.rgb = DARK_ACCENT
    s_card.line.fill.background()

    tf_sc = s_card.text_frame
    tf_sc.word_wrap = True
    p_sch = tf_sc.paragraphs[0]
    p_sch.text = "ANTIGRAVITY 4-PHASE AGILE ROADMAP"
    p_sch.font.name = FONT_MONO
    p_sch.font.size = Pt(11)
    p_sch.font.bold = True
    p_sch.font.color.rgb = LIME

    phases = [
        ("Sprint 1: Design Tokens & Style Guide", "สร้างโทเคนสี ฟอนต์ การ์ดตัดมุม HUD และหน้า Style Guide ทดสอบความคมชัด WCAG"),
        ("Sprint 2: Components & Ingestion Engine", "พัฒนาคอมโพเนนต์ Data Booming และระบบนำเข้าข้อมูลดิบจากฐานข้อมูลภาครัฐ"),
        ("Sprint 3: Thematic Halls & Visual Pages", "เปิดหน้าหลัก Home, HALL 01 Exhibitions, Everyday Lab และหน้า About"),
        ("Sprint 4: Full Graph & Sovereign AI API", "เชื่อมต่อ Force-directed Graph ฉบับเต็ม พร้อมเปิด REST/GraphQL API สู่สาธารณะ")
    ]
    for p_title, p_det in phases:
        pp1 = tf_sc.add_paragraph()
        pp1.text = f"▶ {p_title}"
        pp1.font.name = FONT_BODY
        pp1.font.size = Pt(11)
        pp1.font.bold = True
        pp1.font.color.rgb = LIME

        pp2 = tf_sc.add_paragraph()
        pp2.text = f"   {p_det}"
        pp2.font.name = FONT_BODY
        pp2.font.size = Pt(9.5)
        pp2.font.color.rgb = RGBColor(203, 213, 225)

    # -------------------------------------------------------------
    # SLIDE 10: National Strategic Impact & Next Steps
    # -------------------------------------------------------------
    slide10 = prs.slides.add_slide(blank_layout)
    apply_base_background(slide10)
    add_header(slide10, "GRM-2026-GOV-009", "POLICY & NEXT STEPS",
               "ผลประโยชน์ระดับชาติและการขับเคลื่อนเชิงนโยบาย (National Impact)",
               "ข้อเสนอแนะเพื่อการพิจารณาของ ฯพณฯ รัฐมนตรีว่าการกระทรวงดิจิทัลเพื่อเศรษฐกิจและสังคม")

    # 3 Strategic Pillars
    sp_data = [
        ("1. SOVEREIGN AI FOUNDATION", 
         "ฐานข้อมูลความรู้ภาษาไทยเพื่อ AI ชาติ", 
         CYAN, 
         "สร้างคลัง Knowledge Graph คุณภาพสูงที่มีการระบุแหล่งที่มา 100% เพื่อใช้ Fine-tune และ Grounding โมเดลปัญญาประดิษฐ์สัญชาติไทย ป้องกันปัญหา AI Hallucination ด้านประวัติศาสตร์และภูมิปัญญา"),
        ("2. DIGITAL SOFT POWER", 
         "ส่งออกภูมิปัญญาไทยสู่เวทีสากล", 
         EMBER, 
         "ยกระดับข้อมูลศิลปวัฒนธรรม ชุมชน และประวัติศาสตร์ไทยสู่มาตรฐานสากลในรูปแบบ Visual Storytelling แปล 2 ภาษา พร้อมเปิดให้พิพิธภัณฑ์ทั่วโลกเชื่อมโยงข้อมูล"),
        ("3. CREATIVE ECONOMY HUB", 
         "ต่อยอดเศรษฐกิจสร้างสรรค์และสตาร์ทอัพ", 
         VIOLET, 
         "เปิด Open Knowledge API ให้นักออกแบบ ครีเอเตอร์ ผู้พัฒนาซอฟต์แวร์ และสตาร์ทอัพไทย นำชุดข้อมูลไปพัฒนาเกม แอปพลิเคชัน หรือผลิตภัณฑ์มูลค่าสูง")
    ]

    for sp_i, (sp_code, sp_title, sp_c, sp_desc) in enumerate(sp_data):
        sp_x = Inches(0.8 + sp_i * 4.0)
        sp_shape = slide10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sp_x, Inches(1.9), Inches(3.7), Inches(3.0))
        sp_shape.fill.solid()
        sp_shape.fill.fore_color.rgb = CARD_BG
        sp_shape.line.color.rgb = BORDER_COLOR
        sp_shape.line.width = Pt(1.5)

        tf_sp = sp_shape.text_frame
        tf_sp.word_wrap = True
        p_s1 = tf_sp.paragraphs[0]
        p_s1.text = sp_code
        p_s1.font.name = FONT_MONO
        p_s1.font.size = Pt(9.5)
        p_s1.font.bold = True
        p_s1.font.color.rgb = sp_c

        p_s2 = tf_sp.add_paragraph()
        p_s2.text = sp_title
        p_s2.font.name = FONT_HEADING
        p_s2.font.size = Pt(12.5)
        p_s2.font.bold = True
        p_s2.font.color.rgb = TEXT_MAIN

        p_s3 = tf_sp.add_paragraph()
        p_s3.text = sp_desc
        p_s3.font.name = FONT_BODY
        p_s3.font.size = Pt(10)
        p_s3.font.color.rgb = TEXT_MUTED

    # Bottom Call to Action Card for the Minister
    cta_card = slide10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.1), Inches(11.7), Inches(1.9))
    cta_card.fill.solid()
    cta_card.fill.fore_color.rgb = DARK_ACCENT
    cta_card.line.fill.background()

    tf_cta = cta_card.text_frame
    tf_cta.word_wrap = True
    p_ct1 = tf_cta.paragraphs[0]
    p_ct1.text = "ข้อพิจารณาเพื่อการสนับสนุนและอนุมัติ (RECOMMENDED DECISIONS):"
    p_ct1.font.name = FONT_MONO
    p_ct1.font.size = Pt(10.5)
    p_ct1.font.bold = True
    p_ct1.font.color.rgb = LIME

    actions = [
        "1. เห็นชอบในหลักการจัดตั้งโครงการนำร่อง 'GRURU MUSEUM: National Digital Knowledge Infrastructure'",
        "2. บรรจุโครงการเข้าสู่แผนงานเร่งด่วนด้านโครงสร้างพื้นฐานดิจิทัลและเศรษฐกิจสร้างสรรค์ ประจำปีงบประมาณ 2026",
        "3. อนุมัติการเชื่อมโยงฐานข้อมูลเปิดจากหน่วยงานในกำกับดูแล (สพธอ., สรอ., ดีป้า, หอจดหมายเหตุ) เพื่อร่วมสร้างเครือข่ายความรู้แรกเริ่ม"
    ]
    for act in actions:
        pa = tf_cta.add_paragraph()
        pa.text = act
        pa.font.name = FONT_BODY
        pa.font.size = Pt(11)
        pa.font.bold = True
        pa.font.color.rgb = RGBColor(255, 255, 255)

    output_path = r'c:\Users\Administrator\Downloads\museum\GRURU_Museum_Minister_Presentation.pptx'
    prs.save(output_path)
    print(f"Successfully generated PowerPoint presentation at: {output_path}")

if __name__ == '__main__':
    create_deck()
