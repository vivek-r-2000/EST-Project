"""generate_presentation.py - Unslopped 25-minute academic defense presentation.

Builds TBD2_Final_Project_Presentation.pptx (22 slides, 16:9 widescreen)
with zero empty voids, dynamic aspect-ratio figure scaling, full-width
bottom takeaway banners, bold technical lead-ins, and sharp, emoji-free
speaker notes for an authoritative student-friendly defense.
"""
from pathlib import Path
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

ROOT = Path(__file__).resolve().parent
FIG_DIR = ROOT / "data" / "outputs" / "figures_standalone"
PPTX_OUT = ROOT / "TBD2_Final_Project_Presentation_unslopped.pptx"

# Color Palette (Okabe-Ito + Light Academic)
BG_CANVAS = RGBColor(0xF8, 0xFA, 0xFC)       # Soft off-white
CARD_BG = RGBColor(0xFF, 0xFF, 0xFF)         # Pure white card
CARD_BORDER = RGBColor(0xCB, 0xD5, 0xE1)     # Slate border
CARD_MUTED = RGBColor(0xF1, 0xF5, 0xF9)      # Subtle grey fill

COLOR_INK = RGBColor(0x0F, 0x17, 0x2A)        # Deep slate primary text
COLOR_SLATE = RGBColor(0x47, 0x55, 0x69)      # Secondary text
COLOR_MUTED = RGBColor(0x64, 0x74, 0x8B)      # Meta text
COLOR_BLUE = RGBColor(0x00, 0x72, 0xB2)       # Okabe-Ito Blue
COLOR_VERMILION = RGBColor(0xD5, 0x5E, 0x00)  # Okabe-Ito Vermilion
COLOR_GREEN = RGBColor(0x00, 0x9E, 0x73)      # Okabe-Ito Green
COLOR_ORANGE = RGBColor(0xE6, 0x9F, 0x00)     # Okabe-Ito Orange
COLOR_LINE = RGBColor(0xE2, 0xE8, 0xF0)       # Hairline divider


def create_base_deck():
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs


def add_slide_header(slide, category_tag, title_text, subtitle_text=None):
    """Draws standardized academic header."""
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = BG_CANVAS
    bg.line.fill.background()

    # Category Badge Pill
    pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.38), Inches(3.4), Inches(0.28))
    pill.fill.solid()
    pill.fill.fore_color.rgb = RGBColor(0xE0, 0xF2, 0xFE)
    pill.line.color.rgb = RGBColor(0xBA, 0xE6, 0xFD)
    pill.line.width = Pt(0.75)
    tf = pill.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = category_tag.upper()
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(8.5)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0x03, 0x69, 0xA1)
    p.font.name = "Segoe UI"

    # Slide Title Box
    tx_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.70), Inches(11.733), Inches(0.68))
    tf2 = tx_box.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = title_text
    p2.font.size = Pt(20)
    p2.font.bold = True
    p2.font.color.rgb = COLOR_INK
    p2.font.name = "Segoe UI"

    if subtitle_text:
        p_sub = tf2.add_paragraph()
        p_sub.text = subtitle_text
        p_sub.font.size = Pt(10.5)
        p_sub.font.color.rgb = COLOR_SLATE
        p_sub.font.name = "Segoe UI"
        p_sub.space_before = Pt(2)

    # Accent Rule
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.44), Inches(11.733), Inches(0.02))
    line.fill.solid()
    line.fill.fore_color.rgb = COLOR_LINE
    line.line.fill.background()


def add_bottom_banner(slide, tag_text, statement_text, bg_rgb=RGBColor(0xEE, 0xF2, 0xFF), border_rgb=RGBColor(0xC7, 0xD2, 0xFE), accent_rgb=COLOR_BLUE):
    """Draws full-width bottom takeaway anchor banner to eliminate vertical voids."""
    banner = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.02), Inches(11.733), Inches(1.05))
    banner.fill.solid()
    banner.fill.fore_color.rgb = bg_rgb
    banner.line.color.rgb = border_rgb
    banner.line.width = Pt(1)

    # Accent Stripe
    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(6.02), Inches(0.12), Inches(1.05))
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = accent_rgb
    stripe.line.fill.background()

    tb = slide.shapes.add_textbox(Inches(1.08), Inches(6.06), Inches(11.3), Inches(0.96))
    tf = tb.text_frame
    tf.word_wrap = True
    p_tag = tf.paragraphs[0]
    p_tag.text = tag_text.upper()
    p_tag.font.size = Pt(9.5)
    p_tag.font.bold = True
    p_tag.font.color.rgb = accent_rgb
    p_tag.font.name = "Segoe UI"

    p_txt = tf.add_paragraph()
    p_txt.text = statement_text
    p_txt.font.size = Pt(12)
    p_txt.font.color.rgb = COLOR_INK
    p_txt.font.name = "Segoe UI"
    p_txt.space_before = Pt(3)


def add_card_content(tf, bullets):
    """Parses bullets and renders bold lead-ins."""
    for idx, item in enumerate(bullets):
        p = tf.add_paragraph() if idx > 0 else tf.paragraphs[0]
        p.font.size = Pt(10.5)
        p.font.name = "Segoe UI"
        p.space_after = Pt(6)

        if ":" in item:
            parts = item.split(":", 1)
            lead = p.add_run()
            lead.text = f"•  {parts[0].strip()}:"
            lead.font.bold = True
            lead.font.color.rgb = COLOR_INK
            body = p.add_run()
            body.text = f" {parts[1].strip()}"
            body.font.color.rgb = COLOR_SLATE
        else:
            p.text = f"•  {item}"
            p.font.color.rgb = COLOR_SLATE


def add_card_column(slide, left, top, width, height, title, bullets, badge_text=None, title_color=COLOR_BLUE, bottom_chip=None):
    """Draws a complete vertical card column with bold lead-ins and bottom stat chip."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = CARD_BG
    card.line.color.rgb = CARD_BORDER
    card.line.width = Pt(1)

    tb = slide.shapes.add_textbox(Inches(left + 0.22), Inches(top + 0.18), Inches(width - 0.44), Inches(height - (0.75 if bottom_chip else 0.36)))
    tf = tb.text_frame
    tf.word_wrap = True

    # Title
    p_title = tf.paragraphs[0]
    p_title.text = title
    p_title.font.size = Pt(13)
    p_title.font.bold = True
    p_title.font.color.rgb = title_color
    p_title.font.name = "Segoe UI"
    p_title.space_after = Pt(8)

    # Adaptive bullet font sizing to eliminate vertical card gaps
    b_font = Pt(11) if len(bullets) <= 3 else Pt(10)
    b_space = Pt(9) if len(bullets) <= 3 else Pt(6)

    # Bullets
    for item in bullets:
        p = tf.add_paragraph()
        p.font.size = b_font
        p.font.name = "Segoe UI"
        p.space_after = b_space
        if ":" in item:
            parts = item.split(":", 1)
            r1 = p.add_run()
            r1.text = f"•  {parts[0].strip()}:"
            r1.font.bold = True
            r1.font.color.rgb = COLOR_INK
            r2 = p.add_run()
            r2.text = f" {parts[1].strip()}"
            r2.font.color.rgb = COLOR_SLATE
        else:
            p.text = f"•  {item}"
            p.font.color.rgb = COLOR_SLATE

    # Bottom stat chip if provided
    if bottom_chip:
        chip = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left + 0.2), Inches(top + height - 0.58), Inches(width - 0.4), Inches(0.42))
        chip.fill.solid()
        chip.fill.fore_color.rgb = CARD_MUTED
        chip.line.color.rgb = RGBColor(0xDF, 0xE4, 0xEA)
        chip.line.width = Pt(0.75)
        p_c = chip.text_frame.paragraphs[0]
        p_c.text = bottom_chip
        p_c.alignment = PP_ALIGN.CENTER
        p_c.font.size = Pt(9.5)
        p_c.font.bold = True
        p_c.font.color.rgb = title_color
        p_c.font.name = "Segoe UI"


def add_figure_slide_layout(slide, fig_filename, fig_aspect, card_title, bullets, kpi_stat, kpi_label, bottom_banner_tag, bottom_banner_text, accent_color=COLOR_BLUE):
    """Scales figure according to native aspect ratio and pairs with full-height sidebar and bottom banner."""
    fig_path = FIG_DIR / fig_filename
    top_pos = 1.55
    max_h = 4.32

    # Calculate optimal figure width and sidebar width based on aspect ratio
    if fig_aspect > 1.8:    # Wide figures (Fig 3, Fig 6)
        fig_w = 7.15
        fig_h = min(max_h, fig_w / fig_aspect)
        side_left = 8.12
        side_w = 4.41
    elif fig_aspect < 1.45: # Taller figures (Fig 2, Fig 4, Fig 5)
        fig_h = max_h
        fig_w = fig_h * fig_aspect
        side_left = 0.8 + fig_w + 0.22
        side_w = 11.733 - fig_w - 0.22
    else:                  # Standard figures (Fig 1)
        fig_w = 6.65
        fig_h = min(max_h, fig_w / fig_aspect)
        side_left = 7.65
        side_w = 4.88

    # Add Figure Frame
    frame = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(top_pos), Inches(fig_w), Inches(fig_h))
    frame.fill.solid()
    frame.fill.fore_color.rgb = CARD_BG
    frame.line.color.rgb = CARD_BORDER
    frame.line.width = Pt(1)

    if fig_path.exists():
        slide.shapes.add_picture(str(fig_path), Inches(0.82), Inches(top_pos + 0.02), width=Inches(fig_w - 0.04), height=Inches(fig_h - 0.04))

    # Add Sidebar KPI Box
    kpi_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(side_left), Inches(top_pos), Inches(side_w), Inches(1.10))
    kpi_box.fill.solid()
    kpi_box.fill.fore_color.rgb = CARD_BG
    kpi_box.line.color.rgb = CARD_BORDER
    kpi_box.line.width = Pt(1)

    tb_k = slide.shapes.add_textbox(Inches(side_left + 0.18), Inches(top_pos + 0.12), Inches(side_w - 0.36), Inches(0.86))
    tf_k = tb_k.text_frame
    tf_k.word_wrap = True
    p_stat = tf_k.paragraphs[0]
    p_stat.text = kpi_stat
    p_stat.font.size = Pt(22)
    p_stat.font.bold = True
    p_stat.font.color.rgb = accent_color
    p_stat.font.name = "Segoe UI"

    p_lbl = tf_k.add_paragraph()
    p_lbl.text = kpi_label
    p_lbl.font.size = Pt(10)
    p_lbl.font.bold = True
    p_lbl.font.color.rgb = COLOR_INK
    p_lbl.font.name = "Segoe UI"
    p_lbl.space_before = Pt(1)

    # Add Sidebar Evidence Card
    side_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(side_left), Inches(top_pos + 1.22), Inches(side_w), Inches(max_h - 1.22))
    side_card.fill.solid()
    side_card.fill.fore_color.rgb = CARD_BG
    side_card.line.color.rgb = CARD_BORDER
    side_card.line.width = Pt(1)

    tb_c = slide.shapes.add_textbox(Inches(side_left + 0.18), Inches(top_pos + 1.34), Inches(side_w - 0.36), Inches(max_h - 1.46))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True
    p_t = tf_c.paragraphs[0]
    p_t.text = card_title
    p_t.font.size = Pt(13)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_INK
    p_t.font.name = "Segoe UI"
    p_t.space_after = Pt(7)

    for item in bullets:
        p = tf_c.add_paragraph()
        p.font.size = Pt(10.5)
        p.font.name = "Segoe UI"
        p.space_after = Pt(5)
        if ":" in item:
            parts = item.split(":", 1)
            r1 = p.add_run()
            r1.text = f"•  {parts[0].strip()}:"
            r1.font.bold = True
            r1.font.color.rgb = COLOR_INK
            r2 = p.add_run()
            r2.text = f" {parts[1].strip()}"
            r2.font.color.rgb = COLOR_SLATE
        else:
            p.text = f"•  {item}"
            p.font.color.rgb = COLOR_SLATE

    # Add Full-Width Bottom Banner
    add_bottom_banner(slide, bottom_banner_tag, bottom_banner_text, accent_rgb=accent_color)


def add_table(slide, left, top, width, height, headers, rows, col_widths=None):
    """Draws a clean academic table."""
    num_rows = len(rows) + 1
    num_cols = len(headers)
    table_shape = slide.shapes.add_table(num_rows, num_cols, Inches(left), Inches(top), Inches(width), Inches(height))
    table = table_shape.table

    if col_widths and len(col_widths) == num_cols:
        for idx, w in enumerate(col_widths):
            table.columns[idx].width = w

    for col_idx, header_text in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_INK
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        p.text = header_text
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        p.font.name = "Segoe UI"

    for row_idx, row_data in enumerate(rows):
        bg_col = CARD_BG if row_idx % 2 == 0 else CARD_MUTED
        for col_idx, cell_value in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_col
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.text = str(cell_value)
            p.font.size = Pt(9)
            p.font.name = "Segoe UI"
            if col_idx == 0:
                p.font.bold = True
                p.font.color.rgb = COLOR_INK
            elif col_idx == 1:
                p.font.color.rgb = COLOR_SLATE
            else:
                p.font.color.rgb = COLOR_INK


def set_clean_notes(slide, slide_num, title, target_time, running_time, talk_script, transition, qna_tip=None):
    """Sets clean, unslop-compliant speaker notes without emojis or throat-clearing."""
    notes = slide.notes_slide.notes_text_frame
    text_parts = [
        f"SLIDE {slide_num} OF 22: {title.upper()}",
        f"TIMING: [Target Slide Time: {target_time} | Running Total: {running_time} / 25:00]",
        "",
        "TALKING POINTS & SCRIPT:",
        talk_script.strip(),
        "",
        f"TRANSITION:\n{transition.strip()}"
    ]
    if qna_tip:
        text_parts.extend([
            "",
            f"QUESTIONS & DEFENSE PREP:\n{qna_tip.strip()}"
        ])
    notes.text = "\n".join(text_parts)


# ==============================================================================
# 22 UNSLOPPED SLIDE BUILDERS
# ==============================================================================

def slide_01(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = COLOR_INK
    bg.line.fill.background()

    pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.8), Inches(4.8), Inches(0.36))
    pill.fill.solid()
    pill.fill.fore_color.rgb = RGBColor(0x1E, 0x29, 0x3B)
    pill.line.color.rgb = RGBColor(0x38, 0xBD, 0xF8)
    pill.line.width = Pt(1)
    p = pill.text_frame.paragraphs[0]
    p.text = "ACADEMIC FINAL DEFENSE  ·  COURSE ES1.101"
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0x38, 0xBD, 0xF8)

    tb = slide.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(11.733), Inches(2.2))
    tf = tb.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = "Late-Quaternary Climate–Vegetation Lags,\nNo-Analogue Attribution & 2D Tree Migration"
    p1.font.size = Pt(32)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    p2 = tf.add_paragraph()
    p2.text = "Quantifying 15,000 Years of Forest Dynamics Across Eastern North America Using Fossil Pollen and TraCE-21ka Simulations"
    p2.font.size = Pt(15)
    p2.font.color.rgb = RGBColor(0x94, 0xA3, 0xB8)
    p2.space_before = Pt(8)

    # 3 Summary Cards at Bottom
    cards = [
        ("RQ1: Tree Response Lags", "Tree species ran on distinct clocks, lagging behind climate by 750 to 2,500 years.", COLOR_BLUE),
        ("RQ2: Novel Ancient Forests", "Unusual extinct tree mixtures peaked at 38.2% during rapid warming and dissolved by 500 BP.", COLOR_VERMILION),
        ("RQ3: Migration Deficit", "Past trees moved at only 117–277 m/yr, falling 9.3× slower than modern warming needs.", COLOR_GREEN)
    ]
    for idx, (title, text, col) in enumerate(cards):
        c = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + idx * 3.98), Inches(4.0), Inches(3.75), Inches(1.5))
        c.fill.solid()
        c.fill.fore_color.rgb = RGBColor(0x1E, 0x29, 0x3B)
        c.line.color.rgb = col
        c.line.width = Pt(1)
        tf_c = c.text_frame
        tf_c.word_wrap = True
        pt = tf_c.paragraphs[0]
        pt.text = title
        pt.font.size = Pt(12)
        pt.font.bold = True
        pt.font.color.rgb = col
        pt.space_after = Pt(4)
        px = tf_c.add_paragraph()
        px.text = text
        px.font.size = Pt(10)
        px.font.color.rgb = RGBColor(0xCB, 0xD5, 0xE1)

    # Author Footer
    foot = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.8), Inches(11.733), Inches(0.9))
    foot.fill.solid()
    foot.fill.fore_color.rgb = RGBColor(0x15, 0x1E, 0x2E)
    foot.line.color.rgb = RGBColor(0x33, 0x41, 0x55)
    foot.line.width = Pt(1)
    tf_f = foot.text_frame
    tf_f.word_wrap = True
    pf = tf_f.paragraphs[0]
    pf.text = "Team TBD2 (Project 25): Vivek R  ·  Nandan Rajesh Kamat  ·  Shaurya Gupta"
    pf.font.size = Pt(12)
    pf.font.bold = True
    pf.font.color.rgb = RGBColor(0xF1, 0xF5, 0xF9)
    pf2 = tf_f.add_paragraph()
    pf2.text = "Neotoma Paleoecology Database (69 Cores, 1,706 Bins)  |  TraCE-21ka Paleoclimate  |  LegacyAge 1.0  |  REVEALS RPP  |  Ebisuzaki Phase Tests"
    pf2.font.size = Pt(9.5)
    pf2.font.color.rgb = RGBColor(0x38, 0xBD, 0xF8)
    pf2.space_before = Pt(3)

    set_clean_notes(slide, 1, "Title Slide", "1 min 00s", "01:00",
        "This project looks at how North American forests responded to 15,000 years of postglacial climate warming. "
        "By combining fossil pollen from 69 lake sediment cores in the Neotoma Database with independent supercomputer climate runs from TraCE-21ka, "
        "we answer three main questions: how many centuries different tree species lagged behind warming, why extinct novel forest types appeared, "
        "and how fast tree populations physically moved across the map compared to modern climate change.",
        "We start with why studying ancient forests matters on slide 2.",
        "Question: What is the geographic boundary of the study?\nAnswer: 25°N to 50°N and 95°W to 65°W, spanning eastern North America from Florida up to southern Ontario.")


def slide_02(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "1. Context & Motivation", "Why Study Ancient Forests? A 15,000-Year Natural Experiment",
                     "How trees responded to past warming gives us our only real-world benchmark for modern climate change.")

    cards = [
        ("A 15,000-Year Natural Run", [
            "Big Warming Shift (+9.6°C): As ice sheets melted, regional temperatures rose by nearly 10°C, forcing trees to migrate north.",
            "Long Lifespans Mean Slow Changes: Trees take decades to mature, so whole forests take hundreds of years to reorganize.",
            "Real-World Baseline: This provides the only empirical record of how fast tree species can naturally move under warming."
        ], "15,000-Yr Natural Run", COLOR_BLUE),
        ("Key Questions Left Unanswered", [
            "Forests Don't Move as Blocks: Older studies often assumed entire biomes moved together in neat lockstep.",
            "Unclear Delay Times: Past work did not use detrending or statistical controls to test if lags were real.",
            "Flaws in 1D Northward Models: Simple north-only lines missed how trees actually moved around lakes, valleys, and mountains."
        ], "Need for 2D Quantification", COLOR_VERMILION),
        ("Why This Matters for Today", [
            "Modern Warming Is Much Faster: Climate zones today are shifting at ~1,574 m/yr—nearly five times faster than ancient warming.",
            "Severe Risk of Falling Behind: If ancient trees struggled at 332 m/yr, modern forests will lag behind even further.",
            "Guiding Practical Conservation: Knowing natural speed limits helps us plan where to protect corridors and when to plant trees."
        ], "9.3x Velocity Gap", COLOR_GREEN)
    ]

    for idx, (title, bullets, chip, col) in enumerate(cards):
        add_card_column(slide, 0.8 + idx * 3.98, 1.55, 3.75, 4.30, title, bullets, bottom_chip=chip, title_color=col)

    add_bottom_banner(slide, "Core Takeaway",
                      "Past ice age warming gives us the best real-world benchmark to see if forests can naturally keep pace with rapid climate change.")

    set_clean_notes(slide, 2, "Ecological Context", "1 min 15s", "02:15",
        "To start off, why look back 15,000 years? We can't put a mature forest in a lab to see how it reacts to warming over centuries. "
        "But when the last ice age ended, nature ran that exact experiment. As ice sheets melted across North America, temperatures warmed by about 9.6 degrees Celsius. "
        "Because trees live for decades, they couldn't adjust overnight—it took centuries for seeds to disperse and forests to regrow. "
        "Earlier studies often assumed all trees moved north together in lockstep, and used simple 1D models that forced trees straight north. "
        "Today, human-driven warming is shifting climate zones far faster than ancient warming did. "
        "By measuring how fast trees actually moved in the past, we can figure out whether forests can keep up today.",
        "To solve this, we set up three specific research questions on slide 3.",
        "Question: Why is 500 BP used as the endpoint rather than 0 BP?\nAnswer: 500 BP represents the pre-colonial baseline before European agricultural land clearing disrupted natural forest dynamics with invasive weeds.")


def slide_03(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "1. Context & Motivation", "Three Core Research Questions",
                     "Measuring how long trees lag, why extinct forest types formed, and how fast trees migrated in 2D.")

    rqs = [
        ("RQ1: Tree Response Delays", "How many centuries did different trees lag behind warming?", [
            "Measuring Delays: Use stationary first-differencing to find exact time lags behind temperature.",
            "Speed Hierarchy: Test whether trees group into distinct speed tiers (fast pioneers, canopy trees, climax seeders).",
            "Direction of Change: Check whether each genus retreated under warming (cold-loving) or expanded (warm-loving)."
        ], COLOR_BLUE, "750–2,500 yr Delays"),
        ("RQ2: Extinct Novel Forests", "Why did unusual forest mixtures emerge with no modern counterparts?", [
            "Measuring Novelty: Use pollen dissimilarity scores against 611 modern forests to identify unique ancient mixtures.",
            "Tracking Dissolution: Test if novel forests peaked during rapid warming and dissolved completely before modern times.",
            "Isolating Drivers: Break down total novelty scores to see which specific trees created these extinct mixtures."
        ], COLOR_VERMILION, "38.2% Novelty Peak"),
        ("RQ3: 2D Migration Velocity", "Did postglacial tree migration keep pace with climate speed?", [
            "2D Arrival Maps: Fit 2D arrival surfaces on an equal-area map to track true travel directions and ground speeds.",
            "Past Speed Comparison: Compare empirical tree velocities against ancient deglacial warming speed (332 m/yr).",
            "Modern Climate Gap: Contrast natural tree speeds with modern 21st-century warming velocity (1,574 m/yr)."
        ], COLOR_GREEN, "9.3x Tracking Deficit")
    ]

    for idx, (title, subtitle, bullets, col, badge_text) in enumerate(rqs):
        top_pos = 1.55 + idx * 1.42
        c = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(top_pos), Inches(11.733), Inches(1.30))
        c.fill.solid()
        c.fill.fore_color.rgb = CARD_BG
        c.line.color.rgb = CARD_BORDER
        c.line.width = Pt(1)

        # Left Accent
        st = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(top_pos), Inches(0.10), Inches(1.30))
        st.fill.solid()
        st.fill.fore_color.rgb = col
        st.line.fill.background()

        # Right KPI Badge Pill
        badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.8), Inches(top_pos + 0.12), Inches(2.55), Inches(0.36))
        badge.fill.solid()
        badge.fill.fore_color.rgb = CARD_MUTED
        badge.line.color.rgb = CARD_BORDER
        badge.line.width = Pt(0.75)
        p_bdg = badge.text_frame.paragraphs[0]
        p_bdg.text = badge_text
        p_bdg.alignment = PP_ALIGN.CENTER
        p_bdg.font.size = Pt(10)
        p_bdg.font.bold = True
        p_bdg.font.color.rgb = col
        p_bdg.font.name = "Segoe UI"

        tb = slide.shapes.add_textbox(Inches(1.05), Inches(top_pos + 0.08), Inches(8.65), Inches(1.14))
        tf = tb.text_frame
        tf.word_wrap = True
        pt = tf.paragraphs[0]
        pt.text = f"{title}:  {subtitle}"
        pt.font.size = Pt(12)
        pt.font.bold = True
        pt.font.color.rgb = col
        pt.space_after = Pt(3)

        for b in bullets:
            pb = tf.add_paragraph()
            pb.font.size = Pt(10)
            pb.font.name = "Segoe UI"
            pb.space_after = Pt(2)
            if ":" in b:
                parts = b.split(":", 1)
                r1 = pb.add_run()
                r1.text = f"• {parts[0].strip()}:"
                r1.font.bold = True
                r1.font.color.rgb = COLOR_INK
                r2 = pb.add_run()
                r2.text = f" {parts[1].strip()}"
                r2.font.color.rgb = COLOR_SLATE
            else:
                pb.text = f"•  {b}"
                pb.font.color.rgb = COLOR_SLATE

    add_bottom_banner(slide, "Research Integration",
                      "Together, these three questions show when forests reacted (RQ1), what unique communities formed (RQ2), and how fast trees physically moved across the land (RQ3).")

    set_clean_notes(slide, 3, "Research Questions", "1 min 15s", "03:30",
        "Slide 3 outlines our three research questions. "
        "First, RQ1 asks: How many centuries did different trees lag behind warming, and did they expand or retreat? "
        "Second, RQ2 examines novel extinct forests: When did unusual mixtures of trees appear, and what species caused them? "
        "Third, RQ3 measures real physical speed: When trees traveled across a 2D map, how fast did they move compared to shifting climate zones? "
        "These questions directly test whether forests can naturally keep pace with warming.",
        "Slide 4 introduces our lake sediment dataset.",
        "Question: Why is 2D planar modeling superior to 1D?\nAnswer: 1D models assume due-north travel and invert linear slopes, which artificially inflates speeds. 2D modeling measures true ground gradients along actual migration corridors.")


def slide_04(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "2. Data & Study Corridor", "Our Dataset: 69 Lake Sediment Cores Across North America",
                     "Pulling 15,000 years of fossil pollen records from lake mud stored in the Neotoma Database.")

    bullets = [
        "Study Region (25°–50°N, 95°–65°W): Spans eastern North America from Florida up to southern Canada.",
        "Quality Filtering: Kept 69 lakes with clean, continuous mud layers and at least 150 identified pollen grains per sample.",
        "Consistent Time Slices: Standardized into 1,706 500-year slices running from 14,750 down to 250 years ago.",
        "Modern Baseline Network: 611 modern surface lake samples provide the comparison baseline for our novelty tests."
    ]

    add_figure_slide_layout(slide, "fig1_study_sites.png", 1.68,
                            "Fossil Site Distribution & Quality Criteria", bullets,
                            "69 Lakes / 1,706 Bins", "Quality-Filtered Dataset",
                            "Geographic Coverage",
                            "These 69 continuous lake cores give us an unbroken, high-resolution timeline of forest changes across eastern North America.",
                            COLOR_BLUE)

    set_clean_notes(slide, 4, "Study Sites & Data", "1 min 15s", "04:45",
        "Figure 1 on slide 4 shows our fossil lake network. Each point on the map is a sediment core drilled from the bottom of a lake. "
        "Pollen grains trapped in lake mud provide an unbroken record of past forests. "
        "We filtered the Neotoma Database strictly: only lakes with at least 150 identified pollen grains per sample and continuous coverage were included, giving us 69 high-quality lakes and 1,706 500-year bins. "
        "Blue circles indicate 57 cores with modern LegacyAge 1.0 Bayesian chronologies, while amber circles mark 12 cores with Neotoma depth-scaled dates.",
        "Next, slide 5 introduces the independent climate model we used to track temperature.",
        "Question: How were modern surface samples collected?\nAnswer: Modern samples consist of top-centimeter surface mud from 611 lakes across North America, reflecting pre-industrial forest composition.")


def slide_05(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "2. Data & Study Corridor", "Independent Climate Data: NCAR TraCE-21ka Simulation",
                     "Using a physics-based supercomputer model prevents the trap of using pollen to predict pollen.")

    cards = [
        ("Avoiding Circular Logic", [
            "The Common Trap: Many older studies guessed temperatures from pollen, then used that temperature to explain tree movement.",
            "Physics-Based Model: TraCE-21ka is a global ocean-atmosphere climate model run by NCAR, driven by Earth's orbit, greenhouse gases, and ice melt.",
            "Truly Independent: Vegetation changes are tested against an independent physical climate model, not an estimate made from pollen."
        ], "Zero Circular Reasoning", COLOR_BLUE),
        ("Climate Variables Tracked", [
            "Surface Temperature: Tracked in °C at each of our 69 lake locations across every 500-year time slice.",
            "Annual Rainfall: Measured precipitation (mm/day) to check moisture changes alongside temperature.",
            "Northward Temperature Gradient: Drops by 0.00121°C per km north (~1.2°C per 1,000 km), key for calculating climate speed."
        ], "0.00121 °C/km Northward", COLOR_VERMILION),
        ("Major Ancient Climate Pulses", [
            "Bølling-Allerød Warm Pulse (14.7–12.9 ka BP): Rapid warming surge that kicked off major ice sheet retreat.",
            "Younger Dryas Cold Snap (12.9–11.7 ka BP): Abrupt temporary cooling that froze back early deglacial warmth.",
            "8.2 ka Cold Event: Massive meltwater flood into the North Atlantic that cooled temperatures for a few centuries."
        ], "Rapid Climate Pulses", COLOR_GREEN)
    ]

    for idx, (title, bullets, chip, col) in enumerate(cards):
        add_card_column(slide, 0.8 + idx * 3.98, 1.55, 3.75, 4.30, title, bullets, bottom_chip=chip, title_color=col)

    add_bottom_banner(slide, "Climate Independence",
                      "TraCE-21ka gives us independent temperature and rainfall time series, keeping our climate data cleanly separated from our pollen data.")

    set_clean_notes(slide, 5, "TraCE-21ka Paleoclimate", "1 min 15s", "06:00",
        "Slide 5 outlines our climate dataset. A common flaw in older papers is circular logic: using fossil pollen to reconstruct temperature, and then using that temperature to explain pollen changes. "
        "To avoid that trap, we pulled temperature and rainfall from TraCE-21ka, an independent climate simulation run on NCAR supercomputers. "
        "TraCE models orbital cycles, greenhouse gases, and ice-sheet melting purely from physics principles. "
        "Importantly, TraCE gives us our regional temperature gradient—0.00121 degrees Celsius per kilometer—which is the number we need to calculate how fast climate zones moved.",
        "Slide 6 covers the biological and dating corrections we applied to raw pollen counts.",
        "Question: Does TraCE capture regional topography?\nAnswer: Yes, TraCE incorporates changing Laurentide ice sheet elevation and coastal boundaries at decadal resolution throughout the 22,000-year run.")


def slide_06(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "2. Data & Study Corridor", "Cleaning the Data: Tree Pollen, Weeds & Carbon Dating",
                     "Fixing differences in pollen output between tree species, removing weed spikes, and standardizing dates.")

    cards = [
        ("Adjusting Pollen Output (REVEALS)", [
            "Unequal Pollen Production: Pine and Birch blast huge clouds of airborne pollen; Maple and Tamarack produce far fewer grains.",
            "Productivity Multipliers: We applied REVEALS multipliers to convert raw pollen percentages into true tree abundances on the ground.",
            "Restoring Real Proportions: Scaled down overrepresented pines and brought hidden hardwoods back to their true percentages."
        ], "True Ground Abundances", COLOR_BLUE),
        ("Filtering Farm Weeds (Trees Only)", [
            "Farming Disturbance Spike: European land clearance after 1700 CE caused an explosion of ragweed and grass pollen.",
            "False Novelty Signal: Leaving weeds in makes recent samples look artificially weird just because of European farms.",
            "Keeping Woody Trees Only: We isolated tree and shrub pollen (TRSH-only), revealing true natural forest states."
        ], "Weed Artifacts Removed", COLOR_VERMILION),
        ("Standardizing Dates (LegacyAge 1.0)", [
            "Upgraded Carbon Dates: Replaced older uncalibrated core depths with modern Bayesian age models from Li et al. (2022).",
            "Testing Date Uncertainty: Tested 50 randomized Monte Carlo date draws to ensure age measurement jitter didn't alter our conclusions.",
            "Consistent Calendar Timeline: Standardized every lake core on a common calendar timeline across all 15,000 years."
        ], "Standardized Calendar Ages", COLOR_GREEN)
    ]

    for idx, (title, bullets, chip, col) in enumerate(cards):
        add_card_column(slide, 0.8 + idx * 3.98, 1.55, 3.75, 4.30, title, bullets, bottom_chip=chip, title_color=col)

    add_bottom_banner(slide, "Data Corrections",
                      "Correcting pollen productivity differences, filtering farm weeds, and standardizing radiocarbon dates turned raw pollen counts into reliable forest histories.")

    set_clean_notes(slide, 6, "Data Corrections", "1 min 15s", "07:15",
        "Slide 6 details three essential data corrections we made. "
        "First, pollen grains don't map 1-to-1 to trees on the ground. Pine trees produce huge clouds of light, winged pollen, while maples produce very little. "
        "We used REVEALS productivity multipliers to adjust for these biological differences so raw counts reflect real tree counts. "
        "Second, post-colonial farming caused massive spikes of ragweed and agricultural weeds. If you leave weeds in, recent samples falsely look novel. We isolated strictly woody trees and shrubs. "
        "Third, we standardized all radiocarbon dates using modern LegacyAge 1.0 Bayesian age models and tested them against 50 randomized draws.",
        "Slide 7 explains our statistical tests for significance.",
        "Question: How much does pine pollen over-represent pine trees?\nAnswer: Pine produces roughly 10 times more pollen per unit basal area than sugar maple, which heavily skews raw pollen diagrams if unadjusted.")


def slide_07(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "2. Data & Study Corridor", "Statistical Toolkit: Proving Our Findings Are Real",
                     "Removing false trends, testing against 1,000 random climate series, and tracking trees in 2D.")

    cards = [
        ("First-Differencing (Change vs Change)", [
            "The False Correlation Trap: Because both temperature and hardwoods rose over 15,000 years, raw comparisons show high correlation by accident.",
            "Comparing Steps, Not Trends: First-differencing looks at bin-to-bin changes: does a jump in temperature lead to a jump in trees?",
            "Stationary Testing: Removes overall millennial trends to test real cause-and-effect response timing."
        ], "Bin-to-Bin Changes Tested", COLOR_BLUE),
        ("1,000 Surrogate Tests (Ebisuzaki)", [
            "Climate Memory Problem: Environmental time series have autocorrelation, which can make regular p-values appear falsely significant.",
            "1,000 Synthetic Tests: Created 1,000 randomized climate test series with identical frequency properties to test our lags against.",
            "Strict Pass Rule: A lag is considered real only if its correlation beats 95% of random surrogate runs (p < 0.05)."
        ], "Strict p < 0.05 Significance", COLOR_VERMILION),
        ("2D Equal-Area Migration", [
            "Beyond 1D North Assumptions: Replaced straight-north lines with 2D arrival surfaces on an equal-area map.",
            "Direction and Velocity: Calculated true ground speed and travel direction along actual dispersal paths.",
            "Lake Corridor Boundaries: Masked calculations to the actual lake network so models don't guess speeds outside our data."
        ], "True 2D Landscape Speeds", COLOR_GREEN)
    ]

    for idx, (title, bullets, chip, col) in enumerate(cards):
        add_card_column(slide, 0.8 + idx * 3.98, 1.55, 3.75, 4.30, title, bullets, bottom_chip=chip, title_color=col)

    add_bottom_banner(slide, "Statistical Rigor",
                      "Detrending removed false correlations, 1,000 surrogate runs confirmed statistical significance, and 2D mapping captured real landscape travel.")

    set_clean_notes(slide, 7, "Statistical Methods", "1 min 15s", "08:30",
        "Slide 7 explains our statistical framework. "
        "A common trap in past studies is correlating trending data: because both temperature and oak trees rose over 15,000 years, a basic regression gives a high correlation that is completely accidental. "
        "We eliminated this by first-differencing: looking at bin-to-bin changes instead of raw levels. "
        "Next, climate series have temporal memory. To make sure our correlations were real, we tested every lag against 1,000 randomized surrogate climate series. "
        "Finally, we moved from 1D latitude lines to 2D arrival surfaces to measure real travel speeds across the landscape.",
        "Now let's dive into our results, starting with 15,000 years of climate and forest turnover on slide 8.",
        "Question: How does the Ebisuzaki method generate surrogates?\nAnswer: It computes the discrete Fourier transform of the climate series, randomizes the Fourier phase angles while preserving amplitudes, and inverts back to time domain, maintaining exact spectral properties.")


def slide_08(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "3. Results: Climate & Turnover", "15,000 Years of Climate Warming & Forest Shifts",
                     "Tracking temperature, rainfall, warming rates, and forest turnover across abrupt climate events.")

    bullets = [
        "Big Warming Jump (+9.6°C): Regional average temperature rose from -1.1°C at 15,000 BP up to +8.5°C during the warm mid-Holocene.",
        "Rising Rainfall (+1.2 mm/day): Moisture increased by 60% from dry ice age conditions (2.0 mm/day) to wet Holocene conditions (3.2 mm/day).",
        "Fastest Warming Surge (+0.43°C/century): The most rapid warming burst hit at 14,500 BP during the Bølling-Allerød warm pulse.",
        "Forest Turnover Matched Warming: Community turnover peaked right alongside the fastest warming, showing that temperature shocks directly forced forest changes."
    ]

    add_figure_slide_layout(slide, "fig2_climate_and_turnover.png", 1.38,
                            "Abrupt Climate Shifts & Forest Turnover", bullets,
                            "+9.6 °C Total Warming", "Peak Rate: 0.43 °C/century",
                            "Turnover Matched Warming",
                            "Forest communities reorganized fastest during periods of rapid warming, proving that ancient climate shocks immediately shook up forest composition.",
                            COLOR_VERMILION)

    set_clean_notes(slide, 8, "Climate & Turnover", "1 min 15s", "09:45",
        "Figure 2 shows 15,000 years of climate and forest history side by side. "
        "Up top, we track temperature in orange and rainfall in blue: the region warmed by almost 10°C, and rainfall rose by about 60%. "
        "The yellow and blue shaded bands highlight rapid climate jumps—the warm Bølling-Allerød and the Younger Dryas cold snap. "
        "Look at the bottom two panels: the third panel shows how fast climate warmed (peaking at 0.43°C per century), and the bottom panel shows vegetation turnover. They peak together! "
        "When warming accelerated, forest communities scrambled and reorganized the most.",
        "We now address Research Question 1 on slide 9: How long did trees lag behind these temperature shifts?",
        "Question: Why does turnover drop during the Holocene?\nAnswer: Mid-Holocene climate stabilized, allowing tree populations to establish quasi-equilibrium communities with lower species turnover.")


def slide_09(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "3. Results: RQ1 (Response Lags)", "RQ1: Trees Operate on Three Distinct Speed Clocks",
                     "Forests do not move as a single synchronized unit; tree genera lag by 750 to 2,500 years across three speed tiers.")

    bullets = [
        "Statistically Significant (p = 0.0003): Kruskal-Wallis testing proves that tree genera do not share the same response speed.",
        "Fast Pioneers (<1,000 yr): Tamarack (Larix) reacted first, with a median lag of 750 years [0–2,500 yr CI].",
        "Canopy Hardwoods (1,000–2,000 yr): Major trees—Oak, Birch, Ash, Elm, and Maple—clustered around a 1,500-year delay.",
        "Heavy Climax Seeders (>2,000 yr): Beech and Hemlock took the longest, lagging temperature by 2,500 years [2,000–3,000 yr CI]."
    ]

    add_figure_slide_layout(slide, "fig3_genus_lags_raincloud.png", 1.82,
                            "Three Distinct Ecological Speed Tiers", bullets,
                            "750 to 2,500 yr Lags", "Kruskal-Wallis p = 0.0003",
                            "Ecological Speed Hierarchy",
                            "Trees split into three distinct speed groups: light-seeded pioneers reacted in 750 years, canopy hardwoods took 1,500 years, and heavy-seeded climax trees took 2,500 years.",
                            COLOR_BLUE)

    set_clean_notes(slide, 9, "RQ1 Response Lags", "1 min 15s", "11:00",
        "Here in Figure 3 Panel A, we answer Research Question 1. "
        "Each row is a different tree genus, the dots mark median lag in years, and the horizontal bars show 95% confidence intervals. "
        "The Kruskal-Wallis test gives p = 0.0003, confirming that different trees move on completely different timelines. "
        "We see three clear groups: fast pioneers like Tamarack take about 750 years; canopy hardwoods like Oak, Maple, and Ash cluster at 1,500 years; and slow climax trees like Beech and Hemlock take 2,500 years. "
        "This shows that biomes don't migrate as a single synchronized block—each tree genus follows its own speed clock.",
        "Slide 10 examines whether these lags represented population expansion or collapse.",
        "Question: Why does Larix have a wider confidence interval?\nAnswer: Larix was restricted to northern sites near the ice margin, resulting in a smaller sample size (n=10) than widespread oaks (n=67).")


def slide_10(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "3. Results: RQ1 (Response Lags)", "RQ1: Winners and Losers During Ancient Warming",
                     "Separating cold-loving trees that retreated from warm-adapted trees that expanded northward.")

    bullets = [
        "Boreal Spruce Retreat (95% Negative): Spruce declined as temperatures rose across 95% of lake records (78% passing strict surrogate tests).",
        "Oak Expansion (79% Positive): Warm-loving Oak expanded rapidly northward as melting ice sheets opened new ground.",
        "Hardwood Reorganization: Maple (67% positive) and Birch expanded behind oak along deglaciated corridors.",
        "The Delayed Equilibrium: Even warm-expanding oaks still lagged climate by 1,500 years due to seed dispersal and soil-building limits."
    ]

    add_figure_slide_layout(slide, "fig3_genus_lags_raincloud.png", 1.82,
                            "Biological Direction: Retreat vs. Colonization", bullets,
                            "95% Spruce Retreat", "79% Oak Expansion",
                            "Response Direction",
                            "Deglacial warming triggered massive dieback in cold-adapted spruce (95% negative) while opening the way for temperate hardwoods like oak to colonize north.",
                            COLOR_VERMILION)

    set_clean_notes(slide, 10, "RQ1 Response Polarity", "1 min 15s", "12:15",
        "Slide 10 looks at the second half of Figure 3: which way did each tree react? "
        "Blue bars mean the tree retreated as it got warmer, and orange bars mean it expanded. "
        "Look at Spruce: 95% of lakes show a negative correlation, with 78% passing our strict surrogate tests. Warming caused cold-adapted spruce to collapse across the region. "
        "On the flip side, Oak was 79% positive, expanding northward to take over newly opened ground. "
        "But even for oaks that benefited from warming, they still lagged 1,500 years behind temperature, showing how slowly forests reach equilibrium.",
        "Next, slide 11 explores Research Question 2: the emergence of novel ecosystems during this transition.",
        "Question: Why did spruce decline so consistently across lakes?\nAnswer: Rising summer temperatures exceeded the thermal physiological limit of white and black spruce, increasing respiration costs and pest vulnerability.")


def slide_11(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "3. Results: RQ2 (No-Analogue Forests)", "RQ2: Novel Communities Peaked at 38% and Vanished by 500 BP",
                     "Extinct mixtures of trees peaked during the fastest deglacial warming and disappeared before colonial disturbance.")

    bullets = [
        "Identifying Novel Forests: Any fossil pollen sample with a dissimilarity score above 0.138 SCD compared to 611 modern forests is classified as novel.",
        "Peak During Rapid Warming (38.2%): Novel communities peaked at 14,500 BP right during the Bølling-Allerød warming pulse.",
        "Complete Disappearance by 500 BP (0.0%): Once we filtered out modern agricultural weeds, ancient novel forests dropped to exactly 0.0% by 500 BP.",
        "Direct Link to Turnover: Novel forest prevalence strongly correlated with vegetation turnover (Spearman rho = 0.743, p = 0.021)."
    ]

    add_figure_slide_layout(slide, "fig4_noanalogue_drivers.png", 1.47,
                            "Deglacial Novelty Trajectory & Correlation", bullets,
                            "38.2% Peak Novelty", "0.0% Pre-Colonial Baseline",
                            "Transient Novelty",
                            "Novel forest communities peaked during rapid deglacial warming and dissolved completely to 0.0% by 500 BP once modern climate seasonality stabilized.",
                            COLOR_VERMILION)

    set_clean_notes(slide, 11, "RQ2 No-Analogue Dynamics", "1 min 15s", "13:30",
        "Slide 11 addresses Research Question 2: Why did ancient forests have tree combinations that don't exist today? "
        "We measured novelty using Squared Chord Distance against 611 modern lakes. "
        "The top graph shows that novel communities peaked at 38.2% around 14,500 years ago—right when warming was fastest during the Bølling-Allerød. "
        "But look at where it ends: by 500 years ago, novelty dropped all the way to zero. "
        "Older papers that thought novel forests persisted into the late Holocene were fooled by post-settlement farm weeds. In natural pre-colonial forests, novelty was a temporary phase.",
        "Slide 12 breaks down the specific tree species that generated this novelty.",
        "Question: What is Squared Chord Distance?\nAnswer: SCD is a non-linear metric of ecological dissimilarity; values above 0.138 indicate community compositions with no modern analogue in the reference database.")


def slide_12(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "3. Results: RQ2 (No-Analogue Forests)", "RQ2: What Made These Ancient Forests So Strange?",
                     "Breaking down dissimilarity scores shows that Spruce and Black Ash drove these extinct tree mixtures.")

    bullets = [
        "Decomposing Novelty Scores: Because Squared Chord Distance is additive, we can calculate how much each tree genus contributed to the strange mix.",
        "Spruce Was the #1 Driver (0.161 SCD): Boreal Spruce contributed nearly half (48%) of all community novelty during early deglaciation.",
        "Black Ash Was #2 (0.057 SCD): Temperate wetland Black Ash contributed 17% of the novelty, co-occurring unnaturally with spruce.",
        "The Seasonal Mechanism: High orbital seasonality created hot summers and freezing winters, allowing cold-loving spruce and wetland ash to grow side-by-side."
    ]

    add_figure_slide_layout(slide, "fig4_noanalogue_drivers.png", 1.47,
                            "Taxonomic Dissimilarity Attribution", bullets,
                            "Picea SCD: 0.161", "Fraxinus SCD: 0.057",
                            "Taxonomic Drivers",
                            "Novel forests were created by unusual co-occurrences of boreal spruce and temperate black ash under high seasonal temperature contrast.",
                            COLOR_BLUE)

    set_clean_notes(slide, 12, "RQ2 Taxonomic Drivers", "1 min 15s", "14:45",
        "Figure 4's bottom heatmap shows which trees actually created these strange mixtures. "
        "Because our dissimilarity metric is additive, we broke it down tree by tree. "
        "The darkest row right at the top is Spruce: it contributed almost half of all novelty with an average score of 0.161. Second was Black Ash at 0.057. "
        "Why were they growing together? Back then, orbital cycles produced huge seasonal contrast—hot summers with freezing cold winters. "
        "That strange climate allowed cold-hardy spruce and wetland ash to grow side by side in ways they never do today. Once seasonality moderated, the mix disappeared.",
        "Slide 13 transitions to Research Question 3: mapping 2D tree migration fronts across the map.",
        "Question: Why didn't oak or pine drive novelty?\nAnswer: Oak and pine maintained abundances that closely resembled modern southeastern pine-oak woodlands, contributing minimal dissimilarity.")


def slide_13(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "3. Results: RQ3 (2D Migration)", "RQ3: Mapping 15,000 Years of Tree Travel in 2D",
                     "Replacing straight-north assumptions with planar arrival maps and real travel directions.")

    bullets = [
        "Strict Significance Gating: Tested all 12 genera on 2D arrival surfaces Age(x, y), strictly requiring p <= 0.05 to ensure real directional travel.",
        "5 Validated Migrants: Hickory, Alder, Beech, Tamarack, and Hemlock showed clear, coherent regional migration fronts.",
        "Staying Inside the Data: Arrival contours are masked strictly to the lake network boundary, preventing artificial guesses outside our sites.",
        "Directional Travel Vectors: Black arrows show how trees actually traveled around river valleys, the Great Lakes, and the Appalachians."
    ]

    add_figure_slide_layout(slide, "fig5_migration_isochrones.png", 1.44,
                            "Empirical 2D Colonization Isochrones", bullets,
                            "5 Validated Migrants", "p_2d <= 0.05 Significance",
                            "2D Migration Maps",
                            "Using 2D planar arrival surfaces identified five statistically validated migrating trees and revealed their true paths around major geographic obstacles.",
                            COLOR_GREEN)

    set_clean_notes(slide, 13, "RQ3 2D Arrival Isochrones", "1 min 15s", "16:00",
        "Figure 5 brings us to Research Question 3. "
        "Instead of assuming trees just marched straight north, we fitted 2D planar arrival surfaces across eastern North America. "
        "We ran a strict statistical test (F-test, p <= 0.05) and found that five tree genera had clear, organized migration waves: Hickory, Alder, Beech, Tamarack, and Hemlock. "
        "The remaining seven taxa showed scattered local pockets rather than a single broad wave. "
        "The colored contours show arrival times from 12,000 years ago in dark blue to 4,000 years ago in yellow, and the black arrows show the actual direction trees traveled.",
        "Slide 14 details the specific travel routes and speeds of these five trees.",
        "Question: Why are isochrones clipped to convex hulls?\nAnswer: To prevent mathematical extrapolation into regions where a tree was never recorded, ensuring map integrity.")


def slide_14(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "3. Results: RQ3 (2D Migration)", "RQ3: Real Migration Paths and Travel Speeds",
                     "Topographic barriers like mountains, valleys, and the Great Lakes steered where and how fast trees migrated.")

    headers = ["Genus", "Common Name", "Median Speed", "Full Speed Range", "Travel Heading", "Key Migration Route"]
    rows = [
        ["Carya", "Hickory", "277.0 m/yr", "217–296 m/yr", "65.0° (ENE)", "Mississippi & Ohio valleys toward Atlantic coast"],
        ["Alnus", "Alder", "264.6 m/yr", "254–278 m/yr", "308.2° (WNW)", "New England across Great Lakes into Minnesota"],
        ["Fagus", "Beech", "169.6 m/yr", "157–185 m/yr", "344.4° (NNW)", "Southern Appalachians north into New England"],
        ["Larix", "Tamarack", "124.0 m/yr", "115–129 m/yr", "1.2° (Due North)", "Followed retreating Laurentide ice across Midwest"],
        ["Tsuga", "Hemlock", "116.9 m/yr", "112–120 m/yr", "344.3° (NNW)", "Parallel to Appalachian spine into Great Lakes"]
    ]

    col_w = [Inches(1.6), Inches(1.5), Inches(1.4), Inches(1.8), Inches(1.6), Inches(3.833)]
    add_table(slide, 0.8, 1.55, 11.733, 4.30, headers, rows, col_w)

    add_bottom_banner(slide, "Real Migration Routes",
                      "Trees did not march due North: Hickory traveled East-North-East (65°) along river corridors, Alder headed West-North-West (308°), and only Tamarack moved straight North (1°).")

    set_clean_notes(slide, 14, "RQ3 Dispersal Corridors", "1 min 15s", "17:15",
        "Slide 14 summarizes the exact numbers from our 2D maps. "
        "Look at the compass directions: only Tamarack traveled straight north at 1 degree azimuth, tracking retreating ice. "
        "Hickory moved East-North-East at 65 degrees along the fertile Ohio and Mississippi river valleys at 277 meters per year. "
        "Alder traveled West-North-West at 308 degrees around the Great Lakes at 265 meters per year. "
        "Beech and Hemlock moved at 170 and 117 meters per year along the Appalachian mountain chain. "
        "This proves that 1D north-only models completely miss the real geography of tree migration.",
        "Slide 15 compares these speeds directly to climate change.",
        "Question: Why did Hickory move East-North-East?\nA: Glacial refugia for hickory were concentrated in the lower Mississippi basin; postglacial warming opened river corridors extending toward the Atlantic coast.")


def slide_15(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "3. Results: RQ3 (2D Migration)", "RQ3: The 9.3x Speed Deficit Under Modern Warming",
                     "Past tree migration hit a biological ceiling below 300 m/yr—falling nearly an order of magnitude behind today's warming.")

    bullets = [
        "A Hard Biological Speed Limit (117–277 m/yr): Across all validated trees, real population speeds topped out at 296 m/yr (Hickory max).",
        "Slower Than Ancient Warming: Every single tree moved slower than natural deglacial climate velocity (332 m/yr, blue dashed line).",
        "Modern Warming Is Much Faster: 21st-century warming requires trees to track climate at 1,574 m/yr (orange dashed line).",
        "The 9.3x Tracking Gap: Median past tree migration (170 m/yr) is 9.3 times slower than what is needed to keep pace with modern warming."
    ]

    add_figure_slide_layout(slide, "fig6_velocity_deficit.png", 2.04,
                            "Log-Scale Velocity Comparison", bullets,
                            "9.3x Speed Deficit", "170 m/yr vs. 1,574 m/yr",
                            "Core Conservation Finding",
                            "Past tree migration speeds (117–277 m/yr) were slower than ancient deglacial warming (332 m/yr) and trail modern human-driven warming (1,574 m/yr) by 9.3 times.",
                            COLOR_VERMILION)

    set_clean_notes(slide, 15, "RQ3 Velocity Deficit", "1 min 15s", "18:30",
        "Figure 6 on slide 15 is the core punchline of our entire project. "
        "The horizontal axis shows speed in meters per year on a log scale. The green markers show our 2D tree speeds and their ranges. "
        "Notice two critical reference lines: "
        "First, the blue dashed line at 332 meters per year is how fast climate zones moved during ancient deglaciation. Every single tree is to the left of that line—trees were always slower than climate. "
        "Second, look at the orange dashed line at 1,574 meters per year: that's how fast climate zones are shifting today under 1.9°C of warming per century. "
        "Median tree migration was 170 meters per year. That is a 9.3-fold speed deficit. Natural seed dispersal simply cannot keep up with modern warming on its own.",
        "Slide 16 presents sensitivity checks on our models to ensure these results are solid.",
        "Question: Why did earlier mid-project models cross the 332 m/yr line?\nAnswer: Mid-project 1D models inverted linear regression slopes with wide standard errors, creating mathematical CI blowups that falsely exceeded 500 m/yr.")


def slide_16(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "4. Sensitivity & Robustness", "Checking Our Work: Is the Model Reliable?",
                     "Testing age uncertainty, finer time bins, and 1,000 surrogate climate runs.")

    cards = [
        ("Testing Carbon Dating Jitter (N=50 Draws)", [
            "Random Date Draws: Drew 50 random age-depth realizations from LegacyAge 1.0 Bayesian chronologies across all 69 lakes.",
            "Unchanging Lags: The median response lag across all trees stayed locked at exactly 1,500 years in every single run.",
            "Takeaway: Radiocarbon calibration uncertainty does not change our core conclusions."
        ], "Median Lag = 1,500 yr Locked", COLOR_BLUE),
        ("Finer 250-Year Time Bins", [
            "Halving the Time Step: Re-ran our entire pipeline using finer 250-year slices (66 lakes, 3,365 time levels).",
            "Same Speed Order: Found the identical speed hierarchy—Tamarack fastest, common hardwoods in the middle, Beech and Hemlock slowest.",
            "Takeaway: Using 500-year bins did not blur or distort tree response patterns."
        ], "Same Tree Order Verified", COLOR_VERMILION),
        ("Surrogate Test Pass Rates", [
            "Synthetic Climate Controls: Tested our findings against 1,000 synthetic climate runs to weed out random chance.",
            "Strong Pass Rates: 78.4% of Spruce retreat pairs and 67.2% of Oak expansion pairs passed the strict 95% threshold.",
            "Takeaway: Tree-climate correlations represent real physical reactions, not random mathematical noise."
        ], "78% Spruce / 67% Oak Passed", COLOR_GREEN)
    ]

    for idx, (title, bullets, chip, col) in enumerate(cards):
        add_card_column(slide, 0.8 + idx * 3.98, 1.55, 3.75, 4.30, title, bullets, bottom_chip=chip, title_color=col)

    add_bottom_banner(slide, "Validation Verdict",
                      "Three separate sensitivity stress tests confirm that our tree lag hierarchy, novel forest trends, and 2D speed limits hold up under rigorous testing.")

    set_clean_notes(slide, 16, "Sensitivity Audits", "1 min 15s", "19:45",
        "On slide 16, we stress-tested our models to make sure our findings weren't just artifacts of our code or settings. We tested three things: "
        "First, we re-ran the analysis 50 times with random draws from the radiocarbon age models; the median lag stayed locked at 1,500 years every time. "
        "Second, we doubled the time resolution to 250-year bins across over 3,300 levels, and the exact same tree hierarchy held up. "
        "Third, our surrogate climate tests showed that roughly three-quarters of the spruce and oak signals passed strict significance tests. "
        "This gives us high confidence that the patterns are biologically real.",
        "Slide 17 explores the biological reasons behind these 1,500-year delays.",
        "Question: Did downweighting Pinus alter the lag hierarchy?\nAnswer: No, downweighting Pinus reduced overproduction distortion but preserved the identical pioneer-to-climax lag ordering.")


def slide_17(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "4. Ecological Synthesis", "Why Do Trees Lag? Three Biological Bottlenecks",
                     "Real-world biological constraints stop forests from instantly following shifting climate zones.")

    cards = [
        ("Decades to Mature & Heavy Seeds", [
            "Slow Sexual Maturity: Forest trees take 15 to 40 years before producing their first viable seed crop.",
            "Intermittent Mast Years: Heavy-seeded hardwoods (Oak, Beech, Hickory) produce bumper crops only once every 3 to 7 years.",
            "Limited Seed Travel: Most heavy nuts fall within 30 meters of the parent tree unless carried by animals."
        ], "Natural Seed Limit <300 m/yr", COLOR_BLUE),
        ("Building Soil Takes Centuries", [
            "Bare Glacial Ground: Melting ice sheets left behind raw gravel and crushed stone with zero organic matter or nitrogen.",
            "Decades of Succession: Building rich forest topsoil requires centuries of lichen weathering and nitrogen-fixing plants like Alder.",
            "Climax Trees Need Soil: Demanding hardwoods like Beech and Maple cannot establish until thick humus and fungal networks develop."
        ], "Centuries to Build Forest Soil", COLOR_VERMILION),
        ("Mature Forest Shading & Competition", [
            "Dense Tree Canopies: Established adult trees block out sunlight and buffer forest floor microclimates for centuries.",
            "Seedlings Blocked: Even when regional temperatures warm up, existing spruce and pine stands crowd out invading hardwood seedlings.",
            "Waiting for Disturbances: New hardwood species can only take over when fire, wind, or old age knock down mature trees."
        ], "Canopy Shading & Inertia", COLOR_GREEN)
    ]

    for idx, (title, bullets, chip, col) in enumerate(cards):
        add_card_column(slide, 0.8 + idx * 3.98, 1.55, 3.75, 4.30, title, bullets, bottom_chip=chip, title_color=col)

    add_bottom_banner(slide, "Ecological Bottlenecks",
                      "Delayed maturity, short seed travel distances, slow soil formation, and dense canopy shade explain why tree response lags span centuries to millennia.")

    set_clean_notes(slide, 17, "Ecological Synthesis", "1 min 15s", "21:00",
        "Slide 17 explains the biology behind why trees took 1,500 to 2,500 years to catch up. There are three big bottlenecks: "
        "First, life history. Unlike garden weeds that set seed in a single summer, oak and beech trees take 15 to 40 years to produce viable seeds, bumper crops only happen every few years, and heavy nuts don't travel far. "
        "Second, soil. When glaciers melt, they leave behind bare crushed rock. It takes centuries of lichen, moss, and alder roots to build rich organic topsoil before maple and beech can grow. "
        "Third, canopy competition. An established pine or spruce stand casts dense shade for hundreds of years, physically blocking new hardwood saplings until old trees die off.",
        "Slide 18 connects these historical findings to today's climate challenges.",
        "Question: Can animal dispersal overcome seed shadow limits?\nAnswer: Blue jays and squirrels disperse nuts hundreds of meters, but landscape-scale establishment remains limited by caching mortality and seed predation.")


def slide_18(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "5. Implications & Limitations", "What Does This Mean for Today's Climate Crisis?",
                     "Translating 15,000 years of forest history into practical, real-world conservation strategies.")

    cards = [
        ("Natural Dispersal Will Fall Behind", [
            "Modern Speed Gap: Today's climate warming (1,574 m/yr) is 9.3 times faster than ancient tree migration (170 m/yr).",
            "Trees Lagging Behind: Left on their own, trees cannot keep up with shifting temperatures, leading to severe dieback along their southern borders.",
            "Out-of-Sync Ecosystems: Forests will spend decades or centuries out of equilibrium with regional climate."
        ], "Natural Migration Too Slow", COLOR_VERMILION),
        ("Modern Human Obstacles", [
            "Ancient Wilderness: Postglacial trees migrated through continuous, untouched natural landscapes.",
            "Highways and Sprawl: Eastern North America is now chopped up by highways, cities, suburban sprawl, and huge farm fields.",
            "Blocked Routes: The river valleys that Hickory and Oak once used to move north are now fragmented by human development."
        ], "Migration Routes Blocked", COLOR_ORANGE),
        ("Assisted Migration Needed", [
            "Active Forestry: We cannot simply protect land and assume trees will naturally disperse in time.",
            "Planting Seeds North: Foresters need to actively plant southern tree populations 100 to 300 km north of their current ranges.",
            "Focus on Slow Movers: Priority should go to slow-dispersing climax trees like Beech, Sugar Maple, and Hemlock."
        ], "Assisted Planting Required", COLOR_GREEN)
    ]

    for idx, (title, bullets, chip, col) in enumerate(cards):
        add_card_column(slide, 0.8 + idx * 3.98, 1.55, 3.75, 4.30, title, bullets, bottom_chip=chip, title_color=col)

    add_bottom_banner(slide, "Conservation Strategy",
                      "Because modern warming is nearly 10 times faster than natural tree dispersal across a fragmented map, active assisted planting is required to protect future forests.")

    set_clean_notes(slide, 18, "Conservation Policy", "1 min 00s", "22:00",
        "Slide 18 connects our historical data to today's climate challenges. We see two huge takeaways: "
        "First, natural seed dispersal won't cut it. Ancient trees maxed out below 300 meters per year, but climate zones today are moving at over 1,500 meters per year. "
        "Second, ancient trees moved across an unbroken continent. Today, the landscape is cut up by highways, cities, and farms. "
        "That means modern trees face a faster warming rate across a landscape with more obstacles. "
        "The clear conclusion is that passive conservation isn't enough—we need proactive assisted migration, planting resilient seed sources further north to help forests keep pace.",
        "Slide 19 outlines the limitations of our analysis.",
        "Question: Are there ecological risks to assisted migration?\nAnswer: Yes, risks include introducing novel pathogens or maladapted genetics; planting should focus on assisted population migration within existing ranges first.")


def slide_19(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "5. Implications & Limitations", "Honest Limitations & Next Steps",
                     "Being clear about what our data can and cannot tell us, and where research is heading.")

    cards = [
        ("500-Year Time Bins", [
            "Long Time Slices: 500-year bins smooth out short-term climate events, capturing millennial trends rather than decade-scale droughts.",
            "Missing Extreme Years: Cannot tell if single severe drought years triggered sudden die-off pulses.",
            "Next Step: Annual laminated lake mud cores (varves) and cave stalagmite records to look at year-by-year patterns."
        ], "500-Year Window Resolution", COLOR_BLUE),
        ("2D Planar Surfaces", [
            "Regional Overview: Planar surfaces capture large regional migration waves across the map.",
            "Local Microclimates: Smooths out small local shelters, like deep shady ravines or cool north-facing slopes.",
            "Next Step: Flexible non-linear spatial models (GAMs) that can capture micro-topographic detours."
        ], "Regional Fronts Captured", COLOR_VERMILION),
        ("Pollen Identifies Genera, Not Species", [
            "Genus Limits: Pollen grains under a microscope look nearly identical within a genus (e.g., all Oak pollen looks similar).",
            "Hidden Species Differences: Red Oak and White Oak might have responded with different individual speeds.",
            "Next Step: Ancient environmental DNA (sedaDNA) extracted from lake mud to pinpoint exact tree species."
        ], "Genus-Level Identification", COLOR_GREEN)
    ]

    for idx, (title, bullets, chip, col) in enumerate(cards):
        add_card_column(slide, 0.8 + idx * 3.98, 1.55, 3.75, 4.30, title, bullets, bottom_chip=chip, title_color=col)

    add_bottom_banner(slide, "Future Directions",
                      "Recognizing 500-year time bins, regional planar smoothing, and genus-level pollen limits highlights exciting future work with annual lake cores, non-linear maps, and ancient DNA.")

    set_clean_notes(slide, 19, "Limitations & Horizons", "1 min 00s", "23:00",
        "Slide 19 outlines our limitations. "
        "First, our 500-year time bins capture long-term trends, but they smooth over short-term events like decade-long droughts. "
        "Second, our 2D planar arrival surfaces capture broad regional waves, but they smooth out microclimates in steep mountain gorges. "
        "Third, fossil pollen only identifies trees to the genus level—we can identify Oak, but we can't tell Red Oak from White Oak. "
        "Future work using ancient DNA from lake sediments will help resolve individual tree species.",
        "Slide 20 summarizes our core findings.",
        "Question: How does sedaDNA solve pollen limitations?\nAnswer: Ancient environmental DNA directly extracts chloroplast and nuclear DNA from lake mud, providing species-level and population-level taxonomic identification.")


def slide_20(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "5. Conclusions & Takeaways", "Summary: Three Big Discoveries from 15,000 Years",
                     "Synthesizing 69 lake cores, TraCE-21ka climate, and 2D mapping into three main takeaways.")

    takeaways = [
        ("1. Trees Run on Individual Speed Clocks (RQ1)", [
            "Independent Clocks: Forests do not move as unified biomes; tree genera follow individual speed clocks from 750 to 2,500 years.",
            "Three Speed Tiers: Fast pioneers (Tamarack) react in 750 yr; canopy hardwoods in 1,500 yr; climax seeders in 2,500 yr.",
            "Opposite Reactions: Warming triggered major dieback in cold-loving Spruce (95% retreat) alongside rapid Oak expansion (79% positive)."
        ], COLOR_BLUE, "Non-Unified Biomes"),
        ("2. Novel Forests Were a Temporary Phase (RQ2)", [
            "Peak at 38%: Extinct no-analogue tree mixtures peaked at 38.2% during rapid deglacial warming (14,500 BP).",
            "Spruce & Ash Driven: Driven mainly by abnormal co-occurrences of boreal Spruce (0.161 SCD) and wetland Black Ash (0.057 SCD).",
            "Disappeared Completely: Novel forest prevalence dropped to exactly 0.0% by 500 BP once modern climate seasonality stabilized."
        ], COLOR_VERMILION, "Temporary Disequilibrium"),
        ("3. The 9.3x Modern Velocity Tracking Deficit (RQ3)", [
            "Hard Speed Limit: Real 2D tree population speeds hit an empirical ceiling between 117 and 277 m/yr (Hickory max 296 m/yr).",
            "Slower Than Climate: Past trees moved slower than ancient warming (332 m/yr), and trail modern warming (1,574 m/yr) by 9.3 times.",
            "Active Action Needed: Passive natural dispersal will fail under 21st-century warming; proactive assisted migration is required."
        ], COLOR_GREEN, "Assisted Planting Mandate")
    ]

    for idx, (title, bullets, col, badge_text) in enumerate(takeaways):
        top_pos = 1.55 + idx * 1.42
        c = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(top_pos), Inches(11.733), Inches(1.30))
        c.fill.solid()
        c.fill.fore_color.rgb = CARD_BG
        c.line.color.rgb = CARD_BORDER
        c.line.width = Pt(1)

        st = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(top_pos), Inches(0.10), Inches(1.30))
        st.fill.solid()
        st.fill.fore_color.rgb = col
        st.line.fill.background()

        # Right KPI Badge Pill
        badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.6), Inches(top_pos + 0.12), Inches(2.75), Inches(0.36))
        badge.fill.solid()
        badge.fill.fore_color.rgb = CARD_MUTED
        badge.line.color.rgb = CARD_BORDER
        badge.line.width = Pt(0.75)
        p_bdg = badge.text_frame.paragraphs[0]
        p_bdg.text = badge_text
        p_bdg.alignment = PP_ALIGN.CENTER
        p_bdg.font.size = Pt(10)
        p_bdg.font.bold = True
        p_bdg.font.color.rgb = col
        p_bdg.font.name = "Segoe UI"

        tb = slide.shapes.add_textbox(Inches(1.05), Inches(top_pos + 0.08), Inches(8.45), Inches(1.14))
        tf = tb.text_frame
        tf.word_wrap = True
        pt = tf.paragraphs[0]
        pt.text = title
        pt.font.size = Pt(12)
        pt.font.bold = True
        pt.font.color.rgb = col
        pt.space_after = Pt(3)

        for b in bullets:
            pb = tf.add_paragraph()
            pb.font.size = Pt(10)
            pb.font.name = "Segoe UI"
            pb.space_after = Pt(2)
            if ":" in b:
                parts = b.split(":", 1)
                r1 = pb.add_run()
                r1.text = f"• {parts[0].strip()}:"
                r1.font.bold = True
                r1.font.color.rgb = COLOR_INK
                r2 = pb.add_run()
                r2.text = f" {parts[1].strip()}"
                r2.font.color.rgb = COLOR_SLATE
            else:
                pb.text = f"•  {b}"
                pb.font.color.rgb = COLOR_SLATE

    add_bottom_banner(slide, "Final Conclusion",
                      "Climate change moves in decades, but forests take centuries to reorganize. Understanding these natural historical speed limits is critical to protecting our future forests.")

    set_clean_notes(slide, 20, "Conclusions", "1 min 00s", "24:00",
        "To wrap up our presentation, here are our three core conclusions: "
        "First, trees don't migrate as neat, synchronized biomes—different species move on individual clocks ranging from 750 to 2,500 years. "
        "Second, novel extinct forest mixtures peaked at 38% during rapid warming and disappeared completely under stable pre-colonial conditions. "
        "Third, past tree migration hit a hard ceiling below 300 meters per year, leaving a 9.3-fold speed deficit compared to modern warming. "
        "Natural seed dispersal won't keep pace on its own—we need active assisted migration to safeguard our forests. "
        "Thank you for your time, and we'd be glad to take any questions!",
        "We also have two backup slides prepared covering 1D vs 2D mathematical differences and statistical parameters.",
        "Question: What is the single primary takeaway for foresters?\nAnswer: Natural tree migration cannot outrun modern climate velocity; assisted migration of southern seedlots must be planned now.")


def slide_21(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "Backup Slide (Q&A)", "Backup: Why Did Older 1D Studies Overestimate Speeds?",
                     "Comparing the math: why inverting 1D regression slopes inflated speeds versus realistic 2D gradient vectors.")

    cards = [
        ("The 1D Slope Inversion Glitch", [
            "Old 1D Formula: Modeled arrival age against latitude: Arrival Age = beta_0 + beta_1 * Latitude.",
            "Inverting the Slope: Speed was calculated as 1 / beta_1, with bounds 1 / (beta_1 ± t * SE).",
            "Dividing by Near-Zero: With small sample sizes, the lower slope bound touched near zero, causing 1/slope to explode past 600 m/yr.",
            "False Fast Speeds: Created artificial speeds that mistakenly appeared faster than past climate velocity (332 m/yr)."
        ], "Artifact: 1 / 0 -> Infinity", COLOR_VERMILION),
        ("The Straight-North Blindspot", [
            "North-Only Assumption: 1D models assumed trees only traveled due North (0°).",
            "Missing Sideways Detours: Moving along river valleys or around the Great Lakes produced little latitude change.",
            "Distorted Speeds: Sideways travel was misinterpreted as instant movement, heavily distorting speed estimates."
        ], "Sideways Detours Missed", COLOR_ORANGE),
        ("2D Equal-Area Map Reality", [
            "Planar Surface Age(x, y): Fitted Age(x, y) = beta_0 + beta_x*x + beta_y*y in equal-area coordinates (km).",
            "Real Gradient Vector: Speed = 1 / sqrt(beta_x^2 + beta_y^2) along actual migration corridors (e.g. Hickory at 65° ENE).",
            "Honest Speed Limit: Tree speeds stay bounded between 117 and 277 m/yr, consistently below past climate speed (332 m/yr)."
        ], "Realistic Speed Limit", COLOR_BLUE)
    ]

    for idx, (title, bullets, chip, col) in enumerate(cards):
        add_card_column(slide, 0.8 + idx * 3.98, 1.55, 3.75, 4.30, title, bullets, bottom_chip=chip, title_color=col)

    add_bottom_banner(slide, "Mathematical Resolution",
                      "Earlier reports of trees outrunning past climate velocity were mathematical artifacts of inverting small regression slopes. True 2D mapping reveals a strict biological ceiling under 300 m/yr.")

    set_clean_notes(slide, 21, "Backup 1D vs 2D Math", "Backup", "Buffer",
        "If anyone asks why earlier studies reported tree speeds of 500 to 600 meters per year while we found speeds under 300 meters per year, this slide explains the math. "
        "Older studies used 1D linear regression with latitude, and then inverted the slope: speed = 1 / slope. "
        "When sample sizes are modest, the lower confidence bound of that slope gets close to zero, and dividing by near-zero makes the speed explode to over 600 m/yr. "
        "Plus, 1D models assume trees moved straight north, ignoring detours around lakes. "
        "Our 2D planar method measures true gradient vectors on an equal-area map, keeping speeds physically grounded.",
        "Transition back to Q&A.")


def slide_22(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "Backup Slide (Q&A)", "Backup: Master Parameter Reference Table",
                     "A handy reference of all analytical settings, thresholds, and statistical criteria used in this study.")

    headers = ["Analytical Setting", "Value Used in Study", "Why It Was Chosen & Reference"]
    rows = [
        ["Study Corridor", "25°–50°N, 95°–65°W", "Eastern North American temperate-boreal transition zone"],
        ["Sediment Cores", "69 continuous lakes (1,706 bins)", "Min 150 grains/sample, min 4 continuous 500-yr bins"],
        ["Age Chronologies", "LegacyAge 1.0 (Li et al. 2022)", "Calibrated Bayesian radiocarbon chronologies (PANGAEA Table-S5)"],
        ["Age Uncertainty", "N = 50 Monte Carlo draws", "Resampled depth-age bounds; cross-genus median lag locked at 1,500 yr"],
        ["Pollen Productivity", "REVEALS RPP Multipliers", "Wieczorek & Herzschuh (2020) & Dawson et al. (2016)"],
        ["Taxon Filter", "Arboreal Only (TRSH-only)", "Focuses on trees & shrubs; filters post-settlement farm weeds"],
        ["No-Analogue Cutoff", "Squared Chord Distance (SCD)", "Cutoff = 0.138 (5th percentile of 611 modern surface lakes)"],
        ["Stationary Differencing", "First-differencing (Δy vs ΔT)", "Removes 15,000-year trends; tested against 1,000 Ebisuzaki surrogates"],
        ["2D Arrival Surface", "Age(x, y) = b0 + bx*x + by*y", "Albers Equal-Area (EPSG:5070); gated by planar F-test (p <= 0.05)"],
        ["Modern Climate Velocity", "1,574 m/yr (1,242–2,071 range)", "1.9°C/century warming divided by 0.00121°C/km regional spatial gradient"]
    ]

    col_w = [Inches(2.5), Inches(3.2), Inches(6.033)]
    add_table(slide, 0.8, 1.55, 11.733, 4.30, headers, rows, col_w)

    add_bottom_banner(slide, "Reproducible Research",
                      "All thresholds, settings, and formulas are fully scripted and open-source, ensuring complete end-to-end reproducibility.")

    set_clean_notes(slide, 22, "Backup Parameters", "Backup", "Buffer",
        "If the committee asks for exact calibration numbers, thresholds, or citations, this table has all the key values: "
        "1. RPP multipliers: Wieczorek & Herzschuh (2020). "
        "2. Chronologies: Li et al. (2022) LegacyAge 1.0. "
        "3. No-analogue cutoff: 0.138 Squared Chord Distance based on the 5th percentile of 611 modern surface lakes. "
        "4. Surrogate testing: 1,000 Ebisuzaki phase-randomized surrogate series. "
        "5. Modern climate velocity: 1.9°C/century divided by 0.00121°C/km, yielding 1,574 m/yr.",
        "Transition back to Q&A.")


# ==============================================================================
# MAIN EXECUTION
# ==============================================================================

def main():
    print("[ppt] initializing base presentation (16:9 widescreen)...")
    prs = create_base_deck()

    print("[ppt] building 22 unslopped slides...")
    slide_01(prs)
    slide_02(prs)
    slide_03(prs)
    slide_04(prs)
    slide_05(prs)
    slide_06(prs)
    slide_07(prs)
    slide_08(prs)
    slide_09(prs)
    slide_10(prs)
    slide_11(prs)
    slide_12(prs)
    slide_13(prs)
    slide_14(prs)
    slide_15(prs)
    slide_16(prs)
    slide_17(prs)
    slide_18(prs)
    slide_19(prs)
    slide_20(prs)
    slide_21(prs)
    slide_22(prs)

    try:
        prs.save(str(PPTX_OUT))
        print(f"[ppt] successfully saved unslopped presentation to: {PPTX_OUT}")
    except PermissionError:
        fallback = ROOT / "TBD2_Final_Project_Presentation_unslopped.pptx"
        prs.save(str(fallback))
        print(f"[warning] {PPTX_OUT.name} is currently open in PowerPoint (exclusive write lock).")
        print(f"[ppt] saved presentation to fallback: {fallback}")


if __name__ == "__main__":
    main()
