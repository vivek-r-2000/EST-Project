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
PPTX_OUT = ROOT / "TBD2_Final_Project_Presentation.pptx"

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
        ("RQ1: Response Lags", "Taxon-specific lags span 750 to 2,500 years across 3 ecological speed tiers.", COLOR_BLUE),
        ("RQ2: Novel Forests", "38.2% peak deglacial novelty collapsed to a 0.0% pre-colonial baseline.", COLOR_VERMILION),
        ("RQ3: Migration Deficit", "Real tree velocities (117–277 m/yr) trail modern warming by 9.3×.", COLOR_GREEN)
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
        "This project quantifies how eastern North American forest trees responded to 15,000 years of postglacial climate change. "
        "By integrating 69 radiocarbon-dated sediment cores from the Neotoma Database with transient supercomputer climate simulations from TraCE-21ka, "
        "we measure three concrete parameters: the duration of genus-specific response lags, the drivers of novel no-analogue ecosystems, "
        "and the physical speed of 2D tree migration fronts relative to contemporary climate velocities.",
        "We begin with the ecological context on slide 2.",
        "Question: What is the geographic boundary of the study?\nAnswer: 25°N to 50°N and 95°W to 65°W, spanning the temperate-boreal transition zone from Florida to southern Ontario.")


def slide_02(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "1. Context & Motivation", "The Deglacial Natural Experiment: Why Study Ancient Forests?",
                     "The late Quaternary warming transition is Earth's only empirical record of how long-lived trees track rapid warming.")

    cards = [
        ("The Quaternary Laboratory", [
            "Deglacial Transition (15 to 0.5 ka BP): Global temperatures rose ~9.6°C as continental ice sheets melted across eastern North America.",
            "Long-Lived Organisms: Tree generation times span 15 to 40 years; forests cannot respond instantaneously and require centuries to reorganize.",
            "Natural Empirical Ground-Truth: Provides the sole empirical baseline of natural dispersal limits under climate forcing."
        ], "15,000-Yr Natural Run", COLOR_BLUE),
        ("The Unresolved Scientific Gap", [
            "Synchrony Assumptions: Early literature often treated biomes as coherent units moving north in synchrony.",
            "Uncertain Lag Durations: Previous lag estimates lacked detrending and statistical surrogate controls.",
            "1D Migration Oversimplifications: Models forced trees along straight-north lines, omitting geographic detours around lakes and mountain ranges."
        ], "Need for 2D Quantification", COLOR_VERMILION),
        ("Modern Climate Stakes", [
            "Velocity Disconnect: Contemporary greenhouse warming is shifting climate zones at ~1,574 m/yr.",
            "Tracking Deficit: If past trees failed to track natural warming (332 m/yr), modern forests face severe extinction and disequilibrium risks.",
            "Conservation Imperative: Resolving empirical dispersal speeds informs assisted migration and reserve corridor planning."
        ], "9.3x Velocity Gap", COLOR_GREEN)
    ]

    for idx, (title, bullets, chip, col) in enumerate(cards):
        add_card_column(slide, 0.8 + idx * 3.98, 1.55, 3.75, 4.30, title, bullets, bottom_chip=chip, title_color=col)

    add_bottom_banner(slide, "Core Takeaway",
                      "Past postglacial warming provides the only real-world evidence for tree dispersal limits, establishing that trees do not track climate as unified biomes.")

    set_clean_notes(slide, 2, "Ecological Context", "1 min 15s", "02:15",
        "Slide 2 sets the foundation of this study. Tree generation times span decades, making experimental greenhouse setups impossible for testing multi-century warming responses. "
        "The deglaciation of North America between 15,000 and 500 years ago offers a natural experiment: as ice sheets melted, forests had to disperse northward. "
        "However, historical studies lacked quantitative rigor, frequently assuming that entire forest biomes moved in lockstep. "
        "Crucially, modern anthropogenic warming is shifting climate zones nearly an order of magnitude faster than natural deglaciation, making empirical historical dispersal limits essential for conservation policy.",
        "To address these gaps, we formulated three precise research questions on slide 3.",
        "Question: Why is 500 BP used as the endpoint rather than 0 BP?\nAnswer: 500 BP represents the pre-colonial baseline before European agricultural land clearing disrupted natural forest dynamics with invasive weeds.")


def slide_03(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "1. Context & Motivation", "Three Quantitative Research Questions",
                     "Evaluating response lag duration, community novelty drivers, and 2D landscape migration rates.")

    rqs = [
        ("RQ1: Taxon Response Clocks", "How long do distinct tree genera lag behind climate shifts?", [
            "Stationary Differencing: Quantify genus-level response delays behind TraCE-21ka temperature shifts using first-differencing.",
            "Speed Hierarchies: Test whether genera exhibit distinct lag tiers (pioneers vs. canopy reorganization vs. climax seeders).",
            "Response Polarity: Separate biological response polarity: cold-tracking retreat (-r) versus warm-tracking expansion (+r)."
        ], COLOR_BLUE, "750–2,500 yr Lags"),
        ("RQ2: No-Analogue Emergence", "What caused novel plant communities with no modern counterparts?", [
            "Dissimilarity Metric: Calculate Squared Chord Distance (SCD) dissimilarity against 611 modern surface sample baselines.",
            "Novelty Dissolution: Establish whether novel communities peaked during rapid warming and collapsed pre-colonially to 0%.",
            "Taxonomic Drivers: Decompose additive SCD contributions to isolate individual taxonomic drivers (e.g. Spruce vs. Ash)."
        ], COLOR_VERMILION, "38.2% Novelty Peak"),
        ("RQ3: 2D Dispersal Velocity", "Did postglacial migration rates keep pace with climate velocity?", [
            "Planar Arrival Surfaces: Fit 2D surfaces Age(x, y) in equal-area Albers coordinates to resolve true directional vectors.",
            "Velocity Benchmark: Compare empirical tree dispersal velocities (1 / |∇Age|) against past deglacial climate speed (332 m/yr).",
            "Contemporary Deficit: Evaluate the contemporary tracking deficit against projected 21st-century warming velocity (1,574 m/yr)."
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
                      "These three questions form a cohesive pipeline: RQ1 establishes timing, RQ2 identifies community restructuring, and RQ3 measures physical velocity across the landscape.")

    set_clean_notes(slide, 3, "Research Questions", "1 min 15s", "03:30",
        "Slide 3 outlines our three research questions. "
        "First, RQ1 asks: How many centuries did tree genera lag behind temperature changes, and did they expand or collapse? "
        "Second, RQ2 examines no-analogue forests: When did combinations of trees with no modern counterpart appear, and what species caused them? "
        "Third, RQ3 measures physical velocity: When trees dispersed across a 2D map, how fast did they move relative to shifting climate zones? "
        "These questions directly test whether forests can naturally keep pace with warming.",
        "Slide 4 presents our geographic dataset and core network.",
        "Question: Why is 2D planar modeling superior to 1D?\nAnswer: 1D models assume due-north travel and invert linear slopes, which artificially inflates speeds. 2D modeling measures true ground gradients along actual migration corridors.")


def slide_04(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "2. Data & Study Corridor", "Dataset Harmonization: 69 Continuous Sediment Cores",
                     "Synthesizing 15,000 years of fossil pollen across Eastern North America from the Neotoma Database.")

    bullets = [
        "Geographic Domain (25°–50°N, 95°–65°W): Covers the Eastern North American temperate-boreal transition zone.",
        "Strict Site Filtering: Retained 69 continuous lakes with >=150 pollen grains/sample and >=4 continuous 500-yr bins.",
        "Temporal Windows: 1,706 harmonized 500-year time bins spanning from 14.75 ka BP to 0.25 ka BP.",
        "Modern Reference Network: 611 quality-filtered modern surface samples provide the baseline for no-analogue detection."
    ]

    add_figure_slide_layout(slide, "fig1_study_sites.png", 1.68,
                            "Fossil Site Distribution & Quality Criteria", bullets,
                            "69 Lakes / 1,706 Bins", "Quality-Filtered Dataset",
                            "Geographic Foundation",
                            "The 69 retained cores provide continuous, gap-free sediment records, harmonized with 611 modern reference sites across the Eastern North American corridor.",
                            COLOR_BLUE)

    set_clean_notes(slide, 4, "Study Sites & Data", "1 min 15s", "04:45",
        "Figure 1 on slide 4 shows our fossil core network. Each marker represents a continuous sediment core drilled from a lake basin. "
        "Pollen grains deposited in lake mud provide an unbroken record of past vegetation. "
        "We filtered the Neotoma Database strictly: only sites with at least 150 identified pollen grains per level and continuous coverage were included, yielding 69 high-quality lakes and 1,706 500-year bins. "
        "Blue circles indicate 57 cores directly harmonized to Li et al. (2022) LegacyAge 1.0 Bayesian chronologies, while amber circles mark 12 cores with Neotoma depth-scaled uncertainty fallback.",
        "Next, slide 5 introduces the independent climate model used to track temperature.",
        "Question: How were modern surface samples collected?\nAnswer: Modern samples consist of top-centimeter surface mud from 611 lakes across North America, reflecting pre-industrial forest composition.")


def slide_05(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "2. Data & Study Corridor", "Independent Paleoclimate: The TraCE-21ka Simulation",
                     "Using an independent physics-based climate simulation avoids the circularity of pollen-derived climate proxies.")

    cards = [
        ("Circularity Avoidance", [
            "The Methodological Trap: Many studies infer paleotemperatures from fossil pollen, and then test whether pollen lagged behind that temperature.",
            "Independent Physics Model: TraCE-21ka is a fully coupled ocean-atmosphere GCM (NCAR CCSM3) driven by orbital, greenhouse gas, and ice-sheet boundary conditions.",
            "Decoupled Variables: Ensures vegetation responses are tested against completely independent physical climate drivers."
        ], "Zero Proxy Circularity", COLOR_BLUE),
        ("Extracted Climate Signals", [
            "Decadal Surface Temperature (TREFHT): Tracked in °C across all 69 site coordinates through every 500-yr bin.",
            "Precipitation (PRECT): Annual precipitation rate (mm/day) to monitor moisture constraints alongside thermal shifts.",
            "Spatial Thermal Gradient: 0.00121 °C/km across the study corridor, meaning temperature drops ~1.2°C per 1,000 km north."
        ], "0.00121 °C/km Gradient", COLOR_VERMILION),
        ("Abrupt Events Resolved", [
            "Bølling-Allerød Warm Period (14.7–12.9 ka BP): Rapid thermal surge melting Laurentide ice.",
            "Younger Dryas Cold Snap (12.9–11.7 ka BP): Severe cooling reversing early deglacial warming.",
            "8.2 ka Cooling Pulse: Meltwater drainage into the North Atlantic chilling temperatures for centuries."
        ], "Millennial Shifts Tracked", COLOR_GREEN)
    ]

    for idx, (title, bullets, chip, col) in enumerate(cards):
        add_card_column(slide, 0.8 + idx * 3.98, 1.55, 3.75, 4.30, title, bullets, bottom_chip=chip, title_color=col)

    add_bottom_banner(slide, "Climate Independence",
                      "TraCE-21ka provides physically independent temperature (°C) and precipitation (mm/day) time series, avoiding proxy circularity and establishing exact thermal gradients.")

    set_clean_notes(slide, 5, "TraCE-21ka Paleoclimate", "1 min 15s", "06:00",
        "Slide 5 outlines our climate dataset. A pervasive flaw in older literature is circular reasoning: using pollen to reconstruct temperature, and then using that temperature to explain pollen changes. "
        "To avoid this, we extracted temperature and precipitation from TraCE-21ka, an independent transient climate simulation run on NCAR supercomputers. "
        "TraCE models orbital variations, greenhouse gases, and ice-sheet albedo from first physics principles. "
        "Importantly, TraCE establishes our regional thermal gradient at 0.00121°C per kilometer, which is the denominator needed to calculate climate velocity.",
        "Slide 6 covers the biological and chronological corrections applied to raw pollen counts.",
        "Question: Does TraCE capture regional topography?\nAnswer: Yes, TraCE incorporates changing Laurentide ice sheet elevation and coastal boundaries at decadal resolution throughout the 22,000-year run.")


def slide_06(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "2. Data & Study Corridor", "Data Hygiene: Pollen Productivity, Weeds & Chronologies",
                     "Correcting pollen overproduction, agricultural weed contamination, and radiocarbon dating errors.")

    cards = [
        ("RPP Calibration (REVEALS)", [
            "Pollen Productivity Disparity: Pinus and Betula produce massive airborne pollen clouds; Acer and Larix produce sparse, insect-dispersed grains.",
            "Empirical Multipliers: Applied REVEALS Relative Pollen Productivity coefficients (Wieczorek & Herzschuh 2020) to convert pollen percentages into true ground tree abundances.",
            "Correction Impact: Down-weighted over-represented conifers and recovered suppressed deciduous hardwoods."
        ], "RPP-Adjusted Proportions", COLOR_BLUE),
        ("TRSH-Only Taxon Isolation", [
            "Post-Settlement Weed Spike: European clearance post-1700 CE caused an explosion of agricultural herbs (Ambrosia, Poaceae).",
            "Artificial Novelty Elimination: Including weeds causes late-Holocene samples to falsely appear 'no-analogue' due to farm fields.",
            "Methodological Action: Isolated strictly arboreal woody taxa (TRSH-only), revealing true pre-colonial baseline forest states."
        ], "Weed Artifacts Eliminated", COLOR_VERMILION),
        ("LegacyAge 1.0 Chronologies", [
            "Radiocarbon Model Standardization: Replaced uncalibrated core depths with Li et al. (2022) LegacyAge 1.0 Bayesian age-depth models.",
            "Age Uncertainty Ensembles: 50-member Monte Carlo depth-age resampling to verify that age jitter does not disrupt lag hierarchies.",
            "Core Continuity: Closed interior sedimentary gaps and standardized calendar ages across all 69 basins."
        ], "Bayesian Radiocarbon Dates", COLOR_GREEN)
    ]

    for idx, (title, bullets, chip, col) in enumerate(cards):
        add_card_column(slide, 0.8 + idx * 3.98, 1.55, 3.75, 4.30, title, bullets, bottom_chip=chip, title_color=col)

    add_bottom_banner(slide, "Data Correction Impact",
                      "Applying REVEALS RPP multipliers, TRSH-only woody isolation, and LegacyAge 1.0 chronologies converts noisy pollen counts into accurate, physically grounded forest abundance time series.")

    set_clean_notes(slide, 6, "Data Corrections", "1 min 15s", "07:15",
        "Slide 6 details three essential data corrections. "
        "First, pollen grains do not map 1-to-1 to trees on the ground. Pine trees produce huge clouds of light, winged pollen, while maples produce very little. "
        "We applied REVEALS Relative Pollen Productivity multipliers from Wieczorek & Herzschuh (2020) to adjust for these biological differences. "
        "Second, post-colonial farming created massive spikes of ragweed and agricultural herbs. If left in, pre-industrial forests falsely appear novel. We isolated strictly woody arboreal taxa (TRSH-only). "
        "Third, we harmonized all radiocarbon dates using LegacyAge 1.0 Bayesian age models, backed by 50-member Monte Carlo ensembles.",
        "Slide 7 explains our statistical tests for significance.",
        "Question: How much does pine pollen over-represent pine trees?\nAnswer: Pine produces roughly 10 times more pollen per unit basal area than sugar maple, which heavily skews raw pollen diagrams if unadjusted.")


def slide_07(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "2. Data & Study Corridor", "Statistical Rigor: Detrending, Ebisuzaki & 2D Surfaces",
                     "Eliminating spurious correlations and moving from 1D straight lines to 2D equal-area surfaces.")

    cards = [
        ("Stationary First-Differencing", [
            "The Spurious Trend Trap: Both temperature and hardwood abundance trended upward over the 15,000-year deglaciation.",
            "Mathematical Fix: First-differenced time series (Δy_t vs. ΔT_t-k) evaluate bin-to-bin changes rather than raw levels.",
            "Core Analytical Question: Evaluates whether an abrupt temperature jump in bin t-k caused a corresponding vegetation jump in bin t."
        ], "Trends Removed (Δy vs ΔT)", COLOR_BLUE),
        ("1,000 Ebisuzaki Surrogates", [
            "Autocorrelation Correction: Environmental series possess red-noise memory, inflating standard p-values.",
            "Phase-Randomization: Generated 1,000 synthetic climate series preserving exact spectral power and autocovariance.",
            "Significance Benchmark: Lags are accepted only if empirical correlation exceeds the 95th percentile of surrogate tests (p < 0.05)."
        ], "Non-Parametric Null Tests", COLOR_VERMILION),
        ("2D Equal-Area Migration", [
            "Overcoming 1D North Limits: Replaced 1D latitude regressions with 2D planar arrival surfaces Age(x, y) in EPSG:5070.",
            "Vector Velocity: Speed = 1 / |∇Age| along the actual colonization front, resolving directional azimuths.",
            "Convex Hull Masking: Isochrones are clipped strictly to observed core extents, preventing false map extrapolations."
        ], "True 2D Spatial Gradients", COLOR_GREEN)
    ]

    for idx, (title, bullets, chip, col) in enumerate(cards):
        add_card_column(slide, 0.8 + idx * 3.98, 1.55, 3.75, 4.30, title, bullets, bottom_chip=chip, title_color=col)

    add_bottom_banner(slide, "Statistical Verification",
                      "First-differenced detrending eliminates spurious trends, 1,000 Ebisuzaki surrogates correct for autocorrelation, and 2D planar surfaces calculate true directional migration speeds.")

    set_clean_notes(slide, 7, "Statistical Methods", "1 min 15s", "08:30",
        "Slide 7 explains our statistical framework. "
        "A common error in paleo-studies is correlating trending time series: because both temperature and oak increased over 15,000 years, standard OLS produces high correlations that are completely spurious. "
        "We eliminated this by first-differencing the series: examining change-versus-change. "
        "Furthermore, environmental time series have temporal autocorrelation. We tested every lag correlation against 1,000 phase-randomized Ebisuzaki surrogate time series. "
        "Finally, we transitioned from 1D latitude models to 2D equal-area planar arrival surfaces to calculate real vector speeds.",
        "We now turn to our empirical findings, beginning with climate and vegetation turnover on slide 8.",
        "Question: How does the Ebisuzaki method generate surrogates?\nAnswer: It computes the discrete Fourier transform of the climate series, randomizes the Fourier phase angles while preserving amplitudes, and inverts back to time domain, maintaining exact spectral properties.")


def slide_08(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "3. Results: Climate & Turnover", "15,000-Year Climate Trajectory & Community Turnover",
                     "Tracking temperature, precipitation, warming rates, and vegetation turnover across abrupt events.")

    bullets = [
        "Deglacial Warming (+9.6°C): Regional mean annual temperature increased from -1.1°C at 15 ka BP to +8.5°C during the Holocene thermal optimum.",
        "Precipitation Increase (+1.2 mm/day): Moisture increased from dry glacial conditions (2.0 mm/day) to wet mid-Holocene conditions (3.2 mm/day).",
        "Peak Warming Rate (+0.43°C/century): Reached at 14.5 ka BP during the Bølling-Allerød warming surge.",
        "Coupled Vegetation Turnover: Hellinger turnover peaked synchronously with maximum warming rates, demonstrating that rapid climate shifts directly forced community reorganization."
    ]

    add_figure_slide_layout(slide, "fig2_climate_and_turnover.png", 1.38,
                            "Abrupt Climate Shifts & Turnover Coupling", bullets,
                            "+9.6 °C Total Warming", "Peak Rate: 0.43 °C/century",
                            "Turnover Coupling",
                            "Fossil community turnover peaked synchronously with maximum climate warming rates during the Bølling-Allerød, proving that climate shocks directly forced community upheaval.",
                            COLOR_VERMILION)

    set_clean_notes(slide, 8, "Climate & Turnover", "1 min 15s", "09:45",
        "Figure 2 on slide 8 displays the regional climate and vegetation history. "
        "The top panel tracks temperature in orange and precipitation in blue. Over 15,000 years, the region warmed by 9.6°C, accompanied by a 60% increase in precipitation. "
        "The yellow and blue vertical bars highlight the abrupt Bølling-Allerød warming and Younger Dryas cold reversal. "
        "The third panel tracks warming rate, peaking at 0.43°C per century during deglaciation. "
        "Notice the bottom panel: community turnover peaks in lockstep with the warming rate. When climate shifted rapidly, forests underwent intense compositional restructuring.",
        "We now address Research Question 1 on slide 9: How long did trees lag behind these temperature shifts?",
        "Question: Why does turnover drop during the Holocene?\nAnswer: Mid-Holocene climate stabilized, allowing tree populations to establish quasi-equilibrium communities with lower species turnover.")


def slide_09(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "3. Results: RQ1 (Response Lags)", "RQ1: Trees Operate on Three Distinct Speed Clocks",
                     "Forests do not move as a single synchronized unit; tree genera lag by 750 to 2,500 years across three speed tiers.")

    bullets = [
        "Significant Differentiation (p = 0.0003): Kruskal-Wallis testing confirms robust genus lag differences (H = 34.6).",
        "Fast Pioneer Tier (<1,000 yr): Tamarack (Larix) responded earliest, with a median lag of 750 years [0–2,500 yr CI].",
        "Canopy Reorganization Tier (1,000–2,000 yr): Major hardwoods—Oak (Quercus), Birch (Betula), Ash (Fraxinus), Elm (Ulmus), and Maple (Acer)—clustered at 1,500 years.",
        "Slow Climax Seeders (>2,000 yr): Hemlock (Tsuga) and Beech (Fagus) exhibited extreme lags of 2,500 years [2,000–3,000 yr CI]."
    ]

    add_figure_slide_layout(slide, "fig3_genus_lags_raincloud.png", 1.82,
                            "Three Distinct Ecological Speed Tiers", bullets,
                            "750 to 2,500 yr Lags", "Kruskal-Wallis p = 0.0003",
                            "Ecological Speed Hierarchy",
                            "Tree genera separate into three distinct response tiers: light-seeded pioneers react in 750 yr, broadleaf canopy trees reorganize in 1,500 yr, and heavy-seeded climax trees lag by 2,500 yr.",
                            COLOR_BLUE)

    set_clean_notes(slide, 9, "RQ1 Response Lags", "1 min 15s", "11:00",
        "Figure 3 Panel A on slide 9 presents our core finding for Research Question 1. "
        "Each row shows a tree genus; points indicate median response lag in years, and horizontal lines represent 95% confidence intervals. "
        "Kruskal-Wallis testing strongly rejects the hypothesis of equal lag times across taxa, with p = 0.0003. "
        "Instead, taxa separate into three distinct speed tiers: "
        "First, fast pioneers like Tamarack (Larix) respond in 750 years. "
        "Second, canopy hardwoods—Oak, Birch, Ash, Elm, and Maple—cluster at 1,500 years. "
        "Third, climax seeders—Beech and Hemlock—lag by 2,500 years. "
        "This proves that forests do not migrate as unified biomes; each genus operates on an independent biological clock.",
        "Slide 10 examines whether these lags represented population expansion or collapse.",
        "Question: Why does Larix have a wider confidence interval?\nAnswer: Larix was restricted to northern sites near the ice margin, resulting in a smaller sample size (n=10) than widespread oaks (n=67).")


def slide_10(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "3. Results: RQ1 (Response Lags)", "RQ1: Response Polarity: Boreal Collapse vs. Hardwood Expansion",
                     "Separating cold-tracking boreal dieback from warm-tracking temperate hardwood colonization.")

    bullets = [
        "Boreal Spruce Collapse (94.9% Negative): Picea exhibited negative correlation with temperature across 94.9% of sites, with 78% passing strict Ebisuzaki surrogate tests (p < 0.05).",
        "Temperate Hardwood Expansion (79.1% Positive): Quercus was 79.1% positive-tracking, expanding rapidly as deglacial warming opened new northern territory.",
        "Canopy Reorganization (Acer 67.2% Positive): Maple and Birch expanded behind oak along deglaciated corridors.",
        "The Non-Equilibrium State: Even warm-expanding taxa lagged temperature by 15 centuries due to seed dispersal and soil maturation limits."
    ]

    add_figure_slide_layout(slide, "fig3_genus_lags_raincloud.png", 1.82,
                            "Biological Direction: Retreat vs. Colonization", bullets,
                            "95% Spruce Retreat", "79% Oak Expansion",
                            "Polarity Dynamics",
                            "Deglacial warming triggered catastrophic dieback in cold-adapted boreal spruce (95% negative) alongside aggressive northward colonization by temperate hardwoods (79% positive).",
                            COLOR_VERMILION)

    set_clean_notes(slide, 10, "RQ1 Response Polarity", "1 min 15s", "12:15",
        "Slide 10 presents Panel B of Figure 3, resolving the biological direction of change. "
        "Blue bars indicate negative correlation with warming, while orange and red bars indicate positive correlation. "
        "Look at Spruce (Picea): 95% of lake records show negative correlation, and 78% of sites pass strict Ebisuzaki surrogate tests. Warming caused widespread dieback of cold-adapted spruce across the midwest. "
        "Conversely, Oak (Quercus) shows 79% positive correlation, expanding northward as ice retreated. "
        "Importantly, even warm-adapted oaks lagged warming by 1,500 years, demonstrating persistent disequilibrium.",
        "Next, slide 11 explores Research Question 2: the emergence of novel ecosystems during this transition.",
        "Question: Why did spruce decline so consistently across lakes?\nAnswer: Rising summer temperatures exceeded the thermal physiological limit of white and black spruce, increasing respiration costs and pest vulnerability.")


def slide_11(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "3. Results: RQ2 (No-Analogue Forests)", "RQ2: Novel Communities Peaked at 38.2% and Collapsed to 0%",
                     "No-analogue forest prevalence tracked peak deglacial warming rates and reached zero before colonial disturbance.")

    bullets = [
        "Squared Chord Distance Metric: Fossil samples exceeding the modern 5th-percentile threshold (0.138 SCD across 611 surface sites) are classified as no-analogue.",
        "Peak Deglacial Novelty (38.2%): Novel communities peaked at 14.5 ka BP during the Bølling-Allerød warming surge.",
        "Complete Pre-Colonial Dissolution (0.0%): Under TRSH-only woody filtering, no-analogue prevalence fell to exactly 0.0% by 500 BP.",
        "Coupling to Vegetation Turnover: Novelty correlated significantly with community turnover rate (Spearman rho = 0.743, Ebisuzaki p = 0.021)."
    ]

    add_figure_slide_layout(slide, "fig4_noanalogue_drivers.png", 1.47,
                            "Deglacial Novelty Trajectory & Correlation", bullets,
                            "38.2% Peak Novelty", "0.0% Pre-Colonial Baseline",
                            "Novelty Dissolution",
                            "Novel ecosystems peaked during abrupt deglacial climate turnover and dissolved completely to 0.0% by 500 BP once modern climate seasonality stabilized.",
                            COLOR_VERMILION)

    set_clean_notes(slide, 11, "RQ2 No-Analogue Dynamics", "1 min 15s", "13:30",
        "Slide 11 addresses Research Question 2: What caused no-analogue forests? "
        "We measured community novelty using Squared Chord Distance against 611 modern reference forests. "
        "The top panel of Figure 4 shows that no-analogue communities peaked at 38.2% at 14.5 ka BP—coinciding exactly with the Bølling-Allerød warming pulse. "
        "Crucially, under our clean woody filtering, no-analogue frequency dropped to exactly 0.0% by 500 BP. "
        "Previous studies that reported persistent novelty in the late Holocene were misled by post-settlement agricultural weeds.",
        "Slide 12 breaks down the specific tree species that generated this novelty.",
        "Question: What is Squared Chord Distance?\nAnswer: SCD is a non-linear metric of ecological dissimilarity; values above 0.138 indicate community compositions with no modern analogue in the reference database.")


def slide_12(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "3. Results: RQ2 (No-Analogue Forests)", "RQ2: Spruce & Black Ash Drove Community Novelty",
                     "Additive Squared Chord Distance attribution reveals the taxonomic drivers of extinct forest mixtures.")

    bullets = [
        "Additive SCD Decomposition: The Squared Chord Distance metric can be partitioned taxon-by-taxon to calculate exact novelty contributions.",
        "Primary Driver (Picea, 0.161 Mean SCD): Spruce contributed nearly half (48%) of the total community novelty during early deglaciation.",
        "Secondary Driver (Fraxinus, 0.057 Mean SCD): Black ash contributed 17% of novelty, co-occurring abnormally with spruce.",
        "Ecological Mechanism: Enhanced seasonal insolation created high summer temperatures alongside freezing winters, forcing cold-tolerant spruce and temperate wetland ash into unique co-occurrence."
    ]

    add_figure_slide_layout(slide, "fig4_noanalogue_drivers.png", 1.47,
                            "Taxonomic Dissimilarity Attribution", bullets,
                            "Picea SCD: 0.161", "Fraxinus SCD: 0.057",
                            "Taxonomic Drivers",
                            "Decomposing SCD demonstrates that novel forests were driven by anomalous co-occurrences of boreal spruce and temperate black ash under high seasonal temperature contrast.",
                            COLOR_BLUE)

    set_clean_notes(slide, 12, "RQ2 Taxonomic Drivers", "1 min 15s", "14:45",
        "Figure 4 Bottom Panel on slide 12 breaks down the taxonomic drivers of no-analogue communities. "
        "Because Squared Chord Distance is additive across taxa, we decomposed the total novelty score into genus-specific contributions. "
        "The top row of the heatmap shows that Spruce (Picea) was the dominant driver, contributing a mean SCD of 0.161. Black Ash (Fraxinus) was second at 0.057. "
        "These taxa generated novel communities because high orbital seasonality 14,000 years ago produced hot summers and frigid winters, allowing boreal spruce and wetland ash to grow in the same stands. "
        "Once seasonality moderated, this unique combination dissolved.",
        "Slide 13 transitions to Research Question 3: mapping 2D tree migration fronts.",
        "Question: Why didn't oak or pine drive novelty?\nAnswer: Oak and pine maintained abundances that closely resembled modern southeastern pine-oak woodlands, contributing minimal dissimilarity.")


def slide_13(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "3. Results: RQ3 (2D Migration)", "RQ3: Mapping 15,000 Years of Tree Travel in 2D",
                     "Replacing straight-north 1D assumptions with equal-area planar arrival surfaces and directional vectors.")

    bullets = [
        "Planar 2D F-Test Gating: Tested all 12 genera on a 2D arrival surface Age(x, y); strictly required p <= 0.05 to avoid hallucinating waves.",
        "5 Statistically Validated Migrants: Hickory (Carya), Alder (Alnus), Beech (Fagus), Tamarack (Larix), and Hemlock (Tsuga) exhibited coherent regional dispersal waves.",
        "Convex Hull Boundaries: Isochrones are masked strictly to observed lake core locations, preventing unconstrained map extrapolation.",
        "Directional Azimuth Vectors: Black arrows reveal distinct colonization routes guided by river valleys and mountain ranges."
    ]

    add_figure_slide_layout(slide, "fig5_migration_isochrones.png", 1.44,
                            "Empirical 2D Colonization Isochrones", bullets,
                            "5 Validated Migrants", "p_2d <= 0.05 Significance",
                            "2D Migration Architecture",
                            "Gating by 2D planar F-tests identified five coherent postglacial migrants, revealing distinct topographic dispersal routes around the Great Lakes and Appalachian spine.",
                            COLOR_GREEN)

    set_clean_notes(slide, 13, "RQ3 2D Arrival Isochrones", "1 min 15s", "16:00",
        "Figure 5 on slide 13 presents our 2D migration surfaces, answering Research Question 3. "
        "Earlier studies modeled migration in 1D along latitude lines. We fitted 2D planar arrival surfaces Age(x, y) in equal-area coordinates. "
        "Crucially, we applied an F-test: only genera with statistically significant directional surfaces (p <= 0.05) were mapped. "
        "Five genera passed: Hickory, Alder, Beech, Tamarack, and Hemlock. The remaining seven taxa exhibited diffuse or multi-refugial distributions. "
        "Contour colors indicate arrival ages from 12 ka (dark blue) to 4 ka (yellow), with arrows indicating local migration headings.",
        "Slide 14 details the specific topographic corridors these five taxa followed.",
        "Question: Why are isochrones clipped to convex hulls?\nAnswer: To prevent mathematical extrapolation into regions where a tree was never recorded, ensuring map integrity.")


def slide_14(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "3. Results: RQ3 (2D Migration)", "RQ3: Real Colonization Corridors and Dispersal Headings",
                     "Tree dispersal routes were channeled by river systems, lake basins, and mountain topography.")

    headers = ["Genus", "Common Name", "Median Speed", "Full Surface Range", "Azimuth Heading", "Key Dispersal Corridor"]
    rows = [
        ["Carya", "Hickory", "277.0 m/yr", "217–296 m/yr", "65.0° (ENE)", "Mississippi/Ohio valleys toward Atlantic coast"],
        ["Alnus", "Alder", "264.6 m/yr", "254–278 m/yr", "308.2° (WNW)", "New England across Great Lakes into Minnesota"],
        ["Fagus", "Beech", "169.6 m/yr", "157–185 m/yr", "344.4° (NNW)", "Southern Appalachians north into New England"],
        ["Larix", "Tamarack", "124.0 m/yr", "115–129 m/yr", "1.2° (Due North)", "Tracked retreating Laurentide ice across Midwest"],
        ["Tsuga", "Hemlock", "116.9 m/yr", "112–120 m/yr", "344.3° (NNW)", "Parallel Appalachian spine into Great Lakes"]
    ]

    col_w = [Inches(1.6), Inches(1.5), Inches(1.4), Inches(1.8), Inches(1.6), Inches(3.833)]
    add_table(slide, 0.8, 1.55, 11.733, 4.30, headers, rows, col_w)

    add_bottom_banner(slide, "Corridor Diversity",
                      "Trees did not march due North: Hickory migrated East-North-East (65°) along river corridors, Alder tracked West-North-West (308°), and only Tamarack moved straight North (1°).")

    set_clean_notes(slide, 14, "RQ3 Dispersal Corridors", "1 min 15s", "17:15",
        "Slide 14 quantifies the five migration corridors. "
        "Notice the diversity of compass headings: Hickory (Carya) was our fastest migrant at 277 m/yr, heading East-North-East at 65° along fertile river valleys. "
        "Alder (Alnus) moved at 265 m/yr West-North-West at 308°, colonizing lake margins around the Great Lakes. "
        "Beech and Hemlock moved at 170 and 117 m/yr along parallel North-North-West tracks (344°) along the Appalachian spine. "
        "Tamarack was the only genus moving due North (1.2°), tracking the ice margin. "
        "This proves that 1D north-only models distort the true geometry of postglacial migration.",
        "Slide 15 compares these migration speeds to climate tracking velocities.",
        "Question: Why did Hickory move East-North-East?\nA: Glacial refugia for hickory were concentrated in the lower Mississippi basin; postglacial warming opened river corridors extending toward the Atlantic coast.")


def slide_15(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "3. Results: RQ3 (2D Migration)", "RQ3: The 9.3x Modern Climate Velocity Deficit",
                     "Real tree migration hit a biological ceiling below 300 m/yr—falling nearly an order of magnitude behind modern warming.")

    bullets = [
        "A Hard Biological Ceiling (117–277 m/yr): Across all 2D arrival surfaces, tree speeds reached an empirical ceiling at 296 m/yr (Carya max).",
        "Slower Than Past Deglaciation: Every genus migrated strictly slower than past abrupt deglacial climate velocity (p95 = 332 m/yr, blue dashed line).",
        "Modern Climate Velocity (1,574 m/yr): 21st-century warming (1.9°C/century) divided by regional gradient (0.00121°C/km) requires 1,574 m/yr tracking.",
        "The 9.3x Tracking Deficit: Median tree migration (170 m/yr) falls nearly an order of magnitude short of contemporary warming velocity."
    ]

    add_figure_slide_layout(slide, "fig6_velocity_deficit.png", 2.04,
                            "Log-Scale Velocity Comparison", bullets,
                            "9.3x Velocity Deficit", "170 m/yr vs. 1,574 m/yr",
                            "Velocity Deficit Punchline",
                            "Tree migration speeds (117–277 m/yr) were strictly slower than past deglacial climate velocity (332 m/yr) and trail modern anthropogenic warming (1,574 m/yr) by 9.3×.",
                            COLOR_VERMILION)

    set_clean_notes(slide, 15, "RQ3 Velocity Deficit", "1 min 15s", "18:30",
        "Figure 6 on slide 15 presents the central punchline of our study. "
        "The horizontal axis displays speed in meters per year on a log scale. Green markers show 2D tree migration medians and full surface ranges. "
        "Notice two critical reference lines: "
        "First, the blue dashed line at 332 m/yr represents the 95th percentile velocity of past natural deglacial warming. Not a single tree crossed this line. Past trees were chronically slower than rapid climate shifts. "
        "Second, look at the orange dashed line at 1,574 m/yr: this is the velocity required to track modern anthropogenic warming at 1.9°C per century. "
        "The median tree migration rate of 170 m/yr falls 9.3 times short of modern climate velocity. Trees cannot disperse fast enough naturally.",
        "Slide 16 presents sensitivity testing of our analytical models.",
        "Question: Why did earlier mid-project models cross the 332 m/yr line?\nAnswer: Mid-project 1D models inverted linear regression slopes with wide standard errors, creating mathematical CI blowups that falsely exceeded 500 m/yr.")


def slide_16(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "4. Sensitivity & Robustness", "Model Robustness: Age Uncertainty & Bin Audits",
                     "Monte Carlo age-model draws, temporal bin doubling, and surrogate testing confirm findings.")

    cards = [
        ("LegacyAge N=50 Monte Carlo", [
            "Resampling Protocol: Drew 50 random age-depth realizations from LegacyAge 1.0 Bayesian chronologies across all 69 lakes.",
            "Stability Result: Cross-genus median response lag remained locked at exactly 1,500 years across all 50 iterations.",
            "Conclusion: Proves our lag hierarchies are completely robust to radiocarbon calibration uncertainty."
        ], "Median Lag = 1,500 yr Locked", COLOR_BLUE),
        ("250-Year Binning Check", [
            "Resolution Doubling: Re-ran the complete pipeline on 250-year bins (66 cores, 3,365 binned levels).",
            "Taxonomic Stability: Reproduced the identical three-tier hierarchy—Larix fastest (500-750 yr), Canopy hardwoods intermediate, Climax seeders slowest.",
            "Conclusion: Confirms 500-year binning did not blur or distort underlying biological dynamics."
        ], "Identical Taxonomic Order", COLOR_VERMILION),
        ("Surrogate Significance Rate", [
            "Phase-Randomized Controls: Tested correlations against 1,000 synthetic climate series.",
            "Empirical Pass Rate: 78.4% of Picea retreat pairs and 67.2% of Quercus expansion pairs passed the 95% surrogate significance threshold.",
            "Conclusion: Proves climate-vegetation correlations reflect true physical responses rather than spectral noise."
        ], "78% Picea / 67% Oak Sig.", COLOR_GREEN)
    ]

    for idx, (title, bullets, chip, col) in enumerate(cards):
        add_card_column(slide, 0.8 + idx * 3.98, 1.55, 3.75, 4.30, title, bullets, bottom_chip=chip, title_color=col)

    add_bottom_banner(slide, "Validation Verdict",
                      "Three independent sensitivity stress-tests confirm that our lag rankings, novel community trajectories, and 2D velocity deficits are robust to chronological and binning choices.")

    set_clean_notes(slide, 16, "Sensitivity Audits", "1 min 15s", "19:45",
        "Slide 16 details our sensitivity audits. "
        "To ensure our conclusions were not artifacts of code settings, we performed three tests: "
        "First, a 50-member Monte Carlo simulation resampled radiocarbon dating errors; across all 50 runs, median lag remained locked at 1,500 years. "
        "Second, we doubled temporal resolution using 250-year bins across 3,365 levels; the fast-to-slow taxonomic hierarchy was identical. "
        "Third, our 1,000-surrogate Ebisuzaki tests confirmed that 78% of spruce retreat signals and 67% of oak expansion signals were statistically significant.",
        "Slide 17 explores the biological mechanisms behind these 15-century lags.",
        "Question: Did downweighting Pinus alter the lag hierarchy?\nAnswer: No, downweighting Pinus reduced overproduction distortion but preserved the identical pioneer-to-climax lag ordering.")


def slide_17(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "4. Ecological Synthesis", "Why Do Trees Lag? Three Biological Bottlenecks",
                     "Physical and biological constraints prevent instantaneous forest tracking of climate shifts.")

    cards = [
        ("Generation Time & Dispersal", [
            "Delayed Sexual Maturity: Unlike annual herbs, forest trees require 15 to 40 years before producing their first viable seed crop.",
            "Mast Seeding Periodicity: Heavy-seeded hardwoods (Beech, Oak, Hickory) produce bumper seed crops only every 3 to 7 years.",
            "Local Seed Shadow: Most heavy nuts fall within 30 meters of the parent canopy unless transported by animals."
        ], "Dispersal Ceiling <300 m/yr", COLOR_BLUE),
        ("Slow Pedogenesis (Soil Maturation)", [
            "Glacial Till Substrates: Melting ice sheets left raw gravel and crushed bedrock devoid of nitrogen and organic matter.",
            "Primary Succession: Building organic topsoil requires centuries of lichen weathering and nitrogen-fixing pioneers (Alnus).",
            "Nutrient Requirements: Climax hardwoods (Fagus, Acer) cannot establish until fungal networks and humus layers mature."
        ], "Centuries to Build Topsoil", COLOR_VERMILION),
        ("Canopy Inertia & Shading", [
            "Mature Canopy Buffering: Established trees buffer ground microclimates and cast dense shade for centuries.",
            "Competitive Exclusion: Even when regional climate favors hardwoods, existing spruce and pine canopies block hardwood seedlings.",
            "Disturbance Dependence: Canopy recruitment occurs primarily after fire, windthrow, or pathogen mortality events."
        ], "Competitive Shading Inertia", COLOR_GREEN)
    ]

    for idx, (title, bullets, chip, col) in enumerate(cards):
        add_card_column(slide, 0.8 + idx * 3.98, 1.55, 3.75, 4.30, title, bullets, bottom_chip=chip, title_color=col)

    add_bottom_banner(slide, "Biological Mechanisms",
                      "Delayed maturity, localized seed shadows, centuries of required soil pedogenesis, and mature canopy competition explain why tree lags span 750 to 2,500 years.")

    set_clean_notes(slide, 17, "Ecological Synthesis", "1 min 15s", "21:00",
        "Slide 17 synthesizes the biological mechanisms explaining our lag findings. "
        "Why do trees lag behind climate by 1,500 to 2,500 years? Three bottlenecks operate simultaneously: "
        "First, life history: trees require decades to mature sexually, produce mast crops intermittently, and disperse heavy seeds locally. "
        "Second, soil pedogenesis: retreating glaciers leave barren crushed rock; building topsoil with mycorrhizal fungi takes centuries of succession. "
        "Third, canopy competition: established adult trees shade the ground and buffer local microclimates, physically excluding new colonizers until mortality opens gaps.",
        "Slide 18 translates these findings into modern conservation implications.",
        "Question: Can animal dispersal overcome seed shadow limits?\nAnswer: Blue jays and squirrels disperse nuts hundreds of meters, but landscape-scale establishment remains limited by caching mortality and seed predation.")


def slide_18(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "5. Implications & Limitations", "Conservation Implications for 21st-Century Warming",
                     "Translating empirical paleoecological limits into actionable modern conservation policy.")

    cards = [
        ("Passive Migration Failure", [
            "Velocity Mismatch: Contemporary climate velocity (1,574 m/yr) is 9.3× faster than median historical tree dispersal (170 m/yr).",
            "Lag Accumulation: Left to natural dispersal, long-lived trees will accumulate severe tracking lags, resulting in trailing-edge dieback.",
            "Disequilibrium Ecosystems: Modern forests will enter prolonged structural disequilibrium with climate."
        ], "Natural Tracking Failure", COLOR_VERMILION),
        ("Landscape Fragmentation", [
            "Historical Wild Landscapes: Postglacial trees dispersed across continuous, unfragmented wilderness corridors.",
            "Modern Barriers: Eastern North America is severed by highways, urban sprawl, and monoculture agriculture.",
            "Corridor Severance: The East-North-East river corridor used by Hickory is now heavily fragmented by human infrastructure."
        ], "Dispersal Corridors Blocked", COLOR_ORANGE),
        ("Assisted Migration Imperative", [
            "Proactive Forestry: Conservation managers cannot rely on passive seed dispersal to preserve tree populations.",
            "Assisted Gene Flow & Migration: Actively moving southern genotypes and seedlots 100 to 300 km north of current ranges.",
            "Priority Genera: Priority must be given to slow-dispersing climax hardwoods (Beech, Sugar Maple, Hemlock)."
        ], "Assisted Migration Required", COLOR_GREEN)
    ]

    for idx, (title, bullets, chip, col) in enumerate(cards):
        add_card_column(slide, 0.8 + idx * 3.98, 1.55, 3.75, 4.30, title, bullets, bottom_chip=chip, title_color=col)

    add_bottom_banner(slide, "Management Conclusion",
                      "Because modern warming outpaces natural migration by nearly 10× across a fragmented landscape, active assisted migration is necessary to prevent widespread forest decline.")

    set_clean_notes(slide, 18, "Conservation Policy", "1 min 00s", "22:00",
        "Slide 18 addresses modern conservation policy. "
        "Our empirical results yield two unavoidable conclusions: "
        "First, natural passive migration will fail to track 21st-century warming. Tree migration maxes out below 300 m/yr, while climate zones are shifting at over 1,500 m/yr. "
        "Second, modern landscapes are heavily fragmented by cities and highways, cutting off historical river corridors. "
        "Therefore, proactive assisted migration—planting southern tree seedlots 100 to 300 km north of current boundaries—is necessary to prevent widespread forest collapse.",
        "Slide 19 outlines the limitations of our analysis.",
        "Question: Are there ecological risks to assisted migration?\nAnswer: Yes, risks include introducing novel pathogens or maladapted genetics; planting should focus on assisted population migration within existing ranges first.")


def slide_19(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "5. Implications & Limitations", "Analytical Limitations & Future Directions",
                     "Transparent boundaries regarding temporal discretization, planar models, and taxonomic resolution.")

    cards = [
        ("500-Year Binning Limits", [
            "Temporal Coarseness: 500-year binning integrates over centennial climate variability, smoothing decadal extreme events.",
            "Extreme Event Masking: Cannot resolve whether short-term multi-year mega-droughts triggered localized dieback pulses.",
            "Future Direction: High-resolution annual varved sediment cores and speleothem isotopic records."
        ], "Centennial Discretization", COLOR_BLUE),
        ("Planar Surface Smoothing", [
            "Regional Front Assumption: 2D planar arrival surfaces capture dominant regional fronts across equal-area coordinates.",
            "Microclimatic Detours: Smooths over fine-scale topographic microrefugia, such as sheltered river gorges and north-facing slopes.",
            "Future Direction: Non-linear generalized additive models (GAMs) and thin-plate splines."
        ], "Micro-Topography Smoothed", COLOR_VERMILION),
        ("Taxonomic Identification", [
            "Pollen Morphology Limits: Pollen grains can generally be identified only to genus level (e.g., Quercus, Pinus).",
            "Species-Level Heterogeneity: Red oak (Quercus rubra) and white oak (Quercus alba) may have responded differently.",
            "Future Direction: Ancient sedimentary environmental DNA (sedaDNA) to resolve exact species genotypes."
        ], "Genus-Level Resolution", COLOR_GREEN)
    ]

    for idx, (title, bullets, chip, col) in enumerate(cards):
        add_card_column(slide, 0.8 + idx * 3.98, 1.55, 3.75, 4.30, title, bullets, bottom_chip=chip, title_color=col)

    add_bottom_banner(slide, "Research Horizons",
                      "Acknowledging 500-year temporal limits, planar smoothing, and genus-level pollen constraints points toward future integration of annual varved cores, GAMs, and sedimentary ancient DNA.")

    set_clean_notes(slide, 19, "Limitations & Horizons", "1 min 00s", "23:00",
        "Slide 19 addresses analytical limitations. "
        "First, our 500-year binning resolves millennial trends but smooths decadal drought pulses. "
        "Second, planar arrival surfaces capture regional fronts but simplify micro-topographic detours in mountain valleys. "
        "Third, pollen morphology restricts identification to genus level; red and white oak cannot be distinguished visually. "
        "Future work coupling sedimentary ancient DNA and non-linear GAM surfaces will resolve species-specific genotypes.",
        "Slide 20 summarizes our core findings.",
        "Question: How does sedaDNA solve pollen limitations?\nAnswer: Ancient environmental DNA directly extracts chloroplast and nuclear DNA from lake mud, providing species-level and population-level taxonomic identification.")


def slide_20(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "5. Conclusions & Takeaways", "Summary: Three Empirical Discoveries from 15,000 Years",
                     "Synthesizing 69 lake cores, TraCE-21ka climate, and 2D spatial modeling into three core principles.")

    takeaways = [
        ("1. Distinct Biological Speed Clocks (RQ1)", [
            "Taxon Independence: Trees do not move as synchronized biomes; genera run on individual speed clocks ranging from 750 to 2,500 years.",
            "Tiered Lags: Pioneers (Larix) respond in 750 yr; canopy hardwoods (Quercus, Acer) in 1,500 yr; climax seeders (Fagus, Tsuga) in 2,500 yr.",
            "Response Polarity: Warming caused severe boreal dieback in Spruce (95% -r) alongside aggressive expansion in Oak (79% +r)."
        ], COLOR_BLUE, "Non-Unified Biomes"),
        ("2. Novel Ecosystems Peak During Rapid Shifts (RQ2)", [
            "Transient Peak (38.2%): No-analogue forest prevalence peaked at 14.5 ka BP during peak Bølling-Allerød warming turnover.",
            "Dominant Drivers: Driven primarily by anomalous co-occurrences of boreal Spruce (0.161 SCD) and wetland Black Ash (0.057 SCD).",
            "Pre-Colonial Dissolution: Novel forest prevalence dropped to exactly 0.0% by 500 BP once modern climate seasonality stabilized."
        ], COLOR_VERMILION, "Transient Disequilibrium"),
        ("3. The 9.3x Modern Velocity Tracking Deficit (RQ3)", [
            "Biological Speed Ceiling: Empirical 2D tree dispersal velocities hit a physical ceiling between 117 and 277 m/yr (Carya max 296 m/yr).",
            "Tracking Deficit: Past trees were strictly slower than natural deglacial warming (332 m/yr), and trail modern warming (1,574 m/yr) by 9.3×.",
            "Conservation Action: Passive dispersal will fail under 21st-century warming; proactive assisted migration is required."
        ], COLOR_GREEN, "Assisted Migration Mandate")
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

    add_bottom_banner(slide, "Final Synthesis",
                      "Climate change operates in decades, but forests reorganize across centuries. Understanding empirical historical limits is essential to safeguarding future forest ecosystems.")

    set_clean_notes(slide, 20, "Conclusions", "1 min 00s", "24:00",
        "To conclude, our study provides three empirical principles: "
        "First, tree taxa respond individualistically across 750- to 2,500-year clocks, rather than as synchronized biomes. "
        "Second, novel ecosystems peaked at 38% during rapid warming and dissolved to zero under stable pre-colonial conditions. "
        "Third, tree migration speeds hit an empirical ceiling below 300 m/yr, falling 9.3 times behind contemporary warming velocity. "
        "Active assisted migration is required to manage forest biodiversity through 21st-century climate change. "
        "Thank you. We welcome your questions.",
        "We also have two backup slides prepared covering 1D vs 2D mathematical differences and statistical parameters.",
        "Question: What is the single primary takeaway for foresters?\nAnswer: Natural tree migration cannot outrun modern climate velocity; assisted migration of southern seedlots must be planned now.")


def slide_21(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "Backup Slide (Q&A)", "Why Older 1D Literature Models Overestimated Speeds",
                     "Mathematical comparison: 1D linear regression slope inversion artifacts vs. true 2D vector gradients.")

    cards = [
        ("1D Slope Inversion Glitch", [
            "Linear Regression Formula: Arrival Age = beta_0 + beta_1 * Latitude.",
            "Slope Inversion: Speed was computed as 1 / beta_1, with confidence bounds 1 / (beta_1 ± t * SE).",
            "Asymptotic Explosion: With modest sample sizes, the lower slope bound approached zero, causing 1/slope to explode past 600 m/yr.",
            "Result: Created false speeds that artificially crossed past climate velocity (332 m/yr)."
        ], "Artifact: 1 / 0 -> Infinity", COLOR_VERMILION),
        ("Straight-North Blindspot", [
            "Fixed Azimuth Assumption: 1D models assumed trees only traveled strictly due North (0°).",
            "Topographic Detours Ignored: Lateral movement along river valleys or around the Great Lakes produced near-zero latitude change.",
            "Result: 1D models misinterpreted sideways dispersal as near-instantaneous, heavily biasing speed estimates."
        ], "Lateral Detours Missed", COLOR_ORANGE),
        ("2D Equal-Area Physical Reality", [
            "Planar Arrival Surfaces: Fitted Age(x, y) = beta_0 + beta_x*x + beta_y*y in equal-area Albers coordinates (km).",
            "True Gradient Vector: Speed = 1 / sqrt(beta_x^2 + beta_y^2) along actual migration corridors (e.g. Carya at 65° ENE).",
            "Empirical Ceiling: Bounded by real biological limits between 117 and 277 m/yr, strictly below past climate velocity (332 m/yr)."
        ], "Physically Honest Ceiling", COLOR_BLUE)
    ]

    for idx, (title, bullets, chip, col) in enumerate(cards):
        add_card_column(slide, 0.8 + idx * 3.98, 1.55, 3.75, 4.30, title, bullets, bottom_chip=chip, title_color=col)

    add_bottom_banner(slide, "Mathematical Resolution",
                      "The apparent crossing of past climate velocity in older 1D models was a mathematical artifact of inverting small regression slopes. True 2D equal-area gradients reveal a strict biological ceiling under 300 m/yr.")

    set_clean_notes(slide, 21, "Backup 1D vs 2D Math", "Backup", "Buffer",
        "Use this slide if asked why older literature reported tree speeds of 500 to 600 m/yr while our study finds speeds under 300 m/yr. "
        "Older studies ran 1D linear regressions: Arrival Age = Slope * Latitude, and inverted the slope: Speed = 1 / Slope. "
        "With small sample sizes, the lower bound of the slope touched near zero. Dividing by near-zero caused upper confidence bounds to explode to 600 m/yr. "
        "Furthermore, 1D models forced trees straight north, ignoring lateral routes around lakes. "
        "Our 2D planar arrival surfaces calculate true vector gradients in equal-area coordinates, providing physically honest biological speeds.",
        "Transition back to Q&A.")


def slide_22(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_slide_header(slide, "Backup Slide (Q&A)", "Master Statistical & Methodological Parameter Reference",
                     "Comprehensive reference table of analytical thresholds, calibration parameters, and statistical test criteria.")

    headers = ["Analytical Parameter", "Scientific Value Used", "Methodological Rationale & Reference"]
    rows = [
        ["Study Corridor", "25°–50°N, 95°–65°W", "Eastern North American temperate-boreal transition zone"],
        ["Sediment Cores", "69 continuous lakes (1,706 bins)", "Min 150 grains/sample, min 4 continuous 500-yr bins"],
        ["Chronology Framework", "LegacyAge 1.0 (Li et al. 2022)", "Calibrated Bayesian radiocarbon chronologies (PANGAEA Table-S5)"],
        ["Age Uncertainty", "N = 50 Monte Carlo draws", "Resampled depth-age bounds; cross-genus median lag locked at 1,500 yr"],
        ["Pollen Productivity", "REVEALS RPP Multipliers", "Wieczorek & Herzschuh (2020) & Dawson et al. (2016)"],
        ["Taxon Isolation", "TRSH-only (Woody Taxa)", "Isolates arboreal trees & shrubs; eliminates post-settlement farm weeds"],
        ["No-Analogue Metric", "Squared Chord Distance (SCD)", "Threshold = 0.138 (5th percentile of 611 modern surface sites)"],
        ["Stationary Lag Test", "First-differencing (Δy vs ΔT)", "Removes 15 ka trend; tested against 1,000 Ebisuzaki surrogates"],
        ["2D Arrival Surface", "Age(x, y) = b0 + bx*x + by*y", "Albers Equal-Area (EPSG:5070); gated by planar F-test (p <= 0.05)"],
        ["Modern Climate Velocity", "1,574 m/yr (1,242–2,071 range)", "1.9°C/century warming / 0.00121°C/km regional spatial gradient"]
    ]

    col_w = [Inches(2.5), Inches(3.2), Inches(6.033)]
    add_table(slide, 0.8, 1.55, 11.733, 4.30, headers, rows, col_w)

    add_bottom_banner(slide, "Reproducibility Baseline",
                      "All thresholds, parameters, and algorithms are fully scripted and open-source, ensuring complete analytical reproducibility across all stages of the study.")

    set_clean_notes(slide, 22, "Backup Parameters", "Backup", "Buffer",
        "Use this slide if committee members ask for exact mathematical thresholds, calibration sources, or analytical parameters. "
        "Key points to reference: "
        "1. RPP multipliers: Wieczorek & Herzschuh (2020). "
        "2. Chronologies: Li et al. (2022) LegacyAge 1.0. "
        "3. No-analogue threshold: 0.138 Squared Chord Distance based on the 5th percentile of 611 modern surface samples. "
        "4. Surrogate testing: 1,000 Ebisuzaki phase-randomized surrogate time series. "
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
